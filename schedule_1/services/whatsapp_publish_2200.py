"""22:00 MYT: independent auto-publish and TEST_GROUP_ONLY roster notification.

SAFETY: automatic publishing requires explicit admin switch and full transactional validation.
Do not deploy without staging tests and a database backup.
"""
import logging
import os
from datetime import datetime, timedelta

from db import get_conn
from schedule.builders.time_utils import malaysia_now

LOG = logging.getLogger(__name__)


def _settings(keys):
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute('SELECT key, value FROM schedule_settings WHERE key = ANY(%s)', (list(keys),))
            rows = cur.fetchall()
    values = {r['key'] if isinstance(r, dict) else r[0]:
              r['value'] if isinstance(r, dict) else r[1] for r in rows}
    return {k: str(values.get(k, 'false')).strip().lower() in ('true', '1', 'yes', 'on') for k in keys}


def _scalar(sql, params):
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(sql, params)
            row = cur.fetchone()
    return next(iter(row.values())) if isinstance(row, dict) else row[0]


def _validate_coverage(target):
    """严格检查每一笔有效值班报名的完整时间覆盖；只读，不修改数据库。"""
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                SELECT id, name, start_time, end_time
                FROM volunteer_schedule_signups
                WHERE signup_date=%s AND role='值班'
                  AND COALESCE(status,'pending') <> 'cancelled'
            """, (target,))
            rows = cur.fetchall()
            signup_cols = [c[0] for c in cur.description]
            cur.execute("""
                SELECT signup_id, start_time, end_time
                FROM volunteer_schedule_assignments
                WHERE assignment_date=%s AND role='值班'
                  AND COALESCE(status,'assigned') <> 'cancelled'
            """, (target,))
            assignments = cur.fetchall()
            assignment_cols = [c[0] for c in cur.description]
            cur.execute("""
                SELECT count(*) FROM volunteer_schedule_assignments
                WHERE assignment_date=%s
                  AND COALESCE(status,'assigned') <> 'cancelled'
            """, (target,))
            count_row = cur.fetchone()

    def as_dict(row, cols):
        return dict(row) if isinstance(row, dict) else dict(zip(cols, row))

    by_signup = {}
    for row in assignments:
        item = as_dict(row, assignment_cols)
        if item['signup_id'] is not None:
            by_signup.setdefault(str(item['signup_id']), []).append(item)

    problems = []
    for row in rows:
        signup = as_dict(row, signup_cols)
        required = _dry_interval(signup)
        intervals = [_dry_interval(a) for a in by_signup.get(str(signup['id']), [])]
        intervals = [i for i in intervals if i is not None]
        if not _dry_covers(required, intervals):
            problems.append(f"{signup['id']} {signup['name']} ({signup['start_time']}-{signup['end_time']})")
    if problems:
        raise RuntimeError('值班报名时间未完整覆盖，停止发布及通知：' + '; '.join(problems[:12]))
    total = next(iter(count_row.values())) if isinstance(count_row, dict) else count_row[0]
    if not int(total):
        raise RuntimeError('数据库没有有效排班记录；停止发布及通知')
    return int(total)


def run_2200_tick():
    now = malaysia_now()
    if not (now.hour == 22 and now.minute < 20):
        return {'status': 'outside_window', 'malaysia_time': now.isoformat()}

    flags = _settings(('wa_auto_enabled', 'wa_test_mode', 'wa_roster_2200', 'wa_auto_publish_2200'))
    auto_publish = flags['wa_auto_publish_2200']
    send_roster = flags['wa_auto_enabled'] and flags['wa_test_mode'] and flags['wa_roster_2200']
    if not auto_publish and not send_roster:
        return {'status': 'disabled'}

    target = (now.date() + timedelta(days=1)).isoformat()
    marker = 'auto_roster_2200:' + target
    with get_conn() as lock_conn:
        with lock_conn.cursor() as cur:
            cur.execute('SELECT pg_try_advisory_lock(hashtext(%s), hashtext(%s))', ('gyt_wa_2200', target))
            row = cur.fetchone()
            acquired = row['pg_try_advisory_lock'] if isinstance(row, dict) else row[0]
            if not acquired:
                return {'status': 'busy', 'target_date': target}
            try:
                return _execute_locked(target, marker, auto_publish, send_roster)
            finally:
                cur.execute('SELECT pg_advisory_unlock(hashtext(%s), hashtext(%s))', ('gyt_wa_2200', target))


def _execute_locked(target, marker, auto_publish, send_roster):
    from schedule.services.publish_service import is_schedule_published
    from schedule.builders.schedule_builder import publish_full_schedule_2200_atomic
    from schedule.services.whatsapp_service import build_whatsapp_from_assigned

    published = is_schedule_published(target)
    preserved = bool(_scalar('SELECT count(*) FROM volunteer_schedule_assignments WHERE assignment_date=%s', (target,)))
    # 开启自动发布才允许写入；使用与负责人蓝色按钮相同的排班计算器。
    # 事务中同时写入全部岗位并发布；失败则整体回滚。
    if auto_publish and not published:
        if os.getenv("WA_2200_ALLOW_DB_WRITE", "").strip() != "1":
            return {"status": "db_write_guard_off", "target_date": target,
                    "detail": "22:00 自动排班尚未通过独立测试数据库验证，禁止写入正式数据库"}
        result = publish_full_schedule_2200_atomic(target)
        LOG.info('22:00 full auto schedule/publish result: %s', result)
        published = is_schedule_published(target)
        if not published:
            raise RuntimeError('22:00 publish not confirmed')

    if published:
        from schedule.services.whatsapp_morning_summary import capture_2200_snapshot
        capture_2200_snapshot(target)  # 已发布就建立首次快照，不依赖发送开关。

    if not send_roster:
        return {'status': 'published_no_message' if published else 'no_action',
                'target_date': target, 'preserved_existing': preserved}
    if not published:
        return {'status': 'not_published', 'target_date': target}

    return finalize_published_roster(target)


def finalize_published_roster(target):
    '''首次发布后建立不可覆盖的基准并入队正式表；可由自动或手动发布调用。

    快照和通知 Job 同一事务提交；相同日期的重复调用不会重复发送。
    仅限现有 TEST_GROUP_ONLY 测试模式及已开启的自动正式表通知。
    '''
    from schedule.services.publish_service import is_schedule_published
    from schedule.services.whatsapp_service import build_whatsapp_from_assigned
    from schedule.services.whatsapp_morning_summary import _assignments, _key
    import json

    if not is_schedule_published(target):
        return {'status': 'not_published', 'target_date': target}

    # 通知验证失败也不应影响已发布排班的首次快照。
    from schedule.services.whatsapp_morning_summary import capture_2200_snapshot
    capture_2200_snapshot(target)

    flags = _settings(('wa_auto_enabled', 'wa_test_mode', 'wa_roster_2200'))
    if not all(flags.values()):
        return {'status': 'disabled', 'target_date': target}

    # 先核对实际值班时间覆盖，避免将不完整的值班表发送出去。
    _validate_coverage(target)
    body = build_whatsapp_from_assigned(target)
    if not isinstance(body, str) or not body.strip() or body.lstrip().startswith('❌'):
        raise RuntimeError('正式值班表文案生成失败；停止通知')

    marker = 'auto_roster_2200:' + target  # 沿用旧标记，防止重复发送。
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute('SELECT pg_advisory_xact_lock(hashtext(%s), hashtext(%s))',
                        ('gyt_morning_baseline', target))
            # 在同一事务内建立基准和 Job，防止只建快照却未入队。
            # 首次基准不可覆盖；发布后新增报名只在早上汇总通知。
            cur.execute('SELECT baseline FROM whatsapp_morning_snapshots WHERE target_date=%s', (target,))
            previous_snapshot = cur.fetchone()
            if not previous_snapshot:
                rows = _assignments(cur, target)
                cur.execute('''INSERT INTO whatsapp_morning_snapshots(target_date, baseline)
                    VALUES (%s, %s::jsonb) ON CONFLICT(target_date) DO NOTHING''',
                    (target, json.dumps([list(_key(r)) for r in rows], ensure_ascii=False)))
            cur.execute('''SELECT id FROM whatsapp_notification_jobs
                WHERE created_by=%s AND target_date=%s ORDER BY id LIMIT 1''', (marker, target))
            existing = cur.fetchone()
            if existing:
                job_id = existing['id'] if isinstance(existing, dict) else existing[0]
                result = {'status': 'already_queued', 'target_date': target, 'job_id': job_id}
            else:
                cur.execute('''INSERT INTO whatsapp_notification_jobs
                    (message_type, target_date, destination, body, status, created_by)
                    VALUES ('test', %s, 'TEST_GROUP_ONLY', %s, 'pending', %s)
                    RETURNING id''', (target, body, marker))
                row = cur.fetchone()
                job_id = row['id'] if isinstance(row, dict) else row[0]
                result = {'status': 'queued', 'target_date': target, 'job_id': job_id}
        conn.commit()
    return result


# ---------------------------------------------------------------------------
# 22:00 安全预演：只读数据库，不发布、不入队、不修改任何排班。
# 不调用 insert_2200_assignments_if_empty / sync_patch_assignments。
# ---------------------------------------------------------------------------
def _dry_minutes(value):
    if value is None:
        return None
    value = str(value).strip().lower().replace(' ', '')
    if value.endswith('am') or value.endswith('pm'):
        if ':' not in value:
            value = value[:-2] + ':00' + value[-2:]
    for fmt in ('%H:%M:%S', '%H:%M', '%I:%M%p'):
        try:
            parsed = datetime.strptime(value, fmt)
            return parsed.hour * 60 + parsed.minute
        except ValueError:
            continue
    return None


def _dry_interval(row):
    start = _dry_minutes(row.get('start_time'))
    end = _dry_minutes(row.get('end_time'))
    if start is None or end is None or end <= start:
        return None
    return (start, end)


def _dry_merge(intervals):
    merged = []
    for start, end in sorted(intervals):
        if merged and start <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], end))
        else:
            merged.append((start, end))
    return merged


def _dry_covers(required, existing):
    if not required:
        return False
    merged = _dry_merge(existing)
    return any(start <= required[0] and end >= required[1] for start, end in merged)


def dry_run_2200(target_date=None):
    """安全预演（显式调用），返回可序列化报告；绝不写数据库或发 WhatsApp。

    默认预演马来西亚时间的明天。target_date 可指定 YYYY-MM-DD。
    这里只检查值班报名覆盖与理论/实际排班差异；不会声称卫生已完整验证。
    """
    if target_date is None:
        target_date = (malaysia_now().date() + timedelta(days=1)).isoformat()
    target_date = datetime.strptime(str(target_date), '%Y-%m-%d').date().isoformat()

    from schedule.services.publish_service import is_schedule_published
    from schedule.builders.schedule_builder import calculate_schedule_for_date

    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                SELECT id, name, volunteer_id, role, start_time, end_time, status
                FROM volunteer_schedule_signups
                WHERE signup_date=%s AND role='值班'
                  AND coalesce(status,'pending') <> 'cancelled'
                ORDER BY id
            """, (target_date,))
            signup_rows = cur.fetchall()
            signup_cols = [d[0] for d in cur.description]
            cur.execute("""
                SELECT signup_id, volunteer_id, name, role, shift_label,
                       assigned_place, start_time, end_time
                FROM volunteer_schedule_assignments
                WHERE assignment_date=%s
                ORDER BY id
            """, (target_date,))
            assignment_rows = cur.fetchall()
            assignment_cols = [d[0] for d in cur.description]

    def as_dict(row, cols):
        return dict(row) if isinstance(row, dict) else dict(zip(cols, row))

    signups = [as_dict(row, signup_cols) for row in signup_rows]
    existing = [as_dict(row, assignment_cols) for row in assignment_rows]

    # 正式算法仅做内存计算，不调用写入函数。
    _, _, calculated_rows, _, _ = calculate_schedule_for_date(target_date)
    calculated = [dict(row) for row in calculated_rows]

    by_signup = {}
    for row in existing:
        if row.get('signup_id') is not None and row.get('role') == '值班':
            by_signup.setdefault(str(row['signup_id']), []).append(row)

    theoretical = {}
    for row in calculated:
        if row.get('signup_id') is not None and row.get('role') == '值班':
            theoretical.setdefault(str(int(row['signup_id'])), []).append(row)

    checks = []
    for signup in signups:
        sid = str(signup['id'])
        actual = by_signup.get(sid, [])
        expected = theoretical.get(sid, [])
        required = _dry_interval(signup)
        intervals = [_dry_interval(row) for row in actual]
        intervals = [interval for interval in intervals if interval]
        covered = _dry_covers(required, intervals)
        if required is None:
            status = 'invalid_signup_time'
        elif not actual:
            status = 'missing'
        elif not covered:
            status = 'partial_or_conflicting'
        else:
            status = 'covered'
        checks.append({
            'signup_id': signup['id'],
            'name': signup['name'],
            'status': status,
            'signup_time': [str(signup.get('start_time')), str(signup.get('end_time'))],
            'existing': [{k: str(row.get(k) or '') for k in
                          ('shift_label', 'assigned_place', 'start_time', 'end_time')}
                         for row in actual],
            'theoretical': [{k: str(row.get(k) or '') for k in
                             ('shift_label', 'assigned_place', 'start_time', 'end_time')}
                            for row in expected],
        })

    return {
        'mode': 'DRY_RUN_READ_ONLY',
        'target_date': target_date,
        'published': bool(is_schedule_published(target_date)),
        'existing_assignment_count': len(existing),
        'theoretical_assignment_count': len(calculated),
        'duty_signup_count': len(signups),
        'coverage_checks': checks,
        'missing_count': sum(c['status'] == 'missing' for c in checks),
        'problem_count': sum(c['status'] != 'covered' for c in checks),
        'warnings': [
            '仅预演：不会保存排班、发布、建立 WhatsApp 任务或发送消息。',
            '理论排班没有自动扣除已有安排；不同地点或重叠时间需人工核对。',
            '本报告只验证值班报名的时间覆盖，不验证卫生及其他岗位的完整性。',
            '若计算依赖本地 fixed_schedule.xlsx，请在具备该文件的环境执行。',
        ],
    }
