"""22:00 MYT automatic next-day roster, phase-2 TEST_GROUP_ONLY.
Only runs after 22:00 and before 22:20, and only with all required switches enabled.
"""
from datetime import timedelta
import logging
from db import get_conn
from schedule.builders.time_utils import malaysia_now

LOG = logging.getLogger(__name__)


def _enabled(keys):
    # Read settings directly from DB: settings_service uses a per-process cache.
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute('SELECT key, value FROM schedule_settings WHERE key = ANY(%s)', (list(keys),))
            values = {r[0] if not isinstance(r, dict) else r['key']:
                      r[1] if not isinstance(r, dict) else r['value'] for r in cur.fetchall()}
    return all(str(values.get(k, 'false')).lower().strip() in ('true', '1', 'yes', 'on') for k in keys)


def run_2200_tick():
    now = malaysia_now()
    if not (now.hour == 22 and now.minute < 20):
        return {'status': 'outside_window', 'malaysia_time': now.isoformat()}
    required = ('wa_auto_enabled', 'wa_test_mode', 'wa_roster_2200', 'wa_auto_publish_2200')
    if not _enabled(required):
        return {'status': 'disabled'}

    target = (now.date() + timedelta(days=1)).isoformat()
    marker = 'auto_roster_2200:' + target
    # Session-level advisory lock remains held across schedule calculations, writes,
    # publish and enqueue (each helper opens its own DB connection).
    with get_conn() as lock_conn:
        with lock_conn.cursor() as lock_cur:
            lock_cur.execute('SELECT pg_try_advisory_lock(hashtext(%s), hashtext(%s))',
                             ('gyt_wa_2200', target))
            row = lock_cur.fetchone()
            acquired = row['pg_try_advisory_lock'] if isinstance(row, dict) else row[0]
            if not acquired:
                return {'status': 'busy', 'target_date': target}
            try:
                return _execute_locked(target, marker)
            finally:
                lock_cur.execute('SELECT pg_advisory_unlock(hashtext(%s), hashtext(%s))',
                                 ('gyt_wa_2200', target))


def _execute_locked(target, marker):
    from schedule.services.publish_service import is_schedule_published, publish_schedule_for_date
    from schedule.builders.schedule_builder import calculate_schedule_for_date, update_assigned_places
    from schedule.services.whatsapp_service import build_whatsapp_from_assigned

    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute('SELECT id FROM whatsapp_notification_jobs WHERE created_by=%s AND target_date=%s LIMIT 1',
                        (marker, target))
            row = cur.fetchone()
            if row:
                return {'status': 'already_queued', 'target_date': target}

    was_published = is_schedule_published(target)
    # 方案 A：只要数据库已有该日排班，绝不重新计算或覆盖。
    # 先检查已存记录，避免负责人手工排班尚未发布时被删除。
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                SELECT count(*) FROM volunteer_schedule_assignments
                WHERE assignment_date = %s
            """, (target,))
            result = cur.fetchone()
            existing_count = result['count'] if isinstance(result, dict) else result[0]

    used_existing = existing_count > 0
    if not was_published and not used_existing:
        # 仅在完全没有排班记录时自动生成。
        _, _, rows, _, _ = calculate_schedule_for_date(target)
        if not rows:
            raise RuntimeError('排班结果为空；停止自动发布')
        if not all(isinstance(r, dict) for r in rows):
            raise RuntimeError('排班记录格式异常；停止自动发布')
        # update_assigned_places 在事务中写入；异常直接上抛，禁止发布。
        update_assigned_places(target, rows)
        with get_conn() as conn:
            with conn.cursor() as cur:
                cur.execute("""
                    SELECT count(*) FROM volunteer_schedule_assignments
                    WHERE assignment_date = %s
                """, (target,))
                result = cur.fetchone()
                count = result['count'] if isinstance(result, dict) else result[0]
        if count != len(rows):
            raise RuntimeError(f'排班写入数量不符：预期 {len(rows)}，实际 {count}；停止发布')

    # 无论是负责人手工排班，还是自动排班，都先验证数据库内容。
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                SELECT count(*) FROM volunteer_schedule_assignments
                WHERE assignment_date = %s
            """, (target,))
            result = cur.fetchone()
            final_count = result['count'] if isinstance(result, dict) else result[0]
    if final_count == 0:
        raise RuntimeError('数据库没有正式排班记录；停止发布和通知')

    # 文案生成成功后才发布，防止发布空白值班表。
    body = build_whatsapp_from_assigned(target)
    if not isinstance(body, str) or not body.strip() or body.lstrip().startswith('❌'):
        raise RuntimeError('正式值班表文案生成失败；停止发布和通知')

    if not was_published:
        publish_schedule_for_date(target, published_by='AUTO_2200')

    with get_conn() as conn:
        with conn.cursor() as cur:
            # Dedup under session lock; only TEST_GROUP_ONLY jobs can be claimed by agent.
            cur.execute('SELECT id FROM whatsapp_notification_jobs WHERE created_by=%s AND target_date=%s LIMIT 1',
                        (marker, target))
            if cur.fetchone():
                return {'status': 'already_queued', 'target_date': target}
            cur.execute('''INSERT INTO whatsapp_notification_jobs
                (message_type, target_date, destination, body, status, created_by)
                VALUES ('test', %s, 'TEST_GROUP_ONLY', %s, 'pending', %s) RETURNING id''',
                (target, body, marker))
            row = cur.fetchone()
            job_id = row['id'] if isinstance(row, dict) else row[0]
        conn.commit()
    return {'status': 'queued', 'target_date': target, 'job_id': job_id,
            'already_published': was_published, 'preserved_existing': used_existing}
