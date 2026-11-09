"""Test-group late-signup summary: first at 05:30 (special) / 07:00 (normal), then every 30 minutes until 18:30 MYT.

Requires whatsapp_morning_snapshot.sql and whatsapp_morning_batches.sql.
Fail closed when the 22:00 snapshot or batch ledger is missing.
"""
import json
import re
from db import get_conn
from lunar_rules import get_special_day_info
from schedule.builders.time_utils import malaysia_now

ROLES = ('值班', '卫生')


def _rows(cur):
    data = cur.fetchall()
    cols = [d[0] for d in cur.description]
    return [dict(r) if isinstance(r, dict) else dict(zip(cols, r)) for r in data]


def _assignments(cur, target):
    cur.execute('''SELECT volunteer_id, name, role, assigned_place, start_time, end_time
        FROM volunteer_schedule_assignments
        WHERE assignment_date=%s AND role IN ('值班','卫生')
          AND COALESCE(status,'assigned') <> 'cancelled'
        ORDER BY role, name, start_time, assigned_place''', (target,))
    return _rows(cur)


def _key(row):
    return (str(row.get('volunteer_id') or '').strip(), str(row['role']))


def capture_2200_snapshot(target):
    """Capture once; never overwrite a previous 22:00 baseline."""
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute('SELECT pg_advisory_xact_lock(hashtext(%s), hashtext(%s))',
                        ('gyt_morning_baseline', target))
            rows = _assignments(cur, target)
            cur.execute('''INSERT INTO whatsapp_morning_snapshots(target_date, baseline)
                VALUES (%s, %s::jsonb) ON CONFLICT(target_date) DO NOTHING''',
                        (target, json.dumps([list(_key(r)) for r in rows], ensure_ascii=False)))
        conn.commit()


def _time(value):
    """Render schedule time in 12-hour am/pm format (supports SQL time and am/pm)."""
    if value is None:
        return ''
    raw = str(value).strip()
    match = re.fullmatch(r'(\d{1,2}):(\d{2})(?::\d{2}(?:\.\d+)?)?\s*([ap]\.?m\.?)?', raw, re.I)
    if not match:
        return raw
    hour, minute = int(match.group(1)), int(match.group(2))
    suffix = (match.group(3) or '').lower().replace('.', '')
    if minute > 59 or hour > 23 or (suffix and not 1 <= hour <= 12):
        return raw
    if suffix == 'am':
        hour = hour % 12
    elif suffix == 'pm':
        hour = hour % 12 + 12
    shown_hour = hour % 12 or 12
    meridiem = 'am' if hour < 12 else 'pm'
    return f'{shown_hour}:{minute:02d}{meridiem}'


def _message(target, rows):
    parts = [f'📋 {target} 新增义工安排', '']
    for role in ROLES:
        selected = [r for r in rows if r['role'] == role]
        if not selected:
            continue
        parts.append('🧹 卫生' if role == '卫生' else '🧡 值班')
        by_person = {}
        for r in selected:
            by_person.setdefault(_key(r), []).append(r)
        for entries in by_person.values():
            first = entries[0]
            parts.append(f"• {first.get('name') or first.get('volunteer_id') or '义工'}")
            for r in entries:
                parts.append(f"  {_time(r['start_time'])}–{_time(r['end_time'])}  {r.get('assigned_place') or role}")
        parts.append('')
    parts.append('🙏 感恩各位义工发心护持。')
    return '\n'.join(parts)


def _slot(now):
    """Notify only for today's duty, after the morning opening and before 19:00.

    Previous evening 22:00 through the night is accumulation-only: no messages.
    Normal days start 07:00; special days start 05:30.
    """
    info = get_special_day_info(now.date())
    special = info.get('template_type') != 'normal'
    opening_minutes = 5 * 60 + 30 if special else 7 * 60
    current_minutes = now.hour * 60 + now.minute
    if current_minutes < opening_minutes or current_minutes >= 19 * 60:
        return None
    if not (now.minute < 20 or 30 <= now.minute < 50):
        return None
    return f'{now.hour:02d}:{"00" if now.minute < 20 else "30"}'


def run_morning_tick():
    now = malaysia_now()
    slot = _slot(now)
    # Always today's assignments: after 22:00 yesterday, wait for today's first morning batch.
    target = now.date().isoformat()
    if slot is None:
        return {'status': 'outside_window', 'target_date': target,
                'malaysia_time': now.isoformat()}

    marker = 'auto_morning_late:' + target + ':' + now.date().isoformat() + ':' + slot.replace(':', '')
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute('SELECT pg_advisory_xact_lock(hashtext(%s), hashtext(%s))',
                        ('gyt_morning_notice', target))
            cur.execute("SELECT key, value FROM schedule_settings WHERE key = ANY(%s)",
                        (['wa_auto_enabled', 'wa_test_mode', 'wa_late_signup'],))
            settings = {r['key'] if isinstance(r, dict) else r[0]:
                        str(r['value'] if isinstance(r, dict) else r[1]).lower().strip()
                        for r in cur.fetchall()}
            if not all(settings.get(k) in ('true', '1', 'yes', 'on')
                       for k in ('wa_auto_enabled', 'wa_test_mode', 'wa_late_signup')):
                return {'status': 'disabled'}
            cur.execute('SELECT baseline FROM whatsapp_morning_snapshots WHERE target_date=%s', (target,))
            row = cur.fetchone()
            if not row:
                return {'status': 'missing_2200_snapshot', 'target_date': target}
            baseline = row['baseline'] if isinstance(row, dict) else row[0]
            if isinstance(baseline, str):
                baseline = json.loads(baseline)
            old_keys = {tuple(item) for item in baseline}

            # 兼容旧版07:00通知：只排除旧消息正文明确列出的姓名，
            # 不再因为存在旧 Job 而禁止整天的新批次。
            cur.execute('''SELECT body, status FROM whatsapp_notification_jobs
                WHERE target_date=%s AND created_by=%s
                ORDER BY id''', (target, 'auto_morning_late:' + target))
            legacy_jobs = _rows(cur)
            # 旧版通知正文使用「• 姓名」逐人列出；若格式无法识别，
            # 为避免重复发送，保守停止并要求人工核查。
            legacy_names = set()
            for legacy in legacy_jobs:
                body_text = str(legacy.get('body') or '')
                names = re.findall(r'^\s*•\s*(.+?)\s*$', body_text, re.M)
                if not names:
                    return {'status': 'legacy_unparseable', 'target_date': target}
                legacy_names.update(names)

            # Ledger contains all previously queued person/role keys, even if ACK is uncertain.
            # This prevents re-sending a message after uncertain or a network interruption.
            cur.execute('''SELECT slot_time, notified_keys FROM whatsapp_morning_batches
                WHERE target_date=%s ORDER BY slot_time''', (target,))
            batches = _rows(cur)
            for batch in batches:
                keys = batch['notified_keys']
                if isinstance(keys, str):
                    keys = json.loads(keys)
                old_keys.update(tuple(item) for item in keys)
            if any(str(batch['slot_time'])[:5] == slot for batch in batches):
                return {'status': 'already_queued', 'target_date': target, 'slot': slot}

            current = _assignments(cur, target)
            added = [r for r in current
                     if _key(r) not in old_keys
                     and str(r.get('name') or '').strip() not in legacy_names]
            if not added:
                return {'status': 'no_new_assignments', 'target_date': target, 'slot': slot}
            body = _message(target, added)
            cur.execute('''INSERT INTO whatsapp_notification_jobs
                (message_type, target_date, destination, body, status, created_by)
                VALUES ('test', %s, 'TEST_GROUP_ONLY', %s, 'pending', %s) RETURNING id''',
                        (target, body, marker))
            result = cur.fetchone()
            job_id = result['id'] if isinstance(result, dict) else result[0]
            keys = sorted({ _key(r) for r in added })
            cur.execute('''INSERT INTO whatsapp_morning_batches
                (target_date, slot_time, notified_keys, job_id)
                VALUES (%s, %s, %s::jsonb, %s)''',
                        (target, slot, json.dumps(keys, ensure_ascii=False), job_id))
        conn.commit()
    return {'status': 'queued', 'target_date': target, 'slot': slot,
            'job_id': job_id, 'new_assignment_rows': len(added)}
