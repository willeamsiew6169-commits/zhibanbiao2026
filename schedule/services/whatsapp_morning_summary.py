"""05:30 special day / 07:00 normal day: one test-group late-arrival roster notice.

Requires migration whatsapp_morning_snapshot.sql. Snapshot is captured at 22:00
before the formal roster job is queued. Fail closed when snapshot is missing.
"""
import json
from datetime import datetime, timedelta
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
    return str(value or '')[:5]


def _message(target, rows):
    parts = [f'🌅 {target} 早上新增义工安排', '（昨晚22:00正式值班表后新增）', '']
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


def run_morning_tick():
    now = malaysia_now()
    target = now.date().isoformat()
    info = get_special_day_info(now.date())
    special = info.get('template_type') != 'normal'
    hour, minute = (5, 30) if special else (7, 0)
    start = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
    if not (start <= now < start + timedelta(minutes=20)):
        return {'status': 'outside_window', 'scheduled_time': f'{hour:02d}:{minute:02d}',
                'malaysia_time': now.isoformat()}

    marker = 'auto_morning_late:' + target
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
            cur.execute('SELECT id FROM whatsapp_notification_jobs WHERE created_by=%s AND target_date=%s',
                        (marker, target))
            if cur.fetchone():
                return {'status': 'already_queued', 'target_date': target}
            current = _assignments(cur, target)
            added = [r for r in current if _key(r) not in old_keys]
            if not added:
                return {'status': 'no_new_assignments', 'target_date': target}
            body = _message(target, added)
            cur.execute('''INSERT INTO whatsapp_notification_jobs
                (message_type, target_date, destination, body, status, created_by)
                VALUES ('test', %s, 'TEST_GROUP_ONLY', %s, 'pending', %s) RETURNING id''',
                        (target, body, marker))
            result = cur.fetchone()
            job_id = result['id'] if isinstance(result, dict) else result[0]
        conn.commit()
    return {'status': 'queued', 'target_date': target, 'job_id': job_id,
            'new_assignment_rows': len(added)}
