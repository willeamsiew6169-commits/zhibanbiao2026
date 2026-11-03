"""19:00 MYT signup notice -> TEST_GROUP_ONLY queue; never publishes schedules.
Called by authenticated /schedule/api/whatsapp-agent/notice-1900/tick.
"""
from datetime import timedelta
from schedule.builders.time_utils import malaysia_now
from schedule.services.settings_service import is_schedule_setting_on
from db import get_conn


def enqueue_tomorrow_signup_notice(build_notice):
    now = malaysia_now()
    if not (now.hour == 19 and 0 <= now.minute < 20):
        return {'status': 'outside_window', 'malaysia_time': now.isoformat()}
    required = ('wa_auto_enabled', 'wa_test_mode', 'wa_notice_1900')
    if not all(is_schedule_setting_on(k) for k in required):
        return {'status': 'disabled'}
    target = (now.date() + timedelta(days=1)).isoformat()
    marker = 'auto_signup_1900:' + target
    # Transaction-level advisory lock prevents duplicate jobs across Render workers.
    # The existing job table requires no schema changes.
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute('SELECT pg_advisory_xact_lock(hashtext(%s), hashtext(%s))',
                        ('gyt_whatsapp_signup_1900', target))
            cur.execute('''SELECT id, status FROM whatsapp_notification_jobs
                           WHERE created_by=%s AND target_date=%s
                           ORDER BY id DESC LIMIT 1''', (marker, target))
            previous = cur.fetchone()
            if previous:
                return {'status': 'already_queued', 'target_date': target,
                        'job_id': previous[0] if not isinstance(previous, dict) else previous['id']}
            # Build under lock to avoid duplicate enqueue when tick requests race.
            body = build_notice(target)
            if not isinstance(body, str) or not body.strip():
                raise ValueError('Signup notice builder returned an empty message')
            cur.execute('''INSERT INTO whatsapp_notification_jobs
                (message_type, target_date, destination, body, status, created_by)
                VALUES ('test', %s, 'TEST_GROUP_ONLY', %s, 'pending', %s)
                RETURNING id''', (target, body, marker))
            row = cur.fetchone()
            job_id = row['id'] if isinstance(row, dict) else row[0]
        conn.commit()
    return {'status': 'queued', 'target_date': target, 'job_id': job_id}
