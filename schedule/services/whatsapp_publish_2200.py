"""22:00 MYT: independent auto-publish and TEST_GROUP_ONLY roster notification.

SAFETY: database writes are opt-in via WA_2200_ALLOW_DB_WRITE=1.
Never turn this on while production and test services share the same database.
"""
import logging
import os
from datetime import timedelta

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
    """确保有效值班报名至少有一条对应值班 assignment；不强制目标人数。"""
    missing = _scalar("""
        SELECT count(*) FROM volunteer_schedule_signups s
        WHERE s.signup_date=%s AND s.role='值班'
          AND coalesce(s.status,'pending') <> 'cancelled'
          AND NOT EXISTS (
            SELECT 1 FROM volunteer_schedule_assignments a
            WHERE a.assignment_date=s.signup_date AND a.signup_id=s.id
              AND a.role='值班'
          )
    """, (target,))
    if int(missing):
        raise RuntimeError(f'还有 {missing} 位值班报名未排班；停止自动发布及发送')
    total = _scalar('SELECT count(*) FROM volunteer_schedule_assignments WHERE assignment_date=%s', (target,))
    if not int(total):
        raise RuntimeError('数据库没有排班记录；停止发布及通知')
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

    # Auto-publish is an actual DB write, even when WhatsApp is in test mode.
    if auto_publish and os.getenv('WA_2200_ALLOW_DB_WRITE', '').strip() != '1':
        return {'status': 'db_write_guard_off', 'detail': '需独立测试数据库并设置 WA_2200_ALLOW_DB_WRITE=1'}

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
    from schedule.services.publish_service import is_schedule_published, publish_schedule_for_date
    from schedule.builders.schedule_builder import calculate_schedule_for_date, insert_2200_assignments_if_empty
    from schedule.services.whatsapp_service import build_whatsapp_from_assigned

    published = is_schedule_published(target)
    preserved = bool(_scalar('SELECT count(*) FROM volunteer_schedule_assignments WHERE assignment_date=%s', (target,)))
    if auto_publish and not published:
        if not preserved:
            _, _, rows, _, _ = calculate_schedule_for_date(target)
            count = insert_2200_assignments_if_empty(target, rows)
            LOG.info('22:00 inserted %s assignments for %s', count, target)
        _validate_coverage(target)
        # Already published dates are never republished or reshuffled.
        publish_schedule_for_date(target, published_by='AUTO_2200')
        published = True

    if not send_roster:
        return {'status': 'published_no_message' if published else 'no_action',
                'target_date': target, 'preserved_existing': preserved}
    if not published:
        return {'status': 'not_published', 'target_date': target}

    # Even for previously published dates, do not send incomplete rosters.
    _validate_coverage(target)
    exists = _scalar('SELECT count(*) FROM whatsapp_notification_jobs WHERE created_by=%s AND target_date=%s',
                     (marker, target))
    if int(exists):
        return {'status': 'already_queued', 'target_date': target}

    body = build_whatsapp_from_assigned(target)
    if not isinstance(body, str) or not body.strip() or body.lstrip().startswith('❌'):
        raise RuntimeError('正式值班表文案生成失败；停止通知')
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                INSERT INTO whatsapp_notification_jobs
                (message_type, target_date, destination, body, status, created_by)
                VALUES ('test', %s, 'TEST_GROUP_ONLY', %s, 'pending', %s)
                RETURNING id
            """, (target, body, marker))
            row = cur.fetchone()
            job_id = row['id'] if isinstance(row, dict) else row[0]
        conn.commit()
    return {'status': 'queued', 'target_date': target, 'job_id': job_id,
            'already_published': not auto_publish, 'preserved_existing': preserved}
