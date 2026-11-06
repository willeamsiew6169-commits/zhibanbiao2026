"""Render Cron Job entrypoint for the temple's 22:00 Malaysia-time roster.

Render cron expression: 0 14 * * * (UTC = 22:00 Asia/Kuala_Lumpur)
Command: python render_whatsapp_2200.py

This imports the existing tested scheduler; it does NOT bypass any feature
flags, DB-write guard, locking, validation or idempotency protections.
"""
import json
import logging
import sys
from datetime import datetime
from zoneinfo import ZoneInfo

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
LOG = logging.getLogger('render_whatsapp_2200')


def main():
    now = datetime.now(ZoneInfo('Asia/Kuala_Lumpur'))
    LOG.info('Cron started, Malaysia time: %s', now.isoformat())
    if now.hour != 22 or now.minute >= 20:
        LOG.warning('Outside 22:00–22:19 MYT window; skipping')
        return 0

    # Import lazily so the time check works even without database settings.
    from schedule.services.whatsapp_publish_2200 import run_2200_tick
    result = run_2200_tick()
    LOG.info('22:00 tick result: %s', json.dumps(result, ensure_ascii=False, default=str))
    status = result.get('status') if isinstance(result, dict) else None
    # Surface actual failures in Render Cron Job logs / status.
    if status in {'db_write_guard_off', 'not_published', 'no_action'}:
        LOG.warning('22:00 did not complete publication/queue: %s', status)
    if status in {'db_write_guard_off'}:
        return 2
    return 0


if __name__ == '__main__':
    try:
        sys.exit(main())
    except Exception:
        LOG.exception('22:00 scheduled run failed')
        sys.exit(1)
