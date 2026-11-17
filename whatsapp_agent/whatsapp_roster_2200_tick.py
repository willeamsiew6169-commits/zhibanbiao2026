"""22:00 WhatsApp roster tick client (Windows Task Scheduler).

Default: safe configuration check only (no HTTP request).
--execute: POST to existing authenticated endpoint. This MAY change the
shared database if server-side auto-publish settings permit it.
"""
import argparse
import os
import sys
from datetime import datetime
from zoneinfo import ZoneInfo

import requests
from pathlib import Path
from dotenv import load_dotenv

load_dotenv(Path(__file__).with_name(".env.agent"))


def main():
    parser = argparse.ArgumentParser(description='22:00 MYT roster automation')
    parser.add_argument('--execute', action='store_true', help='Call server; may enqueue/send or publish')
    args = parser.parse_args()

    now = datetime.now(ZoneInfo('Asia/Kuala_Lumpur'))
    print(f'Malaysia time: {now.isoformat()}', flush=True)
    if now.hour != 22 or now.minute >= 20:
        print('Outside 22:00–22:19 MYT; no request sent.', flush=True)
        return 0

    url = (
        os.getenv("WA_SERVER_URL")
        or os.getenv("WA_API_URL")
        or os.getenv("URL")
        or ""
    ).strip().rstrip("/")
    token = (os.getenv('WHATSAPP_AGENT_TOKEN') or '').strip()
    if not url or not url.startswith('https://') or not token:
        print('Missing WA_API_URL (https://...) or WHATSAPP_AGENT_TOKEN.', file=sys.stderr)
        return 2

    if not args.execute:
        print('CHECK ONLY: credentials configured; no request sent. Use --execute only after confirming 22:00 auto-publish is OFF.', flush=True)
        return 0

    endpoint = url + '/schedule/api/whatsapp-agent/roster-2200/tick'
    try:
        response = requests.post(endpoint, json={}, headers={'Authorization': 'Bearer ' + token}, timeout=(10, 75))  # 连接10秒、云端处理最多75秒；Sender外层限时100秒
        print('HTTP:', response.status_code, flush=True)
        print('Result:', response.text[:3000], flush=True)
        response.raise_for_status()
        data = response.json()
        status = data.get('status')
        if status in ('queued', 'already_queued', 'published_no_message', 'disabled'):
            if status == 'disabled':
                print('22:00 automation disabled by settings.', flush=True)
            return 0
        # 其他结果一律不视为完成；由 Sender 在22:00窗口内安全重试。
        print('22:00 roster not confirmed; retry required. status=' + str(status),
              file=sys.stderr, flush=True)
        return 1
    except (requests.RequestException, ValueError) as exc:
        print('Request failed:', str(exc), file=sys.stderr)
        return 1


if __name__ == '__main__':
    sys.exit(main())
