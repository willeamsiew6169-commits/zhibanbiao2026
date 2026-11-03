"""Windows Task Scheduler: run daily at 19:02 Malaysia time.
Uses existing whatsapp_agent/.env.agent; does not send WhatsApp itself.
"""
import os
import json
import urllib.request
from pathlib import Path


def main():
    env_file = Path(__file__).resolve().parent / '.env.agent'
    if env_file.exists():
        for raw in env_file.read_text(encoding='utf-8-sig').splitlines():
            line = raw.strip()
            if not line or line.startswith('#') or '=' not in line:
                continue
            key, value = line.split('=', 1)
            os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))
    base = os.environ.get('WA_SERVER_URL', '').rstrip('/')
    token = os.environ.get('WHATSAPP_AGENT_TOKEN', '')
    if not base or not token:
        raise SystemExit('Missing WA_SERVER_URL or WHATSAPP_AGENT_TOKEN in .env.agent')
    url = base + '/schedule/api/whatsapp-agent/notice-1900/tick'
    req = urllib.request.Request(url, data=b'{}', method='POST', headers={
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json',
    })
    with urllib.request.urlopen(req, timeout=45) as response:
        print(json.dumps(json.load(response), ensure_ascii=False))


if __name__ == '__main__':
    main()
