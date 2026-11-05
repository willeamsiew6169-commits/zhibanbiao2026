"""Windows Task Scheduler: run every minute; server checks MYT 22:00 window."""
import os
from pathlib import Path
import requests
from dotenv import load_dotenv

load_dotenv(Path(__file__).with_name('.env.agent'))
url = os.getenv('WA_SERVER_URL', '').rstrip('/')
token = os.getenv('WHATSAPP_AGENT_TOKEN', '')
if not url or not token:
    raise RuntimeError('Missing WA_SERVER_URL / WHATSAPP_AGENT_TOKEN in .env.agent')
r = requests.post(url + '/schedule/api/whatsapp-agent/roster-2200/tick',
                  headers={'Authorization': 'Bearer ' + token}, json={}, timeout=90)
r.raise_for_status()
print(r.json())
