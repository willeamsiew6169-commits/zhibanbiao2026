"""Windows test agent: claims TEST jobs only. WhatsApp Web automation is unofficial.
First run is manual confirmation; WA_AUTO_SEND=true requires explicit opt-in.
"""
import os
import time
import logging
from pathlib import Path
from dotenv import load_dotenv
import requests
from playwright.sync_api import sync_playwright

BASE = Path(__file__).resolve().parent
load_dotenv(BASE / '.env.agent')
URL = os.getenv('WA_SERVER_URL', '').rstrip('/')
TOKEN = os.getenv('WHATSAPP_AGENT_TOKEN', '')
GROUP = os.getenv('WA_TEST_GROUP', '')
AGENT = os.getenv('WA_AGENT_ID', 'windows-test-pc')
POLL = max(5, int(os.getenv('WA_POLL_SECONDS', '15')))
AUTO = os.getenv('WA_AUTO_SEND', 'false').lower() == 'true'
HEADERS = {'Authorization': 'Bearer ' + TOKEN}
logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')


def post(path, data):
    r = requests.post(URL + path, json=data, headers=HEADERS, timeout=25)
    r.raise_for_status()
    return r.json()


def open_test_chat(page):
    # Chat search UI changes frequently. Fail closed if not exactly one match.
    search = page.get_by_role('textbox', name='Search input textbox')
    if search.count() != 1:
        raise RuntimeError('找不到 WhatsApp 搜索框，请更新选择器')
    search.fill(GROUP)
    page.wait_for_timeout(1800)
    target = page.get_by_text(GROUP, exact=True)
    if target.count() != 1:
        raise RuntimeError(f'测试群匹配数量={target.count()}，为安全起见停止')
    target.click()
    page.wait_for_timeout(700)
    # Verify selected conversation header is the exact group.
    header = page.locator('header').get_by_text(GROUP, exact=True)
    if header.count() != 1:
        raise RuntimeError('当前聊天不是指定测试群，停止发送')


def send_test(page, message):
    open_test_chat(page)
    editor = page.locator('footer [contenteditable="true"][role="textbox"]')
    if editor.count() != 1:
        raise RuntimeError('找不到唯一的消息输入框，停止发送')
    editor.fill(message)
    if not AUTO:
        print('\n消息已填入测试群输入框。请检查群组及内容。')
        answer = input('确认发送？输入 SEND 后回车；其他输入取消：').strip()
        if answer != 'SEND':
            editor.fill('')
            return 'failed', '操作员取消；没有发送'
    # Once Enter is attempted, delivery may be ambiguous. Never automatically retry.
    try:
        editor.press('Enter')
        page.wait_for_timeout(1500)
        return 'uncertain', '已尝试按 Enter；请人工确认测试群收到后处理为 sent。'
    except Exception as exc:
        return 'uncertain', f'发送动作可能已执行：{type(exc).__name__}: {exc}'


def main():
    if not URL.startswith('https://') or not TOKEN or not GROUP:
        raise SystemExit('请先设置 HTTPS WA_SERVER_URL、WHATSAPP_AGENT_TOKEN、WA_TEST_GROUP')
    print('仅领取 message_type=test 的任务；不会发送真实义工通知。')
    print('自动发送模式：', AUTO, '；目标测试群：', GROUP)
    with sync_playwright() as pw:
        browser = pw.chromium.launch_persistent_context(str(BASE / 'wa_browser_profile'), headless=False)
        page = browser.pages[0] if browser.pages else browser.new_page()
        page.goto('https://web.whatsapp.com/')
        input('请先扫码登录 WhatsApp Web，登录成功后回到这里按 Enter 开始轮询：')
        try:
            while True:
                try:
                    job = post('/schedule/api/whatsapp-agent/claim', {'agent_id': AGENT}).get('job')
                    if job:
                        logging.info('Claimed job %s', job['id'])
                        if job.get('destination') != 'TEST_GROUP_ONLY':
                            status, err = 'failed', '非测试目标，已拒绝'
                        else:
                            try:
                                status, err = send_test(page, job['body'])
                            except Exception as exc:
                                status, err = 'uncertain', str(exc)
                        post('/schedule/api/whatsapp-agent/ack', {
                            'id': job['id'], 'agent_id': AGENT,
                            'status': status, 'error': err
                        })
                        logging.info('Job %s -> %s (%s)', job['id'], status, err)
                except KeyboardInterrupt:
                    break
                except Exception:
                    logging.exception('Poll failed; will retry polling, not resend claimed jobs')
                time.sleep(POLL)
        finally:
            browser.close()

if __name__ == '__main__':
    main()
