"""观音堂 WhatsApp Web 发送端 V7：人工模式 + 测试群无人值守模式。

硬性限制：Render API 只领取 message_type=test，且本程序只接受
TEST_GROUP_ONLY。不能发送正式群。发送动作不自动重试。
注意：DOM 中出现自己发出的消息不等于 WhatsApp 服务器或群成员已送达。
"""
import logging
import atexit
import msvcrt
import os
import time
from pathlib import Path

import requests
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, Error as PlaywrightError

BASE = Path(__file__).resolve().parent
load_dotenv(BASE / '.env.agent')
URL = os.getenv('WA_SERVER_URL', '').rstrip('/')
TOKEN = os.getenv('WHATSAPP_AGENT_TOKEN', '').strip()
GROUP = os.getenv('WA_TEST_GROUP', '').strip()
AGENT = os.getenv('WA_AGENT_ID', 'windows-test-pc').strip()[:80]
POLL = max(5, int(os.getenv('WA_POLL_SECONDS', '15')))
AUTO = os.getenv('WA_AUTO_SEND', 'false').strip().lower() == 'true'
AUTO_ACK = os.getenv('WA_TEST_AUTO_ACKNOWLEDGED', '').strip() == 'YES'
HEADERS = {'Authorization': 'Bearer ' + TOKEN}
logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
LOG = logging.getLogger('whatsapp_sender')


def post(path, data):
    response = requests.post(URL + path, json=data, headers=HEADERS, timeout=25)
    response.raise_for_status()
    return response.json()


def ready_page(context):
    pages = [p for p in context.pages if not p.is_closed()
             and p.url.startswith('https://web.whatsapp.com/')]
    if len(pages) != 1:
        raise RuntimeError(f'WhatsApp Web 标签页数量={len(pages)}，要求恰好一个')
    page = pages[0]
    if page.evaluate('() => document.readyState') not in ('interactive', 'complete'):
        raise RuntimeError('WhatsApp Web 未加载完成')
    side = page.locator('#side')
    if side.count() != 1 or not side.is_visible():
        raise RuntimeError('WhatsApp Web 未登录或聊天列表未就绪')
    return page


def check_header(page):
    header = page.locator('header').get_by_text(GROUP, exact=True)
    if header.count() != 1 or not header.is_visible():
        raise RuntimeError('聊天标题不是指定测试群，禁止发送')


def open_test_chat(page):
    page.wait_for_selector('#side', timeout=30000)
    selectors = [
        '#side [contenteditable="true"][data-tab]',
        '#side [contenteditable="true"]',
        '[aria-label="搜索或开始新聊天"]',
        '[aria-label="Search input textbox"]',
    ]
    search = None
    for selector in selectors:
        candidate = page.locator(selector)
        count = candidate.count()
        LOG.info('搜索框检查 %s → %s', selector, count)
        if count == 1 and candidate.is_visible():
            search = candidate
            break
    if search is None:
        raise RuntimeError('找不到 WhatsApp 搜索框')
    search.fill(GROUP)
    page.wait_for_timeout(2000)
    matches = page.locator('#side span[title]').filter(has_text=GROUP)
    exact = [matches.nth(i) for i in range(matches.count())
             if matches.nth(i).get_attribute('title') == GROUP
             and matches.nth(i).is_visible()]
    LOG.info('测试群标题匹配数量=%s', len(exact))
    if len(exact) != 1:
        raise RuntimeError(f'测试群标题匹配数量={len(exact)}，停止发送')
    exact[0].click()
    page.wait_for_timeout(1000)
    check_header(page)


def outgoing_snapshot(page, message):
    """读取本人消息 ID 和正文；同时记录有限诊断样本，绝不把旧消息当新消息。"""
    return page.evaluate(r"""(expected) => {
        const main = document.querySelector('#main');
        if (!main) throw new Error('找不到聊天主区域 #main');
        const normalize = s => (s || '')
            .replace(/\u00a0/g, ' ').replace(/\u200b|\u200c|\u200d|\ufeff/g, '')
            .replace(/\r\n?/g, '\n').replace(/[ \t]+/g, ' ')
            .replace(/[ \t]*\n[ \t]*/g, '\n').trim();
        const want = normalize(expected);
        const matched = [], samples = [], ids = [];
        let total = 0;
        for (const node of main.querySelectorAll('[data-id]')) {
            if (!node.querySelector('[aria-label="你："]')) continue;
            const container = node.querySelector('[data-testid="msg-container"]');
            if (!container) continue;
            const id = node.getAttribute('data-id') || '';
            if (!id) continue;
            total++;
            ids.push(id);
            // WhatsApp 常将正文拆成多个 selectable-text；也可能嵌套。
            // 只选最外层，避免相同文字被拼接两次。
            const all = Array.from(container.querySelectorAll('[data-testid="selectable-text"]'));
            const top = all.filter(el => !el.parentElement?.closest('[data-testid="selectable-text"]'));
            const variants = [];
            for (const el of top) {
                // innerText 与 textContent 在 WhatsApp 的 emoji/隐藏节点上可能不同。
                variants.push(normalize(el.innerText || ''));
                variants.push(normalize(el.textContent || ''));
                // emoji 图片的 alt 是真实消息内容的一部分。
                const copy = el.cloneNode(true);
                for (const img of copy.querySelectorAll('img[alt]')) {
                    img.replaceWith(document.createTextNode(img.getAttribute('alt') || ''));
                }
                variants.push(normalize(copy.textContent || ''));
            }
            const combined = normalize(top.map(el => el.innerText || '').join('\n'));
            variants.push(combined);
            // 只允许整条正文完全相符，绝不做模糊包含判断。
            if (variants.some(body => body === want)) matched.push({id});
            if (samples.length < 5) samples.push({
                id, bodies: [...new Set(variants)].filter(Boolean).slice(0, 4)
                    .map(x => x.slice(0, 160)),
                expectedLength: want.length
            });
        }
        return {total, ids, matched, samples};
    }""", message)


def newly_visible_outgoing(before, after):
    """只有发送前不存在的本人消息 ID 且正文完整相符，才能确认。"""
    old_ids = set(before['ids'])
    return any(item['id'] not in old_ids for item in after['matched'])


def send_test(page, message):
    """返回 status, reason；按 Enter 后异常一律 uncertain。"""
    if not isinstance(message, str) or not message.strip():
        return 'failed', '空消息，拒绝发送'
    open_test_chat(page)
    editor = page.locator('footer [contenteditable="true"][role="textbox"]')
    if editor.count() != 1 or not editor.is_visible():
        raise RuntimeError('找不到唯一消息输入框')

    # 自动模式下先确保输入框为空，避免发送别人的草稿。
    if AUTO and (editor.inner_text().strip() or editor.get_attribute('data-lexical-editor') is None and editor.text_content().strip()):
        return 'failed', '检测到未发送草稿；自动模式拒绝覆盖或发送'

    before = outgoing_snapshot(page, message) if AUTO else None
    if AUTO:
        LOG.info('发送前：聊天区本人消息=%s，完全匹配=%s，匹配ID=%s',
                 before['total'], len(before['matched']),
                 [x['id'] for x in before['matched']][-3:])
    editor.fill(message)
    if AUTO:
        LOG.info('发送前正文样本=%s', before['samples'][-3:])
    if not AUTO:
        print('\n消息已填入测试群，请核对群名和内容。')
        answer = input('确认发送？输入 SEND 后回车；其他输入取消：').strip()
        if answer != 'SEND':
            try:
                editor.fill('')
            except Exception:
                LOG.warning('无法清空草稿，请人工检查')
            return 'failed', '操作员取消；未执行发送'
    try:
        check_header(page)
    except Exception as exc:
        return 'failed', f'发送前群组校验失败：{exc}'

    # 此处起可能已发出：任何异常均不可重试。
    try:
        editor.press('Enter')
        if AUTO:
            deadline = time.monotonic() + 12
            while time.monotonic() < deadline:
                page.wait_for_timeout(500)
                check_header(page)
                after = outgoing_snapshot(page, message)
                LOG.info('发送后检查：本人消息=%s，完全匹配=%s，匹配ID=%s',
                         after['total'], len(after['matched']),
                         [x['id'] for x in after['matched']][-3:])
                if newly_visible_outgoing(before, after):
                    return 'sent', '自动检测到测试群聊天区新增本人消息（仅本地界面确认，非送达回执）'
            LOG.warning('未匹配到新增本人消息；发送后正文样本=%s', after['samples'][-5:])
            return 'uncertain', '已按 Enter，但未能确认新增本人消息；禁止自动重试'
        page.wait_for_timeout(1500)
    except Exception as exc:
        return 'uncertain', f'发送动作结果不明：{type(exc).__name__}: {exc}'
    print('\n请人工检查测试群聊天记录是否出现刚才的完整消息。')
    answer = input('确认出现请输入 CONFIRM；否则直接回车：').strip()
    if answer == 'CONFIRM':
        return 'sent', '操作员确认消息出现在测试群聊天记录（非送达回执）'
    return 'uncertain', '已尝试按 Enter，但操作员未确认；禁止自动重试'


def acquire_single_instance_lock():
    """Windows 上同一资料夹只允许运行一个发送端。"""
    lock_path = BASE / '.whatsapp_sender.lock'
    handle = open(lock_path, 'a+b')
    try:
        handle.seek(0)
        if not handle.read(1):
            handle.seek(0)
            handle.write(b'1')
            handle.flush()
        handle.seek(0)
        msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
    except OSError:
        handle.close()
        raise SystemExit('检测到另一个 WhatsApp 发送程序正在运行，请先关闭旧程序')
    atexit.register(handle.close)
    return handle


def main():
    _lock_handle = acquire_single_instance_lock()
    if not URL.startswith('https://') or not TOKEN or not GROUP:
        raise SystemExit('请配置 HTTPS WA_SERVER_URL、WHATSAPP_AGENT_TOKEN、WA_TEST_GROUP')
    if AUTO and not AUTO_ACK:
        raise SystemExit('自动测试需同时设置 WA_TEST_AUTO_ACKNOWLEDGED=YES；仅限测试群')
    print('仅领取 test / TEST_GROUP_ONLY 消息；正式群发送未启用。')
    print('模式：', '测试群自动发送' if AUTO else '人工 SEND + CONFIRM')
    print('测试群：', GROUP, '；轮询：', POLL, '秒')
    with sync_playwright() as pw:
        context = pw.chromium.launch_persistent_context(
            str(BASE / 'wa_browser_profile'), headless=False)
        try:
            if not context.pages:
                page = context.new_page()
                page.goto('https://web.whatsapp.com/', wait_until='domcontentloaded')
            elif not any(p.url.startswith('https://web.whatsapp.com/') for p in context.pages):
                context.pages[0].goto('https://web.whatsapp.com/', wait_until='domcontentloaded')
            input('确认 WhatsApp Web 已登录，回到 Terminal 按 Enter 开始：')
            while True:
                try:
                    page = ready_page(context)
                except (PlaywrightError, RuntimeError) as exc:
                    LOG.error('浏览器检查失败，暂停领取：%s', exc)
                    break
                try:
                    job = post('/schedule/api/whatsapp-agent/claim', {'agent_id': AGENT}).get('job')
                except requests.RequestException:
                    LOG.exception('Render 连接失败；等待后重试领取')
                    time.sleep(POLL)
                    continue
                if not job:
                    LOG.info('没有待发送的测试消息，%s 秒后再检查', POLL)
                    time.sleep(POLL)
                    continue
                LOG.info('Claimed job %s', job['id'])
                if job.get('destination') != 'TEST_GROUP_ONLY':
                    status, reason = 'failed', '目标非 TEST_GROUP_ONLY，拒绝发送'
                else:
                    try:
                        status, reason = send_test(page, job['body'])
                    except Exception as exc:
                        LOG.exception('准备阶段异常；停止自动发送并人工核对')
                        # 无法证明是否已经按 Enter，保守地记为 uncertain。
                        status, reason = 'uncertain', f'执行过程异常，人工核对：{type(exc).__name__}: {exc}'
                try:
                    post('/schedule/api/whatsapp-agent/ack', {
                        'id': job['id'], 'agent_id': AGENT,
                        'status': status, 'error': reason,
                    })
                    LOG.info('Job %s -> %s (%s)', job['id'], status, reason)
                except Exception:
                    LOG.exception('ACK 失败；请人工检查 Supabase，不要重排任务')
                    break
                if AUTO and status != 'sent':
                    LOG.error('自动测试结果非 sent，暂停继续领取，防止连续失败')
                    break
                time.sleep(POLL)
        except KeyboardInterrupt:
            LOG.info('人工停止')
        finally:
            try:
                context.close()
            except PlaywrightError:
                pass


if __name__ == '__main__':
    main()
