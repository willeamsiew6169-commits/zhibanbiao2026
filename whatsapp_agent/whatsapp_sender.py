"""观音堂 WhatsApp Web 发送端 V9：人工模式 + 测试群无人值守模式。

硬性限制：Render API 只领取 message_type=test，且本程序只接受
TEST_GROUP_ONLY。不能发送正式群。发送动作不自动重试。
注意：DOM 中出现自己发出的消息不等于 WhatsApp 服务器或群成员已送达。
"""
import logging
import atexit
import msvcrt
import os
import time
import subprocess
import sys
from datetime import datetime
from zoneinfo import ZoneInfo
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
# Local setting is an additional safety gate; webpage controls runtime mode.
LOCAL_AUTO_ALLOWED = os.getenv('WA_AUTO_SEND', 'false').strip().lower() == 'true'
AUTO = False
AUTO_ACK = os.getenv('WA_TEST_AUTO_ACKNOWLEDGED', '').strip() == 'YES'
HEADERS = {'Authorization': 'Bearer ' + TOKEN}
logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
LOG = logging.getLogger('whatsapp_sender')


def post(path, data):
    # 连接超时 10 秒；读取超时 90 秒，以容忍 Render 免费实例唤醒。
    # 不在这里自动重试 POST：尤其 /ack 不可盲目重复提交。
    response = requests.post(URL + path, json=data, headers=HEADERS, timeout=(10, 90))
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
    """按 WhatsApp 的 message-out 容器识别本人消息；严格比较正文和新增 ID。"""
    return page.evaluate(r"""(expected) => {
        const main = document.querySelector('#main');
        if (!main) throw new Error('找不到聊天主区域 #main');
        const canonical = s => (s || '')
            .replace(/[\s\u00a0\u200b\u200c\u200d\ufeff\ufe0e\ufe0f]/g, '');
        // WhatsApp Web 的 selectable-text 常把 emoji / 项目符号渲染成独立节点，
        // innerText / textContent 会丢掉它们；仅在严格匹配失败后允许忽略这些装饰符号。
        // 姓名、日期、时间、地点及所有普通文字仍必须完整一致。
        const contentCanonical = s => canonical(s)
            .replace(/\p{Extended_Pictographic}/gu, '')
            .replace(/[•▪●◦]/g, '')
            .replace(/[\u20e3\u200d]/g, '');
        const want = canonical(expected);
        const wantContent = contentCanonical(expected);
        const matched = [], samples = [], ids = [];
        let total = 0;
        // 新版 WhatsApp Web 的本人消息通常使用 .message-out。
        // 旧版仅用 aria-label="你："，可能漏掉刚发送的新消息。
        for (const node of main.querySelectorAll('[data-id]')) {
            const container = node.matches('[data-testid="msg-container"]')
                ? node : node.querySelector('[data-testid="msg-container"]');
            if (!container) continue;
            const outgoing = node.classList.contains('message-out') ||
                !!node.closest('.message-out') ||
                !!node.querySelector('.message-out') ||
                container.classList.contains('message-out') ||
                !!node.querySelector('[aria-label="你："]');
            if (!outgoing) continue;
            const id = node.getAttribute('data-id') || '';
            if (!id || ids.includes(id)) continue;
            total++;
            ids.push(id);
            const all = Array.from(container.querySelectorAll('[data-testid="selectable-text"]'));
            const top = all.filter(el => !el.parentElement?.closest('[data-testid="selectable-text"]'));
            const variants = [];
            for (const el of top) {
                variants.push(el.innerText || '');
                variants.push(el.textContent || '');
                const copy = el.cloneNode(true);
                for (const img of copy.querySelectorAll('img[alt]')) {
                    img.replaceWith(document.createTextNode(img.getAttribute('alt') || ''));
                }
                variants.push(copy.textContent || '');
            }
            variants.push(top.map(el => el.innerText || '').join('\n'));
            variants.push(top.map(el => {
                const copy = el.cloneNode(true);
                for (const img of copy.querySelectorAll('img[alt]')) {
                    img.replaceWith(document.createTextNode(img.getAttribute('alt') || ''));
                }
                return copy.textContent || '';
            }).join('\n'));
            const exact = variants.some(body => canonical(body) === want);
            const contentExact = !exact && wantContent.length >= 30 &&
                variants.some(body => contentCanonical(body) === wantContent);
            if (exact || contentExact) matched.push({id, mode: exact ? 'exact' : 'text-without-decorative-emoji'});
            if (samples.length < 5) samples.push({
                id, bodies: [...new Set(variants)].filter(Boolean).slice(0, 3)
                    .map(x => x.slice(0, 160)),
                expectedLength: expected.length
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
        LOG.info('发送前历史聊天样本（非本次发送内容）=%s', before['samples'][-3:])
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
                         [(x['id'], x.get('mode')) for x in after['matched']][-3:])
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


def run_1900_trigger_if_due(last_completed_date):
    """19:00 MYT：使用现有云端接口入队；云端负责开关、时间及防重复。"""
    now = datetime.now(ZoneInfo('Asia/Kuala_Lumpur'))
    today = now.date().isoformat()
    # 当前后台默认 19:00；如在后台修改通知时间，需同步调整这里的触发时间。
    if not (now.hour == 19 and now.minute < 5) or last_completed_date == today:
        return last_completed_date
    try:
        result = post('/schedule/api/whatsapp-agent/notice-1900/tick', {})
        LOG.info('19:00 报名通知 tick：%s', result)
        if result.get('status') in ('queued', 'already_queued', 'disabled'):
            return today
        # 云端若返回 outside_window，不记完成；仍仅在本地 19:00-19:04 尝试。
    except Exception:
        LOG.exception('19:00 通知触发失败；稍后在时间窗口内重试（云端有防重复）')
    return last_completed_date


def run_2200_trigger_if_due(last_attempt_date):
    """在 Sender 主循环中触发既有 tick；不重复实现云端排班逻辑。"""
    now = datetime.now(ZoneInfo('Asia/Kuala_Lumpur'))
    today = now.date().isoformat()
    if not (now.hour == 22 and now.minute < 20) or last_attempt_date == today:
        return last_attempt_date
    tick = BASE / 'whatsapp_roster_2200_tick.py'
    if not tick.is_file():
        LOG.error('22:00 tick 文件不存在：%s；本次不执行', tick)
        return today
    LOG.info('22:00 MYT：调用既有排班发布 tick（只启动一个 Sender）')
    try:
        result = subprocess.run(
            [sys.executable, str(tick), '--execute'],
            cwd=str(BASE), capture_output=True, text=True, timeout=100,
            check=False,
        )
        LOG.info('22:00 tick 返回码=%s；输出=%s', result.returncode,
                 (result.stdout or '').strip()[:2000])
        if result.stderr:
            LOG.warning('22:00 tick 错误输出=%s', result.stderr.strip()[:2000])
        # 仅在明确成功/已入队时标记完成；失败时窗口内可安全重试。
        # 云端使用事务锁、固定 created_by 和不可覆盖快照防止重复。
        import json
        output = (result.stdout or '').strip()
        for line in reversed(output.splitlines()):
            if line.startswith('Result:'):
                try:
                    data = json.loads(line.partition('Result:')[2].strip())
                except (ValueError, TypeError):
                    data = {}
                if result.returncode == 0 and data.get('status') in (
                        'queued', 'already_queued', 'published_no_message', 'disabled'):
                    return today
                break
    except Exception:
        LOG.exception('22:00 tick 调用异常；在窗口内尝试重试')
    return last_attempt_date


def run_morning_trigger_if_due(last_completed_slot):
    """Ask cloud at half-hour slots from 05:30–18:30 only.

    Cloud checks whether today is a special day (05:30) or normal day (07:00).
    Never trigger after the 22:00 roster; overnight additions wait until morning.
    """
    now = datetime.now(ZoneInfo('Asia/Kuala_Lumpur'))
    today = now.date().isoformat()
    if not (5 <= now.hour < 19 and (now.minute < 20 or 30 <= now.minute < 50)):
        return last_completed_slot
    slot = f'{now.hour:02d}:{"00" if now.minute < 20 else "30"}'
    key = (today, slot)
    if last_completed_slot == key:
        return last_completed_slot
    try:
        result = post('/schedule/api/whatsapp-agent/morning/tick', {})
        LOG.info('早上首次及其后每30分钟新增义工通知 tick：%s', result)
        if result.get('status') in ('queued', 'already_queued', 'no_new_assignments',
                                    'legacy_job_exists', 'disabled', 'outside_window'):
            return key
        # Missing 22:00 snapshot may become available; retry within the slot.
    except Exception:
        LOG.exception('新增义工通知触发失败；在本时段窗口内重试')
    return last_completed_slot


def main():
    _lock_handle = acquire_single_instance_lock()
    if not URL.startswith('https://') or not TOKEN or not GROUP:
        raise SystemExit('请配置 HTTPS WA_SERVER_URL、WHATSAPP_AGENT_TOKEN、WA_TEST_GROUP')
    if LOCAL_AUTO_ALLOWED and not AUTO_ACK:
        LOG.warning('未设置 WA_TEST_AUTO_ACKNOWLEDGED=YES；网页自动模式将保持关闭')
    print('仅领取 test / TEST_GROUP_ONLY 消息；正式群发送未启用。')
    print('模式：由网页开关控制；本机 WA_AUTO_SEND=true 与 WA_TEST_AUTO_ACKNOWLEDGED=YES 才允许自动发送')
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
            # 登录后自动等待 WhatsApp Web 就绪，不再需要人工按 Enter。
            # 只在聊天列表出现后才领取任务；人工 SEND + CONFIRM 仍保留。
            LOG.info('等待 WhatsApp Web 登录及聊天列表就绪（最多 5 分钟）...')
            login_deadline = time.monotonic() + 300
            while True:
                try:
                    ready_page(context)
                    LOG.info('WhatsApp Web 已就绪，开始检查测试消息队列')
                    break
                except (PlaywrightError, RuntimeError) as exc:
                    if time.monotonic() >= login_deadline:
                        raise RuntimeError('等待 WhatsApp Web 就绪超时；请检查登录状态') from exc
                    time.sleep(3)
            last_morning_completed_date = None
            last_1900_completed_date = None
            last_2200_attempt_date = None
            while True:
                last_morning_completed_date = run_morning_trigger_if_due(last_morning_completed_date)
                last_1900_completed_date = run_1900_trigger_if_due(last_1900_completed_date)
                last_2200_attempt_date = run_2200_trigger_if_due(last_2200_attempt_date)
                try:
                    page = ready_page(context)
                except (PlaywrightError, RuntimeError) as exc:
                    LOG.error('浏览器检查失败，暂停领取：%s', exc)
                    break
                # Fetch the webpage switch before claiming each job. Fail closed if unreachable.
                global AUTO
                try:
                    config = post('/schedule/api/whatsapp-agent/config', {})
                    if not isinstance(config.get('auto_send'), bool):
                        raise ValueError('Agent config missing auto_send boolean')
                    AUTO = bool(config['auto_send']) and LOCAL_AUTO_ALLOWED and AUTO_ACK
                except Exception:
                    LOG.exception('无法读取网页自动发送开关；暂停领取，避免模式不明时发送')
                    time.sleep(POLL)
                    continue
                try:
                    job = post('/schedule/api/whatsapp-agent/claim', {'agent_id': AGENT}).get('job')
                except requests.RequestException:
                    LOG.warning('Render 连接超时或网络异常；%s 秒后重新领取（未执行 WhatsApp 发送）', POLL, exc_info=True)
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
