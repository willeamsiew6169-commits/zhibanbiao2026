"""离线综合测试：与 whatsapp_sender.py 放同一目录，python test_whatsapp_popup.py。
使用本机 Chromium 渲染模拟 WhatsApp DOM；所有浏览器网络请求均拦截。
不读取 .env.agent、不运行 main、不访问 Render / Supabase / WhatsApp。
"""
import importlib.util
import shutil
import subprocess  # 在非 Windows 环境先载入 subprocess，再临时模拟 msvcrt
import sys
import types
import unittest
from pathlib import Path
from unittest.mock import patch, Mock
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

with patch('dotenv.load_dotenv', return_value=False):
    if sys.platform != 'win32':
        sys.modules['msvcrt'] = types.ModuleType('msvcrt')
    try:
        spec = importlib.util.spec_from_file_location('sender_under_test', Path(__file__).with_name('whatsapp_sender.py'))
        sender = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(sender)
    finally:
        if sys.platform != 'win32':
            sys.modules.pop('msvcrt', None)

sender.GROUP = '观音堂自动通知测试群'
sender.AUTO = True
HTML = '''<!doctype html><meta charset="utf-8"><style>
[role=dialog] {position:fixed;inset:0;z-index:100;background:#ddd;display:flex;gap:10px;align-items:center;justify-content:center}
#side, #main {padding:12px} [contenteditable] {min-height:24px;border:1px solid #999} span[title] {display:inline-block;padding:6px}
</style><div id="side"><div contenteditable="true" data-tab="1"></div>
<span title="观音堂自动通知测试群" onclick="window.groupClicks++; document.querySelector('#main header span').title=this.title">观音堂自动通知测试群</span></div>
<div id="main"><header><span title="其他群">群标题</span></header></div>
<footer><div contenteditable="true" role="textbox" data-lexical-editor="true"></div></footer>
<script>
window.groupClicks=0; window.sent=0; window.buttonClicks=[];
window.popup = (label='接收 WhatsApp 通知', buttons=['开启通知','暂不'], closes=true) => {
 const d=document.createElement('div'); d.setAttribute('role','dialog'); d.setAttribute('aria-modal','true'); d.setAttribute('aria-label',label);
 for(const label of buttons){const b=document.createElement('button'); b.textContent=label; b.onclick=()=>{window.buttonClicks.push(label); if(closes)d.remove()};d.append(b)}
 document.body.append(d);
};
document.querySelector('footer div').addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();window.sent++; const n=document.createElement('div'); n.className='message-out'; n.dataset.id='msg'+window.sent; const t=document.createElement('span'); t.className='selectable-text'; t.textContent=e.target.innerText; n.append(t); document.querySelector('#main').append(n);e.target.innerText=''}});
</script>'''

class PopupTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.pw = sync_playwright().start()
        cls.addClassCleanup(cls.pw.stop)
        args = {'headless': True}
        if not Path(cls.pw.chromium.executable_path).exists():
            local = shutil.which('chromium') or shutil.which('chromium-browser') or shutil.which('google-chrome')
            if local:
                args['executable_path'] = local
        cls.browser = cls.pw.chromium.launch(**args)
        cls.addClassCleanup(cls.browser.close)

    def setUp(self):
        self.ctx = self.browser.new_context(service_workers='block')
        self.addCleanup(self.ctx.close)
        self.ctx.route('**/*', lambda route: route.abort())
        self.page = self.ctx.new_page()
        self.page.set_content(HTML)
        # 模拟已登录的 WhatsApp URL；真实 DOM 在 about:blank，不导航到公网。
        self.ready_proxy = Mock(wraps=self.page)
        self.ready_proxy.url = 'https://web.whatsapp.com/'
        self.ready_context = types.SimpleNamespace(pages=[self.ready_proxy])
        self.page.set_default_timeout(800)

    def popup(self, label='接收 WhatsApp 通知', buttons=None, closes=True):
        self.page.evaluate('([label,buttons,closes])=>popup(label,buttons,closes)', [label, ['开启通知','暂不'] if buttons is None else buttons, closes])

    def clicks(self):
        return self.page.evaluate('window.buttonClicks')

    def test_01_no_popup(self):
        self.assertFalse(sender.dismiss_notification_popup(self.page))
        self.assertEqual(self.clicks(), [])

    def test_02_chinese_preclaim(self):
        self.popup()
        self.assertIs(sender.ready_page(self.ready_context), self.ready_proxy)
        self.assertEqual(self.clicks(), ['暂不'])

    def test_03_english(self):
        self.popup('Get WhatsApp notifications', ['Turn on','Not now'])
        self.assertTrue(sender.dismiss_notification_popup(self.page))
        self.assertEqual(self.clicks(), ['Not now'])

    def test_04_prefer_decline_to_close(self):
        self.popup(buttons=['关闭','开启通知','暂不接收'])
        self.assertTrue(sender.dismiss_notification_popup(self.page))
        self.assertEqual(self.clicks(), ['暂不接收'])

    def test_05_close_only(self):
        self.popup(buttons=['开启通知','关闭'])
        self.assertTrue(sender.dismiss_notification_popup(self.page))
        self.assertEqual(self.clicks(), ['关闭'])

    def test_06_unknown_dialog(self):
        self.popup('确认删除聊天', ['取消','删除'])
        with self.assertRaises(RuntimeError): sender.ready_page(self.ready_context)
        self.assertEqual(self.clicks(), [])

    def test_07_unknown_buttons(self):
        self.popup(buttons=['开启通知','继续'])
        with self.assertRaises(RuntimeError): sender.dismiss_notification_popup(self.page)
        self.assertEqual(self.clicks(), [])

    def test_08_multiple_dialogs(self):
        self.popup(); self.popup('确认删除聊天')
        with self.assertRaises(RuntimeError): sender.dismiss_notification_popup(self.page)
        self.assertEqual(self.clicks(), [])

    def test_09_multiple_declines(self):
        self.popup(buttons=['暂不','稍后'])
        with self.assertRaises(RuntimeError): sender.dismiss_notification_popup(self.page)
        self.assertEqual(self.clicks(), [])

    def test_10_does_not_close(self):
        self.popup(closes=False)
        with self.assertRaises(PlaywrightTimeoutError): sender.dismiss_notification_popup(self.page)
        self.assertEqual(self.clicks(), ['暂不'])

    def test_11_hidden_dialog(self):
        self.popup('未知弹窗'); self.page.locator('[role=dialog]').evaluate("el=>el.style.display='none'")
        self.assertFalse(sender.dismiss_notification_popup(self.page))

    def test_12_late_popup_real_click_retry(self):
        calls=[]
        def action():
            calls.append(1)
            if len(calls)==1: self.popup()
            self.page.locator('#side span[title]').click()
        sender.pre_send_ui_action(self.page, action)
        self.assertEqual(len(calls),2)
        self.assertEqual(self.clicks(), ['暂不'])
        self.assertEqual(self.page.evaluate('window.groupClicks'),1)

    def test_13_unrelated_timeout_not_retried(self):
        calls=[]
        def action():
            calls.append(1); self.page.locator('#missing').click()
        with self.assertRaises(PlaywrightTimeoutError): sender.pre_send_ui_action(self.page, action)
        self.assertEqual(len(calls),1)

    def test_14_retry_only_once(self):
        calls=[]
        def action():
            calls.append(1); self.popup(); self.page.locator('#side span[title]').click()
        with self.assertRaises(PlaywrightTimeoutError): sender.pre_send_ui_action(self.page, action)
        self.assertEqual(len(calls),2)
        self.assertEqual(self.page.evaluate('window.groupClicks'),0)

    def test_15_normal_send_with_popup(self):
        self.popup()
        status, _ = sender.send_test(self.page,'离线测试信息')
        self.assertEqual(status,'sent')
        self.assertEqual(self.page.evaluate('window.sent'),1)
        self.assertEqual(self.clicks(), ['暂不'])

    def test_16_after_enter_no_retry(self):
        with patch.object(sender,'verify_after_send',side_effect=PlaywrightTimeoutError('模拟核对异常')):
            status,_=sender.send_test(self.page,'离线测试信息')
        self.assertEqual(status,'uncertain')
        self.assertEqual(self.page.evaluate('window.sent'),1)

    def test_17_duplicate_group_stops(self):
        self.page.locator('#side span').evaluate('el=>el.after(el.cloneNode(true))')
        with self.assertRaises(RuntimeError): sender.send_test(self.page,'离线测试信息')
        self.assertEqual(self.page.evaluate('window.sent'),0)

    def test_18_existing_draft_protected(self):
        self.page.locator('footer div').fill('原有草稿')
        status,_=sender.send_test(self.page,'离线测试信息')
        self.assertEqual(status,'failed')
        self.assertEqual(self.page.locator('footer div').inner_text(),'原有草稿')
        self.assertEqual(self.page.evaluate('window.sent'),0)

if __name__=='__main__':
    unittest.main(verbosity=2)
