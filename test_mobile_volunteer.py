"""Offline browser and route checks. Run: python test_mobile_volunteer.py.

Requires Flask, Playwright, Chromium (playwright install chromium), and tzdata
on hosts without IANA time zones. Optionally set PLAYWRIGHT_CHROMIUM_EXECUTABLE.
The app is assembled from real templates and selected functions. No production
modules that connect to a database or send notifications are imported.
"""
import ast
import calendar
from contextlib import contextmanager
from datetime import date, datetime, timedelta, timezone
from functools import wraps
from http import HTTPStatus
import hashlib
import logging
import os
from pathlib import Path
import re
import secrets
import subprocess
import sys
import threading
import time
from types import ModuleType
from unittest.mock import patch
from zoneinfo import ZoneInfo

from flask import Blueprint, Flask, abort, flash, jsonify, redirect, render_template_string, request, session, url_for
from playwright.sync_api import expect, sync_playwright
from werkzeug.serving import make_server

from schedule.builders.time_utils import time_to_minutes
from schedule.constants import ROLE_TEXT
from schedule.volunteer_templates import VOLUNTEER_SIGNUP_HTML
from utils import get_lang, get_text


ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / 'output' / 'mobile-volunteer-tests'
TODAY = date(2026, 11, 25)
SHIFT_TIMES = ['9:00am', '10:00am', '11:00am', '12:00pm']
VOLUNTEER = {'编号': 'CHE-208', '姓名': '测试义工', '状态': '在册', 'PIN': '1234'}
SPECIAL_DAY = {'template_type': 'normal', 'lunar_text': '', 'special_names': []}


def load_parts(file, names, namespace):
    path = ROOT / file
    tree = ast.parse(path.read_text(encoding='utf-8-sig'), filename=str(path))
    nodes = []
    for node in tree.body:
        if isinstance(node, ast.FunctionDef) and node.name in names:
            node.decorator_list = []
            if node.name == 'register_phase3':
                # Keep the routes. Exclude registration of the unrelated reporting module.
                node.body = [child for child in node.body if isinstance(child, ast.FunctionDef)]
            nodes.append(node)
        if isinstance(node, ast.Assign) and any(isinstance(target, ast.Name) and target.id in names
                                                 for target in node.targets):
            nodes.append(node)
    assert len(nodes) == len(names), (file, names)
    exec(compile(ast.Module(body=nodes, type_ignores=[]), str(path), 'exec'), namespace)


class FakeDatabase:
    def __init__(self):
        self.reset()

    def reset(self):
        self.operations = []
        self.saved = []
        self.sync_failed = False
        self.mode = ''
        self.signup_date = TODAY + timedelta(days=1)
        self.close_calls = []

    @contextmanager
    def connect(self):
        self.pending = []
        yield self
        self.commit()

    def commit(self):
        self.saved.extend(self.pending)
        self.pending.clear()

    def cursor(self, **kwargs):
        return self

    def __enter__(self):
        return self

    def __exit__(self, *args):
        return False

    def execute(self, sql, args=()):
        sql = ' '.join(sql.lower().split())
        self.operations.append((sql, args))
        self.row = None
        if sql.startswith(('insert ', 'update ')):
            self.pending.append((sql, args))
        if sql.startswith('insert into volunteer_schedule_signups'):
            self.row = {'id': 41}
        if 'from volunteer_schedule_signups where id = %s for update' in sql:
            self.row = self.signup()
        if sql.startswith('select id,event_date,location_id,status,closed_at from hongfa_events'):
            self.row = {'id': 1, 'event_date': datetime.now(timezone(timedelta(hours=8))).date(),
                        'location_id': 1, 'status': 'published', 'closed_at': None}
        if sql.startswith('insert into hongfa_attendance'):
            self.row = {'id': 51, 'check_in_at': datetime(2026, 11, 25, 4, 17, 32, tzinfo=timezone.utc)}

    def fetchone(self):
        return self.row

    def fetchall(self):
        if self.operations[-1][0].startswith('select * from volunteer_schedule_signups'):
            return [] if any("status = 'cancelled'" in sql for sql, _ in self.saved) else [self.signup()]
        return []

    def signup(self):
        return {'id': 41, 'volunteer_id': 'CHE-208', 'name': VOLUNTEER['姓名'],
                'signup_date': self.signup_date, 'role': '值班', 'status': 'pending',
                'start_time': '10:00am', 'end_time': '11:00am'}

    def sync(self, *args, **kwargs):
        assert self.saved, 'Sync called before saving'
        if self.sync_failed:
            raise RuntimeError('Simulated schedule sync failure')
        return 0


def build_app(db):
    app = Flask(__name__, static_folder=str(ROOT / 'static'))
    app.secret_key = 'only-for-offline-fixtures'
    app.testing = True
    tz = ZoneInfo('Asia/Kuala_Lumpur')
    event = dict(id=1, event_date=datetime.now(tz).date(), location_id=1, status='published',
                 closed_at=None, team_id=1, team_name='测试组', location_name='测试地点')
    dependencies = {
        'request': request, 'render_template_string': render_template_string, 'url_for': url_for, 'jsonify': jsonify,
        'redirect': redirect, 'flash': flash, 'get_lang': get_lang, 'get_text': get_text,
        'session': session, 'abort': abort, 'secrets': secrets, 'hashlib': hashlib, 'TZ': tz, 'wraps': wraps,
        'date': date, 'datetime': datetime, 'timedelta': timedelta, 'timezone': timezone,
        'calendar': calendar, 'logging': logging, 'time': time, 'os': os, 'RealDictCursor': object,
        'TIME_OPTIONS': SHIFT_TIMES, 'time_to_minutes': time_to_minutes,
        'get_conn': db.connect, 'db_query': lambda sql, *args, **kwargs: event if 'SELECT e.*,t.name AS team_name' in sql else None,
        'get_special_day_info': lambda day: SPECIAL_DAY, 'get_daily_dharma': lambda: '',
        'now_date_str': lambda: TODAY.isoformat(), 'now_time_str': lambda: '11:42am',
        'malaysia_today': lambda: TODAY, 'is_schedule_published': lambda day: False,
        '_volunteer_identity': lambda: 'CHE-208',
        '_login_locked': lambda *args: 0,
        '_login_gate': lambda *args, **kwargs: None,
        'find_volunteer_by_keyword': lambda keyword: [{'id': 'CHE-208', 'name': VOLUNTEER['姓名']}],
        'find_duty_time_conflict': lambda *args: None,
        'signup_has_formal_assignment': lambda *args: False,
        'sync_schedule_after_signup_change': db.sync,
        'auto_signout_unfinished_today': lambda: None,
        'find_volunteer': lambda raw_id: VOLUNTEER if raw_id in ('208', 'CHE-208') else None,
        'verify_pin_for_volunteer': lambda vol, pin: pin == '1234',
        'load_today_assignments_for_volunteer': lambda *args: [],
        'get_eligible_signin_roles': lambda *args: ['卫生', '值班'],
        '_close_previous_roles': lambda cur, volunteer, role: db.close_calls.append((volunteer, role)),
        'get_today_stats': lambda: {'total': 0, 'open': 0, 'finished': 0},
        'valid_qr': lambda event, token: token == 'offline-qr',
    }
    load_parts('attendance_service.py', {'sign_in', 'role_label', 'ROLES', 'ROLE_TEXT',
                                        'same_role_contiguous_group', 'format_assignment_for_attendance'}, dependencies)
    app.extensions['offline_handlers'] = dependencies
    dependencies['ROLE_TEXT'] = ROLE_TEXT
    load_parts('app.py', {'PAGE', 'index', 'do_sign_in'}, dependencies)
    load_parts('schedule/volunteer_templates.py', {'VOLUNTEER_PREBOOK_HTML'}, dependencies)
    load_parts('schedule/routes/volunteer_routes.py',
               {'volunteer_signup', 'volunteer_cancel_signup', 'volunteer_my_schedule',
                'volunteer_prebook', 'volunteer_signup_login', '_volunteer_form_identity'},
               dependencies)
    load_parts('hongfa/routes.py', {'BASE', 'render'}, dependencies)
    load_parts('hongfa/phase3.py', {'register_phase3'}, dependencies)
    hongfa = Blueprint('hongfa', __name__, url_prefix='/hongfa')
    dependencies['register_phase3'](hongfa, lambda: {'id': 1, 'team_ids': [1]}, dependencies['render'],
                                   lambda: 'offline-csrf', lambda: request.form.get('csrf_token') == 'offline-csrf')
    app.register_blueprint(hongfa)

    @app.route('/volunteer')
    def daily():
        default = TODAY + timedelta(days=int(request.args.get('tomorrow', '0')))
        html = render_template_string(VOLUNTEER_SIGNUP_HTML, zh=get_lang() == 'zh', times=SHIFT_TIMES,
                                      signup_date=default.isoformat(), default_date=default.isoformat(),
                                      login_id='CHE-208', login_name=VOLUNTEER['姓名'], multi_day_signup_open=True)
        if request.args.get('nav_test'):
            html = html.replace('<body>', '<body><nav data-vol-bottom-nav style="position:fixed;bottom:0;'
                                'width:100%;height:56px;background:#fff;z-index:50">Test navigation</nav>', 1)
        return html

    @app.route('/volunteer/query_volunteer')
    def signup_lookup():
        return jsonify(ok=True, volunteer_id='CHE-208', name=VOLUNTEER['姓名'])

    @app.route('/volunteer/day_info')
    def day_info():
        return jsonify(ok=True, **SPECIAL_DAY)

    @app.route('/api/volunteer/<identifier>', methods=['POST'])
    def lookup(identifier):
        return jsonify(ok=True, volunteer={**VOLUNTEER, 'pin_ok': db.mode != 'lookup-invalid'}, attendance_action='signin',
                       today_assignments=[
                           dict(id=1, role='值班', state='pending', assigned_place='观音堂',
                                start_time='09:00', end_time='11:00'),
                           dict(id=2, role='值班', state='pending', assigned_place='活动中心',
                                start_time='11:00', end_time='12:00'),
                       ],
                       eligible_duty_assignment_ids=[1, 2], open_roles=['佛台'],
                       eligible_roles=[{'role': role, 'label': dependencies['role_label'](role)}
                                       for role in ('卫生', '值班')])

    def delayed(handler):
        def wrapped(**kwargs):
            if request.method == 'POST':
                if db.mode == 'delay':
                    time.sleep(.45)
                if db.mode == 'unknown':
                    response = app.make_response(handler(**kwargs))
                    assert db.saved and response.status_code < 400
                    return 'Simulated lost response after commit', 503
            return handler(**kwargs)
        return wrapped

    for path, name, methods in [
        ('/', 'index', ['GET']), ('/signin', 'do_sign_in', ['POST']),
        ('/volunteer/login', 'volunteer_signup_login', ['GET', 'POST']),
        ('/volunteer/signup', 'volunteer_signup', ['POST']),
        ('/volunteer/prebook', 'volunteer_prebook', ['GET', 'POST']),
        ('/volunteer/cancel/<int:signup_id>', 'volunteer_cancel_signup', ['POST']),
        ('/volunteer/my_schedule', 'volunteer_my_schedule', ['GET']),
    ]:
        endpoint = f'schedule.{name}' if path.startswith('/volunteer') else name
        app.add_url_rule(path, endpoint, delayed(dependencies[name]), methods=methods)

    for endpoint, path in [
        ('set_lang', '/language/<lang>'), ('change_pin_page', '/pin'), ('admin.admin_report', '/admin/report'),
        ('today_attendance_list', '/api/today-attendance-list'), ('today_attendance_stats', '/api/today-attendance-stats'),
    ]:
        app.add_url_rule(path, endpoint, lambda **kwargs: jsonify(rows=[], date=TODAY.isoformat()))
    app.after_request(dependencies['_volunteer_form_identity'])
    return app


def in_view(page, locator):
    expect(locator).to_be_visible()
    box = locator.bounding_box()
    viewport = page.evaluate('({height: visualViewport.height, width: visualViewport.width})')
    assert box['height'] >= 44, box
    assert box['x'] >= -1 and box['x'] + box['width'] <= viewport['width'] + 1, (box, viewport)
    assert box['y'] >= -1 and box['y'] + box['height'] <= viewport['height'] + 1, (box, viewport)


def run_browser(app, db):
    server = make_server('127.0.0.1', 0, app, threaded=True)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    origin = f'http://127.0.0.1:{server.server_port}'
    try:
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(executable_path=os.environ.get('PLAYWRIGHT_CHROMIUM_EXECUTABLE'))
            browser_cases(browser, origin, db)
            browser.close()
    finally:
        server.shutdown()
        server.server_close()
        thread.join()


def browser_cases(browser, origin, db):
    context = browser.new_context(is_mobile=True, has_touch=True, timezone_id='Asia/Kuala_Lumpur',
                                  viewport={'width': 390, 'height': 720})
    context.route('**/*', lambda route: route.continue_() if route.request.url.startswith(origin + '/') else route.abort())
    page = context.new_page()
    page.set_default_timeout(6000)
    page.clock.set_fixed_time(datetime(2026, 11, 25, 0, 1, tzinfo=timezone.utc))
    page.add_init_script("""window.submissionTimeouts = [];
        const nativeTimeout = window.setTimeout;
        window.setTimeout = (callback, delay, ...args) => {
            if (delay >= 30000) window.submissionTimeouts.push(delay);
            return nativeTimeout(callback, delay, ...args);
        };""")
    db.reset()
    page.goto(origin + '/volunteer/login')
    number = page.locator('input[name="number"]')
    pin_field = page.locator('input[name="pin"]')
    expect(number).not_to_be_focused()
    expect(number).to_have_attribute('autocomplete', 'off')
    expect(pin_field).to_have_attribute('autocomplete', 'off')
    expect(pin_field).to_have_attribute('type', 'password')
    number.fill('208')
    pin_field.fill('9999')
    page.get_by_role('button', name='进入义工报名').click()
    expect(page.get_by_role('alert')).to_contain_text('PIN 不正确')
    number.fill('208')
    pin_field.fill('1234')
    page.get_by_role('button', name='进入义工报名').click()
    expect(page.get_by_text('身份已验证，请选择报名方式并提交。')).to_be_visible()
    page.locator('[onclick="showDailySignup()"]').click()
    expect(page.locator('#signupDraftStatus')).to_have_text('尚未提交')
    assert not db.saved
    print('PASS | Shared signup login has no autofocus, accepts correction and does not submit signup')
    errors = []
    page.on('pageerror', lambda error: errors.append(str(error)))
    sizes = [(320, 568), (360, 640), (375, 667), (390, 720), (412, 820), (430, 932), (568, 320), (320, 290)]
    for width, height in sizes:
        page.set_viewport_size({'width': width, 'height': height})
        for path in ('/volunteer', '/volunteer?tomorrow=1', '/volunteer/prebook?year=2026&month=12'):
            page.goto(origin + path)
            if path.startswith('/volunteer?') or path == '/volunteer':
                page.locator('[onclick="showDailySignup()"]').click()
            dock = page.locator('.vol-action-dock')
            button = dock.locator('button:visible')
            in_view(page, button)
            page.evaluate('window.scrollTo(0, document.body.scrollHeight)')
            in_view(page, button)
            last_control = page.locator('form[data-vol-submit]').locator('input:visible, select:visible, a:visible').last
            assert last_control.bounding_box()['y'] + last_control.bounding_box()['height'] <= dock.bounding_box()['y'] + 1
            if path.startswith('/volunteer/prebook'):
                if width in (320, 430) and height > 300:
                    page.screenshot(path=OUTPUT / f'multi-{width}.png')
                continue
            if width in (320, 430) and height > 300 and '?' not in path:
                page.screenshot(path=OUTPUT / f'daily-{width}.png')
            page.locator('#start_time').select_option('10:00am')
            page.locator('#end_time').select_option('11:00am')
            button.click()
            modal = page.locator('#confirm_modal')
            expect(modal).to_be_visible()
            for control in modal.locator('.vol-dialog-actions button').all():
                in_view(page, control)
            expect(modal.get_by_role('heading')).to_be_in_viewport(ratio=1)
            if (width, height) == (320, 568) and '?' not in path:
                page.screenshot(path=OUTPUT / 'daily-confirm-320.png')
        print(f'PASS | {width}×{height} daily, tomorrow, multi-day and signup dialog')

    page.set_viewport_size({'width': 390, 'height': 520})
    page.goto(origin + '/volunteer?nav_test=1')
    page.locator('[onclick="showDailySignup()"]').click()
    dock = page.locator('.vol-action-dock')
    nav = page.locator('[data-vol-bottom-nav]')
    box = dock.bounding_box()
    assert abs(box['y'] + box['height'] - nav.bounding_box()['y']) < 2
    page.evaluate("""() => {
        const oldHeight = visualViewport.height;
        Object.defineProperty(visualViewport, 'height', {value: oldHeight - 210});
        visualViewport.dispatchEvent(new Event('resize'));
    }""")
    box = dock.bounding_box()
    assert abs(box['y'] + box['height'] - (nav.bounding_box()['y'] - 210)) < 2
    print('PASS | Bottom navigation clearance and simulated keyboard viewport')

    page.goto(origin + '/volunteer')
    page.locator('[onclick="showDailySignup()"]').click()
    page.locator('#start_time').select_option('10:00am')
    page.locator('#end_time').select_option('11:00am')
    db.reset()
    db.mode = 'delay'
    page.locator('.vol-action-dock button:visible').click()
    confirm = page.locator('#confirm_modal [onclick="submitSignupForm()"]')
    confirm.evaluate('(button) => { button.click(); button.click(); }')
    expect(page.locator('#signupSubmitStatus')).to_contain_text('正在处理')
    assert 90000 in page.evaluate('window.submissionTimeouts')
    assert 30000 not in page.evaluate('window.submissionTimeouts')
    expect(page.get_by_role('heading', name=re.compile('报名成功'))).to_be_visible()
    assert len([sql for sql, _ in db.saved if sql.startswith('insert into volunteer_schedule_signups')]) == 1
    print('PASS | Daily signup commit, delay and repeated click')

    db.reset()
    page.goto(origin + '/volunteer?tomorrow=1')
    page.locator('[onclick="showDailySignup()"]').click()
    page.locator('#start_time').select_option('10:00am')
    page.locator('#end_time').select_option('11:00am')
    page.route('**/volunteer/signup', lambda route: route.abort('internetdisconnected'))
    page.locator('.vol-action-dock button:visible').click()
    confirm.click()
    expect(page.locator('#signupSubmitStatus')).to_contain_text('暂时无法确认结果')
    expect(page.locator('#signupSubmitStatus a')).to_have_attribute('href', '/volunteer/my_schedule')
    expect(confirm).to_be_disabled()
    assert not db.saved
    page.unroute('**/volunteer/signup')
    print('PASS | Tomorrow signup connection drop, no replay and record link')

    db.reset()
    page.goto(origin + '/volunteer/prebook?year=2026&month=12')
    dates = page.locator('.day-card')
    dates.nth(2).click()
    dates.nth(3).click()
    page.locator('.vol-action-dock button').click()
    expect(page.locator('[data-vol-result]')).to_contain_text('新增 2')
    assert len(db.saved) == 2
    print('PASS | Multi-day calendar selection and two confirmed saves')

    db.reset()
    db.mode = 'lookup-invalid'
    page.goto(origin + '/')
    page.locator('#volunteer_id').fill('208')
    page.locator('#pin').click()
    page.locator('#pin').fill('9999')
    expect(page.locator('#personBox')).to_contain_text(re.compile('PIN.*(?:错|正确)'))
    expect(page.locator('#pin')).to_be_focused()
    expect(page.locator('#signinModal')).to_be_hidden()
    assert not db.saved
    print('PASS | PIN lookup failure leaves focus on PIN and does not open check-in')

    for language, pin, mode in [('zh', '1234', 'delay'), ('en', '9999', ''), ('en', '1234', 'unknown')]:
        context.add_cookies([{'name': 'lang', 'value': language, 'url': origin}])
        db.reset()
        db.mode = mode
        page.goto(origin + '/')
        page.locator('#volunteer_id').fill('208')
        page.locator('#pin').click()
        page.locator('#pin').fill(pin)
        expect(page.locator('.signin-modal-choice').first).to_be_visible()
        expect(page.locator('#pin')).to_have_attribute('autocomplete', 'off')
        expect(page.locator('#pin')).not_to_be_focused()
        expect(page.locator('#signinSubmitStatus')).to_contain_text(re.compile('身份已验证|Identity verified'))
        assert not db.saved
        if mode == 'delay':
            page.locator('.signin-modal-choice').filter(has_text='值班').click()
            expect(page.locator('#signinModalContent')).to_contain_text('观音堂')
            expect(page.locator('#signinModalContent')).to_contain_text('活动中心')
            page.locator('#signinConfirm').click()
            expect(page.locator('#signinSubmitStatus')).to_contain_text('请先填写义工卡号')
            assert not db.saved
            page.get_by_role('button', name='返回选择岗位').click()
        page.locator('.signin-modal-choice').filter(has_text=re.compile('卫生|Cleaning')).click()
        expect(page.locator('#signinModalContent')).to_contain_text(VOLUNTEER['姓名'])
        expect(page.locator('#signinSubmitStatus')).to_contain_text(re.compile('尚未提交|not submitted'))
        in_view(page, page.locator('#signinConfirm'))
        in_view(page, page.locator('.vol-dialog-actions .signin-modal-cancel'))
        if language == 'zh':
            page.screenshot(path=OUTPUT / 'signin-confirm-390.png')
        page.locator('#signinConfirm').evaluate('(button) => { button.click(); button.click(); }')
        if mode == 'delay':
            expect(page.locator('#signinSubmitStatus')).to_contain_text('正在处理')
            expect(page.locator('#signinSuccessNotice')).to_contain_text('11:42am')
            expect(page.locator('#signinSuccessNotice')).to_contain_text('自动签退')
            expect(page.locator('#signinSuccessNotice')).not_to_contain_text('连续换地点')
            attendance_rows = [args for sql, args in db.saved if sql.startswith('insert into attendance')]
            assert len(attendance_rows) == 1 and attendance_rows[0][6] == '11:42am'
            assert db.close_calls == [('CHE-208', '卫生')]
            expect(page.get_by_role('link', name='完成／返回')).to_be_visible()
            page.wait_for_timeout(10200)
            expect(page.locator('#signinSuccessNotice')).to_be_visible()
            page.screenshot(path=OUTPUT / 'signin-success-390.png')
        elif mode == 'unknown':
            expect(page.locator('#signinSubmitStatus')).to_contain_text('result is unknown')
            expect(page.get_by_role('link', name='Check records')).to_be_visible()
            expect(page.locator('#signinConfirm')).to_be_disabled()
            assert db.saved
            page.screenshot(path=OUTPUT / 'signin-unknown-390.png')
        else:
            expect(page.locator('[data-vol-result]')).to_contain_text('Check-in incomplete')
            assert not db.saved
    print('PASS | Check-in selection, review, one POST, saved time, persistent success, failure, unknown result')

    db.reset()
    db.sync_failed = True
    page.goto(origin + '/volunteer/my_schedule')
    page.once('dialog', lambda dialog: dialog.accept())
    page.get_by_role('button', name='Cancel this signup').click()
    expect(page.locator('[data-vol-result]')).to_have_text('Signup cancelled')
    expect(page.get_by_role('alert')).to_contain_text('Schedule sync failed')
    assert len(db.saved) == 1
    page.screenshot(path=OUTPUT / 'cancel-sync-warning-390.png')
    print('PASS | Mobile cancellation confirmation and separate schedule sync warning')

    db.reset()
    context.add_cookies([{'name': 'lang', 'value': 'zh', 'url': origin}])
    page.goto(origin + '/hongfa/join/1?token=offline-qr')
    page.locator('#hf-branch-btn-join').click()
    expect(page.locator('#hf-branch-btn-join')).to_have_text('STW')
    page.locator('#hf-branch-btn-join').click()
    expect(page.locator('#hf-branch-btn-join')).to_have_text('CHE')
    expect(page.locator('#hf-number-join')).to_have_attribute('autocomplete', 'off')
    expect(page.locator('#hf-pin-join')).to_have_attribute('autocomplete', 'off')
    page.locator('#hf-number-join').fill('208')
    page.locator('#hf-pin-join').fill('9999')
    page.locator('.vol-action-dock button').evaluate('(button) => button.click()')
    expect(page.locator('.vol-action-status')).to_contain_text('PIN')
    expect(page.locator('#hf-pin-join')).to_be_focused()
    assert not db.saved
    page.locator('#hf-pin-join').fill('1234')
    page.locator('#hf-pin-join').evaluate("(pin) => pin.addEventListener('blur', () => window.pinBlurred = true)")
    page.locator('.vol-action-dock button').evaluate('(button) => button.click()')
    expect(page.locator('[data-vol-result]')).to_contain_text('签到成功')
    assert page.evaluate('window.pinBlurred')
    db.reset()
    page.goto(origin + '/hongfa/join/1?token=offline-qr')
    page.locator('[data-kind="guest"]').click()
    page.locator('#hf-name').fill('离线访客')
    page.locator('#hf-phone-join').fill('0123456789')
    in_view(page, page.locator('.vol-action-dock button'))
    page.locator('.vol-action-dock button').click()
    expect(page.locator('[data-vol-result]')).to_contain_text('12:17:32')
    expect(page.locator('[data-vol-result]')).to_contain_text('新人签退凭证')
    assert len(db.saved) == 1
    page.screenshot(path=OUTPUT / 'hongfa-guest-390.png')
    assert not errors, errors
    print('PASS | Hongfa branch buttons, guest saved time and checkout code')
    context.close()


def route_cases(app, db):
    client = app.test_client()
    client.set_cookie('lang', 'en')
    form = {'year': 2026, 'month': 12, 'days': '99', 'role': '值班',
            'start_time': '09:00', 'end_time': '11:00'}
    db.reset()
    result = client.post('/volunteer/prebook', data=form)
    assert 'No dates were saved' in result.get_data(as_text=True) and not db.saved
    db.sync_failed = True
    result = client.post('/volunteer/prebook', data={**form, 'days': '3'})
    assert 'Multi-day signup saved' in result.get_data(as_text=True)
    assert 'Schedule sync failures' in result.get_data(as_text=True) and len(db.saved) == 1
    db.reset()
    db.sync_failed = True
    result = client.post('/volunteer/cancel/41')
    assert 'Signup cancelled' in result.get_data(as_text=True)
    assert 'Schedule sync failed' in result.get_data(as_text=True) and len(db.saved) == 1
    db.reset()
    db.signup_date = TODAY
    result = client.post('/volunteer/cancel/41')
    assert '今天不能取消报名' in result.get_data(as_text=True) and not db.saved
    result = client.post('/volunteer/signup', data={'role': '供台'})
    assert result.status_code == HTTPStatus.FORBIDDEN and 'Operation incomplete' in result.get_data(as_text=True)
    db.reset()
    lead = {'csrf_token': 'offline-csrf', 'branch': 'CHE', 'number': '208'}
    leader = '/hongfa/attendance/1/leader-checkin'
    result = client.post(leader, data={**lead, 'action': 'lookup'})
    assert 'Check-in not submitted' in result.get_data(as_text=True) and not db.saved
    result = client.post(leader, data={**lead, 'action': 'confirm', 'confirm_present': 'yes'})
    assert '12:17:32' in result.get_data(as_text=True) and len(db.saved) == 2
    print('PASS | Zero dates, partial sync errors, cancellation saved before failed sync, offering-table rule')

    handlers = app.extensions['offline_handlers']
    shifts = [
        dict(id=index, name=VOLUNTEER['姓名'], role='值班', volunteer_id='CHE-208',
             assignment_date=TODAY, assigned_place=place, start_time=start, end_time=end)
        for index, place, start, end in [(1, '观音堂', '09:00', '11:00'), (2, '活动中心', '11:00', '12:00'),
                                         (3, '观音堂', '13:00', '14:00')]
    ]
    postgres = ModuleType('psycopg2')
    extras = ModuleType('psycopg2.extras')
    extras.execute_values = lambda cur, sql, values: cur.execute(sql, values)
    db.reset()
    with patch.dict(sys.modules, {'psycopg2': postgres, 'psycopg2.extras': extras}):
        with patch.dict(handlers, {'load_today_assignments_for_volunteer': lambda *args: shifts,
                                   '_available_duty_ids': lambda *args: [1, 2]}):
            response = client.post('/signin', data={'volunteer_id': '208', 'pin': '1234', 'role': '值班',
                                                    'card_no': 'CARD-TEST', 'assignment_id': '1'}, follow_redirects=True)
    assert '11:42am' in response.get_data(as_text=True)
    assert 'moving between locations' in response.get_data(as_text=True)
    assert 'automatically check you out' in response.get_data(as_text=True)
    assert db.saved[0][1][6] == '11:42am' and 'assignment_ids=1,2' in db.saved[0][1][-1]
    assert [segment[0] for segment in db.saved[1][1]] == [1, 2]
    assert db.close_calls == [('CHE-208', '值班')]
    print('PASS | Continuous duty writes both matching segments with the saved receipt time')


def main():
    OUTPUT.mkdir(parents=True, exist_ok=True)
    logging.disable(logging.CRITICAL)
    db = FakeDatabase()
    app = build_app(db)
    route_cases(app, db)
    run_browser(app, db)
    for file in ('test_duty_v2.py', 'test_duty_v2_integration.py', 'test_duty_v2_safety.py'):
        subprocess.run([sys.executable, str(ROOT / file)], cwd=ROOT, check=True)
    print(f'PASS | All offline checks. Screenshots: {OUTPUT}')


if __name__ == '__main__':
    main()
