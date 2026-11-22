# volunteer_routes.py

import calendar
import os
import time

from db import get_conn
from utils import apply_branch_prefix
from schedule.blueprint import schedule_bp
from psycopg2.extras import RealDictCursor
from datetime import datetime, timedelta, date, timezone
from flask import request, redirect, url_for, render_template_string, redirect, session, jsonify, abort
from schedule.helpers import (
    find_volunteer_by_keyword,
    build_monthly_signup_text,
)

from lunar_rules import get_special_day_info
from schedule.constants import TIME_OPTIONS
from schedule.services.quote_service import get_daily_dharma
from schedule.services.settings_service import is_schedule_setting_on, current_meal_event_date, is_meal_festival_day

from schedule.builders.schedule_builder import sync_schedule_after_signup_change
from schedule.builders.time_utils import time_to_minutes, malaysia_today, malaysia_now
from schedule.services.publish_service import is_schedule_published
from schedule.services.whatsapp_service import build_whatsapp_from_assigned
from schedule.services.shortage_service import build_signup_shortage_notice
from schedule.services.signup_rules import (
    can_signup_supply_setup,
    can_signup_supply_remove,
)

from schedule.services.supply_service import (
    get_supply_signup_summary,
)

from schedule.volunteer_templates import (
    VOLUNTEER_SIGNUP_HTML,
    VOLUNTEER_PREBOOK_HTML,
)

def signup_has_formal_assignment(cur, signup_id):
    """Only an active assignment locks volunteer self-service changes."""
    cur.execute("""
        SELECT EXISTS (
            SELECT 1 FROM volunteer_schedule_assignments
            WHERE signup_id = %s
              AND coalesce(status, 'assigned') <> 'cancelled'
        ) AS locked
    """, (signup_id,))
    return bool(cur.fetchone()['locked'])


def assigned_signup_locked_message(volunteer_id=None):
    from html import escape
    back = ('/volunteer/my_schedule?keyword=' + escape(str(volunteer_id), quote=True)
            if volunteer_id else '/volunteer')
    return ("<h2>🔒 已正式安排，不能自行修改或取消报名</h2>"
            "<p>如需修改时间或请假，请联系负责人。</p>"
            "<p>Already assigned. Please contact the coordinator for changes.</p>"
            f'<a href="{back}">← 返回我的报名 / Back</a>'), 403


def find_duty_time_conflict(cur, volunteer_id, signup_date, start_time, end_time, exclude_id=None):
    """Same-day duty signup intervals are half-open: [start, end)."""
    new_start = time_to_minutes(start_time)
    new_end = time_to_minutes(end_time)
    if new_start is None or new_end is None or new_end <= new_start:
        return {"invalid": True}
    cur.execute("""
        SELECT id, start_time, end_time
        FROM volunteer_schedule_signups
        WHERE volunteer_id = %s AND signup_date = %s AND role = '值班'
          AND coalesce(status, 'pending') <> 'cancelled'
          AND (%s IS NULL OR id <> %s)
    """, (volunteer_id, signup_date, exclude_id, exclude_id))
    for row in cur.fetchall():
        old_start = time_to_minutes(row['start_time'])
        old_end = time_to_minutes(row['end_time'])
        if old_start is None or old_end is None:
            continue
        if new_start < old_end and old_start < new_end:
            return row
    return None


def duty_conflict_message(row, back='/volunteer'):
    # Bilingual mobile-friendly conflict page; signup rules unchanged.
    from html import escape

    invalid = bool(row.get('invalid'))
    start = escape(str(row.get('start_time') or ''))
    end = escape(str(row.get('end_time') or ''))
    safe_back = escape(str(back), quote=True)
    title = ('时间格式错误 / Invalid time' if invalid
             else '值班时间重叠 / Overlapping Duty Hours')
    detail = ('''<p class="desc">请选择正确的开始和结束时间。<br>
                 <span>Please select a valid start and end time.</span></p>'''
              if invalid else f'''<p class="desc">您已有以下报名时段：<br>
                 <span>You have already signed up for:</span></p>
                 <div class="time-pill">{start} ～ {end}</div>''')
    return f'''<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title}</title>
<style>
* {{ box-sizing: border-box; }}
body {{ margin:0; min-height:100vh; padding:28px 16px; display:grid;
       place-items:center; font-family:"Microsoft YaHei","Noto Sans SC",Arial,sans-serif;
       color:#34292a; background:radial-gradient(circle at 10% 10%,#fff2d9,transparent 40%),
       linear-gradient(160deg,#fffaf1,#f7f0e6); }}
.card {{ width:min(100%,760px); background:#fff; border-radius:24px;
         overflow:hidden; box-shadow:0 14px 42px #654b3a20; border:1px solid #f1e4db; }}
.header {{ padding:26px 26px; text-align:center; background:#fff0f0;
           border-top:5px solid #d9333f; }}
.icon {{ display:inline-grid; place-items:center; width:54px; height:54px;
         background:#d9333f; color:#fff; border-radius:50%; font-size:34px;
         font-weight:bold; margin-bottom:12px; }}
h1 {{ margin:0; font-size:clamp(23px,4vw,31px); color:#a71e29; line-height:1.4; }}
.content {{ padding:28px; text-align:center; }}
.existing {{ padding:22px 18px; border:1px solid #f4c2c6;
             background:#fff2f3; border-radius:16px; }}
.desc {{ margin:0 0 16px; font-size:21px; font-weight:700; line-height:1.6; }}
.desc span {{ font-size:16px; font-weight:400; color:#68575b; }}
.time-pill {{ display:inline-block; max-width:100%; padding:10px 24px;
              border:1px solid #ec8d96; border-radius:50px; background:#fff;
              font-size:clamp(22px,5vw,30px); font-weight:800; color:#b51c2c; }}
.notice {{ margin:20px 0 26px; padding:18px; border-radius:14px;
           background:#fff5d8; border:1px solid #f6db8b; line-height:1.7;
           font-size:19px; font-weight:600; }}
.notice small {{ display:block; font-size:15px; font-weight:400; color:#68572f; }}
.back {{ display:inline-block; min-width:220px; padding:14px 25px;
         border-radius:40px; background:linear-gradient(135deg,#9147bd,#632a9b);
         color:white; text-decoration:none; font-size:20px; font-weight:700;
         box-shadow:0 5px 12px #63358a30; }}
.back:hover {{ filter:brightness(1.08); }}
.back:focus-visible {{ outline:3px solid #4c217b; outline-offset:4px; }}
@media(max-width:480px) {{ .content {{padding:18px;}} .header {{padding:22px 14px;}}
 .desc {{font-size:18px;}} .notice {{font-size:17px;}} .back {{width:100%;}} }}
</style>
</head>
<body>
<main class="card">
  <header class="header"><div class="icon" aria-hidden="true">×</div>
    <h1>{title}</h1></header>
  <section class="content">
    <div class="existing">{detail}</div>
    <div class="notice">⚠️ 同一天可报名多个不重叠时段，请更换时间。
      <small>Multiple non-overlapping shifts are allowed on the same day.<br>
      Please choose another time slot.</small></div>
    <a class="back" href="{safe_back}">← 返回 / Back</a>
  </section>
</main>
</body>
</html>'''


def find_next_meal_signup_date(base_date, days_ahead=7):
    """膳食组仅返回当前开放窗口对应的佛日，禁止跳到未来尚未开放的日期。"""
    event_date = current_meal_event_date()
    if event_date and is_schedule_setting_on("meal_signup_open"):
        return event_date, get_special_day_info(event_date)
    return None, None


def cleaning_signup_window(target_date):
    """按马来西亚时间判定卫生报名是否截止；未来日期始终开放。"""
    now = malaysia_now()
    info = get_special_day_info(target_date)
    special = info.get("template_type") in ("lunar_1_15", "buddhist_festival")
    cutoff_hour = 8 if special else 10
    is_open = target_date > now.date() or (
        target_date == now.date() and now.hour < cutoff_hour
    )
    return is_open, cutoff_hour, special


def cleaning_signup_closed_message(cutoff_hour):
    time_text = "8:00am" if cutoff_hour == 8 else "10:00am"
    return (
        "<h2>❌ 卫生报名时间已截止 / Cleaning signup closed</h2>"
        f"<p>当天卫生报名于 {time_text} 截止。</p>"
        "<p>请改报明天或未来日期。 / Please select a future date.</p>"
        '<a href="/volunteer">返回义工报名 / Back</a>'
    )

def cleaning_full_message():
    """卫生报名满额：中英双语、手机及电脑自适应提示页。"""
    return """<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>卫生岗位已满 / Cleaning duty is full</title>
<style>
* { box-sizing: border-box; }
body { margin: 0; min-height: 100vh; padding: 24px 16px;
       display: flex; justify-content: center; align-items: center;
       font-family: "Microsoft YaHei", "Noto Sans SC", Arial, sans-serif;
       background: linear-gradient(145deg, #fff8ef 0%, #fff0ee 54%, #f5f7ff 100%);
       color: #332b2b; }
.card { width: 100%; max-width: 640px; padding: 38px 34px 32px;
        border-radius: 26px; background: white; text-align: center;
        box-shadow: 0 18px 55px rgba(112, 66, 57, .12);
        border: 1px solid #f8e4e0; }
.icon { width: 82px; height: 82px; margin: 0 auto 22px;
        display: grid; place-items: center; border-radius: 50%;
        background: #fff0ed; color: #d94840; font-size: 44px; }
h1 { font-size: clamp(26px, 5vw, 34px); color: #a72f31;
     line-height: 1.35; margin: 0 0 8px; }
.en-title { font-size: 17px; color: #856a6a; margin-bottom: 26px; }
.count { display: inline-block; border-radius: 999px;
         background: #fff0ee; border: 1px solid #ffd4ce;
         color: #c43837; font-weight: 800; font-size: 24px;
         padding: 9px 28px; margin-bottom: 24px; }
.notice { background: #fff9ed; border: 1px solid #f5dfb0;
          border-radius: 17px; padding: 22px 18px; line-height: 1.8; }
.notice .zh { font-size: 21px; font-weight: 700; }
.notice .en { font-size: 15px; color: #71665b; margin-top: 8px; }
.tip { font-size: 16px; color: #6b655f; margin: 22px 0; }
.back { display: block; width: 100%; padding: 16px 20px;
        background: linear-gradient(135deg, #228b66, #157457);
        color: #fff; font-size: 20px; font-weight: 700;
        border-radius: 14px; text-decoration: none;
        box-shadow: 0 5px 16px rgba(24, 121, 86, .2); }
.back:hover { filter: brightness(1.06); }
.back:focus-visible { outline: 3px solid #1c7557; outline-offset: 4px; }
@media (max-width: 480px) {
  .card { padding: 30px 20px 24px; border-radius: 20px; }
  .notice .zh { font-size: 18px; }
  .back { font-size: 18px; }
}
</style>
</head>
<body>
<main class="card">
  <div class="icon" aria-hidden="true">✕</div>
  <h1>卫生岗位已满</h1>
  <div class="en-title">Cleaning duty is full</div>
  <div class="count">3 / 3 人 · FULL</div>
  <div class="notice">
    <div class="zh">当天已有 3 位义工成功报名，<br>卫生岗位暂时没有空缺。</div>
    <div class="en">All 3 cleaning places have been taken.<br>Please choose another date.</div>
  </div>
  <div class="tip">🙏 感恩您的发心，欢迎选择其他日期继续护持。</div>
  <a class="back" href="/volunteer">← 返回义工报名 / Back to Signup</a>
</main>
</body>
</html>"""


def lock_and_count_cleaning(cur, signup_date):
    """同一日期串行检查容量，防止多人同时报名超过3位。"""
    cur.execute("SELECT pg_advisory_xact_lock(hashtext('cleaning_signup'), hashtext(%s))",
                (str(signup_date),))
    cur.execute("""
        SELECT count(*) AS cnt FROM volunteer_schedule_signups
        WHERE signup_date = %s AND role = '卫生'
          AND coalesce(status, 'pending') <> 'cancelled'
    """, (signup_date,))
    return int(cur.fetchone()['cnt'] or 0)


# 义工报名独立登录；不改变签到系统及负责人登录权限。
from attendance_service import find_volunteer, verify_pin_for_volunteer
from datetime import timezone as _login_timezone

# PostgreSQL shared across Render workers; table is installed separately via SQL.
# No automatic schema changes and no changes to attendance or signup data.
def _login_gate(identifier, ip, *, success=False):
    """Atomic account/IP login throttle. Return seconds remaining when blocked."""
    from datetime import datetime, timedelta
    now = datetime.now(_login_timezone.utc)
    keys = [('account', identifier, 5), ('ip', ip, 20)]
    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            # Deterministic lock order prevents deadlocks on concurrent logins.
            for kind, value, limit in sorted(keys):
                cur.execute("""INSERT INTO volunteer_login_limits (scope, subject, failures, window_start, locked_until)
                               VALUES (%s,%s,0,%s,NULL) ON CONFLICT (scope,subject) DO NOTHING""",
                            (kind, value, now))
            records = {}
            for kind, value, limit in sorted(keys):
                cur.execute("""SELECT failures, window_start, locked_until FROM volunteer_login_limits
                               WHERE scope=%s AND subject=%s FOR UPDATE""", (kind, value))
                records[(kind, value)] = cur.fetchone()
            remaining = max((max(0, int((r['locked_until'] - now).total_seconds()) + 1)
                             for r in records.values() if r['locked_until'] and r['locked_until'] > now), default=0)
            if remaining:
                return remaining
            for kind, value, limit in keys:
                r = records[(kind, value)]
                if success:
                    # Only clear the authenticated account, not the shared IP budget.
                    if kind == 'account':
                        cur.execute("""UPDATE volunteer_login_limits SET failures=0, window_start=%s, locked_until=NULL
                                       WHERE scope=%s AND subject=%s""", (now, kind, value))
                    continue
                recent = r['window_start'] and r['window_start'] > now - timedelta(minutes=15)
                failures = (int(r['failures']) if recent else 0) + 1
                until = now + timedelta(minutes=15) if failures >= limit else None
                cur.execute("""UPDATE volunteer_login_limits SET failures=%s, window_start=%s, locked_until=%s
                               WHERE scope=%s AND subject=%s""",
                            (failures, r['window_start'] if recent else now, until, kind, value))
            return 0


def _login_locked(identifier, ip):
    """Read shared lock state before checking a PIN; no failure is counted here."""
    from datetime import datetime
    now = datetime.now(_login_timezone.utc)
    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""SELECT locked_until FROM volunteer_login_limits
                           WHERE (scope='account' AND subject=%s) OR (scope='ip' AND subject=%s)""",
                        (identifier, ip))
            return max((max(0, int((r['locked_until'] - now).total_seconds()) + 1)
                        for r in cur.fetchall() if r['locked_until'] and r['locked_until'] > now), default=0)


def _volunteer_identity():
    return str(session.get("volunteer_signup_id") or "")


def _volunteer_self_row(signup_id):
    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("SELECT volunteer_id FROM volunteer_schedule_signups WHERE id=%s", (signup_id,))
            return cur.fetchone()


@schedule_bp.before_request
def _protect_volunteer_pages():
    # 仅保护义工端；不改变 /schedule/admin 等管理员路径。
    if not request.path.startswith('/volunteer'):
        return None
    if request.path in ('/volunteer/login', '/volunteer/logout'):
        return None
    if not _volunteer_identity():
        if request.method != 'GET' or request.path == '/volunteer/query_volunteer':
            return jsonify(ok=False, message='请先登录义工报名系统'), 401
        return redirect('/volunteer/login')
    if request.method == 'POST':
        # 所有义工端写入都要求同一登录会话产生的 CSRF token。
        import hmac
        if not hmac.compare_digest(str(request.form.get('_csrf') or ''), str(session.get('volunteer_signup_csrf') or '')):
            return '❌ 页面已过期，请返回报名首页重新操作', 403
    if request.endpoint in ('schedule.volunteer_edit_signup', 'schedule.volunteer_edit_signup_post', 'schedule.volunteer_cancel_signup'):
        row = _volunteer_self_row(request.view_args['signup_id'])
        if not row or str(row['volunteer_id']) != _volunteer_identity():
            abort(403)
    return None


@schedule_bp.route('/volunteer/login', methods=['GET', 'POST'])
def volunteer_signup_login():
    from html import escape
    import secrets
    error = ''
    if request.method == 'POST':
        branch = request.form.get('branch', 'CHE').upper()
        digits = request.form.get('number', '').strip()
        pin = request.form.get('pin', '').strip()
        # Use only the trusted server peer address; do not trust user-supplied X-Forwarded-For.
        ip = request.remote_addr or 'unknown'
        valid_number = branch in ('CHE', 'STW') and digits.isdigit() and 1 <= len(digits) <= 4
        volunteer_id = (('STW-' if branch == 'STW' else 'CHE-') + str(int(digits))) if valid_number else 'invalid'
        if valid_number and branch == 'CHE' and int(digits) >= 800:
            volunteer_id = str(int(digits))
        # Hash identifiers at rest so this security-only table stores no plain volunteer numbers/IPs.
        import hashlib
        account_key = hashlib.sha256(volunteer_id.encode('utf-8')).hexdigest()
        ip_key = hashlib.sha256(ip.encode('utf-8')).hexdigest()
        try:
            wait = _login_locked(account_key, ip_key)
            if wait:
                error = f'登录暂时受限，请约 {max(1, (wait + 59) // 60)} 分钟后重试。'
            else:
                authenticated = False
                if valid_number and pin:
                    volunteer = find_volunteer(volunteer_id)
                    has_personal_credential = bool(str(volunteer.get('PIN') or '').strip() or ''.join(c for c in str(volunteer.get('电话号码') or '') if c.isdigit())) if volunteer else False
                    if volunteer and volunteer.get('状态') in ('在册', '', None) and has_personal_credential and verify_pin_for_volunteer(volunteer, pin):
                        canonical_id = str(volunteer.get('编号') or '')
                        matches = find_volunteer_by_keyword(canonical_id)
                        if len(matches) == 1 and str(matches[0]['id']) == canonical_id:
                            authenticated = True
                if authenticated:
                    _login_gate(account_key, ip_key, success=True)
                    session['volunteer_signup_id'] = canonical_id
                    session['volunteer_signup_name'] = str(matches[0]['name'])
                    session['volunteer_signup_csrf'] = secrets.token_urlsafe(32)
                    session.permanent = True
                    return redirect('/volunteer')
                _login_gate(account_key, ip_key)
                error = '编号或 PIN 不正确，请重新输入。'
        except Exception:
            # Fail closed if database/table is unavailable; never bypass lock checks.
            import logging
            logging.exception('Volunteer login security storage unavailable')
            error = '登录验证暂时无法使用，请稍后重试或联系负责人。'
    return render_template_string('''<!doctype html><html lang="zh"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>义工报名身份验证</title><style>body{font-family:Arial,"Microsoft YaHei",sans-serif;background:#f0f8f1;padding:20px}.box{max-width:480px;margin:7vh auto;background:white;padding:28px;border-radius:20px;box-shadow:0 5px 24px #0001}h1{text-align:center;color:#227a4d;font-size:28px}label{display:block;margin:18px 0 8px;font-size:20px}input,button{box-sizing:border-box;font-size:23px;padding:15px;border-radius:12px}input{width:100%;border:1px solid #bbb}.row{display:flex;gap:10px}.row input{flex:1;min-width:0}.branch{background:#28a745;color:white;border:0;min-width:90px}button.submit{width:100%;background:#16834a;color:white;border:0;margin-top:24px;font-weight:bold}.error{color:#bd2636;font-size:18px}</style></head><body><main class="box"><h1>🪷 义工报名身份验证</h1><p style="text-align:center">请输入自己的编号和 PIN，验证一次即可报名。</p>{% if error %}<p class="error">{{error}}</p>{% endif %}<form method="post"><label>义工编号</label><div class="row"><button class="branch" type="button" id="branchBtn" onclick="let b=document.getElementById('branch');b.value=b.value==='CHE'?'STW':'CHE';this.textContent=b.value;this.style.background=b.value==='CHE'?'#28a745':'#dc3545'">CHE</button><input type="hidden" name="branch" id="branch" value="CHE"><input name="number" inputmode="numeric" pattern="[0-9]+" required autofocus placeholder="例如：208"></div><label>个人 PIN</label><input name="pin" type="password" inputmode="numeric" autocomplete="off" required><button class="submit" type="submit">进入义工报名</button></form></main></body></html>''', error=error)


@schedule_bp.route('/volunteer/logout', methods=['POST'])
def volunteer_signup_logout():
    import hmac
    if not hmac.compare_digest(str(request.form.get('_csrf') or ''), str(session.get('volunteer_signup_csrf') or '')):
        abort(403)
    for key in ('volunteer_signup_id', 'volunteer_signup_name', 'volunteer_signup_csrf'):
        session.pop(key, None)
    return redirect('/volunteer/login')


@schedule_bp.after_request
def _volunteer_form_identity(response):
    if not request.path.startswith('/volunteer') or request.path in ('/volunteer/login', '/volunteer/logout'):
        return response
    if not _volunteer_identity():
        return response
    response.headers['Cache-Control'] = 'private, no-store'
    if not response.content_type.startswith('text/html'):
        return response
    import re
    from html import escape
    html = response.get_data(as_text=True)
    token = escape(str(session.get('volunteer_signup_csrf') or ''), quote=True)
    html = re.sub(r'(<form\b[^>]*\bmethod=["\']post["\'][^>]*>)', lambda m: m.group(1) + '<input type="hidden" name="_csrf" value="'+token+'">', html, flags=re.I)
    response.set_data(html)
    response.headers['Cache-Control'] = 'private, no-store'
    return response


@schedule_bp.route("/volunteer")
def volunteer_home():

    now = malaysia_now()
    today = now.date()

    multi_day_signup_open = is_schedule_setting_on("multi_day_signup_open")
    meal_signup_open = is_schedule_setting_on("meal_signup_open")
    food_offering_open = is_schedule_setting_on("food_offering_open")

    if today.month == 12:
        prebook_year = today.year + 1
        prebook_month = 1
    else:
        prebook_year = today.year
        prebook_month = today.month + 1

    if now.hour >= 19:
        default_date = (now + timedelta(days=1)).strftime("%Y-%m-%d")
        schedule_label = "明日值班报名情况"
    else:
        default_date = now.strftime("%Y-%m-%d")
        schedule_label = "今日值班报名情况"

    meal_button_show = False
    meal_status_text = ""
    meal_date = ""
    meal_button_date = ""
    meal_count = 0

    base_date = datetime.strptime(default_date, "%Y-%m-%d").date()

    active_meal_date = current_meal_event_date(now)
    if meal_signup_open and active_meal_date:
        meal_button_show = True
        meal_date = active_meal_date.strftime("%Y-%m-%d")
        check_date = active_meal_date

    if meal_button_show:
        with get_conn() as conn:
            with conn.cursor(cursor_factory=RealDictCursor) as cur:
                cur.execute("""
                    select count(*) as cnt
                    from volunteer_schedule_signups
                    where signup_date = %s
                    and role = '膳食'
                    and coalesce(status, 'pending') <> 'cancelled'
                """, (meal_date,))

                total_meal = cur.fetchone()["cnt"]

        meal_count = min(total_meal, 9)

        meal_button_date = f"{check_date.month}月{check_date.day}日（星期{'一二三四五六日'[check_date.weekday()]}）"

        meal_status_text = "已满" if meal_count >= 9 else f"{meal_count}/9"

    target_date = datetime.strptime(
        default_date,
        "%Y-%m-%d"
    ).date()

    supply_summary = get_supply_signup_summary(default_date)

    supply_setup_open = (
        can_signup_supply_setup(target_date)
        and not supply_summary["setup"]["is_full"]
    )

    supply_remove_open = (
        can_signup_supply_remove(target_date)
        and not supply_summary["remove"]["is_full"]
    )

    # 设供台或收供台只要还有一个可报名，就显示“供台”按钮
    supply_signup_available = (
        supply_setup_open
        or supply_remove_open
    )
        
    return render_template_string(
    VOLUNTEER_SIGNUP_HTML,
    login_id=_volunteer_identity(), login_name=session.get("volunteer_signup_name", ""),
    default_date=default_date,
    schedule_label=schedule_label,
    times=TIME_OPTIONS,
    signup_date=default_date,
    prebook_year=prebook_year,
    prebook_month=prebook_month,
    meal_button_show=meal_button_show,
    meal_status_text=meal_status_text,
    meal_date=meal_date,
    meal_button_date=meal_button_date,
    meal_count=meal_count,
    multi_day_signup_open=multi_day_signup_open,
    meal_signup_open=meal_signup_open,
    food_offering_open=food_offering_open,

    # 供台报名状态
    supply_signup_available=supply_signup_available,
    supply_setup_open=supply_setup_open,
    supply_remove_open=supply_remove_open,
    supply_summary=supply_summary,
)

@schedule_bp.route("/volunteer/signup", methods=["POST"])
def volunteer_signup():

    _signup_timing = os.environ.get("SCHEDULE_SIGNUP_TIMING") == "1"
    _signup_t0 = time.perf_counter()
    def _signup_mark(stage):
        if _signup_timing:
            print(f"[signup-timing] {stage}: {time.perf_counter() - _signup_t0:.3f}s", flush=True)

    keyword = _volunteer_identity()
    signup_date = request.form.get("signup_date", "").strip()
    role = request.form.get("role", "").strip()

    supply_task = (
        request.form.get("supply_task", "")
        .strip()
        .lower()
    )

    food_offering_task = (
        request.form.get("food_offering_task", "")
        .strip()
        .lower()
    )

    food_offering_item = (
        request.form.get("food_offering_item", "")
        .strip()
    )

    start_time = request.form.get("start_time", "").strip()
    end_time = request.form.get("end_time", "").strip()
    branch = request.form.get("branch", "CHE").strip().upper()

    keyword = _volunteer_identity()
    daily_quote = get_daily_dharma()

    if keyword.isdigit() and branch == "STW":
        keyword = f"STW-{keyword}"

    matches = find_volunteer_by_keyword(keyword)
    _signup_mark("volunteer_lookup_done")

    if not matches:
        return "❌ 找不到义工，请检查编号 / 姓名<br><a href='/volunteer'>返回</a>"

    if len(matches) > 1:
        return "❌ 找到多个同名义工，请用义工编号报名<br><a href='/volunteer'>返回</a>"

    vol = matches[0]
    vol_id = str(vol["id"])
    name = str(vol["name"])

    meal_role = None

    # =========================
    # 处理岗位时间
    # =========================

    if role == "值班":

        s_min = time_to_minutes(start_time)
        e_min = time_to_minutes(end_time)

        if s_min is None or e_min is None:
            return "❌ 时间格式错误，请重新选择<br><a href='/volunteer'>返回</a>"

        if e_min <= s_min:
            return "❌ 结束时间必须比开始时间迟<br><a href='/volunteer'>返回</a>"

    elif role == "卫生":

        try:
            cleaning_date = datetime.strptime(signup_date, "%Y-%m-%d").date()
        except ValueError:
            return "❌ 日期格式错误 / Invalid date<br><a href='/volunteer'>返回</a>"
        cleaning_open, cutoff_hour, cleaning_special = cleaning_signup_window(cleaning_date)
        if not cleaning_open:
            return cleaning_signup_closed_message(cutoff_hour)

        date_obj = datetime.strptime(signup_date, "%Y-%m-%d").date()
        special_info = get_special_day_info(date_obj)

        if special_info.get("template_type") in ("lunar_1_15", "buddhist_festival"):
            start_time = "6:00am"
            end_time = "8:00am"
        else:
            start_time = "8:00am"
            end_time = "10:00am"

    elif role == "供台":

        date_obj = datetime.strptime(
            signup_date,
            "%Y-%m-%d"
        ).date()

        special_info = get_special_day_info(date_obj)

        if special_info.get("template_type") not in [
            "lunar_1_15",
            "buddhist_festival"
        ]:

            return """
            <h1>❌ 这一天不需要供台报名</h1>

            <p>
            供台通常只开放于初一、
            十五及佛诞大日子。
            </p>

            <a href="/volunteer">
            返回重新报名
            </a>
            """

        if supply_task not in [
            "setup",
            "remove"
        ]:

            return """
            <h1>❌ 请先选择供台工作</h1>

            <p>
            请选择：
            设供台 或 收供台
            </p>

            <a href="/volunteer">
            返回
            </a>
            """

        # 连续供奉期间不得重复设台或提前收台；其余日期原逻辑不变。
        if (special_info.get("lunar_month"), special_info.get("lunar_day")) in ((11, 15), (11, 17)):
            allowed = can_signup_supply_setup(date_obj) if supply_task == "setup" else can_signup_supply_remove(date_obj)
            if not allowed:
                return '❌ 当天不开放此供台工作 / Not available<br><a href="/volunteer">返回</a>', 400

        if supply_task == "setup":

            start_time = "6:00am"
            end_time = "8:00am"

        else:

            start_time = "12:00pm"
            end_time = "1:00pm"

    elif role == "膳食":

        if not is_schedule_setting_on("meal_signup_open"):
            return """
            <h2>❌ 膳食组报名尚未开放 / Meal signup closed</h2>
            <p>只在指定佛日前一天至佛日当天 2:00pm 开放，负责人亦可临时关闭。</p>
            <a href="/volunteer">返回义工报名 / Back</a>
            """

        selected_date = datetime.strptime(
            signup_date,
            "%Y-%m-%d"
        ).date()

        meal_date, special_info = find_next_meal_signup_date(
            selected_date,
            days_ahead=7
        )

        if not meal_date:
            return """
            <h1>❌ 近期没有开放膳食组报名</h1>
            <p>当前不在指定膳食报名时段。</p>
            <p>请检查日期是否选错。</p>
            <a href="/volunteer">返回重新报名</a>
            """

        signup_date = meal_date.strftime("%Y-%m-%d")
        start_time = "8:00am"
        end_time = "2:00pm"


    elif role == "素食结缘":

        if not is_schedule_setting_on(
            "food_offering_open"
        ):
            return """
            <h1>❌ 素食结缘报名尚未开放</h1>
            <p>请等待负责人开放报名。</p>
            <a href="/volunteer">返回</a>
            """

        valid_food_offering_tasks = {
            "main_food",
            "vegetable",
            "dessert",
            "fruit",
            "drink",
        }

        if food_offering_task not in valid_food_offering_tasks:
            return """
            <h1>❌ 请先选择素食结缘项目</h1>

            <p>
            请选择主食、菜肴、甜品、水果或饮料。
            </p>

            <a href="/volunteer">
            返回重新报名
            </a>
            """

        # 素食结缘暂时没有固定值班时间
        start_time = None
        end_time = None


    else:
        return "❌ 岗位错误，请重新选择<br><a href='/volunteer'>返回</a>"

    # =========================
    # 新增 / 更新报名
    # =========================
    if role != "供台":
        supply_task = None

    if role != "素食结缘":
        food_offering_task = None

    _signup_mark("before_signup_db")
    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:

            # 卫生名额必须与写入同一个数据库事务，避免并发超额。
            if role == "卫生":
                cleaning_count = lock_and_count_cleaning(cur, signup_date)
                cur.execute("""
                    SELECT id FROM volunteer_schedule_signups
                    WHERE volunteer_id = %s AND signup_date = %s
                      AND role = '卫生'
                      AND coalesce(status, 'pending') <> 'cancelled'
                    LIMIT 1
                """, (vol_id, signup_date))
                already_cleaning = cur.fetchone() is not None
                if not already_cleaning and cleaning_count >= 3:
                    return cleaning_full_message()

            # 默认值，避免非膳食报名时 meal_role 未定义
            meal_role = None

            # 只有对应岗位才保留细项
            if role != "供台":
                supply_task = None

            if role != "素食结缘":
                food_offering_task = None
                food_offering_item = None

            if role == "值班":
                conflict = find_duty_time_conflict(cur, vol_id, signup_date, start_time, end_time)
                if conflict:
                    # Exact same time is an idempotent repeat; other overlaps are blocked.
                    if str(conflict.get('start_time')) != start_time or str(conflict.get('end_time')) != end_time:
                        return duty_conflict_message(conflict)

            # 同一天的膳食报名先加锁，避免并发重复提交。
            if role == "膳食":
                cur.execute("SELECT pg_advisory_xact_lock(hashtext('meal_signup'), hashtext(%s))", (str(signup_date),))

            # ============================================================
            # 检查重复报名
            #
            # 值班：同一天允许多轮；只有开始、结束时间都相同才算重复。
            # 其他岗位：维持原本规则（同岗位 / 同细项只保留一笔）。
            # ============================================================
            if role == "值班":
                cur.execute("""
                    select id
                    from volunteer_schedule_signups
                    where volunteer_id = %s
                    and signup_date = %s
                    and role = '值班'
                    and start_time = %s
                    and end_time = %s
                    and coalesce(status, 'pending') <> 'cancelled'
                    limit 1
                """, (
                    vol_id,
                    signup_date,
                    start_time,
                    end_time
                ))
            else:
                cur.execute("""
                    select id
                    from volunteer_schedule_signups
                    where volunteer_id = %s
                    and signup_date = %s
                    and role = %s
                    and (
                            role <> '供台'
                            or coalesce(supply_task, 'setup')
                            = coalesce(%s, 'setup')
                        )
                    and (
                            role <> '素食结缘'
                            or food_offering_task = %s
                        )
                    and coalesce(status, 'pending') <> 'cancelled'
                    limit 1
                """, (
                    vol_id,
                    signup_date,
                    role,
                    supply_task,
                    food_offering_task
                ))

            existing = cur.fetchone()

            # 派餐义工普通报名最多9位；负责人可从后台特批第10位。
            if role == "膳食":
                if existing:
                    return """<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>派餐义工 · 已报名</title>
<style>
* { box-sizing:border-box; }
body { margin:0; min-height:100vh; padding:24px 16px; display:flex;
       align-items:center; justify-content:center; font-family:"Microsoft YaHei","Noto Sans SC",Arial,sans-serif;
       color:#344054; background:linear-gradient(145deg,#edf7ff,#f6f9ff 55%,#eaf5f0); }
.card { width:100%; max-width:620px; background:white; border-radius:24px;
        box-shadow:0 16px 48px rgba(41,80,110,.13); overflow:hidden; text-align:center; }
.head { padding:30px 22px 22px; background:#fff8ec; border-top:5px solid #e8a83d; }
.icon { display:grid; place-items:center; width:74px; height:74px; margin:0 auto 14px;
        border-radius:50%; background:#ffedc9; font-size:40px; }
h1 { margin:0; color:#9a5a0c; font-size:clamp(25px,5vw,32px); line-height:1.4; }
.content { padding:26px 28px 30px; }
.confirm { padding:18px 16px; border-radius:16px; background:#edf9f2; border:1px solid #bfe7cd;
           color:#176b3d; font-size:21px; font-weight:700; line-height:1.65; }
.tip { margin:20px 0; font-size:18px; line-height:1.8; color:#566273; }
.thanks { margin:0 0 22px; font-size:18px; color:#90612e; }
.actions { display:grid; gap:12px; }
.btn { display:block; padding:16px 14px; border-radius:14px; color:#fff; text-decoration:none;
       font-size:20px; font-weight:700; line-height:1.4; }
.btn-green { background:#21885a; }
.btn-blue { background:#2779ba; }
.btn:hover { filter:brightness(1.06); }
.btn:focus-visible { outline:3px solid #2b6089; outline-offset:3px; }
@media(max-width:480px) { body {padding:14px;} .head {padding:24px 14px 20px;}
  .content {padding:20px 16px 24px;} .confirm {font-size:18px;}
  .tip,.thanks {font-size:16px;} .btn {font-size:18px;} }
</style>
</head>
<body>
<main class="card">
  <div class="head"><div class="icon" aria-hidden="true">⚠️</div>
    <h1>您已经提交报名了！</h1></div>
  <div class="content">
    <div class="confirm">✅ 您已成功报名当天的派餐义工，<br>无需重复报名。</div>
    <p class="tip">如需更改，请先进入「我的报名」取消原有报名，再重新报名。</p>
    <p class="thanks">🙏 感恩您的发心！</p>
    <div class="actions">
      <a class="btn btn-green" href="/volunteer/my_schedule_search">📋 查看／修改／取消</a>
      <a class="btn btn-blue" href="/volunteer">← 返回义工报名</a>
    </div>
  </div>
</main>
</body>
</html>"""
                cur.execute("""
                    select count(*) as cnt
                    from volunteer_schedule_signups
                    where signup_date = %s
                    and role = '膳食'
                    and coalesce(status, 'pending') <> 'cancelled'
                """, (signup_date,))

                meal_count = int(cur.fetchone()["cnt"] or 0)

                if existing:
                    # 更新已有报名不改变原来的派餐/候补身份。
                    cur.execute("""
                        SELECT meal_role FROM volunteer_schedule_signups WHERE id = %s
                    """, (existing["id"],))
                    saved_meal = cur.fetchone()
                    meal_role = (saved_meal or {}).get("meal_role")
                    if not meal_role:
                        # 兼容旧记录：按创建顺序重新判定身份。
                        cur.execute("""
                            SELECT id FROM volunteer_schedule_signups
                            WHERE signup_date = %s AND role = '膳食'
                              AND coalesce(status, 'pending') <> 'cancelled'
                            ORDER BY created_at, id
                        """, (signup_date,))
                        ids_in_order = [r['id'] for r in cur.fetchall()]
                        meal_role = ('派餐义工' if existing['id'] in ids_in_order[:9]
                                     else '候补义工')
                elif meal_count < 9:
                    meal_role = "派餐义工"
                else:
                    return "❌ 派餐义工9个报名名额已满；特殊第10位请联系负责人。<br><a href='/volunteer'>返回</a>"

            if existing:
                if signup_has_formal_assignment(cur, existing['id']):
                    return assigned_signup_locked_message(vol_id)

                signup_id = existing["id"]

                cur.execute("""
                    update volunteer_schedule_signups
                    set supply_task = %s,
                        food_offering_task = %s,
                        food_offering_item = %s,
                        start_time = %s,
                        end_time = %s,
                        status = 'pending',
                        assigned_place = null,
                        remarks = '义工网页更新报名',
                        meal_role = %s
                    where id = %s
                """, (
                    supply_task,
                    food_offering_task,
                    food_offering_item,
                    start_time,
                    end_time,
                    meal_role,
                    signup_id
                ))

                result_title = "✅ 已更新报名"

            else:

                cur.execute("""
                    insert into volunteer_schedule_signups
                    (
                        volunteer_id,
                        name,
                        signup_date,
                        role,
                        supply_task,
                        food_offering_task,
                        food_offering_item,
                        start_time,
                        end_time,
                        status,
                        remarks,
                        meal_role
                    )
                    values
                    (
                        %s,
                        %s,
                        %s,
                        %s,
                        %s,
                        %s,
                        %s,
                        %s,
                        %s,
                        'pending',
                        '义工报名',
                        %s
                    )
                    returning id
                """, (
                    vol_id,
                    name,
                    signup_date,
                    role,
                    supply_task,
                    food_offering_task,
                    food_offering_item,
                    start_time,
                    end_time,
                    meal_role
                ))

                signup_id = cur.fetchone()["id"]

                result_title = "🎉 报名成功"

            conn.commit()
    _signup_mark("signup_db_committed")
    
    # =========================
    # V2 同步：只走统一函数
    # =========================

    sync_schedule_after_signup_change(
        signup_id,
        action="upsert",
        changed_by="volunteer"
    )
    _signup_mark("schedule_sync_done")

    # =========================
    # 判断发布状态与安排结果
    # =========================

    published = is_schedule_published(signup_date)
    _signup_mark("publish_status_done")
    
    assignment_rows = []

    # 无论是否发布，都查询真实 assignments；当天补排无需等待发布。
    if role == "值班":

        with get_conn() as conn:
            with conn.cursor(cursor_factory=RealDictCursor) as cur:

                cur.execute("""
                    select
                        assigned_place,
                        shift_label,
                        start_time,
                        end_time
                    from volunteer_schedule_assignments
                    where signup_id = %s
                    order by start_time, id
                """, (signup_id,))

                assignment_rows = cur.fetchall()
    _signup_mark("assignment_lookup_done")
    if assignment_rows:

        assignment_lines = ""

        for a in assignment_rows:
            place = a.get("assigned_place") or ""
            s = a.get("start_time") or ""
            e = a.get("end_time") or ""

            assignment_lines += f"""
            <div class="assignment-line">
                ✅ {place}<br>
                <span>{s} ~ {e}</span>
            </div>
            """

        status_html = f"""
        <div class="alert alert-success">
            🟢 状态：已加入正式值班表<br><br>
            {assignment_lines}
            <br>
            请依照以上岗位值班。若负责人之后有更新，请以最新正式值班表为准。
        </div>
        """

    elif published or (role == "值班" and signup_date == str(malaysia_today())):

        status_html = """
        <div class="alert alert-warning">
            🟡 状态：报名成功，等待系统安排<br><br>
            系统暂时还没有找到适合岗位。<br>
            负责人会查看后再处理，请以最新正式值班表为准。
        </div>
        """

    else:

        status_html = """
        <div class="alert alert-warning">
            🟡 状态：等待负责人安排<br><br>
            ⚠️ 当前属于报名阶段。<br>
            最终岗位安排请以负责人公布的正式值班表为准。<br>
            请多留意义工群信息，感恩大家护持观音堂 🙏
        </div>
        """

    # 报名成功页：复用确认报名的日期资料，不影响已提交的报名。
    signup_weekday = ""
    signup_lunar = ""
    signup_festivals = ""
    try:
        success_date = datetime.strptime(signup_date, "%Y-%m-%d").date()
        signup_weekday = "星期" + "一二三四五六日"[success_date.weekday()]
        success_info = get_special_day_info(success_date)
        lunar = str(success_info.get("lunar_text") or "")
        names = [str(n) for n in (success_info.get("special_names") or []) if n]
        if success_info.get("template_type") == "lunar_1_15" and not names:
            names.append("十五" if "十五" in lunar else "初一" if "初一" in lunar else "初一 / 十五")
        signup_lunar = lunar if lunar.startswith("农历") else ("农历" + lunar if lunar else "")
        signup_festivals = "、".join(names)
    except Exception:
        # 日期附加显示失败时仍正常呈现报名成功页。
        signup_lunar = ""
        signup_festivals = ""

    # =========================
    # 膳食报名成功页
    # =========================

    if role == "膳食":

        with get_conn() as conn:
            with conn.cursor(cursor_factory=RealDictCursor) as cur:
                cur.execute("""
                    select id, name
                    from volunteer_schedule_signups
                    where signup_date = %s
                    and role = '膳食'
                    and coalesce(status, 'pending') <> 'cancelled'
                    order by created_at, id
                """, (signup_date,))

                all_rows = cur.fetchall()

        my_no = None

        for i, row in enumerate(all_rows, start=1):
            if row["id"] == signup_id:
                my_no = i
                break

        my_meal_role = "派餐义工"
        meal_rows = all_rows[:10]
        backup_rows = []

        lines = ""

        for i in range(1, 10):
            person = meal_rows[i - 1]["name"] if i <= len(meal_rows) else ""
            lines += f"<div class='name-line'>{i}）{person}</div>"
        if len(meal_rows) == 10:
            lines += f"<div class='name-line'>10）{meal_rows[9]['name']}</div>"

        backup_lines = ""



        return render_template_string("""
        <!doctype html>
        <html>
        <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">

        <style>
        body {
            font-family:"Microsoft YaHei", Arial;
            background:#f7f2e8;
            padding:20px;
            font-size:22px;
        }

        .card {
            max-width:760px;
            margin:auto;
            background:white;
            border-radius:20px;
            padding:28px;
            box-shadow:0 4px 18px rgba(0,0,0,0.12);
        }

        h1 {
            font-size:34px;
            margin-bottom:10px;
        }

        .big-green {
            font-size:30px;
            color:#168a3a;
            font-weight:bold;
            line-height:1.5;
        }

        .info {
            background:#fff7dc;
            border-radius:16px;
            padding:18px;
            margin:18px 0;
            line-height:1.7;
        }

        .section {
            margin-top:24px;
            padding-top:18px;
            border-top:2px solid #ddd;
        }

        .name-line {
            font-size:24px;
            padding:8px 0;
        }

        .btn {
            display:block;
            margin-top:18px;
            background:#1f9d55;
            color:white;
            text-align:center;
            padding:14px;
            border-radius:14px;
            text-decoration:none;
            font-size:22px;
            font-weight:bold;
        }
        </style>
        </head>

        <body>

        <div class="card">

            <h1>🍱 素食结缘报名成功</h1>

            <div class="info">
                <div><b>日期：</b>{{ signup_date }}　<span style="color:#b91c1c;font-weight:bold;font-size:0.85em;">{{ signup_weekday }}</span>{% if signup_lunar %}<br><span style="font-size:0.85em;color:#8b5a2b;">🌙 {{ signup_lunar }}</span>{% endif %}{% if signup_festivals %}<br><span style="font-size:0.85em;color:#b91c1c;font-weight:bold;">🪷 {{ signup_festivals }}</span>{% endif %}</div>
                <div><b>开放堂食：</b>11:00am</div>
            </div>

            <div class="big-green">
                您是第 {{ my_no }} 位膳食组义工<br>
                您的岗位：{{ my_meal_role }}
            </div>

            <div class="section">
                <h2>🍱 派餐义工（最多9位）</h2>
                {{ lines|safe }}
            </div>


            <div class="section">
                <p>✅ 派餐义工最迟 <b>10:45am</b> 开始站岗</p>
                <p>✅ 所有报名义工需要在 <b>8:30am</b> 前报到</p>
                <p>✅ 活动结束后协助清理场地及餐具</p>
            </div>

            <a class="btn" href="/volunteer">返回义工首页</a>

        </div>

        </body>
        </html>
        """,
        signup_date=signup_date,
        signup_weekday=signup_weekday,
        signup_lunar=signup_lunar,
        signup_festivals=signup_festivals,
        my_no=my_no,
        my_meal_role=my_meal_role,
        lines=lines,
        backup_lines=backup_lines)

    # =========================
    # 普通报名成功页
    # =========================

    return render_template_string("""
    <!doctype html>
    <html>
    <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet"
          href="{{ url_for('static', filename='css/toolbox.css') }}">

    <style>
    body {
        font-family:"Microsoft YaHei";
        background:#f5f5f5;
        padding:20px;
    }

    .box {
        max-width:760px;
        margin:30px auto;
        background:white;
        border-radius:20px;
        padding:30px;
        text-align:center;
    }

    .success {
        background:#e8f5e9;
        border:2px solid #4CAF50;
        border-radius:15px;
        padding:25px;
        margin-bottom:25px;
    }

    .success h1 {
        color:#2e7d32;
        margin-top:0;
        font-size:38px;
    }

    .info {
        font-size:25px;
        line-height:1.9;
    }

    .quote-box {
        margin-top:20px;
        padding:18px;
        background:#fff5f5;
        border:2px solid #efc8c8;
        border-radius:18px;
        text-align:center;
    }

    .quote-title {
        font-size:21px;
        font-weight:bold;
        color:#9b6b00;
    }

    .quote-content {
        font-size:23px;
        line-height:1.8;
        color:#b84a4a;
    }

    .notice {
        border-radius:16px;
        padding:22px;
        margin:22px 0;
        font-size:23px;
        line-height:1.8;
    }

    .yellow-notice {
        background:#fff8e1;
        color:#8d6e00;
    }

    .green-notice {
        background:#e8f5e9;
        color:#1b6e34;
        border:2px solid #7acb8a;
    }

    .assignment-line {
        background:white;
        border-radius:12px;
        padding:12px;
        margin:10px 0;
        font-weight:bold;
    }

    .assignment-line span {
        font-weight:normal;
        color:#555;
    }
                                  
    .daily-quote-card{
        background:#fff8f7;
        border:1px solid #f3c7c3;
        border-radius:18px;
        padding:24px 28px;
        text-align:center;
    }

    .daily-quote-title{
        color:#b56a00;
        font-size:24px;
        font-weight:bold;
        margin-bottom:16px;
    }

    .quote-text{
        white-space: pre-line;
        line-height: 1.9;
        font-size:22px;
        color:#c94840;
        text-align:left;
        display:inline-block;
        margin:auto;
    }
    
    @media (max-width:700px) {
        body {
            padding:12px;
        }

        .box {
            padding:20px;
            margin:15px auto;
        }

        .success h1 {
            font-size:32px;
        }

        .info,
        .notice {
            font-size:21px;
        }

        .btn {
            font-size:20px;
            padding:13px;
        }
    }
    </style>
    </head>

    <body>

    <div class="box">

        <div class="success">

            <h1>{{ result_title }}</h1>

            <div class="info">

                🙏 感恩发心护持观音堂<br><br>

                义工：{{ name }}<br>
                日期：{{ signup_date }}　<span style="color:#b91c1c;font-weight:bold;font-size:0.85em;">{{ signup_weekday }}</span>{% if signup_lunar %}<br><span style="font-size:0.85em;color:#8b5a2b;">🌙 {{ signup_lunar }}</span>{% endif %}{% if signup_festivals %}<br><span style="font-size:0.85em;color:#b91c1c;font-weight:bold;">🪷 {{ signup_festivals }}</span>{% endif %}<br>
                岗位：{{ role }}

                {% if role == "值班" %}
                <br>
                时间：{{ start_time }} ~ {{ end_time }}
                {% endif %}

            </div>

        </div>

        <div class="daily-quote-card">

            <div class="daily-quote-title">
                🌸 每日佛言佛语
            </div>

            <div class="quote-text">{{ daily_quote }}</div>

        </div>

        {{ status_html|safe }}

        <div class="btn-row">

            <a class="btn-tool btn-blue btn-full"
            href="/volunteer/my_schedule_search">
                📋 查看我的报名
            </a>

            <a class="btn-tool btn-orange btn-full"
            href="/volunteer/day_schedule?date={{ signup_date }}">
                📅 查看当天值班表
            </a>

            <a class="btn-tool btn-green btn-full"
            href="/volunteer">
                ➕ 继续报名
            </a>

        </div>

    </div>

    </body>
    </html>
    """,
    result_title=result_title,
    name=name,
    signup_date=signup_date,
    signup_weekday=signup_weekday,
    signup_lunar=signup_lunar,
    signup_festivals=signup_festivals,
    role=role,
    start_time=start_time,
    end_time=end_time,
    daily_quote=daily_quote,
    status_html=status_html)


@schedule_bp.route("/volunteer/query_volunteer")
def query_volunteer_api():

    keyword = _volunteer_identity()
    branch = request.args.get("branch", "CHE").strip()

    if not keyword:
        return jsonify({
            "ok": False,
            "message": "请输入编号 / 姓名 / 电话"
        })
    
    matches = find_volunteer_by_keyword(keyword)

    if not matches:
        return jsonify({
            "ok": False,
            "message": "找不到义工"
        })

    if len(matches) > 1:
        return jsonify({
            "ok": False,
            "message": "找到多个义工，请输入完整编号"
        })

    v = matches[0]

    return jsonify({
        "ok": True,
        "volunteer_id": v.get("member_id") or v.get("id") or v.get("volunteer_id"),
        "name": v.get("name")
    })


@schedule_bp.route("/schedule/meal_extra_tenth", methods=["POST"])
def meal_extra_tenth():
    """负责人特批第10位派餐义工；不修改正常9位报名机制。"""
    if not session.get("schedule_login"):
        return "❌ 请先登录负责人系统", 403
    target = request.form.get("date", "").strip()
    keyword = request.form.get("volunteer_id", "").strip()
    try:
        datetime.strptime(target, "%Y-%m-%d")
    except ValueError:
        return "❌ 日期无效", 400
    if not keyword:
        return "❌ 请输入义工编号", 400
    matches = find_volunteer_by_keyword(keyword)
    if len(matches) != 1:
        return "❌ 请填写唯一的义工正式编号", 400
    vol = matches[0]
    vol_id, name = str(vol["id"]), str(vol["name"])
    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("SELECT pg_advisory_xact_lock(hashtext('meal_signup'), hashtext(%s))", (target,))
            cur.execute("""SELECT COUNT(*) AS cnt FROM volunteer_schedule_signups
                WHERE signup_date=%s AND role='膳食'
                AND coalesce(status,'pending') <> 'cancelled'""", (target,))
            count = int(cur.fetchone()["cnt"] or 0)
            if count != 9:
                return "❌ 仅已有9位时才能特批第10位；当前人数：" + str(count), 409
            cur.execute("""SELECT 1 FROM volunteer_schedule_signups
                WHERE signup_date=%s AND role='膳食' AND volunteer_id=%s
                AND coalesce(status,'pending') <> 'cancelled' LIMIT 1""", (target, vol_id))
            if cur.fetchone():
                return "❌ 该义工已报名", 409
            cur.execute("""INSERT INTO volunteer_schedule_signups
                (volunteer_id,name,signup_date,role,start_time,end_time,status,remarks,meal_role)
                VALUES (%s,%s,%s,'膳食','8:00am','2:00pm','pending','负责人特批第10位','派餐义工')""",
                (vol_id,name,target))
        conn.commit()
    return redirect('/volunteer/meal_status?date=' + target)


@schedule_bp.route("/volunteer/meal_status")
def volunteer_meal_status():

    signup_date = request.args.get("date") or malaysia_today().strftime("%Y-%m-%d")

    date_obj = datetime.strptime(signup_date, "%Y-%m-%d").date()
    special_info = get_special_day_info(date_obj)

    lunar_text = special_info.get("lunar_text") or special_info.get("lunar_date") or "农历初一 / 十五"
    festival_name = special_info.get("festival_name") or special_info.get("name") or ""

    page_title = f"🍱 {lunar_text} · 素食结缘"
    festival_line = f"🙏 {festival_name}" if festival_name else ""

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
            select id,name
            from volunteer_schedule_signups
            where signup_date=%s
            and role='膳食'
            and coalesce(status,'pending')<>'cancelled'
            order by created_at,id
            """,(signup_date,))

            rows = cur.fetchall()
            print("MEAL STATUS ROWS:", rows)
            

    meal_names = [r["name"] for r in rows[:10]]
    backup_names = []

    lines = ""
    for i in range(1, 10):
        person = meal_names[i - 1] if i <= len(meal_names) else ""
        lines += f"<div class='name-line'>{i}）{person}</div>"

    # 第10位若由负责人特批，作为正式派餐义工列出。
    if len(meal_names) == 10:
        lines += f"<div class='name-line'>10）{meal_names[9]}</div>"
    backup_lines = ""

    meal_count = len(meal_names)
    backup_count = len(backup_names)

    return render_template_string("""
    <style>
    body {
        font-family: Arial, "Microsoft YaHei", sans-serif;
        background:#f7f2e8;
        padding:20px;
        font-size:24px;
    }

    .card {
        max-width:820px;
        margin:auto;
        background:white;
        border-radius:22px;
        padding:30px;
        box-shadow:0 4px 18px rgba(0,0,0,0.12);
    }

    h1 {
        font-size:36px;
        color:#8b5a2b;
        margin-top:0;
        line-height:1.4;
    }

    .info-box {
        background:#fff8e6;
        border:2px solid #f3d9a5;
        border-radius:18px;
        padding:18px;
        margin:20px 0;
        font-size:24px;
        line-height:1.8;
        color:#5c3b1e;
    }

    .status-box {
        background:#eaf8ef;
        border:2px solid #b8e0c4;
        border-radius:18px;
        padding:18px;
        margin:20px 0;
        font-size:28px;
        font-weight:bold;
        color:#167a35;
        text-align:center;
    }

    .section {
        margin-top:24px;
        padding-top:18px;
        border-top:2px solid #eee;
    }

    h2 {
        font-size:30px;
        color:#5c3b1e;
    }

    .name-line {
        font-size:28px;
        padding:10px 0;
        border-bottom:1px dashed #eee;
    }

    .empty {
        color:#999;
        font-size:24px;
        padding:10px 0;
    }

    .note {
        background:#fff3f0;
        border:2px solid #f5c2bd;
        border-radius:16px;
        padding:16px;
        margin-top:20px;
        font-size:22px;
        line-height:1.7;
        color:#7a2e24;
    }

    .btn {
        display:block;
        margin-top:24px;
        background:#1f9d55;
        color:white;
        text-align:center;
        padding:18px;
        border-radius:14px;
        text-decoration:none;
        font-size:26px;
        font-weight:bold;
    }

    @media (max-width:700px) {
        body {
            padding:10px;
        }

        .card {
            padding:22px;
        }

        h1 {
            font-size:30px;
        }

        .name-line {
            font-size:25px;
        }
    }
    </style>

    <div class="card">

        <h1>{{ page_title }}</h1>

        <div class="info-box">
            <div><b>📅 日期：</b>{{ signup_date }}</div>
            <div><b>🌙 农历：</b>{{ lunar_text }}</div>

            {% if festival_line %}
            <div><b>🙏 佛诞：</b>{{ festival_line }}</div>
            {% endif %}

            <div><b>🍱 开放堂食：</b>11:00am</div>
        </div>

        <div class="status-box">
            派餐义工：{{ meal_count }} 位（通常9位）
        </div>

        <div class="section">
            <h2>🍱 派餐义工（最多9位）</h2>
            {{ lines|safe }}
        </div>


        <div class="note">
            ⚠️ 派餐义工最多 9 位。<br>
            <br>
            ✅ 派餐义工最迟 10:45am 开始站岗。<br>
            ✅ 所有报名义工需要在 8:30am 前报到。<br>
            ✅ 活动结束后协助清理场地及餐具。
        </div>

        {% if admin_logged_in and meal_count == 9 %}
        <div class="section">
            <h2>🔐 负责人特批第10位</h2>
            <form method="post" action="/schedule/meal_extra_tenth">
                <input type="hidden" name="date" value="{{ signup_date }}">
                <input name="volunteer_id" required placeholder="输入义工正式编号" style="width:70%;padding:12px;font-size:20px;">
                <button type="submit" style="padding:12px;font-size:20px;">特批加入第10位</button>
            </form>
        </div>
        {% endif %}
        <a class="btn" href="/volunteer">返回义工首页</a>

    </div>
    """,
    page_title=page_title,
    signup_date=signup_date,
    lunar_text=lunar_text,
    festival_line=festival_line,
    meal_count=meal_count,
    backup_count=backup_count,
    lines=lines,
    backup_lines=backup_lines,
    admin_logged_in=bool(session.get("schedule_login")))


@schedule_bp.route("/schedule/meal_leader/add", methods=["POST"])
def add_meal_leader():

    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    signup_date = request.form.get("date")
    keyword = request.form.get("volunteer_id", "").strip()

    matches = find_volunteer_by_keyword(keyword)

    if not matches:
        return "❌ 找不到义工<br><a href='/schedule'>返回</a>"

    if len(matches) > 1:
        return "❌ 找到多个义工，请用义工编号报名<br><a href='/schedule'>返回</a>"

    vol = matches[0]
    vol_id = str(vol["id"])
    name = vol["name"]

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:

            cur.execute("""
                insert into volunteer_schedule_signups
                (volunteer_id, name, signup_date, role, start_time, end_time, status, remarks, meal_role)
                values (%s, %s, %s, '膳食', '8:00am', '2:00pm', 'pending', '负责人加入膳食组长', '组长')
            """, (
                vol_id,
                name,
                signup_date
            ))

        conn.commit()

    return redirect(f"/schedule/notice_center?date={signup_date}")


@schedule_bp.route("/volunteer/my_schedule")
def volunteer_my_schedule():
    keyword = _volunteer_identity()

    if not keyword:
        return "❌ 请输入义工编号 / 姓名 / 电话<br><a href='/volunteer'>返回</a>"

    matches = find_volunteer_by_keyword(keyword)

    if not matches:
        return "❌ 找不到义工<br><a href='/volunteer'>返回</a>"

    if len(matches) > 1:
        return "❌ 找到多个同名义工，请用义工编号查询<br><a href='/volunteer'>返回</a>"

    vol = matches[0]
    vol_id = str(vol["id"])
    name = str(vol["name"])

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                select *
                from volunteer_schedule_signups
                where volunteer_id = %s
                and coalesce(status, 'pending') <> 'cancelled'
                order by signup_date DESC, start_time
            """, (vol_id,))
            rows = cur.fetchall()

            # 正式排班必须以 assignments 为准，不能只看 signup.status / assigned_place
            cur.execute("""
                select
                    signup_id,
                    role,
                    shift_label,
                    assigned_place,
                    start_time,
                    end_time
                from volunteer_schedule_assignments
                where volunteer_id = %s
                and coalesce(status, 'assigned') <> 'cancelled'
                order by assignment_date, start_time, id
            """, (vol_id,))

            assignment_rows = cur.fetchall()

    assignments_by_signup = {}
    for a in assignment_rows:
        sid = a.get("signup_id")
        if sid is None:
            continue
        assignments_by_signup.setdefault(int(sid), []).append(a)

    html = f"""
    <!doctype html>
    <html>
    <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <style>
    body {{
        font-family:"Microsoft YaHei";
        background:#f5f5f5;
        padding:18px;
        font-size:26px;
    }}

    .box {{
        max-width:800px;
        margin:auto;
    }}

    .card {{
        background:white;
        border-radius:18px;
        padding:24px;
        margin:18px 0;
        box-shadow:0 3px 12px rgba(0,0,0,0.08);
        font-size:28px;
        line-height:1.8;
    }}

    .status {{
        background:#fff3cd;
        color:#8a5a00;
        padding:10px 14px;
        border-radius:10px;
        display:inline-block;
        font-weight:bold;
    }}

    .big-btn {{
        width:100%;
        border:0;
        border-radius:15px;
        padding:22px;
        margin-top:14px;
        font-size:30px;
        font-weight:bold;
        color:white;
    }}

    .edit-btn {{
        background:#4CAF50;
    }}

    .cancel-btn {{
        background:#d9534f;
    }}

    .back-btn {{
        display:block;
        background:#607D8B;
        color:white;
        text-align:center;
        text-decoration:none;
        padding:18px;
        border-radius:15px;
        font-size:28px;
        font-weight:bold;
        margin:18px 0;
    }}
    </style>
    </head>
    <body>
    <div class="box">

    <h1>📋 我的报名记录</h1>
    <h2>{name}</h2>

    <a class="back-btn" href="/volunteer">⬅ 返回报名</a>
    """

    if not rows:
        html += """
        <div class="card">
            暂时没有报名记录。
        </div>
        """
    else:
        for r in rows:
            signup_id = int(r["id"])
            formal_rows = assignments_by_signup.get(signup_id, [])
            status = str(r.get("status") or "pending")

            # ===============================================
            # 我的报名：正式安排一律读取 assignments
            # ===============================================
            if formal_rows:
                status_text = "🟢 已正式安排"

                assignment_html = ""
                for a in formal_rows:
                    shift = a.get("shift_label") or ""
                    place = a.get("assigned_place") or ""
                    a_start = a.get("start_time") or ""
                    a_end = a.get("end_time") or ""

                    assignment_html += f"""
                    <div style="
                        margin-top:10px;
                        padding:12px 14px;
                        background:#e8f5e9;
                        border-radius:10px;
                        line-height:1.5;
                    ">
                        📍 <b>{place}</b><br>
                        🕐 {a_start} ~ {a_end}
                    </div>
                    """

                system_arrangement_html = f"""
                    <div style="margin-top:8px;">
                        <b>系统正式安排：</b>
                        {assignment_html}
                    </div>
                """

            else:
                # 没有 assignment 才是真正未安排
                if status == "pending":
                    status_text = "🟡 等待正式安排"
                elif status == "assigned":
                    # 防止 signup.status 已是 assigned，但 assignment 实际不存在
                    status_text = "🟠 正式安排尚未建立"
                else:
                    status_text = status

                system_arrangement_html = """
                    <div style="margin-top:8px;">
                        <b>系统正式安排：</b>尚未安排
                    </div>
                """

            html += f"""
            <div class="card">
                📅 <b>{r["signup_date"]}</b><br>
                岗位：<b>{r["role"]}</b><br>
                报名时间：{r["start_time"]} ~ {r["end_time"]}<br>
                {system_arrangement_html}
                状态：<span class="status">{status_text}</span>
            """

            signup_date_obj = r["signup_date"]

            if str(signup_date_obj) != str(malaysia_today()) and not formal_rows:
                html += f"""
                <form method="get" action="/volunteer/edit_signup/{r['id']}">
                    <button class="big-btn edit-btn" type="submit">
                        ✏️ 修改这次报名
                    </button>
                </form>

                <form method="post"
                    action="/volunteer/cancel/{r['id']}"
                    onsubmit="return confirm('确定不参加这次报名了吗？');">
                    <button class="big-btn cancel-btn" type="submit">
                        ❌ 不参加了，取消报名
                    </button>
                </form>
                """
            elif formal_rows:
                html += """
                <p style="color:#9b5b00; font-size:24px;">
                    🔒 已正式安排，如需修改时间或请假，请联系负责人。<br>
                    Already assigned. Please contact the coordinator.
                </p>
                """
            else:
                html += """
                <p style="color:#b36b00; font-size:26px;">
                    🚫 今天值班如需更改或取消，请务必通知负责人。
                </p>
                """

            html += """
            </div>
            """

    html += """
    </div>
    </body>
    </html>
    """

    return html


@schedule_bp.route("/volunteer/guide")
def volunteer_guide():
    return render_template_string("""
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>义工须知</title>

<link rel="stylesheet"
      href="{{ url_for('static', filename='css/toolbox.css') }}">
<style>
.guide-wrap{
    max-width:900px;
    margin:auto;
    padding:24px;
}

.guide-content{
    font-size:26px;
    line-height:1.8;
}

.guide-content h2{
    margin-top:30px;
    font-size:32px;
}
</style>
</head>

<body>

<div class="guide-wrap">

    <div class="card">

        <h1 class="page-title">📖 义工须知</h1>

        <p class="page-subtitle">
            感恩您发心护持观音堂，请先了解报名与值班规则。
        </p>

        <div class="guide-content">

            <h2>🙏 感恩发心护持道场</h2>
            <p>
                感恩您发心参与观音堂义工服务。每一位义工都是护持道场的重要力量，
                愿大家互相配合，共同成就道场。
            </p>

            <h2>📝 关于报名</h2>
            <p>
                报名后，系统会根据当天需要自动安排岗位。负责人也可能因现场情况调整岗位或时间。
                报名成功并不代表最终岗位。
            </p>

            <h2>📢 正式值班表</h2>
            <p>
                每天约 <b>10:00pm</b> 发布正式值班表。请以最新正式值班表为准。
                发布后如有调整，请以负责人通知为准。
            </p>

            <h2>✏️ 修改或取消报名</h2>
            <p>
                尚未正式安排岗位时，可自行修改或取消未来日期的报名；一旦正式安排岗位，请联络负责人处理。
            </p>

            <h2>🚫 值班当天无法出席</h2>
            <p>
                若当天因突发情况无法值班，请尽快通知负责人，以便安排其他义工补位。
                请不要直接缺席，以免影响当天道场运作。
            </p>

            <h2>⏰ 请准时签到</h2>
            <p>
                到达观音堂后，请先签到。签到后请依照系统显示的岗位值班。
                如负责人现场有调整，请以负责人安排为准。
            </p>

            <h2>🤝 值班期间</h2>
            <p>
                请礼貌待人，配合负责人安排。如有任何疑问，可随时向负责人请教。
                请穿着整齐仪容：有义工服者请穿义工服；尚未领取义工服者，请穿红色上衣及黑色长裤。
            </p>

            <h2>❤️ 感恩大家</h2>
            <p>
                愿大家发心护持道场，广结善缘，福慧增长。
                🙏 感恩每一位义工。
            </p>

        </div>

        <div class="btn-row">
            <a class="btn-tool btn-gray btn-full" href="/volunteer">
                ⬅ 返回义工首页
            </a>
        </div>

    </div>

</div>

</body>
</html>
""")


@schedule_bp.route("/volunteer/edit_signup/<int:signup_id>", methods=["GET"])
def volunteer_edit_signup(signup_id):

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                select *
                from volunteer_schedule_signups
                where id = %s
                and coalesce(status, 'pending') <> 'cancelled'
            """, (signup_id,))
            row = cur.fetchone()

    if not row:
        return "❌ 找不到这笔报名，或已取消<br><a href='/volunteer'>返回</a>"

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            if signup_has_formal_assignment(cur, signup_id):
                return assigned_signup_locked_message(row['volunteer_id'])

    return render_template_string("""
    <!doctype html>
    <html>
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>修改报名</title>
        <link rel="stylesheet"
              href="{{ url_for('static', filename='css/toolbox.css') }}">
    </head>

    <body>

    <div class="page-wrap">

        <div class="card">

            <div class="result-icon" style="text-align:center;">✏️</div>
            <div class="result-title" style="text-align:center;">修改报名</div>

            <div class="name" style="text-align:center;font-size:32px;font-weight:bold;margin-bottom:20px;">
                {{ row.name }}
            </div>

            <form method="post"
                action="/volunteer/edit_signup/{{ row.id }}"
                onsubmit="return confirm('确定修改这次报名？');">

                <div class="form-grid">

                    <div class="form-group">
                        <label class="form-label">📅 日期</label>

                        <input
                            class="form-input"
                            type="date"
                            name="signup_date"
                            value="{{ row.signup_date }}"
                            required>
                    </div>

                    <div class="form-group">
                        <label class="form-label">👷 岗位</label>

                        <select
                            class="form-select"
                            name="role">

                            <option value="值班"
                                {% if row.role=="值班" %}selected{% endif %}>
                                值班
                            </option>

                            <option value="卫生"
                                {% if row.role=="卫生" %}selected{% endif %}>
                                卫生
                            </option>

                            <option value="供台"
                                {% if row.role=="供台" %}selected{% endif %}>
                                供台
                            </option>

                        </select>
                    </div>

                    <div class="form-group">
                        <label class="form-label">🕙 开始时间</label>

                        <input
                            class="form-input"
                            name="start_time"
                            value="{{ row.start_time }}">
                    </div>

                    <div class="form-group">
                        <label class="form-label">🕒 结束时间</label>

                        <input
                            class="form-input"
                            name="end_time"
                            value="{{ row.end_time }}">
                    </div>

                </div>

                <div class="btn-row">

                    <button
                        class="btn-tool btn-green btn-full"
                        type="submit">

                        ✅ 确认修改报名

                    </button>

                    <a
                        class="btn-tool btn-gray btn-full"
                        href="/volunteer/my_schedule?keyword={{ row.volunteer_id }}">

                        ⬅ 返回我的报名

                    </a>

                </div>

            </form>

        </div>

    </div>

    </body>
    </html>
    """, row=row)


@schedule_bp.route("/volunteer/edit_signup/<int:signup_id>", methods=["POST"])
def volunteer_edit_signup_post(signup_id):

    new_date = request.form.get("signup_date", "").strip()
    new_role = request.form.get("role", "").strip()
    new_start = request.form.get("start_time", "").strip()
    new_end = request.form.get("end_time", "").strip()

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:

            cur.execute("""
                select *
                from volunteer_schedule_signups
                where id = %s
                and coalesce(status, 'pending') <> 'cancelled'
                FOR UPDATE
            """, (signup_id,))
            old = cur.fetchone()

            if not old:

                return render_template_string("""
                <!DOCTYPE html>
                <html>
                <head>
                    <meta charset="utf-8">
                    <title>找不到报名记录</title>
                    <link rel="stylesheet"
                          href="{{ url_for('static', filename='css/toolbox.css') }}">
                </head>
                <body>

                <div class="page-wrap">

                    <div class="card result-card">

                        <div class="result-icon">❌</div>

                        <div class="result-title">
                            找不到报名记录
                        </div>

                        <p>
                            这笔报名不存在，或已经取消，无法修改。
                        </p>

                        <div class="btn-row">
                            <a class="btn-tool btn-gray" href="/volunteer">
                                返回首页
                            </a>
                        </div>

                    </div>

                </div>

                </body>
                </html>
                """)
            
            if signup_has_formal_assignment(cur, signup_id):
                return assigned_signup_locked_message(old['volunteer_id'])

            today = malaysia_today()

            if str(old["signup_date"]) == str(today):

                return render_template_string("""
                <!DOCTYPE html>
                <html>
                <head>
                    <meta charset="utf-8">
                    <title>今天不能修改报名</title>
                    <link rel="stylesheet"
                          href="{{ url_for('static', filename='css/toolbox.css') }}">
                </head>
                <body>

                <div class="page-wrap">

                    <div class="card result-card">

                        <div class="result-icon">🚫</div>

                        <div class="result-title">
                            今天不能修改报名
                        </div>

                        <div class="info-box">

                            <div class="info-row">
                                <span class="info-label">姓名</span>
                                <span class="info-value">{{ old.name }}</span>
                            </div>

                            <div class="info-row">
                                <span class="info-label">日期</span>
                                <span class="info-value">{{ old.signup_date }}</span>
                            </div>

                            <div class="info-row">
                                <span class="info-label">岗位</span>
                                <span class="info-value">{{ old.role }}</span>
                            </div>

                        </div>

                        <p class="text-red" style="font-weight:bold;">
                            今天已经是值班当天，如需修改，请通知负责人处理。
                        </p>

                        <div class="btn-row">
                            <a class="btn-tool btn-gray"
                            href="/volunteer/my_schedule?keyword={{ old.volunteer_id }}">
                                返回我的报名
                            </a>
                        </div>

                    </div>

                </div>

                </body>
                </html>
                """, old=old)

            if new_role == "卫生":
                try:
                    target_cleaning_date = datetime.strptime(new_date, "%Y-%m-%d").date()
                except ValueError:
                    return "❌ 日期格式错误 / Invalid date"
                cleaning_open, cutoff_hour, _ = cleaning_signup_window(target_cleaning_date)
                if not cleaning_open:
                    return cleaning_signup_closed_message(cutoff_hour)

            if new_role == "卫生":
                cleaning_count = lock_and_count_cleaning(cur, new_date)
                same_slot = (old["role"] == "卫生" and str(old["signup_date"]) == new_date)
                if cleaning_count >= 3 and not same_slot:
                    return cleaning_full_message()

            if new_role == "值班":
                conflict = find_duty_time_conflict(cur, old['volunteer_id'], new_date,
                                                   new_start, new_end, exclude_id=signup_id)
                if conflict:
                    return duty_conflict_message(conflict)

            cur.execute("""
                update volunteer_schedule_signups
                set status = 'cancelled'
                where id = %s
            """, (signup_id,))

            cur.execute("""
                insert into volunteer_schedule_signups
                    (volunteer_id, name, signup_date, role, start_time, end_time, status, remarks)
                values
                    (%s, %s, %s, %s, %s, %s, 'pending', %s)
                returning id
            """, (
                old["volunteer_id"],
                old["name"],
                new_date,
                new_role,
                new_start,
                new_end,
                f"由报名 #{signup_id} 修改而来"
            ))

            new_signup = cur.fetchone()
            new_signup_id = new_signup["id"]

        conn.commit()

    sync_schedule_after_signup_change(
        signup_id,
        action="cancel",
        changed_by="volunteer"
    )

    sync_schedule_after_signup_change(
        new_signup_id,
        action="upsert",
        changed_by="volunteer"
    )
    
    return redirect(f"/volunteer/my_schedule?keyword={old['volunteer_id']}")

@schedule_bp.route(
    "/volunteer/cancel/<int:signup_id>",
    methods=["POST"]
)
def volunteer_cancel_signup(signup_id):

    today = malaysia_today()

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:

            cur.execute("""
                select
                    id,
                    volunteer_id,
                    name,
                    signup_date,
                    role,
                    coalesce(status, 'pending') as status
                from volunteer_schedule_signups
                where id = %s
                FOR UPDATE
            """, (signup_id,))

            row = cur.fetchone()

            if not row:

                return render_template_string("""
                <!DOCTYPE html>
                <html>
                <head>
                    <meta charset="utf-8">
                    <title>找不到报名记录</title>
                    <link rel="stylesheet"
                          href="{{ url_for('static', filename='css/toolbox.css') }}">
                </head>
                <body>

                <div class="page-wrap">

                    <div class="card result-card">

                        <div class="result-icon">❌</div>

                        <div class="result-title">
                            找不到报名记录
                        </div>

                        <p>
                            这笔报名记录不存在，可能已经删除或取消。
                        </p>

                        <div class="btn-row">

                            <a class="btn-tool btn-gray"
                            href="/volunteer">
                                返回首页
                            </a>

                        </div>

                    </div>

                </div>

                </body>
                </html>
                """)

            if signup_has_formal_assignment(cur, signup_id):
                return assigned_signup_locked_message(row['volunteer_id'])

            signup_date = row["signup_date"]

            if str(signup_date) == str(today):

                return render_template_string("""
                <!DOCTYPE html>
                <html>
                <head>
                    <meta charset="utf-8">
                    <title>今天不能取消报名</title>
                    <link rel="stylesheet"
                          href="{{ url_for('static', filename='css/toolbox.css') }}">
                </head>
                <body>

                <div class="page-wrap">

                    <div class="card result-card">

                        <div class="result-icon">🚫</div>

                        <div class="result-title">
                            今天不能取消报名
                        </div>

                        <div class="info-box">

                            <div class="info-row">
                                <span class="info-label">姓名</span>
                                <span class="info-value">{{ row.name }}</span>
                            </div>

                            <div class="info-row">
                                <span class="info-label">日期</span>
                                <span class="info-value">{{ row.signup_date }}</span>
                            </div>

                            <div class="info-row">
                                <span class="info-label">岗位</span>
                                <span class="info-value">{{ row.role }}</span>
                            </div>

                        </div>

                        <p style="color:#dc2626;font-weight:bold;">
                            今天已经是值班当天，如需取消，请通知负责人处理。
                        </p>

                        <div class="btn-row" style="justify-content:center;">

                            <a class="btn-tool btn-gray"
                            href="/volunteer/my_schedule?keyword={{ row.volunteer_id }}">
                                返回我的报名
                            </a>

                        </div>

                    </div>

                </div>

                </body>
                </html>
                """, row=row)

            cur.execute("""
                update volunteer_schedule_signups
                set
                    status = 'cancelled',
                    assigned_place = null,
                    remarks = '义工自行取消报名'
                where id = %s
            """, (signup_id,))

            conn.commit()

    sync_schedule_after_signup_change(
        signup_id,
        action="cancel",
        changed_by="volunteer"
    )

    return render_template_string("""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <title>已取消报名</title>
        <link rel="stylesheet"
              href="{{ url_for('static', filename='css/toolbox.css') }}">
    </head>
    <body>

    <div class="page-wrap">

        <div class="card result-card">

            <div class="result-icon">✅</div>
            <div class="result-title">已取消报名</div>

            <div class="info-box">

                <div class="info-row">
                    <span class="info-label">姓名</span>
                    <span class="info-value">{{ row.name }}</span>
                </div>

                <div class="info-row">
                    <span class="info-label">日期</span>
                    <span class="info-value">{{ row.signup_date }}</span>
                </div>

                <div class="info-row">
                    <span class="info-label">岗位</span>
                    <span class="info-value">{{ row.role }}</span>
                </div>

                <div class="info-row">
                    <span class="info-label">状态</span>
                    <span class="info-value">已取消</span>
                </div>

            </div>

            <p>你的报名已经取消，系统会自动同步正式值班表。</p>

            <div class="btn-row" style="justify-content:center;">
                <a class="btn-tool btn-blue"
                href="/volunteer/my_schedule?keyword={{ row.volunteer_id }}">
                    返回我的报名
                </a>
            </div>

        </div>

    </div>

    </body>
    </html>
    """, row=row)

@schedule_bp.route("/volunteer/today_schedule")
def volunteer_today_schedule():

    now = malaysia_now()

    if now.hour >= 19:
        target_date = (now + timedelta(days=1)).strftime("%Y-%m-%d")
    else:
        target_date = now.strftime("%Y-%m-%d")

    return redirect(url_for(
        "schedule.volunteer_day_schedule",
        date=target_date
    ))

@schedule_bp.route("/volunteer/day_schedule")
def volunteer_day_schedule():
    """义工只读电子通告；复用现有正式值班表来源，不修改排班或发送。"""
    from html import escape

    # 未指定日期时，19:00 后默认查看明天；明确选择的日期保持不变。
    now = malaysia_now()
    default_schedule_date = (now.date() + timedelta(days=1)) if now.hour >= 19 else now.date()
    signup_date = request.args.get("date", "").strip() or str(default_schedule_date)
    try:
        selected_day = datetime.strptime(signup_date, "%Y-%m-%d").date()
        if selected_day.year < 2020 or selected_day.year > 2100:
            raise ValueError("date out of range")
    except ValueError:
        return "日期格式不正确。<a href='/volunteer'>返回</a>", 400

    published = is_schedule_published(signup_date)
    entries = []
    if published:
        try:
            message = build_whatsapp_from_assigned(signup_date)
            # 仅在本页面合并同一人、同一段落内连续的值班时段；不修改正式排班原文。
            import re
            interval_re = re.compile(
                r"^\s*(\d{1,2}:\d{2}\s*(?:am|pm))\s*[~～至-]+\s*"
                r"(\d{1,2}:\d{2}\s*(?:am|pm))\s*$", re.I
            )

            def minutes_of(clock):
                clock = clock.strip().lower().replace(" ", "")
                hour, rest = clock.split(":", 1)
                minute = int(rest[:2])
                h = int(hour) % 12 + (12 if rest[2:] == "pm" else 0)
                return h * 60 + minute

            def merge_display_lines(lines):
                # 仅合并相邻的「姓名 / 时间 / 同姓名 / 时间」；不同地点分段不合并。
                merged = []
                i = 0
                while i < len(lines):
                    if i + 1 < len(lines) and interval_re.fullmatch(lines[i + 1]):
                        name = lines[i]
                        first = interval_re.fullmatch(lines[i + 1])
                        start, end = first.group(1), first.group(2)
                        end_minutes = minutes_of(end)
                        j = i + 2
                        while (j + 1 < len(lines) and lines[j] == name
                               and (next_match := interval_re.fullmatch(lines[j + 1]))):
                            next_start = minutes_of(next_match.group(1))
                            next_end = minutes_of(next_match.group(2))
                            if next_start > end_minutes or next_end < end_minutes:
                                break
                            end, end_minutes = next_match.group(2), next_end
                            j += 2
                        merged.extend((name, f"{start}~{end}" if j > i + 2 else lines[i + 1]))
                        i = j
                    else:
                        merged.append(lines[i])
                        i += 1
                return merged

            for paragraph in message.replace("\r\n", "\n").split("\n\n"):
                lines = [line.strip() for line in paragraph.splitlines() if line.strip()]
                if lines:
                    entries.append(merge_display_lines(lines))
        except Exception:
            return "暂时无法读取正式值班表，请稍后再试。<a href='/volunteer'>返回</a>", 503
    else:
        with get_conn() as conn:
            with conn.cursor(cursor_factory=RealDictCursor) as cur:
                cur.execute("""
                    SELECT name, role, start_time, end_time, supply_task
                    FROM volunteer_schedule_signups
                    WHERE signup_date = %s
                      AND coalesce(status, 'pending') <> 'cancelled'
                    ORDER BY role, start_time, name
                """, (signup_date,))
                signups = cur.fetchall()
        for role in ("值班", "卫生", "供台", "膳食"):
            group = [r for r in signups if r.get("role") == role]
            if group:
                lines = [role]
                for r in group:
                    name = str(r.get("name") or "")
                    start = str(r.get("start_time") or "")
                    end = str(r.get("end_time") or "")
                    task = str(r.get("supply_task") or "")
                    if role == "供台":
                        task = {"setup": "设供台", "remove": "收供台"}.get(task, task)
                    lines.append(" ".join(x for x in (name, task, f"{start}～{end}" if start and end else "") if x))
                entries.append(lines)
        if not entries:
            entries = [["目前尚无报名资料"]]

    # 单张通告：完整保留正式 WhatsApp 文案的换行与空行；
    # 只消除农历行开头的多余缩进，不更改原始 WhatsApp 文案。
    if published:
        display_text = message.replace("\r\n", "\n")
        display_text = "\n".join(
            line.lstrip() if line.lstrip().startswith(("（农历", "(农历")) else line
            for line in display_text.split("\n")
        )
    else:
        display_text = "\n\n".join("\n".join(str(x) for x in group) for group in entries)

    # 仅义工网页隐藏报名链接提示；不改 WhatsApp 正式文案。
    if published:
        import re as _web_re
        display_text = _web_re.sub(
            r"(?m)^[ \t]*(?:📝[ \t]*)?义工报名请点击以下链接[^\n]*\n"
            r"[ \t]*https?://[^\s]+/volunteer[ \t]*\n?",
            "", display_text,
        )

    # 只影响义工网页显示：文字原样保留，按岗位给标题柔和着色。
    import html
    from markupsafe import Markup
    import re as _re
    _role_colors = {
        "卫生": "clean", "打扫卫生": "clean", "佛台": "altar", "整理佛台": "altar",
        "供花": "flower", "供果": "fruit", "膳食": "meal", "派餐": "meal",
        "供台": "offering", "设供台": "offering", "收供台": "offering",
        "佛学班": "classroom", "值班": "duty",
    }
    def _line_class(line):
        stripped = line.strip().replace("［", "[").replace("］", "]").replace("〔", "[").replace("〕", "]")
        if not stripped:
            return ""
        if "观音堂:" in stripped or "观音堂：" in stripped:
            return "place-temple"
        if "活动中心:" in stripped or "活动中心：" in stripped:
            return "place-center"
        for label, cls in _role_colors.items():
            if stripped.startswith((f"[{label}]", f"【{label}】", f"[{label}：", f"【{label}：")):
                return "role-" + cls
        if stripped in _role_colors:
            return "role-" + _role_colors[stripped]
        for label, cls in _role_colors.items():
            if stripped.startswith((f"[{label}", f"【{label}", f"{label}：", f"{label}:")):
                return "role-" + cls
        return ""
    styled_lines = []
    for _line in display_text.split("\n"):
        _cls = _line_class(_line)
        _safe = html.escape(_line)
        styled_lines.append(f'<span class="{_cls}">{_safe}</span>' if _cls else _safe)
    styled_text = Markup("\n".join(styled_lines))
    # 星期一到星期日固定轮换；以马来西亚当天为准，不影响查看的排班日期。
    background_index = malaysia_today().weekday() + 1

    return render_template_string("""
    <!doctype html><html lang="zh"><head><meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>义工值班安排</title>
    <style>
      *{box-sizing:border-box}
      body{margin:0;min-height:100vh;color:#344037;
        font-family:'Microsoft YaHei','Noto Sans SC',Arial,sans-serif;
        background:#f8f2e7}
      body::before{content:"";position:fixed;inset:0;z-index:-1;pointer-events:none;
        background:#f8f2e7}
      .page{max-width:900px;margin:24px auto;background:rgba(255,251,244,.88);
        padding:32px 42px 90px;border:1px solid #eee1cb;
        border-radius:19px;box-shadow:0 8px 28px #6e593014}
      h1{text-align:center;font-size:30px;color:#77562d;margin:4px 0 6px}
      .sub{text-align:center;color:#82745f;margin:0 0 23px}
      .toolbar{display:flex;align-items:center;justify-content:center;gap:10px;
        flex-wrap:wrap;margin:0 0 18px}
      input[type=date]{font-size:17px;padding:10px;border:1px solid #d7c6aa;border-radius:9px}
      button,.link{border:0;border-radius:9px;background:#39845e;color:white;
        padding:11px 17px;font-size:17px;text-decoration:none;cursor:pointer}
      .status{text-align:center;border-radius:12px;padding:12px;margin:0 0 22px;
        background:{{ '#e7f1e7' if published else '#fff0d5' }};
        color:{{ '#237344' if published else '#93620b' }};font-weight:bold}
      .notice-sheet{white-space:pre-wrap;overflow-wrap:anywhere;
        font-size:20px;line-height:1.72;margin:0;padding:22px 24px;
        border-radius:15px;border:1px solid #ecdfcb;position:relative;isolation:isolate;
        background:#fff9ee;font-family:inherit;color:#344037}
      .notice-sheet::before{content:"";position:absolute;inset:0;z-index:-1;
        border-radius:inherit;pointer-events:none;
        background-image:linear-gradient(rgba(255,249,238,.93),rgba(255,249,238,.93)),
          url("{{ url_for('static',filename='images/duty_bg_' ~ background_index ~ '.jpg') }}");
        background-position:center top;background-size:100% auto;background-repeat:no-repeat}
      .notice-sheet span[class^="role-"]{font-weight:800;font-size:1.17em;line-height:1.9}
      .role-clean{color:#087966}.role-altar{color:#6b3fa0}
      .role-flower{color:#b64e80}.role-fruit{color:#46834a}
      .role-meal{color:#a6652a}.role-offering{color:#7553a0}
      .role-classroom{color:#326f9d}.role-duty{color:#b16b30}
      .place-temple{color:#b55c20;font-weight:800;font-size:1.08em}
      .place-center{color:#927100;font-weight:800;font-size:1.08em}
      .footer{position:fixed;right:20px;bottom:20px;z-index:1000;margin:0;display:flex}
      .secondary{background:#687a70;box-shadow:0 4px 14px #0003}
      .blessing{text-align:center;color:#927b5d;font-size:15px;margin:22px 0 0}
      @media(max-width:600px){.page{margin:0;padding:22px 12px 90px;border:0;
        border-radius:0;box-shadow:none;background:rgba(255,251,244,.82)}
        h1{font-size:25px}.notice-sheet{font-size:18px;padding:16px 14px}
        .footer{right:12px;bottom:calc(12px + env(safe-area-inset-bottom))}
        .footer .link{font-size:15px;padding:10px 13px}}
    </style></head><body><main class="page">
      <h1>🪷 蕉赖共修会</h1><p class="sub">义工值班安排</p>
      <form class="toolbar" method="get" action="/volunteer/day_schedule">
        <input type="date" name="date" value="{{ signup_date }}" aria-label="选择日期">
        <button type="submit">查看</button>
      </form>
      <div class="status">{{ '🟢 正式值班安排（已发布）' if published else '🟡 当天报名情况（尚未正式排班）' }}</div>
      <div class="notice-sheet">{{ styled_text }}</div>
      <p class="blessing">🪷 感恩每一位义工的发心护持</p>
      <div class="footer"><a class="link secondary" href="/volunteer">⬅ 返回义工首页</a></div>
    </main></body></html>
    """, signup_date=signup_date, published=published, styled_text=styled_text,
         background_index=background_index)

@schedule_bp.route("/volunteer/monthly_signup_list")
def volunteer_monthly_signup_list():
    today = malaysia_today()

    year = int(request.args.get("year", today.year))
    month = int(request.args.get("month", today.month))

    text = build_monthly_signup_text(year, month)

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                select
                    count(*) as total_signup,
                    count(distinct signup_date) as total_days
                from volunteer_schedule_signups
                where extract(year from signup_date) = %s
                  and extract(month from signup_date) = %s
                  and coalesce(status, 'pending') <> 'cancelled'
                  and role in ('值班', '卫生')
            """, (year, month))
            summary = cur.fetchone()

            cur.execute("""
                select role, count(*) as cnt
                from volunteer_schedule_signups
                where extract(year from signup_date) = %s
                  and extract(month from signup_date) = %s
                  and coalesce(status, 'pending') <> 'cancelled'
                  and role in ('值班', '卫生')
                group by role
            """, (year, month))
            role_rows = cur.fetchall()

    total_signup = summary["total_signup"] or 0
    total_days = summary["total_days"] or 0

    role_counts = {r["role"]: r["cnt"] for r in role_rows}
    duty_count = role_counts.get("值班", 0)
    cleaning_count = role_counts.get("卫生", 0)

    return render_template_string("""
    <style>
    body {
        font-family:"Microsoft YaHei", Arial;
        background:#f8f3ea;
        padding:20px;
        font-size:22px;
    }

    .page {
        max-width:1200px;
        margin:20px auto;
        background:white;
        padding:35px;
        border-radius:20px;
        box-shadow:0 4px 15px rgba(0,0,0,0.08);
    }

    h1 {
        color:#8b5a2b;
        font-size:40px;
        margin-top:0;
        margin-bottom:25px;
    }

    .dashboard-header {
        display:flex;
        gap:25px;
        margin-bottom:25px;
    }

    .filter-panel,
    .summary-panel {
        background:#fff8e6;
        border:2px solid #f3d9a5;
        border-radius:18px;
        padding:25px;
        box-sizing:border-box;
    }

    .filter-panel {
        flex:1;
    }

    .summary-panel {
        width:430px;
    }

    .row {
        display:flex;
        align-items:center;
        gap:15px;
        margin-bottom:18px;
    }

    .row label {
        min-width:100px;
        white-space:nowrap;
        font-size:28px;
        font-weight:bold;
        color:#5c3b1e;
    }

    .row select {
        width:180px;
        min-width:0;
        max-width:100%;
        font-size:24px;
        padding:10px;
        border-radius:10px;
        border:1px solid #d6c7b0;
    }

    .search-btn {
        margin-top:8px;
        width:280px;
        height:65px;
        background:#4CAF50;
        color:white;
        border:none;
        border-radius:12px;
        font-size:28px;
        font-weight:bold;
        cursor:pointer;
    }

    .top-summary {
        display:flex;
        justify-content:space-around;
        margin-bottom:16px;
    }

    .summary-item {
        text-align:center;
    }

    .summary-icon {
        font-size:30px;
        margin-bottom:4px;
    }

    .summary-num {
        font-size:50px;
        color:#1f6feb;
        font-weight:bold;
        line-height:1.1;
    }

    .summary-label {
        font-size:20px;
        color:#666;
    }

    .role-grid {
        display:grid;
        grid-template-columns:1fr 1fr;
        gap:14px 22px;
        margin-top:16px;
        font-size:23px;
        color:#5c3b1e;
    }

    .role-grid div {
        display:flex;
        justify-content:space-between;
        gap:12px;
    }

    .role-grid b {
        color:#1f6feb;
        font-size:24px;
    }

    textarea {
        width:100%;
        height:520px;
        font-size:22px;
        line-height:1.8;
        padding:22px;
        border-radius:18px;
        border:2px solid #ddd;
        box-sizing:border-box;
        background:#fffdf8;
        color:#222;
    }

    .action-row {
        display:flex;
        gap:18px;
        margin-top:22px;
        flex-wrap:wrap;
    }

    .big-btn {
        width:300px;
        padding:18px;
        border-radius:14px;
        border:0;
        color:white;
        font-size:26px;
        font-weight:bold;
        text-align:center;
        text-decoration:none;
        box-sizing:border-box;
    }

    .copy-btn { background:#2baa4b; }
    .home-btn { background:#666; }

    @media (max-width:900px) {
        .dashboard-header {
            flex-direction:column;
        }

        .summary-panel {
            width:auto;
        }

        .big-btn {
            width:100%;
        }

        h1 {
            font-size:32px;
        }
    }
    </style>

    <div class="page">

        <h1>📖 {{ year }}年{{ month }}月预报名名单</h1>

        <div class="dashboard-header">

            <div class="filter-panel">
                <form method="get">
                    <div class="row">
                        <label>年份：</label>
                        <select name="year" onchange="this.form.submit()">
                            {% for y in range(year - 1, year + 2) %}
                            <option value="{{ y }}" {% if y == year %}selected{% endif %}>
                                {{ y }}
                            </option>
                            {% endfor %}
                        </select>
                    </div>

                    <div class="row">
                        <label>月份：</label>
                        <select name="month" onchange="this.form.submit()">
                            {% for m in range(1, 13) %}
                            <option value="{{ m }}" {% if m == month %}selected{% endif %}>
                                {{ m }}月
                            </option>
                            {% endfor %}
                        </select>
                    </div>

                    <noscript><button type="submit" class="search-btn">🔍 查看</button></noscript>
                </form>
            </div>

            <div class="summary-panel">
                <div class="top-summary">
                    <div class="summary-item">
                        <div class="summary-icon">👥</div>
                        <div class="summary-num">{{ total_signup }}</div>
                        <div class="summary-label">值班及卫生人次</div>
                    </div>

                    <div class="summary-item">
                        <div class="summary-icon">📅</div>
                        <div class="summary-num">{{ total_days }}</div>
                        <div class="summary-label">值班及卫生报名天数</div>
                    </div>
                </div>

                <hr>

                <div class="role-grid">
                    <div><span>👤 值班</span><b>{{ duty_count }}</b></div>
                    <div><span>🧹 卫生</span><b>{{ cleaning_count }}</b></div>
                </div>
            </div>

        </div>

        <textarea id="signupText">{{ text }}</textarea>

        <div class="action-row">
            <button class="big-btn copy-btn" onclick="copySignupText()">
                📋 复制 WhatsApp 格式
            </button>

            <a class="big-btn home-btn" href="/volunteer">
                🏠 返回首页
            </a>
        </div>

    </div>

    <script>
    function copySignupText() {
        const text = document.getElementById("signupText");
        text.select();
        text.setSelectionRange(0, 999999);
        document.execCommand("copy");
        alert("✅ 已复制，可以贴去 WhatsApp");
    }
    </script>
    """,
    year=year,
    month=month,
    text=text,
    total_signup=total_signup,
    total_days=total_days,
    duty_count=duty_count,
    cleaning_count=cleaning_count)

@schedule_bp.route("/volunteer/prebook", methods=["GET", "POST"])
def volunteer_prebook():

    if request.method == "GET":
        today = malaysia_today()

        # Default to next month, including December-to-January rollover.
        next_month = today.replace(day=1) + timedelta(days=32)
        year = int(request.args.get("year", next_month.year))
        month = int(request.args.get("month", next_month.month))

        keyword = _volunteer_identity()
        branch = request.args.get("branch", "CHE").strip().upper()

        special_days = {}
        days_in_month = calendar.monthrange(year, month)[1]

        for d in range(1, days_in_month + 1):
            date_obj = date(year, month, d)
            info = get_special_day_info(date_obj)

            if info["template_type"] == "lunar_1_15":
                special_days[d] = "lunar"
            elif info["template_type"] == "buddhist_festival":
                special_days[d] = "festival"

        return render_template_string(
            VOLUNTEER_PREBOOK_HTML,
            login_id=_volunteer_identity(), login_name=session.get("volunteer_signup_name", ""),
            default_year=year,
            default_month=month,
            times=TIME_OPTIONS,
            keyword=keyword,
            branch=branch,
            special_days=special_days,
        )

    # POST 提交报名
    keyword = _volunteer_identity()
    branch = request.form.get("branch", "CHE").strip().upper()
    keyword = _volunteer_identity()

    year = int(request.form.get("year"))
    month = int(request.form.get("month"))
    days = request.form.getlist("days")
    role = request.form.get("role", "").strip()
    start_time = request.form.get("start_time", "").strip()
    end_time = request.form.get("end_time", "").strip()

    if role not in ("值班", "卫生"):
        return "❌ 多日报名仅开放值班及卫生<br><a href='/volunteer/prebook'>返回</a>", 400

    if not days:
        return "❌ 请选择至少一个日期<br><a href='/volunteer/prebook'>返回</a>"

    matches = find_volunteer_by_keyword(keyword)

    if not matches:
        return "❌ 找不到义工，请检查编号 / 姓名<br><a href='/volunteer/prebook'>返回</a>"

    if len(matches) > 1:
        return "❌ 找到多个同名义工，请用义工编号报名<br><a href='/volunteer/prebook'>返回</a>"

    vol = matches[0]
    vol_id = str(vol["id"])
    name = str(vol["name"])

    if role == "值班":
        s_min = time_to_minutes(start_time)
        e_min = time_to_minutes(end_time)

        if s_min is None or e_min is None:
            return "❌ 时间格式错误，请重新选择<br><a href='/volunteer/prebook'>返回</a>"

        if e_min <= s_min:
            return "❌ 结束时间必须比开始时间迟<br><a href='/volunteer/prebook'>返回</a>"
    else:
        start_time = None
        end_time = None

    inserted = 0
    updated = 0
    skipped = 0
    new_signup_items = []

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:

            for d in days:
                try:
                    signup_date = date(year, month, int(d))
                except ValueError:
                    skipped += 1
                    continue

                if role == "卫生":
                    cleaning_open, _, _ = cleaning_signup_window(signup_date)
                    if not cleaning_open:
                        skipped += 1
                        continue

                if role == "卫生":
                    cleaning_count = lock_and_count_cleaning(cur, signup_date)
                    cur.execute("""
                        SELECT 1 FROM volunteer_schedule_signups
                        WHERE volunteer_id = %s AND signup_date = %s
                          AND role = '卫生'
                          AND coalesce(status, 'pending') <> 'cancelled'
                        LIMIT 1
                    """, (vol_id, signup_date))
                    if cleaning_count >= 3 and cur.fetchone() is None:
                        skipped += 1
                        continue

                if role == "值班":
                    conflict = find_duty_time_conflict(cur, vol_id, signup_date, start_time, end_time)
                    if conflict and (str(conflict.get('start_time')) != start_time or
                                     str(conflict.get('end_time')) != end_time):
                        skipped += 1
                        continue

                cur.execute("""
                    select id
                    from volunteer_schedule_signups
                    where volunteer_id = %s
                    and signup_date = %s
                    and role = %s
                    and (%s <> '值班' OR (start_time = %s AND end_time = %s))
                    and coalesce(status, 'pending') <> 'cancelled'
                    limit 1
                """, (vol_id, signup_date, role, role, start_time, end_time))

                existing = cur.fetchone()

                if existing:
                    if signup_has_formal_assignment(cur, existing['id']):
                        skipped += 1
                        continue
                    cur.execute("""
                        update volunteer_schedule_signups
                        set start_time = %s,
                            end_time = %s,
                            status = 'pending',
                            assigned_place = null,
                            remarks = '义工多日报名更新'
                        where id = %s
                    """, (
                        start_time,
                        end_time,
                        existing["id"]
                    ))

                    updated += 1

                    new_signup_items.append((existing["id"], signup_date))
                else:
                    cur.execute("""
                        insert into volunteer_schedule_signups
                        (volunteer_id, name, signup_date, role, start_time, end_time, status, remarks)
                        values (%s, %s, %s, %s, %s, %s, 'pending', '义工多日报名')
                        returning id
                    """, (
                        vol_id,
                        name,
                        signup_date,
                        role,
                        start_time,
                        end_time
                    ))

                    new_signup_id = cur.fetchone()["id"]
                    new_signup_items.append((new_signup_id, signup_date))
                    inserted += 1

            conn.commit()

    for signup_id, signup_date in new_signup_items:

        try:
            synced = sync_schedule_after_signup_change(
                signup_id,
                action="upsert",
                changed_by="volunteer_multi_signup"
            )

            print(
                f"✅ {signup_date} 同步完成，共新增 {synced} 笔 assignment"
            )

        except Exception as e:
            print(f"多日报名同步失败：{signup_date}：{e}")

    return f"""
    <h1>✅ 多日报名完成</h1>
    <p>义工：{name}</p>
    <p>岗位：{role}</p>
    <p>新增：{inserted} 笔</p>
    <p>更新：{updated} 笔</p>
    <p>跳过无效日期：{skipped} 笔</p>
    <a href="/volunteer/prebook">继续多日报名</a><br>
    <a href="/volunteer">返回首页</a>
    """

@schedule_bp.route("/volunteer/my_schedule_search")
def volunteer_my_schedule_search():
    return redirect("/volunteer/my_schedule")
    return render_template_string("""
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>我的报名</title>

<link rel="stylesheet"
      href="{{ url_for('static', filename='css/toolbox.css') }}">

<style>
.search-wrap{
    max-width:760px;
    margin:auto;
    padding:24px;
}

.branch-search-row{
    display:grid;
    grid-template-columns:110px 1fr;
    gap:12px;
    align-items:center;
}

.branch-toggle-btn{
    height:66px;
    font-size:26px;
    font-weight:bold;
    color:white;
    border:none;
    border-radius:16px;
    cursor:pointer;
}

@media (max-width:700px){
    .branch-search-row{
        grid-template-columns:1fr;
    }
}
</style>

<script>
window.addEventListener("DOMContentLoaded", function(){
    document.getElementById("keyword").focus();
});
</script>

</head>

<body>

<div class="search-wrap">

    <div class="card">

        <h1 class="page-title">🔍 我的报名</h1>

        <p class="page-subtitle">
            查看已报名日期、修改报名、取消报名或查看正式岗位。
        </p>

        <div class="alert alert-info">
            💡 可输入义工编号、电话或姓名查询。
        </div>

        <form method="get" action="/volunteer/my_schedule">

            <div class="form-group">

                <label class="form-label">
                    义工编号 / 电话 / 姓名
                </label>

                <div class="branch-search-row">

                    <button
                        type="button"
                        id="branch_btn"
                        onclick="toggleBranch()"
                        class="branch-toggle-btn"
                        style="background:#28a745;">
                        CHE
                    </button>

                    <input
                        type="hidden"
                        id="branch"
                        name="branch"
                        value="CHE">

                    <input
                        class="form-input"
                        name="keyword"
                        id="keyword"
                        required
                        placeholder="例如：108 / 姓名 / 电话">

                </div>

            </div>

            <div class="btn-row">

                <button
                    class="btn-tool btn-blue btn-full"
                    type="submit">
                    📋 查询我的报名
                </button>

                <a
                    class="btn-tool btn-gray btn-full"
                    href="/volunteer">
                    ⬅ 返回义工首页
                </a>

            </div>

        </form>

    </div>

</div>

<script>
function toggleBranch(){

    const btn = document.getElementById("branch_btn");
    const branch = document.getElementById("branch");

    if(branch.value === "CHE"){
        branch.value = "STW";
        btn.innerText = "STW";
        btn.style.background = "#dc3545";
    }else{
        branch.value = "CHE";
        btn.innerText = "CHE";
        btn.style.background = "#28a745";
    }
}
</script>

</body>
</html>
""")

@schedule_bp.route("/volunteer/day_info")
def volunteer_day_info():
    date_str = request.args.get("date", "").strip()

    try:
        date_obj = datetime.strptime(date_str, "%Y-%m-%d").date()
    except:
        try:
            date_obj = datetime.strptime(date_str, "%Y/%m/%d").date()
        except:
            return {
                "ok": False,
                "template_type": "normal",
                "duty_start": "10:00am"
            }

    info = get_special_day_info(date_obj)

    template_type = info.get("template_type", "normal")

    if template_type == "lunar_1_15":
        duty_start = "8:00am"
    elif template_type == "buddhist_festival":
        duty_start = "8:00am"
    else:
        duty_start = "10:00am"

    cleaning_open, cleaning_cutoff_hour, _ = cleaning_signup_window(date_obj)

    # 卫生有效报名满3位时隐藏按钮；沿用报名提交时相同的人数口径。
    if cleaning_open:
        with get_conn() as conn:
            with conn.cursor(cursor_factory=RealDictCursor) as cur:
                cur.execute("""
                    SELECT count(*) AS cnt
                    FROM volunteer_schedule_signups
                    WHERE signup_date = %s AND role = '卫生'
                      AND coalesce(status, 'pending') <> 'cancelled'
                """, (date_obj,))
                cleaning_open = int(cur.fetchone()['cnt'] or 0) < 3

    return {
        "ok": True,
        "cleaning_signup_open": cleaning_open,
        "cleaning_cutoff": "8:00am" if cleaning_cutoff_hour == 8 else "10:00am",
        "template_type": template_type,
        "duty_start": duty_start,
        "lunar_text": info.get("lunar_text", ""),
        "special_names": info.get("special_names", [])
    }


