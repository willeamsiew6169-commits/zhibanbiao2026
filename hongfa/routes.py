"""弘法系统第一批：独立权限、负责人/秘书只读首页、已发布活动列表。
尚未实现签到码、签到、活动发布、库存写入；不要作为正式签到系统使用。
"""
from datetime import datetime, timedelta
from functools import wraps
import hmac
import secrets
from zoneinfo import ZoneInfo

from flask import Blueprint, abort, redirect, render_template_string, request, session, url_for
from attendance_service import find_volunteer, verify_pin_for_volunteer
from db import db_query

hongfa_bp = Blueprint('hongfa', __name__, url_prefix='/hongfa')
TZ = ZoneInfo('Asia/Kuala_Lumpur')

BASE = '''<!doctype html><html lang="zh"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>蕉赖观音堂｜弘法</title><style>
body{font-family:system-ui,sans-serif;background:#f5f7f5;color:#20352c;margin:0}
main{max-width:760px;margin:25px auto;padding:0 14px}.card{background:white;border-radius:16px;padding:20px;margin:14px 0;box-shadow:0 3px 12px #0001}
input,button{font-size:18px;padding:12px;border-radius:9px;box-sizing:border-box}input{width:100%;border:1px solid #ccc;margin:7px 0 15px}
button,.btn{background:#16854e;color:white;border:0;display:inline-block;text-decoration:none;padding:13px 18px;border-radius:9px;cursor:pointer}
small{color:#66776f}.warn{color:#ad3434}
/* 仅首页登录入口：固定右下角，滚动时保持可见。 */
.hf-floating-login{position:fixed;right:max(16px,env(safe-area-inset-right));bottom:max(18px,env(safe-area-inset-bottom));z-index:1000;font-size:19px;font-weight:800;box-shadow:0 4px 16px #0003;max-width:calc(100vw - 32px);text-align:center}
.hf-id-row{display:flex;gap:10px;align-items:center;margin:7px 0 10px}.hf-id-row input{flex:1;min-width:0;margin:0}.hf-branch-toggle{width:94px;flex:none;font-weight:800;color:white}.hf-branch-toggle.che{background:#17834d}.hf-branch-toggle.stw{background:#c73545}
select,textarea{font:inherit;border:1px solid #cddbd3;border-radius:10px;padding:11px;background:#fff;max-width:100%}
.hf-form-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px}.hf-field{display:flex;flex-direction:column;gap:5px}.hf-field input,.hf-field select{width:100%;margin:0}.hf-wide{grid-column:1/-1}.hf-form-grid label{font-weight:600}.hf-event{border:1px solid #e0eae3;background:#fbfdfb;border-radius:12px;padding:14px;margin:12px 0}.hf-event form{display:inline-block;margin:6px 5px 0 0}.hf-event button{font-size:15px;padding:9px 13px}.hf-muted{color:#63746a;font-size:14px}.hf-actions{margin-top:18px}.hf-actions button{font-weight:700}@media(max-width:520px){.hf-form-grid{grid-template-columns:1fr}.card{padding:16px}}</style></head><body><main>
<h1>🪷 蕉赖观音堂｜弘法</h1>''' + '{% block content %}{% endblock %}' + '</main></body></html>'


def render(content, **kwargs):
    return render_template_string(BASE.replace('{% block content %}{% endblock %}', content), **kwargs)


def csrf_token():
    if 'hongfa_csrf' not in session:
        session['hongfa_csrf'] = secrets.token_urlsafe(32)
    return session['hongfa_csrf']


def valid_csrf():
    return hmac.compare_digest(str(session.get('hongfa_csrf', '')),
                               str(request.form.get('csrf_token', ''))) and bool(session.get('hongfa_csrf'))


def actor():
    vid = session.get('hongfa_volunteer_id')
    if not vid:
        return None
    # 权限每次从数据库读取，撤销后即时失效。
    roles = db_query('''SELECT role,team_id FROM hongfa_roles
        WHERE volunteer_id=%s AND active=true''', (vid,), fetchall=True) or []
    if not roles:
        session.pop('hongfa_volunteer_id', None)
        return None
    return {'id':vid,'secretary':any(r['role']=='secretary' for r in roles),
            'team_ids':sorted({r['team_id'] for r in roles if r['role']=='leader' and r['team_id'] is not None})}


def require_login(fn):
    @wraps(fn)
    def inner(*args, **kwargs):
        if actor() is None:
            return redirect(url_for('hongfa.login'))
        return fn(*args, **kwargs)
    return inner


@hongfa_bp.route('/')
def index():
    today = datetime.now(TZ).date()
    # 首页仅展示已发布且有地点的活动；未来六天只预览，不开放提前签到。
    rows = db_query('''SELECT e.id,e.event_date,e.starts_at,e.ends_at,t.name AS team_name,
        l.name AS location_name,l.maps_url FROM hongfa_events e
        JOIN hongfa_teams t ON t.id=e.team_id
        JOIN hongfa_locations l ON l.id=e.location_id
        WHERE e.event_date >= %s AND e.event_date < %s AND e.status='published'
        ORDER BY e.event_date,e.starts_at,t.name''',
        (today, today + timedelta(days=7)), fetchall=True) or []
    current = [e for e in rows if e['event_date'] == today]
    upcoming = [e for e in rows if e['event_date'] > today]
    return render('''<div class="card"><h2>📍 今日弘法地点</h2>
        {% for e in current %}<div class="hf-event"><b>{{e.location_name}}</b>｜{{e.team_name}}<br>
        ⏰ {{(e.starts_at|string)[:5]}}–{{(e.ends_at|string)[:5]}}<br>
        {% if e.maps_url %}<a href="{{e.maps_url}}" rel="noopener noreferrer" target="_blank">📍 地图导航</a>　{% endif %}
        <a class="btn" href="{{url_for('hongfa.join',event_id=e.id)}}">🪷 现场签到</a></div>
        {% else %}<p>今日暂无已发布的弘法活动。</p>{% endfor %}
        <p class="hf-muted">现场签到仅限活动当天，并须在负责人设定的签到时间内输入有效现场码。</p></div>
        <div class="card"><h2>📅 未来六天弘法安排</h2>
        {% for e in upcoming %}<div class="hf-event"><b>{{e.event_date}}｜{{e.team_name}}</b><br>
        📍 {{e.location_name}}<br>⏰ {{(e.starts_at|string)[:5]}}–{{(e.ends_at|string)[:5]}}
        {% if e.maps_url %}<br><a href="{{e.maps_url}}" rel="noopener noreferrer" target="_blank">📍 地图导航</a>{% endif %}</div>
        {% else %}<p>未来六天暂无已发布的弘法活动。</p>{% endfor %}
        <p class="hf-muted">未来活动仅供查看，活动当天才会出现签到入口。</p></div>
        <a class="btn hf-floating-login" href="{{url_for('hongfa.login')}}">🔐 负责人／秘书登录</a>''',
        current=current, upcoming=upcoming)


@hongfa_bp.route('/login', methods=['GET','POST'])
def login():
    error = ''
    if request.method == 'POST':
        if not valid_csrf():
            abort(400)
        branch = request.form.get('branch','CHE').upper()
        number = request.form.get('number','').strip()
        pin = request.form.get('pin','')
        if branch not in ('CHE','STW') or not number.isascii() or not number.isdigit() or not (1 <= len(number) <= 6) or int(number) == 0:
            error = '编号格式不正确。'
        else:
            raw_id = str(int(number)) if branch == 'CHE' and 800 <= int(number) <= 899 else f'{branch}-{int(number)}'
            vol = find_volunteer(raw_id)
            if (not vol or vol.get('状态') not in ('在册','',None)
                    or str(vol.get('编号', '')).strip().upper() != raw_id.upper()
                    or not str(vol.get('PIN') or '').strip()
                    or not verify_pin_for_volunteer(vol,pin)):
                error = '编号或 PIN 不正确。'
            else:
                vid = str(vol['编号'])
                roles = db_query('SELECT 1 FROM hongfa_roles WHERE volunteer_id=%s AND active=true LIMIT 1',
                                 (vid,),fetchone=True)
                if not roles:
                    error = '此编号尚未获得弘法负责人／秘书权限。'
                else:
                    # 只设置弘法身份，不清空其他独立系统的 Session。
                    session['hongfa_volunteer_id'] = vid
                    session.permanent = True
                    return redirect(url_for('hongfa.dashboard'))
    return render('''<div class="card"><h2>🔐 弘法管理登录</h2>
        {% if error %}<p class="warn">{{error}}</p>{% endif %}
        <form method="post"><input type="hidden" name="csrf_token" value="{{csrf}}"><label>义工编号</label>
        <div class="hf-id-row"><input type="hidden" name="branch" id="hf-branch" value="CHE">
        <button type="button" id="hf-branch-toggle" class="hf-branch-toggle che" aria-label="切换 CHE 或 STW" aria-pressed="false">CHE</button>
        <input name="number" inputmode="numeric" pattern="[0-9]+" maxlength="6" placeholder="只输入数字，例如 69 或 819" required></div>
        <small>🟢 CHE 包含 8xx 纯义工；🔴 STW 为沙登编号。</small>
        <script>(function(){var b=document.getElementById('hf-branch-toggle'),i=document.getElementById('hf-branch');b.addEventListener('click',function(){var s=i.value==='CHE'?'STW':'CHE';i.value=s;b.textContent=s;b.className='hf-branch-toggle '+s.toLowerCase();b.setAttribute('aria-pressed',String(s==='STW'));});})();</script>
        <label>PIN</label><input name="pin" type="password" autocomplete="current-password" required>
        <button type="submit">登录</button></form></div>''',error=error,csrf=csrf_token())


@hongfa_bp.route('/dashboard')
@require_login
def dashboard():
    who = actor()
    if who['secretary']:
        rows = db_query('''SELECT e.event_date,t.name AS team_name,e.status,l.name AS location_name
            FROM hongfa_events e JOIN hongfa_teams t ON t.id=e.team_id
            LEFT JOIN hongfa_locations l ON l.id=e.location_id
            WHERE e.event_date >= %s ORDER BY e.event_date,t.name LIMIT 100''',
            (datetime.now(TZ).date(),),fetchall=True) or []
    elif who['team_ids']:
        rows = db_query('''SELECT e.event_date,t.name AS team_name,e.status,l.name AS location_name
            FROM hongfa_events e JOIN hongfa_teams t ON t.id=e.team_id
            LEFT JOIN hongfa_locations l ON l.id=e.location_id
            WHERE e.event_date >= %s AND e.team_id=ANY(%s)
            ORDER BY e.event_date,t.name LIMIT 100''',
            (datetime.now(TZ).date(),who['team_ids']),fetchall=True) or []
    else:
        abort(403)
    volunteer = find_volunteer(who['id']) or {}
    display_name = next((str(volunteer.get(k)).strip() for k in ('中文姓名', '姓名', '名字', 'name') if volunteer.get(k)), '')
    if not display_name:
        display_name = {'CHE-50':'俊达师兄','CHE-69':'Alex师兄','CHE-228':'诗霞师兄'}.get(who['id'], who['id'])
    team_names = db_query('SELECT name FROM hongfa_teams WHERE id=ANY(%s) ORDER BY id', (who['team_ids'],), fetchall=True) or []
    role_label = '弘法秘书' if who['secretary'] else '、'.join(t['name'] for t in team_names) + '负责人'
    today = datetime.now(TZ).date()
    today_events = db_query('''SELECT e.id,t.name AS team_name,l.name AS location_name FROM hongfa_events e
        JOIN hongfa_teams t ON t.id=e.team_id LEFT JOIN hongfa_locations l ON l.id=e.location_id
        WHERE e.event_date=%s AND e.status='published' AND e.closed_at IS NULL
        AND e.location_id IS NOT NULL AND e.team_id=ANY(%s) ORDER BY t.id''',
        (today, who['team_ids']), fetchall=True) or []
    return render('''<div class="card"><div style="font-size:27px;font-weight:800;line-height:1.4">👤 {{display_name}}</div>
        <div style="font-size:21px;font-weight:700;margin-top:6px">{{role_label}}</div>
        <p style="color:#63746a">编号：{{who.id}}</p>
        <p>身份：{{'秘书（跨组查看）' if who.secretary else '负责人（仅自己的组）'}}</p>
        {% for e in today_events %}<p><a class="btn" href="{{url_for('hongfa.event_code',event_id=e.id)}}">🔐 生成今日签到码｜{{e.team_name}}</a></p>{% endfor %}
        <h2>📅 弘法安排</h2>
        {% for e in rows %}<p>{{e.event_date}}｜{{e.team_name}}｜{{e.location_name or 'Off'}}｜{{e.status}}</p>
        {% else %}<p>暂无安排。</p>{% endfor %}
        <p><a class="btn" href="/hongfa/plans">📅 活动提交／审核</a> <a class="btn" href="/hongfa/week">📋 每周表</a> <a class="btn" href="/hongfa/locations">📍 地点</a> <a class="btn" href="/hongfa/attendance">👥 出席</a> <a class="btn" href="/hongfa/books">📚 快速扣书／点货</a> <a class="btn" href="/hongfa/stock">📚 原库存</a> <a class="btn" href="/hongfa/report">📊 报表</a> {% if who.secretary %}<a class="btn" href="/hongfa/items">📖 经书目录</a>{% endif %}</p><small>第三批本地测试中，请勿直接用于正式签到。</small></div>
        <form method="post" action="{{url_for('hongfa.logout')}}"><input type="hidden" name="csrf_token" value="{{csrf}}"><button>退出登录</button></form>''',who=who,rows=rows,csrf=csrf_token(),display_name=display_name,role_label=role_label,today_events=today_events)


@hongfa_bp.route('/logout',methods=['POST'])
def logout():
    if not valid_csrf():
        abort(400)
    session.pop('hongfa_volunteer_id',None)
    return redirect(url_for('hongfa.index'))

# 弘法第二批：独立扩展，不修改现有系统。
from .phase2 import register_phase2
register_phase2(hongfa_bp, actor, render, csrf_token, valid_csrf)

# 弘法第三批：独立扩展，不影响现有系统。
from .phase3 import register_phase3
register_phase3(hongfa_bp, actor, render, csrf_token, valid_csrf)

# 经书快速登记独立模块：保留原库存及报表逻辑。
from .hongfa_books import register_books
register_books(hongfa_bp, actor, render, csrf_token, valid_csrf)
