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
from db import db_query, get_conn

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
.hf-form-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px}.hf-field{display:flex;flex-direction:column;gap:5px}.hf-field input,.hf-field select{width:100%;margin:0}.hf-wide{grid-column:1/-1}.hf-form-grid label{font-weight:600}.hf-event{border:1px solid #e0eae3;background:#fbfdfb;border-radius:12px;padding:14px;margin:12px 0}.hf-event form{display:inline-block;margin:6px 5px 0 0}.hf-event button{font-size:15px;padding:9px 13px}.hf-muted{color:#63746a;font-size:14px}.hf-actions{margin-top:18px}.hf-actions button{font-weight:700}@media(max-width:520px){.hf-form-grid{grid-template-columns:1fr}.card{padding:16px}}
main{padding-bottom:100px;font-size:18px;line-height:1.55}h1{font-size:26px}h2{font-size:23px;margin:0 0 14px}a{color:#145b87;font-weight:700}button,.btn{margin:3px 0} .hf-menu{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px}.hf-tile{display:flex;flex-direction:column;align-items:center;justify-content:center;min-height:94px;padding:12px 6px;border-radius:12px;color:white;text-decoration:none;text-align:center;font-weight:750;font-size:19px;box-sizing:border-box;background:#16854e}.hf-tile span{font-size:27px}.hf-blue{background:#246b9b}.hf-purple{background:#704ba0}.hf-brown{background:#805b36}.hf-navy{background:#344e79}.hf-gray{background:#52616a}.hf-orange{background:#a65312}.hf-day{font-weight:800;margin:18px 0 8px;color:#254b39}.hf-agenda{border:1px solid #d2e2d8;border-radius:12px;padding:12px;margin:8px 0;background:#f8fbf9;overflow-wrap:anywhere}.hf-agenda-head{display:flex;gap:8px;align-items:center;justify-content:space-between;flex-wrap:wrap}.hf-status{font-size:15px;font-weight:700;padding:4px 9px;border-radius:8px;background:#e1f3e7;color:#17633b}.hf-status.draft{background:#fff0cb;color:#855000}.hf-status.rejected{background:#ffe3e3;color:#a22}.hf-back{position:fixed;right:16px;bottom:18px;z-index:40;background:#344e79!important;color:white!important;padding:13px 17px;border-radius:12px;text-decoration:none;box-shadow:0 3px 12px #0003;font-size:17px;font-weight:800}.hf-logout{background:#f9e4e4;color:#952b2b}.hf-nav-buttons{gap:12px!important}.hf-quicknav{display:flex;flex-wrap:wrap;gap:10px;margin:12px 0}.hf-quicknav a{padding:10px 14px;background:#e2eee7;border-radius:9px;text-decoration:none}.hf-event{overflow-wrap:anywhere} @media(max-width:520px){main{margin:14px auto}.hf-menu{gap:10px}.hf-tile{font-size:17px;min-height:86px}.hf-back{bottom:14px}h1{font-size:23px}}
</style></head><body><main>
<h1>🪷 蕉赖观音堂｜弘法</h1>''' + '{% block content %}{% endblock %}' + '</main></body></html>'


def render(content, **kwargs):
    back = {
        'hongfa.plans': ('/hongfa/dashboard','🏠 返回管理首页'),
        'hongfa.week': ('/hongfa/dashboard','🏠 返回管理首页'),
        'hongfa.locations': ('/hongfa/settings-menu','⬅ 返回系统设置'),
        'hongfa.team_admins': ('/hongfa/settings-menu','⬅ 返回系统设置'),
        'hongfa.event_code': ('/hongfa/plans','⬅ 返回活动安排'),
        'hongfa.attendance': ('/hongfa/dashboard','🏠 返回管理首页'),
        'hongfa.event_attendance': ('/hongfa/attendance','⬅ 返回出席记录'),
        'hongfa.stock': ('/hongfa/book-menu','⬅ 返回书籍管理'),
        'hongfa.items': ('/hongfa/book-menu','⬅ 返回书籍管理'),
        'hongfa.report': ('/hongfa/dashboard','🏠 返回管理首页'),
        'hongfa.books_manage': ('/hongfa/book-menu','⬅ 返回书籍管理'),
        'hongfa.books_add': ('/hongfa/books/manage','⬅ 返回管理书籍'),
    }.get(request.endpoint)
    if request.endpoint == 'hongfa.books' and not kwargs.get('is_volunteer'):
        back = ('/hongfa/book-menu','⬅ 返回书籍管理')
        content += '<style>.hf-back{bottom:90px}</style>'
    if back:
        content += '<a class="hf-back" href="'+back[0]+'">'+back[1]+'</a>'
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
    admins = db_query("SELECT a.team_id FROM hongfa_team_admins a JOIN hongfa_teams t ON t.id=a.team_id WHERE a.volunteer_id=%s AND a.active=true AND t.active=true", (vid,), fetchall=True) or []
    secretary = any(x['role']=='secretary' for x in roles)
    leaders = {x['team_id'] for x in roles if x['role']=='leader' and x['team_id'] is not None}
    teams = leaders | {x['team_id'] for x in admins}
    if not secretary and not teams:
        session.pop('hongfa_volunteer_id', None)
        return None
    return {'id':vid,'secretary':secretary,'leader_team_ids':sorted(leaders),
            'team_ids':sorted(teams),'small_admin':bool(teams and not leaders and not secretary)}



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
        <a class="btn hf-floating-login" href="{{url_for('hongfa.login')}}">🔐 负责人／小管理员／秘书登录</a>''',
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
                roles = db_query('SELECT 1 FROM hongfa_roles WHERE volunteer_id=%s AND active=true UNION ALL SELECT 1 FROM hongfa_team_admins a JOIN hongfa_teams t ON t.id=a.team_id WHERE a.volunteer_id=%s AND a.active=true AND t.active=true LIMIT 1',
                                 (vid,vid),fetchone=True)
                if not roles:
                    error = '此编号尚未获得弘法管理权限。'
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
    role_label = '弘法秘书' if who['secretary'] else '、'.join(t['name'] for t in team_names) + ('小管理员' if who['small_admin'] else '负责人')
    today = datetime.now(TZ).date()
    today_events = db_query('''SELECT e.id,t.name AS team_name,l.name AS location_name FROM hongfa_events e
        JOIN hongfa_teams t ON t.id=e.team_id LEFT JOIN hongfa_locations l ON l.id=e.location_id
        WHERE e.event_date=%s AND e.status='published' AND e.closed_at IS NULL
        AND e.location_id IS NOT NULL AND e.team_id=ANY(%s) ORDER BY t.id''',
        (today, who['team_ids']), fetchall=True) or []
    return render('''<div class="card"><b style="font-size:24px">👤 {{display_name}}</b><div>{{role_label}} · {{who.id}}</div></div>
        {% for e in today_events %}<a class="hf-tile hf-orange" style="margin:12px 0;min-height:65px" href="{{url_for('hongfa.event_code',event_id=e.id)}}">🔐 今日签到码｜{{e.team_name}}</a>{% endfor %}
        <div class="card"><h2>常用功能</h2><div class="hf-menu">
        <a class="hf-tile" href="/hongfa/plans"><span>📅</span>{{'审核发布' if who.secretary else '提交安排'}}</a>
        <a class="hf-tile hf-blue" href="/hongfa/week"><span>📋</span>每周时间表</a>
        <a class="hf-tile hf-purple" href="/hongfa/attendance"><span>👥</span>出席记录</a>
        <a class="hf-tile hf-brown" href="{{url_for('hongfa.book_menu')}}"><span>📚</span>书籍管理</a>
        <a class="hf-tile hf-navy" href="/hongfa/report"><span>📊</span>弘法报表</a>
        <a class="hf-tile hf-gray" href="{{url_for('hongfa.settings_menu')}}"><span>⚙️</span>系统设置</a></div></div>
        <div class="card"><h2>📅 近期弘法安排</h2>
        {% for e in rows %}{% if loop.changed(e.event_date) %}<div class="hf-day">{{e.event_date.strftime('%m月%d日')}}（星期{{'一二三四五六日'[e.event_date.weekday()]}}）</div>{% endif %}
        <div class="hf-agenda"><div class="hf-agenda-head"><b>{{e.team_name}}</b><span class="hf-status {{e.status}}">{{{'published':'🟢 已发布','draft':'🟠 待审核','rejected':'🔴 已退回'}.get(e.status,e.status)}}</span></div><div>{{'📍 '+e.location_name if e.location_name else '休息 / Off'}}</div></div>
        {% else %}<p>暂无安排。</p>{% endfor %}</div>
        <small>系统试用中，请勿直接用于正式签到。</small>
        <form method="post" action="{{url_for('hongfa.logout')}}"><input type="hidden" name="csrf_token" value="{{csrf}}"><button class="hf-logout">退出登录</button></form>''',who=who,rows=rows,csrf=csrf_token(),display_name=display_name,role_label=role_label,today_events=today_events)


@hongfa_bp.route('/settings-menu')
@require_login
def settings_menu():
    return render('''<div class="card"><h2>⚙️ 系统设置</h2><div class="hf-menu"><a class="hf-tile" href="/hongfa/locations"><span>📍</span>弘法地点</a>{% if who.secretary or who.leader_team_ids %}<a class="hf-tile hf-purple" href="{{url_for('hongfa.team_admins')}}"><span>👥</span>小管理员设置</a>{% endif %}</div></div><a class="hf-back" href="/hongfa/dashboard">🏠 返回管理首页</a>''',who=actor())


@hongfa_bp.route('/book-menu')
@require_login
def book_menu():
    return render('''<div class="card"><h2>📚 书籍管理</h2><div class="hf-menu">{% if who.team_ids %}<a class="hf-tile" href="/hongfa/books"><span>📖</span>结缘／点货补货</a>{% endif %}<a class="hf-tile hf-blue" href="/hongfa/stock"><span>📦</span>查看库存</a><a class="hf-tile hf-brown" href="/hongfa/books/manage"><span>⚙️</span>管理书籍</a>{% if who.secretary %}<a class="hf-tile hf-gray" href="/hongfa/items"><span>📖</span>原书籍目录</a>{% endif %}{% if who.team_ids %}<a class="hf-tile hf-gray" href="/hongfa/stock"><span>📋</span>原库存操作</a>{% endif %}</div></div><a class="hf-back" href="/hongfa/dashboard">🏠 返回管理首页</a>''',who=actor())


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


@hongfa_bp.route('/team-admins', methods=['GET','POST'])
@require_login
def team_admins():
    who = actor()
    teams = db_query('SELECT id,name FROM hongfa_teams WHERE active=true ORDER BY id',fetchall=True) or []
    teams = [t for t in teams if who['secretary'] or t['id'] in who['leader_team_ids']]
    if not teams: abort(403)
    allowed = {t['id'] for t in teams}
    selected = (request.form if request.method == 'POST' else request.args).get('team_id',type=int)
    if selected is None: selected = teams[0]['id']
    if selected not in allowed: abort(403)
    error = ''; candidate = None
    if request.method=='POST':
        if not valid_csrf(): abort(400)
        action = request.form.get('action')
        if action in ('lookup','add'):
            branch=request.form.get('branch','CHE'); number=request.form.get('number','').strip()
            if branch not in ('CHE','STW') or not number.isascii() or not number.isdigit() or not 1<=len(number)<=6 or int(number)==0:
                error='编号格式不正确。'
            else:
                n=int(number); vid=str(n) if branch=='CHE' and 800<=n<=899 else f'{branch}-{n}'
                v=find_volunteer(vid)
                if not v or str(v.get('编号','')).strip().upper()!=vid.upper() or v.get('状态') not in ('在册','',None) or not str(v.get('PIN') or '').strip():
                    error='找不到有效在册义工，或尚未设置 PIN。'
                elif db_query("SELECT 1 FROM hongfa_roles WHERE volunteer_id=%s AND active=true LIMIT 1",(vid,),fetchone=True):
                    error='此义工已有负责人或秘书身份。'
                elif action=='lookup':
                    candidate={'id':vid,'name': admin_name(v,vid),'branch':branch,'number':n}
                else:
                    db_query("INSERT INTO hongfa_team_admins(team_id,volunteer_id,created_by) VALUES (%s,%s,%s) ON CONFLICT (team_id,volunteer_id) DO UPDATE SET active=true",(selected,vid,who['id']))
                    return redirect(url_for('hongfa.team_admins',team_id=selected))
        elif action=='remove':
            db_query('UPDATE hongfa_team_admins SET active=false WHERE team_id=%s AND volunteer_id=%s',(selected,request.form.get('vid','')))
            return redirect(url_for('hongfa.team_admins',team_id=selected))
        else: abort(400)
    rows=db_query('SELECT volunteer_id FROM hongfa_team_admins WHERE team_id=%s AND active=true ORDER BY volunteer_id',(selected,),fetchall=True) or []
    for row in rows: row['name']=admin_name(find_volunteer(row['volunteer_id']) or {},row['volunteer_id'])
    return render('''<div class="card"><h2>👥 小管理员设置</h2><p><a href="{{url_for('hongfa.dashboard')}}">🏠 返回管理首页</a></p>
    <label>弘法组</label><select onchange="location.href='?team_id='+this.value">{% for t in teams %}<option value="{{t.id}}" {% if t.id==selected %}selected{% endif %}>{{t.name}}</option>{% endfor %}</select>
    {% if error %}<p class="warn">{{error}}</p>{% endif %}
    <form method="post"><input type="hidden" name="csrf_token" value="{{csrf}}"><input type="hidden" name="team_id" value="{{selected}}"><div class="hf-id-row"><input type="hidden" id="ha-branch" name="branch" value="CHE"><button type="button" class="hf-branch-toggle che" onclick="var i=document.getElementById('ha-branch');i.value=i.value==='CHE'?'STW':'CHE';this.textContent=i.value;this.className='hf-branch-toggle '+i.value.toLowerCase()">CHE</button><input name="number" required inputmode="numeric" pattern="[0-9]+" maxlength="6" placeholder="义工编号（数字）"></div><button name="action" value="lookup">查询义工</button></form>
    {% if candidate %}<form method="post"><p><b>{{candidate.name}}｜{{candidate.id}}</b></p><input type="hidden" name="csrf_token" value="{{csrf}}"><input type="hidden" name="team_id" value="{{selected}}"><input type="hidden" name="branch" value="{{candidate.branch}}"><input type="hidden" name="number" value="{{candidate.number}}"><button name="action" value="add">➕ 添加小管理员</button></form>{% endif %}
    <h3>现有小管理员</h3>{% for row in rows %}<form method="post" class="hf-event" onsubmit="return confirm('确认解除权限？')"><b>{{row.name}}｜{{row.volunteer_id}}</b><input type="hidden" name="csrf_token" value="{{csrf}}"><input type="hidden" name="team_id" value="{{selected}}"><input type="hidden" name="vid" value="{{row.volunteer_id}}"><button name="action" value="remove" style="background:#a33">解除权限</button></form>{% else %}<p>暂无小管理员。</p>{% endfor %}</div>''', teams=teams,selected=selected,rows=rows,error=error,candidate=candidate,csrf=csrf_token())


def admin_name(v,vid):
    return next((str(v[k]).strip() for k in ('中文姓名','姓名','名字','name') if v.get(k)),vid)
