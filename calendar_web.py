"""年度日历领取与库存；不改动原有签到和报表数据。"""
from datetime import datetime
from flask import Blueprint, request, redirect, render_template_string, jsonify, session
from psycopg2.extras import RealDictCursor
from db import get_conn, db_query
from utils import MY_TZ
from attendance_service import find_volunteer, verify_pin_for_volunteer

calendar_bp = Blueprint('calendar', __name__)


def _year(cur):
    cur.execute('SELECT year,initial_stock,is_open FROM calendar_campaigns WHERE is_open=true ORDER BY year DESC LIMIT 1')
    return cur.fetchone()


def _stock(cur, year, initial):
    cur.execute('SELECT COALESCE(SUM(quantity),0) AS delta FROM calendar_movements WHERE year=%s', (year,))
    return int(initial) + int(cur.fetchone()['delta'])


def _identity(number):
    volunteer = find_volunteer(str(number or '').strip())
    if not volunteer or volunteer.get('状态') not in ('在册', '', None):
        return None
    return volunteer


def _admin():
    return session.get('calendar_admin') is True


def _reply(ok, msg, **extra):
    return jsonify(ok=ok, message=msg, **extra), (200 if ok else 400)


@calendar_bp.route('/calendar')
def home():
    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            campaign = _year(cur)
            stock = _stock(cur, campaign['year'], campaign['initial_stock']) if campaign else 0
    return render_template_string(PAGE, campaign=campaign, stock=stock, admin=_admin())


@calendar_bp.route('/calendar/lookup', methods=['POST'])
def lookup():
    number = request.form.get('number', '')
    pin = request.form.get('pin', '')
    v = _identity(number)
    if not v or not verify_pin_for_volunteer(v, pin):
        return _reply(False, '编号或 PIN 不正确。')
    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            c = _year(cur)
            if not c:
                return _reply(False, '目前没有开放日历领取。')
            cur.execute('SELECT status FROM calendar_personal WHERE year=%s AND volunteer_id=%s', (c['year'], v['编号']))
            r = cur.fetchone()
            return _reply(True, '查询成功', name=v['姓名'], status=r['status'] if r else '', year=c['year'])


@calendar_bp.route('/calendar/personal', methods=['POST'])
def personal():
    v = _identity(request.form.get('number'))
    if not v or not verify_pin_for_volunteer(v, request.form.get('pin', '')):
        return _reply(False, '编号或 PIN 不正确。')
    status = request.form.get('status')
    if status != 'claimed':
        return _reply(False, '状态无效。')
    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute('SELECT * FROM calendar_campaigns WHERE is_open=true ORDER BY year DESC LIMIT 1 FOR UPDATE')
            c = cur.fetchone()
            if not c:
                return _reply(False, '领取已关闭。')
            cur.execute('SELECT status FROM calendar_personal WHERE year=%s AND volunteer_id=%s', (c['year'], v['编号']))
            if cur.fetchone():
                return _reply(False, '今年已经登记过，无需重复登记。')
            if _stock(cur, c['year'], c['initial_stock']) < 1:
                return _reply(False, '日历库存不足。')
            cur.execute('INSERT INTO calendar_personal(year,volunteer_id,volunteer_name,status) VALUES (%s,%s,%s,%s)', (c['year'], v['编号'], v['姓名'], status))
            cur.execute("INSERT INTO calendar_movements(year,kind,quantity,volunteer_id,volunteer_name) VALUES (%s,'personal',-1,%s,%s)", (c['year'], v['编号'], v['姓名']))
    return _reply(True, '已登记领取1本。')


@calendar_bp.route('/calendar/admin/quick-toggle', methods=['POST'])
def quick_toggle():
    """仅管理员工具使用：按年度切换日历领取开放状态。"""
    from admin_web import ADMIN_PIN
    from flask import url_for, flash
    pin = (request.form.get('admin_pin') or '').strip()
    if pin != ADMIN_PIN:
        return '无权限', 403
    year = request.form.get('year', type=int)
    if year is None:
        return '年度无效', 400
    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute('SELECT year,is_open FROM calendar_campaigns WHERE year=%s FOR UPDATE', (year,))
            campaign = cur.fetchone()
            if not campaign:
                return '年度不存在', 404
            if not campaign['is_open']:
                cur.execute('UPDATE calendar_campaigns SET is_open=false')
            cur.execute('UPDATE calendar_campaigns SET is_open=%s WHERE year=%s',
                        (not campaign['is_open'], year))
    flash('日历领取开关已更新。', 'ok')
    return redirect(url_for('admin.admin_report', pin=pin))


@calendar_bp.route('/calendar/admin-login', methods=['POST'])
def admin_login():
    from admin_web import ADMIN_PIN
    if request.form.get('admin_pin', '') != ADMIN_PIN:
        return redirect('/calendar?error=pin')
    session['calendar_admin'] = True
    return redirect('/calendar/admin')


@calendar_bp.route('/calendar/admin')
def admin_page():
    if not _admin():
        return redirect('/calendar')
    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute('SELECT * FROM calendar_campaigns ORDER BY year DESC')
            years = cur.fetchall()
            selected = request.args.get('year', type=int) or (years[0]['year'] if years else None)
            c = next((r for r in years if r['year'] == selected), None)
            stock = _stock(cur, selected, c['initial_stock']) if c else 0
            records = []
            moves = []
            if c:
                cur.execute('SELECT volunteer_id,volunteer_name,status,updated_at FROM calendar_personal WHERE year=%s ORDER BY volunteer_name', (selected,))
                records = cur.fetchall()
                cur.execute('SELECT kind,quantity,volunteer_name,note,created_at FROM calendar_movements WHERE year=%s ORDER BY id DESC LIMIT 100', (selected,))
                moves = cur.fetchall()
    volunteers = db_query("SELECT id,name FROM volunteers WHERE status='在册' ORDER BY name", fetchall=True) or []
    claimed = {r['volunteer_id'] for r in records if r['status'] == 'claimed'}
    pending = [v for v in volunteers if v['id'] not in claimed]
    return render_template_string(ADMIN, years=years, campaign=c, selected=selected, stock=stock,
                                  records=records, moves=moves, pending=pending, tz=MY_TZ)


@calendar_bp.route('/calendar/admin/action', methods=['POST'])
def action():
    if not _admin():
        return '无权限', 403
    op = request.form.get('op', '')
    year = request.form.get('year', type=int)
    qty = request.form.get('quantity', type=int)
    note = request.form.get('note', '').strip()[:200]
    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            if op == 'new':
                if not year or year < 2026 or year > 2200 or qty is None or qty < 0:
                    return '年度或库存数量不正确', 400
                cur.execute('SELECT 1 FROM calendar_campaigns WHERE year=%s', (year,))
                if cur.fetchone():
                    return '该年度已存在', 400
                cur.execute('UPDATE calendar_campaigns SET is_open=false')
                cur.execute('INSERT INTO calendar_campaigns(year,initial_stock,is_open) VALUES (%s,%s,true)', (year,qty))
            else:
                cur.execute('SELECT * FROM calendar_campaigns WHERE year=%s FOR UPDATE', (year,))
                c = cur.fetchone()
                if not c:
                    return '年度不存在', 404
                if op == 'toggle':
                    if not c['is_open']:
                        cur.execute('UPDATE calendar_campaigns SET is_open=false')
                    cur.execute('UPDATE calendar_campaigns SET is_open=%s WHERE year=%s', (not c['is_open'],year))
                elif op in ('restock','friend','outreach','return'):
                    if not c['is_open']:
                        return '该年度已关闭', 400
                    if qty is None or qty <= 0:
                        return '数量必须大于0', 400
                    vid = request.form.get('number','').strip()
                    v = _identity(vid) if op in ('outreach','return') else None
                    if op in ('outreach','return') and not v:
                        return '弘法领出或归还必须输入有效义工编号', 400
                    if op == 'return':
                        cur.execute("SELECT COALESCE(-SUM(quantity),0) AS issued FROM calendar_movements WHERE year=%s AND volunteer_id=%s AND kind IN ('outreach','return')", (year,v['编号']))
                        if qty > int(cur.fetchone()['issued']):
                            return '归还数量超过该义工尚未归还的领出数量', 400
                    delta = qty if op in ('restock','return') else -qty
                    if delta < 0 and _stock(cur,year,c['initial_stock']) < qty:
                        return '库存不足', 400
                    cur.execute('INSERT INTO calendar_movements(year,kind,quantity,volunteer_id,volunteer_name,note) VALUES (%s,%s,%s,%s,%s,%s)', (year,op,delta,v['编号'] if v else None,v['姓名'] if v else None,note))
                elif op == 'reset':
                    vid = request.form.get('number','').strip()
                    cur.execute('SELECT status FROM calendar_personal WHERE year=%s AND volunteer_id=%s', (year,vid))
                    r = cur.fetchone()
                    if not r:
                        return '找不到登记', 404
                    if r['status'] == 'claimed':
                        cur.execute("INSERT INTO calendar_movements(year,kind,quantity,volunteer_id,note) VALUES (%s,'correction',1,%s,'负责人撤销个人领取')", (year,vid))
                    cur.execute('DELETE FROM calendar_personal WHERE year=%s AND volunteer_id=%s', (year,vid))
                else:
                    return '操作无效', 400
    return redirect('/calendar/admin?year='+str(year))


PAGE = '''<!doctype html><html lang="zh"><meta name="viewport" content="width=device-width,initial-scale=1"><meta charset="utf-8"><title>日历领取</title>
<style>body{font-family:Arial,"Microsoft YaHei",sans-serif;background:#fff7f9;color:#253b2e;margin:0;padding:14px}.card{max-width:650px;margin:16px auto;background:white;border-radius:18px;padding:22px;box-shadow:0 3px 14px #0001}h1{font-size:25px}input,button{font-size:19px;padding:13px;border-radius:10px;border:1px solid #ccd6cd;box-sizing:border-box}input{width:100%;margin:7px 0 13px}button{background:#269651;color:white;border:0;font-weight:bold;cursor:pointer}.row{display:flex;gap:10px;flex-wrap:wrap}.row button{flex:1}.muted{color:#68766b}.danger{background:#777}a{color:#176a41}</style>
<div class="card"><a href="/">← 返回签到首页</a><h1>📅 {{ campaign.year if campaign else '' }}年日历领取</h1>
{% if campaign %}<p>📦 目前剩余：<b>{{stock}}</b> 本</p><p class="muted">请先从柜台领取日历，再登记「已领取1本」。登记不影响签到。</p>
<label>分会 / 编号</label><div class="row"><button id="branch" type="button" onclick="branch=this.textContent==='CHE'?'STW':'CHE';this.textContent=branch;this.style.background=branch==='CHE'?'#269651':'#d9534f'">CHE</button><input id="number" inputmode="numeric" placeholder="只输入数字，例如208" style="flex:4;min-width:130px"></div>
<label>义工 PIN</label><input id="pin" type="password" inputmode="numeric" placeholder="输入义工 PIN"><button onclick="lookup()">🔍 查询姓名</button><p id="message"></p><div class="row" id="choices" hidden><button onclick="submitStatus('claimed')">✅ 已领取1本</button></div>
{% else %}<p>目前未开放日历领取。</p>{% endif %}</div>
<div class="card"><h2>🔐 负责人管理</h2><form method="post" action="/calendar/admin-login"><input type="password" name="admin_pin" placeholder="管理员 PIN" required><button>进入管理</button></form>{% if request.args.get('error') %}<p>管理员 PIN 不正确</p>{% endif %}</div>
<script>let branch='CHE';function data(){let f=new FormData();f.set('number',branch+'-'+document.getElementById('number').value.trim());f.set('pin',document.getElementById('pin').value);return f}
async function lookup(){document.getElementById('choices').hidden=true;try{let r=await fetch('/calendar/lookup',{method:'POST',body:data()});let x=await r.json();document.getElementById('message').textContent=x.ok?'👤 '+x.name+'　'+({claimed:'✅ 已领取'}[x.status]||'尚未领取'):x.message;document.getElementById('choices').hidden=!(x.ok&&!x.status)}catch(e){document.getElementById('message').textContent='网络错误，请重试'}}
async function submitStatus(status){let f=data();f.set('status',status);document.getElementById('choices').hidden=true;try{let r=await fetch('/calendar/personal',{method:'POST',body:f});let x=await r.json();document.getElementById('message').textContent=x.message;if(!x.ok)await lookup()}catch(e){document.getElementById('message').textContent='网络错误，请重新查询确认是否已登记'}}
</script></html>'''

ADMIN = '''<!doctype html><html lang="zh"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>日历管理</title><style>body{font-family:Arial,"Microsoft YaHei",sans-serif;background:#f4f7f3;margin:0;padding:16px;color:#25372a}main{max-width:1000px;margin:auto}.card{background:white;border-radius:14px;padding:18px;margin:14px 0}input,button,select{font-size:17px;padding:10px;border:1px solid #ccd5cb;border-radius:9px;margin:4px}button{background:#258c4a;color:white;border:0;cursor:pointer}.names{display:grid;grid-template-columns:repeat(auto-fill,minmax(145px,1fr));gap:8px}.person{padding:10px;background:#edf6ef;border-radius:9px}.small{font-size:13px;color:#667}h2{font-size:21px}form{margin:9px 0}table{border-collapse:collapse;width:100%}td,th{border-bottom:1px solid #ddd;padding:8px;text-align:left}</style><main><a href="/calendar">← 返回日历领取</a><h1>📅 日历管理</h1>
<div class="card"><h2>开启新年度</h2><form method="post" action="/calendar/admin/action"><input type="hidden" name="op" value="new"><input type="number" name="year" min="2026" max="2200" placeholder="年度，如2027" required><input type="number" name="quantity" min="0" placeholder="初始库存" required><button onclick="return confirm('确定开启新年度？当前年度会关闭。')">开启年度</button></form></div>
{% if campaign %}<div class="card"><form method="get"><select name="year" onchange="this.form.submit()">{% for y in years %}<option value="{{y.year}}" {% if y.year==selected %}selected{% endif %}>{{y.year}}年</option>{% endfor %}</select></form><h2>{{selected}}年　{{'🟢 开放' if campaign.is_open else '🔴 已关闭'}}</h2><h2>📦 剩余 {{stock}} 本</h2><p>初始 {{campaign.initial_stock}} 本 · 已领取 {{records|selectattr('status','equalto','claimed')|list|length}} 人 · 未领取 {{pending|length}} 人</p><form method="post" action="/calendar/admin/action"><input type="hidden" name="op" value="toggle"><input type="hidden" name="year" value="{{selected}}"><button>{{'🔒 关闭领取' if campaign.is_open else '🟢 重新开放'}}</button></form></div>
<div class="card"><h2>📦 库存与发放</h2><form method="post" action="/calendar/admin/action"><input type="hidden" name="year" value="{{selected}}"><select name="op"><option value="restock">补充库存</option><option value="friend">佛友领取（无需姓名）</option><option value="outreach">弘法领出</option><option value="return">弘法归还</option></select><input name="number" placeholder="弘法义工完整编号 CHE-208"><input name="quantity" type="number" min="1" placeholder="本数" required><input name="note" placeholder="备注（可选）"><button>确认登记</button></form><p class="small">弘法领出与归还须填写义工编号；归还不能超过该义工未归还数量。</p></div>
<div class="card"><h2>🟢 已领取</h2><div class="names">{% for r in records if r.status=='claimed' %}<div class="person">{{r.volunteer_name}}<div class="small">{{r.volunteer_id}}</div><form method="post" action="/calendar/admin/action" onsubmit="return confirm('确定撤销登记并退回1本库存？')"><input type="hidden" name="op" value="reset"><input type="hidden" name="year" value="{{selected}}"><input type="hidden" name="number" value="{{r.volunteer_id}}"><button>撤销</button></form></div>{% endfor %}</div></div>
<div class="card"><h2>🟠 未领取（{{pending|length}}）</h2><div class="names">{% for v in pending %}<div class="person">{{v.name}}<div class="small">{{v.id}}</div></div>{% endfor %}</div></div>
<div class="card"><h2>📋 最近库存记录</h2><table><tr><th>时间</th><th>类别</th><th>数量变化</th><th>义工</th></tr>{% for m in moves %}<tr><td>{{m.created_at.astimezone(tz).strftime('%m/%d %H:%M') if tz else m.created_at}}</td><td>{{m.kind}}</td><td>{{m.quantity}}</td><td>{{m.volunteer_name or ''}}</td></tr>{% endfor %}</table></div>
{% endif %}</main></html>'''
