"""XFZ 经文组合登记 / Offering and burning ledger. Register as Flask blueprint."""
from datetime import date, datetime
from zoneinfo import ZoneInfo
import os
import hmac
from flask import Blueprint, request, redirect, flash, render_template_string, jsonify, send_file
from db import db_query
from attendance_service import find_volunteer, verify_pin_for_volunteer
from io import BytesIO

xfz_bp = Blueprint('xfz', __name__, url_prefix='/xfz')
TZ = ZoneInfo('Asia/Kuala_Lumpur')
WORDS = {
 'zh': {'title':'经文组合记录','received':'累计结缘','burned':'累计烧送','balance':'目前结存','offer':'结缘登记','burn':'烧送登记','records':'记录查询','name':'结缘者编号／姓名','date':'日期','qty':'张数','witness':'见证人义工编号','witness_name':'见证人姓名','officer':'烧送负责人编号','pin':'负责人 PIN','notes':'备注','submit':'确认登记','type':'类型','person':'姓名','back':'返回首页','download':'下载 Excel','initial':'期初库存','admin':'管理员调整期初库存','adminpin':'管理员 PIN','save':'保存','no_records':'暂无记录','invalid':'请检查输入内容','verified':'在册义工','filter':'筛选','all':'全部','start':'开始日期','end':'结束日期','search':'查询','diet':'饮食类别','vegan':'全素','nonvegan':'非全素'},
 'en': {'title':'XFZ Records','received':'Total Received','burned':'Total Burned','balance':'Current Balance','offer':'Offering Registration','burn':'Burning Registration','records':'Records','name':'Contributor ID / Name','date':'Date','qty':'Sheets','witness':'Witness Volunteer ID','witness_name':'Witness Name','officer':'Burning Officer ID','pin':'Officer PIN','notes':'Notes','submit':'Submit','type':'Type','person':'Name','back':'Back Home','download':'Export Excel','initial':'Opening Balance','admin':'Set Opening Balance (Admin)','adminpin':'Admin PIN','save':'Save','no_records':'No records yet','invalid':'Please check your input','verified':'Active volunteer','filter':'Filter','all':'All','start':'Start Date','end':'End Date','search':'Search','diet':'Diet Category','vegan':'Fully Vegetarian','nonvegan':'Not Fully Vegetarian'}
}
PAGE = """<!doctype html><html lang="{{lang}}"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{{t.title}}</title><style>\nbody{background:#fffaf0;font-family:Arial,\'Microsoft YaHei\',sans-serif;color:#3b3321;margin:0}.wrap{max-width:1000px;margin:20px auto;padding:12px}.card{background:white;border:1px solid #f2dfa7;border-radius:16px;padding:22px;margin-bottom:18px;box-shadow:0 3px 12px #0001}h1{color:#8b6000}h2{margin-top:0}button,.btn{background:#f4c84a;border:0;border-radius:9px;padding:12px 18px;color:#342900;font-size:16px;font-weight:bold;cursor:pointer;text-decoration:none;display:inline-block}input,select{padding:11px;border:1px solid #c9bd9f;border-radius:8px;font-size:16px;max-width:100%;box-sizing:border-box}label{display:block;margin:12px 0 5px}form{margin:0}.stats{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}.stat{text-align:center;background:#fff4cf;border-radius:12px;padding:16px}.stat strong{font-size:30px;display:block;margin-top:8px}.grid{display:grid;grid-template-columns:1fr 1fr;gap:16px}.wide{width:100%}.msg{padding:12px;border-radius:8px;background:#fff0c2;margin-bottom:12px}table{width:100%;border-collapse:collapse}td,th{border-bottom:1px solid #eee;padding:10px;text-align:left}th{background:#fff4cf}.scroll{overflow:auto}.muted{color:#73684d;font-size:14px}.idrow{display:flex;gap:9px;align-items:stretch}.idrow input{flex:1;min-width:0}.branch{width:79px;flex-shrink:0;color:white;padding:10px 5px}.branch.che{background:#159447}.branch.stw{background:#dc3545}.verified{min-height:22px;margin:6px 0 0;color:#198754;font-weight:bold}.verified.bad{color:#c0392b}.choice{display:flex;gap:10px;flex-wrap:wrap;margin:9px 0}.choice button{font-size:14px;padding:9px 12px;background:#f3ead1}.choice button.selected{background:#f4c84a}.diet-options{display:flex;gap:12px;flex-wrap:wrap;margin:12px 0 16px}.diet-options input{position:absolute;opacity:0;width:1px;height:1px}.diet-options label{display:inline-flex;align-items:center;justify-content:center;min-width:130px;margin:0;padding:12px 18px;border:2px solid #159447;border-radius:10px;background:#fff;color:#159447;cursor:pointer;font-weight:bold}.diet-options label:nth-child(2){border-color:#dc3545;color:#dc3545}.diet-options label:has(input:checked){background:#159447;border-color:#159447;color:#fff}.diet-options label:nth-child(2):has(input:checked){background:#dc3545;border-color:#dc3545;color:#fff}.diet-options label:has(input:focus-visible){outline:3px solid #7e9eff;outline-offset:3px}.hide{display:none}@media(max-width:700px){.grid,.stats{grid-template-columns:1fr}.card{padding:15px}}\n.record-filters{display:grid;grid-template-columns:1fr 1.35fr 1.35fr auto;gap:12px;align-items:end;margin:10px 0}.record-filter label{font-weight:bold;margin:0 0 7px}.record-submit button{min-height:44px;white-space:nowrap}.record-filter input,.record-filter select{min-height:44px}@media(max-width:700px){.record-filters{grid-template-columns:1fr 1fr}.record-filters .record-filter:first-child,.record-submit{grid-column:1/-1}.record-submit button{width:100%}}\n/* 记录表：整行按进出着色；默认仅展示最近八笔 */\n.xfz-records tbody tr.xfz-offer{background:#e5f6e9;color:#146b37}.xfz-records tbody tr.xfz-burn{background:#fde7e7;color:#a32626}.xfz-records tbody tr.xfz-offer td,.xfz-records tbody tr.xfz-burn td{border-bottom:2px solid #fff}.xfz-records .xfz-qty{font-weight:800;font-size:17px}.xfz-records tbody tr.xfz-extra{display:none}.xfz-records.xfz-expanded tbody tr.xfz-extra{display:table-row}.xfz-record-toggle{margin-top:12px;background:#fff4cf;border:1px solid #e6cb78;color:#6c4d00}\n\n</style></head><body><main class="wrap"><div style="display:flex;justify-content:space-between;align-items:center;gap:8px"><h1>📜 {{t.title}}</h1><span><a href="?lang=zh">中文</a> | <a href="?lang=en">English</a></span></div><p><a class="btn" href="/">← {{t.back}}</a></p>\n{% with messages=get_flashed_messages() %}{% for m in messages %}<div class="msg">{{m}}</div>{% endfor %}{% endwith %}\n<div class="card stats"><div class="stat">{{t.received}}<strong>{{received}}</strong></div><div class="stat">{{t.burned}}<strong>{{burned}}</strong></div><div class="stat">{{t.balance}}<strong>{{balance}}</strong></div></div>\n<div class="grid"><section class="card"><h2>🟢 {{t.offer}}</h2><form action="/xfz/offer" method="post" onsubmit="return validateOffer()"><input type="hidden" name="lang" value="{{lang}}"><label>{{t.date}}</label><input class="wide" type="date" name="entry_date" value="{{today}}" required>\n<label>{{\'结缘者\' if lang==\'zh\' else \'Contributor\'}}</label><div class="choice"><button type="button" id="mode_vol" class="selected" onclick="setMode(\'vol\')">{{\'义工编号\' if lang==\'zh\' else \'Volunteer ID\'}}</button><button type="button" id="mode_name" onclick="setMode(\'name\')">{{\'佛友姓名\' if lang==\'zh\' else \'Contributor Name\'}}</button></div><input type="hidden" name="person_mode" id="person_mode" value="vol"><div id="contributor_vol"><div class="idrow"><button type="button" class="branch che" data-for="person_id" data-branch="CHE" onclick="toggleBranch(this)">CHE</button><input id="person_id" name="person_id" inputmode="numeric" autocomplete="off" placeholder="{{\'只输入数字\' if lang==\'zh\' else \'Numbers only\'}}"></div><p class="verified" id="person_result"></p></div><div id="contributor_name" class="hide"><input class="wide" name="person_name" id="person_name" maxlength="120" placeholder="{{\'输入佛友姓名\' if lang==\'zh\' else \'Enter contributor name\'}}"></div>\n<div class="diet-options" role="group" aria-label="{{t.diet}}"><label><input type="radio" name="diet_type" value="vegan" required><span>{{t.vegan}}</span></label><label><input type="radio" name="diet_type" value="nonvegan" required><span>{{t.nonvegan}}</span></label></div><label>{{t.qty}}</label><input class="wide" type="number" name="quantity" min="1" max="100000" value="1" required><label>{{t.witness}}</label><div class="idrow"><button type="button" class="branch che" data-for="witness_id" data-branch="CHE" onclick="toggleBranch(this)">CHE</button><input name="witness_id" id="witness_id" inputmode="numeric" autocomplete="off" placeholder="{{\'只输入数字\' if lang==\'zh\' else \'Numbers only\'}}" required></div><p class="verified" id="witness_result"></p><label>{{t.notes}}</label><input class="wide" name="notes" maxlength="500"><p><button type="submit">{{t.submit}} +</button></p></form></section>\n<section class="card"><h2>🔴 {{t.burn}}</h2><p class="muted">{% if lang==\'zh\' %}每星期三通常烧送7张；在册义工输入本人编号及 PIN 即可登记。{% else %}Usually 7 sheets each Wednesday; active volunteers can record with their own ID and PIN.{% endif %}</p><form action="/xfz/burn" method="post" onsubmit="normalizeSubmittedId(\'officer_id\')"><input type="hidden" name="lang" value="{{lang}}"><label>{{t.date}}</label><input class="wide" type="date" name="entry_date" value="{{today}}" required><label>{{t.officer}}</label><div class="idrow"><button type="button" class="branch che" data-for="officer_id" data-branch="CHE" onclick="toggleBranch(this)">CHE</button><input name="officer_id" id="officer_id" inputmode="numeric" autocomplete="off" placeholder="{{\'只输入数字\' if lang==\'zh\' else \'Numbers only\'}}" required></div><p class="verified" id="officer_result"></p><label>{{t.pin}}</label><input class="wide" type="password" name="officer_pin" required><label>{{t.qty}}</label><input class="wide" type="number" name="quantity" min="1" max="100000" value="7" required><label>{{t.notes}}</label><input class="wide" name="notes" maxlength="500"><p><button type="submit">{{t.submit}} −</button></p></form></section></div>\n<section class="card"><h2>📋 {{t.records}}</h2><form method="get" onsubmit="return validateRecordDates(this)"><input type="hidden" name="lang" value="{{lang}}"><div class="record-filters"><div class="record-filter"><label for="record-kind">{{t.type}} / {{'Type' if lang=='zh' else '类型'}}</label><select id="record-kind" name="kind" class="wide"><option value="">{{t.all}}</option><option value="offer" {% if kind==\'offer\' %}selected{% endif %}>{{t.offer}}</option><option value="burn" {% if kind==\'burn\' %}selected{% endif %}>{{t.burn}}</option></select></div><div class="record-filter"><label for="record-start">{{t.start}} / {{'From' if lang=='zh' else '开始日期'}}</label><input id="record-start" class="wide" type="date" name="start" value="{{start}}"></div><div class="record-filter"><label for="record-end">{{t.end}} / {{'To' if lang=='zh' else '结束日期'}}</label><input id="record-end" class="wide" type="date" name="end" value="{{end}}"></div><div class="record-filter record-submit"><button type="submit">🔎 {{t.search}}</button></div></div><p class="muted">{{'日期留空表示不限制范围。' if lang=='zh' else 'Leave dates blank to search without date limits.'}}</p></form><p><a class="btn" href="/xfz/export?lang={{lang}}">{{t.download}}</a></p><div class="scroll"><table class="xfz-records" id="xfz-records"><thead><tr><th>{{t.date}}</th><th>{{t.type}}</th><th>{{t.person}}</th><th>{{t.diet}}</th><th>{{t.qty}}</th><th>{{t.witness_name}}</th><th>{{t.notes}}</th></tr></thead><tbody>{% for r in rows %}<tr class="{{'xfz-offer' if r.kind=='offer' else 'xfz-burn'}}{{' xfz-extra' if loop.index > 8 else ''}}"><td>{{r.entry_date}}</td><td>{{t.offer if r.kind==\'offer\' else t.burn}}</td><td>{{r.person_name}}</td><td>{{(t.vegan if r.diet_type=='vegan' else t.nonvegan if r.diet_type=='nonvegan' else '—') if r.kind=='offer' else '—'}}</td><td class="xfz-qty">{{\'+\' if r.kind==\'offer\' else \'-\'}}{{r.quantity}}</td><td>{{r.witness_name or \'—\'}}</td><td>{{r.notes or \'\'}}</td></tr>{% else %}<tr><td colspan="7">{{t.no_records}}</td></tr>{% endfor %}</tbody></table></div>{% if rows|length > 8 %}<button type="button" id="xfz-record-toggle" class="xfz-record-toggle" aria-expanded="false" onclick="toggleXfzRecords()">{{'显示其余记录（共' ~ (rows|length) ~ '笔）▼' if lang=='zh' else 'Show all records (' ~ (rows|length) ~ ') ▼'}}</button>{% endif %}</section>\n<details class="card"><summary>{{t.admin}}</summary><form action="/xfz/opening" method="post"><input type="hidden" name="lang" value="{{lang}}"><label>{{t.initial}}</label><input name="opening" type="number" min="0" value="{{opening}}" required><label>{{t.adminpin}}</label><input type="password" name="admin_pin" required><p><button>{{t.save}}</button></p></form></details></main><script>\nconst isEn={{\'true\' if lang==\'en\' else \'false\'}};const valid={};const timers={};\nfunction branchOf(id){return document.querySelector(\'[data-for="\'+id+\'"]\')}\nfunction fullId(id){const el=document.getElementById(id);const v=el.value.trim();if(!v)return \'\';if(/^(CHE|STW)-/i.test(v))return v.toUpperCase();if(!/^\\d+$/.test(v))return \'\';if(/^8\\d{2}$/.test(v))return v;return branchOf(id).dataset.branch+\'-\'+Number(v)}\nfunction toggleBranch(btn){const next=btn.dataset.branch===\'CHE\'?\'STW\':\'CHE\';btn.dataset.branch=next;btn.textContent=next;btn.classList.toggle(\'che\',next===\'CHE\');btn.classList.toggle(\'stw\',next===\'STW\');lookup(btn.dataset.for)}\nfunction setMode(mode){document.getElementById(\'person_mode\').value=mode;document.getElementById(\'contributor_vol\').classList.toggle(\'hide\',mode!==\'vol\');document.getElementById(\'contributor_name\').classList.toggle(\'hide\',mode!==\'name\');document.getElementById(\'mode_vol\').classList.toggle(\'selected\',mode===\'vol\');document.getElementById(\'mode_name\').classList.toggle(\'selected\',mode===\'name\');document.getElementById(\'person_id\').required=mode===\'vol\';document.getElementById(\'person_name\').required=mode===\'name\'}\nasync function lookup(id){clearTimeout(timers[id]);valid[id]=false;const result=document.getElementById(id.replace(\'_id\',\'_result\'));result.textContent=\'\';result.classList.remove(\'bad\');const value=fullId(id);if(!value)return;timers[id]=setTimeout(async()=>{try{const r=await fetch((id===\'person_id\'?\'/xfz/contributor/\':\'/xfz/witness/\')+encodeURIComponent(value));const d=await r.json();if(fullId(id)!==value)return;valid[id]=!!d.ok;result.textContent=d.ok?\'✓ \'+d.name:(id===\'person_id\'?(isEn?\'✗ Contributor not found\':\'✗ 找不到佛友或义工\'):(isEn?\'✗ Volunteer not found\':\'✗ 找不到在册义工\'));result.classList.toggle(\'bad\',!d.ok)}catch(e){result.textContent=isEn?\'Lookup failed\':\'查询失败\';result.classList.add(\'bad\')}},320)}\n[\'person_id\',\'witness_id\',\'officer_id\'].forEach(id=>document.getElementById(id).addEventListener(\'input\',()=>lookup(id)));\nfunction normalizeSubmittedId(id){const value=fullId(id);if(value)document.getElementById(id).value=value}\nfunction validateOffer(){if(document.getElementById(\'person_mode\').value===\'vol\'&&!valid.person_id){alert(isEn?\'Please verify contributor ID\':\'请先确认结缘者编号\');return false}if(!valid.witness_id){alert(isEn?\'Please verify witness ID\':\'请先确认见证人编号\');return false}if(document.getElementById('person_mode').value==='vol')normalizeSubmittedId('person_id');normalizeSubmittedId('witness_id');return true}\nfunction toggleXfzRecords(){const table=document.getElementById('xfz-records');const btn=document.getElementById('xfz-record-toggle');const expanded=table.classList.toggle('xfz-expanded');btn.setAttribute('aria-expanded',String(expanded));btn.textContent=expanded?(isEn?'Show latest 8 ▲':'收起记录，只显示最近8笔 ▲'):(isEn?'Show all records ▼':'显示其余记录 ▼')}\nfunction validateRecordDates(form){const start=form.elements.start.value;const end=form.elements.end.value;if(start&&end&&start>end){alert(isEn?'Start date cannot be later than end date.':'开始日期不能晚于结束日期。');return false}return true}\nsetMode(\'vol\');\n</script></body></html>"""

def language():
    lang=request.values.get('lang') or request.cookies.get('lang','zh')
    return lang if lang in WORDS else 'zh'

def rows_query(sql, params=(), one=False):
    return db_query(sql,params,fetchone=True) if one else (db_query(sql,params,fetchall=True) or [])

def valid_date(raw):
    return date.fromisoformat(raw)

def positive(raw):
    n=int(raw)
    if not 1<=n<=100000: raise ValueError('quantity')
    return n

def active_volunteer(raw):
    v=find_volunteer((raw or '').strip())
    if not v: return None
    vid=str(v.get('编号') or '').strip()
    if not vid: return None
    check=rows_query("SELECT id,name FROM volunteers WHERE id=%s AND status='在册'",(vid,),True)
    return (v,check) if check else None

def find_contributor(raw):
    """结缘者可为月费佛友或非在册的历史义工；不授予见证人资格。"""
    member_id = str(raw or '').strip().upper()
    if not member_id:
        return None
    member = rows_query(
        "SELECT member_id AS id, name FROM members WHERE upper(member_id)=upper(%s) LIMIT 1",
        (member_id,), True
    )
    if member:
        return member
    person = rows_query(
        "SELECT id, name FROM volunteers WHERE upper(id)=upper(%s) "
        "OR upper(coalesce(member_id,''))=upper(%s) LIMIT 1",
        (member_id, member_id), True
    )
    return person


@xfz_bp.get('/contributor/<path:person_id>')
def contributor_lookup(person_id):
    found = find_contributor(person_id)
    return jsonify(ok=bool(found), name=found['name'] if found else '')


@xfz_bp.get('/witness/<path:vol_id>')
def witness_lookup(vol_id):
    found=active_volunteer(vol_id)
    return jsonify(ok=bool(found),name=found[1]['name'] if found else '')

@xfz_bp.get('')
@xfz_bp.get('/')
def index():
    lang=language()
    opening=rows_query("SELECT opening_balance FROM xfz_settings WHERE id=1",(),True)
    totals=rows_query("SELECT COALESCE(SUM(quantity) FILTER (WHERE kind='offer'),0) AS received, COALESCE(SUM(quantity) FILTER (WHERE kind='burn'),0) AS burned FROM xfz_entries",(),True)
    received=int(totals['received']);burned=int(totals['burned']);initial=int(opening['opening_balance']) if opening else 0
    kind=request.args.get('kind','');start=request.args.get('start','');end=request.args.get('end','')
    where=[];params=[]
    if kind in ('offer','burn'): where.append('kind=%s');params.append(kind)
    try:
        if start: where.append('entry_date >= %s');params.append(valid_date(start))
        if end: where.append('entry_date <= %s');params.append(valid_date(end))
    except ValueError:
        flash('❌ 日期格式不正确' if lang=='zh' else '❌ Invalid date format')
        return redirect('/xfz?lang='+lang)
    if start and end and valid_date(start) > valid_date(end):
        flash('❌ 开始日期不能晚于结束日期' if lang=='zh' else '❌ Start date cannot be later than end date')
        return redirect('/xfz?lang='+lang)
    sql='SELECT * FROM xfz_entries'+(' WHERE '+' AND '.join(where) if where else '')+' ORDER BY entry_date DESC,id DESC LIMIT 300'
    return render_template_string(PAGE,t=WORDS[lang],lang=lang,today=datetime.now(TZ).date().isoformat(),opening=initial,received=received,burned=burned,balance=initial+received-burned,rows=rows_query(sql,tuple(params)),kind=kind,start=start,end=end)

@xfz_bp.post('/offer')
def offer():
    lang=language()
    try:
        day=valid_date(request.form['entry_date']);quantity=positive(request.form['quantity'])
        diet_type=request.form.get('diet_type','').strip()
        if diet_type not in ('vegan','nonvegan'): raise ValueError('invalid diet type')
        mode=request.form.get('person_mode','vol')
        if mode=='vol':
            contributor=find_contributor(request.form.get('person_id',''))
            if not contributor: raise ValueError('invalid contributor')
            name=contributor['name']
        elif mode=='name':
            name=request.form.get('person_name','').strip()
        else: raise ValueError('invalid mode')
        witness=request.form.get('witness_id','').strip()
        found=active_volunteer(witness)
        if not name or len(name)>120 or not found: raise ValueError('name or witness')
        db_query('INSERT INTO xfz_entries (entry_date,kind,person_name,quantity,witness_id,witness_name,notes,diet_type) VALUES (%s,%s,%s,%s,%s,%s,%s,%s)',(day,'offer',name,quantity,found[1]['id'],found[1]['name'],request.form.get('notes','')[:500],diet_type))
        flash('✅ 结缘登记成功' if lang=='zh' else '✅ Offering recorded')
    except (ValueError,KeyError) as e: flash('❌ 资料不正确，请确认见证人是在册义工' if lang=='zh' else '❌ Invalid entry or witness')
    return redirect('/xfz?lang='+lang)

@xfz_bp.post('/burn')
def burn():
    lang=language()
    try:
        day=valid_date(request.form['entry_date']);qty=positive(request.form['quantity'])
        found=active_volunteer(request.form.get('officer_id',''))
        if not found:
            flash('❌ 找不到在册义工，请检查负责人编号' if lang=='zh' else '❌ Active volunteer not found');return redirect('/xfz?lang='+lang)
        if not verify_pin_for_volunteer(found[0],request.form.get('officer_pin','')):
            flash('❌ 负责人 PIN 错误' if lang=='zh' else '❌ Incorrect officer PIN');return redirect('/xfz?lang='+lang)
        # One transaction with a row lock, so two simultaneous requests cannot overspend stock.
        from db import get_conn
        with get_conn() as conn:
            with conn.cursor() as cur:
                cur.execute('SELECT opening_balance FROM xfz_settings WHERE id=1 FOR UPDATE')
                initial=cur.fetchone()
                if not initial: raise ValueError('settings missing')
                cur.execute("SELECT COALESCE(SUM(CASE WHEN kind='offer' THEN quantity ELSE -quantity END),0) AS net_quantity FROM xfz_entries")
                # db.py may return dict rows (RealDictCursor) rather than tuples.
                # Use explicit aliases and read by column name in either case.
                summed = cur.fetchone()
                opening_qty = initial['opening_balance'] if isinstance(initial, dict) else initial[0]
                net_qty = summed['net_quantity'] if isinstance(summed, dict) else summed[0]
                stock = int(opening_qty) + int(net_qty)
                if qty > stock:
                    raise ValueError(('库存不足：目前只有 %s 张' if lang == 'zh' else 'Insufficient stock: %s sheets available') % stock)
                cur.execute('SELECT COUNT(*) AS burn_count FROM xfz_entries WHERE kind=%s AND entry_date=%s', ('burn', day))
                count_row = cur.fetchone()
                burn_count = count_row['burn_count'] if isinstance(count_row, dict) else count_row[0]
                if int(burn_count):
                    raise ValueError('该日期已有烧送记录，请勿重复登记' if lang == 'zh' else 'A burning record already exists for this date')
                cur.execute('INSERT INTO xfz_entries (entry_date,kind,person_name,quantity,officer_id,notes) VALUES (%s,%s,%s,%s,%s,%s)',(day,'burn',found[1]['name'],qty,found[1]['id'],request.form.get('notes','')[:500]))
        flash('✅ 烧送登记成功' if lang=='zh' else '✅ Burning recorded')
    except (ValueError,KeyError) as e:
        flash(('❌ 烧送失败：'+str(e)) if lang=='zh' else ('❌ Could not record burning: '+str(e)))
    return redirect('/xfz?lang='+lang)

@xfz_bp.post('/opening')
def opening():
    lang=language()
    # 与签到管理员中心共用同一 PIN 来源，不再使用 XFZ_ADMIN_PIN。
    # 在请求时读取，避免模块初始化顺序引起循环导入。
    from admin_web import ADMIN_PIN
    submitted_pin = request.form.get('admin_pin', '').strip()
    if not ADMIN_PIN or not hmac.compare_digest(submitted_pin, str(ADMIN_PIN)):
        flash('❌ 签到管理员 PIN 错误' if lang=='zh' else '❌ Incorrect attendance admin PIN')
    else:
        try:
            val=int(request.form['opening'])
            if val<0: raise ValueError()
            db_query('UPDATE xfz_settings SET opening_balance=%s WHERE id=1',(val,))
            flash('✅ 已更新期初库存' if lang=='zh' else '✅ Opening balance updated')
        except ValueError: flash('❌ 数量错误' if lang=='zh' else '❌ Invalid quantity')
    return redirect('/xfz?lang='+lang)

@xfz_bp.get('/export')
def export():
    import pandas as pd
    data=rows_query('SELECT entry_date,kind,person_name,quantity,witness_id,witness_name,officer_id,notes,created_at,diet_type FROM xfz_entries ORDER BY entry_date,id')
    columns=['entry_date','kind','person_name','quantity','witness_id','witness_name','officer_id','notes','created_at','diet_type']
    df=pd.DataFrame(data,columns=columns)
    for col in ('entry_date','created_at'):
        if col in df: df[col]=df[col].astype(str)
    output=BytesIO()
    with pd.ExcelWriter(output,engine='openpyxl') as writer: df.to_excel(writer,index=False,sheet_name='XFZ')
    output.seek(0)
    return send_file(output,as_attachment=True,download_name='XFZ_records.xlsx',mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
