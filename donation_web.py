"""Independent cleaning-supplies donation signup module.

Register in app.py only after testing:
    from donation_web import donation_bp
    app.register_blueprint(donation_bp)

Set DONATION_ADMIN_PIN in Render environment before enabling admin actions.
Existing attendance, schedule and database modules are not modified.
"""
from __future__ import annotations

import os
import re
import hmac
import io
from PIL import Image, UnidentifiedImageError
from datetime import datetime, timedelta, time as dt_time
from zoneinfo import ZoneInfo

from flask import Blueprint, jsonify, request, render_template_string, session, redirect, url_for, Response
from psycopg2 import IntegrityError
from db import get_conn

MY_TZ = ZoneInfo('Asia/Kuala_Lumpur')
donation_bp = Blueprint('donation', __name__, url_prefix='/donation')

STYLE = '''<style>
*{box-sizing:border-box}body{margin:0;background:#f6f8f4;color:#25352a;font-family:Arial,"Microsoft YaHei",sans-serif;font-size:18px}
main{max-width:1040px;margin:auto;padding:18px}h1{text-align:center;font-size:clamp(26px,5vw,36px)}h2{font-size:23px}
.card{background:white;border-radius:17px;padding:20px;margin:16px 0;box-shadow:0 3px 12px #0001}
.grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:14px}.item{overflow:hidden;padding:0}
.item img,.placeholder{width:100%;height:260px;object-fit:contain;background:#f3f5f0;display:grid;place-items:center;font-size:54px}
.inner{padding:16px}.names{font-size:19px;color:#1457a5;font-weight:700;line-height:1.7;min-height:32px;overflow-wrap:anywhere}.names .names-label{color:#657168;font-weight:400}.names .empty{color:#89938c;font-weight:400}.name-list{margin-top:5px}.name-entry{display:block;color:#1457a5;font-weight:700;line-height:1.65}.lookup-result{min-height:30px;font-size:21px;font-weight:bold;color:#087a36;margin:5px 0 12px}.lookup-result.error{color:#b74343}.id-row{display:flex;gap:10px;align-items:center}.id-row .branch-option{flex:0 0 115px;margin:0}.id-row input{flex:1;min-width:0;margin:0}.admin-link{text-align:right;margin:0 0 8px}.admin-link .btn{font-size:15px;padding:8px 12px}.branch-options{display:flex;gap:12px;margin:12px 0}.branch-option{flex:1;background:#e8eee9;color:#315445;border:2px solid #b8c9bc}.branch-option.active{background:#21865a;color:#fff;border-color:#21865a}
button,.btn{display:inline-block;border:0;border-radius:12px;padding:13px 17px;font-size:18px;font-weight:bold;cursor:pointer;background:#21865a;color:white;text-decoration:none;text-align:center}
button:disabled{background:#b8b8b8;cursor:not-allowed}.danger{background:#b74343}.secondary{background:#657581}
input,select,textarea{padding:12px;font-size:18px;border:1px solid #b7c9ba;border-radius:10px;width:100%;margin:7px 0 12px}
label{display:block;font-weight:bold;margin-top:10px}.muted{color:#637168;font-size:15px}.notice{background:#fff6dc;padding:15px;border-radius:12px}
.row{display:flex;gap:12px;flex-wrap:wrap;align-items:center}.row>*{flex:1}.status{font-weight:bold;color:#187c4b}.full{color:#b63b3b}
@media(max-width:850px){.grid{grid-template-columns:repeat(2,minmax(0,1fr))}}@media(max-width:360px){.grid{grid-template-columns:1fr}}
</style>'''

VOL_HTML = '''<!doctype html><html lang="zh"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>卫生用品结缘</title>''' + STYLE + '''<main><p class="admin-link"><a class="btn secondary" href="/donation/admin">⚙️ 负责人入口</a></p><h1>🌸 观音堂卫生用品结缘</h1>
<div class="card" id="login"><label for="vid">义工编号 / Volunteer ID</label><div class="id-row"><button type="button" class="branch-option active" id="branchBtn" onclick="switchBranch()">CHE</button><input id="vid" placeholder="输入号码，例如 208 或 821" inputmode="numeric" autocomplete="off" maxlength="12" oninput="scheduleLookup()" onkeydown="if(event.key==='Enter'){event.preventDefault();identify()}"></div>
<div id="lookupResult" class="lookup-result" aria-live="polite"></div><button id="continueBtn" onclick="identify()">确认进入结缘 / Continue</button><p class="muted">点击 CHE 可切换 STW；输入号码后自动显示姓名，按 Enter 即可进入。不需要 PIN。</p></div>
<div class="card" id="welcome" style="display:none"><h2 id="who"></h2><button class="secondary" onclick="logout()">更换编号 / Change ID</button></div>
<div class="notice" id="message">请输入义工编号后查看结缘物品。</div><div id="items" class="grid"></div>
<p class="muted" style="text-align:center">页面会自动更新名额和结缘名单，无需刷新。</p></main>
<script>
let person=null, busy=false, branch='CHE', lookupTimer=null, lookupSeq=0, found=null, foundKey='';
function switchBranch(){branch=branch==='CHE'?'STW':'CHE';document.getElementById('branchBtn').textContent=branch;resetLookup();scheduleLookup()}
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
function msg(s){document.getElementById('message').textContent=s}
async function api(url,options){const r=await fetch(url,{cache:'no-store',...options});return r.json()}
function currentId(){const raw=document.getElementById('vid').value.trim();if(!/^\\d+$/.test(raw))return '';return /^8\\d{2}$/.test(raw)?raw:branch+'-'+raw}
function resetLookup(){lookupSeq++;clearTimeout(lookupTimer);found=null;foundKey='';document.getElementById('lookupResult').textContent='';document.getElementById('lookupResult').classList.remove('error')}
function scheduleLookup(){resetLookup();const id=currentId();if(!id)return;lookupTimer=setTimeout(()=>lookupName(id),350)}
async function lookupName(id){const seq=++lookupSeq;const result=document.getElementById('lookupResult');result.textContent='正在查询姓名…';
try{const d=await api('/donation/api/identify',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({volunteer_id:id})});
if(seq!==lookupSeq||currentId()!==id)return;
if(!d.ok){result.textContent=d.message||'找不到义工';result.classList.add('error');return}
found=d.volunteer;foundKey=id;result.classList.remove('error');result.textContent='✅ '+found.id+'　'+found.name;
}catch(e){if(seq===lookupSeq){result.textContent='查询失败，请稍后重试';result.classList.add('error')}}}
async function identify(){const id=currentId();if(!id)return msg('请输入正确的义工号码');clearTimeout(lookupTimer);
if(!found||foundKey!==id)await lookupName(id);
if(!found||foundKey!==id)return;
person=found;document.getElementById('login').style.display='none';document.getElementById('welcome').style.display='block';document.getElementById('who').textContent='🙏 '+person.name+'（'+person.id+'）';await refresh();}
function logout(){person=null;document.getElementById('login').style.display='block';document.getElementById('welcome').style.display='none';document.getElementById('items').innerHTML='';msg('请输入义工编号后查看结缘物品。')}
async function refresh(){if(!person)return;try{const d=await api('/donation/api/items');if(!d.ok){msg(d.message);return}
msg(d.message);document.getElementById('items').innerHTML=d.items.map(x=>`<article class="card item">${x.image_url?`<img src="${esc(x.image_url)}" alt="${esc(x.name)}" loading="lazy">`:'<div class="placeholder">🧹</div>'}
<div class="inner"><h2>${esc(x.name)}</h2><p>${esc(x.specification)}</p><p><b>每位结缘：${esc(x.per_volunteer_quantity)}</b></p>
<p class="${x.remaining?'status':'full'}">${x.remaining?'剩余 '+x.remaining+' / '+x.slots+' 位':'🔴 已满额'}</p>
<div class="names"><span class="names-label">已结缘：</span>${x.names.length?'<div class="name-list">'+x.names.map((name,i)=>'<div class="name-entry">'+(i+1)+'. '+esc(name)+'</div>').join('')+'</div>':'<span class="empty">暂无</span>'}</div><br>
<button ${!x.remaining||busy?'disabled':''} onclick="claim(${x.id})">${x.remaining?'＋ 我要结缘':'已满额'}</button></div></article>`).join('')}
catch(e){msg('更新暂时失败，系统会自动重试。')}}
async function claim(id){if(!person||busy)return;if(!confirm('确认以「'+person.name+'」的名义结缘这项物品吗？'))return;
busy=true;try{const d=await api('/donation/api/claim',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({campaign_item_id:id,volunteer_id:person.id})});alert(d.message);await refresh()}catch(e){alert('网络不稳定，请先查看名单确认是否报名成功，避免重复操作。')}finally{busy=false}}
setInterval(refresh,4000);
</script></html>'''

ADMIN_HTML = '''<!doctype html><html lang="zh"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>卫生用品结缘管理</title>''' + STYLE + '''<style>
.photo-card{border:3px solid #e3e9e4;border-radius:16px;overflow:hidden;background:#fff;cursor:pointer;position:relative}
.photo-card.selected{border-color:#15905c;background:#effaf2}.photo-card img,.photo-card .placeholder{height:175px;width:100%;object-fit:contain;background:#f4f5f1}
.photo-card .inner{padding:12px}.photo-card h3{margin:5px 0;font-size:20px}.photo-card .qty{font-size:22px;font-weight:bold}
.select-product{display:block;width:100%;margin:12px 0;background:#657581;font-size:18px;padding:13px 8px}.select-product.chosen{background:#21865a}.photo-card input[type=number]{font-size:22px;max-width:150px}.edit-product{background:#657581;font-size:16px;padding:9px 12px;margin-top:8px}.check{font-size:18px;font-weight:bold;color:#127345}
.delete-product{background:#b74343;font-size:16px;padding:9px 12px;margin-top:8px;margin-left:6px}
 .sticky-save{position:sticky;bottom:0;background:#fff;padding:14px;box-shadow:0 -4px 20px #0002;border-radius:12px;z-index:2}
.sticky-save button{width:100%;font-size:23px}.upload-row{display:grid;grid-template-columns:1fr 1fr;gap:10px}
@media(max-width:600px){.upload-row{grid-template-columns:1fr}}
</style><main><h1>🌸 卫生用品结缘管理</h1>
{% if not authorized %}<div class="card"><form method="post" action="{{url_for('donation.admin_login')}}"><label>负责人专用 PIN</label><input name="pin" type="password" required><button>进入管理</button></form></div>
{% else %}<p><a class="btn secondary" href="{{url_for('donation.admin_logout')}}">退出管理</a> <a class="btn" href="{{url_for('donation.home')}}">查看义工页面</a></p>
<div class="notice" id="notice">正在读取用品…</div>
<div class="card"><h2>① 第一次使用：上传用品照片</h2><p class="muted">每件用品上传一次，以后可重复使用。照片会保存在 Supabase 数据库，不怕 Render 重新部署。</p>
<form id="uploadForm" enctype="multipart/form-data"><div class="upload-row"><div><label>用品名称</label><input name="name" required maxlength="150" placeholder="例如：洗洁精"></div><div><label>规格</label><input name="specification" maxlength="300" placeholder="例如：5公升"></div></div>
<label>从电脑／手机选择照片</label><input type="file" name="photo" accept="image/jpeg,image/png,image/webp" required><button type="submit">📷 上传并保存用品</button></form></div>
<div class="card"><h2>② 点击照片，输入需要几份</h2><p class="muted">每1份由1位义工结缘。没有勾选的用品不会出现在本轮义工页面。已有人报名的用品不可直接移除，请先处理报名。</p>
<div id="products" class="grid"></div></div>
<div class="sticky-save"><div id="campaign" class="muted" style="margin-bottom:8px"></div><button id="saveBtn" onclick="saveAndOpen()">🟢 保存并开放结缘</button></div>
<div class="card"><h2>③ 结缘名单（每4秒更新）</h2><div id="claims"></div><button class="danger" onclick="closeCampaign()">结束本轮结缘</button></div>
<script>
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
let data=null,editing=false,saving=false;
async function api(url,body){const opt=body===undefined?{cache:'no-store'}:{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)};const r=await fetch(url,opt);return r.json()}
function getSelection(){return [...document.querySelectorAll('.photo-card')].map(el=>({product_id:Number(el.dataset.id),selected:el.querySelector('.selected-box').checked,slots:Number(el.querySelector('.qty').value)}))}
function toggle(id){const el=document.querySelector('.photo-card[data-id="'+id+'"]');const check=el.querySelector('.selected-box');check.checked=!check.checked;updateCard(el);editing=true}
function updateCard(el){const yes=el.querySelector('.selected-box').checked;el.classList.toggle('selected',yes);el.querySelector('.qty').disabled=!yes;const btn=el.querySelector('.select-product');btn.classList.toggle('chosen',yes);btn.textContent=yes?'✅ 已选择此用品':'＋ 选择此用品';btn.setAttribute('aria-pressed',String(yes))}
function drawProducts(d){const existing=editing?new Map(getSelection().map(x=>[x.product_id,x])):new Map();
document.getElementById('products').innerHTML=d.products.map(p=>{const saved=existing.get(p.id),selected=saved?saved.selected:!!p.selected,slots=saved?saved.slots:(p.slots||1);return `<div class="photo-card ${selected?'selected':''}" data-id="${p.id}">
${p.has_photo?`<img src="/donation/photo/${p.id}?v=${encodeURIComponent(p.photo_version||'')}" alt="${esc(p.name)}">`:'<div class="placeholder">🧹</div>'}
<div class="inner"><h3>${esc(p.name)}</h3><div class="muted">${esc(p.specification)}</div>
<input class="selected-box" type="checkbox" ${selected?'checked':''} hidden><button type="button" class="select-product ${selected?'chosen':''}" aria-pressed="${selected?'true':'false'}" onclick="toggle(${p.id})">${selected?'✅ 已选择此用品':'＋ 选择此用品'}</button>
<label>需要几份（几位义工）</label><input class="qty" type="number" min="1" max="10000" value="${slots}" ${selected?'':'disabled'} oninput="editing=true">
${p.claimed?`<div class="muted">已结缘 ${p.claimed} 位</div>`:''}<button type="button" class="edit-product" onclick="editProduct(${p.id})">✏️ 编辑名称／规格</button><button type="button" class="delete-product" onclick="deleteProduct(${p.id})">🗑️ 删除用品</button></div></div>`}).join('')||'<p>还没有用品，请先上传照片。</p>'}
async function editProduct(id){const p=data?.products.find(x=>x.id===id);if(!p)return;const name=prompt('修改用品名称：',p.name);if(name===null)return;const spec=prompt('修改规格：',p.specification||'');if(spec===null)return;if(!name.trim())return alert('用品名称不能为空');try{const d=await api('/donation/admin/api/edit-product',{product_id:id,name:name.trim(),specification:spec.trim()});alert(d.message);if(d.ok){const wasEditing=editing;editing=false;await refresh();editing=wasEditing}}catch(e){alert('修改失败，请稍后重试')}}
async function deleteProduct(id){const p=data?.products.find(x=>x.id===id);if(!p)return;
if(p.claimed){alert('这件用品已有义工结缘，不能删除。');return}
if(p.selected){alert('请先取消选择此用品，并点击「保存并开放结缘」，再删除。');return}
if(!confirm('确定停用「'+p.name+'」吗？\\n用品将从资料库列表隐藏，历史记录保留。'))return;
try{const d=await api('/donation/admin/api/delete-product',{product_id:id});alert(d.message);if(d.ok){editing=false;await refresh()}}catch(e){alert('删除失败，请稍后重试')}}
async function refresh(){if(saving)return;try{const d=await api('/donation/admin/api/state');if(!d.ok)return;data=d;
document.getElementById('notice').textContent='📷 已保存 '+d.products.length+' 种用品 · 每4秒自动更新名单';
if(!editing)drawProducts(d);
document.getElementById('campaign').textContent=d.campaign?'本轮状态：'+d.campaign.status+'（编号 '+d.campaign.id+'）':'目前没有开放的结缘活动';
document.getElementById('claims').innerHTML=d.items.map(x=>`<div class="card"><h3>${esc(x.name)}（${x.claims.length}/${x.slots}）</h3>${x.claims.map(c=>`<p>${esc(c.volunteer_name)} (${esc(c.volunteer_id)}) <button class="danger" onclick="removeClaim(${c.id})">删除报名</button></p>`).join('')||'<p>暂无报名</p>'}</div>`).join('')||'<p>暂无本轮用品</p>';
}catch(e){document.getElementById('notice').textContent='网络暂时不稳定，稍后自动重试'}}
document.getElementById('uploadForm').onsubmit=async e=>{e.preventDefault();const form=e.target;const fd=new FormData(form);const btn=form.querySelector('button');btn.disabled=true;try{const r=await fetch('/donation/admin/api/upload-product',{method:'POST',body:fd});const d=await r.json();alert(d.message);if(d.ok){form.reset();editing=false;await refresh()}}catch(e){alert('上传失败，请检查网络和照片大小')}finally{btn.disabled=false}};
async function saveAndOpen(){if(saving)return;const items=getSelection();if(!items.some(x=>x.selected))return alert('请至少选择一种用品');if(items.some(x=>x.selected&&(!Number.isInteger(x.slots)||x.slots<1||x.slots>10000)))return alert('请输入正确份数（1至10000）');
if(!confirm('确认保存所选用品及份数，并立即开放给义工结缘？'))return;
saving=true;document.getElementById('saveBtn').disabled=true;
try{const d=await api('/donation/admin/api/save-and-open',{items});alert(d.message);if(d.ok){editing=false;await refresh()}}catch(e){alert('网络错误，请先检查本轮状态，避免重复操作')}finally{saving=false;document.getElementById('saveBtn').disabled=false}}
async function closeCampaign(){if(!confirm('确认结束本轮结缘？结束后义工不能再报名。'))return;const d=await api('/donation/admin/api/close',{});alert(d.message);editing=false;await refresh()}
async function removeClaim(id){if(!confirm('确认取消这位义工的报名？'))return;const d=await api('/donation/admin/api/remove-claim',{claim_id:id});alert(d.message);await refresh()}
refresh();setInterval(refresh,4000);
</script>{% endif %}</main></html>'''


def _json():
    return request.get_json(silent=True) or {}


def _admin_ok():
    return bool(os.getenv('FLASK_SECRET_KEY') and session.get('donation_admin') is True)


def _admin_required():
    if not _admin_ok():
        return jsonify(ok=False, message='请先登录负责人后台'), 403
    return None


def _current_campaign(cur, lock=False):
    cur.execute('''SELECT * FROM donation_campaigns WHERE status IN ('open','scheduled','draft')
                   ORDER BY id DESC LIMIT 1''' + (' FOR UPDATE' if lock else ''))
    return cur.fetchone()


def _live_campaign(cur, lock=False):
    cur.execute('''SELECT * FROM donation_campaigns WHERE status='open' ORDER BY id DESC LIMIT 1''' + (' FOR UPDATE' if lock else ''))
    return cur.fetchone()


def _find_volunteer(cur, raw):
    raw = str(raw or '').strip().upper()
    if not re.fullmatch(r'(?:CHE-\d+|STW-\d+|8\d\d)', raw):
        return None
    cur.execute("SELECT id,name FROM volunteers WHERE UPPER(id)=%s AND status='在册' LIMIT 1", (raw,))
    return cur.fetchone()


@donation_bp.route('/')
def home():
    return render_template_string(VOL_HTML)


@donation_bp.route('/api/identify', methods=['POST'])
def identify():
    with get_conn() as conn:
        with conn.cursor() as cur:
            v = _find_volunteer(cur, _json().get('volunteer_id'))
    if not v:
        return jsonify(ok=False, message='找不到在册义工，请检查完整编号。')
    return jsonify(ok=True, volunteer={'id':v['id'], 'name':v['name']})


@donation_bp.route('/api/items')
def public_items():
    with get_conn() as conn:
        with conn.cursor() as cur:
            campaign = _live_campaign(cur)
            if not campaign or not campaign['open_at'] or campaign['open_at'] > datetime.now(MY_TZ):
                return jsonify(ok=True, message='目前尚未开放卫生用品结缘，感恩。', items=[])
            cur.execute('''SELECT ci.id,ci.slots,ci.per_volunteer_quantity,p.name,p.specification,p.image_url
                FROM donation_campaign_items ci JOIN donation_products p ON p.id=ci.product_id
                WHERE ci.campaign_id=%s AND ci.is_visible=true ORDER BY ci.sort_order,ci.id''', (campaign['id'],))
            rows = [dict(r) for r in cur.fetchall()]
            for row in rows:
                cur.execute('''SELECT volunteer_name FROM donation_claims WHERE campaign_item_id=%s AND status='active' ORDER BY claimed_at,id''',(row['id'],))
                row['names']=[r['volunteer_name'] for r in cur.fetchall()]
                row['remaining']=max(0,row['slots']-len(row['names']))
    return jsonify(ok=True, message='🙏 感恩大家发心护持，点击物品即可结缘。', items=rows)


@donation_bp.route('/api/claim', methods=['POST'])
def claim():
    d=_json()
    try: item_id=int(d.get('campaign_item_id'))
    except (TypeError,ValueError): return jsonify(ok=False,message='请选择正确物品'),400
    if item_id <= 0: return jsonify(ok=False,message='请选择正确物品'),400
    with get_conn() as conn:
        with conn.cursor() as cur:
            # Lock the campaign row, serializing ALL claims and admin mutations in the round.
            campaign=_live_campaign(cur,lock=True)
            if not campaign or not campaign['open_at'] or campaign['open_at'] > datetime.now(MY_TZ):
                return jsonify(ok=False,message='目前尚未开放结缘。'),409
            v=_find_volunteer(cur,d.get('volunteer_id'))
            if not v: return jsonify(ok=False,message='找不到在册义工，请重新输入编号。'),403
            cur.execute('''SELECT id,slots FROM donation_campaign_items WHERE id=%s AND campaign_id=%s AND is_visible=true''',(item_id,campaign['id']))
            item=cur.fetchone()
            if not item:return jsonify(ok=False,message='该物品目前不开放结缘。'),409
            cur.execute("SELECT campaign_item_id FROM donation_claims WHERE campaign_id=%s AND volunteer_id=%s AND status='active'",(campaign['id'],v['id']))
            claimed=[r['campaign_item_id'] for r in cur.fetchall()]
            if item_id in claimed:return jsonify(ok=False,message='您已结缘过这项物品，每轮不能重复结缘。'),409
            if datetime.now(MY_TZ) < campaign['first_midnight_at'] and claimed:
                return jsonify(ok=False,message='今天每位义工限结缘一种物品；过了今晚12点可选择其他物品，感恩。'),409
            cur.execute("SELECT count(*) AS n FROM donation_claims WHERE campaign_item_id=%s AND status='active'",(item_id,))
            if cur.fetchone()['n']>=item['slots']:
                return jsonify(ok=False,message='🙏 感恩您的发心，该物品已满额，请选择其他物品。'),409
            cur.execute('''INSERT INTO donation_claims(campaign_id,campaign_item_id,volunteer_id,volunteer_name)
                VALUES (%s,%s,%s,%s)''',(campaign['id'],item_id,v['id'],v['name']))
    return jsonify(ok=True,message='🎉 结缘成功！感恩您的发心护持。')


@donation_bp.route('/admin')
def admin():
    return render_template_string(ADMIN_HTML, authorized=_admin_ok())


@donation_bp.route('/admin/login',methods=['POST'])
def admin_login():
    if not os.getenv('FLASK_SECRET_KEY'):
        return '请先设置 FLASK_SECRET_KEY，负责人后台暂不开放。', 503
    configured=os.getenv('DONATION_ADMIN_PIN','')
    entered=request.form.get('pin','')
    if not configured or not hmac.compare_digest(entered,configured):
        return '负责人 PIN 未设置或不正确。请设置 DONATION_ADMIN_PIN。<br><a href="/donation/admin">返回</a>',403
    session['donation_admin']=True
    return redirect(url_for('donation.admin'))


@donation_bp.route('/admin/logout')
def admin_logout():
    session.pop('donation_admin',None)
    return redirect(url_for('donation.admin'))


@donation_bp.route('/photo/<int:product_id>')
def product_photo(product_id):
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute('SELECT image_data,mime_type,updated_at FROM donation_product_photos WHERE product_id=%s',(product_id,))
            row=cur.fetchone()
    if not row:
        return '照片不存在',404
    return Response(bytes(row['image_data']),mimetype=row['mime_type'],headers={'Cache-Control':'public, max-age=3600','X-Content-Type-Options':'nosniff'})


@donation_bp.route('/admin/api/upload-product',methods=['POST'])
def admin_upload_product():
    error=_admin_required()
    if error:return error
    name=(request.form.get('name') or '').strip()
    spec=(request.form.get('specification') or '').strip()
    if not name or len(name)>150 or len(spec)>300:
        return jsonify(ok=False,message='用品名称或规格不正确'),400
    file=request.files.get('photo')
    if not file:return jsonify(ok=False,message='请选择用品照片'),400
    raw=file.stream.read(5*1024*1024+1)
    if len(raw)>5*1024*1024:return jsonify(ok=False,message='照片不能超过5MB'),400
    try:
        with Image.open(io.BytesIO(raw)) as img:
            if img.format not in ('JPEG','PNG','WEBP'):
                return jsonify(ok=False,message='只支持 JPG、PNG、WEBP 照片'),400
            if img.width*img.height>25_000_000:
                return jsonify(ok=False,message='照片像素过大，请缩小后上传'),400
            img.thumbnail((1200,1200))
            if img.mode not in ('RGB','L'):
                img=img.convert('RGB')
            elif img.mode=='L':img=img.convert('RGB')
            out=io.BytesIO();img.save(out,format='JPEG',quality=82,optimize=True)
            photo=out.getvalue()
    except (UnidentifiedImageError,OSError,ValueError):
        return jsonify(ok=False,message='照片格式无法读取'),400
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute('INSERT INTO donation_products(name,specification) VALUES (%s,%s) RETURNING id',(name,spec))
            pid=cur.fetchone()['id']
            cur.execute('INSERT INTO donation_product_photos(product_id,image_data,mime_type) VALUES (%s,%s,%s)',(pid,photo,'image/jpeg'))
            cur.execute('UPDATE donation_products SET image_url=%s WHERE id=%s',(f'/donation/photo/{pid}',pid))
    return jsonify(ok=True,message='📷 用品照片和资料已永久保存。',product_id=pid)


@donation_bp.route('/admin/api/edit-product', methods=['POST'])
def admin_edit_product():
    error = _admin_required()
    if error:
        return error
    d = _json()
    try:
        product_id = int(d.get('product_id'))
    except (TypeError, ValueError):
        return jsonify(ok=False, message='用品编号不正确'), 400
    name = str(d.get('name') or '').strip()
    specification = str(d.get('specification') or '').strip()
    if product_id <= 0 or not name or len(name) > 150 or len(specification) > 300:
        return jsonify(ok=False, message='用品名称或规格不正确'), 400
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute('UPDATE donation_products SET name=%s,specification=%s WHERE id=%s AND is_active=true',
                        (name, specification, product_id))
            if cur.rowcount != 1:
                return jsonify(ok=False, message='找不到用品'), 404
    return jsonify(ok=True, message='✅ 用品名称和规格已更新，照片保持不变。')


@donation_bp.route('/admin/api/delete-product', methods=['POST'])
def admin_delete_product():
    error = _admin_required()
    if error:
        return error
    try:
        product_id = int(_json().get('product_id'))
    except (TypeError, ValueError):
        return jsonify(ok=False, message='用品编号不正确'), 400
    if product_id <= 0:
        return jsonify(ok=False, message='用品编号不正确'), 400
    with get_conn() as conn:
        with conn.cursor() as cur:
            # Use the same campaign lock as claims and admin updates.
            cur.execute("SELECT pg_advisory_xact_lock(hashtext('donation_campaign_create'))")
            campaign = _current_campaign(cur, lock=True)
            cur.execute('SELECT id FROM donation_products WHERE id=%s AND is_active=true FOR UPDATE', (product_id,))
            if not cur.fetchone():
                return jsonify(ok=False, message='用品不存在或已删除'), 404
            cur.execute('''SELECT 1 FROM donation_claims dc
                JOIN donation_campaign_items ci ON ci.id=dc.campaign_item_id
                WHERE ci.product_id=%s LIMIT 1''', (product_id,))
            if cur.fetchone():
                return jsonify(ok=False, message='这件用品已有结缘记录，为保护历史资料，不能删除。'), 409
            if campaign:
                cur.execute('''SELECT 1 FROM donation_campaign_items
                    WHERE campaign_id=%s AND product_id=%s AND is_visible=true LIMIT 1''',
                    (campaign['id'], product_id))
                if cur.fetchone():
                    return jsonify(ok=False, message='请先取消选择此用品并保存本轮，再删除。'), 409
            # Soft-delete: preserve photo and historical campaign references.
            cur.execute('UPDATE donation_products SET is_active=false WHERE id=%s', (product_id,))
    return jsonify(ok=True, message='✅ 用品已从资料库隐藏，照片及历史资料保留。')


@donation_bp.route('/admin/api/save-and-open',methods=['POST'])
def admin_save_and_open():
    error=_admin_required()
    if error:return error
    items=_json().get('items')
    if not isinstance(items,list) or len(items)>500:
        return jsonify(ok=False,message='用品资料格式错误'),400
    try:
        requested={}
        for x in items:
            pid=int(x['product_id']);slots=int(x['slots']);selected=x['selected']
            if pid<=0 or slots<1 or slots>10000 or not isinstance(selected,bool) or pid in requested:
                raise ValueError()
            requested[pid]=(selected,slots)
    except (TypeError,ValueError,KeyError):
        return jsonify(ok=False,message='用品或份数不正确'),400
    if not any(x[0] for x in requested.values()):
        return jsonify(ok=False,message='请至少选择一种用品'),400
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT pg_advisory_xact_lock(hashtext('donation_campaign_create'))")
            campaign=_current_campaign(cur,lock=True)
            if campaign and campaign['status']=='scheduled':
                return jsonify(ok=False,message='已有预约活动，暂时不能立即开放'),409
            if campaign is None:
                cur.execute("INSERT INTO donation_campaigns(title,status) VALUES ('观音堂卫生用品结缘','draft') RETURNING *")
                campaign=cur.fetchone()
            cur.execute('SELECT id FROM donation_products WHERE is_active=true')
            valid={r['id'] for r in cur.fetchall()}
            if any(pid not in valid for pid in requested):
                return jsonify(ok=False,message='有用品已不存在，请刷新页面'),409
            cur.execute('''SELECT ci.id,ci.product_id,ci.slots,ci.is_visible,
                (SELECT count(*) FROM donation_claims dc WHERE dc.campaign_item_id=ci.id AND dc.status='active') AS claimed
                FROM donation_campaign_items ci WHERE ci.campaign_id=%s''',(campaign['id'],))
            old={r['product_id']:r for r in cur.fetchall()}
            for pid,record in old.items():
                selected,slots=requested.get(pid,(False,record['slots']))
                if slots<record['claimed'] and selected:
                    return jsonify(ok=False,message='份数不能少于已结缘人数'),409
                if not selected and record['claimed']:
                    return jsonify(ok=False,message='有用品已有义工结缘，不能直接取消选择。请先删除相关报名。'),409
            for pid,(selected,slots) in requested.items():
                if pid in old:
                    cur.execute('''UPDATE donation_campaign_items SET slots=%s,is_visible=%s,per_volunteer_quantity='1份',updated_at=now()
                        WHERE id=%s''',(slots,selected,old[pid]['id']))
                elif selected:
                    cur.execute('''INSERT INTO donation_campaign_items(campaign_id,product_id,slots,per_volunteer_quantity)
                        VALUES (%s,%s,%s,'1份')''',(campaign['id'],pid,slots))
            for pid,record in old.items():
                if pid not in requested:
                    if record['claimed']:
                        return jsonify(ok=False,message='页面资料已变化，请刷新后重试'),409
                    cur.execute('UPDATE donation_campaign_items SET is_visible=false WHERE id=%s',(record['id'],))
            if campaign['status']=='draft':
                now=datetime.now(MY_TZ)
                midnight=datetime.combine(now.date()+timedelta(days=1),dt_time.min,tzinfo=MY_TZ)
                cur.execute("UPDATE donation_campaigns SET status='open',open_at=%s,first_midnight_at=%s,updated_at=now() WHERE id=%s",(now,midnight,campaign['id']))
    return jsonify(ok=True,message='✅ 已保存并开放结缘！义工页面马上可以看到所选用品。')


@donation_bp.route('/admin/api/state')
def admin_state():
    error=_admin_required()
    if error:return error
    with get_conn() as conn:
        with conn.cursor() as cur:
            campaign=_current_campaign(cur)
            cur.execute('''SELECT p.id,p.name,p.specification,ph.product_id IS NOT NULL AS has_photo,
                ph.updated_at AS photo_version FROM donation_products p LEFT JOIN donation_product_photos ph ON ph.product_id=p.id
                WHERE p.is_active=true ORDER BY p.sort_order,p.id''')
            products=[dict(r) for r in cur.fetchall()]
            items=[]
            if campaign:
                cur.execute('''SELECT ci.id,ci.slots,p.name FROM donation_campaign_items ci
                    JOIN donation_products p ON p.id=ci.product_id WHERE ci.campaign_id=%s ORDER BY ci.sort_order,ci.id''',(campaign['id'],))
                items=[dict(r) for r in cur.fetchall()]
                for x in items:
                    cur.execute("SELECT id,volunteer_id,volunteer_name FROM donation_claims WHERE campaign_item_id=%s AND status='active' ORDER BY claimed_at,id",(x['id'],))
                    x['claims']=[dict(r) for r in cur.fetchall()]
    # 依据 product_id 而非名称匹配，避免同名用品互相覆盖。
    if campaign:
        with get_conn() as conn:
            with conn.cursor() as cur:
                cur.execute('''SELECT product_id,slots,is_visible,id FROM donation_campaign_items WHERE campaign_id=%s''',(campaign['id'],))
                item_by_product={r['product_id']:r for r in cur.fetchall()}
        for p in products:
            r=item_by_product.get(p['id'])
            p['selected']=bool(r and r['is_visible'])
            p['slots']=r['slots'] if r else 1
            p['claimed']=len(next((x['claims'] for x in items if x['id']==r['id']),[])) if r else 0
    return jsonify(ok=True,message='每4秒自动更新，无需F5。',campaign=dict(campaign) if campaign else None,products=products,items=items)


@donation_bp.route('/admin/api/product',methods=['POST'])
def admin_product():
    error=_admin_required()
    if error:return error
    d=_json();name=str(d.get('name') or '').strip()
    if not name or len(name)>150:return jsonify(ok=False,message='请输入物品名称（最多150字）'),400
    image=str(d.get('image_url') or '').strip()
    if image and (not image.startswith('https://') or len(image)>2000):
        return jsonify(ok=False,message='照片必须使用 HTTPS 图片网址'),400
    try:price=float(d['reference_price']) if str(d.get('reference_price') or '').strip() else None
    except ValueError:return jsonify(ok=False,message='参考价格格式错误'),400
    if price is not None and (price<0 or price>9999999999):return jsonify(ok=False,message='参考价格不正确'),400
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute('''INSERT INTO donation_products(name,specification,description,image_url,reference_price)
                VALUES (%s,%s,%s,%s,%s)''',(name,str(d.get('specification') or '')[:300],str(d.get('description') or '')[:2000],image or None,price))
    return jsonify(ok=True,message='物品已加入资料库。')


@donation_bp.route('/admin/api/campaign',methods=['POST'])
def admin_campaign():
    error=_admin_required()
    if error:return error
    with get_conn() as conn:
        with conn.cursor() as cur:
            # Global transaction advisory lock also protects against simultaneous admins creating campaigns.
            cur.execute("SELECT pg_advisory_xact_lock(hashtext('donation_campaign_create'))")
            existing=_current_campaign(cur)
            if existing:return jsonify(ok=False,message='已有未结束的结缘活动，请先使用或结束该轮。'),409
            cur.execute("INSERT INTO donation_campaigns(title,status) VALUES ('观音堂卫生用品结缘','draft')")
    return jsonify(ok=True,message='已建立新一轮草稿。')


@donation_bp.route('/admin/api/item',methods=['POST'])
def admin_item():
    error=_admin_required()
    if error:return error
    d=_json()
    try:product_id=int(d.get('product_id'));slots=int(d.get('slots'))
    except (TypeError,ValueError):return jsonify(ok=False,message='物品或名额不正确'),400
    if product_id<=0 or slots<=0 or slots>10000:return jsonify(ok=False,message='名额必须为1至10000'),400
    quantity=str(d.get('per_volunteer_quantity') or '1份').strip()[:100]
    with get_conn() as conn:
        with conn.cursor() as cur:
            campaign=_current_campaign(cur,lock=True)
            if not campaign or campaign['status']=='scheduled':return jsonify(ok=False,message='请先建立草稿。'),409
            cur.execute('SELECT id FROM donation_products WHERE id=%s AND is_active=true',(product_id,))
            if not cur.fetchone():return jsonify(ok=False,message='物品不存在'),404
            cur.execute('''SELECT ci.id,(SELECT count(*) FROM donation_claims dc WHERE dc.campaign_item_id=ci.id AND dc.status='active') AS n
                FROM donation_campaign_items ci WHERE ci.campaign_id=%s AND ci.product_id=%s''',(campaign['id'],product_id))
            old=cur.fetchone()
            if old:
                if slots<old['n']:return jsonify(ok=False,message='新名额不能少于已经报名的人数。'),409
                cur.execute('''UPDATE donation_campaign_items SET slots=%s,per_volunteer_quantity=%s,updated_at=now()
                    WHERE id=%s''',(slots,quantity,old['id']))
            else:
                cur.execute('''INSERT INTO donation_campaign_items(campaign_id,product_id,slots,per_volunteer_quantity)
                    VALUES (%s,%s,%s,%s)''',(campaign['id'],product_id,slots,quantity))
    return jsonify(ok=True,message='本轮物品及名额已保存。')


@donation_bp.route('/admin/api/open',methods=['POST'])
def admin_open():
    error=_admin_required()
    if error:return error
    with get_conn() as conn:
        with conn.cursor() as cur:
            campaign=_current_campaign(cur,lock=True)
            if not campaign or campaign['status']!='draft':return jsonify(ok=False,message='请先建立尚未开放的草稿。'),409
            cur.execute('SELECT count(*) AS n FROM donation_campaign_items WHERE campaign_id=%s AND is_visible=true',(campaign['id'],))
            if not cur.fetchone()['n']:return jsonify(ok=False,message='请先加入至少一种物品。'),409
            now=datetime.now(MY_TZ)
            midnight=datetime.combine(now.date()+timedelta(days=1),dt_time.min,tzinfo=MY_TZ)
            cur.execute("UPDATE donation_campaigns SET status='open',open_at=%s,first_midnight_at=%s,updated_at=now() WHERE id=%s",(now,midnight,campaign['id']))
    return jsonify(ok=True,message='已立即开放结缘。')


@donation_bp.route('/admin/api/close',methods=['POST'])
def admin_close():
    error=_admin_required()
    if error:return error
    with get_conn() as conn:
        with conn.cursor() as cur:
            campaign=_current_campaign(cur,lock=True)
            if not campaign:return jsonify(ok=False,message='目前没有活动。'),409
            cur.execute("UPDATE donation_campaigns SET status='closed',closed_at=now(),updated_at=now() WHERE id=%s",(campaign['id'],))
    return jsonify(ok=True,message='本轮已结束，历史记录仍保留。')


@donation_bp.route('/admin/api/remove-claim',methods=['POST'])
def admin_remove_claim():
    error=_admin_required()
    if error:return error
    try:claim_id=int(_json().get('claim_id'))
    except (TypeError,ValueError):return jsonify(ok=False,message='报名编号错误'),400
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute('SELECT campaign_id FROM donation_claims WHERE id=%s',(claim_id,))
            record=cur.fetchone()
            if not record:return jsonify(ok=False,message='找不到报名'),404
            cur.execute('SELECT id FROM donation_campaigns WHERE id=%s FOR UPDATE',(record['campaign_id'],))
            cur.execute("UPDATE donation_claims SET status='cancelled',cancelled_at=now(),cancelled_by='donation_admin' WHERE id=%s AND status='active'",(claim_id,))
            if cur.rowcount!=1:return jsonify(ok=False,message='报名已被取消'),409
    return jsonify(ok=True,message='已取消该义工的结缘报名。')
