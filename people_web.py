# people_web.py
# 人员管理中心 v6.7.4
# 直接管理 Supabase/PostgreSQL 的 members + volunteers
# 功能：总名单、搜索筛选、个人资料编辑、义工/月费状态切换、Excel 下载

from io import BytesIO
import re
import json
import secrets
from datetime import datetime
import pandas as pd
from psycopg2.extras import RealDictCursor
from db import db_query, get_conn
from werkzeug.security import generate_password_hash, check_password_hash

from flask import (
    Blueprint,
    request,
    redirect,
    url_for,
    flash,
    abort,
    Response,
    render_template_string,
    send_file,
    session,
)

people_bp = Blueprint("people", __name__, url_prefix="/people")


def clean(v):
    return str(v or "").strip()



# 人员姓名搜索用：将常见繁体字统一成简体后再比较。
# 只影响搜索，不修改数据库中本人正式姓名。
_TRAD = "蕭倫麗儀蘇葉賜陳劉黃吳張楊趙鄭謝鍾馮許羅鄧盧鍾龔龍鳳華國學會義體關愛歡樂禮佛經號聯絡電話門東萬與為來時後裡麼這個們從還過見長開發現實應當讓點無處間問題資料檔案審核轉會費續狀態舊新寫報名號碼備註職業區歲歷師專屬請鏈結個資緊急親係單雙總數個人檔案"
_SIMP = "萧伦丽仪苏叶赐陈刘黄吴张杨赵郑谢钟冯许罗邓卢钟龚龙凤华国学会义体关爱欢乐礼佛经号联络电话门东万与为来时后里么这个们从还过见长开发现实应当让点无处间问题资料档案审核转会费续状态旧新写报名号码备注职业区岁历师专属请链结个资紧急亲系单双总数个人档案"
_ZH_SEARCH_TRANS = str.maketrans({t: s for t, s in zip(_TRAD, _SIMP)})

def normalize_zh_search(v):
    return clean(v).lower().translate(_ZH_SEARCH_TRANS)

def digits(v):
    return "".join(ch for ch in clean(v) if ch.isdigit())


def normalize_phone(v):
    """电话号码入库统一去掉空格、横线等；例如 016-2038622 -> 0162038622。"""
    raw = clean(v)
    if not raw:
        return ""
    d = digits(raw)
    # 马来西亚本地手机号码统一存纯数字；其它特殊格式保留原值，避免误改。
    return d if d.startswith("0") and 9 <= len(d) <= 11 else raw


def format_phone(v):
    """网页显示统一为 016-2038622；非本地/异常格式不强行改变。"""
    raw = clean(v)
    d = digits(raw)
    if d.startswith("0") and 9 <= len(d) <= 11:
        return d[:3] + "-" + d[3:]
    return raw


def default_pin(phone):
    p = digits(phone)
    return p[-4:] if len(p) >= 4 else ""


def norm_branch(v):
    return "STW" if clean(v).upper() == "STW" else "CHE"


def natural_id_key(value):
    s = clean(value).upper()
    m = re.search(r"(\d+)$", s)
    n = int(m.group(1)) if m else 999999
    if s.startswith("CHE-"):
        group = 0
    elif s.startswith("STW-"):
        group = 1
    elif s.isdigit():
        group = 2
    else:
        group = 3
    return (group, n, s)



def next_volunteer_id():
    rows = db_query("select id from volunteers", fetchall=True) or []
    nums = []
    for r in rows:
        value = clean(r.get("id"))
        if value.isdigit():
            n = int(value)
            if n >= 800:
                nums.append(n)
    return str(max(nums, default=799) + 1)


def next_member_id(branch):
    branch = norm_branch(branch)
    rows = db_query("""
        select member_id
        from members
        where upper(member_id) like %s
    """, (branch + "-%",), fetchall=True) or []

    nums = []
    for r in rows:
        mid = clean(r.get("member_id")).upper()
        m = re.fullmatch(r"(?:CHE|STW)-(\d+)", mid)
        if m:
            nums.append(int(m.group(1)))
    return f"{branch}-{max(nums, default=0) + 1}"


def normalize_member_id(branch, value):
    branch = norm_branch(branch)
    value = clean(value).upper().replace(" ", "")
    if not value:
        return ""
    if re.fullmatch(r"(CHE|STW)-\d+", value):
        return value
    if value.isdigit():
        return f"{branch}-{int(value)}"
    return value


def member_exists(member_id):
    row = db_query("select 1 from members where member_id=%s", (member_id,), fetchone=True)
    return bool(row)


def volunteer_exists(volunteer_id):
    row = db_query("select 1 from volunteers where id=%s", (volunteer_id,), fetchone=True)
    return bool(row)

def load_people():
    members = db_query("""
        select member_id, name, english_name, phone, pin, branch,
               status, remark, member_status, ic_number
        from members
    """, fetchall=True) or []

    volunteers = db_query("""
        select id, branch, name, status, phone,
               is_volunteer, is_member, member_id, pin, remark
        from volunteers
    """, fetchall=True) or []

    # 一个人优先以 member_id 作为合并键。
    # 纯义工没有 member_id，则以 V:<义工编号> 单独保留。
    people = {}

    for m in members:
        mid = clean(m.get("member_id"))
        key = "M:" + mid.upper()
        people[key] = {
            "key": key,
            "display_id": mid,
            "member_id": mid,
            "volunteer_id": "",
            "branch": norm_branch(m.get("branch")),
            "member_branch": norm_branch(m.get("branch")),
            "volunteer_branch": "",
            "name": clean(m.get("name")),
            "english_name": clean(m.get("english_name")),
            "phone": format_phone(m.get("phone")),
            "pin": clean(m.get("pin")),
            "ic_number": clean(m.get("ic_number")),
            "remark": clean(m.get("remark")),
            "is_member": True,
            "member_status": clean(m.get("member_status") or m.get("status") or "在供"),
            "is_volunteer": False,
            "volunteer_status": "",
        }

    for v in volunteers:
        vid = clean(v.get("id"))
        linked_mid = clean(v.get("member_id"))

        # 旧资料有时 is_member=true 但 member_id 为空；若义工编号本身就是 CHE/STW 编号，
        # 尝试与 members 同编号连接。
        candidate_mid = linked_mid
        if not candidate_mid and bool(v.get("is_member")) and (
            vid.upper().startswith("CHE-") or vid.upper().startswith("STW-")
        ):
            candidate_mid = vid

        mkey = "M:" + candidate_mid.upper() if candidate_mid else ""
        if mkey and mkey in people:
            p = people[mkey]
            p["volunteer_id"] = vid
            p["volunteer_branch"] = norm_branch(v.get("branch"))
            p["is_volunteer"] = bool(v.get("is_volunteer")) if v.get("is_volunteer") is not None else True
            p["volunteer_status"] = clean(v.get("status") or "在册")
            # members 是个人资料主来源；空白时才从 volunteers 补。
            p["name"] = p["name"] or clean(v.get("name"))
            p["phone"] = p["phone"] or format_phone(v.get("phone"))
            p["pin"] = p["pin"] or clean(v.get("pin"))
            p["remark"] = p["remark"] or clean(v.get("remark"))
        else:
            key = "V:" + vid.upper()
            people[key] = {
                "key": key,
                "display_id": vid,
                "member_id": linked_mid,
                "volunteer_id": vid,
                "branch": norm_branch(v.get("branch")),
                "member_branch": "",
                "volunteer_branch": norm_branch(v.get("branch")),
                "name": clean(v.get("name")),
                "english_name": "",
                "phone": format_phone(v.get("phone")),
                "pin": clean(v.get("pin")),
                "ic_number": "",
                "remark": clean(v.get("remark")),
                "is_member": False,
                "member_status": "",
                "is_volunteer": bool(v.get("is_volunteer")) if v.get("is_volunteer") is not None else True,
                "volunteer_status": clean(v.get("status") or "在册"),
            }

    rows = list(people.values())
    rows.sort(key=lambda x: natural_id_key(x["display_id"]))
    return rows


def find_person(person_key):
    for p in load_people():
        if p["key"] == person_key:
            return p
    return None


def filtered_people():
    rows = load_people()
    q = clean(request.args.get("q")).lower()
    branch = clean(request.args.get("branch")).upper()
    kind = clean(request.args.get("kind") or "all")

    if q:
        qdigits = digits(q)
        qnorm = normalize_zh_search(q)
        def match(p):
            hay = " ".join([
                p["display_id"], p["member_id"], p["volunteer_id"],
                p["name"], p["english_name"], p["phone"]
            ]).lower()
            haynorm = normalize_zh_search(hay)
            return q in hay or qnorm in haynorm or (qdigits and qdigits in digits(p["phone"]))
        rows = [p for p in rows if match(p)]

    if branch in ("CHE", "STW"):
        if kind in ("volunteer", "volunteer_only", "volunteer_inactive"):
            rows = [p for p in rows if p["is_volunteer"] and p["volunteer_branch"] == branch]
        elif kind in ("member", "member_only", "member_inactive"):
            rows = [p for p in rows if p["is_member"] and p["member_branch"] == branch]
        elif kind == "volunteer_member":
            # 「义工＋月费」以义工所属分会筛选；月费归属另外显示。
            rows = [p for p in rows if p["is_volunteer"] and p["volunteer_branch"] == branch and p["is_member"]]
        else:
            # 全部人员：任一身份属于该分会即可出现。
            rows = [p for p in rows if p["volunteer_branch"] == branch or p["member_branch"] == branch]

    if kind == "volunteer":
        rows = [p for p in rows if p["is_volunteer"]]
    elif kind == "volunteer_member":
        rows = [p for p in rows if p["is_volunteer"] and p["is_member"]]
    elif kind == "volunteer_only":
        rows = [p for p in rows if p["is_volunteer"] and not p["is_member"]]
    elif kind == "member":
        rows = [p for p in rows if p["is_member"]]
    elif kind == "member_only":
        rows = [p for p in rows if p["is_member"] and not p["is_volunteer"]]
    elif kind == "volunteer_inactive":
        rows = [p for p in rows if p["is_volunteer"] and p["volunteer_status"] != "在册"]
    elif kind == "member_inactive":
        rows = [p for p in rows if p["is_member"] and p["member_status"] != "在供"]

    return rows


PAGE = r"""
<!doctype html>
<html lang="zh">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>人员管理中心</title>
<link rel="stylesheet" href="{{ url_for('static', filename='css/toolbox.css') }}">
<style>
body{background:#f5f7fb}.people-page{max-width:1280px;margin:0 auto;padding:22px}
.hero{background:linear-gradient(135deg,#6157e7,#8a6cf0);color:#fff;border-radius:22px;padding:25px 28px;margin-bottom:18px;display:flex;justify-content:space-between;align-items:center;gap:18px}.hero-main{min-width:0}.hero-account{display:flex;gap:9px;flex-wrap:wrap;justify-content:flex-end}.hero-account .btn{background:rgba(255,255,255,.94);color:#333}.hero-account .password-btn{background:#d9363e;color:#fff}
.hero h1{margin:0 0 7px;font-size:32px}.hero p{margin:0;opacity:.94}
.stats{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin:18px 0}
.stat{background:#fff;border-radius:16px;padding:16px;text-align:center;box-shadow:0 2px 10px #0000000d}
.stat b{display:block;font-size:27px;margin-top:5px}.stat small{display:block;color:#667085;margin-top:4px;font-weight:700}
.toolbar,.person-card{background:#fff;border-radius:18px;padding:16px;margin-bottom:13px;box-shadow:0 2px 10px #0000000d}
.filters{display:grid;grid-template-columns:2fr 1fr 1.4fr auto;gap:10px}
.input{width:100%;padding:12px;border:1px solid #d7dce5;border-radius:11px;box-sizing:border-box;font-size:16px}
.actions{display:flex;gap:9px;flex-wrap:wrap;margin-top:12px}
.btn{display:inline-block;padding:10px 15px;border:0;border-radius:10px;text-decoration:none;cursor:pointer;font-weight:700}
.primary{background:#6558e8;color:white}.secondary{background:#eef0f5;color:#333}.success{background:#e8f7ee;color:#14723c}.danger{background:#fdecec;color:#a32626}
.person-head{display:flex;justify-content:space-between;gap:16px;align-items:flex-start}
.person-id{font-weight:900;font-size:18px}.person-name{font-size:22px;font-weight:900;margin:4px 0}
.meta{color:#667085;line-height:1.7}.badges{display:flex;gap:7px;flex-wrap:wrap;margin-top:8px}
.badge{padding:5px 9px;border-radius:999px;font-size:13px;font-weight:800;background:#f0f2f6}
.on{background:#e7f7ed;color:#18753e}.off{background:#fbeaea;color:#9b2c2c}
.alert{padding:12px 14px;border-radius:10px;margin-bottom:12px}.alert-success{background:#e8f7ee}.alert-danger{background:#fdecec}
@media(max-width:760px){.hero{display:block}.hero-account{justify-content:flex-start;margin-top:16px}.stats{grid-template-columns:1fr 1fr}.filters{grid-template-columns:1fr}.person-head{display:block}}
</style>
</head>
<body><div class="people-page">
<div class="hero">
 <div class="hero-main"><h1>👥 人员管理中心</h1><p>统一管理义工与月费佛友资料</p></div>
 <div class="hero-account">
  <a class="btn password-btn" href="{{ url_for('people.admin_change_password') }}">🔑 修改密码</a>
  <a class="btn" href="{{ url_for('people.admin_logout') }}">🔒 登出</a>
 </div>
</div>

{% with messages = get_flashed_messages(with_categories=true) %}
{% for cat,msg in messages %}<div class="alert {{ 'alert-success' if cat=='ok' else 'alert-danger' }}">{{ msg }}</div>{% endfor %}
{% endwith %}

<div class="stats">
 <div class="stat">总人数<b>{{ stats.total }}</b></div>
 <div class="stat">义工<b>{{ stats.volunteers }}</b><small>已更新 {{ stats.profile_done }} · 待审核 {{ stats.profile_pending }} · 尚未更新 {{ stats.profile_missing }}</small></div>
 <div class="stat">月费佛友<b>{{ stats.members }}</b><small>CHE {{ stats.che_members }} · STW {{ stats.stw_members }}</small></div>
 <div class="stat">义工＋月费<b>{{ stats.both }}</b></div>
</div>

<div class="toolbar">
<form method="get">
<div class="filters">
<input class="input" name="q" value="{{ request.args.get('q','') }}" placeholder="编号 / 姓名 / 英文名 / 电话">
<select class="input" name="branch">
<option value="">全部分会</option><option value="CHE" {% if request.args.get('branch')=='CHE' %}selected{% endif %}>CHE</option>
<option value="STW" {% if request.args.get('branch')=='STW' %}selected{% endif %}>STW</option>
</select>
<select class="input" name="kind">
{% for value,label in kinds %}
<option value="{{ value }}" {% if request.args.get('kind','all')==value %}selected{% endif %}>{{ label }}</option>
{% endfor %}
</select>
<button class="btn primary">🔍 查询</button>
</div>
</form>
<div class="actions">
<a class="btn primary" href="{{ url_for('people.add') }}">💰 新增月费佛友</a>
<a class="btn success" href="{{ url_for('people.profile_reviews') }}">📝 待审核资料更新</a>
<a class="btn success" href="{{ url_for('people.applications') }}">🆕 新义工申请</a>
<a class="btn secondary" href="{{ url_for('people.self_update_login') }}">🔗 义工资料更新入口</a>
<a class="btn secondary" href="{{ url_for('people.index') }}">清除筛选</a>
<a class="btn success" href="{{ url_for('people.download', q=request.args.get('q',''), branch=request.args.get('branch',''), kind=request.args.get('kind','all')) }}">📥 下载当前名单</a>
</div>
</div>

{% for p in rows %}
<div class="person-card">
 <div class="person-head">
  <div>
   <div class="person-id">{{ p.display_id }}</div>
   <div class="person-name">{{ p.name or '未填写姓名' }}</div>
   <div class="meta">
    {% if p.english_name %}{{ p.english_name }} · {% endif %}📞 {{ p.phone or '-' }}
   </div>
   <div class="badges">
    {% if p.is_volunteer %}<span class="badge {{ 'on' if p.volunteer_status=='在册' else 'off' }}">👥 义工：{{ p.volunteer_branch }} · {{ p.volunteer_status }}</span>{% else %}<span class="badge">非义工</span>{% endif %}
    {% if p.is_member %}<span class="badge {{ 'on' if p.member_status=='在供' else 'off' }}">💰 月费：{{ p.member_branch }} · {{ p.member_status }}</span>{% else %}<span class="badge">非月费</span>{% endif %}
   </div>
  </div>
  <div class="actions" style="margin-top:0"><a class="btn success" href="{{ url_for('people.profile_view', volunteer_id=p.volunteer_id) }}" {% if not p.is_volunteer %}style="display:none"{% endif %}>📁 个人档案</a><a class="btn primary" href="{{ url_for('people.edit', person_key=p.key) }}">✏️ 编辑资料</a></div>
 </div>
</div>
{% else %}
<div class="person-card">没有符合条件的人员。</div>
{% endfor %}
</div></body></html>
"""

EDIT_PAGE = r"""
<!doctype html><html lang="zh"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>编辑人员资料</title>
<link rel="stylesheet" href="{{ url_for('static', filename='css/toolbox.css') }}">
<style>
body{background:#f5f7fb}.wrap{max-width:850px;margin:0 auto;padding:22px}.card{background:white;border-radius:20px;padding:22px;margin-bottom:15px}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:13px}.full{grid-column:1/-1}
label{font-weight:800;display:block;margin-bottom:6px}.input{width:100%;padding:12px;border:1px solid #d7dce5;border-radius:10px;box-sizing:border-box;font-size:16px}
.btn{padding:11px 16px;border:0;border-radius:10px;text-decoration:none;font-weight:800;cursor:pointer}.primary{background:#6558e8;color:white}.secondary{background:#e9ebf0;color:#333}.danger{background:#fde8e8;color:#a32323}.success{background:#e6f7ed;color:#14733c}
.actions{display:flex;gap:10px;flex-wrap:wrap;margin-top:16px}.note{background:#fff8df;padding:12px;border-radius:10px;margin-bottom:14px}
@media(max-width:700px){.grid{grid-template-columns:1fr}.full{grid-column:auto}}
</style></head><body><div class="wrap">
<div class="card"><h1>✏️ 编辑个人资料</h1>
<div class="note">编号不会在这里修改。CHE/STW 转分会或更换编号会另外做专用功能，避免影响历史财政与签到记录。</div>
<form method="post">
<div class="grid">
<div><label>显示编号</label><input class="input" value="{{ p.display_id }}" disabled></div>
{% if p.is_volunteer %}<div><label>义工分会</label><input class="input" value="{{ p.volunteer_branch }} · {{ p.volunteer_status }}" disabled><small>义工分会不能在个人资料页直接修改。</small></div>{% endif %}
{% if p.is_member %}<div><label>月费资料</label><input class="input" value="{{ p.member_id }} · {{ p.member_branch }} · {{ p.member_status }}" disabled><small>月费转会请使用下方专用按钮。</small></div>{% endif %}
<div><label>姓名</label><input class="input" name="name" value="{{ p.name }}" required></div>
<div><label>英文名</label><input class="input" name="english_name" value="{{ p.english_name }}"></div>
<div><label>电话号码</label><input class="input" name="phone" value="{{ p.phone }}"></div>
<div><label>PIN</label><input class="input" name="pin" value="{{ p.pin }}" inputmode="numeric"></div>
{% if p.is_member %}<div><label>身份证号码</label><input class="input" name="ic_number" value="{{ p.ic_number }}"></div>{% endif %}
<div class="full"><label>备注</label><textarea class="input" name="remark" rows="3">{{ p.remark }}</textarea></div>
</div>
<label style="margin-top:13px"><input type="checkbox" name="pin_from_phone" value="1"> 电话更改后，PIN 自动使用新电话号码最后4位</label>
<div class="actions"><button class="btn primary" type="submit">💾 保存资料</button><a class="btn secondary" href="{{ url_for('people.index') }}">返回</a></div>
</form></div>

<div class="card"><h2>身份状态</h2>
<div class="actions">
{% if p.is_volunteer %}
<form method="post" action="{{ url_for('people.toggle_volunteer', person_key=p.key) }}">
<button class="btn {{ 'danger' if p.volunteer_status=='在册' else 'success' }}" type="submit">{{ '停止义工' if p.volunteer_status=='在册' else '恢复义工' }}</button>
</form>
{% endif %}
{% if p.is_member %}
<form method="post" action="{{ url_for('people.toggle_member', person_key=p.key) }}">
<button class="btn {{ 'danger' if p.member_status=='在供' else 'success' }}" type="submit">{{ '⏸️ 停止月费' if p.member_status=='在供' else '▶️ 恢复月费' }}</button>
</form>
<a class="btn secondary" href="{{ url_for('people.transfer_member', person_key=p.key) }}">🔄 月费转会</a>
{% endif %}
{% if p.is_member and not p.is_volunteer %}
<a class="btn success" href="{{ url_for('people.start_member_volunteer_application', person_key=p.key) }}">🆕 申请成为义工</a>
{% endif %}
{% if not p.is_member %}
<a class="btn success" href="{{ url_for('people.add_member_role', person_key=p.key) }}">➕ 加入月费</a>
{% endif %}
</div></div>
</div></body></html>
"""


ADD_PAGE = r"""
<!doctype html><html lang="zh"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>新增月费佛友</title><style>
body{background:#f5f7fb;font-family:Arial,"Microsoft YaHei",sans-serif}.wrap{max-width:760px;margin:0 auto;padding:22px}.card{background:#fff;border-radius:20px;padding:24px;box-shadow:0 2px 10px #0000000d}.grid{display:grid;grid-template-columns:1fr 1fr;gap:14px}.full{grid-column:1/-1}label{display:block;font-weight:800;margin-bottom:6px}.input{width:100%;padding:12px;border:1px solid #d7dce5;border-radius:10px;box-sizing:border-box;font-size:16px}.btn{padding:11px 17px;border:0;border-radius:10px;text-decoration:none;font-weight:800;cursor:pointer}.primary{background:#6558e8;color:#fff}.secondary{background:#e9ebf0;color:#333}.note{background:#eef7ff;padding:12px;border-radius:10px;margin:14px 0;line-height:1.6}.alert{background:#fdecec;color:#8f2424;padding:12px;border-radius:10px;margin-bottom:12px}@media(max-width:700px){.grid{grid-template-columns:1fr}.full{grid-column:auto}}
</style></head><body><div class="wrap"><div class="card"><h1>💰 新增月费佛友</h1>
<div class="note">这里现在只建立月费佛友。新义工必须使用「新义工申请」流程，经面试、试用及负责人批准后才建立正式义工身份。</div>
{% if error %}<div class="alert">{{ error }}</div>{% endif %}
<form method="post"><div class="grid">
<div><label>月费分会</label><select class="input" name="branch" id="branch"><option value="CHE" {% if form.get('branch','CHE')=='CHE' %}selected{% endif %}>CHE</option><option value="STW" {% if form.get('branch')=='STW' %}selected{% endif %}>STW</option></select></div>
<div><label>月费编号</label><input class="input" name="member_id" id="member_id" value="{{ form.get('member_id','') }}" required></div>
<div><label>姓名</label><input class="input" name="name" value="{{ form.get('name','') }}" required></div>
<div><label>英文名</label><input class="input" name="english_name" value="{{ form.get('english_name','') }}"></div>
<div><label>电话号码</label><input class="input" name="phone" value="{{ form.get('phone','') }}" placeholder="例如 016-2038622"></div>
<div><label>PIN</label><input class="input" name="pin" value="{{ form.get('pin','') }}" placeholder="留空 = 电话最后4位"></div>
<div><label>身份证号码</label><input class="input" name="ic_number" value="{{ form.get('ic_number','') }}"></div>
<div class="full"><label>备注</label><textarea class="input" name="remark" rows="3">{{ form.get('remark','') }}</textarea></div>
</div><div style="margin-top:18px;display:flex;gap:10px"><button class="btn primary">💾 建立月费佛友</button><a class="btn secondary" href="{{ url_for('people.index') }}">取消</a></div></form></div></div>
<script>const suggestions={CHE:{{ suggested_che|tojson }},STW:{{ suggested_stw|tojson }}};const b=document.getElementById('branch'),m=document.getElementById('member_id');function setSuggested(){m.value=suggestions[b.value]}b.addEventListener('change',setSuggested);if(!m.value)setSuggested();</script>
</body></html>
"""

ADD_MEMBER_ROLE_PAGE = r"""
<!doctype html><html lang="zh"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>加入月费</title><style>
body{background:#f5f7fb;font-family:Arial,"Microsoft YaHei",sans-serif}.wrap{max-width:650px;margin:0 auto;padding:24px}.card{background:#fff;border-radius:18px;padding:24px}
.input{width:100%;padding:12px;border:1px solid #d7dce5;border-radius:10px;box-sizing:border-box;font-size:16px;margin:7px 0 15px}
.btn{padding:11px 16px;border:0;border-radius:10px;text-decoration:none;font-weight:800}.primary{background:#6558e8;color:#fff}.secondary{background:#e9ebf0;color:#333}.alert{background:#fdecec;padding:12px;border-radius:10px}
</style></head><body><div class="wrap"><div class="card">
<h1>💰 加入月费</h1><p><b>{{ p.name }}</b>（现有义工编号：{{ p.volunteer_id }}）</p>
<p>原来的义工编号不会改，过去签到/排班记录继续保留；新月费编号会通过 member_id 与这位义工连接。</p>
{% if error %}<div class="alert">{{ error }}</div>{% endif %}
<form method="post">
<label>分会</label><select class="input" name="branch"><option value="CHE" {% if p.branch=='CHE' %}selected{% endif %}>CHE</option><option value="STW" {% if p.branch=='STW' %}selected{% endif %}>STW</option></select>
<label>月费编号</label><input class="input" name="member_id" value="{{ suggested }}" required>
<label>英文名</label><input class="input" name="english_name">
<label>身份证号码</label><input class="input" name="ic_number">
<button class="btn primary">确认加入月费</button> <a class="btn secondary" href="{{ url_for('people.edit', person_key=p.key) }}">取消</a>
</form></div></div></body></html>
"""

KINDS = [
    ("all", "全部人员"),
    ("volunteer", "所有义工"),
    ("volunteer_member", "义工＋月费"),
    ("volunteer_only", "纯义工（没有月费）"),
    ("member", "所有月费佛友"),
    ("member_only", "月费但没做义工"),
    ("volunteer_inactive", "已停止义工"),
    ("member_inactive", "已停供月费"),
]


@people_bp.route("/")
def index():
    ensure_people_profile_tables()
    all_rows = load_people()
    rows = filtered_people()
    profile_rows = db_query("select volunteer_id from people_profiles", fetchall=True) or []
    pending_rows = db_query("select distinct volunteer_id from people_profile_updates where status='pending'", fetchall=True) or []
    profile_ids = {clean(r.get('volunteer_id')).upper() for r in profile_rows}
    pending_ids = {clean(r.get('volunteer_id')).upper() for r in pending_rows}
    active_volunteer_ids = {clean(p.get('volunteer_id')).upper() for p in all_rows if p['is_volunteer'] and p['volunteer_status']=='在册'}
    stats = {
        "total": len(all_rows),
        "volunteers": sum(1 for p in all_rows if p["is_volunteer"] and p["volunteer_status"] == "在册"),
        "members": sum(1 for p in all_rows if p["is_member"] and p["member_status"] == "在供"),
        "che_members": sum(1 for p in all_rows if p["is_member"] and p["member_status"] == "在供" and p["member_branch"] == "CHE"),
        "stw_members": sum(1 for p in all_rows if p["is_member"] and p["member_status"] == "在供" and p["member_branch"] == "STW"),
        "both": sum(1 for p in all_rows if p["is_volunteer"] and p["volunteer_status"] == "在册"
                    and p["is_member"] and p["member_status"] == "在供"),
        "profile_done": len(active_volunteer_ids & profile_ids),
        "profile_pending": len(active_volunteer_ids & pending_ids),
        "profile_missing": len(active_volunteer_ids - profile_ids - pending_ids),
    }
    return render_template_string(PAGE, rows=rows, stats=stats, kinds=KINDS)


@people_bp.route("/edit/<path:person_key>", methods=["GET", "POST"])
def edit(person_key):
    p = find_person(person_key)
    if not p:
        flash("找不到这位人员。", "bad")
        return redirect(url_for("people.index"))

    if request.method == "POST":
        name = clean(request.form.get("name"))
        english_name = clean(request.form.get("english_name"))
        phone = normalize_phone(request.form.get("phone"))
        pin = clean(request.form.get("pin"))
        volunteer_branch = norm_branch(p.get("volunteer_branch") or "CHE")
        member_branch = norm_branch(p.get("member_branch") or "CHE")
        remark = clean(request.form.get("remark"))
        ic_number = clean(request.form.get("ic_number"))

        if not name:
            flash("姓名不能为空。", "bad")
            return render_template_string(EDIT_PAGE, p=p)

        if request.form.get("pin_from_phone") == "1":
            pin = default_pin(phone)

        # 同一个人同时存在两张表时，用一个 transaction 同步更新。
        with get_conn() as conn:
            with conn.cursor(cursor_factory=RealDictCursor) as cur:
                if p["is_member"]:
                    cur.execute("""
                        update members
                        set name=%s, english_name=%s, phone=%s, pin=%s,
                            branch=%s, remark=%s, ic_number=%s
                        where member_id=%s
                    """, (name, english_name, phone, pin, member_branch, remark,
                          ic_number, p["member_id"]))

                if p["is_volunteer"]:
                    cur.execute("""
                        update volunteers
                        set name=%s, phone=%s, pin=%s, branch=%s, remark=%s
                        where id=%s
                    """, (name, phone, pin, volunteer_branch, remark, p["volunteer_id"]))

        flash(f"✅ {name} 的个人资料已更新。", "ok")
        return redirect(url_for("people.edit", person_key=person_key))

    return render_template_string(EDIT_PAGE, p=p)



@people_bp.route("/add", methods=["GET", "POST"])
def add():
    suggested_che = next_member_id("CHE")
    suggested_stw = next_member_id("STW")
    form = request.form.to_dict()
    error = ""
    if request.method == "POST":
        branch = norm_branch(request.form.get("branch"))
        name = clean(request.form.get("name"))
        english_name = clean(request.form.get("english_name"))
        phone = normalize_phone(request.form.get("phone"))
        pin = clean(request.form.get("pin")) or default_pin(phone)
        ic_number = clean(request.form.get("ic_number"))
        remark = clean(request.form.get("remark"))
        member_id = normalize_member_id(branch, request.form.get("member_id"))
        if not name:
            error = "姓名不能为空。"
        elif not member_id:
            error = "月费佛友必须有月费编号。"
        elif member_exists(member_id):
            error = f"月费编号 {member_id} 已经存在。"
        if not error:
            with get_conn() as conn:
                with conn.cursor() as cur:
                    cur.execute("""insert into members(member_id,name,english_name,phone,pin,branch,status,remark,member_status,ic_number) values(%s,%s,%s,%s,%s,%s,'在供',%s,'在供',%s)""",(member_id,name,english_name,phone,pin,branch,remark,ic_number))
            flash(f"✅ 已新增月费佛友：{name}（{member_id}）", "ok")
            return redirect(url_for("people.edit", person_key="M:" + member_id.upper()))
    return render_template_string(ADD_PAGE,form=form,error=error,suggested_che=suggested_che,suggested_stw=suggested_stw)


@people_bp.route("/add-volunteer-role/<path:person_key>", methods=["GET", "POST"])
def add_volunteer_role(person_key):
    p = find_person(person_key)
    if not p or not p["is_member"]:
        flash("找不到这位月费佛友。", "bad")
        return redirect(url_for("people.index"))
    if p["is_volunteer"]:
        return redirect(url_for("people.edit", person_key=person_key))

    # 有月费者加入义工：义工编号直接跟月费编号一致。
    volunteer_id = p["member_id"]
    if volunteer_exists(volunteer_id):
        flash(f"义工编号 {volunteer_id} 已经存在，暂时没有写入。", "bad")
        return redirect(url_for("people.edit", person_key=person_key))

    if request.method == "POST":
        with get_conn() as conn:
            with conn.cursor(cursor_factory=RealDictCursor) as cur:
                cur.execute("""
                    insert into volunteers
                    (id, branch, name, status, phone, is_volunteer,
                     is_member, member_id, pin, remark)
                    values (%s,%s,%s,'在册',%s,true,true,%s,%s,%s)
                """, (
                    volunteer_id, "CHE", p["name"], p["phone"],
                    p["member_id"], p["pin"], p["remark"]
                ))
        flash(f"✅ {p['name']} 已加入义工，义工编号：{volunteer_id}", "ok")
        return redirect(url_for("people.edit", person_key=person_key))

    return render_template_string(r"""
    <!doctype html><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
    <style>body{background:#f5f7fb;font-family:Arial,"Microsoft YaHei",sans-serif}.c{max-width:620px;margin:35px auto;background:white;padding:25px;border-radius:18px}.b{padding:11px 16px;border:0;border-radius:10px;text-decoration:none;font-weight:800}.ok{background:#6558e8;color:white}.no{background:#e9ebf0;color:#333}</style>
    <div class="c"><h1>👥 加入义工</h1>
    <p><b>{{ p.name }}</b> 已经是月费佛友。</p>
    <p>系统会使用月费编号 <b>{{ p.member_id }}</b> 作为义工编号，不另外建立 800+ 编号。</p>
    <form method="post"><button class="b ok">确认加入义工</button> <a class="b no" href="{{ url_for('people.edit', person_key=p.key) }}">取消</a></form></div>
    """, p=p)


@people_bp.route("/add-member-role/<path:person_key>", methods=["GET", "POST"])
def add_member_role(person_key):
    p = find_person(person_key)
    if not p or not p["is_volunteer"]:
        flash("找不到这位义工。", "bad")
        return redirect(url_for("people.index"))
    if p["is_member"]:
        return redirect(url_for("people.edit", person_key=person_key))

    branch = norm_branch(request.form.get("branch") or p["branch"])
    suggested = next_member_id(branch)
    error = ""

    if request.method == "POST":
        member_id = normalize_member_id(branch, request.form.get("member_id"))
        english_name = clean(request.form.get("english_name"))
        ic_number = clean(request.form.get("ic_number"))

        if not member_id:
            error = "请输入月费编号。"
        elif member_exists(member_id):
            error = f"月费编号 {member_id} 已经存在。"
        else:
            with get_conn() as conn:
                with conn.cursor(cursor_factory=RealDictCursor) as cur:
                    cur.execute("""
                        insert into members
                        (member_id, name, english_name, phone, pin, branch,
                         status, remark, member_status, ic_number)
                        values (%s,%s,%s,%s,%s,%s,'在供',%s,'在供',%s)
                    """, (
                        member_id, p["name"], english_name, p["phone"], p["pin"],
                        branch, p["remark"], ic_number
                    ))
                    # 保留原 800+ 义工编号，避免旧签到/排班记录断掉。
                    cur.execute("""
                        update volunteers
                        set is_member=true, member_id=%s
                        where id=%s
                    """, (member_id, p["volunteer_id"]))

            new_key = "M:" + member_id.upper()
            flash(f"✅ {p['name']} 已加入月费，月费编号：{member_id}", "ok")
            return redirect(url_for("people.edit", person_key=new_key))

    return render_template_string(
        ADD_MEMBER_ROLE_PAGE, p=p, suggested=suggested, error=error
    )



@people_bp.route("/toggle-volunteer/<path:person_key>", methods=["POST"])
def toggle_volunteer(person_key):
    p = find_person(person_key)
    if not p or not p["is_volunteer"]:
        flash("找不到义工资料。", "bad")
        return redirect(url_for("people.index"))

    new_status = "退出" if p["volunteer_status"] == "在册" else "在册"
    db_query("""
        update volunteers
        set status=%s, is_volunteer=true
        where id=%s
    """, (new_status, p["volunteer_id"]))
    flash(f"义工状态已改为：{new_status}", "ok")
    return redirect(url_for("people.edit", person_key=person_key))



TRANSFER_MEMBER_PAGE = r"""
<!doctype html><html lang="zh"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>月费转会</title>
<style>body{background:#f5f7fb;font-family:Arial,"Microsoft YaHei"}.wrap{max-width:680px;margin:35px auto;padding:20px}.card{background:#fff;border-radius:20px;padding:24px}.row{background:#f6f7fb;border-radius:12px;padding:14px;margin:10px 0}.input{width:100%;padding:12px;border:1px solid #d7dce5;border-radius:10px;box-sizing:border-box;font-size:16px}.warn{background:#fff4dd;padding:13px;border-radius:11px;line-height:1.65;margin:15px 0}.btn{display:inline-block;padding:11px 16px;border:0;border-radius:10px;text-decoration:none;font-weight:800;cursor:pointer}.primary{background:#6558e8;color:#fff}.secondary{background:#e9ebf0;color:#333}</style></head><body><div class="wrap"><div class="card"><h1>🔄 月费转会</h1><div class="row"><b>{{ p.name }}</b><br>目前月费：{{ p.member_id }} · {{ p.member_branch }} · {{ p.member_status }}{% if p.is_volunteer %}<br>义工：{{ p.volunteer_id }} · {{ p.volunteer_branch }} · {{ p.volunteer_status }}{% endif %}</div><div class="warn">⚠️ 此操作只处理月费身份。义工编号、义工分会、签到及排班身份全部保持不变；过去使用旧月费编号的财政历史也不会被改写。</div><form method="post"><label><b>转往月费分会</b></label><input class="input" value="{{ target_branch }}" disabled><br><br><label><b>新月费编号</b></label><input class="input" name="new_member_id" value="{{ suggested }}" required><p><button class="btn primary" onclick="return confirm('确认将月费从 {{ p.member_id }} 转为 {{ suggested }}？义工身份不会改变。')">确认月费转会</button> <a class="btn secondary" href="{{ url_for('people.edit',person_key=p.key) }}">取消</a></p></form></div></div></body></html>
"""

@people_bp.route('/member-volunteer-application/<path:person_key>')
def start_member_volunteer_application(person_key):
    ensure_people_admin_and_application_tables()
    p=find_person(person_key)
    if not p or not p.get('is_member'):
        flash('找不到这位月费佛友。','bad'); return redirect(url_for('people.index'))
    if p.get('is_volunteer'):
        flash('这位佛友已经有义工身份，不需要重新申请。','bad'); return redirect(url_for('people.edit',person_key=person_key))
    existing=db_query("select id,status,invite_token from people_applications where upper(existing_member_id)=upper(%s) and status in ('draft','submitted','interview','probation') order by id desc limit 1",(p['member_id'],),fetchone=True)
    if existing:
        flash('这位佛友已经有进行中的义工申请，已带你到申请工作台。','ok'); return redirect(url_for('people.applications'))
    token=secrets.token_urlsafe(24)
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""insert into people_applications(chinese_name,english_name,ic_number,phone,existing_member_id,preferred_branch,status,invite_token,invited_at,updated_by) values(%s,%s,%s,%s,%s,'CHE','draft',%s,current_timestamp,%s)""",(p['name'],p.get('english_name',''),p.get('ic_number',''),normalize_phone(p.get('phone')),p['member_id'],token,session.get('people_admin_name')))
    flash(f"✅ 已为 {p['name']} 建立义工申请。请到「新义工申请」复制专属 WhatsApp 链接。",'ok')
    return redirect(url_for('people.applications'))

@people_bp.route('/transfer-member/<path:person_key>',methods=['GET','POST'])
def transfer_member(person_key):
    p=find_person(person_key)
    if not p or not p.get('is_member'):
        flash('找不到月费佛友资料。','bad'); return redirect(url_for('people.index'))
    target_branch='STW' if p['member_branch']=='CHE' else 'CHE'
    suggested=next_member_id(target_branch)
    if request.method=='POST':
        new_mid=normalize_member_id(target_branch,request.form.get('new_member_id'))
        if not new_mid or not new_mid.startswith(target_branch+'-'):
            flash(f'新月费编号必须是 {target_branch}-xxx 格式。','bad'); return redirect(url_for('people.transfer_member',person_key=person_key))
        if member_exists(new_mid):
            flash(f'月费编号 {new_mid} 已经存在。','bad'); return redirect(url_for('people.transfer_member',person_key=person_key))
        old_mid=p['member_id']
        with get_conn() as conn:
            with conn.cursor(cursor_factory=RealDictCursor) as cur:
                cur.execute("select * from members where member_id=%s for update",(old_mid,))
                old=cur.fetchone()
                if not old:
                    flash('原月费资料已经不存在，请重新载入。','bad'); return redirect(url_for('people.index'))
                # 先建立新月费身份，再把义工的 member_id 指向新编号；最后移除旧的当前会员记录。
                cur.execute("""insert into members(member_id,name,english_name,phone,pin,branch,status,remark,member_status,ic_number) values(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""",(new_mid,old.get('name'),old.get('english_name'),old.get('phone'),old.get('pin'),target_branch,old.get('status'),old.get('remark'),old.get('member_status'),old.get('ic_number')))
                cur.execute("update volunteers set member_id=%s where member_id=%s",(new_mid,old_mid))
                cur.execute("update people_applications set existing_member_id=%s where upper(existing_member_id)=upper(%s) and status!='formal'",(new_mid,old_mid))
                cur.execute("delete from members where member_id=%s",(old_mid,))
        flash(f'✅ 月费已从 {old_mid} 转为 {new_mid}。义工编号及义工分会没有改变。','ok')
        return redirect(url_for('people.edit',person_key='M:'+new_mid.upper()))
    return render_template_string(TRANSFER_MEMBER_PAGE,p=p,target_branch=target_branch,suggested=suggested)

@people_bp.route("/toggle-member/<path:person_key>", methods=["POST"])
def toggle_member(person_key):
    p = find_person(person_key)
    if not p or not p["is_member"]:
        flash("找不到月费佛友资料。", "bad")
        return redirect(url_for("people.index"))

    new_status = "停供" if p["member_status"] == "在供" else "在供"
    db_query("""
        update members
        set status=%s, member_status=%s
        where member_id=%s
    """, (new_status, new_status, p["member_id"]))
    flash(f"月费状态已改为：{new_status}", "ok")
    return redirect(url_for("people.edit", person_key=person_key))


@people_bp.route("/download")
def download():
    rows = filtered_people()
    data = []
    for p in rows:
        data.append({
            "编号": p["display_id"],
            "义工分会": p["volunteer_branch"],
            "月费分会": p["member_branch"],
            "姓名": p["name"],
            "英文名": p["english_name"],
            "电话号码": p["phone"],
            "PIN": p["pin"],
            "是否义工": "是" if p["is_volunteer"] else "否",
            "义工状态": p["volunteer_status"],
            "义工编号": p["volunteer_id"],
            "是否月费": "是" if p["is_member"] else "否",
            "月费状态": p["member_status"],
            "月费编号": p["member_id"],
            "身份证号码": p["ic_number"],
            "备注": p["remark"],
        })

    df = pd.DataFrame(data)
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="人员名单")
        ws = writer.book["人员名单"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for col in ws.columns:
            max_len = max((len(str(c.value or "")) for c in col), default=8)
            ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 3, 10), 28)

    output.seek(0)
    return send_file(
        output,
        as_attachment=True,
        download_name="人员名单.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ============================================================
# 义工个人档案 / 自助资料更新（v2）
# ============================================================

def ensure_people_profile_tables():
    """首次进入新功能时自动建立资料表；不修改现有 members / volunteers 结构。"""
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                create table if not exists people_profile_updates (
                    id bigserial primary key,
                    volunteer_id text not null,
                    name text,
                    english_name text,
                    ic_number text,
                    birth_date text,
                    phone text,
                    address text,
                    marital_status text,
                    occupation text,
                    education text,
                    skills text,
                    emergency_name text,
                    emergency_relation text,
                    emergency_phone text,
                    vegetarian_info text,
                    cooks_meat text,
                    other_practice text,
                    xlfm_source text,
                    xlfm_years text,
                    disciple_status text,
                    disciple_no text,
                    dabei text,
                    xin_jing text,
                    li_fo text,
                    small_mantras text,
                    xfz_weekly text,
                    xfz_total text,
                    original_join_date text,
                    note text,
                    photo_data bytea,
                    photo_mimetype text,
                    status text not null default 'pending',
                    submitted_at timestamp not null default current_timestamp,
                    reviewed_at timestamp,
                    review_note text
                )
            """)
            cur.execute("""
                create table if not exists people_profiles (
                    volunteer_id text primary key,
                    profile_json text not null default '{}',
                    photo_data bytea,
                    photo_mimetype text,
                    updated_at timestamp not null default current_timestamp
                )
            """)


def get_volunteer_for_self_login(volunteer_id, pin):
    volunteer_id = clean(volunteer_id).upper()
    raw = clean(volunteer_id)
    candidates = [raw]
    if raw.isdigit():
        candidates += [f"CHE-{int(raw)}", f"STW-{int(raw)}", str(int(raw))]
    rows = db_query("""
        select id, branch, name, phone, pin, status, is_volunteer, member_id
        from volunteers
        where upper(id)=any(%s)
    """, ([x.upper() for x in candidates],), fetchall=True) or []
    for row in rows:
        if clean(row.get('pin')) == clean(pin) and bool(row.get('is_volunteer')):
            return row
    return None


SELF_LOGIN_PAGE = r"""
<!doctype html><html lang="zh"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>义工个人资料更新</title><style>
body{background:#f5f7fb;font-family:Arial,"Microsoft YaHei",sans-serif}.wrap{max-width:620px;margin:30px auto;padding:18px}.card{background:#fff;border-radius:20px;padding:24px;box-shadow:0 2px 12px #00000010}
.input{width:100%;padding:13px;border:1px solid #d7dce5;border-radius:11px;box-sizing:border-box;font-size:17px;margin:7px 0 15px}.btn{width:100%;padding:13px;border:0;border-radius:11px;background:#6558e8;color:#fff;font-weight:800;font-size:17px}.note{background:#fff8df;padding:12px;border-radius:10px;line-height:1.55}.err{background:#fdecec;color:#922;padding:12px;border-radius:10px;margin-bottom:12px}
</style></head><body><div class="wrap"><div class="card"><h1>🙏 义工个人资料更新</h1>
<p>请输入您的义工编号与 PIN。验证成功后即可填写或更新个人档案。</p>
<div class="note">提交后不会直接修改正式资料，会先交由负责人审核确认。</div>
{% if error %}<div class="err">{{ error }}</div>{% endif %}
<form method="post"><label>义工编号</label><input class="input" name="volunteer_id" placeholder="例如 CHE-208 / 208 / 800" required>
<label>PIN</label><input class="input" name="pin" type="password" inputmode="numeric" required>
<button class="btn">进入我的资料</button></form></div></div></body></html>
"""

SELF_FORM_PAGE = r"""
<!doctype html><html lang="zh"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>更新个人资料</title>
<style>body{background:#f5f7fb;font-family:Arial,"Microsoft YaHei",sans-serif}.wrap{max-width:900px;margin:0 auto;padding:20px}.card{background:#fff;border-radius:20px;padding:22px;margin-bottom:15px}.grid{display:grid;grid-template-columns:1fr 1fr;gap:13px}.full{grid-column:1/-1}label{display:block;font-weight:800;margin-bottom:6px}.input{width:100%;padding:11px;border:1px solid #d7dce5;border-radius:10px;box-sizing:border-box;font-size:16px}.section{margin:5px 0 16px;padding-bottom:8px;border-bottom:2px solid #eee}.btn{padding:13px 18px;border:0;border-radius:11px;background:#6558e8;color:#fff;font-weight:800;font-size:16px}.locked{background:#f1f3f6;padding:12px;border-radius:10px;margin-bottom:14px}@media(max-width:700px){.grid{grid-template-columns:1fr}.full{grid-column:auto}}</style></head>
<body><div class="wrap"><div class="card"><h1>👤 {{ v.name }} 的个人档案</h1><div class="locked">义工编号：<b>{{ v.id }}</b>　义工分会：<b>{{ v.branch }}</b>（这两项不能自行修改）</div>
<form method="post" enctype="multipart/form-data">
<h2 class="section">基本资料</h2><div class="grid">
<div><label>中文姓名</label><input class="input" name="name" value="{{ d.get('name',v.name) }}" required></div><div><label>英文姓名</label><input class="input" name="english_name" value="{{ d.get('english_name','') }}"></div>
<div><label>身份证号码</label><input class="input" name="ic_number" value="{{ d.get('ic_number','') }}"></div><div><label>出生日期</label><input class="input" type="date" name="birth_date" value="{{ d.get('birth_date','') }}"></div>
<div><label>电话号码</label><input class="input" name="phone" value="{{ d.get('phone',v.phone or '') }}"></div><div><label>婚姻状况</label><select class="input" name="marital_status"><option value="">请选择</option>{% for x in ['单身','已婚','丧偶','离异','其他'] %}<option value="{{ x }}" {% if d.get('marital_status','')==x %}selected{% endif %}>{{ x }}</option>{% endfor %}</select></div>
<div class="full"><label>目前住址</label><textarea class="input" name="address" rows="2">{{ d.get('address','') }}</textarea></div>
<div><label>职业</label><input class="input" name="occupation" value="{{ d.get('occupation','') }}" placeholder="例如：电工、会计、教师、家庭主妇、退休、自由业"></div><div><label>最高学历</label><input class="input" name="education" value="{{ d.get('education','') }}"></div>
<div class="full"><label>个人技术 / 专长</label><input class="input" name="skills" value="{{ d.get('skills','') }}"></div>
<div><label>紧急联系人姓名</label><input class="input" name="emergency_name" value="{{ d.get('emergency_name','') }}"></div><div><label>关系</label><input class="input" name="emergency_relation" value="{{ d.get('emergency_relation','') }}"></div><div><label>紧急联系电话</label><input class="input" name="emergency_phone" value="{{ d.get('emergency_phone','') }}"></div>
<div><label>个人照片</label><input class="input" type="file" name="photo" accept="image/jpeg,image/png,image/webp"></div></div>
<h2 class="section">修学资料</h2><div class="grid">
<div><label>许愿全素 / 多久</label><input class="input" name="vegetarian_info" value="{{ d.get('vegetarian_info','') }}"></div><div><label>怎样认识 XLFM</label><input class="input" name="xlfm_source" value="{{ d.get('xlfm_source','') }}"></div>
<div><label>修 XLFM 多久</label><input class="input" name="xlfm_years" value="{{ d.get('xlfm_years','') }}"></div><div><label>弟子 / 师兄</label><input class="input" name="disciple_status" value="{{ d.get('disciple_status','') }}"></div><div><label>弟子编号</label><input class="input" name="disciple_no" value="{{ d.get('disciple_no','') }}"></div></div>
<h2 class="section">义工资料</h2><div class="grid"><div><label>大约什么时候开始做义工？</label><input class="input" name="original_join_date" value="{{ d.get('original_join_date','') }}" placeholder="例如：2018年、大约2019年、约10年前、不记得"><small style="color:#667085">不记得准确日期没关系，填写大概年份或时间即可。</small></div><div class="full"><label>补充说明</label><textarea class="input" name="note" rows="3">{{ d.get('note','') }}</textarea></div></div>
<p><label><input type="checkbox" required> 我确认以上资料由本人填写并正确。</label></p><button class="btn">📨 提交给负责人审核</button></form></div></div></body></html>
"""

SUCCESS_PAGE = r"""<!doctype html><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><style>body{font-family:Arial,"Microsoft YaHei";background:#f5f7fb}.c{max-width:600px;margin:60px auto;background:#fff;padding:30px;border-radius:20px;text-align:center}</style><div class="c"><h1>✅ 已提交</h1><p>感恩，您的个人资料已经送交负责人审核。</p><p>审核前不会直接更改正式义工资料。</p></div>"""

PROFILE_FIELDS = [
 ('name','中文姓名'),('english_name','英文姓名'),('ic_number','身份证号码'),('birth_date','出生日期'),('phone','电话号码'),('address','地址'),('marital_status','婚姻状况'),('occupation','职业'),('education','最高学历'),('skills','个人技术 / 专长'),('emergency_name','紧急联系人'),('emergency_relation','关系'),('emergency_phone','紧急联系电话'),('vegetarian_info','许愿全素'),('xlfm_source','认识 XLFM'),('xlfm_years','修学多久'),('disciple_status','弟子 / 师兄'),('disciple_no','弟子编号'),('original_join_date','大约开始做义工时间'),('note','补充说明')
]

REVIEW_PAGE = r"""
<!doctype html><html lang="zh"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>待审核资料更新</title><style>body{background:#f5f7fb;font-family:Arial,"Microsoft YaHei"}.wrap{max-width:1100px;margin:auto;padding:22px}.card{background:white;padding:18px;border-radius:16px;margin-bottom:12px}.btn{padding:9px 13px;border-radius:9px;text-decoration:none;font-weight:800;background:#6558e8;color:white}.muted{color:#667085}</style></head><body><div class="wrap"><h1>📝 待审核资料更新</h1><p><a href="{{ url_for('people.index') }}">← 返回人员管理中心</a></p>{% for r in rows %}<div class="card"><b>{{ r.volunteer_id }}　{{ r.name }}</b><div class="muted">提交：{{ r.submitted_at }}</div><p><a class="btn" href="{{ url_for('people.profile_review_detail', update_id=r.id) }}">查看及审核</a></p></div>{% else %}<div class="card">目前没有等待审核的资料。</div>{% endfor %}</div></body></html>
"""

REVIEW_DETAIL_PAGE = r"""
<!doctype html><html lang="zh"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>审核义工资料</title><style>body{background:#f5f7fb;font-family:Arial,"Microsoft YaHei"}.wrap{max-width:950px;margin:auto;padding:22px}.card{background:#fff;padding:22px;border-radius:18px}.row{display:grid;grid-template-columns:190px 1fr;gap:12px;padding:9px;border-bottom:1px solid #eee}.label{font-weight:800}.btn{padding:12px 18px;border:0;border-radius:10px;font-weight:800;cursor:pointer}.ok{background:#16834b;color:white}.no{background:#eee;color:#333}.photo{max-width:220px;max-height:260px;border-radius:12px}@media(max-width:650px){.row{grid-template-columns:1fr}}</style></head><body><div class="wrap"><div class="card"><h1>审核：{{ r.volunteer_id }} {{ r.name }}</h1>{% if r.photo_data %}<p><img class="photo" src="{{ url_for('people.profile_update_photo', update_id=r.id) }}"></p>{% endif %}{% for key,label in fields %}<div class="row"><div class="label">{{ label }}</div><div>{{ format_phone(r[key]) if key in ('phone','emergency_phone') else (r[key] or '-') }}</div></div>{% endfor %}<form method="post" style="margin-top:18px"><button class="btn ok" name="action" value="approve">✅ 批准更新</button> <button class="btn no" name="action" value="reject">退回 / 不采用</button></form></div></div></body></html>
"""

@people_bp.route('/self-update', methods=['GET','POST'])
def self_update_login():
    ensure_people_profile_tables()
    error=''
    if request.method=='POST':
        v=get_volunteer_for_self_login(request.form.get('volunteer_id'), request.form.get('pin'))
        if v:
            session['people_self_volunteer_id']=v['id']
            session['people_self_pin']=clean(request.form.get('pin'))
            return redirect(url_for('people.self_update_form'))
        error='义工编号或 PIN 不正确，请重新输入。'
    return render_template_string(SELF_LOGIN_PAGE,error=error)

@people_bp.route('/self-update/form', methods=['GET','POST'])
def self_update_form():
    ensure_people_profile_tables()
    volunteer_id=clean(session.get('people_self_volunteer_id'))
    pin=clean(session.get('people_self_pin'))
    v=get_volunteer_for_self_login(volunteer_id,pin)
    if not v: abort(403)
    existing=db_query("select profile_json from people_profiles where volunteer_id=%s",(v['id'],),fetchone=True)
    d={}
    if existing and existing.get('profile_json'):
        try:d=json.loads(existing['profile_json'])
        except Exception:d={}
        d['phone']=format_phone(d.get('phone'))
        d['emergency_phone']=format_phone(d.get('emergency_phone'))
    # 第一次填写时，尽量从现有 members / volunteers 带入已有资料，减少重复输入。
    if not d:
        d={'name': clean(v.get('name')), 'phone': format_phone(v.get('phone'))}
        member_id=clean(v.get('member_id'))
        if member_id:
            m=db_query("select english_name,ic_number,phone from members where member_id=%s",(member_id,),fetchone=True)
            if m:
                d['english_name']=clean(m.get('english_name'))
                d['ic_number']=clean(m.get('ic_number'))
                d['phone']=format_phone(m.get('phone')) or d['phone']
    if request.method=='POST':
        vals={k:clean(request.form.get(k)) for k,_ in PROFILE_FIELDS}
        vals['phone']=normalize_phone(vals.get('phone'))
        vals['emergency_phone']=normalize_phone(vals.get('emergency_phone'))
        photo=request.files.get('photo'); pdata=None; pmime=None
        if photo and photo.filename:
            pmime=clean(photo.mimetype).lower()
            if pmime not in ('image/jpeg','image/png','image/webp'):
                return '只接受 JPG、PNG 或 WEBP 照片。',400
            pdata=photo.read()
            if len(pdata)>5*1024*1024:return '照片不能超过 5MB。',400
        cols=[k for k,_ in PROFILE_FIELDS]
        sql_cols=', '.join(cols)
        placeholders=', '.join(['%s']*len(cols))
        with get_conn() as conn:
            with conn.cursor() as cur:
                cur.execute(f"""insert into people_profile_updates (volunteer_id,{sql_cols},photo_data,photo_mimetype,status) values (%s,{placeholders},%s,%s,'pending')""", [v['id']]+[vals[k] for k in cols]+[pdata,pmime])
        return render_template_string(SUCCESS_PAGE)
    return render_template_string(SELF_FORM_PAGE,v=v,d=d)

@people_bp.route('/profile-reviews')
def profile_reviews():
    ensure_people_profile_tables()
    rows=db_query("select id,volunteer_id,name,submitted_at from people_profile_updates where status='pending' order by submitted_at desc",fetchall=True) or []
    return render_template_string(REVIEW_PAGE,rows=rows)

@people_bp.route('/profile-review/<int:update_id>',methods=['GET','POST'])
def profile_review_detail(update_id):
    ensure_people_profile_tables()
    r=db_query('select * from people_profile_updates where id=%s',(update_id,),fetchone=True)
    if not r: abort(404)
    if request.method=='POST':
        action=request.form.get('action')
        if action=='approve':
            data={k:clean(r.get(k)) for k,_ in PROFILE_FIELDS}
            payload=json.dumps(data,ensure_ascii=False)
            with get_conn() as conn:
                with conn.cursor() as cur:
                    cur.execute("""insert into people_profiles(volunteer_id,profile_json,photo_data,photo_mimetype,updated_at) values(%s,%s,%s,%s,current_timestamp) on conflict(volunteer_id) do update set profile_json=excluded.profile_json, photo_data=coalesce(excluded.photo_data,people_profiles.photo_data), photo_mimetype=coalesce(excluded.photo_mimetype,people_profiles.photo_mimetype), updated_at=current_timestamp""",(r['volunteer_id'],payload,r.get('photo_data'),r.get('photo_mimetype')))
                    cur.execute("update people_profile_updates set status='approved',reviewed_at=current_timestamp where id=%s",(update_id,))
                    # 只同步安全的现有基础字段；编号、分会、身份、状态绝不由自助表修改。
                    # 姓名属于同一个人的基础资料：若义工连接了月费 member_id，批准后同步两边姓名。
                    # 注意：这里只同步姓名/电话，不会修改 volunteers.branch、members.branch 或任何编号。
                    phone = normalize_phone(data['phone'])
                    cur.execute(
                        "update volunteers set name=%s,phone=%s where id=%s",
                        (data['name'], phone, r['volunteer_id'])
                    )
                    cur.execute(
                        "select member_id from volunteers where id=%s",
                        (r['volunteer_id'],)
                    )
                    linked = cur.fetchone()
                    member_id = clean(linked.get("member_id")) if linked else None
                    if member_id:
                        cur.execute(
                            "update members set name=%s,phone=%s where member_id=%s",
                            (data['name'], phone, member_id)
                        )
            flash('✅ 义工个人档案已审核并更新。','ok')
        elif action=='reject':
            db_query("update people_profile_updates set status='rejected',reviewed_at=current_timestamp where id=%s",(update_id,))
            flash('该次提交已标记为不采用。','ok')
        return redirect(url_for('people.profile_reviews'))
    return render_template_string(REVIEW_DETAIL_PAGE, r=r, fields=PROFILE_FIELDS, format_phone=format_phone)

@people_bp.route('/profile-update-photo/<int:update_id>')
def profile_update_photo(update_id):
    ensure_people_profile_tables()
    r=db_query('select photo_data,photo_mimetype from people_profile_updates where id=%s',(update_id,),fetchone=True)
    if not r or not r.get('photo_data'): abort(404)
    return Response(bytes(r['photo_data']),mimetype=r.get('photo_mimetype') or 'image/jpeg',headers={'Cache-Control':'private, no-store'})


PROFILE_VIEW_PAGE = r"""
<!doctype html><html lang="zh"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>义工个人档案</title>
<style>body{background:#f5f7fb;font-family:Arial,"Microsoft YaHei",sans-serif}.wrap{max-width:1000px;margin:auto;padding:22px}.hero,.card{background:#fff;border-radius:18px;padding:22px;margin-bottom:14px}.head{display:flex;gap:20px;align-items:flex-start}.photo{width:150px;height:190px;object-fit:cover;border-radius:14px;background:#eee}.title{font-size:28px;font-weight:900}.muted{color:#667085}.section{font-size:20px;font-weight:900;margin:0 0 10px}.grid{display:grid;grid-template-columns:190px 1fr;gap:0}.label,.value{padding:9px;border-bottom:1px solid #eee}.label{font-weight:800}.btn{display:inline-block;padding:10px 15px;border-radius:10px;text-decoration:none;font-weight:800;background:#6558e8;color:#fff;margin-right:8px;border:0;cursor:pointer}.btn.green{background:#16834b}@media(max-width:650px){.head{display:block}.photo{margin-bottom:14px}.grid{grid-template-columns:1fr}.label{padding-bottom:2px;border-bottom:0}.value{padding-top:2px}}</style></head>
<body><div class="wrap">
<div class="hero"><div class="head">{% if has_photo %}<img class="photo" src="{{ url_for('people.profile_photo', volunteer_id=v.id) }}">{% endif %}<div><div class="title">{{ v.name }}</div><p><b>{{ v.id }}</b> · 义工分会 {{ v.branch }}</p><p class="muted">档案最后更新：{{ profile.updated_at or '-' }}</p><a class="btn" href="{{ url_for('people.index') }}">← 返回人员管理</a>{% if profile %}<button class="btn green" id="copyProfileWhatsAppBtn" type="button" onclick="copyProfileWhatsApp()">📱 复制 WhatsApp 个人资料</button>{% endif %}</div></div></div>
{% if not profile %}<div class="card"><h2>尚未建立完整个人档案</h2><p>这位义工还没有通过负责人审核的新版资料。</p></div>{% else %}
<div id="profileData">{% for title, keys in sections %}<div class="card profile-section"><div class="section">{{ title }}</div><div class="grid">{% for key in keys %}<div class="label">{{ labels[key] }}</div><div class="value">{{ d.get(key) or '-' }}</div>{% endfor %}</div></div>{% endfor %}</div>
<script>
function profileWhatsAppText(){
  const name={{ v.name|tojson }};
  const volunteerId={{ v.id|tojson }};
  const branch={{ v.branch|tojson }};
  let blocks=[];
  document.querySelectorAll('#profileData .profile-section').forEach(section=>{
    const title=(section.querySelector('.section')?.innerText || '').trim();
    const rows=section.querySelectorAll('.grid .label, .grid .value');
    let lines=[];
    for(let i=0;i<rows.length;i+=2){
      const label=(rows[i]?.innerText || '').trim();
      const value=(rows[i+1]?.innerText || '').trim() || '—';
      if(label) lines.push(`${label}：${value}`);
    }
    if(title) blocks.push(`【${title}】\n${lines.join('\n')}`);
  });
  return `📋 义工个人资料

👤 ${name || '—'}
🆔 义工编号：${volunteerId || '—'}
🏠 义工分会：${branch || '—'}

${blocks.join('\n\n')}`;
}
async function copyProfileWhatsApp(){
  const text=profileWhatsAppText();
  const btn=document.getElementById('copyProfileWhatsAppBtn');
  let ok=false;
  try{
    if(navigator.clipboard && window.isSecureContext){
      await navigator.clipboard.writeText(text);
      ok=true;
    }
  }catch(e){}
  if(!ok){
    try{
      const ta=document.createElement('textarea');
      ta.value=text; ta.style.position='fixed'; ta.style.opacity='0';
      document.body.appendChild(ta); ta.focus(); ta.select();
      ok=document.execCommand('copy'); ta.remove();
    }catch(e){}
  }
  if(ok){
    const old=btn.innerText;
    btn.innerText='✅ 已复制 WhatsApp 个人资料';
    setTimeout(()=>btn.innerText=old,1800);
  }else{
    window.prompt('浏览器无法自动复制，请按 Ctrl+C 复制以下资料：',text);
  }
}
</script>
{% endif %}</div></body></html>
"""

PROFILE_LABELS = dict(PROFILE_FIELDS)
PROFILE_SECTIONS = [
    ('基本资料', ['name','english_name','ic_number','birth_date','phone','address','marital_status','occupation','education','skills']),
    ('紧急联系人', ['emergency_name','emergency_relation','emergency_phone']),
    ('修学资料', ['vegetarian_info','xlfm_source','xlfm_years','disciple_status','disciple_no']),
    ('义工资料', ['original_join_date','note']),
]

@people_bp.route('/profile/<path:volunteer_id>')
def profile_view(volunteer_id):
    ensure_people_profile_tables()
    v=db_query("select id,branch,name,phone,status from volunteers where id=%s",(volunteer_id,),fetchone=True)
    if not v: abort(404)
    profile=db_query("select profile_json,photo_data,updated_at from people_profiles where volunteer_id=%s",(volunteer_id,),fetchone=True)
    d={}
    if profile and profile.get('profile_json'):
        try:d=json.loads(profile['profile_json'])
        except Exception:d={}
    return render_template_string(PROFILE_VIEW_PAGE,v=v,profile=profile,d=d,has_photo=bool(profile and profile.get('photo_data')),labels=PROFILE_LABELS,sections=PROFILE_SECTIONS)

@people_bp.route('/profile-photo/<path:volunteer_id>')
def profile_photo(volunteer_id):
    ensure_people_profile_tables()
    r=db_query('select photo_data,photo_mimetype from people_profiles where volunteer_id=%s',(volunteer_id,),fetchone=True)
    if not r or not r.get('photo_data'): abort(404)
    return Response(bytes(r['photo_data']),mimetype=r.get('photo_mimetype') or 'image/jpeg',headers={'Cache-Control':'private, no-store'})


# ============================================================
# 负责人权限 + 新义工申请（v4）
# ============================================================

ADMIN_ACCOUNT_SLOTS = [
    ('Willeam', 'Willeam'),
    ('康芬', '康芬'),
    ('俊达', '俊达'),
]


def ensure_people_admin_and_application_tables():
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                create table if not exists people_admin_users (
                    username text primary key,
                    display_name text not null,
                    password_hash text not null,
                    active boolean not null default true,
                    created_at timestamp not null default current_timestamp,
                    last_login_at timestamp
                )
            """)
            cur.execute("""
                create table if not exists people_applications (
                    id bigserial primary key,
                    chinese_name text not null,
                    english_name text,
                    age text,
                    ic_number text,
                    phone text,
                    address text,
                    occupation text,
                    education text,
                    skills text,
                    marital_status text,
                    emergency_name text,
                    emergency_relation text,
                    emergency_phone text,
                    recommender text,
                    existing_member_id text,
                    preferred_branch text not null default 'CHE',
                    vegetarian_info text,
                    cooks_meat text,
                    other_practice text,
                    xlfm_source text,
                    xlfm_years text,
                    disciple_status text,
                    disciple_no text,
                    previous_no_volunteer_reason text,
                    dabei text,
                    xin_jing text,
                    li_fo text,
                    small_mantras text,
                    xfz_weekly text,
                    xfz_total text,
                    photo_data bytea,
                    photo_mimetype text,
                    status text not null default 'submitted',
                    interview_date text,
                    interview_time text,
                    interviewers text,
                    interview_note text,
                    probation_start_date text,
                    volunteer_start_date text,
                    formal_volunteer_id text,
                    submitted_at timestamp not null default current_timestamp,
                    updated_at timestamp not null default current_timestamp,
                    updated_by text
                )
            """)
            # v6：新义工申请专属问题。旧数据库自动补栏，不影响已有申请。
            for col in (
                "application_reason text",
                "volunteer_experience text",
                "previous_volunteer_place text",
                "preferred_tasks text",
                "available_time text",
                "invite_token text",
                "applicant_submitted_at timestamp",
                "invited_at timestamp"
            ):
                cur.execute(f"alter table people_applications add column if not exists {col}")
            cur.execute("create unique index if not exists people_applications_invite_token_uq on people_applications(invite_token) where invite_token is not null")


def people_admin_count():
    ensure_people_admin_and_application_tables()
    r=db_query("select count(*) as n from people_admin_users where active=true",fetchone=True) or {}
    return int(r.get('n') or 0)


def people_admin_logged_in():
    username=clean(session.get('people_admin_username'))
    if not username:
        return False
    r=db_query("select username from people_admin_users where username=%s and active=true",(username,),fetchone=True)
    return bool(r)


ADMIN_LOGIN_PAGE = r"""
<!doctype html><html lang="zh"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>负责人登录</title>
<style>body{background:#f5f7fb;font-family:Arial,"Microsoft YaHei",sans-serif}.wrap{max-width:560px;margin:45px auto;padding:18px}.card{background:#fff;border-radius:20px;padding:25px;box-shadow:0 2px 12px #00000010}.input{width:100%;padding:13px;border:1px solid #d7dce5;border-radius:11px;box-sizing:border-box;font-size:17px;margin:7px 0 15px}.btn{width:100%;padding:13px;border:0;border-radius:11px;background:#6558e8;color:#fff;font-weight:800;font-size:17px}.err{background:#fdecec;color:#922;padding:12px;border-radius:10px;margin-bottom:12px}.note{background:#eef7ff;padding:12px;border-radius:10px;line-height:1.6}</style></head><body><div class="wrap"><div class="card"><h1>🔐 人员管理负责人登录</h1><div class="note">此处只供三位负责人使用。义工更新资料与新义工申请不需要进入这里。</div>{% if error %}<div class="err">{{ error }}</div>{% endif %}<form method="post"><label>账号</label><input class="input" name="username" autocomplete="username" required><label>密码</label><input class="input" name="password" type="password" autocomplete="current-password" required><button class="btn">登录人员管理中心</button></form></div></div></body></html>
"""

ADMIN_SETUP_PAGE = r"""
<!doctype html><html lang="zh"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>首次建立负责人账号</title>
<style>body{background:#f5f7fb;font-family:Arial,"Microsoft YaHei",sans-serif}.wrap{max-width:760px;margin:25px auto;padding:18px}.card{background:#fff;border-radius:20px;padding:25px}.grid{display:grid;grid-template-columns:1fr 1fr;gap:12px}.full{grid-column:1/-1}.input{width:100%;padding:12px;border:1px solid #d7dce5;border-radius:10px;box-sizing:border-box;font-size:16px}.slot{background:#f7f8fb;padding:15px;border-radius:13px;margin-bottom:12px}.btn{padding:13px 18px;border:0;border-radius:11px;background:#6558e8;color:#fff;font-weight:800}.err{background:#fdecec;color:#922;padding:12px;border-radius:10px;margin-bottom:12px}@media(max-width:650px){.grid{grid-template-columns:1fr}.full{grid-column:auto}}</style></head><body><div class="wrap"><div class="card"><h1>🔐 首次建立三位负责人账号</h1><p>三个人权限完全相同。请各自使用自己的账号和密码，方便以后知道是谁审核资料。</p>{% if error %}<div class="err">{{ error }}</div>{% endif %}<form method="post">{% for key,label in slots %}<div class="slot"><h3>{{ label }}</h3><div class="grid"><div><label>登录账号</label><input class="input" name="username_{{ loop.index }}" value="{{ key }}" required></div><div><label>密码（至少8位）</label><input class="input" type="password" name="password_{{ loop.index }}" required></div></div></div>{% endfor %}<button class="btn">建立三个负责人账号</button></form></div></div></body></html>
"""

@people_bp.route('/admin/setup',methods=['GET','POST'])
def admin_setup():
    ensure_people_admin_and_application_tables()
    if people_admin_count()>0:
        return redirect(url_for('people.admin_login'))
    error=''
    if request.method=='POST':
        users=[]; seen=set()
        for i,(_,display_name) in enumerate(ADMIN_ACCOUNT_SLOTS,1):
            username=clean(request.form.get(f'username_{i}'))
            password=request.form.get(f'password_{i}') or ''
            if not username or len(password)<8:
                error='三个账号都必须填写；密码至少 8 位。'; break
            if username.lower() in seen:
                error='三个登录账号不能重复。'; break
            seen.add(username.lower())
            users.append((username,display_name,generate_password_hash(password)))
        if not error:
            with get_conn() as conn:
                with conn.cursor() as cur:
                    for row in users:
                        cur.execute("insert into people_admin_users(username,display_name,password_hash) values(%s,%s,%s)",row)
            flash('✅ 三位负责人账号已经建立，请登录。','ok')
            return redirect(url_for('people.admin_login'))
    return render_template_string(ADMIN_SETUP_PAGE,error=error,slots=ADMIN_ACCOUNT_SLOTS)

@people_bp.route('/admin/login',methods=['GET','POST'])
def admin_login():
    ensure_people_admin_and_application_tables()
    if people_admin_count()==0:
        return redirect(url_for('people.admin_setup'))
    error=''
    if request.method=='POST':
        username=clean(request.form.get('username'))
        password=request.form.get('password') or ''
        r=db_query("select username,display_name,password_hash from people_admin_users where username=%s and active=true",(username,),fetchone=True)
        if r and check_password_hash(r['password_hash'],password):
            session.pop('people_self_volunteer_id',None); session.pop('people_self_pin',None)
            session['people_admin_username']=r['username']; session['people_admin_name']=r['display_name']
            db_query("update people_admin_users set last_login_at=current_timestamp where username=%s",(r['username'],))
            return redirect(url_for('people.index'))
        error='账号或密码不正确。'
    return render_template_string(ADMIN_LOGIN_PAGE,error=error)

ADMIN_CHANGE_PASSWORD_PAGE = r"""
<!doctype html><html lang="zh"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>修改我的密码</title>
<style>body{background:#f5f7fb;font-family:Arial,"Microsoft YaHei",sans-serif}.wrap{max-width:580px;margin:40px auto;padding:18px}.card{background:#fff;border-radius:20px;padding:25px;box-shadow:0 2px 12px #00000010}.input{width:100%;padding:13px;border:1px solid #d7dce5;border-radius:11px;box-sizing:border-box;font-size:17px;margin:7px 0 15px}.btn{padding:12px 17px;border:0;border-radius:11px;font-weight:800;font-size:16px;cursor:pointer;text-decoration:none;display:inline-block}.primary{background:#6558e8;color:#fff}.secondary{background:#eef0f5;color:#333}.note{background:#eef7ff;padding:12px;border-radius:10px;line-height:1.6;margin-bottom:14px}.err{background:#fdecec;color:#922;padding:12px;border-radius:10px;margin-bottom:12px}</style></head>
<body><div class="wrap"><div class="card"><h1>🔑 修改我的密码</h1><div class="note">当前负责人：<b>{{ admin_name }}</b><br>这里只会修改您自己的登录密码，不会影响另外两位负责人。</div>{% if error %}<div class="err">{{ error }}</div>{% endif %}<form method="post"><label>目前密码</label><input class="input" type="password" name="current_password" autocomplete="current-password" required><label>新密码（至少 8 位）</label><input class="input" type="password" name="new_password" autocomplete="new-password" minlength="8" required><label>再次输入新密码</label><input class="input" type="password" name="confirm_password" autocomplete="new-password" minlength="8" required><button class="btn primary" type="submit">💾 更新密码</button> <a class="btn secondary" href="{{ url_for('people.index') }}">取消</a></form></div></div></body></html>
"""

@people_bp.route('/admin/change-password', methods=['GET','POST'])
def admin_change_password():
    ensure_people_admin_and_application_tables()
    username=clean(session.get('people_admin_username'))
    if not username:
        return redirect(url_for('people.admin_login'))
    r=db_query("select username,display_name,password_hash from people_admin_users where username=%s and active=true",(username,),fetchone=True)
    if not r:
        session.pop('people_admin_username',None); session.pop('people_admin_name',None)
        return redirect(url_for('people.admin_login'))
    error=''
    if request.method=='POST':
        current_password=request.form.get('current_password') or ''
        new_password=request.form.get('new_password') or ''
        confirm_password=request.form.get('confirm_password') or ''
        if not check_password_hash(r['password_hash'],current_password):
            error='目前密码不正确。'
        elif len(new_password)<8:
            error='新密码至少需要 8 位。'
        elif new_password!=confirm_password:
            error='两次输入的新密码不一致。'
        elif new_password==current_password:
            error='新密码不能与目前密码相同。'
        else:
            db_query("update people_admin_users set password_hash=%s where username=%s",(generate_password_hash(new_password),username))
            flash('✅ 您的负责人登录密码已经更新。','ok')
            return redirect(url_for('people.index'))
    return render_template_string(ADMIN_CHANGE_PASSWORD_PAGE,error=error,admin_name=r.get('display_name') or username)

@people_bp.route('/admin/logout')
def admin_logout():
    session.pop('people_admin_username',None); session.pop('people_admin_name',None)
    return redirect(url_for('people.admin_login'))

# Blueprint 统一保护：只有下列公开入口不需要负责人登录。
@people_bp.before_request
def protect_people_admin_routes():
    endpoint=(request.endpoint or '').split('.')[-1]
    public={
        'admin_setup','admin_login','self_update_login','self_update_form',
        'new_application','invited_application','application_success','static'
    }
    if endpoint in public:
        return None
    if not people_admin_logged_in():
        return redirect(url_for('people.admin_login',next=request.path))


APPLICATION_FIELDS = [
 ('chinese_name','中文姓名'),('english_name','英文姓名'),('age','年龄'),('ic_number','身份证号码'),('phone','电话号码'),('address','住址'),('occupation','职业'),('education','最高学历'),('skills','技术 / 专长'),('marital_status','婚姻状况'),('emergency_name','紧急联系人'),('emergency_relation','关系'),('emergency_phone','紧急联系电话'),('recommender','推荐人'),('existing_member_id','现有月费编号（如有）'),('vegetarian_info','许愿全素 / 多久'),('cooks_meat','是否煮荤'),('other_practice','其它法门'),('xlfm_source','怎样认识 XLFM'),('xlfm_years','修 XLFM 多久'),('disciple_status','弟子 / 师兄'),('disciple_no','弟子编号'),('previous_no_volunteer_reason','之前没有做义工原因'),('application_reason','为什么想申请做义工'),('volunteer_experience','过去义工经验'),('previous_volunteer_place','曾在哪个共修会 / 观音堂做过义工'),('preferred_tasks','希望参与的岗位'),('available_time','通常可以参与义工的时间'),('dabei','大悲咒'),('xin_jing','心经'),('li_fo','礼佛'),('small_mantras','小咒文'),('xfz_weekly','每星期 XFZ'),('xfz_total','累计 XFZ')
]

NEW_APPLICATION_PAGE = r"""
<!doctype html><html lang="zh"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>新义工申请</title>
<style>body{background:#f5f7fb;font-family:Arial,"Microsoft YaHei",sans-serif}.wrap{max-width:900px;margin:auto;padding:20px}.card{background:#fff;border-radius:20px;padding:22px}.grid{display:grid;grid-template-columns:1fr 1fr;gap:13px}.full{grid-column:1/-1}.section{margin:22px 0 13px;padding-bottom:8px;border-bottom:2px solid #eee}.input{width:100%;padding:11px;border:1px solid #d7dce5;border-radius:10px;box-sizing:border-box;font-size:16px}label{display:block;font-weight:800;margin-bottom:6px}.btn{margin-top:18px;padding:13px 20px;border:0;border-radius:11px;background:#6558e8;color:white;font-weight:800;font-size:17px}.note{background:#fff8df;padding:12px;border-radius:10px;line-height:1.6}.err{background:#fdecec;color:#922;padding:12px;border-radius:10px;margin-bottom:12px}@media(max-width:700px){.grid{grid-template-columns:1fr}.full{grid-column:auto}}</style></head><body><div class="wrap"><div class="card"><h1>🙏 新义工申请表</h1><div class="note">提交申请不会马上成为正式义工。负责人会先安排面试，之后再进入试用及正式加入流程。</div>{% if error %}<div class="err">{{ error }}</div>{% endif %}<form method="post" enctype="multipart/form-data">
<h2 class="section">基本资料</h2><div class="grid"><div><label>推荐人</label><input class="input" name="recommender" value="{{ d.get('recommender','') }}"></div><div><label>中文姓名 *</label><input class="input" name="chinese_name" value="{{ d.get('chinese_name','') }}" required></div><div><label>英文姓名</label><input class="input" name="english_name" value="{{ d.get('english_name','') }}"></div><div><label>年龄</label><input class="input" name="age" inputmode="numeric"></div><div><label>身份证号码</label><input class="input" name="ic_number" value="{{ d.get('ic_number','') }}"></div><div><label>电话号码 *</label><input class="input" name="phone" value="{{ format_phone(d.get('phone','')) }}" placeholder="例如 016-2038622" required></div><div class="full"><label>住址</label><textarea class="input" name="address" rows="2"></textarea></div><div><label>职业</label><input class="input" name="occupation" value="{{ d.get('occupation','') }}" placeholder="例如：电工、会计、教师、家庭主妇、退休、自由业"></div><div><label>最高学历</label><input class="input" name="education"></div><div><label>技术 / 专长</label><input class="input" name="skills"></div><div><label>婚姻状况</label><select class="input" name="marital_status"><option value="">请选择</option>{% for x in ['单身','已婚','丧偶','离异','其他'] %}<option value="{{ x }}" {% if d.get('marital_status','')==x %}selected{% endif %}>{{ x }}</option>{% endfor %}</select></div><div><label>紧急联系人</label><input class="input" name="emergency_name"></div><div><label>关系</label><input class="input" name="emergency_relation"></div><div><label>紧急联系电话</label><input class="input" name="emergency_phone" placeholder="例如 016-2038622"></div><div><label>现有月费编号</label>{% if d.get('existing_member_id') %}<input class="input" name="existing_member_id" value="{{ d.get('existing_member_id') }}" readonly><small>负责人已从人员档案连接此月费编号，请勿更改。</small>{% else %}<input class="input" name="existing_member_id" placeholder="如已经是月费佛友，例如 CHE-208">{% endif %}</div><div><label>个人照片 *</label><input class="input" type="file" name="photo" accept="image/jpeg,image/png,image/webp" required></div></div>
<h2 class="section">修学资料</h2><div class="grid"><div><label>许愿全素 / 多久</label><input class="input" name="vegetarian_info"></div><div><label>是否有煮荤给家人</label><input class="input" name="cooks_meat"></div><div><label>是否同时修其它法门</label><input class="input" name="other_practice"></div><div><label>怎样认识 XLFM</label><input class="input" name="xlfm_source"></div><div><label>修 XLFM 多久</label><input class="input" name="xlfm_years"></div><div><label>弟子 / 师兄</label><input class="input" name="disciple_status"></div><div><label>弟子编号</label><input class="input" name="disciple_no"></div><div class="full"><label>之前没有做义工的原因</label><textarea class="input" name="previous_no_volunteer_reason" rows="2"></textarea></div></div>
<h2 class="section">义工申请资料</h2><div class="grid"><div class="full"><label>为什么想申请做义工？</label><textarea class="input" name="application_reason" rows="2"></textarea></div><div class="full"><label>过去有没有做过义工？有什么经验？</label><textarea class="input" name="volunteer_experience" rows="2"></textarea></div><div class="full"><label>曾在哪个共修会 / 观音堂做过义工？</label><input class="input" name="previous_volunteer_place"></div><div><label>希望参与哪些岗位？</label><input class="input" name="preferred_tasks" placeholder="例如：值班、佛台、供花、膳食"></div><div><label>通常什么时间可以参与？</label><input class="input" name="available_time" placeholder="例如：星期日 / 晚上 / 大日子"></div></div>
<h2 class="section">每日功课（供负责人面试了解）</h2><div class="grid"><div><label>大悲咒</label><input class="input" name="dabei"></div><div><label>心经</label><input class="input" name="xin_jing"></div><div><label>礼佛</label><input class="input" name="li_fo"></div><div><label>小咒文</label><input class="input" name="small_mantras"></div><div><label>每星期 XFZ</label><input class="input" name="xfz_weekly"></div><div><label>累计 XFZ</label><input class="input" name="xfz_total"></div></div><p><label><input type="checkbox" required> 我确认以上资料由本人填写并正确。</label></p><button class="btn">📨 提交新义工申请</button></form></div></div></body></html>
"""

APPLICATION_SUCCESS_PAGE = r"""<!doctype html><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><style>body{font-family:Arial,"Microsoft YaHei";background:#f5f7fb}.c{max-width:600px;margin:60px auto;background:#fff;padding:30px;border-radius:20px;text-align:center}</style><div class="c"><h1>✅ 申请已提交</h1><p>感恩，负责人会查看资料并安排后续面试。</p><p>提交申请并不代表已经成为正式义工。</p></div>"""

@people_bp.route('/apply')
def new_application():
    return render_template_string(r"""<!doctype html><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><style>body{font-family:Arial,"Microsoft YaHei";background:#f5f7fb}.c{max-width:620px;margin:60px auto;background:#fff;padding:28px;border-radius:20px;text-align:center}.n{background:#eef7ff;padding:14px;border-radius:11px;line-height:1.7}</style><div class="c"><h1>🙏 新义工申请</h1><div class="n">新义工申请采用专属邀请方式。<br>请使用负责人通过 WhatsApp 发给您的个人申请链接填写资料。</div></div>""")

@people_bp.route('/apply/<token>',methods=['GET','POST'])
def invited_application(token):
    ensure_people_admin_and_application_tables()
    r=db_query("select * from people_applications where invite_token=%s",(clean(token),),fetchone=True)
    if not r: abort(404)
    if clean(r.get("status")) != "draft":
        return render_template_string(APPLICATION_SUCCESS_PAGE)
    error=""
    if request.method=="POST":
        vals={k:clean(request.form.get(k)) for k,_ in APPLICATION_FIELDS}
        vals["phone"]=normalize_phone(vals.get("phone"))
        vals["emergency_phone"]=normalize_phone(vals.get("emergency_phone"))
        vals["preferred_branch"]="CHE"
        if not vals["chinese_name"] or not vals["phone"]:
            error="中文姓名和电话号码必须填写。"
        photo=request.files.get("photo"); pdata=None; pmime=None
        if not error:
            if not photo or not photo.filename:
                error="请上传个人照片。"
            else:
                pmime=clean(photo.mimetype).lower()
                if pmime not in ("image/jpeg","image/png","image/webp"):
                    error="照片只接受 JPG、PNG 或 WEBP。"
                else:
                    pdata=photo.read()
                    if len(pdata)>5*1024*1024: error="照片不能超过 5MB。"
        if not error:
            cols=[k for k,_ in APPLICATION_FIELDS]
            sets=",".join([f"{k}=%s" for k in cols])
            params=[vals[k] for k in cols]+[pdata,pmime,token]
            db_query(f"update people_applications set {sets},photo_data=%s,photo_mimetype=%s,status='submitted',applicant_submitted_at=current_timestamp,updated_at=current_timestamp where invite_token=%s and status='draft'",params)
            return redirect(url_for("people.application_success"))
        d=vals
    else:
        d=dict(r)
    return render_template_string(NEW_APPLICATION_PAGE,error=error,d=d,format_phone=format_phone)

@people_bp.route('/apply/success')
def application_success():
    return render_template_string(APPLICATION_SUCCESS_PAGE)

APPLICATION_LIST_PAGE = r"""
<!doctype html><html lang="zh"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>新义工申请</title><style>body{background:#f5f7fb;font-family:Arial,"Microsoft YaHei"}.wrap{max-width:1100px;margin:auto;padding:22px}.top{display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap}.card{background:#fff;border-radius:16px;padding:18px;margin-bottom:12px}.btn{display:inline-block;padding:10px 14px;border-radius:9px;text-decoration:none;font-weight:800;background:#6558e8;color:#fff;border:0;cursor:pointer}.green{background:#16834b}.badge{padding:5px 9px;border-radius:999px;background:#eef0f5;font-weight:800}.muted{color:#667085}.invite{background:#eef7ff;border-radius:12px;padding:12px;margin-top:10px;line-height:1.6}.urlbox{width:100%;box-sizing:border-box;padding:9px;border:1px solid #ccd6e3;border-radius:8px;background:#fff}</style></head><body><div class="wrap"><div class="top"><div><h1>🆕 新义工申请</h1><p><a href="{{ url_for('people.index') }}">← 返回人员管理中心</a></p></div><a class="btn green" href="{{ url_for('people.create_application_invite') }}">➕ 建立新义工申请</a></div>
{% for r in rows %}<div class="card"><b>{{ r.chinese_name }}</b>　<span class="badge">{{ labels.get(r.status,r.status) }}</span><div class="muted">{{ format_phone(r.phone) }} · CHE · {% if r.status=='draft' %}等待申请人填写{% else %}提交 {{ r.applicant_submitted_at or r.submitted_at }}{% endif %}</div>
{% if r.status=='draft' %}<div class="invite"><b>📲 专属申请链接</b><br><input class="urlbox" id="url{{ r.id }}" readonly value="https://gyt-checkin.onrender.com{{ url_for('people.invited_application',token=r.invite_token) }}"><p><button class="btn" id="copybtn{{ r.id }}" type="button" data-name="{{ r.chinese_name|e }}" onclick="copyInvite({{ r.id }}, this.dataset.name)">📋 复制 WhatsApp 邀请</button></p></div>{% else %}<p><a class="btn" href="{{ url_for('people.application_detail', application_id=r.id) }}">查看申请档案</a></p>{% endif %}</div>{% else %}<div class="card">目前没有新义工申请。点击右上方「建立新义工申请」开始。</div>{% endfor %}</div>
<script>
async function copyInvite(id,name){
  const u=document.getElementById('url'+id).value;
  const btn=document.getElementById('copybtn'+id);
  const t=`${name} 师兄/佛友您好 🙏\n\n请点击以下专属链接填写新义工申请资料：\n${u}\n\n此链接为您的个人申请链接，请勿转发给其他人。\n填写完成并提交后，负责人会再与您联系安排面试。感恩 🙏`;
  let ok=false;
  try{
    if(navigator.clipboard && window.isSecureContext){
      await navigator.clipboard.writeText(t);
      ok=true;
    }
  }catch(e){}
  if(!ok){
    const ta=document.createElement('textarea');
    ta.value=t;
    ta.setAttribute('readonly','');
    ta.style.position='fixed';
    ta.style.left='-9999px';
    document.body.appendChild(ta);
    ta.focus();
    ta.select();
    try{ ok=document.execCommand('copy'); }catch(e){ ok=false; }
    document.body.removeChild(ta);
  }
  if(ok){
    const old=btn.textContent;
    btn.textContent='✅ 已复制';
    setTimeout(()=>{btn.textContent=old;},1800);
  }else{
    window.prompt('浏览器无法自动复制，请按 Ctrl+C 复制以下 WhatsApp 邀请文字：',t);
  }
}
</script></body></html>
"""

CREATE_INVITE_PAGE = r"""
<!doctype html><html lang="zh"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>建立新义工申请</title><style>body{background:#f5f7fb;font-family:Arial,"Microsoft YaHei"}.wrap{max-width:650px;margin:35px auto;padding:20px}.card{background:#fff;border-radius:18px;padding:24px}.input{width:100%;padding:12px;border:1px solid #d7dce5;border-radius:10px;box-sizing:border-box;font-size:16px;margin:7px 0 15px}.btn{padding:12px 16px;border:0;border-radius:10px;background:#6558e8;color:#fff;font-weight:800;cursor:pointer}.note{background:#eef7ff;padding:12px;border-radius:10px;line-height:1.6}.err{background:#fdecec;color:#922;padding:12px;border-radius:10px;margin-bottom:12px}</style></head><body><div class="wrap"><div class="card"><p><a href="{{ url_for('people.applications') }}">← 返回新义工申请</a></p><h1>➕ 建立新义工申请</h1><div class="note">负责人先输入申请人的姓名和电话。建立后系统会产生一个专属链接，再通过 WhatsApp 发给申请人本人填写完整资料。</div>{% if error %}<div class="err">{{ error }}</div>{% endif %}<form method="post"><label>申请人中文姓名 *</label><input class="input" name="chinese_name" value="{{ request.form.get('chinese_name','') }}" required><label>电话号码 *</label><input class="input" name="phone" value="{{ request.form.get('phone','') }}" placeholder="例如 016-2038622" required><button class="btn">建立并产生专属链接</button></form></div></div></body></html>
"""

APPLICATION_STATUS_LABELS={'draft':'待填写','submitted':'待处理','interview':'待面试','probation':'试用期','formal':'正式加入','rejected':'不采用'}

APPLICATION_DETAIL_PAGE = r"""
<!doctype html><html lang="zh"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>新义工申请档案</title><style>body{background:#f5f7fb;font-family:Arial,"Microsoft YaHei"}.wrap{max-width:1000px;margin:auto;padding:22px}.card{background:#fff;border-radius:18px;padding:22px;margin-bottom:14px}.head{display:flex;gap:18px}.photo{width:150px;height:190px;object-fit:cover;border-radius:13px}.grid{display:grid;grid-template-columns:210px 1fr}.label,.value{padding:8px;border-bottom:1px solid #eee}.label{font-weight:800}.input{width:100%;padding:10px;border:1px solid #d7dce5;border-radius:9px;box-sizing:border-box}.actions{display:flex;gap:8px;flex-wrap:wrap;margin-top:14px}.btn{padding:11px 14px;border:0;border-radius:9px;font-weight:800;cursor:pointer}.purple{background:#6558e8;color:#fff}.green{background:#16834b;color:#fff}.grey{background:#eee}.red{background:#fde8e8;color:#922}@media(max-width:650px){.head{display:block}.grid{grid-template-columns:1fr}.label{border-bottom:0;padding-bottom:2px}.value{padding-top:2px}}</style></head><body><div class="wrap"><p><a href="{{ url_for('people.applications') }}">← 返回新义工申请</a></p><div class="card"><div class="head"><img class="photo" src="{{ url_for('people.application_photo',application_id=r.id) }}"><div><h1>{{ r.chinese_name }}</h1><p>状态：<b>{{ status_labels.get(r.status,r.status) }}</b></p><p>申请日期：{{ r.submitted_at }}</p>{% if r.formal_volunteer_id %}<p>正式义工编号：<b>{{ r.formal_volunteer_id }}</b></p>{% endif %}</div></div></div><div class="card"><h2>申请资料</h2><div class="actions" style="margin:0 0 14px"><button class="btn green" id="copyWhatsAppBtn" type="button" onclick="copyApplicationWhatsApp()">📱 复制 WhatsApp 资料</button></div><div class="grid" id="applicationData">{% for key,label in fields %}<div class="label">{{ label }}</div><div class="value">{{ r[key] or '-' }}</div>{% endfor %}</div></div><div class="card"><h2>负责人处理</h2><form method="post"><div class="grid"><div class="label">面试日期</div><div class="value"><input class="input" type="date" name="interview_date" value="{{ r.interview_date or '' }}"></div><div class="label">面试时间</div><div class="value"><input class="input" type="time" name="interview_time" value="{{ r.interview_time or '' }}"></div><div class="label">面试负责人</div><div class="value"><input class="input" name="interviewers" value="{{ r.interviewers or '' }}" placeholder="例如 Willeam、康芬"></div><div class="label">面试备注</div><div class="value"><textarea class="input" name="interview_note" rows="3">{{ r.interview_note or '' }}</textarea></div><div class="label">试用开始日期</div><div class="value"><input class="input" type="date" name="probation_start_date" value="{{ r.probation_start_date or '' }}"></div><div class="label">正式开始义工日期</div><div class="value"><input class="input" type="date" name="volunteer_start_date" value="{{ r.volunteer_start_date or '' }}"></div><div class="label">正式义工分会</div><div class="value"><b>CHE</b><br><small>新义工申请统一属于 CHE；不会改变申请人原有的月费分会。</small></div><div class="label">指定义工编号（特殊情况）</div><div class="value"><input class="input" name="manual_volunteer_id" placeholder="一般留空；系统会按规则自动决定义工编号"><small>CHE 月费佛友可沿用 CHE 月费编号；STW 月费佛友或纯新义工会使用 800+ 义工编号。</small></div></div><div class="actions"><button class="btn grey" name="action" value="save">💾 保存记录</button><button class="btn purple" name="action" value="interview">🗣️ 设为待面试</button><button class="btn purple" name="action" value="probation">🟡 进入试用期</button>{% if r.status!='formal' %}<button class="btn green" name="action" value="formal" onclick="return confirm('确认正式加入义工？系统将建立正式义工编号。')">✅ 正式加入义工</button>{% endif %}<button class="btn red" name="action" value="reject">不采用</button></div></form></div></div>
<script>
function applicationWhatsAppText(){
  const name={{ r.chinese_name|tojson }};
  const status={{ status_labels.get(r.status,r.status)|tojson }};
  const rows=document.querySelectorAll('#applicationData .label, #applicationData .value');
  let lines=[];
  for(let i=0;i<rows.length;i+=2){
    const label=(rows[i]?.innerText || '').trim();
    const value=(rows[i+1]?.innerText || '').trim() || '—';
    if(label) lines.push(`${label}：${value}`);
  }
  return `🆕 新义工申请资料

👤 ${name || '—'}
📌 申请状态：${status || '—'}

${lines.join('\n')}`;
}
async function copyApplicationWhatsApp(){
  const text=applicationWhatsAppText();
  const btn=document.getElementById('copyWhatsAppBtn');
  let ok=false;
  try{
    if(navigator.clipboard && window.isSecureContext){ await navigator.clipboard.writeText(text); ok=true; }
  }catch(e){}
  if(!ok){
    try{
      const ta=document.createElement('textarea'); ta.value=text; ta.style.position='fixed'; ta.style.opacity='0';
      document.body.appendChild(ta); ta.focus(); ta.select(); ok=document.execCommand('copy'); ta.remove();
    }catch(e){}
  }
  if(ok){
    const old=btn.innerText; btn.innerText='✅ 已复制 WhatsApp 资料';
    setTimeout(()=>btn.innerText=old,1800);
  }else{
    window.prompt('浏览器无法自动复制，请按 Ctrl+C 复制以下资料：',text);
  }
}
</script></body></html>
"""

@people_bp.route('/application-invite/new',methods=['GET','POST'])
def create_application_invite():
    ensure_people_admin_and_application_tables()
    error=''
    if request.method=='POST':
        name=clean(request.form.get('chinese_name'))
        phone=normalize_phone(request.form.get('phone'))
        branch='CHE'
        if not name or not phone:
            error='姓名和电话号码必须填写。'
        if not error:
            token=secrets.token_urlsafe(24)
            with get_conn() as conn:
                with conn.cursor() as cur:
                    cur.execute("""insert into people_applications(chinese_name,phone,preferred_branch,status,invite_token,invited_at,updated_by) values(%s,%s,%s,'draft',%s,current_timestamp,%s)""",(name,phone,branch,token,session.get('people_admin_name')))
            flash('✅ 已建立申请，请复制专属 WhatsApp 邀请给申请人。','ok')
            return redirect(url_for('people.applications'))
    return render_template_string(CREATE_INVITE_PAGE,error=error)

@people_bp.route('/applications')
def applications():
    ensure_people_admin_and_application_tables()
    rows=db_query("select id,chinese_name,phone,preferred_branch,status,invite_token,submitted_at,applicant_submitted_at from people_applications order by case status when 'draft' then 1 when 'submitted' then 2 when 'interview' then 3 when 'probation' then 4 when 'formal' then 5 else 6 end, submitted_at desc",fetchall=True) or []
    return render_template_string(APPLICATION_LIST_PAGE,rows=rows,labels=APPLICATION_STATUS_LABELS,format_phone=format_phone)

@people_bp.route('/application-photo/<int:application_id>')
def application_photo(application_id):
    r=db_query("select photo_data,photo_mimetype from people_applications where id=%s",(application_id,),fetchone=True)
    if not r or not r.get('photo_data'):abort(404)
    return Response(bytes(r['photo_data']),mimetype=r.get('photo_mimetype') or 'image/jpeg',headers={'Cache-Control':'private, no-store'})

@people_bp.route('/application/<int:application_id>',methods=['GET','POST'])
def application_detail(application_id):
    ensure_people_admin_and_application_tables(); ensure_people_profile_tables()
    r=db_query("select * from people_applications where id=%s",(application_id,),fetchone=True)
    if not r:abort(404)
    if request.method=='POST':
        action=clean(request.form.get('action'))
        interview_date=clean(request.form.get('interview_date')); interview_time=clean(request.form.get('interview_time')); interviewers=clean(request.form.get('interviewers')); interview_note=clean(request.form.get('interview_note')); probation_start_date=clean(request.form.get('probation_start_date')); volunteer_start_date=clean(request.form.get('volunteer_start_date'))
        status=r['status']
        if action=='interview':status='interview'
        elif action=='probation':status='probation'
        elif action=='reject':status='rejected'
        elif action=='formal':
            if clean(r.get('status'))=='formal':
                flash('这份申请已经正式加入义工，不会重复建立。','bad')
                return redirect(url_for('people.application_detail',application_id=application_id))
            branch='CHE'
            manual_vid=clean(request.form.get('manual_volunteer_id')).upper()
            existing_mid=clean(r.get('existing_member_id')).upper()
            member=None
            if existing_mid:
                member=db_query("select member_id,name,phone,pin,remark,branch from members where upper(member_id)=upper(%s)",(existing_mid,),fetchone=True)
                if not member:
                    flash(f'找不到月费编号 {existing_mid}，请先确认编号。','bad'); return redirect(url_for('people.application_detail',application_id=application_id))
            member_branch=norm_branch(member.get('branch')) if member else ''
            # 同分会的月费佛友沿用月费编号；不同分会绝不自动借用月费编号。
            if manual_vid:
                volunteer_id=manual_vid
            elif member and member_branch==branch:
                volunteer_id=clean(member.get('member_id')).upper()
            elif branch=='CHE':
                volunteer_id=next_volunteer_id()
            else:
                flash('这位申请者的月费编号不能直接作为 STW 义工编号。请在「指定义工编号」填写正确的 STW-xxx 编号后再正式加入。','bad')
                return redirect(url_for('people.application_detail',application_id=application_id))
            if branch=='STW' and not volunteer_id.startswith('STW-'):
                flash('STW 正式义工必须使用 STW-xxx 义工编号，请重新确认。','bad'); return redirect(url_for('people.application_detail',application_id=application_id))
            if volunteer_exists(volunteer_id):
                flash(f'义工编号 {volunteer_id} 已经存在，不能重复建立。','bad'); return redirect(url_for('people.application_detail',application_id=application_id))
            pin=(clean(member.get('pin')) if member else '') or default_pin(r.get('phone'))
            with get_conn() as conn:
                with conn.cursor() as cur:
                    cur.execute("""insert into volunteers(id,branch,name,status,phone,is_volunteer,is_member,member_id,pin,remark) values(%s,%s,%s,'在册',%s,true,%s,%s,%s,%s)""",(volunteer_id,branch,r['chinese_name'],normalize_phone(r['phone']),bool(member),member['member_id'] if member else None,pin,'由新义工申请转为正式义工'))
                    profile_data={
                        'name':clean(r.get('chinese_name')),'english_name':clean(r.get('english_name')),'ic_number':clean(r.get('ic_number')),'birth_date':'','phone':normalize_phone(r.get('phone')),'address':clean(r.get('address')),'marital_status':clean(r.get('marital_status')),'occupation':clean(r.get('occupation')),'education':clean(r.get('education')),'skills':clean(r.get('skills')),'emergency_name':clean(r.get('emergency_name')),'emergency_relation':clean(r.get('emergency_relation')),'emergency_phone':normalize_phone(r.get('emergency_phone')),'vegetarian_info':clean(r.get('vegetarian_info')),'cooks_meat':clean(r.get('cooks_meat')),'other_practice':clean(r.get('other_practice')),'xlfm_source':clean(r.get('xlfm_source')),'xlfm_years':clean(r.get('xlfm_years')),'disciple_status':clean(r.get('disciple_status')),'disciple_no':clean(r.get('disciple_no')),'dabei':clean(r.get('dabei')),'xin_jing':clean(r.get('xin_jing')),'li_fo':clean(r.get('li_fo')),'small_mantras':clean(r.get('small_mantras')),'xfz_weekly':clean(r.get('xfz_weekly')),'xfz_total':clean(r.get('xfz_total')),'original_join_date':volunteer_start_date,'note':('推荐人：'+clean(r.get('recommender'))+'；申请原因：'+clean(r.get('application_reason'))+'；义工经验：'+clean(r.get('volunteer_experience'))+'；希望岗位：'+clean(r.get('preferred_tasks'))+'；可参与时间：'+clean(r.get('available_time'))).strip('；')
                    }
                    cur.execute("""insert into people_profiles(volunteer_id,profile_json,photo_data,photo_mimetype,updated_at) values(%s,%s,%s,%s,current_timestamp) on conflict(volunteer_id) do update set profile_json=excluded.profile_json,photo_data=excluded.photo_data,photo_mimetype=excluded.photo_mimetype,updated_at=current_timestamp""",(volunteer_id,json.dumps(profile_data,ensure_ascii=False),r.get('photo_data'),r.get('photo_mimetype')))
                    cur.execute("update people_applications set status='formal',preferred_branch=%s,formal_volunteer_id=%s,volunteer_start_date=%s,updated_at=current_timestamp,updated_by=%s where id=%s",(branch,volunteer_id,volunteer_start_date,session.get('people_admin_name'),application_id))
            flash(f'✅ 已正式加入 {branch} 义工，义工编号：{volunteer_id}','ok')
            return redirect(url_for('people.application_detail',application_id=application_id))
        db_query("""update people_applications set status=%s,interview_date=%s,interview_time=%s,interviewers=%s,interview_note=%s,probation_start_date=%s,volunteer_start_date=%s,updated_at=current_timestamp,updated_by=%s where id=%s""",(status,interview_date,interview_time,interviewers,interview_note,probation_start_date,volunteer_start_date,session.get('people_admin_name'),application_id))
        flash('✅ 申请记录已更新。','ok')
        return redirect(url_for('people.application_detail',application_id=application_id))
    return render_template_string(APPLICATION_DETAIL_PAGE,r=r,fields=APPLICATION_FIELDS,status_labels=APPLICATION_STATUS_LABELS,format_phone=format_phone)
