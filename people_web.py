# people_web.py
# 人员管理中心 v1
# 直接管理 Supabase/PostgreSQL 的 members + volunteers
# 功能：总名单、搜索筛选、个人资料编辑、义工/月费状态切换、Excel 下载

from io import BytesIO
import re
import pandas as pd
from flask import Blueprint, request, redirect, url_for, flash, render_template_string, send_file
from psycopg2.extras import RealDictCursor

from db import db_query, get_conn

people_bp = Blueprint("people", __name__, url_prefix="/people")


def clean(v):
    return str(v or "").strip()


def digits(v):
    return "".join(ch for ch in clean(v) if ch.isdigit())


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


def load_people():
    members = db_query("""
        select member_id, name, english_name, phone, pin, branch,
               status, remark, member_status, ic_number, fee_start_month
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
            "name": clean(m.get("name")),
            "english_name": clean(m.get("english_name")),
            "phone": clean(m.get("phone")),
            "pin": clean(m.get("pin")),
            "ic_number": clean(m.get("ic_number")),
            "remark": clean(m.get("remark")),
            "is_member": True,
            "member_status": clean(m.get("member_status") or m.get("status") or "在供"),
            "fee_start_month": m.get("fee_start_month").strftime("%Y-%m") if m.get("fee_start_month") else "",
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
            p["is_volunteer"] = bool(v.get("is_volunteer")) if v.get("is_volunteer") is not None else True
            p["volunteer_status"] = clean(v.get("status") or "在册")
            # members 是个人资料主来源；空白时才从 volunteers 补。
            p["name"] = p["name"] or clean(v.get("name"))
            p["phone"] = p["phone"] or clean(v.get("phone"))
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
                "name": clean(v.get("name")),
                "english_name": "",
                "phone": clean(v.get("phone")),
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
        def match(p):
            hay = " ".join([
                p["display_id"], p["member_id"], p["volunteer_id"],
                p["name"], p["english_name"], p["phone"]
            ]).lower()
            return q in hay or (qdigits and qdigits in digits(p["phone"]))
        rows = [p for p in rows if match(p)]

    if branch in ("CHE", "STW"):
        rows = [p for p in rows if p["branch"] == branch]

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
.hero{background:linear-gradient(135deg,#6157e7,#8a6cf0);color:#fff;border-radius:22px;padding:25px 28px;margin-bottom:18px}
.hero h1{margin:0 0 7px;font-size:32px}.hero p{margin:0;opacity:.94}
.stats{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin:18px 0}
.stat{background:#fff;border-radius:16px;padding:16px;text-align:center;box-shadow:0 2px 10px #0000000d}
.stat b{display:block;font-size:27px;margin-top:5px}
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
@media(max-width:760px){.stats{grid-template-columns:1fr 1fr}.filters{grid-template-columns:1fr}.person-head{display:block}}
</style>
</head>
<body><div class="people-page">
<div class="hero"><h1>👥 人员管理中心</h1><p>统一管理义工与月费佛友资料</p></div>

{% with messages = get_flashed_messages(with_categories=true) %}
{% for cat,msg in messages %}<div class="alert {{ 'alert-success' if cat=='ok' else 'alert-danger' }}">{{ msg }}</div>{% endfor %}
{% endwith %}

<div class="stats">
 <div class="stat">总人数<b>{{ stats.total }}</b></div>
 <div class="stat">义工<b>{{ stats.volunteers }}</b></div>
 <div class="stat">月费佛友<b>{{ stats.members }}</b></div>
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
    {% if p.english_name %}{{ p.english_name }} · {% endif %}{{ p.branch }} · 📞 {{ p.phone or '-' }}
   </div>
   <div class="badges">
    {% if p.is_volunteer %}<span class="badge {{ 'on' if p.volunteer_status=='在册' else 'off' }}">👥 义工：{{ p.volunteer_status }}</span>{% else %}<span class="badge">非义工</span>{% endif %}
    {% if p.is_member %}<span class="badge {{ 'on' if p.member_status=='在供' else 'off' }}">💰 月费：{{ p.member_status }}</span>{% else %}<span class="badge">非月费</span>{% endif %}
   </div>
  </div>
  <a class="btn primary" href="{{ url_for('people.edit', person_key=p.key) }}">✏️ 编辑资料</a>
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
<div><label>分会</label><select class="input" name="branch"><option value="CHE" {% if p.branch=='CHE' %}selected{% endif %}>CHE</option><option value="STW" {% if p.branch=='STW' %}selected{% endif %}>STW</option></select></div>
<div><label>姓名</label><input class="input" name="name" value="{{ p.name }}" required></div>
<div><label>英文名</label><input class="input" name="english_name" value="{{ p.english_name }}"></div>
<div><label>电话号码</label><input class="input" name="phone" value="{{ p.phone }}"></div>
<div><label>PIN</label><input class="input" name="pin" value="{{ p.pin }}" inputmode="numeric"></div>
{% if p.is_member %}<div><label>身份证号码</label><input class="input" name="ic_number" value="{{ p.ic_number }}"></div>{% endif %}
{% if p.is_member %}<div><label>月费开始月份</label><input class="input" type="month" name="fee_start_month" value="{{ p.fee_start_month }}"><small>新佛友可设置未来月份；留空沿用旧规则。开始月份之前不计算迟付。</small></div>{% endif %}
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
<button class="btn {{ 'danger' if p.member_status=='在供' else 'success' }}" type="submit">{{ '停止月费' if p.member_status=='在供' else '恢复月费' }}</button>
</form>
{% endif %}
</div></div>
</div></body></html>
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
    all_rows = load_people()
    rows = filtered_people()
    stats = {
        "total": len(all_rows),
        "volunteers": sum(1 for p in all_rows if p["is_volunteer"] and p["volunteer_status"] == "在册"),
        "members": sum(1 for p in all_rows if p["is_member"] and p["member_status"] == "在供"),
        "both": sum(1 for p in all_rows if p["is_volunteer"] and p["volunteer_status"] == "在册"
                    and p["is_member"] and p["member_status"] == "在供"),
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
        phone = clean(request.form.get("phone"))
        pin = clean(request.form.get("pin"))
        branch = norm_branch(request.form.get("branch"))
        remark = clean(request.form.get("remark"))
        ic_number = clean(request.form.get("ic_number"))
        fee_start_month = clean(request.form.get("fee_start_month")) if p["is_member"] else ""
        if fee_start_month and not re.fullmatch(r"\d{4}-(0[1-9]|1[0-2])", fee_start_month):
            flash("月费开始月份格式错误，请使用 YYYY-MM。", "bad")
            return render_template_string(EDIT_PAGE, p=p)

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
                            branch=%s, remark=%s, ic_number=%s, fee_start_month=%s
                        where member_id=%s
                    """, (name, english_name, phone, pin, branch, remark,
                          ic_number, (fee_start_month + "-01") if fee_start_month else None, p["member_id"]))

                if p["is_volunteer"]:
                    cur.execute("""
                        update volunteers
                        set name=%s, phone=%s, pin=%s, branch=%s, remark=%s
                        where id=%s
                    """, (name, phone, pin, branch, remark, p["volunteer_id"]))

        flash(f"✅ {name} 的个人资料已更新。", "ok")
        return redirect(url_for("people.edit", person_key=person_key))

    return render_template_string(EDIT_PAGE, p=p)


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
            "分会": p["branch"],
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
