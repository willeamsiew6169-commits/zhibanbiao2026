# admin_web.py

import io
import pandas as pd


from datetime import datetime
from io import BytesIO
from db import db_query
from utils import get_text
from openpyxl import Workbook
from utils import now_date_str, get_text
from utils import get_display_today_code
from psycopg2.extras import RealDictCursor
from attendance_service import auto_signout_unfinished_today
from utils import MY_TZ, get_today_code, now_date_str, calc_hours
from flask import Blueprint, request, redirect, url_for, render_template_string, flash, send_file, session
from attendance_service import (
    role_label,
    get_today_open_records,
    get_today_records,
    find_volunteer,
    sign_in,
    sign_out,
)

admin_bp = Blueprint("admin", __name__)

# 管理员 PIN：你可以改成自己的
ADMIN_PIN = "8888"

@admin_bp.route("/admin_add_record", methods=["GET", "POST"])
def admin_add_record():
    from flask import request, redirect, url_for, render_template_string

    pin = request.args.get("pin") or request.form.get("pin")

    if pin != ADMIN_PIN:
        return "无权限"

    if request.method == "POST":
        date = request.form.get("date")
        vid = request.form.get("id")
        name = request.form.get("name")
        role = request.form.get("role")
        start = request.form.get("start")
        end = request.form.get("end")
        remark = request.form.get("remark")

        hours = ""
        if start and end:
            hours = calc_hours(start, end)

        db_query("""
            insert into attendance
            (date, volunteer_id, name, role, start_time, end_time, hours, remark)
            values (%s,%s,%s,%s,%s,%s,%s,%s)
        """, (date, vid, name, role, start, end, hours, remark))

        return redirect(url_for("admin.admin_add_record", pin=pin))

    return render_template_string("""
    <h1>🛠 补录签到</h1>

    <form method="post">
    <input type="hidden" name="pin" value="{{ pin }}">

    日期：
    <input type="date" name="date" value="{{ today }}" required><br><br>

    选择义工：
    <select name="id" id="vol_id" onchange="fillName()" required>
        <option value="">--请选择--</option>
        {% for v in volunteers %}
        <option value="{{ v.id }}">{{ v.name }} ({{ v.id }})</option>
        {% endfor %}
    </select><br><br>

    姓名：
    <input type="text" name="name" id="vol_name" required><br><br>

    岗位：
    <input type="text" name="role"><br><br>

    开始时间：
    <input type="text" name="start" value="{{ now }}"><br><br>

    结束时间：
    <input type="text" name="end"><br><br>

    备注：
    <input type="text" name="remark"><br><br>

    <button style="font-size:20px;">➕ 添加记录</button>
    </form>

    <script>
    function fillName() {
    const sel = document.getElementById("vol_id");
    const text = sel.options[sel.selectedIndex].text;
    const name = text.split(" (")[0];
    document.getElementById("vol_name").value = name;
    }
    </script>

    <br>
    <a href="/admin_report?pin={{ pin }}">⬅ 返回</a>
    """, 
    pin=pin,
    today=now_date_str(),
    now=datetime.now(MY_TZ).strftime("%I:%M%p").lower(),
    volunteers=db_query("select id, name from volunteers where status='在册'", fetchall=True)
    )

@admin_bp.route("/admin_today_open")
def admin_today_open():

    return render_template_string("""
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">

<style>
body {
    font-family:"Microsoft YaHei", Arial;
    background:#f4f6f8;
    margin:0;
    padding:18px;
    font-size:24px;
}

.wrap {
    max-width:800px;
    margin:auto;
}

.card {
    background:white;
    border-radius:20px;
    padding:22px;
    margin-bottom:18px;
    box-shadow:0 4px 14px rgba(0,0,0,.08);
}

h1 {
    font-size:34px;
}

.name {
    font-size:32px;
    font-weight:bold;
    margin-bottom:12px;
}

.info {
    line-height:1.8;
    font-size:24px;
}

.btn-out {
    width:100%;
    background:#dc3545;
    color:white;
    border:0;
    border-radius:16px;
    padding:18px;
    font-size:30px;
    font-weight:bold;
    margin-top:18px;
}

.btn-edit {
    display:block;
    text-align:center;
    background:#0d6efd;
    color:white;
    text-decoration:none;
    border-radius:16px;
    padding:16px;
    font-size:26px;
    font-weight:bold;
    margin-top:12px;
}

.back {
    display:block;
    text-align:center;
    background:#6c757d;
    color:white;
    text-decoration:none;
    border-radius:16px;
    padding:18px;
    font-size:28px;
    font-weight:bold;
    margin-top:20px;
}

.empty {
    background:#d1e7dd;
    color:#0f5132;
    border-radius:18px;
    padding:24px;
    font-size:28px;
    font-weight:bold;
    text-align:center;
}
</style>
</head>

<body>
<div class="wrap">

<h1>🔴 今日进行中（未签退）</h1>

{% if open_records %}

    {% for r in open_records %}
    <div class="card">

        <div class="name">
            {{ r.get('姓名','') }}
        </div>

        <div class="info">
            岗位：<b>{{ role_label(r.get('岗位','')) }}</b><br>
            签到时间：<b>{{ r.get('开始时间','') }}</b><br>

            {% if r.get('编号') %}
                义工编号：{{ r.get('编号','') }}<br>
            {% endif %}

            {% if r.get('card_no') %}
                卡号：{{ r.get('card_no') }}<br>
            {% endif %}
        </div>

        <form method="post"
              action="{{ url_for('do_sign_out') }}"
              onsubmit="return askSignOutPin(this);">

            <input type="hidden" name="row_number" value="{{ r.get('_row') }}">
            <input type="hidden" name="pin" value="">

            <button class="btn-out" type="submit">
                ⛔ 帮他签退
            </button>
        </form>

        <a class="btn-edit"
           href="{{ url_for('edit_page', row_number=r.get('_row')) }}">
            ✏ 修改记录
        </a>

    </div>
    {% endfor %}

{% else %}

    <div class="empty">
        ✅ 现在没有未签退的义工
    </div>

{% endif %}

<a class="back" href="/admin_report?pin={{ pin }}">
    ⬅ 返回管理员
</a>

</div>

<script>
function askSignOutPin(form) {
    const pin = prompt("请输入管理员 PIN");
    if (pin === null) return false;

    if (!pin.trim()) {
        alert("PIN 不能为空");
        return false;
    }

    form.querySelector('input[name="pin"]').value = pin.trim();
    return true;
}
</script>

</body>
</html>
""",
    pin=request.args.get("pin",""),
    open_records=get_today_open_records(),
    role_label=role_label,
    t=get_text())

@admin_bp.route("/admin/auto_signout")
def admin_auto_signout():
    if not session.get("schedule_login"):
        return redirect(url_for("index"))

    result = auto_signout_unfinished_today()

    flash(result["msg"], "ok" if result["ok"] else "bad")
    return redirect(url_for("index"))

@admin_bp.route("/admin_today_records")
def admin_today_records():

    return render_template_string("""
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">

<style>
body {
    font-family:"Microsoft YaHei", Arial;
    background:#f4f6f8;
    margin:0;
    padding:18px;
    font-size:24px;
}

.wrap {
    max-width:800px;
    margin:auto;
}

.card {
    background:white;
    border-radius:20px;
    padding:22px;
    margin-bottom:18px;
    box-shadow:0 4px 14px rgba(0,0,0,.08);
}

h1 {
    font-size:34px;
}

.name {
    font-size:32px;
    font-weight:bold;
    margin-bottom:12px;
}

.info {
    line-height:1.8;
    font-size:24px;
}

.time {
    font-size:28px;
    font-weight:bold;
    color:#0d6efd;
}

.btn-edit {
    display:block;
    text-align:center;
    background:#0d6efd;
    color:white;
    text-decoration:none;
    border-radius:16px;
    padding:16px;
    font-size:26px;
    font-weight:bold;
    margin-top:14px;
}

.back {
    display:block;
    text-align:center;
    background:#6c757d;
    color:white;
    text-decoration:none;
    border-radius:16px;
    padding:18px;
    font-size:28px;
    font-weight:bold;
    margin-top:20px;
}

.empty {
    background:#d1e7dd;
    color:#0f5132;
    border-radius:18px;
    padding:24px;
    font-size:28px;
    font-weight:bold;
    text-align:center;
}
</style>
</head>

<body>
<div class="wrap">

<h1>📋 今日签到记录</h1>

{% if today_records %}

    {% for r in today_records %}
    <div class="card">

        <div class="name">
            {{ r.get('姓名','') }}
        </div>

        <div class="info">
            岗位：<b>{{ role_label(r.get('岗位','')) }}</b><br>

            时间：<br>
            <span class="time">
                {{ r.get('开始时间','') }}
                {% if r.get('结束时间') %}
                    ~ {{ r.get('结束时间','') }}
                {% else %}
                    ~ 未签退
                {% endif %}
            </span><br>

            时数：<b>{{ r.get('时数','') or '-' }}</b><br>

            {% if r.get('card_no') %}
                卡号：{{ r.get('card_no') }}<br>
            {% endif %}
        </div>

        <a class="btn-edit"
           href="{{ url_for('edit_page', row_number=r.get('_row')) }}">
            ✏ 修改记录
        </a>

    </div>
    {% endfor %}

{% else %}

    <div class="empty">
        今天还没有签到记录
    </div>

{% endif %}

<a class="back" href="/admin_report?pin={{ pin }}">
    ⬅ 返回管理员
</a>

</div>
</body>
</html>
""",
    pin=request.args.get("pin",""),
    today_records=get_today_records(limit=200),
    role_label=role_label,
    t=get_text())

@admin_bp.route("/download_data")
def download_data():

    try:
        old_att_rows = db_query("""
            select
                date as "日期",
                volunteer_id as "编号",
                name as "姓名",
                signup as "报名",
                signin as "签到",
                role as "岗位",
                start_time as "开始时间",
                end_time as "结束时间",
                hours as "时数",
                remark as "备注"
            from attendance
            order by date, name, start_time
        """, fetchall=True)

        att_logs = db_query("""
            select
                attendance_date as "日期",
                volunteer_id as "编号",
                name as "姓名",
                actual_role as "岗位",
                actual_place as "地点",
                to_char(check_in_time, 'HH12:MIam') as "签到时间",
                to_char(check_out_time, 'HH12:MIam') as "签退时间",
                case when walk_in then '是' else '' end as "现场签到",
                remarks as "备注"
            from volunteer_attendance_logs
            order by attendance_date, name, check_in_time
        """, fetchall=True)

        vol_rows = db_query("""
            select
                id as "编号",
                name as "姓名",
                status as "状态",
                phone as "电话号码",
                pin as "PIN"
            from volunteers
            order by id
        """, fetchall=True)

        reading_rows = db_query("""
            select
                date as "日期",
                name as "姓名",
                identity as "身份",
                topic as "主题",
                session as "场次",
                time as "时间"
            from reading
            order by date, time
        """, fetchall=True)

        old_df = pd.DataFrame(old_att_rows)
        logs_df = pd.DataFrame(att_logs)

        if not logs_df.empty:
            new_att_df = pd.DataFrame({
                "日期": logs_df.get("日期", ""),
                "编号": logs_df.get("编号", ""),
                "姓名": logs_df.get("姓名", ""),
                "报名": logs_df["现场签到"].apply(lambda x: 0 if str(x).strip() == "是" else 1),
                "签到": 1,
                "岗位": logs_df.get("岗位", ""),
                "开始时间": logs_df.get("签到时间", ""),
                "结束时间": logs_df.get("签退时间", ""),
                "时数": "",
                "备注": logs_df.get("备注", ""),
            })
        else:
            new_att_df = pd.DataFrame(columns=[
                "日期", "编号", "姓名", "报名", "签到",
                "岗位", "开始时间", "结束时间", "时数", "备注"
            ])

        attendance_df = pd.concat(
            [old_df, new_att_df],
            ignore_index=True
        )

        if not attendance_df.empty:
            attendance_df = attendance_df.drop_duplicates(
                subset=["日期", "编号", "姓名", "岗位", "开始时间"],
                keep="last"
            )

            attendance_df = attendance_df.sort_values(
                by=["日期", "姓名", "开始时间"],
                ascending=[True, True, True]
            )

        output = io.BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:

            attendance_df.to_excel(
                writer,
                index=False,
                sheet_name="attendance"
            )

            logs_df.to_excel(
                writer,
                index=False,
                sheet_name="att_logs"
            )

            pd.DataFrame(vol_rows).to_excel(
                writer,
                index=False,
                sheet_name="volunteers"
            )

            pd.DataFrame(reading_rows).to_excel(
                writer,
                index=False,
                sheet_name="reading"
            )

        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name="zhibanbiao_data.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        return f"下载失败：{e}"
    
@admin_bp.route("/admin_edit_record/<int:record_id>", methods=["GET", "POST"])
def admin_edit_record(record_id):
    pin = request.args.get("pin") or request.form.get("pin")

    if pin != ADMIN_PIN:
        return "无权限"

    row = db_query("""
        select *
        from attendance
        where id = %s
    """, (record_id,), fetchone=True)

    if not row:
        return "找不到记录"

    if request.method == "POST":
        role = request.form.get("role", "").strip()
        start_time = request.form.get("start_time", "").strip()
        end_time = request.form.get("end_time", "").strip()
        remark = request.form.get("remark", "").strip()

        hours = None
        if start_time and end_time:
            hours = calc_hours(start_time, end_time)

        db_query("""
            update attendance
            set role = %s,
                start_time = %s,
                end_time = %s,
                hours = %s,
                remark = %s
            where id = %s
        """, (role, start_time, end_time, hours, remark, record_id))

        return redirect(url_for("admin.admin_records", pin=pin))

    return render_template_string("""
<h1>✏️ 修改记录</h1>

<form method="post">
  <input type="hidden" name="pin" value="{{ pin }}">

  <p>姓名：<b>{{ row.name }}</b></p>
  <p>日期：{{ row.date }}</p>

  岗位：<br>
  <input name="role" value="{{ row.role or '' }}" style="font-size:22px;"><br><br>

  开始时间：<br>
  <input name="start_time" value="{{ row.start_time or '' }}" style="font-size:22px;"><br><br>

  结束时间：<br>
  <input name="end_time" value="{{ row.end_time or '' }}" style="font-size:22px;"><br><br>

  备注：<br>
  <input name="remark" value="{{ row.remark or '' }}" style="font-size:22px;"><br><br>

  <button style="font-size:22px;">保存修改</button>
</form>

<br>
<a href="/admin_records?pin={{ pin }}">⬅ 返回</a>
""", row=row, pin=pin)

@admin_bp.route("/admin_delete_record/<int:record_id>")
def admin_delete_record(record_id):
    pin = request.args.get("pin", "")

    if pin != ADMIN_PIN:
        return "无权限"

    db_query("""
        delete from attendance
        where id = %s
          and date = %s
    """, (record_id, now_date_str()))

    return redirect(url_for("admin.admin_records", pin=pin))
    
@admin_bp.route("/admin_records")
def admin_records():
    pin = request.args.get("pin", "")

    if pin != ADMIN_PIN:
        return "无权限"

    rows = db_query("""
        select *
        from attendance
        where date = %s
        order by id desc
    """, (now_date_str(),), fetchall=True)

    return render_template_string("""
<h1>✏️ 今日记录管理</h1>

<a href="/admin_report?pin={{ pin }}">⬅ 返回管理员</a>
<br><br>

<table border="1" cellpadding="8" style="border-collapse:collapse;font-size:18px;">
<tr>
  <th>姓名</th>
  <th>岗位</th>
  <th>开始</th>
  <th>结束</th>
  <th>时数</th>
  <th>操作</th>
</tr>

{% for r in rows %}
<tr>
  <td>{{ r.name }}</td>
  <td>{{ r.role }}</td>
  <td>{{ r.start_time }}</td>
  <td>{{ r.end_time }}</td>
  <td>{{ r.hours }}</td>
  <td>
    <a href="/admin_edit_record/{{ r.id }}?pin={{ pin }}">修改</a>
    |
    <a href="/admin_delete_record/{{ r.id }}?pin={{ pin }}"
       onclick="return confirm('确定删除这笔记录吗？');">
       删除
    </a>
  </td>
</tr>
{% endfor %}
</table>
""", rows=rows, pin=pin)


@admin_bp.route("/admin_report", methods=["GET", "POST"])
def admin_report():
    pin = (
        request.form.get("admin_pin")
        or request.args.get("pin")
        or ""
    ).strip()

    print("输入 pin =", repr(pin))
    print("系统 ADMIN_PIN =", repr(ADMIN_PIN))

    t = get_text()

    if pin != ADMIN_PIN:
        flash(t["admin_pin_wrong"], "bad")
        return redirect(url_for("index"))

    code = get_display_today_code()

    return f"""
    <!doctype html>
    <html lang="zh">
    <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>管理员工具</title>

    <style>
    body {{
        margin:0;
        background:#f4f6f8;
        font-family:"Microsoft YaHei", Arial, sans-serif;
        font-size:22px;
    }}

    .wrap {{
        max-width:760px;
        margin:0 auto;
        padding:24px;
    }}

    .card {{
        background:white;
        border-radius:20px;
        padding:24px;
        margin-bottom:18px;
        box-shadow:0 4px 14px rgba(0,0,0,.08);
    }}

    h1 {{
        font-size:32px;
        margin:0 0 18px;
    }}

    h2 {{
        font-size:24px;
        margin:0 0 14px;
    }}

    .code {{
        font-size:64px;
        font-weight:bold;
        color:#dc3545;
        text-align:center;
        padding:20px;
        background:#fff3cd;
        border-radius:18px;
    }}

    .btn {{
        display:block;
        width:100%;
        box-sizing:border-box;
        text-align:center;
        text-decoration:none;
        font-size:26px;
        font-weight:bold;
        padding:18px;
        border-radius:16px;
        margin-top:14px;
        color:white;
    }}

    .blue {{ background:#0d6efd; }}
    .green {{ background:#198754; }}
    .orange {{ background:#fd7e14; }}
    .red {{ background:#dc3545; }}
    .purple {{ background:#6f42c1; }}
    .gray {{ background:#6c757d; }}

    .note {{
        color:#666;
        font-size:18px;
        margin-top:10px;
    }}
    </style>
    </head>

    <body>
    <div class="wrap">

    <div class="card">
        <h1>🔐 {t["admin_title"]}</h1>
    </div>

    <div class="card">
        <h2>{t["today_code_big"]}</h2>
        <div class="code">{code}</div>
        <div class="note">⚠ 请只写在观音堂现场，不要发群</div>
    </div>

    <div class="card">
        <h2>📋 今日签到管理</h2>

        <a class="btn red" href="/admin_today_open?pin={pin}">
            ⛔ 今日进行中（未签退）
        </a>

        <a class="btn blue" href="/admin_today_records?pin={pin}">
            📋 查看今日记录
        </a>
    </div>

    <div class="card">
        <h2>🛠 管理工具</h2>

        <a class="btn green" href="/download_data">
            📥 下载签到日志 Excel
        </a>

        <a class="btn orange" href="/admin_add_record?pin={pin}">
            ✍️ 补录签到
        </a>

        <a class="btn purple" href="/admin_records?pin={pin}">
            📝 修改 / 删除今日记录
        </a>
    </div>

    <div class="card">
        <a class="btn gray" href="/">
            ⬅ 返回首页
        </a>
    </div>

    </div>
    </body>
    </html>
    """

ADMIN_HOME_HTML = """
<!doctype html>
<html lang="zh">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">

<link rel="manifest" href="/manifest.json">

<title>观音堂管理员入口</title>

<style>

body{
    margin:0;
    padding:0;
    background:#f5efe3;
    font-family:"Microsoft YaHei";
    text-align:center;
}

.container{
    max-width:500px;
    margin:auto;
    padding:30px 20px;
}

h1{
    color:#8b5a2b;
}

.card{
    background:white;
    border-radius:20px;
    padding:25px;
    margin-bottom:25px;
    box-shadow:0 4px 12px rgba(0,0,0,0.08);
}

.btn{
    display:block;
    background:#b67b3d;
    color:white;
    text-decoration:none;
    padding:16px;
    border-radius:14px;
    font-size:18px;
    font-weight:bold;
}

</style>
</head>

<body>

<div class="container">

<h1>🙏 观音堂管理员入口</h1>

<div class="card">
    <h2>📅 今日签到码</h2>

    <div style="
        font-size:48px;
        font-weight:bold;
        text-align:center;
        padding:20px;
        background:#fff3cd;
        border-radius:16px;
        color:#856404;
    ">
        {{ today_code }}
    </div>
</div>

<div class="card">
<h2>📋 义工报名系统</h2>
<a class="btn" href="/volunteer">
进入系统
</a>
</div>

<div class="card">
<h2>💰 月费管理员系统</h2>
<a class="btn" href="/member">
进入系统
</a>
</div>

<div class="card">
<h2>✅ 义工签到系统</h2>
<a class="btn" href="/">
进入系统
</a>
</div>

<div class="card">
<h2>📥 数据下载</h2>

<a class="btn" href="{{ url_for('admin.download_data') }}">
下载旧签到数据 Excel
</a>

<br>

<a class="btn" href="{{ url_for('admin.download_att_logs') }}">
下载新签到日志 att_logs
</a>

</div>

</div>

</body>
</html>
"""

@admin_bp.route("/admin-home")
def admin_home():
    return render_template_string(
        ADMIN_HOME_HTML,
        today_code=get_today_code()
    )

@admin_bp.route("/download_att_logs")
def download_att_logs():

    try:
        att_logs = db_query("""
            select
                attendance_date as "日期",
                volunteer_id as "编号",
                name as "姓名",
                actual_role as "岗位",
                actual_place as "地点",
                to_char(check_in_time, 'HH12:MIam') as "签到时间",
                to_char(check_out_time, 'HH12:MIam') as "签退时间",
                case when walk_in then '是' else '' end as "临时报到",
                remarks as "备注"
            from volunteer_attendance_logs
            order by attendance_date, name, check_in_time
        """, fetchall=True)

        output = io.BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            pd.DataFrame(att_logs).to_excel(
                writer,
                index=False,
                sheet_name="att_logs"
            )

        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name="att_logs.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        return f"下载签到日志失败：{e}"