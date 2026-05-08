# admin_web.py

import io
import pandas as pd

from flask import Blueprint, request, redirect, url_for, render_template_string, flash, send_file
from datetime import datetime

from db import db_query
from utils import get_text
from utils import now_date_str
from utils import get_display_today_code
from utils import MY_TZ, get_today_code, now_date_str, calc_hours


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

@admin_bp.route("/download_data")
def download_data():
    
    try:
        att_rows = db_query("""
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
            order by date, start_time
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

        output = io.BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            pd.DataFrame(att_rows).to_excel(writer, index=False, sheet_name="attendance")
            pd.DataFrame(vol_rows).to_excel(writer, index=False, sheet_name="volunteers")
            pd.DataFrame(reading_rows).to_excel(writer, index=False, sheet_name="reading")

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


@admin_bp.route("/admin_report", methods=["POST"])
def admin_report():
    pin = str(request.form.get("admin_pin", "")).strip()

    t = get_text()  # ⭐ 加这一行

    if pin != ADMIN_PIN:
        flash(t["admin_pin_wrong"], "bad")
        return redirect(url_for("index"))

    code = get_display_today_code()

    return f"""
<h1>{t["admin_title"]}</h1>

<h2>{t["today_code_big"]}</h2>
<div style="font-size:60px;font-weight:bold;color:#dc3545;">
    {code}
</div>

<p>{t["today_code_warning"]}</p>

<a href="/download_data" style="display:block;margin-top:12px;font-size:24px;">
    {t["download_data"]}
</a>

<a href="/admin_add_record?pin={pin}" style="display:block;margin-top:12px;font-size:24px;">
  {t["admin_add_record"]}
</a>

<a href="/admin_records?pin={pin}" style="display:block;margin-top:12px;font-size:24px;">
  {t["admin_records"]}
</a>

<br>
<a href="/" style="font-size:20px;">⬅ {t["back_home"]}</a>
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
<h2>📋 值班表生成系统</h2>
<a class="btn" href="/schedule">
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

<a class="btn" href="/admin/download_data">📥 下载签到数据 Excel</a>

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