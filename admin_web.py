# admin_web.py

from flask import Blueprint, request, redirect, url_for, render_template_string
from datetime import datetime

from db import db_query
from utils import MY_TZ, get_today_code, now_date_str, calc_hours

admin_bp = Blueprint("admin", __name__)

# 管理员 PIN：你可以改成自己的
ADMIN_PIN = "8888"

@app.route("/admin_add_record", methods=["GET", "POST"])
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

        return redirect(url_for("admin_add_record", pin=pin))

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