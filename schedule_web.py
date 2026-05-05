# schedule_web.py

from flask import Blueprint, request, session, redirect, url_for, render_template_string

schedule_bp = Blueprint("schedule", __name__)

SCHEDULE_PIN = "1234"

TIME_OPTIONS = [
    "8:00am", "8:30am", "9:00am", "9:30am",
    "10:00am", "10:30am", "11:00am", "11:30am",
    "12:00pm", "12:30pm", "1:00pm", "1:30pm",
    "2:00pm", "2:30pm", "3:00pm", "3:30pm",
    "4:00pm", "4:30pm", "5:00pm", "5:30pm", "6:00pm"
]

ROLE_OPTIONS = ["值班", "卫生", "佛台", "供台"]

schedule_records = []


@schedule_bp.route("/schedule", methods=["GET", "POST"])
def schedule():
    if not session.get("schedule_login"):
        if request.method == "POST":
            pin = request.form.get("pin", "").strip()
            if pin == SCHEDULE_PIN:
                session["schedule_login"] = True
                return redirect(url_for("schedule.schedule"))

            return "❌ PIN 错误<br><a href='/schedule'>返回</a>"

        return """
        <h2>📅 负责人排班系统</h2>
        <form method="post">
            <input type="password" name="pin" placeholder="请输入负责人PIN" style="font-size:24px;padding:10px;">
            <button type="submit" style="font-size:24px;padding:10px;">进入</button>
        </form>
        """

    return render_template_string(SCHEDULE_HTML, times=TIME_OPTIONS, roles=ROLE_OPTIONS, records=schedule_records)


@schedule_bp.route("/schedule/add", methods=["POST"])
def schedule_add():
    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule"))

    vol_id = request.form.get("vol_id", "").strip()
    month = request.form.get("month", "").strip()
    days = request.form.getlist("days")
    roles = request.form.getlist("roles")
    start_time = request.form.get("start_time", "")
    end_time = request.form.get("end_time", "")

    for day in days:
        for role in roles:
            schedule_records.append({
                "编号": vol_id,
                "月份": month,
                "日期": day,
                "岗位": role,
                "开始时间": start_time,
                "结束时间": end_time,
            })

    return redirect(url_for("schedule.schedule"))


@schedule_bp.route("/schedule/clear", methods=["POST"])
def schedule_clear():
    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule"))

    schedule_records.clear()
    return redirect(url_for("schedule.schedule"))


SCHEDULE_HTML = """
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>负责人排班系统</title>
<style>
body {
    font-family: "Microsoft YaHei", Arial;
    background: #f5f5f5;
    padding: 20px;
}
.box {
    background: white;
    max-width: 900px;
    margin: auto;
    padding: 25px;
    border-radius: 15px;
}
input, select, button {
    font-size: 22px;
    padding: 10px;
    margin: 6px;
}
.day-box label, .role-box label {
    display: inline-block;
    font-size: 22px;
    padding: 8px 12px;
    margin: 5px;
    background: #eee;
    border-radius: 8px;
}
table {
    width: 100%;
    border-collapse: collapse;
    margin-top: 20px;
    font-size: 20px;
}
th, td {
    border: 1px solid #aaa;
    padding: 8px;
    text-align: center;
}
th {
    background: #d9ead3;
}
</style>
</head>
<body>
<div class="box">

<h1>📅 负责人排班系统</h1>

<form method="post" action="/schedule/add">

<h2>1. 输入义工编号</h2>
<input name="vol_id" placeholder="例如 208 / 0160 / 803" required>

<h2>2. 选择月份</h2>
<input name="month" placeholder="例如 2026-05" required>

<h2>3. 多选日期</h2>
<div class="day-box">
{% for d in range(1, 32) %}
<label>
    <input type="checkbox" name="days" value="{{ d }}"> {{ d }}
</label>
{% endfor %}
</div>

<h2>4. 选择岗位</h2>
<div class="role-box">
{% for role in roles %}
<label>
    <input type="checkbox" name="roles" value="{{ role }}"> {{ role }}
</label>
{% endfor %}
</div>

<h2>5. 选择时间</h2>
开始：
<select name="start_time">
{% for t in times %}
<option value="{{ t }}">{{ t }}</option>
{% endfor %}
</select>

结束：
<select name="end_time">
{% for t in times %}
<option value="{{ t }}">{{ t }}</option>
{% endfor %}
</select>

<br><br>
<button type="submit">➕ 加入报名</button>

</form>

<form method="post" action="/schedule/clear">
<button type="submit">🗑 清空名单</button>
</form>

<h2>已加入报名名单</h2>

<table>
<tr>
    <th>编号</th>
    <th>月份</th>
    <th>日期</th>
    <th>岗位</th>
    <th>时间</th>
</tr>
{% for r in records %}
<tr>
    <td>{{ r["编号"] }}</td>
    <td>{{ r["月份"] }}</td>
    <td>{{ r["日期"] }}</td>
    <td>{{ r["岗位"] }}</td>
    <td>{{ r["开始时间"] }} ~ {{ r["结束时间"] }}</td>
</tr>
{% endfor %}
</table>

</div>
</body>
</html>
"""