# schedule_web.py

import os
import pandas as pd
from openpyxl import load_workbook
from sqlalchemy import create_engine, text
from schedule_builder import run_schedule_for_date
from monthly_prebook_message import generate_monthly_prebook_message
from flask import Blueprint, request, session, redirect, url_for, render_template_string

schedule_bp = Blueprint("schedule", __name__)


DATABASE_URL = os.environ.get("DATABASE_URL")
engine = create_engine(DATABASE_URL, pool_pre_ping=True)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

MASTER_FILE = os.path.join(BASE_DIR, "master_volunteers.xlsx")
PREBOOK_FILE = os.path.join(BASE_DIR, "prebook_schedule.xlsx")

SCHEDULE_PIN = "1234"

TIME_OPTIONS = [
    "6:00am", "6:30am", "7:00am", "7:30am",
    "8:00am", "8:30am", "9:00am", "9:30am",
    "10:00am", "10:30am", "11:00am", "11:30am",
    "12:00pm", "12:30pm", "1:00pm", "1:30pm",
    "2:00pm", "2:30pm", "3:00pm", "3:30pm",
    "4:00pm", "4:30pm", "5:00pm", "5:30pm", "6:00pm"
]

ROLE_OPTIONS = ["值班", "卫生", "佛台", "供台"]

schedule_records = []

def find_name_by_id(vol_id):
    raw_id = str(vol_id or "").strip()

    if not raw_id:
        return raw_id

    ids = [raw_id]

    if "-" in raw_id:
        branch, num = raw_id.split("-", 1)
        branch = branch.strip().upper()
        num = num.strip()

        ids.append(num)

        if branch == "STW" and num.isdigit():
            ids.append("0" + num)

    else:
        ids.append(f"CHE-{raw_id}")

        if raw_id.startswith("0") and raw_id[1:].isdigit():
            ids.append(f"STW-{raw_id[1:]}")

    ids = list(dict.fromkeys(ids))

    try:
        placeholders = ", ".join([f":id{i}" for i in range(len(ids))])
        params = {f"id{i}": v for i, v in enumerate(ids)}

        sql = text(f"""
            select id, name, status, branch
            from volunteers
            where id in ({placeholders})
            limit 1
        """)

        with engine.connect() as conn:
            row = conn.execute(sql, params).mappings().first()

        if not row:
            return raw_id

        return str(row["name"]).strip()

    except Exception as e:
        print("find_name_by_id error:", e)
        return raw_id

def get_default_time_by_role(role, start_time, end_time):
    role = str(role).strip()

    # 值班才用负责人选择的时间
    if role == "值班":
        return start_time, end_time

    # 第一版先固定普通日时间
    if role in ["卫生", "佛台"]:
        return "8:00am", "10:00am"

    if role == "供台":
        return "6:00am", "8:00am"

    return start_time, end_time

def save_prebook_record(record):
    new_df = pd.DataFrame([record])

    if os.path.exists(PREBOOK_FILE):
        try:
            old_df = pd.read_excel(PREBOOK_FILE, sheet_name="预报名")
            old_df.columns = old_df.columns.astype(str).str.strip()
            df = pd.concat([old_df, new_df], ignore_index=True)
        except Exception:
            df = new_df
    else:
        df = new_df

    df = df.drop_duplicates(
        subset=["日期", "姓名", "岗位", "开始时间", "结束时间"],
        keep="first"
    )

    with pd.ExcelWriter(PREBOOK_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="预报名", index=False)

@schedule_bp.route("/schedule/generate_day", methods=["POST"])
def schedule_generate_day():
    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule"))

    date = request.form.get("date", "").strip()

    try:
        output = run_schedule_for_date(date)
    except Exception as e:
        output = f"❌ 生成失败：{e}"

    return render_template_string(DAY_OUTPUT_HTML, output=output)


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

    name = find_name_by_id(vol_id)

    for day in days:
        for role in roles:
            date_text = f"{month}-{int(day):02d}"

            job_start, job_end = get_default_time_by_role(role, start_time, end_time)

            record = {
                "日期": date_text,
                "编号": vol_id,
                "姓名": name,
                "岗位": role,
                "开始时间": job_start,
                "结束时间": job_end,
                "备注": "网页录入",
            }

            schedule_records.append(record)
            save_prebook_record(record)

    return redirect(url_for("schedule.schedule"))

@schedule_bp.route("/schedule/monthly_prebook", methods=["POST"])
def schedule_monthly_prebook():
    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule"))

    year = request.form.get("year", "").strip()
    month = request.form.get("month", "").strip()

    try:
        output = generate_monthly_prebook_message(int(year), int(month))
    except Exception as e:
        output = f"❌ 生成失败：{e}"

    return render_template_string(MONTHLY_PREBOOK_HTML, output=output)


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

<hr>
<h2>📢 生成月预报名表</h2>

<form method="post" action="/schedule/monthly_prebook">
    年份：
    <input name="year" value="2026" style="width:120px;" required>

    月份：
    <select name="month">
        {% for m in range(1, 13) %}
        <option value="{{ m }}">{{ m }}月</option>
        {% endfor %}
    </select>

    <button type="submit">📢 生成预报名表</button>
</form>
<hr>

<hr>
<h2>📅 生成当天排班</h2>

<form method="post" action="/schedule/generate_day">
    日期：
    <input type="date" name="date" required>

    <button type="submit">⚡ 生成排班</button>
</form>
<hr>

</form>

<form method="post" action="/schedule/clear">
<button type="submit">🗑 清空名单</button>
</form>

<h2>已加入报名名单</h2>

<table>
<tr>
    <th>日期</th>
    <th>编号</th>
    <th>姓名</th>
    <th>岗位</th>
    <th>时间</th>
</tr>
{% for r in records %}
<tr>
    <td>{{ r["日期"] }}</td>
    <td>{{ r["编号"] }}</td>
    <td>{{ r["姓名"] }}</td>
    <td>{{ r["岗位"] }}</td>
    <td>{{ r["开始时间"] }} ~ {{ r["结束时间"] }}</td>
</tr>
{% endfor %}
</table>

</div>
</body>
</html>
"""

MONTHLY_PREBOOK_HTML = """
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>月预报名表</title>
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
textarea {
    width: 100%;
    height: 650px;
    font-size: 20px;
    padding: 15px;
    box-sizing: border-box;
}
button, a {
    font-size: 22px;
    padding: 10px 18px;
    margin: 8px;
}
</style>
</head>
<body>
<div class="box">
<h1>📢 月预报名表</h1>

<a href="/schedule">⬅ 返回排班系统</a>

<br><br>

<textarea readonly>{{ output }}</textarea>

</div>
</body>
</html>
"""

DAY_OUTPUT_HTML = """
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>排班结果</title>
<style>
body {
    font-family: "Microsoft YaHei";
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
textarea {
    width: 100%;
    height: 650px;
    font-size: 20px;
}
</style>
</head>
<body>
<div class="box">
<h1>📅 排班结果</h1>

<a href="/schedule">⬅ 返回</a>

<br><br>

<textarea readonly>{{ output }}</textarea>

</div>
</body>
</html>
"""