# schedule_web.py

import os
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime, timedelta
from sqlalchemy import create_engine, text
from schedule_builder import run_schedule_for_date
from monthly_prebook_message import generate_monthly_prebook_message
from flask import Blueprint, request, session, redirect, url_for, render_template_string


schedule_bp = Blueprint("schedule", __name__)

DATABASE_URL = os.environ.get("DATABASE_URL")
engine = create_engine(DATABASE_URL, pool_pre_ping=True) if DATABASE_URL else None

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
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
        if raw_id.startswith("0") and raw_id[1:].isdigit():
            ids.append(f"STW-{raw_id[1:]}")
        else:
            ids.append(f"CHE-{raw_id}")

    ids = list(dict.fromkeys(ids))

    if not engine:
        return raw_id

    try:
        placeholders = ", ".join([f":id{i}" for i in range(len(ids))])
        params = {f"id{i}": v for i, v in enumerate(ids)}

        sql = text(f"""
            SELECT id, name, status, branch
            FROM volunteers
            WHERE id IN ({placeholders})
            LIMIT 1
        """)

        with engine.connect() as conn:
            row = conn.execute(sql, params).mappings().first()

        if not row:
            return raw_id

        return str(row["name"]).strip()

    except Exception as e:
        print("find_name_by_id error:", e)
        return raw_id
    
def load_buddha_name_options():
    file = os.path.join(BASE_DIR, "fixed_schedule.xlsx")

    if not os.path.exists(file):
        return []

    try:
        df = pd.read_excel(file, sheet_name="佛台固定")
        df.columns = df.columns.astype(str).str.strip()

        names = []
        for col in ["姓名1", "姓名2", "姓名3"]:
            if col in df.columns:
                for v in df[col].dropna():
                    name = str(v).strip()
                    if name and name != "nan" and name not in names:
                        names.append(name)

        return names
    except Exception as e:
        print("load_buddha_name_options error:", e)
        return []

def get_fixed_buddha_for_date(date_str):
    file = os.path.join(BASE_DIR, "fixed_schedule.xlsx")

    if not os.path.exists(file):
        return []

    try:
        date_obj = pd.to_datetime(date_str)
        weekday_map = {
            0: "星期一",
            1: "星期二",
            2: "星期三",
            3: "星期四",
            4: "星期五",
            5: "星期六",
            6: "星期日",
        }
        weekday = weekday_map[date_obj.weekday()]

        df = pd.read_excel(file, sheet_name="佛台固定")
        df.columns = df.columns.astype(str).str.strip()

        row = df[df["星期"].astype(str).str.strip() == weekday]

        if row.empty:
            return []

        names = []
        r = row.iloc[0]

        for col in ["姓名1", "姓名2", "姓名3"]:
            if col in df.columns:
                name = str(r.get(col, "")).strip()
                if name and name != "nan":
                    names.append(name)

        return names

    except Exception as e:
        print("get_fixed_buddha_for_date error:", e)
        return []

def get_default_time_by_role(role, start_time, end_time):
    role = str(role).strip()

    if role == "值班":
        return start_time, end_time

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

def save_buddha_override(date_str, final_names):
    file = os.path.join(BASE_DIR, "buddha_override.xlsx")

    names = [str(n).strip() for n in final_names if str(n).strip()]

    new_row = {
        "日期": date_str,
        "姓名1": names[0] if len(names) > 0 else "",
        "姓名2": names[1] if len(names) > 1 else "",
        "姓名3": names[2] if len(names) > 2 else "",
    }

    if os.path.exists(file):
        df = pd.read_excel(file)
        df.columns = df.columns.astype(str).str.strip()
        df = df[df["日期"].astype(str) != date_str]
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    else:
        df = pd.DataFrame([new_row])

    with pd.ExcelWriter(file, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)

def delete_prebook_record(record):
    if not os.path.exists(PREBOOK_FILE):
        return

    try:
        df = pd.read_excel(PREBOOK_FILE, sheet_name="预报名")
        df.columns = df.columns.astype(str).str.strip()

        cond = (
            (df["日期"].astype(str) == str(record["日期"])) &
            (df["姓名"].astype(str) == str(record["姓名"])) &
            (df["岗位"].astype(str) == str(record["岗位"])) &
            (df["开始时间"].astype(str) == str(record["开始时间"])) &
            (df["结束时间"].astype(str) == str(record["结束时间"]))
        )

        df = df[~cond]

        with pd.ExcelWriter(PREBOOK_FILE, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="预报名", index=False)

    except Exception as e:
        print("delete_prebook_record error:", e)


@schedule_bp.route("/schedule", methods=["GET", "POST"])
def schedule():
    if not session.get("schedule_login"):
        if request.method == "POST":
            pin = request.form.get("pin", "").strip()
            if pin == SCHEDULE_PIN:
                session["schedule_login"] = True
                return redirect(url_for("schedule.schedule"))

            return "❌ PIN 错误<br><a href='/schedule'>返回</a>"

        return LOGIN_HTML
    
    tomorrow = (datetime.now() + timedelta(days=1)).strftime("%Y-%m-%d")

    mode = request.args.get("mode", "")

    buddha_names = load_buddha_name_options()

    override_date = request.args.get("override_date", tomorrow)
    fixed_buddha_today = get_fixed_buddha_for_date(override_date)

    return render_template_string(
        SCHEDULE_HTML,
        mode=mode,
        times=TIME_OPTIONS,
        roles=ROLE_OPTIONS,
        records=schedule_records,
        tomorrow=tomorrow,
        buddha_names=buddha_names,
        override_date=override_date,
        fixed_buddha_today=fixed_buddha_today,
    )


@schedule_bp.route("/schedule/add", methods=["POST"])
def schedule_add():
    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule"))

    mode = request.form.get("mode", "").strip()

    vol_id = request.form.get("vol_id", "").strip()
    roles = request.form.getlist("roles")
    start_time = request.form.get("start_time", "").strip()
    end_time = request.form.get("end_time", "").strip()

    if not vol_id:
        return "❌ 请填写义工编号<br><a href='/schedule'>返回</a>"

    if not roles:
        return "❌ 请至少选择一个岗位<br><a href='/schedule'>返回</a>"

    name = find_name_by_id(vol_id)

    if mode == "day":
        single_date = request.form.get("single_date", "").strip()
        if not single_date:
            return "❌ 请选择日期<br><a href='/schedule?mode=day'>返回</a>"
        date_list = [single_date]
    else:
        month = request.form.get("month", "").strip()
        days = request.form.getlist("days")

        if not month:
            return "❌ 请填写月份，例如 2026-05<br><a href='/schedule?mode=prebook'>返回</a>"

        if not days:
            return "❌ 请至少选择一天<br><a href='/schedule?mode=prebook'>返回</a>"

        date_list = [f"{month}-{int(day):02d}" for day in days]

    for date_text in date_list:
        for role in roles:
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

    return redirect(url_for("schedule.schedule", mode=mode))


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

@schedule_bp.route("/schedule/delete/<int:index>", methods=["POST"])
def schedule_delete(index):
    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule"))

    mode = request.form.get("mode", "")

    if 0 <= index < len(schedule_records):
        record = schedule_records[index]
        delete_prebook_record(record)
        schedule_records.pop(index)

    return redirect(url_for("schedule.schedule", mode=mode))

@schedule_bp.route("/schedule/override", methods=["POST"])
def schedule_override():
    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule"))

    date = request.form.get("date", "").strip()

    original_names = request.form.getlist("original_name")
    replacement_names = request.form.getlist("replacement_name")

    final_names = []

    for original, replacement in zip(original_names, replacement_names):
        original = str(original).strip()
        replacement = str(replacement).strip()

        if replacement:
            final_names.append(replacement)
        elif original:
            final_names.append(original)

    save_buddha_override(date, final_names)

    return redirect(url_for("schedule.schedule", mode="day", override_date=date))


LOGIN_HTML = """
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>负责人排班系统</title>
<style>
body { font-family: "Microsoft YaHei", Arial; background:#f5f5f5; padding:30px; }
.box { background:white; max-width:600px; margin:auto; padding:30px; border-radius:15px; text-align:center; }
input, button { font-size:24px; padding:12px; margin:8px; }
</style>
</head>
<body>
<div class="box">
<h1>📅 负责人排班系统</h1>
<form method="post">
    <input type="password" name="pin" placeholder="请输入负责人PIN">
    <button type="submit">进入</button>
</form>
</div>
</body>
</html>
"""


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
    max-width: 950px;
    margin: auto;
    padding: 25px;
    border-radius: 15px;
}
input, select, button {
    font-size: 22px;
    padding: 10px;
    margin: 6px;
}
.big-btn {
    font-size: 28px;
    padding: 18px 28px;
    margin: 12px;
    border-radius: 12px;
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

<div style="text-align:center;">
    <a href="/schedule?mode=day">
        <button type="button" class="big-btn">📋 生成当天排班</button>
    </a>

    <a href="/schedule?mode=prebook">
        <button type="button" class="big-btn">📢 生成月预报名表</button>
    </a>
</div>

<hr>

{% if mode == "day" %}


<h2>📋 当天排班模式</h2>

<form method="post" action="/schedule/add">
    <h3>1. 选择日期</h3>
    <input type="date" name="single_date" value="{{ tomorrow }}" required>

    <h3>2. 输入义工编号</h3>
    <input name="vol_id" placeholder="例如 208 / 0160 / 803" required>

    <h3>3. 选择岗位</h3>
    <div class="role-box">
    {% for role in roles %}
    <label>
        <input type="checkbox" name="roles" value="{{ role }}"> {{ role }}
    </label>
    {% endfor %}
    </div>

    <h3>4. 选择时间（只给值班用）</h3>
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

    <input type="hidden" name="mode" value="day">

    <br><br>
    <button type="submit">➕ 加入当天名单</button>
</form>

<hr>

<form method="post" action="/schedule/generate_day">
    <h3>5. 输出 WhatsApp 值班表</h3>
    日期：
    <input type="date" name="date" value="{{ tomorrow }}" required>
    <button type="submit">⚡ 生成 WhatsApp 值班表</button>
</form>

<hr>

<h2>🙏 佛台请假 / 换人</h2>

<form method="get" action="/schedule">
    <input type="hidden" name="mode" value="day">
    日期：
    <input type="date" name="override_date" value="{{ override_date }}" required>
    <button type="submit">查看当天佛台</button>
</form>

<form method="post" action="/schedule/override">
    <input type="hidden" name="date" value="{{ override_date }}">

    <h3>原本佛台：</h3>

    {% if fixed_buddha_today %}
        {% for old_name in fixed_buddha_today %}
            <div style="font-size:22px; margin:10px 0;">
                {{ old_name }}
                <input type="hidden" name="original_name" value="{{ old_name }}">

                换成：
                <select name="replacement_name">
                    <option value="">不换，保留原本</option>
                    {% for n in buddha_names %}
                    <option value="{{ n }}">{{ n }}</option>
                    {% endfor %}
                </select>
            </div>
        {% endfor %}

        <button type="submit">💾 保存佛台请假 / 换人</button>
    {% else %}
        <p style="font-size:22px;">这一天没有找到固定佛台名单。</p>
    {% endif %}
</form>

<hr>

<h2>🙏 佛台请假 / 换人</h2>

<form method="post" action="/schedule/override">

    日期：
    <input type="date" name="date" value="{{ tomorrow }}" required>

    <br><br>

    替换义工：
    <br>

    <input name="id1" placeholder="编号1（例如 208）">
    <input name="id2" placeholder="编号2">
    <input name="id3" placeholder="编号3">

    <br><br>

    <button type="submit">💾 保存佛台替换</button>

</form>

<hr>

<h2>🙏 佛台请假 / 换人</h2>

<form method="post" action="/schedule/override">
    日期：
    <input type="date" name="date" value="{{ tomorrow }}" required>

    <br><br>

    佛台位置1：
    <select name="name1">
        <option value="">-- 留空 --</option>
        {% for n in buddha_names %}
        <option value="{{ n }}">{{ n }}</option>
        {% endfor %}
    </select>

    <br>

    佛台位置2：
    <select name="name2">
        <option value="">-- 留空 --</option>
        {% for n in buddha_names %}
        <option value="{{ n }}">{{ n }}</option>
        {% endfor %}
    </select>

    <br>

    佛台位置3：
    <select name="name3">
        <option value="">-- 留空 --</option>
        {% for n in buddha_names %}
        <option value="{{ n }}">{{ n }}</option>
        {% endfor %}
    </select>

    <br><br>

    <button type="submit">💾 保存佛台替换</button>
</form>

{% elif mode == "prebook" %}

<h2>📢 月预报名模式</h2>

<form method="post" action="/schedule/add">

    <h3>1. 输入义工编号</h3>
    <input name="vol_id" placeholder="例如 208 / 0160 / 803" required>

    <h3>2. 选择月份</h3>
    <input name="month" placeholder="例如 2026-05" required>

    <h3>3. 多选日期</h3>
    <div class="day-box">
    {% for d in range(1, 32) %}
    <label>
        <input type="checkbox" name="days" value="{{ d }}"> {{ d }}
    </label>
    {% endfor %}
    </div>

    <h3>4. 选择岗位</h3>
    <div class="role-box">
    {% for role in roles %}
    <label>
        <input type="checkbox" name="roles" value="{{ role }}"> {{ role }}
    </label>
    {% endfor %}
    </div>

    <h3>5. 选择时间（只给值班用）</h3>
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

    <input type="hidden" name="mode" value="prebook">

    <br><br>
    <button type="submit">➕ 加入预报名</button>
</form>

<hr>

<h2>📢 输出月预报名表</h2>

<form method="post" action="/schedule/monthly_prebook">
    年份：
    <input name="year" value="2026" style="width:120px;" required>

    月份：
    <select name="month">
        {% for m in range(1, 13) %}
        <option value="{{ m }}">{{ m }}月</option>
        {% endfor %}
    </select>

    <button type="submit">📢 生成月预报名表</button>
</form>

{% else %}

<h2 style="text-align:center;">请选择要做的功能</h2>

{% endif %}

<hr>

<form method="post" action="/schedule/clear">
<button type="submit">🗑 清空下方显示名单</button>
</form>

<h2>已加入名单</h2>

<table>
<tr>
    <th>日期</th>
    <th>编号</th>
    <th>姓名</th>
    <th>岗位</th>
    <th>时间</th>
    <th>操作</th>
</tr>

{% for r in records %}
<tr>
    <td>{{ r["日期"] }}</td>
    <td>{{ r["编号"] }}</td>
    <td>{{ r["姓名"] }}</td>
    <td>{{ r["岗位"] }}</td>
    <td>{{ r["开始时间"] }} ~ {{ r["结束时间"] }}</td>

    <td>
        <form method="post" action="/schedule/delete/{{ loop.index0 }}">
            <input type="hidden" name="mode" value="{{ mode }}">
            <button type="submit">❌</button>
        </form>
    </td>
</tr>
{% endfor %}
</table>

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
    height: 600px;
    font-size: 20px;
    padding: 15px;
}
button, a {
    font-size: 22px;
    padding: 10px 18px;
    margin: 8px;
}
.copy-btn {
    background: #4CAF50;
    color: white;
    border: none;
}
</style>
</head>
<body>
<div class="box">

<h1>📋 WhatsApp 值班表</h1>

<a href="/schedule?mode=day">⬅ 返回</a>

<br><br>

<button class="copy-btn" onclick="copyText()">📋 一键复制</button>

<br><br>

<textarea id="output">{{ output }}</textarea>

</div>

<script>
function copyText() {
    var text = document.getElementById("output");
    text.select();
    text.setSelectionRange(0, 99999);
    document.execCommand("copy");
    alert("✅ 已复制，可以直接贴去 WhatsApp");
}
</script>

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
body { font-family: "Microsoft YaHei", Arial; background:#f5f5f5; padding:20px; }
.box { background:white; max-width:900px; margin:auto; padding:25px; border-radius:15px; }
textarea { width:100%; height:650px; font-size:20px; padding:15px; box-sizing:border-box; }
a, button { font-size:22px; padding:10px 18px; margin:8px; }
</style>
</head>
<body>
<div class="box">
<h1>📢 月预报名表</h1>
<a href="/schedule?mode=prebook">⬅ 返回月预报名</a>
<br><br>
<textarea readonly>{{ output }}</textarea>
</div>
</body>
</html>
"""

