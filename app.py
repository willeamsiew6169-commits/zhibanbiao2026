# app.py
# 蕉赖观音堂义工签到系统 v3
# 功能：
# 1) 编号查找、PIN验证、签到、签退
# 2) 今日记录修改/删除、自动备份、秒按缓存
# 3) 中文 / English 双语网页（Excel 内部仍保存中文）
# 4) 输入管理员 PIN 
# 5) 修改 PIN

from __future__ import annotations

import os
import io
import sys
import shutil
import atexit
import socket
import qrcode
import base64
import psycopg2
import threading
import subprocess
import pandas as pd

from io import BytesIO
from pathlib import Path
from typing import Optional
from zoneinfo import ZoneInfo
from pypinyin import lazy_pinyin
from schedule_web import schedule_bp
from id_utils import normalize_member_id
from sqlalchemy import create_engine, text
from psycopg2.extras import RealDictCursor
from openpyxl import Workbook, load_workbook
from psycopg2.pool import SimpleConnectionPool
from datetime import datetime, date, time, timedelta
from excel_style_utils import beautify_attendance_file
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from flask import (
    Flask, request, redirect, url_for,
    render_template_string, flash, jsonify,
    make_response, send_file,
)

DATABASE_URL = os.environ.get("DATABASE_URL")
MY_TZ = ZoneInfo("Asia/Kuala_Lumpur")
today = datetime.now(MY_TZ).strftime("%Y-%m-%d")

def get_db():
    if not DATABASE_URL:
        raise RuntimeError("DATABASE_URL 没有设置")
    return psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor)

pool = SimpleConnectionPool(
    1,
    20,
    dsn=DATABASE_URL,
    cursor_factory=RealDictCursor
)

def db_query(sql, params=None, fetchone=False, fetchall=False):
    import psycopg2

    conn = pool.getconn()
    try:
        with conn.cursor() as cur:
            cur.execute(sql, params or ())
            
            if fetchone:
                result = cur.fetchone()
            elif fetchall:
                result = cur.fetchall()
            else:
                result = None

            conn.commit()
            return result

    except psycopg2.OperationalError:
        # 🔥 连接坏了 → 重试一次
        conn = pool.getconn()
        with conn.cursor() as cur:
            cur.execute(sql, params or ())
            
            if fetchone:
                result = cur.fetchone()
            elif fetchall:
                result = cur.fetchall()
            else:
                result = None

            conn.commit()
            return result

    finally:
        pool.putconn(conn)


def get_today_reading_rows():
    rows = db_query("""
        select *
        from reading
        where date = %s
        order by id desc
    """, (now_date_str(),), fetchall=True)

    return rows or []

def add_reading_record(name, identity, topic, session, time_text):
    db_query("""
        insert into reading
        (date, name, identity, topic, session, time)
        values (%s, %s, %s, %s, %s, %s)
    """, (
        now_date_str(),
        name,
        identity,
        topic,
        session,
        time_text
    ))

def get_reading_topics():
    rows = db_query("""
        select distinct topic
        from reading
        where topic is not null and topic <> ''
        order by topic
    """, fetchall=True)

    return [r["topic"] for r in rows if r.get("topic")]

# =========================
# 1) 基本设定
# =========================
BASE_DIR = Path(__file__).resolve().parent
VOLUNTEERS_FILE = BASE_DIR / "volunteers.xlsx"
ATTENDANCE_FILE = BASE_DIR / "attendance.xlsx"
BACKUP_DIR = BASE_DIR / "backups"

VOLUNTEERS_SHEET = "volunteers"
ATTENDANCE_SHEET = "records"

# 管理员 PIN：你可以改成自己的
ADMIN_PIN = "8888"

# 你的月报/年报脚本名称
REPORT_SCRIPT = BASE_DIR / "zhibanbiao2026.py"

# 保持你旧 attendance.xlsx 格式，避免月报错位
ATTENDANCE_HEADERS = [
    "日期", "姓名", "报名", "签到", "岗位",
    "开始时间", "结束时间", "时数", "备注"
]

TODAY_CODE_ENABLED = True

TODAY_CODE_LIST = [
    "2580", "7312", "4901", "8625", "1047",
    "3698", "5206", "9174", "6842", "0359",
    "2468", "1357", "8080", "1122", "5566",
    "7788", "9090", "3145", "6721", "4826",
]

# 系统内部岗位永远用中文；英文只用于网页显示
ROLES = ["值班", "卫生", "佛台", "供台", "供花", "供果", "膳食", "佛学班"]

ROLE_TEXT = {
    "值班": {"zh": "值班", "en": "Duty"},
    "卫生": {"zh": "卫生", "en": "Cleaning"},
    "佛台": {"zh": "佛台", "en": "Altar"},
    "供台": {"zh": "供台", "en": "Offering Table"},
    "供花": {"zh": "供花", "en": "Flowers"},
    "供果": {"zh": "供果", "en": "Fruit Offering"},
    "膳食": {"zh": "膳食", "en": "Meal"},
    "佛学班": {"zh": "佛学班", "en": "Buddhist Class"},
}

TEXT = {
    "zh": {
        "html_lang": "zh-Hans",
        "system_title": "蕉赖观音堂义工签到系统",
        "check_in": "签到",
        "enter_id": "输入编号",
        "id_placeholder": "例如：123",
        "pin": "PIN",
        "pin_placeholder": "输入PIN",
        "find_volunteer": "查找义工",
        "name": "姓名",
        "phone": "电话",
        "status": "状态",
        "role": "选择岗位",
        "open_records": "今日进行中（未签退）",
        "all_today_records": "今日记录",
        "show_today_records": "查看今日记录",
        "latest_records_note": "只显示最新 20 条，Excel 资料不会被删除。",
        "no_open": "现在没有未签退的义工。",
        "no_today": "今天还没有记录。",
        "start": "开始",
        "time": "时间",
        "hours": "时数",
        "action": "操作",
        "sign_out": "签退",
        "edit": "修改",
        "row_id": "编号",
        "admin_tools": "管理员工具",
        "admin_pin": "请输入管理员 PIN",
        "generate_report": "管理员登录",
        "change_pin": "修改 PIN",
        "language": "语言",
        "chinese": "中文",
        "english": "English",
        "not_found_id": "找不到编号",
        "enter_id_first": "请先输入编号。",
        "lookup_first": "请先按【查找义工】确认姓名。",
        "enter_pin": "请输入 PIN。",
        "signout_prompt": "请输入 PIN 才能签退：",
        "pin_empty": "PIN 不能为空。",
        "edit_title": "修改今日记录",
        "back_home": "返回首页",
        "date": "日期",
        "start_time": "开始时间",
        "end_time": "结束时间",
        "remark": "备注",
        "save_edit": "保存修改",
        "delete_record": "删除这笔误按记录",
        "delete_confirm": "确定删除这笔记录吗？删除前系统会自动备份。",
        "change_pin_title": "修改 PIN",
        "old_pin": "旧 PIN",
        "new_pin": "新 PIN",
        "confirm_new_pin": "确认新 PIN",
        "save": "保存",
        "bhff_title": "📖 白话佛法共修",
        "bhff_enter": "📖 进入共修记录",
        "bhff_desc": "👉 记录今日共修（最少2人）",
        "bhff_record_title": "📖 白话佛法共修记录",
        "today_code": "今日签到码",
        "today_code_placeholder": "请输入现场今日码",
        "admin_title": "🔐 管理员工具",
        "today_code_big": "今日签到码",
        "today_code_warning": "⚠ 请只写在观音堂现场，不要发群",
        "download_data": "📥 下载签到数据",
        "admin_add_record": "🛠 补录签到",
        "admin_records": "✏️ 修改 / 删除今日记录",
        "today_stats": "📊 今日统计",
        "today_checkin": "今日签到",
        "today_not_checkout": "目前未签退",
        "today_checkout_done": "已完成签退",
        "people_count": "人次",
        "people": "人",
        "paid_until": "月费已供养至",
        "pin_wrong": "PIN 不正确，无法显示个人资料",
        "today_topic": "今日主题",
        "topic_placeholder": "例如：佛陀的大智慧",
        "session_remark": "场次 / 备注",
        "session_placeholder": "例如：早上共修 / 晚上共修",
        "volunteer_list": "义工名单",
        "select_all_volunteers": "✅ 全选全部义工",
        "record_study": "✅ 记录共修",
        "extra_friend_name": "非义工佛友姓名",
        "extra_friend_placeholder": "例如：王小明、李美玲、陈先生",
        "extra_friend_tip": "多个名字用逗号、空格或顿号隔开",
        "today_recorded": "今日已记录",
        "identity": "身份",
        "topic": "主题",
        "operation": "操作",
        "today_study_count": "📊 今日共修次数",
        "count": "次数",
        "name": "姓名",
        "status": "状态",
        "phone": "电话",
        "not_registered": "未登记",
        "no_contribution": "暂无月费记录",


    },
    "en": {
        "html_lang": "en",
        "system_title": "Cheras Guan Yin Citta Volunteer Check-in System",
        "check_in": "Check In",
        "enter_id": "Volunteer ID",
        "id_placeholder": "Example: 123",
        "pin": "PIN",
        "pin_placeholder": "Enter PIN",
        "find_volunteer": "Find Volunteer",
        "name": "Name",
        "phone": "Phone",
        "status": "Status",
        "role": "Select Role",
        "open_records": "Currently Checked In",
        "all_today_records": "Today’s Records",
        "show_today_records": "View Today’s Records",
        "latest_records_note": "Only the latest 20 records are shown. Excel data is not deleted.",
        "no_open": "No volunteer is currently checked in.",
        "no_today": "No records yet today.",
        "start": "Start",
        "time": "Time",
        "hours": "Hours",
        "action": "Action",
        "sign_out": "Check Out",
        "edit": "Edit",
        "row_id": "ID",
        "admin_tools": "Admin Tools",
        "admin_pin": "Enter Admin PIN",
        "generate_report": "🔐 Enter Admin Panel",
        "change_pin": "Change PIN",
        "language": "Language",
        "chinese": "中文",
        "english": "English",
        "not_found_id": "Volunteer ID not found",
        "enter_id_first": "Please enter the volunteer ID first.",
        "lookup_first": "Please find the volunteer first.",
        "enter_pin": "Please enter PIN.",
        "signout_prompt": "Enter PIN to check out:",
        "pin_empty": "PIN cannot be empty.",
        "edit_title": "Edit Today’s Record",
        "back_home": "Back to Home",
        "date": "Date",
        "start_time": "Start Time",
        "end_time": "End Time",
        "remark": "Remark",
        "save_edit": "Save Changes",
        "delete_record": "Delete This Record",
        "delete_confirm": "Are you sure you want to delete this record? A backup will be created first.",
        "change_pin_title": "Change PIN",
        "old_pin": "Old PIN",
        "new_pin": "New PIN",
        "confirm_new_pin": "Confirm New PIN",
        "save": "Save",
        "bhff_title": "📖 Bai Hua Fo Fa (BHFF) Group Study",
        "bhff_enter": "📖 Enter Study Records",
        "bhff_desc": "👉 Record today’s group study (minimum 2 participants)",
        "bhff_record_title": "📖 Bai Hua Fo Fa (BHFF) Study Records",
        "today_code": "Today Code",
        "today_code_placeholder": "Enter today's code (on-site)",
        "admin_title": "🔐 Admin Tools",
        "today_code_big": "Today Code",
        "today_code_warning": "⚠ Display this only on-site. Do not share it in group chats.",
        "download_data": "📥 Download Check-in Data",
        "admin_add_record": "🛠 Add Record",
        "admin_records": "✏️ Edit / Delete Today’s Records",
        "today_stats": "📊 Today Summary",
        "today_checkin": "Check-ins",
        "today_not_checkout": "Not Signed Out",
        "today_checkout_done": "Completed",
        "people_count": "times",
        "people": "people",
        "paid_until": "Paid Until",
        "pin_wrong": "PIN incorrect. Cannot display personal info.",
        "today_topic": "Today’s Topic",
        "topic_placeholder": "Example: The Buddha’s Great Wisdom",
        "session_remark": "Session / Remark",
        "session_placeholder": "Example: Morning Study / Evening Study",
        "volunteer_list": "Volunteer List",
        "select_all_volunteers": "✅ Select All Volunteers",
        "record_study": "✅ Save Study Record",
        "extra_friend_name": "Non-volunteer Names",
        "extra_friend_placeholder": "Example: Wang Xiao Ming, Li Mei Ling",
        "extra_friend_tip": "Separate multiple names with commas, spaces, or new lines.",
        "today_recorded": "Today’s Records",
        "identity": "Identity",
        "topic": "Topic",
        "operation": "Action",
        "today_study_count": "📊 Today’s Study Count",
        "count": "Count",
        "name": "Name",
        "status": "Status",
        "phone": "Phone",
        "not_registered": "Not Registered",
        "no_contribution": "No Contribution Record",

    }
}


app = Flask(__name__)
app.secret_key = "change-this-simple-secret"
app.register_blueprint(schedule_bp)

READING_FILE = "reading.xlsx"

# =========================
# 工具函数
# =========================
def get_today_code():
    today = datetime.now(MY_TZ)
    day_index = today.toordinal() % len(TODAY_CODE_LIST)
    return TODAY_CODE_LIST[day_index]


def get_display_today_code():
    """管理员页面显示用：晚上7点后显示明天的现场码。"""
    now = datetime.now(MY_TZ)

    if now.hour >= 19:
        tomorrow = now + timedelta(days=1)
        day_index = tomorrow.toordinal() % len(TODAY_CODE_LIST)
        return TODAY_CODE_LIST[day_index]

    return get_today_code()

def beautify_reading_excel():
    if not os.path.exists(READING_FILE):
        return
    
    print("🔥 正在执行 beautify")

    wb = load_workbook(READING_FILE)
    ws = wb.active

    header_fill = PatternFill("solid", fgColor="D9EAD3")
    header_font = Font(name="Microsoft YaHei", size=16, bold=True)
    body_font = Font(name="Microsoft YaHei", size=15)

    thin = Side(style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for row in ws.iter_rows():
        for cell in row:
            cell.font = header_font if cell.row == 1 else body_font
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )
            cell.border = border
            if cell.row == 1:
                cell.fill = header_fill

    # ✅ 行高：重点
    ws.row_dimensions[1].height = 32

    for r in range(2, ws.max_row + 1):
        ws.row_dimensions[r].height = 30

    # ✅ 列宽：放大
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 45
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 18

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:E{ws.max_row}"

    wb.save(READING_FILE)



def parse_time_to_datetime(t):
    today = datetime.now(MY_TZ)
    s = str(t).strip().lower()

    for fmt in ["%H:%M", "%I:%M%p", "%I:%M %p"]:
        try:
            parsed = datetime.strptime(s, fmt)
            return today.replace(hour=parsed.hour, minute=parsed.minute, second=0, microsecond=0)
        except ValueError:
            pass

    raise ValueError(f"不能解析时间：{t}")

def ensure_reading_file():
    required_cols = ["日期", "姓名", "主题", "场次", "时间"]

    if not os.path.exists(READING_FILE):
        df = pd.DataFrame(columns=required_cols)
        df.to_excel(READING_FILE, index=False)
        return

    df = pd.read_excel(READING_FILE)

    for col in required_cols:
        if col not in df.columns:
            df[col] = ""

    df = df[required_cols]
    df.to_excel(READING_FILE, index=False)
    beautify_reading_excel()


def get_today():
    return datetime.now(MY_TZ).strftime("%Y-%m-%d")

# 👉 从 attendance.xlsx 拿今天签到名单
def get_today_attendees():
    rows = db_query("""
        select distinct name
        from attendance
        where date = %s and signin = 1
    """, (now_date_str(),), fetchall=True)

    return sorted([r["name"] for r in rows if r.get("name")])

def get_today_stats():
    rows = db_query("""
        select *
        from attendance
        where date = %s
    """, (now_date_str(),), fetchall=True)

    total = len(rows)
    open_count = sum(1 for r in rows if not str(r.get("end_time") or "").strip())
    finished = total - open_count

    return {
        "total": total,
        "open": open_count,
        "finished": finished,
    }

# =========================
# 页面：共修记录
# =========================
@app.route("/reading", methods=["GET", "POST"])
def reading():
    t = get_text()
    attendees = get_today_attendees()
    today = now_date_str()

    if request.method == "POST":
        topic = request.form.get("topic", "").strip()
        names = request.form.getlist("names")
        session = request.form.get("session", "").strip()
        extra_text = request.form.get("extra_names", "").strip()

        extra_names = []
        if extra_text:
            for sep in ["，", "、", ",", "\n"]:
                extra_text = extra_text.replace(sep, " ")
            extra_names = [x.strip() for x in extra_text.split(" ") if x.strip()]

        if len(names) + len(extra_names) < 2:
            return f"❌ {t['need_two_people']}<br><a href='/reading'>{t['back_home']}</a>"

        if not topic:
            return f"❌ {t['enter_topic']}<br><a href='/reading'>{t['back_home']}</a>"

        now_time = datetime.now(MY_TZ).strftime("%I:%M %p").lstrip("0")

        for name in names:
            add_reading_record(
                name=name,
                identity=t["volunteer_identity"],
                topic=topic,
                session=session,
                time_text=now_time
            )

        for name in extra_names:
            add_reading_record(
                name=name,
                identity=t["friend_identity"],
                topic=topic,
                session=session,
                time_text=now_time
            )

        return redirect(url_for("reading"))

    rows = get_today_reading_rows()

    today_records = []
    for r in rows:
        today_records.append({
            "id": r.get("id"),
            "日期": r.get("date"),
            "姓名": r.get("name"),
            "身份": r.get("identity"),
            "主题": r.get("topic"),
            "场次": r.get("session"),
            "时间": r.get("time"),
        })

    summary_map = {}
    for r in today_records:
        key = (r["姓名"], r["身份"])
        summary_map[key] = summary_map.get(key, 0) + 1

    today_summary_records = [
        {
            "姓名": name,
            "身份": identity,
            "共修次数": count
        }
        for (name, identity), count in summary_map.items()
    ]

    topic_options = get_reading_topics()

    html = """
    <!doctype html>
    <html lang="{{ t.html_lang }}">
    <head>
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <style>
            body {
                font-family: "Microsoft YaHei", Arial;
                background: #f6f7f2;
                padding: 18px;
                font-size: 24px;
            }
            .card {
                background: white;
                padding: 28px;
                border-radius: 16px;
                margin-bottom: 18px;
                box-shadow: 0 2px 8px #ccc;
            }
            input[type=text] {
                width: 95%;
                font-size: 32px;
                padding: 20px;
                border-radius: 10px;
                border: 1px solid #aaa;
            }
            label {
                display: block;
                padding: 8px;
                font-size: 21px;
            }
            button {
                font-size: 32px;
                padding: 12px 22px;
                border-radius: 18px;
                border: none;
                background: #4CAF50;
                color: white;
                margin: 5px;
            }
            .delete {
                background: #d9534f;
            }
            .edit {
                background: #f0ad4e;
            }
            a {
                text-decoration: none;
            }
            table {
                width: 100%;
                border-collapse: collapse;
                background: white;
                font-size: 18px;
            }
            th, td {
                border: 1px solid #ccc;
                padding: 8px;
                text-align: center;
            }
            th {
                background: #d9ead3;
            }
        </style>
    </head>
    <body>

    <a href="/"><button>⬅ {{ t.back_home }}</button></a>

    <a href="/download_reading">
        <button style="background:#2196F3;">{{ t.download_report }}</button>
    </a>

    <div class="card">
        <h2>{{ t.bhff_record_title }}</h2>

        <form method="post">
            <p>{{ t.today_topic }}：</p>
            <input type="text" name="topic" list="topicList" placeholder="{{ t.topic_placeholder }}">

            <datalist id="topicList">
            {% for topic in topic_options %}
                <option value="{{ topic }}">
            {% endfor %}
            </datalist>

            <p>{{ t.session_remark }}：</p>
            <input type="text" name="session" placeholder="{{ t.session_placeholder }}">

            <p>{{ t.volunteer_list }}：</p>

            <button type="button" onclick="selectAllNames()" style="margin-bottom:10px;">
                {{ t.select_all_volunteers }}
            </button>

            {% for name in attendees %}
                <label>
                    <input type="checkbox" name="names" value="{{name}}">
                    {{name}}
                </label>
            {% endfor %}

            <br>
            <button type="submit">{{ t.record_study }}</button>

            <p>{{ t.extra_friend_name }}：</p>
            <input type="text" name="extra_names" placeholder="{{ t.extra_friend_placeholder }}"
                style="font-size:20px;width:420px;padding:8px;">
            <p style="color:#777;">{{ t.extra_friend_tip }}</p>
        </form>
    </div>

    <div class="card">
        <h2>{{ t.today_recorded }}</h2>

        <table>
            <tr>
                <th>{{ t.name }}</th>
                <th>{{ t.identity }}</th>
                <th>{{ t.topic }}</th>
                <th>{{ t.time }}</th>
                <th>{{ t.operation }}</th>
            </tr>

            {% for r in today_records %}
            <tr>
                <td>{{ r["姓名"] }}</td>
                <td>{{ r["身份"] }}</td>
                <td>{{ r["主题"] }}</td>
                <td>{{ r["时间"] }}</td>
                <td>
                    <a href="/reading_edit/{{ r['id'] }}"><button class="edit">{{ t.edit }}</button></a>
                    <a href="/reading_delete/{{ r['id'] }}" onclick="return confirm('{{ t.delete_confirm_simple }}')">
                        <button class="delete">{{ t.delete_record }}</button>
                    </a>
                </td>
            </tr>
            {% endfor %}
        </table>

        <h3 style="margin-top:20px;">{{ t.today_study_count }}</h3>

        <table>
            <tr>
                <th>{{ t.name }}</th>
                <th>{{ t.identity }}</th>
                <th>{{ t.count }}</th>
            </tr>

            {% for r in today_summary %}
            <tr>
                <td>{{ r["姓名"] }}</td>
                <td>{{ r["身份"] }}</td>
                <td>{{ r["共修次数"] }}</td>
            </tr>
            {% endfor %}
        </table>
    </div>

    <script>
    function selectAllNames() {
        const boxes = document.querySelectorAll('input[name="names"]');
        boxes.forEach(b => b.checked = true);
    }
    </script>

    </body>
    </html>
    """

    return render_template_string(
        html,
        t=t,
        attendees=attendees,
        today_records=today_records,
        today_summary=today_summary_records,
        topic_options=topic_options
    )

@app.route("/reading_delete/<int:record_id>")
def reading_delete(record_id):
    db_query("""
        delete from reading
        where id = %s
    """, (record_id,))

    return redirect(url_for("reading"))


@app.route("/reading_edit/<int:record_id>", methods=["GET", "POST"])
def reading_edit(record_id):
    row = db_query("""
        select *
        from reading
        where id = %s
    """, (record_id,), fetchone=True)

    if not row:
        return "找不到这笔记录<br><a href='/reading'>返回</a>"

    if request.method == "POST":
        new_topic = request.form.get("topic", "").strip()
        new_session = request.form.get("session", "").strip()

        db_query("""
            update reading
            set topic = %s,
                session = %s
            where id = %s
        """, (new_topic, new_session, record_id))

        return redirect(url_for("reading"))

    html = """
    <h2>修改白话佛法记录</h2>
    <form method="post">
        <p>姓名：{{ name }}</p>
        <p>身份：{{ identity }}</p>

        <p>主题：</p>
        <input name="topic" value="{{ topic }}" style="font-size:22px;width:320px;">

        <p>场次 / 备注：</p>
        <input name="session" value="{{ session }}" style="font-size:22px;width:320px;">

        <br><br>
        <button type="submit" style="font-size:22px;">保存修改</button>
    </form>
    <br>
    <a href="/reading">返回</a>
    """

    return render_template_string(
        html,
        name=row.get("name"),
        identity=row.get("identity"),
        topic=row.get("topic") or "",
        session=row.get("session") or ""
    )

# =========================
# 已弃用（改用 Supabase）
# =========================
ATT_CACHE = []
ATT_CACHE_LOADED = True
ATT_DIRTY = False
ATT_LOCK = threading.Lock()
SAVE_INTERVAL_SEC = 5


# =========================
# 2) 语言工具
# =========================
def get_lang() -> str:
    lang = request.cookies.get("lang", "zh")
    return lang if lang in TEXT else "zh"


def get_text() -> dict:
    return TEXT[get_lang()]


def role_label(role: str, lang: str | None = None) -> str:
    lang = lang or get_lang()
    return ROLE_TEXT.get(role, {}).get(lang, role)


@app.route("/set_lang/<lang>")
def set_lang(lang):
    if lang not in TEXT:
        lang = "zh"
    resp = make_response(redirect(request.referrer or url_for("index")))
    resp.set_cookie("lang", lang, max_age=60 * 60 * 24 * 365)
    return resp


# =========================
# 3) 工具函数
# =========================
def now_date_str() -> str:
    return datetime.now(MY_TZ).strftime("%Y-%m-%d")


def now_time_str() -> str:
    return datetime.now(MY_TZ).strftime("%I:%M%p").lstrip("0").lower()


def only_digits(value) -> str:
    return "".join(ch for ch in str(value or "") if ch.isdigit())


def parse_time(value: str) -> Optional[datetime]:
    s = str(value or "").strip().lower().replace(" ", "")
    if not s:
        return None
    for fmt in ["%I:%M%p", "%I%p", "%H:%M"]:
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            pass
    return None


def calc_hours(start_time: str, end_time: str) -> float:
    st = parse_time(start_time)
    et = parse_time(end_time)
    if not st or not et:
        return 0.0
    diff = (et - st).total_seconds() / 3600
    if diff < 0:
        return 0.0
    return round(diff, 2)


def backup_attendance() -> None:
    if ATTENDANCE_FILE.exists():
        BACKUP_DIR.mkdir(exist_ok=True)
        ts = datetime.now(MY_TZ).strftime("%Y%m%d_%H%M%S")
        shutil.copy2(ATTENDANCE_FILE, BACKUP_DIR / f"attendance_{ts}.xlsx")


def safe_save_workbook(wb, path: Path) -> None:
    tmp_path = path.with_suffix(".tmp.xlsx")
    wb.save(tmp_path)
    os.replace(tmp_path, path)


def reload_attendance_cache() -> None:
    return


# =========================
# 4) Excel 读取 / 建立
# =========================
def ensure_attendance_file() -> None:
    return


def load_volunteers() -> list[dict]:
    rows = db_query("""
        select
            id as "编号",
            name as "姓名",
            phone as "电话号码",
            status as "状态",
            pin as "PIN"
        from volunteers
        where status = '在册'
        order by id
    """, fetchall=True)

    return rows or []

def verify_today_code(input_code):
    now = datetime.now(MY_TZ)

    # 1️⃣ 当前有效码
    today_code = get_today_code()

    # 2️⃣ 如果超过晚上7点 → 用“明天的码”
    if now.hour >= 19:
        tomorrow = now + timedelta(days=1)
        day_index = tomorrow.toordinal() % len(TODAY_CODE_LIST)
        today_code = TODAY_CODE_LIST[day_index]

    return str(input_code).strip() == str(today_code)


def find_volunteer(volunteer_id: str):
    raw_id = str(volunteer_id or "").strip()

    if not raw_id:
        return None

    ids = [raw_id]

    # CHE-208 / STW-160 → 同时尝试数字部分
    if "-" in raw_id:
        branch, num = raw_id.split("-", 1)
        branch = branch.strip().upper()
        num = num.strip()

        ids.append(num)

        # STW-160 也尝试 0160
        if branch == "STW" and num.isdigit():
            ids.append("0" + num)

    else:
        # 208 → 尝试 CHE-208
        ids.append(f"CHE-{raw_id}")

        # 0160 → 尝试 STW-160
        if raw_id.startswith("0") and raw_id[1:].isdigit():
            ids.append(f"STW-{raw_id[1:]}")

    # 去重
    ids = list(dict.fromkeys(ids))

    placeholders = ",".join(["%s"] * len(ids))

    result = db_query(f"""
        select
            id as "编号",
            name as "姓名",
            status as "状态",
            phone as "电话号码",
            pin as "PIN",
            branch as "分会"
        from volunteers
        where id in ({placeholders})
        limit 1
    """, tuple(ids), fetchone=True)

    return result

def to_member_id(volunteer):
    branch = str(volunteer.get("分会") or "CHE").strip()
    vol_id = str(volunteer.get("编号") or "").strip()

    if not vol_id:
        return ""

    if "-" in vol_id:
        return vol_id

    return f"{branch}-{vol_id}"

def verify_pin_for_volunteer(volunteer, pin):
    input_pin = str(pin).strip()

    # 1️⃣ 先用数据库 PIN（如果有）
    real_pin = volunteer.get("PIN")

    if real_pin:
        return input_pin == str(real_pin).strip()

    # 2️⃣ fallback → 电话后4位 or 0000
    phone = only_digits(volunteer.get("电话号码", ""))

    if not phone:
        return input_pin == "0000"   # ✅ 关键在这里

    return input_pin == phone[-4:]

def get_member_paid_until(member_id):
    row = db_query("""
        select max(paid_month) as paid_until
        from member_payments
        where member_id = %s
    """, (member_id,), fetchone=True)

    return row.get("paid_until") if row else ""

def get_member_paid_until(member_id):
    row = db_query("""
        select max(paid_month) as paid_until
        from member_payments
        where member_id = %s
    """, (member_id,), fetchone=True)

    return row.get("paid_until") if row else ""

def _load_attendance_rows_from_excel() -> list[dict]:
    return []

def load_attendance_rows() -> list[dict]:
    rows = db_query("""
        select *
        from attendance
        order by id desc
        limit 200
    """, fetchall=True)

    result = []
    for r in rows:
        result.append({
            "日期": r.get("date"),
            "编号": r.get("volunteer_id"),
            "姓名": r.get("name"),
            "报名": r.get("signup"),
            "签到": r.get("signin"),
            "岗位": r.get("role"),
            "开始时间": r.get("start_time"),
            "结束时间": r.get("end_time") or "",
            "时数": r.get("hours") or "",
            "备注": r.get("remark") or "",
            "_row": r.get("id"),
        })

    return result


def mark_attendance_dirty() -> None:
    global ATT_DIRTY
    ATT_DIRTY = True


def flush_attendance_to_excel(force: bool = False) -> None:
    return


def background_saver():
    while True:
        threading.Event().wait(SAVE_INTERVAL_SEC)
        try:
            flush_attendance_to_excel(force=False)
        except Exception as e:
            print("后台保存 attendance.xlsx 失败：", e)


def start_background_saver_once():
    t = threading.Thread(target=background_saver, daemon=True)
    t.start()


def format_date_value(d) -> str:
    if isinstance(d, datetime):
        return d.strftime("%Y-%m-%d")
    if isinstance(d, date):
        return d.strftime("%Y-%m-%d")
    return str(d or "").strip()


def get_today_open_records() -> list[dict]:
    today = now_date_str()

    rows = db_query("""
        select *
        from attendance
        where date = %s
          and (end_time is null or end_time = '')
        order by id desc
    """, (today,), fetchall=True)

    result = []
    for r in rows:
        result.append({
            "日期": r.get("date"),
            "编号": r.get("volunteer_id"),
            "姓名": r.get("name"),
            "报名": r.get("signup"),
            "签到": r.get("signin"),
            "岗位": r.get("role"),
            "开始时间": r.get("start_time"),
            "结束时间": r.get("end_time") or "",
            "时数": r.get("hours") or "",
            "备注": r.get("remark") or "",
            "_row": r.get("id"),
        })

    return result


def get_today_records(limit: int | None = None) -> list[dict]:
    today = now_date_str()
    out = []
    for item in load_attendance_rows():
        if format_date_value(item.get("日期")) == today:
            out.append(item)
    if limit is not None:
        return out[-limit:]
    return out


# =========================
# 5) 签到 / 签退 / 修改
# =========================
def sign_in(volunteer_id: str, pin: str, role: str) -> tuple[bool, str]:
    volunteer_id = normalize_member_id(volunteer_id)
    role = str(role or "").strip()

    if not volunteer_id:
        return False, "请输入编号。"
    if not pin:
        return False, "请输入 PIN。"
    if role not in ROLES:
        return False, "请选择正确岗位。"

    volunteer = find_volunteer(volunteer_id)
    if not volunteer:
        return False, f"找不到编号：{volunteer_id}"

    if volunteer.get("状态") not in ["在册", "", None]:
        return False, f"此义工状态不是在册：{volunteer.get('状态')}"

    if not verify_pin_for_volunteer(volunteer, pin):
        return False, "PIN 不正确。"

    opened = db_query("""
        select id
        from attendance
        where date = %s
          and volunteer_id = %s
          and (end_time is null or end_time = '')
        order by id desc
        limit 1
    """, (now_date_str(), volunteer["编号"]), fetchone=True)

    if opened:
        return False, f"{volunteer['姓名']} 今天已经签到，还没签退。请先签退。"

    db_query("""
        insert into attendance
        (date, volunteer_id, name, signup, signin, role, start_time, end_time, hours, remark)
        values (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
    """, (
        now_date_str(),
        volunteer["编号"],
        volunteer["姓名"],
        0,
        1,
        role,
        now_time_str(),
        "",
        None,
        "iPad签到"
    ))

    phone = volunteer.get("电话号码") or ""
    paid_until = get_member_paid_until(to_member_id(volunteer))

    # 电话显示
    if phone:
        phone_text = phone
    else:
        phone_text = "未登记"

    extra = f"\n电话：{phone_text}"

    # 月费显示（只在有记录时才显示）
    if paid_until:
        extra += f"\n月费已供养至：{paid_until}"

    return True, f"{volunteer['姓名']} 已签到：{role}{extra}"

def sign_out(volunteer_id: str, pin: str) -> tuple[bool, str]:
    volunteer_id = normalize_member_id(volunteer_id)

    if not volunteer_id:
        return False, "请输入编号。"
    if not pin:
        return False, "请输入 PIN。"

    volunteer = find_volunteer(volunteer_id)
    if not volunteer:
        return False, f"找不到编号：{volunteer_id}"

    if not verify_pin_for_volunteer(volunteer, pin):
        return False, "PIN 不正确。"

    row = db_query("""
        select *
        from attendance
        where date = %s
          and volunteer_id = %s
          and (end_time is null or end_time = '')
        order by id desc
        limit 1
    """, (now_date_str(), volunteer["编号"]), fetchone=True)

    if not row:
        return False, f"{volunteer['姓名']} 今天没有未签退记录。"

    end_time = now_time_str()

    try:
        start_dt = parse_time_to_datetime(row["start_time"])
        end_dt = parse_time_to_datetime(end_time)
        hours = round((end_dt - start_dt).total_seconds() / 3600, 2)
        if hours < 0:
            hours = 0
    except Exception:
        hours = None

    db_query("""
        update attendance
        set end_time = %s,
            hours = %s
        where id = %s
    """, (end_time, hours, row["id"]))

    return True, f"{volunteer['姓名']} 已签退。"


def update_record(row_number: int, role: str, start_time: str, end_time: str, remark: str) -> tuple[bool, str]:
    if role not in ROLES:
        return False, "请选择正确岗位。"

    row = db_query("""
        select *
        from attendance
        where id = %s
    """, (row_number,), fetchone=True)

    if not row:
        return False, "找不到这笔记录。"

    if row.get("date") != now_date_str():
        return False, "为了安全，页面只允许修改今天的记录。"

    start_time = str(start_time or "").strip()
    end_time = str(end_time or "").strip()
    remark = str(remark or "").strip()

    if start_time and end_time:
        try:
            hours = calc_hours(start_time, end_time)
        except Exception:
            hours = None
    else:
        hours = None

    db_query("""
        update attendance
        set role = %s,
            start_time = %s,
            end_time = %s,
            hours = %s,
            remark = %s
        where id = %s
    """, (
        role,
        start_time,
        end_time,
        hours,
        remark,
        row_number
    ))

    return True, "记录已修改。"


def delete_record(row_number: int) -> tuple[bool, str]:
    row = db_query("""
        select *
        from attendance
        where id = %s
    """, (row_number,), fetchone=True)

    if not row:
        return False, "找不到这笔记录。"

    if row.get("date") != now_date_str():
        return False, "为了安全，页面只允许删除今天的记录。"

    name = row.get("name") or ""

    db_query("""
        delete from attendance
        where id = %s
    """, (row_number,))

    return True, f"已删除 {name} 的这笔记录。"

def get_member_payment(member_id):
    rows = db_query("""
        select paid_month
        from member_payments
        where member_id = %s
        order by paid_month
    """, (member_id,), fetchall=True)

    months = [r["paid_month"] for r in rows]
    latest = months[-1] if months else ""

    return months, latest  

# =========================
# 6) 页面
# =========================
PAGE = """
<!doctype html>
<html lang="{{ t.html_lang }}">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{{ t.system_title }}</title>
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, "Microsoft YaHei", Arial, sans-serif; background:#f6f6f6; margin:0; font-size: 22px;}
    .wrap { max-width: 900px; margin: 0 auto; padding: 18px; }
    .card { background:white; border-radius:18px; padding:24px; margin-bottom:18px; box-shadow:0 2px 12px rgba(0,0,0,.08); }
    h1 { font-size: 30px; margin: 6px 0 18px; }
    h2 { font-size: 24px; margin: 0 0 14px; }
    label { font-size: 20px; font-weight: 700; display:block; margin-bottom:8px; }
    input, select { width:100%; font-size: 30px; padding: 16px; border-radius: 14px; border:1px solid #ccc; box-sizing:border-box; }
    .row { display:grid; grid-template-columns: 1fr 1fr; gap:14px; }
    button { font-size: 28px; font-weight: 800; padding: 14px 18px; border:0; border-radius: 16px; cursor:pointer; }
    .btn-find { background:#0d6efd; color:white; width:100%; margin-top:16px; }
    .btn-in { background:#198754; color:white; width:100%; margin-top:16px; font-size:30px; }
    .btn-out { background:#dc3545; color:white; }
    .btn-admin { background:#6f42c1; color:white; width:100%; margin-top:16px; }
    .btn-lang { display:inline-block; font-size:18px; padding:8px 12px; background:white; border-radius:999px; text-decoration:none; color:#333; border:1px solid #ddd; margin-right:6px; }
    .msg { padding:18px; border-radius:14px; font-size:20px; margin-bottom:14px; white-space:pre-wrap; }
    .ok { background:#d1e7dd; color:#0f5132; }
    .bad { background:#f8d7da; color:#842029; }
    .person { font-size: 26px; line-height:1.8; background:#f8f9fa; padding:16px; border-radius:14px; margin-top:12px; }
    table { width:100%; border-collapse: collapse; font-size:20px; }
    th, td { padding:12px; border-bottom:1px solid #eee; text-align:left; vertical-align:top; }
    th { background:#f1f1f1; }
    .muted { color:#666; font-size:18px; }
    .pill { display:inline-block; background:#eee; border-radius:999px; padding:4px 10px; }
    .topbar { margin-bottom:12px; }
    @media (max-width: 700px) { .row { grid-template-columns: 1fr; } h1 {font-size:26px;} table {font-size:17px;} button{font-size:22px;} }
  </style>
</head>
<body>
<div class="wrap">

  <div class="topbar">
    <span class="muted">{{ t.language }}：</span>
    <a class="btn-lang" href="{{ url_for('set_lang', lang='zh') }}">{{ t.chinese }}</a>
    <a class="btn-lang" href="{{ url_for('set_lang', lang='en') }}">{{ t.english }}</a>
    <a class="btn-lang" href="{{ url_for('change_pin_page') }}">{{ t.change_pin }}</a>
  </div>

  <h1>{{ t.system_title }}</h1>

  {% with messages = get_flashed_messages(with_categories=true) %}
    {% for category, message in messages %}
      <div class="msg {{ 'ok' if category == 'ok' else 'bad' }}">{{ message }}</div>
    {% endfor %}
  {% endwith %}

  <div class="card">
    <h2>{{ t.today_stats }}</h2>

    {{ t.today_checkin }}：{{ today_count }} {{ t.people_count }}<br>
    {{ t.today_not_checkout }}：{{ not_out }} {{ t.people }}<br>
    {{ t.today_checkout_done }}：{{ done_out }} {{ t.people }}
  </div>

  <div class="card">
    <h2>✅ {{ t.check_in }}</h2>
    <form method="post" action="{{ url_for('do_sign_in') }}" onsubmit="return quickSignIn();">
      <div class="row">
        <div>
          <label>{{ t.enter_id }}</label>
          <input id="volunteer_id" name="volunteer_id" inputmode="numeric" autocomplete="off" placeholder="{{ t.id_placeholder }}" required>
        </div>
        <div>
          <label>{{ t.pin }}</label>
          <input id="pin" name="pin" type="password" inputmode="numeric" pattern="[0-9]*" autocomplete="off" placeholder="{{ t.pin_placeholder }}" required>
        </div>
      </div>

      {% if today_code_enabled %}
      <div>
        <label>{{ t.today_code }}</label>
        <input 
          type="tel" 
          name="today_code" 
          inputmode="numeric" 
          pattern="[0-9]*"
          placeholder="{{ t.today_code_placeholder }}" 
          required
        
        >
      </div>
      {% endif %}

      <button type="button" class="btn-find" onclick="lookupVolunteer()">
        {{ t.find_volunteer }}
      </button>
      <div id="personBox" class="person" style="display:none;"></div>

      <label style="margin-top:16px;">{{ t.role }}</label>
      <select id="role" name="role" required>
        {% for role in roles %}
          <option value="{{ role }}">{{ role_label(role) }}</option>
        {% endfor %}
      </select>

      <button id="signInBtn" class="btn-in" type="submit" disabled>✅ {{ t.check_in }}</button>
    </form>
  </div>

  <div class="card">
    <h2>⛔ {{ t.open_records }}</h2>
    {% if open_records %}
      <table>
        <thead><tr><th>{{ t.name }}</th><th>{{ t.role }}</th><th>{{ t.start }}</th><th>{{ t.action }}</th></tr></thead>
        <tbody>
        {% for r in open_records %}
          <tr>
            <td><b>{{ r.get('姓名','') }}</b>{% if r.get('编号') %}<br><span class="muted">{{ t.row_id }}：{{ r.get('编号','') }}</span>{% endif %}</td>
            <td><span class="pill">{{ role_label(r.get('岗位','')) }}</span></td>
            <td>{{ r.get('开始时间','') }}</td>
            <td>
              <form method="post" action="{{ url_for('do_sign_out') }}" onsubmit="return askSignOutPin(this);">
                <input type="hidden" name="row_number" value="{{ r.get('_row') }}">
                <input type="hidden" name="pin" value="">
                <button class="btn-out" type="submit">{{ t.sign_out }}</button>
              </form>
              <a href="{{ url_for('edit_page', row_number=r.get('_row')) }}" style="display:inline-block;margin-top:8px;font-size:20px;">{{ t.edit }}</a>
            </td>
          </tr>
        {% endfor %}
        </tbody>
      </table>
    {% else %}
      <div class="muted">{{ t.no_open }}</div>
    {% endif %}
  </div>

  <div class="card">
    <details>
      <summary style="font-size:24px;font-weight:800;cursor:pointer;">📋 {{ t.show_today_records }}</summary>
      <div class="muted" style="margin:10px 0;">{{ t.latest_records_note }}</div>
      {% if today_records %}
        <table>
          <thead><tr><th>{{ t.name }}</th><th>{{ t.role }}</th><th>{{ t.time }}</th><th>{{ t.hours }}</th></tr></thead>
          <tbody>
          {% for r in today_records %}
            <tr>
              <td>{{ r.get('姓名','') }}</td>
              <td>{{ role_label(r.get('岗位','')) }}</td>
              <td>{{ r.get('开始时间','') }} ~ {{ r.get('结束时间','') }}</td>
              <td>{{ r.get('时数','') }}<br><a href="{{ url_for('edit_page', row_number=r.get('_row')) }}" style="font-size:18px;">{{ t.edit }}</a></td>
            </tr>
          {% endfor %}
          </tbody>
        </table>
      {% else %}
        <div class="muted">{{ t.no_today }}</div>
      {% endif %}
    </details>
  </div>

  <div class="card">
    <h2>{{ t.bhff_title }}</h2>

    <a href="/reading">
      <button style="width:100%;font-size:26px;padding:18px;border-radius:14px;background:#0d6efd;color:white;">
        {{ t.bhff_enter }}
      </button>
    </a>

    <div style="font-size:16px;color:#666;margin-top:10px;">
      {{ t.bhff_desc }}
    </div>
  </div>

  <div class="card">
    <h2>🔐 {{ t.admin_tools }}</h2>

    <form method="post" action="{{ url_for('admin_report') }}">
      <label>{{ t.admin_pin }}</label>
      <input name="admin_pin" type="password" inputmode="numeric" pattern="[0-9]*" autocomplete="off" placeholder="{{ t.admin_pin }}" required>
      <button class="btn-admin" type="submit">📊 {{ t.generate_report }}</button>
    </form>
    
  </div>

<script>
let currentVolunteerName = '';
const TXT = {
  enter_id_first: {{ t.enter_id_first|tojson }},
  not_found_id: {{ t.not_found_id|tojson }},
  name: {{ t.name|tojson }},
  phone: {{ t.phone|tojson }},
  status: {{ t.status|tojson }},
  lookup_first: {{ t.lookup_first|tojson }},
  enter_pin: {{ t.enter_pin|tojson }},
  signout_prompt: {{ t.signout_prompt|tojson }},
  pin_empty: {{ t.pin_empty|tojson }},
  paid_until: {{ t.paid_until|tojson }},
  pin_wrong: {{ t.pin_wrong|tojson }},
};

async function lookupVolunteer() {
  const id = document.getElementById('volunteer_id').value.trim();
  const box = document.getElementById('personBox');
  const btn = document.getElementById('signInBtn');

  currentVolunteerName = '';
  btn.disabled = true;

  if (!id) {
    box.style.display = 'block';
    box.innerHTML = '<span style="color:#842029;">' + TXT.enter_id_first + '</span>';
    return;
  }

  const pin = document.getElementById('pin').value.trim();

  const formData = new FormData();
  formData.append('pin', pin);

  const res = await fetch('/api/volunteer/' + encodeURIComponent(id), {
    method: 'POST',
    body: formData
  });

  const data = await res.json();
  box.style.display = 'block';

  if (data.ok) {
    currentVolunteerName = data.volunteer.姓名;

    let html =
      `${TXT.name}：<b>${data.volunteer.姓名}</b><br>` +
      `${TXT.status}：${data.volunteer.状态 || '-'}`;

    if (data.volunteer.pin_ok) {
      html += `<br>${TXT.phone}：${data.volunteer.电话号码 || TXT.not_registered}`;

      const paidUntil = data.volunteer["月费已供养至"];

      if (paidUntil && paidUntil !== "-") {
        html += `<br>${TXT.paid_until}：${paidUntil}`;
      } else {
        html += `<br>${TXT.no_contribution}`;
      }

    } else if (pin) {
      html += `<br><span style="color:#842029;">${TXT.pin_wrong}</span>`;
    }

    box.innerHTML = html;
    btn.disabled = false;

  } else {
    box.innerHTML = `<span style="color:#842029;">${TXT.not_found_id}：${id}</span>`;
  }
}

function quickSignIn() {
  const pin = document.getElementById('pin').value.trim();
  if (!currentVolunteerName) {
    alert(TXT.lookup_first);
    return false;
  }
  if (!pin) {
    alert(TXT.enter_pin);
    return false;
  }
  return true;
}

function askSignOutPin(form) {
  const pin = prompt(TXT.signout_prompt);
  if (pin === null) return false;
  if (!pin.trim()) {
    alert(TXT.pin_empty);
    return false;
  }
  form.querySelector('input[name="pin"]').value = pin.trim();
  return true;
}

setTimeout(() => {
  document.querySelectorAll('.msg.ok').forEach(el => {
    el.style.display = 'none';
  });
}, 8000);

</script>
</body>
</html>
"""


EDIT_PAGE = """
<!doctype html>
<html lang="{{ t.html_lang }}">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{{ t.edit_title }}</title>
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, "Microsoft YaHei", Arial, sans-serif; background:#f6f6f6; margin:0; }
    .wrap { max-width: 760px; margin: 0 auto; padding: 18px; }
    .card { background:white; border-radius:18px; padding:20px; box-shadow:0 2px 12px rgba(0,0,0,.08); }
    h1 { font-size: 28px; }
    label { font-size: 20px; font-weight: 700; display:block; margin-top:14px; margin-bottom:8px; }
    input, select { width:100%; font-size: 24px; padding: 14px; border-radius: 12px; border:1px solid #ccc; box-sizing:border-box; }
    button { font-size: 22px; font-weight: 800; padding: 14px 18px; border:0; border-radius: 14px; cursor:pointer; margin-top:16px; }
    .save { background:#198754; color:white; width:100%; }
    .delete { background:#dc3545; color:white; width:100%; }
    .back { display:inline-block; margin-bottom:14px; font-size:20px; }
    .info { font-size:22px; line-height:1.7; background:#f8f9fa; padding:12px; border-radius:12px; }
  </style>
</head>
<body>
<div class="wrap">
  <a class="back" href="{{ url_for('index') }}">← {{ t.back_home }}</a>
  <div class="card">
    <h1>{{ t.edit_title }}</h1>
    <div class="info">
      {{ t.name }}：<b>{{ record.get('姓名','') }}</b><br>
      {{ t.date }}：{{ record.get('日期','') }}
    </div>

    <form method="post" action="{{ url_for('save_edit') }}">
      <input type="hidden" name="row_number" value="{{ row_number }}">

      <label>{{ t.role }}</label>
      <select name="role" required>
        {% for role in roles %}
          <option value="{{ role }}" {% if role == record.get('岗位') %}selected{% endif %}>{{ role_label(role) }}</option>
        {% endfor %}
      </select>

      <label>{{ t.start_time }}</label>
      <input name="start_time" value="{{ record.get('开始时间','') }}" placeholder="10:00am">

      <label>{{ t.end_time }}</label>
      <input name="end_time" value="{{ record.get('结束时间','') }}" placeholder="">

      <label>{{ t.remark }}</label>
      <input name="remark" value="{{ record.get('备注','') }}">

      <button class="save" type="submit">{{ t.save_edit }}</button>
    </form>

    <form method="post" action="{{ url_for('delete_edit') }}" onsubmit="return confirm({{ t.delete_confirm|tojson }});">
      <input type="hidden" name="row_number" value="{{ row_number }}">
      <button class="delete" type="submit">{{ t.delete_record }}</button>
    </form>
  </div>
</div>
</body>
</html>
"""


PIN_PAGE = """
<!doctype html>
<html lang="{{ t.html_lang }}">
<head>
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{{ t.change_pin_title }}</title>
<style>
body{font-family:-apple-system,BlinkMacSystemFont,"Microsoft YaHei",Arial,sans-serif;background:#f6f6f6;padding:20px}
.card{max-width:600px;margin:auto;background:white;border-radius:18px;padding:20px;box-shadow:0 2px 12px rgba(0,0,0,.08)}
input{width:100%;font-size:24px;margin:10px 0;padding:14px;border-radius:12px;border:1px solid #ccc;box-sizing:border-box}
button{font-size:24px;padding:14px;width:100%;border:0;border-radius:14px;background:#198754;color:white;font-weight:800}
a{font-size:20px}
.msg { padding:14px; border-radius:12px; font-size:20px; margin-bottom:14px; }
.ok { background:#d1e7dd; color:#0f5132; }
.bad { background:#f8d7da; color:#842029; }
</style>
</head>
<body>
<div class="card">
<a href="{{ url_for('index') }}">← {{ t.back_home }}</a>
<h2>{{ t.change_pin_title }}</h2>

{% with messages = get_flashed_messages(with_categories=true) %}
  {% for category, message in messages %}
    <div class="msg {{ 'ok' if category == 'ok' else 'bad' }}">{{ message }}</div>
  {% endfor %}
{% endwith %}

<form method="post">
<input name="id" placeholder="{{ t.enter_id }}" inputmode="numeric" required>
<input name="old" type="password" placeholder="{{ t.old_pin }}" inputmode="numeric" required>
<input name="new" type="password" placeholder="{{ t.new_pin }}" inputmode="numeric" required>
<input name="confirm" type="password" placeholder="{{ t.confirm_new_pin }}" inputmode="numeric" required>
<button type="submit">{{ t.save }}</button>
</form>
</div>
</body>
</html>
"""

def get_local_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
    except:
        ip = "127.0.0.1"
    finally:
        s.close()
    return ip


def make_qr_base64(text):
    img = qrcode.make(text)
    buffer = BytesIO()
    img.save(buffer, format="PNG")
    return base64.b64encode(buffer.getvalue()).decode()
# =========================
# 7) 路由
# =========================
@app.route("/")
def index():
    return render_template_string(
        PAGE,
        t=get_text(),
        roles=ROLES,
        role_label=role_label,
        open_records=get_today_open_records(),
        today_records=get_today_records(limit=20),
        today_code_enabled=TODAY_CODE_ENABLED,
        stats=get_today_stats(),
    )

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

@app.route("/qr")
def qr_page():
    ip = get_local_ip()
    url = f"http://{ip}:5000"
    qr_base64 = make_qr_base64(url)

    return f"""
    <!DOCTYPE html>
    <html lang="zh">
    <head>
        <meta charset="UTF-8">
        <title>扫码签到</title>
        <style>
            body {{
                font-family: "Microsoft YaHei", Arial, sans-serif;
                text-align: center;
                background: #fff8e7;
                padding: 30px;
            }}
            .card {{
                max-width: 520px;
                margin: auto;
                background: white;
                border-radius: 24px;
                padding: 30px;
                box-shadow: 0 8px 30px rgba(0,0,0,0.12);
            }}
            h1 {{
                font-size: 34px;
                margin-bottom: 8px;
            }}
            h2 {{
                font-size: 22px;
                color: #8a5a00;
                margin-top: 0;
            }}
            img {{
                width: 320px;
                height: 320px;
                margin: 20px 0;
            }}
            .url {{
                font-size: 18px;
                word-break: break-all;
                color: #333;
            }}
            .tip {{
                font-size: 20px;
                margin-top: 18px;
                color: #555;
            }}
        </style>
    </head>
    <body>
        <div class="card">
            <h1>义工值班签到系统</h1>
            <h2>蕉赖观音堂分会</h2>

            <img src="data:image/png;base64,{qr_base64}">

            <div class="url">{url}</div>
            <div class="tip">请用手机扫码签到 / 签退</div>
            <div class="tip">请输入义工编号与 PIN</div>
        </div>
    </body>
    </html>
    """

@app.route("/signin", methods=["POST"])
def do_sign_in():
    if TODAY_CODE_ENABLED:
        input_code = request.form.get("today_code", "").strip()
        if not verify_today_code(input_code):
            flash("今日签到码错误，请看现场公布的号码", "bad")
            return redirect(url_for("index"))

    ok, msg = sign_in(
        request.form.get("volunteer_id", ""),
        request.form.get("pin", ""),
        request.form.get("role", ""),
    )
    flash(msg, "ok" if ok else "bad")
    return redirect(url_for("index"))


@app.route("/signout", methods=["POST"])
def do_sign_out():
    try:
        row_number = int(request.form.get("row_number", "0"))
    except Exception:
        row_number = 0

    row = db_query("select * from attendance where id=%s", (row_number,), fetchone=True)

    if not row:
        flash("记录不存在", "bad")
        return redirect(url_for("index"))

    ok, msg = sign_out(row["volunteer_id"], request.form.get("pin", ""))
    flash(msg, "ok" if ok else "bad")
    return redirect(url_for("index"))

@app.route("/edit/<int:row_number>")
def edit_page(row_number):
    record = None
    for r in get_today_records():
        if int(r.get("_row", 0)) == row_number:
            record = r
            break
    if not record:
        flash("找不到今天这笔记录。", "bad")
        return redirect(url_for("index"))
    return render_template_string(
        EDIT_PAGE,
        t=get_text(),
        record=record,
        row_number=row_number,
        roles=ROLES,
        role_label=role_label
    )


@app.route("/edit/save", methods=["POST"])
def save_edit():
    try:
        row_number = int(request.form.get("row_number", "0"))
    except Exception:
        row_number = 0
    ok, msg = update_record(
        row_number=row_number,
        role=request.form.get("role", ""),
        start_time=request.form.get("start_time", ""),
        end_time=request.form.get("end_time", ""),
        remark=request.form.get("remark", ""),
    )
    flash(msg, "ok" if ok else "bad")
    return redirect(url_for("index"))


@app.route("/edit/delete", methods=["POST"])
def delete_edit():
    try:
        row_number = int(request.form.get("row_number", "0"))
    except Exception:
        row_number = 0
    ok, msg = delete_record(row_number)
    flash(msg, "ok" if ok else "bad")
    return redirect(url_for("index"))

@app.route("/download_data")
def download_data():
    from flask import send_file
    import pandas as pd
    import io

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

@app.route("/download_reading")
def download_reading():
    
    rows = db_query("""
        select *
        from reading
        order by date, time
    """, fetchall=True)

    if not rows:
        return "没有数据"

    df = pd.DataFrame(rows)

    # 改中文列名
    df = df.rename(columns={
        "date": "日期",
        "name": "姓名",
        "identity": "身份",
        "topic": "主题",
        "session": "场次",
        "time": "时间"
    })

    # 只保留需要的列
    df = df[["日期", "姓名", "身份", "主题", "场次", "时间"]]

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="白话佛法记录")

    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="reading_report.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

def normalize_member_query_id(value: str) -> str:
    s = str(value or "").strip().upper()

    if not s:
        return ""

    if s.endswith(".0"):
        s = s[:-2]

    if "-" in s:
        return s

    if s.startswith("0") and s.isdigit():
        return f"STW-{int(s)}"

    if s.isdigit():
        return f"CHE-{int(s)}"

    return s


def verify_member_pin(member, input_pin):
    input_pin = str(input_pin or "").strip()

    real_pin = str(member.get("pin") or "").strip()

    if real_pin:
        return input_pin == real_pin

    phone = only_digits(member.get("phone") or "")
    return input_pin == phone[-4:]


def get_member_payment(member_id):
    rows = db_query("""
        select paid_month
        from member_payments
        where member_id = %s
        order by paid_month
    """, (member_id,), fetchall=True)

    months = [r["paid_month"] for r in (rows or [])]
    latest = months[-1] if months else ""

    return months, latest


@app.route("/member", methods=["GET", "POST"])
def member_page():
    result = None
    error = ""

    if request.method == "POST":
        raw_id = request.form.get("member_id", "")
        pin = request.form.get("pin", "")
        member_id = normalize_member_query_id(raw_id)

        member = db_query("""
            select *
            from members
            where member_id = %s
        """, (member_id,), fetchone=True)

        if not member:
            error = "❌ 找不到月费编号"
        elif not verify_member_pin(member, pin):
            error = "❌ PIN 不正确"
        else:
            months, latest = get_member_payment(member_id)
            result = {
                "member_id": member_id,
                "name": member.get("name") or "",
                "english_name": member.get("english_name") or "",
                "phone": member.get("phone") or "",
                "months": months,
                "latest": latest,
            }

    return render_template_string("""
<!doctype html>
<html>
<head>
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>月费查询</title>
<style>
body {
    font-family: "Microsoft YaHei", Arial, sans-serif;
    background:#f6f6f6;
    padding:20px;
}
.card {
    max-width:650px;
    margin:auto;
    background:white;
    border-radius:18px;
    padding:22px;
    box-shadow:0 2px 12px rgba(0,0,0,.08);
}
input {
    width:100%;
    font-size:26px;
    padding:14px;
    margin:8px 0 16px;
    border-radius:12px;
    border:1px solid #ccc;
    box-sizing:border-box;
}
button {
    width:100%;
    font-size:26px;
    padding:14px;
    border:0;
    border-radius:14px;
    background:#198754;
    color:white;
    font-weight:bold;
}
.result {
    margin-top:20px;
    font-size:22px;
    line-height:1.8;
    background:#f8f9fa;
    padding:16px;
    border-radius:12px;
}
.error {
    margin-top:20px;
    font-size:22px;
    color:#842029;
    background:#f8d7da;
    padding:14px;
    border-radius:12px;
}
a { font-size:20px; }
</style>
</head>
<body>
<div class="card">
<a href="/">← 返回签到首页</a>
<h1>月费查询</h1>

<form method="post">
    <label>月费编号</label>
    <input name="member_id" placeholder="例如：208 / CHE-208 / 0208" required>

    <label>PIN</label>
    <input name="pin" type="password" inputmode="numeric" required>

    <button type="submit">查询</button>
</form>

{% if error %}
<div class="error">{{ error }}</div>
{% endif %}

{% if result %}
<div class="result">
    姓名：<b>{{ result.name }}</b><br>
    英文名：{{ result.english_name or "-" }}<br>
    电话：{{ result.phone or "-" }}<br>
    月费编号：{{ result.member_id }}<br>
    已供养月份：{{ ", ".join(result.months) if result.months else "暂无记录" }}<br>
    已供养至：<b>{{ result.latest or "暂无记录" }}</b>
</div>
{% endif %}
</div>
</body>
</html>
""", result=result, error=error)


@app.route("/admin_report", methods=["POST"])
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

@app.route("/admin_records")
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

@app.route("/admin_edit_record/<int:record_id>", methods=["GET", "POST"])
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

        return redirect(url_for("admin_records", pin=pin))

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

@app.route("/admin_delete_record/<int:record_id>")
def admin_delete_record(record_id):
    pin = request.args.get("pin", "")

    if pin != ADMIN_PIN:
        return "无权限"

    db_query("""
        delete from attendance
        where id = %s
          and date = %s
    """, (record_id, now_date_str()))

    return redirect(url_for("admin_records", pin=pin))

@app.route("/api/volunteer/<volunteer_id>", methods=["POST"])
def api_volunteer(volunteer_id):
    try:
        v = find_volunteer(volunteer_id)
        if not v:
            return jsonify({"ok": False})

        pin = request.form.get("pin", "").strip()
        pin_ok = verify_pin_for_volunteer(v, pin) if pin else False

        safe_v = {
            "编号": v.get("编号", ""),
            "姓名": v.get("姓名", ""),
            "状态": v.get("状态", ""),
            "pin_ok": pin_ok,
        }

        if pin_ok:
            safe_v["电话号码"] = v.get("电话号码", "")
            safe_v["月费已供养至"] = get_member_paid_until(v.get("编号"))

        return jsonify({"ok": True, "volunteer": safe_v})

    except Exception as e:
        return jsonify({"ok": False, "error": str(e)})
    
@app.route("/member", methods=["GET", "POST"])
def member():
    if request.method == "POST":
        member_id = request.form.get("member_id")
        pin = request.form.get("pin")

        # 查人
        m = db_query("""
            select * from members where member_id = %s
        """, (member_id,), fetchone=True)

        if not m:
            return "❌ 找不到佛友"

        # PIN 验证
        real_pin = m.get("pin") or m.get("phone")[-4:]
        if str(pin) != str(real_pin):
            return "❌ PIN 错误"

        months, latest = get_member_payment(member_id)

        return f"""
        姓名：{m['name']}<br>
        已供养月份：{', '.join(months)}<br>
        已供养至：{latest}
        """

    return """
    <form method="post">
        编号：<input name="member_id"><br>
        PIN：<input name="pin"><br>
        <button>查询</button>
    </form>
    """

# =========================
# 8) 修改 PIN
# =========================
def change_pin(volunteer_id: str, old_pin: str, new_pin: str, confirm_pin: str):
    volunteer = find_volunteer(volunteer_id)
    if not volunteer:
        return False, "找不到编号"

    if not verify_pin_for_volunteer(volunteer, old_pin):
        return False, "旧 PIN 不正确"

    new_pin = str(new_pin or "").strip()
    confirm_pin = str(confirm_pin or "").strip()

    if not new_pin or len(new_pin) < 4:
        return False, "新 PIN 至少4位"

    if not new_pin.isdigit():
        return False, "新 PIN 只能用数字"

    if new_pin != confirm_pin:
        return False, "两次 PIN 不一致"

    db_query("""
        update volunteers
        set pin = %s
        where id = %s
    """, (new_pin, normalize_member_id(volunteer_id)))

    return True, "PIN 已更新"


@app.route("/change_pin", methods=["GET", "POST"])
def change_pin_page():
    if request.method == "POST":
        ok, msg = change_pin(
            request.form.get("id"),
            request.form.get("old"),
            request.form.get("new"),
            request.form.get("confirm")
        )
        flash(msg, "ok" if ok else "bad")
        return redirect(url_for("change_pin_page"))
    return render_template_string(PIN_PAGE, t=get_text())

if __name__ == "__main__":
    print("====================================")
    print("义工签到系统已启动")
    print(f"管理员 PIN：{ADMIN_PIN}")
    print("====================================")

    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)