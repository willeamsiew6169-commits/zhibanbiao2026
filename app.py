# app.py
# 蕉赖观音堂义工签到系统 v3
# 功能：
# 1) 编号查找、PIN验证、签到、签退
# 2) 今日记录修改/删除、自动备份、秒按缓存
# 3) 中文 / English 双语网页（Excel 内部仍保存中文）
# 4) 管理员 PIN 网页按钮生成月报/年报
# 5) 修改 PIN

from __future__ import annotations

import os
import sys
import shutil
import threading
import pandas as pd
import atexit
import subprocess
import socket
import qrcode
import base64

from io import BytesIO
from pypinyin import lazy_pinyin
from datetime import datetime, date
from pathlib import Path
from typing import Optional
from openpyxl import Workbook, load_workbook
from excel_style_utils import beautify_attendance_file
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

from flask import (
    Flask, request, redirect, url_for,
    render_template_string, flash, jsonify,
    make_response
)


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
        "admin_pin": "管理员 PIN",
        "generate_report": "生成月报 / 年报",
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
        "admin_pin": "Admin PIN",
        "generate_report": "Generate Monthly / Yearly Report",
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
    }
}


app = Flask(__name__)
app.secret_key = "change-this-simple-secret"

READING_FILE = "reading.xlsx"

# =========================
# 工具函数
# =========================
def get_today_code():
    today = date.today()
    day_index = today.toordinal() % len(TODAY_CODE_LIST)
    return TODAY_CODE_LIST[day_index]

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
    return datetime.now().strftime("%Y-%m-%d")

# 👉 从 attendance.xlsx 拿今天签到名单
def get_today_attendees():
    try:
        df = pd.read_excel("attendance.xlsx")
    except:
        return []

    if "日期" not in df.columns or "姓名" not in df.columns:
        return []

    today = get_today()
    df["日期"] = pd.to_datetime(df["日期"], errors="coerce").dt.strftime("%Y-%m-%d")

    # 👉 筛选：今天 + 已签到=1
    if "签到" in df.columns:
        df = df[(df["日期"] == today) & (df["签到"] == 1)]
    else:
        df = df[df["日期"] == today]

    return sorted(df["姓名"].dropna().unique().tolist())


# =========================
# 页面：共修记录
# =========================
@app.route("/reading", methods=["GET", "POST"])
def reading():
    ensure_reading_file()

    attendees = get_today_attendees()
    today = get_today()

    if request.method == "POST":
        topic = request.form.get("topic", "").strip()
        names = request.form.getlist("names")
        session = request.form.get("session", "").strip()
        extra_names = request.form.get("extra_names", "").strip()
        extra_text = request.form.get("extra_names", "").strip()

        extra_names = []
        if extra_text:
            for sep in ["，", "、", ",", "\n"]:
                extra_text = extra_text.replace(sep, " ")
            extra_names = [x.strip() for x in extra_text.split(" ") if x.strip()]

        if extra_names:
            extra_list = [x.strip() for x in extra_names.split(",") if x.strip()]
            names.extend(extra_list)

        if len(names) < 2:
            return "❌ 至少需要2位义工共修<br><a href='/reading'>返回</a>"

        if not topic:
            return "❌ 请输入主题<br><a href='/reading'>返回</a>"

        df = pd.read_excel(READING_FILE)

        for col in ["日期", "姓名", "主题", "时间"]:
            if col not in df.columns:
                df[col] = ""

        now_time = datetime.now().strftime("%I:%M %p").lstrip("0")
        new_rows = []

        today_attendees = set(get_today_attendees())

        for name in names:
            new_rows.append({
                "日期": today,
                "姓名": name,
                "身份": "义工",
                "主题": topic,
                "场次": session,
                "时间": now_time,
            })

        for name in extra_names:
            new_rows.append({
                "日期": today,
                "姓名": name,
                "身份": "佛友",
                "主题": topic,
                "场次": session,
                "时间": now_time,
            })

        if new_rows:
            df = pd.concat([df, pd.DataFrame(new_rows)], ignore_index=True)

            # ✅ reading.xlsx 自动按日期排序：最早 → 最近
            df["日期"] = pd.to_datetime(df["日期"], errors="coerce")
            df = df.sort_values(by="日期", ascending=True)
            df["日期"] = df["日期"].dt.strftime("%Y-%m-%d")

            df.to_excel(READING_FILE, index=False)
            beautify_reading_excel()

        return redirect(url_for("reading"))

    df = pd.read_excel(READING_FILE)
    for col in ["日期", "姓名", "身份", "主题", "场次", "时间"]:
        if col not in df.columns:
            df[col] = ""

    df["日期"] = df["日期"].astype(str)
    today_df = df[df["日期"] == today].copy()
    today_records = today_df.to_dict("records")

    today_summary = (
        today_df.groupby(["姓名", "身份"], as_index=False)
                .size()
                .rename(columns={"size": "共修次数"})
    )

    today_summary_records = today_summary.to_dict("records")

    topic_options = sorted(
        [x for x in df["主题"].dropna().astype(str).unique().tolist() if x.strip()]
    )

    html = """
    <!doctype html>
    <html>
    <head>
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <style>
            body {
                font-family: "Microsoft YaHei", Arial;
                background: #f6f7f2;
                padding: 18px;
                font-size: 20px;
            }
            .card {
                background: white;
                padding: 18px;
                border-radius: 16px;
                margin-bottom: 18px;
                box-shadow: 0 2px 8px #ccc;
            }
            input[type=text] {
                width: 95%;
                font-size: 22px;
                padding: 12px;
                border-radius: 10px;
                border: 1px solid #aaa;
            }
            label {
                display: block;
                padding: 8px;
                font-size: 21px;
            }
            button {
                font-size: 22px;
                padding: 12px 22px;
                border-radius: 12px;
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

    <a href="/"><button>⬅ 返回签到首页</button></a>

    <div class="card">
        <h2>{{ t.bhff_record_title }}</h2>

        <form method="post">
            <p>今日主题：</p>
            <input type="text" name="topic" list="topicList" placeholder="例如：佛陀的大智慧">

            <datalist id="topicList">
            {% for topic in topic_options %}
                <option value="{{ topic }}">
            {% endfor %}
            </datalist>

            <p>场次 / 备注：</p>
            <input type="text" name="session" placeholder="例如：早上共修 / 晚上共修">

            <p>共修人员：</p>

            <p>义工名单：</p>

            <button type="button" onclick="selectAllNames()" style="margin-bottom:10px;">
            ✅ 全选全部义工
            </button>

            {% for name in attendees %}
                <label>
                    <input type="checkbox" name="names" value="{{name}}">
                    {{name}}
                </label>
            {% endfor %}

            <br>
            <button type="submit">✅ 记录共修</button>

            <p>非义工佛友姓名：</p>
            <input type="text" name="extra_names" placeholder="例如：王小明、李美玲、陈先生"
                style="font-size:20px;width:420px;padding:8px;">
            <p style="color:#777;">多个名字用逗号、空格或顿号隔开</p>
        </form>
    </div>

    <div class="card">
        <h2>今日已记录</h2>

        <table>
            <tr>
                <th>姓名</th>
                <th>主题</th>
                <th>时间</th>
                <th>操作</th>
            </tr>

            {% for r in today_records %}
            <tr>
                <td>{{ r["姓名"] }}</td>
                <td>{{ r["主题"] }}</td>
                <td>{{ r["时间"] }}</td>
                <td>
                    <button>修改</button>
                    <button>删除</button>
                </td>
            </tr>
            {% endfor %}
        </table>

        <!-- ⭐ 加在这里 -->
        <h3 style="margin-top:20px;">📊 今日共修次数</h3>

        <table>
            <tr>
                <th>姓名</th>
                <th>身份</th>
                <th>次数</th>
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
    today_records=today_records,
    today_summary=today_summary_records,
    topic_options=topic_options
)

@app.route("/reading_delete")
def reading_delete():
    date = request.args.get("date", "")
    name = request.args.get("name", "")
    topic = request.args.get("topic", "")

    df = pd.read_excel(READING_FILE)

    mask = (
        (df["日期"].astype(str) == date) &
        (df["姓名"].astype(str) == name) &
        (df["主题"].astype(str) == topic)
    )

    df = df[~mask].copy()
    df.to_excel(READING_FILE, index=False)
    #beautify_reading_excel()

    return redirect(url_for("reading"))


@app.route("/reading_edit", methods=["GET", "POST"])
def reading_edit():
    old_date = request.args.get("date", "")
    old_name = request.args.get("name", "")
    old_topic = request.args.get("topic", "")

    df = pd.read_excel(READING_FILE)

    if request.method == "POST":
        new_topic = request.form.get("topic", "").strip()

        mask = (
            (df["日期"].astype(str) == old_date) &
            (df["姓名"].astype(str) == old_name) &
            (df["主题"].astype(str) == old_topic)
        )

        df.loc[mask, "主题"] = new_topic
        df.to_excel(READING_FILE, index=False)
        #beautify_reading_excel()

        return redirect(url_for("reading"))

    html = """
    <h2>修改白话佛法主题</h2>
    <form method="post">
        <p>姓名：{{name}}</p>
        <p>原主题：{{topic}}</p>
        <input name="topic" value="{{topic}}" style="font-size:22px;width:300px;">
        <br><br>
        <button type="submit" style="font-size:22px;">保存修改</button>
    </form>
    <br>
    <a href="/reading">返回</a>
    """

    return render_template_string(
        html,
        name=old_name,
        topic=old_topic
    )

# =========================
# 秒按模式：内存缓存
# =========================
ATT_CACHE = []
ATT_CACHE_LOADED = False
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
    return datetime.now().strftime("%Y-%m-%d")


def now_time_str() -> str:
    return datetime.now().strftime("%I:%M%p").lstrip("0").lower()


def normalize_id(value) -> str:
    if value is None:
        return ""
    s = str(value).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s.strip()


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
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        shutil.copy2(ATTENDANCE_FILE, BACKUP_DIR / f"attendance_{ts}.xlsx")


def safe_save_workbook(wb, path: Path) -> None:
    tmp_path = path.with_suffix(".tmp.xlsx")
    wb.save(tmp_path)
    os.replace(tmp_path, path)


def reload_attendance_cache() -> None:
    global ATT_CACHE, ATT_CACHE_LOADED, ATT_DIRTY
    with ATT_LOCK:
        ATT_CACHE = _load_attendance_rows_from_excel()
        ATT_CACHE_LOADED = True
        ATT_DIRTY = False


# =========================
# 4) Excel 读取 / 建立
# =========================
def ensure_attendance_file() -> None:
    if ATTENDANCE_FILE.exists():
        wb = load_workbook(ATTENDANCE_FILE)
        if ATTENDANCE_SHEET not in wb.sheetnames:
            ws = wb.create_sheet(ATTENDANCE_SHEET)
            ws.append(ATTENDANCE_HEADERS)
            safe_save_workbook(wb, ATTENDANCE_FILE)
        return

    wb = Workbook()
    ws = wb.active
    ws.title = ATTENDANCE_SHEET
    ws.append(ATTENDANCE_HEADERS)
    safe_save_workbook(wb, ATTENDANCE_FILE)


def load_volunteers() -> list[dict]:
    if not VOLUNTEERS_FILE.exists():
        raise FileNotFoundError(f"找不到 {VOLUNTEERS_FILE.name}")

    wb = load_workbook(VOLUNTEERS_FILE, data_only=True)
    if VOLUNTEERS_SHEET not in wb.sheetnames:
        raise ValueError(f"{VOLUNTEERS_FILE.name} 找不到工作表：{VOLUNTEERS_SHEET}")

    ws = wb[VOLUNTEERS_SHEET]
    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]

    def col(name: str) -> Optional[int]:
        return headers.index(name) + 1 if name in headers else None

    id_col = col("编号")
    name_col = col("姓名")
    phone_col = col("电话号码") or col("联络号码")
    status_col = col("状态")
    pin_col = col("PIN") or col("pin") or col("密码")

    if not id_col or not name_col:
        raise ValueError("volunteers.xlsx 至少需要栏位：编号、姓名")

    records = []
    for row in range(2, ws.max_row + 1):
        vid = normalize_id(ws.cell(row, id_col).value)
        name = str(ws.cell(row, name_col).value or "").strip()
        phone = str(ws.cell(row, phone_col).value or "").strip() if phone_col else ""
        status = str(ws.cell(row, status_col).value or "在册").strip() if status_col else "在册"
        pin_raw = str(ws.cell(row, pin_col).value or "").strip() if pin_col else ""

        if not vid or not name:
            continue

        phone_digits = only_digits(phone)
        default_pin = phone_digits[-4:] if len(phone_digits) >= 4 else "0000"
        pin = pin_raw if pin_raw else default_pin
        if pin.endswith(".0"):
            pin = pin[:-2]

        records.append({
            "编号": vid,
            "姓名": name,
            "电话号码": phone,
            "状态": status,
            "PIN": str(pin).strip(),
        })

    return records


def find_volunteer(volunteer_id: str) -> Optional[dict]:
    target = normalize_id(volunteer_id)
    volunteers = load_volunteers()

    for v in volunteers:
        if normalize_id(v["编号"]) == target:
            return v

    if target.isdigit():
        target_int = int(target)
        for v in volunteers:
            vid = normalize_id(v["编号"])
            if vid.isdigit() and int(vid) == target_int:
                return v
    return None


def verify_pin_for_volunteer(volunteer: dict, pin: str) -> bool:
    saved = str(volunteer.get("PIN") or "").strip()
    typed = str(pin or "").strip()
    return bool(saved) and saved == typed


def _load_attendance_rows_from_excel() -> list[dict]:
    ensure_attendance_file()
    wb = load_workbook(ATTENDANCE_FILE, data_only=True)
    ws = wb[ATTENDANCE_SHEET]
    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]

    rows = []
    for r in range(2, ws.max_row + 1):
        item = {headers[i]: ws.cell(r, i + 1).value for i in range(len(headers)) if headers[i]}
        item["_row"] = r
        rows.append(item)
    return rows


def load_attendance_rows() -> list[dict]:
    global ATT_CACHE_LOADED, ATT_CACHE
    with ATT_LOCK:
        if not ATT_CACHE_LOADED:
            ATT_CACHE = _load_attendance_rows_from_excel()
            ATT_CACHE_LOADED = True
        return [dict(r) for r in ATT_CACHE]


def mark_attendance_dirty() -> None:
    global ATT_DIRTY
    ATT_DIRTY = True


def flush_attendance_to_excel(force: bool = False) -> None:
    global ATT_DIRTY, ATT_CACHE_LOADED
    with ATT_LOCK:
        if not ATT_CACHE_LOADED:
            return
        if not ATT_DIRTY and not force:
            return
        rows = [dict(r) for r in ATT_CACHE]
        ATT_DIRTY = False

    ensure_attendance_file()
    wb = load_workbook(ATTENDANCE_FILE)
    ws = wb[ATTENDANCE_SHEET]
    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]

    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)

    for item in rows:
        ws.append([item.get(h, "") for h in headers])

    from openpyxl.styles import Font
    font = Font(name="Microsoft YaHei", size=14)

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font = font

    safe_save_workbook(wb, ATTENDANCE_FILE)
    beautify_attendance_file(ATTENDANCE_FILE, ATTENDANCE_SHEET)

    with ATT_LOCK:
        for idx, item in enumerate(ATT_CACHE, start=2):
            item["_row"] = idx


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
    open_rows = []
    for item in load_attendance_rows():
        d = format_date_value(item.get("日期"))
        end_time = str(item.get("结束时间") or "").strip()
        if d == today and end_time == "":
            open_rows.append(item)
    return open_rows


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
    volunteer_id = normalize_id(volunteer_id)
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

    with ATT_LOCK:
        global ATT_CACHE_LOADED, ATT_CACHE
        if not ATT_CACHE_LOADED:
            ATT_CACHE = _load_attendance_rows_from_excel()
            ATT_CACHE_LOADED = True

        for item in ATT_CACHE:
            d = format_date_value(item.get("日期"))
            end_time = str(item.get("结束时间") or "").strip()
            same_id = "编号" in item and normalize_id(item.get("编号")) == normalize_id(volunteer["编号"])
            same_name = str(item.get("姓名") or "").strip() == volunteer["姓名"]
            if d == now_date_str() and end_time == "" and (same_id or same_name):
                return False, f"{volunteer['姓名']} 今天已经签到，还没签退。请先签退。"

        next_row = max([int(r.get("_row", 1)) for r in ATT_CACHE] + [1]) + 1
        ATT_CACHE.append({
            "日期": now_date_str(),
            "编号": volunteer["编号"],
            "姓名": volunteer["姓名"],
            "报名": 0,
            "签到": 1,
            "岗位": role,
            "开始时间": now_time_str(),
            "结束时间": "",
            "时数": "",
            "备注": "iPad签到",
            "_row": next_row,
        })
        mark_attendance_dirty()

    return True, f"{volunteer['姓名']} 已签到：{role}"


def sign_out(row_number: int, pin: str) -> tuple[bool, str]:
    with ATT_LOCK:
        global ATT_CACHE_LOADED, ATT_CACHE
        if not ATT_CACHE_LOADED:
            ATT_CACHE = _load_attendance_rows_from_excel()
            ATT_CACHE_LOADED = True

        target = None
        for item in ATT_CACHE:
            if int(item.get("_row", 0)) == row_number and format_date_value(item.get("日期")) == now_date_str() and str(item.get("结束时间") or "").strip() == "":
                target = item
                break

        if not target:
            return False, "找不到这笔进行中的签到记录，可能已经签退了。"

        name = str(target.get("姓名") or "").strip()
        volunteer = None
        for v in load_volunteers():
            if v.get("姓名") == name:
                volunteer = v
                break

        if not volunteer:
            return False, "找不到此义工资料，无法验证 PIN。"

        if not verify_pin_for_volunteer(volunteer, pin):
            return False, "PIN 不正确，不能签退。"

        end = now_time_str()
        start = str(target.get("开始时间") or "").strip()
        hours = calc_hours(start, end)
        target["结束时间"] = end
        target["时数"] = hours
        role = target.get("岗位", "")
        mark_attendance_dirty()

    return True, f"{name} 已签退：{role}，共 {hours} 小时。"


def update_record(row_number: int, role: str, start_time: str, end_time: str, remark: str) -> tuple[bool, str]:
    flush_attendance_to_excel(force=True)
    if role not in ROLES:
        return False, "请选择正确岗位。"

    ensure_attendance_file()
    wb = load_workbook(ATTENDANCE_FILE)
    ws = wb[ATTENDANCE_SHEET]

    if row_number < 2 or row_number > ws.max_row:
        return False, "找不到这笔记录。"

    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]

    def has_col(name: str) -> bool:
        return name in headers

    def col(name: str) -> int:
        return headers.index(name) + 1

    today = now_date_str()
    d = format_date_value(ws.cell(row_number, col("日期")).value if has_col("日期") else "")
    if d != today:
        return False, "为了安全，页面只允许修改今天的记录。"

    backup_attendance()
    ws.cell(row_number, col("岗位")).value = role
    ws.cell(row_number, col("开始时间")).value = start_time.strip()
    ws.cell(row_number, col("结束时间")).value = end_time.strip()
    if has_col("备注"):
        ws.cell(row_number, col("备注")).value = remark.strip()

    if has_col("时数"):
        if start_time.strip() and end_time.strip():
            ws.cell(row_number, col("时数")).value = calc_hours(start_time.strip(), end_time.strip())
        else:
            ws.cell(row_number, col("时数")).value = ""

    safe_save_workbook(wb, ATTENDANCE_FILE)
    beautify_attendance_file(ATTENDANCE_FILE, ATTENDANCE_SHEET)
    reload_attendance_cache()
    return True, "记录已修改。"


def delete_record(row_number: int) -> tuple[bool, str]:
    flush_attendance_to_excel(force=True)
    ensure_attendance_file()
    wb = load_workbook(ATTENDANCE_FILE)
    ws = wb[ATTENDANCE_SHEET]

    if row_number < 2 or row_number > ws.max_row:
        return False, "找不到这笔记录。"

    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]

    def col(name: str) -> int:
        return headers.index(name) + 1

    today = now_date_str()
    d = format_date_value(ws.cell(row_number, col("日期")).value)
    if d != today:
        return False, "为了安全，页面只允许删除今天的记录。"

    name = ws.cell(row_number, col("姓名")).value if "姓名" in headers else ""
    backup_attendance()
    ws.delete_rows(row_number, 1)
    safe_save_workbook(wb, ATTENDANCE_FILE)
    beautify_attendance_file(ATTENDANCE_FILE, ATTENDANCE_SHEET)
    reload_attendance_cache()
    return True, f"已删除 {name} 的这笔记录。"


def run_report_script() -> tuple[bool, str]:
    flush_attendance_to_excel(force=True)

    if not REPORT_SCRIPT.exists():
        return False, f"找不到报表脚本：{REPORT_SCRIPT.name}"

    try:
        result = subprocess.run(
            [sys.executable, str(REPORT_SCRIPT)],
            cwd=str(BASE_DIR),
            capture_output=True,
            text=True,
            timeout=300
        )
        if result.returncode != 0:
            err = result.stderr.strip() or result.stdout.strip() or "未知错误"
            return False, "生成报表失败：" + err[-500:]
        return True, "报表已生成。"
    except subprocess.TimeoutExpired:
        return False, "生成报表超时，请用 BAT 手动生成。"
    except Exception as e:
        return False, f"生成报表失败：{e}"


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
    body { font-family: -apple-system, BlinkMacSystemFont, "Microsoft YaHei", Arial, sans-serif; background:#f6f6f6; margin:0; }
    .wrap { max-width: 900px; margin: 0 auto; padding: 18px; }
    .card { background:white; border-radius:18px; padding:20px; margin-bottom:16px; box-shadow:0 2px 12px rgba(0,0,0,.08); }
    h1 { font-size: 30px; margin: 6px 0 18px; }
    h2 { font-size: 24px; margin: 0 0 14px; }
    label { font-size: 20px; font-weight: 700; display:block; margin-bottom:8px; }
    input, select { width:100%; font-size: 26px; padding: 14px; border-radius: 12px; border:1px solid #ccc; box-sizing:border-box; }
    .row { display:grid; grid-template-columns: 1fr 1fr; gap:14px; }
    button { font-size: 24px; font-weight: 800; padding: 14px 18px; border:0; border-radius: 14px; cursor:pointer; }
    .btn-find { background:#0d6efd; color:white; width:100%; margin-top:16px; }
    .btn-in { background:#198754; color:white; width:100%; margin-top:16px; }
    .btn-out { background:#dc3545; color:white; }
    .btn-admin { background:#6f42c1; color:white; width:100%; margin-top:16px; }
    .btn-lang { display:inline-block; font-size:18px; padding:8px 12px; background:white; border-radius:999px; text-decoration:none; color:#333; border:1px solid #ddd; margin-right:6px; }
    .msg { padding:14px; border-radius:12px; font-size:20px; margin-bottom:14px; white-space:pre-wrap; }
    .ok { background:#d1e7dd; color:#0f5132; }
    .bad { background:#f8d7da; color:#842029; }
    .person { font-size: 22px; line-height:1.7; background:#f8f9fa; padding:12px; border-radius:12px; margin-top:12px; }
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
        <label>今日签到码</label>
        <input type="text" name="today_code" placeholder="请输入现场今日码" required>
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

  const res = await fetch('/api/volunteer/' + encodeURIComponent(id));
  const data = await res.json();
  box.style.display = 'block';

  if (data.ok) {
    currentVolunteerName = data.volunteer.姓名;
    box.innerHTML =
      `${TXT.name}：<b>${data.volunteer.姓名}</b><br>` +
      `${TXT.phone}：${data.volunteer.电话号码 || '-'}<br>` +
      `${TXT.status}：${data.volunteer.状态 || '-'}`;
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
        real_code = get_today_code()

        if input_code != real_code:
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
    ok, msg = sign_out(row_number, request.form.get("pin", ""))
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


@app.route("/admin_report", methods=["POST"])
def admin_report():
    pin = str(request.form.get("admin_pin", "")).strip()
    if pin != ADMIN_PIN:
        flash("管理员 PIN 不正确。", "bad")
        return redirect(url_for("index"))

    ok, msg = run_report_script()
    flash(msg, "ok" if ok else "bad")
    return redirect(url_for("index"))


@app.route("/api/volunteer/<volunteer_id>")
def api_volunteer(volunteer_id):
    try:
        v = find_volunteer(volunteer_id)
        if not v:
            return jsonify({"ok": False})
        safe_v = {k: v.get(k, "") for k in ["编号", "姓名", "电话号码", "状态"]}
        return jsonify({"ok": True, "volunteer": safe_v})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)})


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

    wb = load_workbook(VOLUNTEERS_FILE)
    ws = wb[VOLUNTEERS_SHEET]
    headers = [str(c.value).strip() if c.value else "" for c in ws[1]]

    def col(name):
        return headers.index(name) + 1 if name in headers else None

    id_col = col("编号")
    pin_col = col("PIN")

    if not pin_col:
        ws.cell(1, ws.max_column + 1).value = "PIN"
        pin_col = ws.max_column

    for r in range(2, ws.max_row + 1):
        vid = normalize_id(ws.cell(r, id_col).value)
        if vid == normalize_id(volunteer_id):
            ws.cell(r, pin_col).value = new_pin
            safe_save_workbook(wb, VOLUNTEERS_FILE)
            return True, "PIN 已更新"

    return False, "更新失败"


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


# =========================
# 9) 启动
# =========================
if __name__ == "__main__":
    ensure_attendance_file()
    load_attendance_rows()
    start_background_saver_once()
    atexit.register(lambda: flush_attendance_to_excel(force=True))

    import os

if __name__ == "__main__":
    print("====================================")
    print("义工签到系统已启动")
    print(f"管理员 PIN：{ADMIN_PIN}")
    print("====================================")

    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)