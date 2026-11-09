# app.py
# 蕉赖共修会义工签到系统 v3
# 功能：
# 1) 编号查找、PIN验证、签到、签退
# 2) 今日记录修改/删除、自动备份、秒按缓存
# 3) 中文 / English 双语网页（Excel 内部仍保存中文）
# 4) 输入管理员 PIN 
# 5) 修改 PIN

from __future__ import annotations

import os
import time
import io
import sys
import psycopg2
import pandas as pd

from pathlib import Path
from typing import Optional
from zoneinfo import ZoneInfo
from admin_web import admin_bp
from member_web import member_bp
from people_web import people_bp
from pypinyin import lazy_pinyin
from manifest import manifest_bp
from db import db_query, get_conn
from library_web import library_bp
from reading_web import reading_bp
from xfz_web import xfz_bp

from finance_v7 import finance_v7_bp
from finance_web import finance_bp
import finance_export

from psycopg2.extras import RealDictCursor
from finance_import import finance_import_bp
from lunar_rules import get_special_day_info
from dharma_class_web import dharma_class_bp
from schedule.schedule_web import schedule_bp
from datetime import datetime, date, timedelta
from finance_engine import finance_engine_api_bp
from finance_month_end import finance_month_end_bp
from finance_audit import finance_audit_bp
from flask import (
    Flask, request, redirect, url_for,
    render_template_string, flash, jsonify, 
    make_response, send_file, send_from_directory,
)

from attendance_service import (
    ROLE_TEXT,
    ROLES,
    role_label,
    sign_in,
    sign_out,
    find_volunteer,
    verify_pin_for_volunteer,
    load_attendance_rows,
    get_today_open_records,
    get_today_records,
    only_digits,
    get_assignment_id_candidates,
    get_today_assignments,
    ENABLE_SIGNIN_TIME_LIMIT,
    SIGNIN_EARLY_MINUTES,
    REQUIRE_ASSIGNMENT_FOR_SIGNIN,
    ENABLE_AUTO_SIGNOUT,
    AUTO_SIGNOUT_TIME,
    AUTO_SIGNOUT_DISPLAY,
    auto_signout_unfinished_today,
    get_eligible_signin_roles,
)

from utils import (
    MY_TZ,
    now_date_str,
    now_time_str,
    parse_time,
    parse_time_to_datetime,
    calc_hours,
    get_text,
    get_lang,
    normalize_member_id,
    TEXT,
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


# 你的月报/年报脚本名称
REPORT_SCRIPT = BASE_DIR / "zhibanbiao2026.py"

# 保持你旧 attendance.xlsx 格式，避免月报错位
ATTENDANCE_HEADERS = [
    "日期", "姓名", "报名", "签到", "岗位",
    "开始时间", "结束时间", "时数", "备注"
]


app = Flask(__name__)
app.secret_key = "change-this-simple-secret"
app.permanent_session_lifetime = timedelta(hours=2)
app.register_blueprint(finance_engine_api_bp)
app.register_blueprint(finance_month_end_bp)
app.register_blueprint(finance_import_bp)
app.register_blueprint(finance_audit_bp)
app.register_blueprint(dharma_class_bp)
app.register_blueprint(finance_v7_bp)
app.register_blueprint(schedule_bp)
app.register_blueprint(manifest_bp)
app.register_blueprint(finance_bp)
app.register_blueprint(reading_bp)
app.register_blueprint(library_bp)
app.register_blueprint(member_bp)
app.register_blueprint(people_bp)
app.register_blueprint(admin_bp)
app.register_blueprint(xfz_bp)


def get_today_stats():
    """在数据库内汇总，避免将当天所有签到记录传回 Python。"""
    row = db_query("""
        SELECT COUNT(*) AS total,
               COUNT(*) FILTER (WHERE end_time IS NULL OR end_time = '') AS open
        FROM attendance
        WHERE date = %s
    """, (now_date_str(),), fetchone=True) or {}
    total = int(row.get("total") or 0)
    open_count = int(row.get("open") or 0)
    return {"total": total, "open": open_count, "finished": total - open_count}


# =========================
# 2) 语言工具
# =========================
@app.route("/set_lang/<lang>")
def set_lang(lang):
    if lang not in TEXT:
        lang = "zh"
    resp = make_response(redirect(request.referrer or url_for("index")))
    resp.set_cookie("lang", lang, max_age=60 * 60 * 24 * 365)
    return resp


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

def to_member_id(volunteer):
    vol_id = str(volunteer.get("编号") or "").strip()
    return vol_id

def get_member_paid_until(member_id):
    row = db_query("""
        select max(end_month) as paid_until
        from member_payments
        where member_id = %s
    """, (member_id,), fetchone=True)

    return row.get("paid_until") if row else ""


def format_date_value(d) -> str:
    if isinstance(d, datetime):
        return d.strftime("%Y-%m-%d")
    if isinstance(d, date):
        return d.strftime("%Y-%m-%d")
    return str(d or "").strip()

def check_in_assignment(assignment):

    db_query("""
        insert into volunteer_attendance_logs (
            assignment_id,
            volunteer_id,
            name,
            attendance_date,
            actual_role,
            actual_place,
            walk_in,
            remarks
        )
        values (%s, %s, %s, %s, %s, %s, false, %s)
    """, (
        assignment["id"],
        assignment["volunteer_id"],
        assignment["name"],
        assignment["assignment_date"],
        assignment["role"],
        assignment["assigned_place"],
        "按排班签到"
    ))


def load_today_assignments_for_volunteer(volunteer, raw_id):
    ids = get_assignment_id_candidates(volunteer, raw_id)

    if not ids:
        return []

    return db_query("""
        select
            id,
            volunteer_id,
            name,
            assignment_date,
            role,
            shift_label,
            assigned_place,
            start_time,
            end_time
        from volunteer_schedule_assignments
        where volunteer_id = any(%s)
        and assignment_date = %s
        and coalesce(status, 'assigned') <> 'cancelled'
        order by start_time, role, assigned_place
    """, (ids, now_date_str()), fetchall=True) or []


def format_assignment_for_attendance(a):
    role = a.get("role")
    place = a.get("assigned_place") or ""
    shift = a.get("shift_label") or ""
    start = a.get("start_time") or ""
    end = a.get("end_time") or ""

    if role == "值班":
        time_text = f"{start} ~ {end}" if start and end else ""
        return f"{time_text} {shift} {place}".strip()

    return place or role




def update_record(row_number: int, role: str, card_no: str, start_time: str, end_time: str, remark: str) -> tuple[bool, str]:
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

    card_no = str(card_no or "").strip()
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
            card_no = %s,
            start_time = %s,
            end_time = %s,
            hours = %s,
            remark = %s
        where id = %s
    """, (
        role,
        card_no,
        start_time,
        end_time,
        hours,
        remark,
        row_number
    ))

    return True, "记录已修改。"


def delete_record(row_number: int, pin: str) -> tuple[bool, str]:
    pin = (pin or "").strip()

    row = db_query("""
        select *
        from attendance
        where id = %s
    """, (row_number,), fetchone=True)

    if not row:
        return False, "找不到这笔记录。"

    if row.get("date") != now_date_str():
        return False, "为了安全，页面只允许删除今天的记录。"

    volunteer_id = (row.get("volunteer_id") or "").strip()
    name = row.get("name") or ""

    if not volunteer_id:
        return False, "这笔记录没有义工编号，不能用本人 PIN 删除。"

    vol = db_query("""
        select pin, phone
        from volunteers
        where id = %s
    """, (volunteer_id,), fetchone=True)

    if not vol:
        return False, "找不到这位义工资料，不能删除。"

    correct_pin = (vol.get("pin") or "").strip()

    if not correct_pin:
        phone = (vol.get("phone") or "").strip()
        correct_pin = phone[-4:] if len(phone) >= 4 else ""

    if not correct_pin:
        return False, "这位义工没有 PIN，也没有电话号码后4位，不能删除。"

    if pin != correct_pin:
        return False, "PIN 不正确，不能删除。"

    db_query("""
        delete from attendance
        where id = %s
    """, (row_number,))

    return True, f"已删除 {name} 的这笔记录。"


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

  <link rel="stylesheet" href="{{ url_for('static', filename='css/toolbox.css') }}">

  <style>
    .center { text-align:center; }
    .small-note { font-size:16px; color:#666; margin-top:10px; text-align:center; }
    .two-card-grid {
      display:grid;
      grid-template-columns:1fr 1fr;
      gap:18px;
    }
    .branch-row {
        display: grid;
        grid-template-columns: 72px 1fr;
        gap: 10px;
        align-items: end;
    }
    .branch-btn {
      width:100px;
      flex-shrink:0;
    }
    .footer {
      text-align:center;
      color:#777;
      font-size:16px;
      margin:20px 0;
    }
    .sign-page .btn-row {
    display: flex;
    gap: 16px;
    width: 100%;
    }

    .sign-page .btn-row .btn-tool {
    flex: 1;
    width: 100%;
    justify-content: center;
    }

    .sign-page .single-btn-row .btn-tool {
    max-width: none;
    width: 100%;
    }

    .sign-page .small-action-btn {
    max-width: 260px;
    margin: 0 auto;
    }
    .login-row {
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 18px;
        align-items: end;
    }

    .form-two-col {
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 12px;
        align-items: end;
    }

    .branch-input-row {
        display: flex;
        align-items: stretch;
        gap: 8px;
    }

    .branch-btn {
        width: 72px;
        height: 72px;
        padding: 0;
        font-size: 20px;
    }

    .branch-row .form-input,
    .login-row > .form-group > .form-input {
        height: 72px;
        box-sizing: border-box;
    }

    .branch-input-row .form-input,
    .form-two-col .form-input {
        height: 44px;
        box-sizing: border-box;
    }

    @media (max-width: 700px) {
        .form-two-col {
            grid-template-columns: 1fr;
        }
    }

    @media (max-width: 700px) {
    .login-row {
        grid-template-columns: 1fr;
    }
    }
    @media (max-width:700px) {
      .two-card-grid {
        grid-template-columns:1fr;
      }
    }

    .hero-card{
      background: linear-gradient(135deg, #6d5dfc, #8f6ff7, #ffffff);
      color:white;
      border:none;
      overflow:hidden;
      position:relative;
  }

  .hero-card::after{
      content:"🪷";
      position:absolute;
      right:28px;
      top:22px;
      font-size:92px;
      opacity:.12;
  }

  .hero-icon{
      font-size:46px;
      margin-bottom:8px;
  }

  .hero-title{
      font-size:34px;
      font-weight:900;
      margin:6px 0;
      color:white;
  }

  .hero-subtitle{
      font-size:17px;
      line-height:1.6;
      opacity:.95;
  }

  .hero-links{
      margin-top:16px;
      font-size:15px;
  }

  .hero-links a{
      color:white;
      font-weight:700;
      text-decoration:none;
  }

  .hero-date{
      margin-top:18px;
      background:rgba(255,255,255,.20);
      border-radius:16px;
      padding:12px;
      font-size:17px;
      font-weight:700;
  }

  .stat-box-1{
      background:#eefdf4;
  }

  .stat-box-2{
      background:#fff8e8;
  }

  .stat-box-3{
      background:#eef4ff;
  }

  .summary-box{
      border:none;
  }

  .stats-header{
      display:flex;
      justify-content:space-between;
      align-items:flex-start;
      gap:24px;
      margin-bottom:20px;
  }

  .today-info{
      text-align:right;
      font-size:15px;
      color:#666;
      line-height:1.7;
  }

  .festival-text{
      color:#8c3eff;
      font-weight:700;
  }

  @media(max-width:700px){

      .stats-header{
          flex-direction:column;
          gap:10px;
      }

      .today-info{
          text-align:left;
      }

  }
  </style>
</head>

<body>

<div class="page sign-page">

  <div class="card center hero-card">
      <div class="hero-icon">🌸</div>

      <div class="hero-title">
          {{ t.system_title }}
      </div>

      <div class="hero-subtitle">
          {{ t.header_subtitle_1 }}
          {{ t.header_subtitle_2 }}
      </div>

      <div class="hero-date">
          🙏 {{ t.welcome_home }}
      </div>

      <div class="hero-links">
          {{ t.language }}：
          <a href="{{ url_for('set_lang', lang='zh') }}">{{ t.chinese }}</a>
          |
          <a href="{{ url_for('set_lang', lang='en') }}">{{ t.english }}</a>
          |
          <a href="{{ url_for('change_pin_page') }}">{{ t.change_pin }}</a>
      </div>
  </div>

  <div class="card">
    <div class="stats-header">

      <h2 class="section-title">
          📊 {{ t.today_stats }}
      </h2>

      <div class="today-info">

          <div>
              📅 {{ today_text }}
          </div>

          <div>
              🌙 {{ lunar_text }}
          </div>

          {% if buddhist_day %}
          <div class="festival-text">
              🪷 {{ buddhist_day }}
          </div>
          {% endif %}

          <div>
          🕘 <span id="clock"></span>
          </div>

      </div>

  </div>

    <div class="summary-grid">
      <div class="summary-box stat-box-1">
        <div class="summary-title">👥 {{ t.checked_in }}</div>
        <div class="summary-value">{{ today_count }}</div>
      </div>

      <div class="summary-box stat-box-2">
        <div class="summary-title">🕒 {{ t.on_duty }}</div>
        <div class="summary-value">{{ not_out }}</div>
      </div>

      <div class="summary-box stat-box-3">
        <div class="summary-title">✅ {{ t.checked_out }}</div>
        <div class="summary-value">{{ done_out }}</div>
      </div>
    </div>
  </div>

  {% with messages = get_flashed_messages(with_categories=true) %}
    {% for category, message in messages %}
      <div class="alert {{ 'alert-success' if category == 'ok' else 'alert-danger' }}">
        {{ message }}
      </div>
    {% endfor %}
  {% endwith %}

  <div class="card">
    <h2 class="section-title">✅ {{ t.check_in }}</h2>

    <form method="post"
          action="{{ url_for('do_sign_in') }}"
          onsubmit="return quickSignIn();">

      <input type="hidden" id="assignment_id" name="assignment_id" value="">

      <div class="login-row">

        <div class="form-group">
          <label class="form-label">{{ t.enter_id }}</label>

          <div class="branch-row">
            <button
              type="button"
              id="branch_btn"
              onclick="toggleBranch()"
              class="btn-tool btn-success branch-btn">
              CHE
            </button>

            <input type="hidden" id="branch" name="branch" value="CHE">

            <input
              class="form-input"
              id="volunteer_id"
              name="volunteer_id"
              inputmode="numeric"
              autocomplete="off"
              placeholder="{{ t.id_placeholder }}"
              required
            >
          </div>
        </div>

        <div class="form-group">
          <label class="form-label">{{ t.pin }}</label>
          <input
            class="form-input"
            id="pin"
            name="pin"
            type="password"
            inputmode="numeric"
            pattern="[0-9]*"
            autocomplete="new-password"
            autocorrect="off"
            autocapitalize="off"
            spellcheck="false"
            value=""
            placeholder="{{ t.pin_placeholder }}"
            readonly
            onfocus="this.removeAttribute('readonly');"
            required
          >
        </div>

      </div>
      <div class="btn-row single-btn-row">
        <button type="button"
                class="btn-tool btn-primary"
                onclick="lookupVolunteer()">
            {{ t.search_volunteer }}
        </button>
        </div>

      <div id="personBox"
           class="alert alert-success"
           style="display:none;">
      </div>

      <div class="form-row" id="signinOnlyFields">

        <div class="form-group form-col">

            <label class="form-label">
                {{ t.role }}
            </label>

            <select
                class="form-input"
                name="role"
                id="role"
                onchange="toggleCardNo(); refreshRoleAction()">

                <option value="">{{ t.lookup_placeholder }}</option>

            </select>

        </div>

        <div
            class="form-group form-col"
            id="card_no_group">

            <label class="form-label">
                {{ t.card_no_duty_only }}
            </label>

            <input
                class="form-input"
                type="text"
                name="card_no"
                id="card_no"
                placeholder="{{ t.card_no_duty_placeholder }}">

        </div>

    </div>

      <div class="form-group" id="signoutRoleGroup" style="display:none;margin:12px 0;">
        <label class="form-label">{{ '选择签退岗位' if t.html_lang.startswith('zh') else 'Select role to check out' }}</label>
        <select id="signoutRole" class="form-input" onchange="refreshRoleAction()"></select>
      </div>
      <div class="btn-row">
        <button id="signInBtn"
                class="btn-tool btn-success"
                type="submit"
                style="display:none;">
            {{ t.confirm_checkin }}
        </button>

        <button id="signOutBtn"
                class="btn-tool btn-danger"
                type="button"
                style="display:none;"
                onclick="submitQuickSignOut();">
            {{ t.confirm_checkout }}
        </button>
      </div>

    </form>

    <form id="quickSignOutForm"
          method="post"
          action="{{ url_for('do_sign_out') }}">
      <input type="hidden" id="signout_volunteer_id" name="volunteer_id">
      <input type="hidden" id="signout_pin" name="pin">
      <input type="hidden" id="signout_assignment_id" name="assignment_id">
      <input type="hidden" id="signout_role" name="role">
    </form>

    <div class="small-note" id="signinOnlyTip">
      💡 {{ t.card_no_duty_tip }}。
    </div>
  </div>

    <div class="card" style="margin-bottom:14px;padding:12px 18px;text-align:right">




      <div class="btn-row single-btn-row">
        <button
            type="button"
            class="btn-tool btn-purple" style="width:auto;padding:10px 16px"
            onclick="toggleAdminBox()">
            {{ t.admin_entry }}
        </button>
      </div>

      <div id="adminBox" style="display:none; margin-top:16px;">

        <form method="post" action="{{ url_for('admin.admin_report') }}">

          <div class="form-group">
            <label class="form-label">{{ t.admin_pin }}</label>
            <input
              class="form-input"
              id="admin_pin"
              name="admin_pin"
              type="password"
              inputmode="numeric"
              pattern="[0-9]*"
              autocomplete="new-password"
              autocorrect="off"
              autocapitalize="off"
              spellcheck="false"
              value=""
              placeholder="{{ t.pin_placeholder }}"
              readonly
              onfocus="this.removeAttribute('readonly');"
              required
            >
          </div>

          <div class="btn-row">
            <button class="btn-tool btn-purple" type="submit">
              {{ t.admin_login }}
            </button>
          </div>

        </form>

      </div>


    </div>

  <div class="two-card-grid">

    <div class="card center">
      <h2 class="section-title">📖 {{ t.bhff_title }}</h2>

      <div class="page-subtitle">
        {{ t.study_records_desc }}
      </div>

      <div class="btn-row single-btn-row">
        <a class="btn-tool btn-primary small-action-btn" href="/reading">
            {{ t.bhff_enter }}
        </a>
      </div>

      <div class="small-note">
        👉 {{ t.bhff_desc }}
      </div>
    </div>

    <div class="card center">
      <h2 class="section-title">📜 {{ '经文组合记录' if t.html_lang.startswith('zh') else 'XFZ Records' }}</h2>
      <div class="page-subtitle">{{ '结缘、烧送与库存管理' if t.html_lang.startswith('zh') else 'Offerings, burning & balance' }}</div>
      <div class="btn-row single-btn-row"><a class="btn-tool small-action-btn" style="background:#f4c84a;color:#3d3000;border:0" href="/xfz">{{ '进入经文组合记录' if t.html_lang.startswith('zh') else 'Enter XFZ Records' }}</a></div>
    </div>

  </div>

  <div class="footer">
    © {{ t.footer_text }} 💚
  </div>

</div>

<script>
const ROLE_LABELS = {{ role_text_labels|tojson }};
window.ROLE_LABELS = ROLE_LABELS;
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
  lookup_placeholder: {{ t.lookup_placeholder|tojson }},
  no_eligible_role: {{ t.no_eligible_role|tojson }},
  verify_pin_first: {{ t.verify_pin_first|tojson }},
  today_duties: {{ t.today_duties|tojson }},
  duty_completed: {{ t.duty_completed|tojson }},
  duty_auto_out: {{ t.duty_auto_out|tojson }},
  duty_active: {{ t.duty_active|tojson }},
  duty_pending: {{ t.duty_pending|tojson }},
  duty_place: {{ t.duty_place|tojson }},
  duty_role: {{ t.duty_role|tojson }},
  duty_time: {{ t.duty_time|tojson }},
  no_formal_duty: {{ t.no_formal_duty|tojson }},
  currently_checked_in: {{ t.currently_checked_in|tojson }},
  currently_can_checkin: {{ t.currently_can_checkin|tojson }},
  formal_duty_finished: {{ t.formal_duty_finished|tojson }},
  pin_search_again: {{ t.pin_search_again|tojson }},
  lookup_required_alert: {{ t.lookup_required_alert|tojson }},
  confirm_checkin: {{ t.confirm_checkin|tojson }},
  confirm_checkout: {{ t.confirm_checkout|tojson }},
  welcome_home: {{ t.welcome_home|tojson }},
  card_no_tip: {{ t.card_no_duty_tip|tojson }}
};

let verifiedLookupKey = null;
let lookupRequestSerial = 0;
function currentLookupKey() {
  return JSON.stringify([
    document.getElementById('branch').value,
    document.getElementById('volunteer_id').value.trim(),
    document.getElementById('pin').value.trim()
  ]);
}
function invalidateLookup() {
  verifiedLookupKey = null;
  window.lastAttendanceLookup = null;
  document.getElementById('signInBtn').style.display = 'none';
  document.getElementById('signOutBtn').style.display = 'none';
  document.getElementById('signoutRoleGroup').style.display = 'none';
  document.getElementById('assignment_id').value = '';
  document.getElementById('personBox').style.display = 'none';
  const roleSelect = document.getElementById('role');
  roleSelect.innerHTML = '';
  roleSelect.add(new Option(TXT.lookup_placeholder, ''));
}
async function lookupVolunteer() {
  invalidateLookup();
  const requestSerial = ++lookupRequestSerial;
  const lookupKey = currentLookupKey();
  const id = document.getElementById('volunteer_id').value.trim();
  const branch = document.getElementById('branch').value;
  let finalId;

  if (branch === "STW") {
    finalId = "STW-" + id;
  } else {
    finalId = id;
  }

  const box = document.getElementById('personBox');
  const btn = document.getElementById('signInBtn');

  if (!id) {
    box.style.display = 'block';
    box.className = 'alert alert-danger';
    box.innerHTML = TXT.enter_id_first;
    return;
  }

  const pin = document.getElementById('pin').value.trim();

  const formData = new FormData();
  formData.append('pin', pin);

  const res = await fetch('/api/volunteer/' + encodeURIComponent(finalId), {
    method: 'POST',
    body: formData
  });

  const data = await res.json();
  if (currentLookupKey() !== lookupKey || requestSerial !== lookupRequestSerial) return;
  box.style.display = 'block';

  if (data.ok) {

    box.className = 'alert alert-success';

    let html =
  `${TXT.name}：<b>${data.volunteer.姓名}</b>`;

    if (data.volunteer.pin_ok) {
       const phoneDigits = String(data.volunteer.电话号码 || '').replace(/\D/g, '');
       const displayPhone = /^0\d{8,10}$/.test(phoneDigits)
         ? phoneDigits.slice(0, 3) + '-' + phoneDigits.slice(3)
         : (data.volunteer.电话号码 || '-');
       html += `<br>${TXT.phone}：${displayPhone}`;

      const paidUntil = data.volunteer["月费已缴费至"];

      if (paidUntil && paidUntil !== "-") {
        html += `<br>${TXT.paid_until}：${paidUntil}`;
      }
    } else if (pin) {
      html += `<br><span style="color:#842029;">${TXT.pin_wrong}</span>`;
    }

    if (data.today_assignments && data.today_assignments.length > 0) {
      html += `<br><hr><b>${TXT.today_duties}</b>`;
      data.today_assignments.forEach(a => {
        const mark = a.state === "completed" ? (a.auto_signed_out ? "⚠️ " + TXT.duty_auto_out : "✅ " + TXT.duty_completed) :
                     a.state === "active" ? "🟢 " + TXT.duty_active : "🕒 " + TXT.duty_pending;
        html += `
          <div style="margin-top:10px;">
            ${mark}<br>
            ${TXT.duty_place}：<b>${a.assigned_place || "-"}</b><br>
            ${TXT.duty_role}：${(window.ROLE_LABELS || {})[a.role] || a.role || "-"}<br>
            ${TXT.duty_time}：${a.start_time || "-"} ～ ${a.end_time || "-"}
          </div>`;
      });
    } else {
      html += `<br><hr><b style="color:#b02a37;">${TXT.no_formal_duty}</b>`;
    }

    const signInBtn = document.getElementById('signInBtn');
    const signOutBtn = document.getElementById('signOutBtn');
    const assignmentInput = document.getElementById('assignment_id');
    signInBtn.style.display = 'none';
    signOutBtn.style.display = 'none';
    assignmentInput.value = '';

    window.lastAttendanceLookup = data;
    const roleSelect = document.getElementById('role');
    roleSelect.innerHTML = '';
    if (pin && data.volunteer.pin_ok) {
      (data.eligible_roles || []).forEach(item => {
        const option = document.createElement('option');
        option.value = item.role;
        option.textContent = item.label;
        roleSelect.appendChild(option);
      });
    }
    if (!roleSelect.options.length) {
      const option = document.createElement('option');
      option.value = '';
      option.textContent = (pin && data.volunteer.pin_ok)
        ? TXT.no_eligible_role
        : TXT.verify_pin_first;
      roleSelect.appendChild(option);
    }
    toggleCardNo();
    // 空 PIN 仅查看排班；填写 PIN 时必须验证正确才能操作。
    verifiedLookupKey = (pin && data.volunteer.pin_ok) ? lookupKey : null;
     const openRoles = data.open_roles || [];
     const outSelect = document.getElementById('signoutRole');
     outSelect.innerHTML = '';
     openRoles.forEach(r => outSelect.add(new Option(ROLE_LABELS[r] || r, r)));
     document.getElementById('signoutRoleGroup').style.display = openRoles.length ? '' : 'none';
     updateSigninOnlyVisibility(false);
    if (data.attendance_action === 'signout') {
      signOutBtn.style.display = '';
      assignmentInput.value = data.current_assignment_id || '';
      html += `<br><hr><b style="color:#198754;">${TXT.currently_checked_in}</b>`;
    } else if (data.attendance_action === 'signin') {
      signInBtn.style.display = '';
      assignmentInput.value = data.current_assignment_id || '';
      html += `<br><hr><b style="color:#0d6efd;">${TXT.currently_can_checkin}：${data.current_assignment_text || ''}</b>`;
    } else if (data.attendance_action === 'completed') {
      html += `<br><hr><b style="color:#198754;">✅ ${TXT.formal_duty_finished}</b>`;
    } else {
      // 试用期：没有正式排班时仍保留现场签到。
      signInBtn.style.display = '';
    }

    box.innerHTML = html;
    refreshRoleAction();
    if (!pin) {
      signInBtn.style.display = 'none';
      signOutBtn.style.display = 'none';
      html += `<br><span style="color:#666;">${TXT.pin_search_again}</span>`;
      box.innerHTML = html;
    }
    if (pin && !data.volunteer.pin_ok) {
      signInBtn.style.display = 'none';
      signOutBtn.style.display = 'none';
      html += `<br><b style="color:#842029;">${TXT.pin_wrong}</b>`;
      box.innerHTML = html;
    }
    btn.disabled = false;

  } else {
    box.className = 'alert alert-danger';
    updateSigninOnlyVisibility(false);
    window.lastAttendanceLookup = null;
    box.innerHTML = `${TXT.not_found_id}：${finalId}`;
  }
}

function updateClock(){

    const now = new Date();

    const text = new Intl.DateTimeFormat(
        "en-MY",
        {
            timeZone:"Asia/Kuala_Lumpur",
            hour:"2-digit",
            minute:"2-digit",
            second:"2-digit",
            hour12:true
        }
    ).format(new Date());

    document.getElementById("clock").innerText=text;

}

updateClock();

setInterval(updateClock,1000);

function toggleBranch() {
  const btn = document.getElementById("branch_btn");
  const branch = document.getElementById("branch");

  if (branch.value === "CHE") {
    branch.value = "STW";
    invalidateLookup();
    btn.innerText = "STW";
    btn.className = "btn-tool btn-danger branch-btn";
  } else {
    branch.value = "CHE";
    invalidateLookup();
    btn.innerText = "CHE";
    btn.className = "btn-tool btn-success branch-btn";
  }
}

function updateSigninOnlyVisibility(isSigningOut) {
  const fields = document.getElementById('signinOnlyFields');
  const tip = document.getElementById('signinOnlyTip');
  if (fields) fields.style.display = isSigningOut ? 'none' : '';
  if (tip) tip.style.display = isSigningOut ? 'none' : '';
}

function refreshRoleAction() {
  const data = window.lastAttendanceLookup;
  if (!data || !data.ok) return;
  const role = document.getElementById('role').value;
  const pinEntered = !!document.getElementById('pin').value.trim();
  const pinValid = !pinEntered || (verifiedLookupKey === currentLookupKey() && data.volunteer.pin_ok);
  const signin = document.getElementById('signInBtn');
  const signout = document.getElementById('signOutBtn');
  const assignmentInput = document.getElementById('assignment_id');
  signin.style.display = 'none';
  signout.style.display = 'none';
  assignmentInput.value = '';
  if (!pinValid || !pinEntered) return;
  const openRoles = data.open_roles || [];
  if (openRoles.length) signout.style.display = '';
  // 值班与供台可并行，其余岗位仍须先签退。
  if (!role || !(data.eligible_roles || []).some(item => item.role === role)) return;
  if (openRoles.includes(role)) return;
  if (openRoles.length && !(('值班' === role || '供台' === role) &&
      openRoles.every(r => (r === '值班' || r === '供台') && r !== role))) return;
  // 已完成某一轮值班，不影响同一天其他岗位签到。
  const pending = (data.today_assignments || []).find(a => a.role === role && a.state === 'pending');
  if (pending && pending.id && !pending.display_only) assignmentInput.value = pending.id;
  signin.style.display = '';
}

function toggleCardNo() {
    // 岗位改变后，不沿用另一岗位的排班编号。
    document.getElementById("assignment_id").value = "";
    const role = document.getElementById("role").value;
    const group = document.getElementById("card_no_group");

    if (role === "值班") {
        group.style.display = "block";
    } else {
        group.style.display = "none";
    }

}

function toggleAdminBox() {
  const box = document.getElementById("adminBox");

  if (box.style.display === "none") {
    box.style.display = "block";
  } else {
    box.style.display = "none";
  }
}

function quickSignIn() {
  if (verifiedLookupKey !== currentLookupKey()) {
    alert(TXT.lookup_required_alert);
    return false;
  }
  const pin = document.getElementById('pin').value.trim();

  if (!pin) {
    alert(TXT.enter_pin);
    return false;
  }

  const idInput = document.getElementById("volunteer_id");
  const rawId = idInput.value.trim();
  const branch = document.getElementById("branch").value;

  if (branch === "STW" && rawId && !rawId.toUpperCase().startsWith("STW")) {
    idInput.value = "STW-" + rawId;
  }

  return true;
}

let pendingSignOutId = "";

function submitQuickSignOut() {
  if (verifiedLookupKey !== currentLookupKey()) {
    alert(TXT.lookup_required_alert);
    return;
  }
  const pin = document.getElementById('pin').value.trim();
  const volunteerId = document.getElementById("volunteer_id").value.trim();

  if (!volunteerId) {
    alert(TXT.enter_id_first);
    return;
  }

  const branch = document.getElementById("branch").value;

  pendingSignOutId = volunteerId;

  if (
    branch === "STW" &&
    volunteerId &&
    !volunteerId.toUpperCase().startsWith("STW")
  ) {
    pendingSignOutId = "STW-" + volunteerId;
  }

  // 查询时已验证 PIN；直接提交，后台 sign_out() 仍会再次验证。
  document.getElementById('signout_volunteer_id').value = pendingSignOutId;
  document.getElementById('signout_pin').value = pin;
  document.getElementById('signout_assignment_id').value = '';
  document.getElementById('signout_role').value = document.getElementById('signoutRole').value || '';
  document.getElementById('quickSignOutForm').submit();
}

function closeSignOutDialog() {
  document.getElementById("signOutDialog").style.display = "none";
}

function confirmSignOut() {
  const pin = document.getElementById("dialog_pin").value.trim();

  if (!pin) {
    alert("请输入 PIN");
    document.getElementById("dialog_pin").focus();
    return;
  }

  document.getElementById("signout_volunteer_id").value = pendingSignOutId;
  document.getElementById("signout_pin").value = pin;
  document.getElementById("signout_assignment_id").value =
      document.getElementById("assignment_id").value || "";

  document.getElementById("quickSignOutForm").submit();
}

function toggleDialogPin(cb) {
  const input = document.getElementById("dialog_pin");
  input.type = cb.checked ? "text" : "password";
}

setTimeout(() => {
  document.querySelectorAll('.alert-success').forEach(el => {
    if (el.id !== 'personBox') {
      el.style.display = 'none';
    }
  });
}, 10000);

const volunteerIdInput = document.getElementById('volunteer_id');
const pinInput = document.getElementById('pin');

if (volunteerIdInput && pinInput) {
  volunteerIdInput.addEventListener('keydown', function(e) {
    if (e.key === 'Enter') {
      e.preventDefault();
      pinInput.focus();
    }
  });
}

window.addEventListener("load", function() {
  const adminPin = document.getElementById("admin_pin");

  if (pinInput) pinInput.value = "";
  ['volunteer_id', 'pin'].forEach(id => {
    document.getElementById(id).addEventListener('input', invalidateLookup);
  });
  if (adminPin) adminPin.value = "";

  toggleCardNo();
});
</script>

<div id="signOutDialog" class="dialog-mask" style="display:none;">

    <div class="dialog-card">

        <h2 class="section-title">
            {{ t.confirm_signout }}
        </h2>

        <div class="page-subtitle">
            {{ t.enter_volunteer_pin }}
        </div>

        <div class="form-group">

            <input
                id="dialog_pin"
                class="form-input"
                type="password"
                placeholder="{{ t.enter_pin_for_checkout }}">

        </div>

        <label style="display:flex;align-items:center;gap:10px;margin-bottom:20px;">

            <input
                type="checkbox"
                onchange="toggleDialogPin(this)">

            {{ t.show_pin }}

        </label>

        <div class="btn-row">

            <button
                class="btn-tool btn-secondary"
                onclick="closeSignOutDialog()">

                {{ t.cancel }}

            </button>

            <button
                class="btn-tool btn-danger"
                onclick="confirmSignOut()">

                {{ t.confirm_checkout }}

            </button>

        </div>

    </div>

</div>

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

  <link rel="stylesheet" href="{{ url_for('static', filename='css/toolbox.css') }}">
</head>

<body>

<div class="page">

  <div class="card">

    <h1 class="page-title">✏️ {{ t.edit_title }}</h1>

    <div class="summary-grid">
      <div class="summary-box">
        <div class="summary-title">{{ t.name }}</div>
        <div class="summary-value">{{ record.get('姓名','') }}</div>
      </div>

      <div class="summary-box">
        <div class="summary-title">{{ t.date }}</div>
        <div class="summary-value">{{ record.get('日期','') }}</div>
      </div>
    </div>

    <form method="post" action="{{ url_for('save_edit', lang=lang) }}">
      <input type="hidden" name="row_number" value="{{ row_number }}">

      <div class="form-group">
        <label class="form-label">{{ t.role }}</label>
        <select class="form-input" name="role" required>
          {% for role in roles %}
            <option value="{{ role }}"
              {% if role == record.get('岗位') %}selected{% endif %}>
              {{ role_label(role) }}
            </option>
          {% endfor %}
        </select>
      </div>

      <div class="form-group">
        <label class="form-label">{{ t.card_no_label }}</label>
        <input
          class="form-input"
          type="text"
          name="card_no"
          value="{{ record.get('card_no','') }}"
          placeholder="{{ t.card_no_placeholder }}"
        >
      </div>

      <div class="form-group">
        <label class="form-label">{{ t.start_time }}</label>
        <input
          class="form-input"
          name="start_time"
          value="{{ record.get('开始时间','') }}"
          placeholder="10:00am"
        >
      </div>

      <div class="form-group">
        <label class="form-label">{{ t.end_time }}</label>
        <input
          class="form-input"
          name="end_time"
          value="{{ record.get('结束时间','') }}"
        >
      </div>

      <div class="form-group">
        <label class="form-label">{{ t.remark }}</label>
        <input
          class="form-input"
          name="remark"
          value="{{ record.get('备注','') }}"
        >
      </div>

      <div class="btn-row">
        <button class="btn-tool btn-success" type="submit">
          {{ t.save_edit }}
        </button>
      </div>

    </form>

    <hr>

    <form method="post"
          action="{{ url_for('delete_edit', lang=lang) }}"
          onsubmit="return confirm({{ t.delete_confirm|tojson }});">

      <input type="hidden" name="row_number" value="{{ row_number }}">

      <div class="form-group">
        <label class="form-label">{{ t.admin_pin }}</label>
        <input
          class="form-input"
          type="password"
          name="pin"
          placeholder="{{ t.pin_placeholder }}"
          required
        >
      </div>

      <div class="btn-row">
        <button class="btn-tool btn-danger" type="submit">
          {{ t.delete_record }}
        </button>
      </div>

    </form>

    <div class="btn-row">
      <a class="btn-tool btn-secondary"
         href="{{ url_for('index', lang=lang) }}">
        {{ t.back_home }}
      </a>
    </div>

  </div>

</div>

</body>
</html>
"""


PIN_PAGE = """
<!doctype html>
<html lang="{{ t.html_lang }}">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">

<title>{{ t.change_pin_title }}</title>

<link rel="stylesheet"
      href="{{ url_for('static', filename='css/toolbox.css') }}">

</head>

<body>

<div class="page">

    <div class="card">

        <h1 class="page-title">
            🔑 {{ t.change_pin_title }}
        </h1>

        <div class="page-subtitle">
            为了帐号安全，请定期更换 PIN。
        </div>

        {% with messages = get_flashed_messages(with_categories=true) %}
            {% for category, message in messages %}
                <div class="alert {{ 'alert-success' if category == 'ok' else 'alert-danger' }}">
                    {{ message }}
                </div>
            {% endfor %}
        {% endwith %}

        <form method="post">

            <div class="form-group">
                <label class="form-label">
                    {{ t.enter_id }}
                </label>

                <input
                    class="form-input"
                    name="id"
                    inputmode="numeric"
                    placeholder="{{ t.enter_id }}"
                    required>
            </div>

            <div class="form-group">
                <label class="form-label">
                    {{ t.old_pin }}
                </label>

                <input
                    class="form-input"
                    name="old"
                    type="password"
                    inputmode="numeric"
                    placeholder="{{ t.old_pin }}"
                    required>
            </div>

            <div class="form-group">
                <label class="form-label">
                    {{ t.new_pin }}
                </label>

                <input
                    class="form-input"
                    name="new"
                    type="password"
                    inputmode="numeric"
                    placeholder="{{ t.new_pin }}"
                    required>
            </div>

            <div class="form-group">
                <label class="form-label">
                    {{ t.confirm_new_pin }}
                </label>

                <input
                    class="form-input"
                    name="confirm"
                    type="password"
                    inputmode="numeric"
                    placeholder="{{ t.confirm_new_pin }}"
                    required>
            </div>

            <div class="btn-row">

                <button
                    class="btn-tool btn-success"
                    type="submit">

                    {{ t.save }}

                </button>

            </div>

        </form>

        <div class="btn-row">

            <a
                class="btn-tool btn-secondary"
                href="{{ url_for('index') }}">

                {{ t.back_home }}

            </a>

        </div>

    </div>

</div>

</body>
</html>
"""

# =========================
# 7) 路由
# =========================
@app.route("/")
def index():
    auto_signout_unfinished_today()
    stats = get_today_stats()

    from zoneinfo import ZoneInfo
    from datetime import datetime

    now = datetime.now(ZoneInfo("Asia/Kuala_Lumpur"))

    weekday_map = {
        0: "星期一",
        1: "星期二",
        2: "星期三",
        3: "星期四",
        4: "星期五",
        5: "星期六",
        6: "星期日",
    }

    today_text = (now.strftime("%Y年%m月%d日") + " " + weekday_map[now.weekday()]
                  if get_lang() == "zh" else now.strftime("%d %B %Y (%A)"))
    malaysia_time = now.strftime("%I:%M %p")

    special_day_info = get_special_day_info(now.date())

    lunar_text = special_day_info.get("lunar_text", "")

    special_names = special_day_info.get("special_names", [])
    buddhist_day = " / ".join(special_names)

    return render_template_string(
        PAGE,
        t=get_text(),
        roles=ROLES,
        role_label=role_label,
        role_text_labels={r: role_label(r) for r in ROLES},
        today_count=stats["total"],
        not_out=stats["open"],
        done_out=stats["finished"],

        today_text=today_text,
        lunar_text=lunar_text,
        buddhist_day=buddhist_day,
        malaysia_time=malaysia_time,
    )

@app.before_request
def _attendance_speed_start():
    if request.path in ("/", "/signin", "/signout") or request.path.startswith("/api/volunteer/"):
        request.environ["attendance_speed_start"] = time.perf_counter()


@app.after_request
def _attendance_speed_report(response):
    started = request.environ.get("attendance_speed_start")
    if started is not None:
        elapsed_ms = (time.perf_counter() - started) * 1000
        route = "/api/volunteer/<id>" if request.path.startswith("/api/volunteer/") else request.path
        app.logger.info("[ATTENDANCE_SPEED] route=%s status=%s elapsed_ms=%.1f", route, response.status_code, elapsed_ms)
        response.headers["Server-Timing"] = f"app;dur={elapsed_ms:.1f}"
    return response


@app.route("/signin", methods=["POST"])
def do_sign_in():
    auto_signout_unfinished_today()

    volunteer_id = request.form.get("volunteer_id", "").strip()
    role = request.form.get("role", "").strip()
    # 三种岗位限制由 attendance_service.sign_in() 统一判断。
    # 不在路由重复检查，以免卫生报名被错误要求正式排班。
    ok, msg = sign_in(
        volunteer_id,
        request.form.get("pin", ""),
        role,
        request.form.get("card_no", ""),
        request.form.get("assignment_id", ""),
    )

    if ok:
        flash(f"{msg}｜签到时间：{now_time_str()}", "ok")
    else:
        flash(msg, "bad")

    return redirect(url_for("index"))


@app.route("/signout", methods=["POST"])
def do_sign_out():
    auto_signout_unfinished_today()
    pin = request.form.get("pin", "").strip()

    # 旧方式：从今日进行中表格签退
    row_number_raw = request.form.get("row_number", "").strip()

    if row_number_raw:
        try:
            row_number = int(row_number_raw)
        except Exception:
            row_number = 0

        row = db_query(
            "select * from attendance where id=%s",
            (row_number,),
            fetchone=True
        )

        if not row:
            flash("记录不存在", "bad")
            return redirect(url_for("index"))

        volunteer_id = str(row["volunteer_id"]).strip()

    else:
        # 核心服务统一验证编号、PIN 和进行中的签到记录。
        volunteer_id = request.form.get("volunteer_id", "").strip()

    ok, msg = sign_out(
        volunteer_id,
        pin,
        request.form.get("assignment_id", ""),
        role=request.form.get("role", "").strip() or (str(row.get("role") or "") if row_number_raw and row else None)
    )

    if ok:
        flash(f"{msg}｜签退时间：{now_time_str()}", "ok")
    else:
        flash(msg, "bad")

    return redirect(url_for("index"))

@app.route("/edit/<int:row_number>")
def edit_page(row_number):
    lang = request.args.get("lang", "zh")
    t = TEXT.get(lang, TEXT["zh"])

    record = None
    for r in get_today_records():
        if int(r.get("_row", 0)) == row_number:
            record = r
            break

    if not record:
        flash(t["record_not_found"], "bad")
        return redirect(url_for("index", lang=lang))

    return render_template_string(
        EDIT_PAGE,
        t=t,
        lang=lang,
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
        card_no=request.form.get("card_no", ""),
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
    pin = request.form.get("pin", "")
    ok, msg = delete_record(row_number, pin)
    flash(msg, "ok" if ok else "bad")
    return redirect(url_for("index"))


@app.route("/download_reading")
def download_reading():

    try:
        start_date = request.args.get("start_date", "").strip()
        end_date = request.args.get("end_date", "").strip()

        use_date_filter = bool(start_date and end_date)

        if use_date_filter:
            rows = db_query("""
                select *
                from reading
                where date between %s and %s
                order by date, time
            """, (start_date, end_date), fetchall=True)

            file_date_text = f"{start_date}_to_{end_date}"

        else:
            rows = db_query("""
                select *
                from reading
                order by date, time
            """, fetchall=True)

            file_date_text = "all"

        if not rows:
            return "没有数据"

        df = pd.DataFrame(rows)

        df = df.rename(columns={
            "date": "日期",
            "name": "姓名",
            "identity": "身份",
            "topic": "主题",
            "session": "场次",
            "time": "时间"
        })

        df = df[
            ["日期", "姓名", "身份", "主题", "场次", "时间"]
        ]

        output = io.BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(
                writer,
                index=False,
                sheet_name="白话佛法记录"
            )

        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name=f"reading_report_{file_date_text}.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        return f"下载失败：{e}"


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


@app.route("/api/volunteer/<volunteer_id>", methods=["POST"])
def api_volunteer(volunteer_id):
    try:
        auto_signout_unfinished_today()
        v = find_volunteer(volunteer_id)
        if not v:
            return jsonify({
                "ok": False,
                "error": "find_volunteer returned None"
            })

        pin = request.form.get("pin", "").strip()
        pin_ok = verify_pin_for_volunteer(v, pin) if pin else False

        today = datetime.now(ZoneInfo("Asia/Kuala_Lumpur")).date()

        # 同一个数据库连接读取排班及签到状态，避免反复借还连接。
        with get_conn() as conn:
            with conn.cursor(cursor_factory=RealDictCursor) as cur:
                cur.execute("""
                    select id, role, assigned_place, start_time, end_time, shift_label
                    from volunteer_schedule_assignments
                    where volunteer_id = %s and assignment_date = %s
                      and coalesce(status, 'assigned') <> 'cancelled'
                    order by start_time, assigned_place
                """, (v.get("编号"), today))
                today_assignments = cur.fetchall()
                # 负责人指定的设师父供台人员保存在每日标记中，仅用于页面显示。
                cur.execute("""
                    SELECT setup_people FROM schedule_day_flags
                    WHERE flag_date = %s
                """, (today,))
                setup_flag = cur.fetchone()
                setup_people = (setup_flag.get('setup_people') or '') if setup_flag else ''
                # 膳食组报名独立于正式排班；只补显示，不能伪造 assignment_id。
                cur.execute("""
                    SELECT id, meal_role, start_time, end_time
                    FROM volunteer_schedule_signups
                    WHERE volunteer_id = ANY(%s) AND signup_date = %s
                      AND role = '膳食'
                      AND coalesce(status, 'pending') <> 'cancelled'
                    ORDER BY created_at, id
                """, (get_assignment_id_candidates(v, volunteer_id), today))
                meal_signups = cur.fetchall() or []
                cur.execute("""
                    SELECT role, end_time FROM attendance
                    WHERE volunteer_id = %s AND date = %s
                    ORDER BY id DESC
                """, (v.get('编号'), now_date_str()))
                today_attendance_roles = cur.fetchall() or []
                cur.execute("""
                    select assignment_id from volunteer_attendance_logs
                    where volunteer_id = %s and attendance_date = %s
                      and check_out_time is null
                    order by id desc limit 1
                """, (v.get("编号"), today))
                open_log = cur.fetchone()
                cur.execute("""
                    select assignment_id, remarks from volunteer_attendance_logs
                    where volunteer_id = %s and attendance_date = %s
                      and check_out_time is not null and assignment_id is not null
                """, (v.get("编号"), today))
                completed_rows = cur.fetchall() or []
                cur.execute("""
                    select id, role from attendance
                    where volunteer_id = %s and date = %s
                      and (end_time is null or end_time = '')
                    order by id desc
                """, (v.get("编号"), now_date_str()))
                actual_open_rows = cur.fetchall() or []
                actual_open = actual_open_rows[0] if actual_open_rows else None
                paid_until = None
                if pin_ok:
                    cur.execute("""
                        select max(end_month) as paid_until
                        from member_payments where member_id = %s
                    """, (v.get("编号"),))
                    paid_row = cur.fetchone()
                    paid_until = paid_row.get("paid_until") if paid_row else None
        # 固定佛台由排班系统的「Excel 固定轮值 + 当日云端覆盖」决定。
        # 不将它插入正式排班表，避免破坏发布、自动补排逻辑。
        from schedule.schedule_web import get_fixed_buddha_for_date
        from opencc import OpenCC
        _cc = OpenCC('t2s')
        _normalize_name = lambda name: ''.join(_cc.convert(str(name or '')).split())
        fixed_buddha_names = get_fixed_buddha_for_date(today.strftime('%Y-%m-%d'))
        is_fixed_buddha = _normalize_name(v.get('姓名')) in {
            _normalize_name(name) for name in fixed_buddha_names
        }

        def _extra_state(role):
            records = [r for r in today_attendance_roles if r.get('role') == role]
            if any(not r.get('end_time') for r in records):
                return 'active'
            if records:
                return 'completed'
            return 'pending'

        if is_fixed_buddha and not any(a.get('role') in ('佛台', '整理佛台') for a in today_assignments):
            today_assignments.append({
                'id': None, 'role': '佛台', 'assigned_place': '整理佛台',
                'start_time': '8:00am', 'end_time': '完成佛台工作',
                'shift_label': '', 'state': _extra_state('佛台'),
                'display_only': True,
            })
        # 设师父供台名单以姓名换行保存；与固定佛台一样只补显示，不伪造排班 ID。
        setup_names = {_normalize_name(name) for name in setup_people.splitlines() if name.strip()}
        if _normalize_name(v.get('姓名')) in setup_names and not any(
            a.get('role') == '供台' and a.get('assigned_place') == '设师父供台'
            for a in today_assignments
        ):
            today_assignments.append({
                'id': None, 'role': '供台', 'assigned_place': '设师父供台',
                'start_time': '', 'end_time': '', 'shift_label': '',
                'state': _extra_state('供台'), 'display_only': True,
            })
        if meal_signups and not any(a.get('role') == '膳食' for a in today_assignments):
            meal = meal_signups[0]
            meal_position = next((i for i, r in enumerate(meal_signups)
                                  if r['id'] == meal['id']), 0)
            meal_role = meal.get('meal_role') or ('派餐义工' if meal_position < 9 else '候补义工')
            today_assignments.append({
                'id': None, 'role': '膳食', 'assigned_place': meal_role,
                'start_time': meal.get('start_time') or '8:00am',
                'end_time': meal.get('end_time') or '2:00pm',
                'shift_label': '', 'state': _extra_state('膳食'),
                'display_only': True,
            })
        today_assignments.sort(key=lambda x: str(x.get('start_time') or ''))

        completed_ids = {int(x["assignment_id"]) for x in completed_rows if x.get("assignment_id") is not None}
        auto_ids = {int(x["assignment_id"]) for x in completed_rows
                    if x.get("assignment_id") is not None and "系统自动签退" in str(x.get("remarks") or "")}
        open_assignment_id = int(open_log["assignment_id"]) if open_log and open_log.get("assignment_id") is not None else None

        attendance_action = "walkin"
        current_assignment_id = None
        current_assignment_text = ""

        for a in today_assignments:
            if a.get('display_only'):
                continue
            aid = int(a.get("id") or 0)
            if aid == open_assignment_id:
                a["state"] = "active"
            elif aid in completed_ids:
                a["state"] = "completed"
                a["auto_signed_out"] = aid in auto_ids
            else:
                a["state"] = "pending"

        if actual_open:
            attendance_action = "signout"
            current_assignment_id = open_assignment_id
        elif open_assignment_id is not None:
            attendance_action = "signout"
            current_assignment_id = open_assignment_id
        elif not today_assignments:
            attendance_action = "walkin"
        elif today_assignments:
            pending = next((a for a in today_assignments if a.get("state") == "pending" and not a.get('display_only')), None)
            if pending:
                attendance_action = "signin"
                current_assignment_id = pending.get("id")
                current_assignment_text = f"{pending.get('start_time') or ''} ～ {pending.get('end_time') or ''} {pending.get('assigned_place') or ''}".strip()
            else:
                attendance_action = ("walkin" if any(a.get('display_only') and a.get('state') == 'pending'
                                                       for a in today_assignments) else "completed")

        safe_v = {
            "编号": v.get("编号", ""),
            "姓名": v.get("姓名", ""),
            "状态": v.get("状态", ""),
            "pin_ok": pin_ok,
        }

        if pin_ok:
            safe_v["电话号码"] = v.get("电话号码", "")

            if paid_until:
                try:
                    paid_until = f"{paid_until.year}年{paid_until.month}月"
                except Exception:
                    paid_until = str(paid_until)

            safe_v["月费已缴费至"] = paid_until

        eligible_roles = ([{"role": role, "label": role_label(role)}
                           for role in get_eligible_signin_roles(v, volunteer_id)]
                          if pin_ok else [])

        return jsonify({
            "ok": True,
            "volunteer": safe_v,
            "eligible_roles": eligible_roles,
            "today_assignments": today_assignments,
            "attendance_action": attendance_action,
            "current_assignment_id": current_assignment_id,
            "current_assignment_text": current_assignment_text,
            "open_roles": [r["role"] for r in actual_open_rows]
        })

    except Exception as e:
        print("API ERROR =", e)
        return jsonify({
            "ok": False,
            "error": str(e)
        })
    
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
    print("====================================")

    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)

    