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
import io
import sys
import psycopg2
import pandas as pd

from pathlib import Path
from typing import Optional
from zoneinfo import ZoneInfo
from admin_web import admin_bp
from member_web import member_bp
from pypinyin import lazy_pinyin
from manifest import manifest_bp
from db import db_query, get_conn
from library_web import library_bp
from reading_web import reading_bp

from finance_web import finance_bp
import finance_export

from psycopg2.extras import RealDictCursor
from finance_import import finance_import_bp
from lunar_rules import get_special_day_info
from dharma_class_web import dharma_class_bp
from schedule.schedule_web import schedule_bp
from datetime import datetime, date, timedelta
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
    TODAY_CODE_ENABLED,
)

from utils import (
    MY_TZ,
    TODAY_CODE_ENABLED,
    TODAY_CODE_LIST,
    now_date_str,
    now_time_str,
    get_today_code,
    get_display_today_code,
    verify_today_code,
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
app.register_blueprint(finance_month_end_bp)
app.register_blueprint(finance_import_bp)
app.register_blueprint(finance_audit_bp)
app.register_blueprint(dharma_class_bp)
app.register_blueprint(schedule_bp)
app.register_blueprint(manifest_bp)
app.register_blueprint(reading_bp)
app.register_blueprint(finance_bp)
app.register_blueprint(library_bp)
app.register_blueprint(member_bp)
app.register_blueprint(admin_bp)


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
          🙏 {{ t.get("welcome_home", "欢迎回来，感恩您的发心护持") }}
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
      {% if today_code %}
      {% if today_code_enabled %}
      <div class="form-group">
        <label class="form-label">{{ t.today_code }}</label>
        <input
          class="form-input"
          id="today_code"
          type="tel"
          name="today_code"
          inputmode="numeric"
          pattern="[0-9]*"
          placeholder="{{ t.today_code_placeholder }}"
          required
        >
      </div>
      {% endif %}
      {% endif %}
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

      <div class="form-row">

        <div class="form-group form-col">

            <label class="form-label">
                {{ t.role }}
            </label>

            <select
                class="form-input"
                name="role"
                id="role"
                onchange="toggleCardNo()">

                {% for role in roles %}
                    <option value="{{ role }}">{{ role_label(role) }}</option>
                {% endfor %}

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

      <div class="btn-row">
        <button id="signInBtn"
                class="btn-tool btn-success"
                type="submit">
            {{ t.check_in }}
        </button>

        <button class="btn-tool btn-danger"
                type="button"
                onclick="submitQuickSignOut();">
            {{ t.sign_out }}
        </button>
      </div>

    </form>

    <form id="quickSignOutForm"
          method="post"
          action="{{ url_for('do_sign_out') }}">
      <input type="hidden" id="signout_volunteer_id" name="volunteer_id">
      <input type="hidden" id="signout_pin" name="pin">
    </form>

    <div class="small-note">
      💡 {{ t.card_no_duty_tip }}。
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
      <h2 class="section-title">🔐 {{ t.admin_entry }}</h2>

      <div class="page-subtitle">
        {{ t.admin_entry_desc }}
      </div>

      <div class="btn-row single-btn-row">
        <button
            type="button"
            class="btn-tool btn-purple small-action-btn"
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

      <div class="small-note">
        🛡️ {{ t.admin_pin_required }}
      </div>
    </div>

  </div>

  <div class="footer">
    © {{ t.footer_text }} 💚
  </div>

</div>

<script>
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
  pin_wrong: {{ t.pin_wrong|tojson }}
};

async function lookupVolunteer() {
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
  box.style.display = 'block';

  if (data.ok) {

    box.className = 'alert alert-success';

    let html =
      `${TXT.name}：<b>${data.volunteer.姓名}</b><br>` +
      `${TXT.status}：${data.volunteer.状态 || '-'}`;

    if (data.volunteer.pin_ok) {
      html += `<br>${TXT.phone}：${data.volunteer.电话号码 || '-'}`;

      const paidUntil = data.volunteer["月费已缴费至"];

      if (paidUntil && paidUntil !== "-") {
        html += `<br>${TXT.paid_until}：${paidUntil}`;
      }
    } else if (pin) {
      html += `<br><span style="color:#842029;">${TXT.pin_wrong}</span>`;
    }

    if (data.today_assignments && data.today_assignments.length > 0) {

      html += `<br><hr>`;
      html += `<b>📅 今天正式安排</b>`;

      data.today_assignments.forEach(a => {
        html += `
          <div style="margin-top:10px;">
            地点：<b>${a.assigned_place || "-"}</b><br>
            岗位：${a.role || "-"}<br>
            时间：${a.start_time || "-"} ～ ${a.end_time || "-"}
          </div>
        `;
      });

    } else {
      html += `
        <br><hr>
        <b style="color:#b02a37;">今天没有正式安排</b>
      `;
    }

    box.innerHTML = html;
    btn.disabled = false;

  } else {
    box.className = 'alert alert-danger';
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
    btn.innerText = "STW";
    btn.className = "btn-tool btn-danger branch-btn";
  } else {
    branch.value = "CHE";
    btn.innerText = "CHE";
    btn.className = "btn-tool btn-success branch-btn";
  }
}

function toggleCardNo() {
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
  const volunteerId = document.getElementById("volunteer_id").value.trim();

  if (!volunteerId) {
    alert("请先输入义工编号");
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

  document.getElementById("dialog_pin").value = "";
  document.getElementById("signOutDialog").style.display = "flex";

  setTimeout(function() {
    document.getElementById("dialog_pin").focus();
  }, 100);
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
const todayCodeInput = document.getElementById('today_code');

if (volunteerIdInput && pinInput) {
  volunteerIdInput.addEventListener('keydown', function(e) {
    if (e.key === 'Enter') {
      e.preventDefault();
      pinInput.focus();
    }
  });
}

if (pinInput && todayCodeInput) {
  pinInput.addEventListener('keydown', function(e) {
    if (e.key === 'Enter') {
      e.preventDefault();
      todayCodeInput.focus();
    }
  });
}

if (todayCodeInput) {
  todayCodeInput.addEventListener('keydown', function(e) {
    if (e.key === 'Enter') {
      e.preventDefault();
      lookupVolunteer();
    }
  });
}

window.addEventListener("load", function() {
  const adminPin = document.getElementById("admin_pin");

  if (pinInput) pinInput.value = "";
  if (todayCodeInput) todayCodeInput.value = "";
  if (adminPin) adminPin.value = "";

  toggleCardNoBox();
});
</script>

<div id="signOutDialog" class="dialog-mask" style="display:none;">

    <div class="dialog-card">

        <h2 class="section-title">
            {{ t.confirm_signout }}
        </h2>

        <div class="page-subtitle">
            请输入该义工 PIN
        </div>

        <div class="form-group">

            <input
                id="dialog_pin"
                class="form-input"
                type="password"
                placeholder="请输入 PIN">

        </div>

        <label style="display:flex;align-items:center;gap:10px;margin-bottom:20px;">

            <input
                type="checkbox"
                onchange="toggleDialogPin(this)">

            显示 PIN

        </label>

        <div class="btn-row">

            <button
                class="btn-tool btn-secondary"
                onclick="closeSignOutDialog()">

                取消

            </button>

            <button
                class="btn-tool btn-danger"
                onclick="confirmSignOut()">

                确认签退

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

    today_text = now.strftime("%Y年%m月%d日") + " " + weekday_map[now.weekday()]
    malaysia_time = now.strftime("%I:%M %p")

    special_day_info = get_special_day_info(now.date())

    lunar_text = special_day_info.get("lunar_text", "")
    buddhist_day = special_day_info.get("festival_name", "")

    return render_template_string(
        PAGE,
        t=get_text(),
        roles=ROLES,
        role_label=role_label,
        open_records=get_today_open_records(),
        today_records=get_today_records(limit=20),
        today_code_enabled=TODAY_CODE_ENABLED,
        today_count=stats["total"],
        not_out=stats["open"],
        done_out=stats["finished"],

        today_text=today_text,
        lunar_text=lunar_text,
        buddhist_day=buddhist_day,
        malaysia_time=malaysia_time,
    )

@app.route("/signin", methods=["POST"])
def do_sign_in():

    volunteer_id = request.form.get("volunteer_id", "").strip()
    role = request.form.get("role", "").strip()
    # 值班、卫生、膳食必须有正式安排

    v = find_volunteer(volunteer_id)

    if not v:
        flash("❌ 找不到义工编号", "bad")
        return redirect(url_for("index"))

    volunteer_id = str(v.get("编号") or volunteer_id).strip()

    if role in ["值班", "卫生", "膳食"]:

        today_assignments = get_today_assignments(
            volunteer_id,
            role,
            current_time=True
        )
        
        if not today_assignments:

          if REQUIRE_ASSIGNMENT_FOR_SIGNIN:

              if ENABLE_SIGNIN_TIME_LIMIT:
                  flash("❌ 现在不是可签到时间，或今天没有正式安排。", "bad")
              else:
                  flash("❌ 今天没有正式安排这个岗位，不能签到。", "bad")

              return redirect(url_for("index"))
    
    if TODAY_CODE_ENABLED:
      input_code = request.form.get("today_code", "").strip()

      if not verify_today_code(input_code):
          flash("今日签到码错误，请看现场公布的号码", "bad")
          return redirect(url_for("index"))

    ok, msg = sign_in(
        volunteer_id,
        request.form.get("pin", ""),
        role,
        request.form.get("card_no", ""),
    )

    if ok:
        flash(f"{msg}｜签到时间：{now_time_str()}", "ok")
    else:
        flash(msg, "bad")

    return redirect(url_for("index"))


@app.route("/signout", methods=["POST"])
def do_sign_out():
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
        # 新方式：首页直接用义工编号签退
        raw_id = request.form.get("volunteer_id", "").strip()
        volunteer = find_volunteer(raw_id)

        if not volunteer:
            flash("找不到这个义工编号", "bad")
            return redirect(url_for("index"))

        volunteer_id = str(volunteer.get("编号") or raw_id).strip()

        if not volunteer_id:
            flash("请输入义工编号", "bad")
            return redirect(url_for("index"))

        row = db_query("""
            select *
            from attendance
            where volunteer_id = %s
              and date = %s
              and (end_time is null or end_time = '')
            order by id desc
            limit 1
        """, (
            volunteer_id,
            now_date_str()
        ), fetchone=True)

        if not row:
            flash("今天没有找到未签退记录", "bad")
            return redirect(url_for("index"))

    ok, msg = sign_out(volunteer_id, pin)

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
        print("API volunteer_id =", volunteer_id)

        v = find_volunteer(volunteer_id)
        print("API result =", v)

        if not v:
            return jsonify({
                "ok": False,
                "error": "find_volunteer returned None"
            })

        pin = request.form.get("pin", "").strip()
        pin_ok = verify_pin_for_volunteer(v, pin) if pin else False

        today = date.today()

        today_assignments = []

        with get_conn() as conn:
            with conn.cursor(cursor_factory=RealDictCursor) as cur:
                cur.execute("""
                    select
                        role,
                        assigned_place,
                        start_time,
                        end_time,
                        shift_label
                    from volunteer_schedule_assignments
                    where volunteer_id = %s
                      and assignment_date = %s
                      and coalesce(status, 'assigned') <> 'cancelled'
                    order by start_time, assigned_place
                """, (
                    v.get("编号"),
                    today
                ))

                today_assignments = cur.fetchall()

        safe_v = {
            "编号": v.get("编号", ""),
            "姓名": v.get("姓名", ""),
            "状态": v.get("状态", ""),
            "pin_ok": pin_ok,
        }

        if pin_ok:
            safe_v["电话号码"] = v.get("电话号码", "")

            paid_until = get_member_paid_until(v.get("编号"))

            if paid_until:
                try:
                    paid_until = f"{paid_until.year}年{paid_until.month}月"
                except Exception:
                    paid_until = str(paid_until)

            safe_v["月费已缴费至"] = paid_until

        return jsonify({
            "ok": True,
            "volunteer": safe_v,
            "today_assignments": today_assignments
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

    