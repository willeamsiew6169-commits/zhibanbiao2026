# admin_web.py

import io
import pandas as pd

from io import BytesIO
from db import db_query
from utils import get_text
from openpyxl import Workbook
from datetime import date, datetime
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

    volunteers = db_query(
        "select id, name from volunteers where status='在册' order by id",
        fetchall=True
    )

    return render_template_string("""
    <!doctype html>
    <html lang="zh">
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>补录签到</title>

        <link rel="stylesheet"
              href="{{ url_for('static', filename='css/toolbox.css') }}">
    </head>

    <body>

    <div class="page">

        <div class="card">

            <h1 class="page-title">🛠 补录签到</h1>
            <div class="page-subtitle">
                手动新增义工签到记录
            </div>

            <form method="post">

                <input type="hidden" name="pin" value="{{ pin }}">

                <div class="form-group">
                    <label class="form-label">日期</label>
                    <input
                        class="form-input"
                        type="date"
                        name="date"
                        value="{{ today }}"
                        required
                    >
                </div>

                <div class="form-group">
                    <label class="form-label">选择义工</label>
                    <select
                        class="form-input"
                        name="id"
                        id="vol_id"
                        onchange="fillName()"
                        required
                    >
                        <option value="">请选择义工</option>
                        {% for v in volunteers %}
                            <option value="{{ v.id }}">
                                {{ v.name }}（{{ v.id }}）
                            </option>
                        {% endfor %}
                    </select>
                </div>

                <div class="form-group">
                    <label class="form-label">姓名</label>
                    <input
                        class="form-input"
                        type="text"
                        name="name"
                        id="vol_name"
                        required
                    >
                </div>

                <div class="form-group">
                    <label class="form-label">岗位</label>
                    <input
                        class="form-input"
                        type="text"
                        name="role"
                        placeholder="例如：值班 / 卫生 / 佛台"
                    >
                </div>

                <div class="form-group">
                    <label class="form-label">开始时间</label>
                    <input
                        class="form-input"
                        type="text"
                        name="start"
                        value="{{ now }}"
                        placeholder="例如：10:00am"
                    >
                </div>

                <div class="form-group">
                    <label class="form-label">结束时间</label>
                    <input
                        class="form-input"
                        type="text"
                        name="end"
                        placeholder="例如：2:00pm"
                    >
                </div>

                <div class="form-group">
                    <label class="form-label">备注</label>
                    <input
                        class="form-input"
                        type="text"
                        name="remark"
                        placeholder="可留空"
                    >
                </div>

                <div class="btn-row">
                    <button class="btn-tool btn-success" type="submit">
                        添加记录
                    </button>

                    <a class="btn-tool btn-secondary"
                       href="{{ url_for('admin.admin_report', pin=pin) }}">
                        返回
                    </a>
                </div>

            </form>

        </div>

    </div>

    <script>
    function fillName() {
        const sel = document.getElementById("vol_id");
        const text = sel.options[sel.selectedIndex].text;
        const name = text.split("（")[0].trim();
        document.getElementById("vol_name").value = name;
    }
    </script>

    </body>
    </html>
    """,
    pin=pin,
    today=now_date_str(),
    now=datetime.now(MY_TZ).strftime("%I:%M%p").lower(),
    volunteers=volunteers
    )

@admin_bp.route("/admin_today_open")
def admin_today_open():

    pin = request.args.get("pin", "")

    return render_template_string("""
    <!doctype html>
    <html lang="zh">
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>今日进行中</title>

        <link rel="stylesheet" href="{{ url_for('static', filename='css/toolbox.css') }}">
    </head>

    <body>

    <div class="page">

        <div class="card">

            <h1 class="page-title">🔴 今日进行中</h1>
            <div class="page-subtitle">
                尚未签退的义工记录
            </div>

            {% if open_records %}

                {% for r in open_records %}
                <div class="card">

                    <h2 class="section-title">
                        {{ r.get('姓名','') }}
                    </h2>

                    <div class="summary-grid">

                        <div class="summary-box">
                            <div class="summary-title">岗位</div>
                            <div class="summary-value">
                                {{ role_label(r.get('岗位','')) }}
                            </div>
                        </div>

                        <div class="summary-box">
                            <div class="summary-title">签到时间</div>
                            <div class="summary-value">
                                {{ r.get('开始时间','') }}
                            </div>
                        </div>

                        {% if r.get('编号') %}
                        <div class="summary-box">
                            <div class="summary-title">义工编号</div>
                            <div class="summary-value">
                                {{ r.get('编号','') }}
                            </div>
                        </div>
                        {% endif %}

                        {% if r.get('card_no') %}
                        <div class="summary-box">
                            <div class="summary-title">卡号</div>
                            <div class="summary-value">
                                {{ r.get('card_no') }}
                            </div>
                        </div>
                        {% endif %}

                    </div>

                    <form method="post"
                          action="{{ url_for('do_sign_out') }}"
                          onsubmit="return askSignOutPin(this);">

                        <input type="hidden" name="row_number" value="{{ r.get('_row') }}">
                        <input type="hidden" name="pin" value="">

                        <div class="btn-row">
                            <button class="btn-tool btn-danger" type="submit">
                                帮他签退
                            </button>

                            <a class="btn-tool btn-warning"
                               href="{{ url_for('edit_page', row_number=r.get('_row')) }}">
                                修改记录
                            </a>
                        </div>

                    </form>

                </div>
                {% endfor %}

            {% else %}

                <div class="empty-state">
                    现在没有未签退的义工
                </div>

            {% endif %}

            <div class="btn-row">
                <a class="btn-tool btn-secondary"
                   href="{{ url_for('admin.admin_report', pin=pin) }}">
                    返回管理员
                </a>
            </div>

        </div>

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
    pin=pin,
    open_records=get_today_open_records(),
    role_label=role_label,
    t=get_text()
    )

@admin_bp.route("/admin/auto_signout")
def admin_auto_signout():
    if not session.get("schedule_login"):
        return redirect(url_for("index"))

    result = auto_signout_unfinished_today()

    flash(result["msg"], "ok" if result["ok"] else "bad")
    return redirect(url_for("index"))

@admin_bp.route("/admin_today_records")
def admin_today_records():

    pin = request.args.get("pin", "")

    return render_template_string("""
    <!doctype html>
    <html lang="zh">
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>今日签到记录</title>

        <link rel="stylesheet" href="{{ url_for('static', filename='css/toolbox.css') }}">
        <style>
        .today-record-card .section-title {
            font-size: 28px;
            margin-bottom: 18px;
        }

        .today-record-card .summary-title {
            font-size: 18px;
        }

        .today-record-card .summary-value {
            font-size: 24px;
            line-height: 1.4;
            font-weight: 700;
            word-break: break-word;
        }

        .today-record-card .summary-box {
            padding: 20px 16px;
        }
        </style>
    </head>

    <body>

    <div class="page">

        <div class="card">

            <h1 class="page-title">📋 今日签到记录</h1>
            <div class="page-subtitle">
                查看今日所有签到与签退记录
            </div>

            {% if today_records %}

                {% for r in today_records %}
                <div class="card today-record-card">

                    <h2 class="section-title">
                        {{ r.get('姓名','') }}
                    </h2>

                    <div class="summary-grid">

                        <div class="summary-box">
                            <div class="summary-title">岗位</div>
                            <div class="summary-value">
                                {{ role_label(r.get('岗位','')) }}
                            </div>
                        </div>

                        <div class="summary-box">
                            <div class="summary-title">时间</div>
                            <div class="summary-value">
                                {{ r.get('开始时间','') }}
                                {% if r.get('结束时间') %}
                                    ~ {{ r.get('结束时间','') }}
                                {% else %}
                                    ~ 未签退
                                {% endif %}
                            </div>
                        </div>

                        <div class="summary-box">
                            <div class="summary-title">时数</div>
                            <div class="summary-value">
                                {{ r.get('时数','') or '-' }}
                            </div>
                        </div>

                        {% if r.get('card_no') %}
                        <div class="summary-box">
                            <div class="summary-title">卡号</div>
                            <div class="summary-value">
                                {{ r.get('card_no') }}
                            </div>
                        </div>
                        {% endif %}

                    </div>

                    <div class="btn-row">
                        <a class="btn-tool btn-warning"
                           href="{{ url_for('edit_page', row_number=r.get('_row')) }}">
                            修改记录
                        </a>
                    </div>

                </div>
                {% endfor %}

            {% else %}

                <div class="empty-state">
                    今天还没有签到记录
                </div>

            {% endif %}

            <div class="btn-row">
                <a class="btn-tool btn-secondary"
                   href="{{ url_for('admin.admin_report', pin=pin) }}">
                    返回管理员
                </a>
            </div>

        </div>

    </div>

    </body>
    </html>
    """,
    pin=pin,
    today_records=get_today_records(limit=200),
    role_label=role_label,
    t=get_text()
    )

@admin_bp.route("/download_data")
def download_data():

    try:
        start_date = request.args.get("start_date", "").strip()
        end_date = request.args.get("end_date", "").strip()

        if start_date and end_date:
            rows = db_query("""
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
                where date between %s and %s
                order by date, name, start_time
            """, (start_date, end_date), fetchall=True)

            file_date_text = f"{start_date}_to_{end_date}"

        else:
            rows = db_query("""
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

            file_date_text = "all"

        df = pd.DataFrame(rows)

        output = io.BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(
                writer,
                index=False,
                sheet_name="attendance"
            )

        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name=f"old_attendance_{file_date_text}.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        return f"下载旧签到失败：{e}"
    
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
    <!doctype html>
    <html lang="zh">
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>修改记录</title>

        <link rel="stylesheet" href="{{ url_for('static', filename='css/toolbox.css') }}">
    </head>

    <body>

    <div class="page">

        <div class="card">

            <h1 class="page-title">✏️ 修改记录</h1>
            <div class="page-subtitle">
                修改义工签到资料
            </div>

            <div class="summary-grid">
                <div class="summary-box">
                    <div class="summary-title">姓名</div>
                    <div class="summary-value">{{ row.name }}</div>
                </div>

                <div class="summary-box">
                    <div class="summary-title">日期</div>
                    <div class="summary-value">{{ row.date }}</div>
                </div>
            </div>

            <form method="post">
                <input type="hidden" name="pin" value="{{ pin }}">

                <div class="form-group">
                    <label class="form-label">岗位</label>
                    <input
                        class="form-input"
                        name="role"
                        value="{{ row.role or '' }}"
                    >
                </div>

                <div class="form-group">
                    <label class="form-label">开始时间</label>
                    <input
                        class="form-input"
                        name="start_time"
                        value="{{ row.start_time or '' }}"
                    >
                </div>

                <div class="form-group">
                    <label class="form-label">结束时间</label>
                    <input
                        class="form-input"
                        name="end_time"
                        value="{{ row.end_time or '' }}"
                    >
                </div>

                <div class="form-group">
                    <label class="form-label">备注</label>
                    <input
                        class="form-input"
                        name="remark"
                        value="{{ row.remark or '' }}"
                    >
                </div>

                <div class="btn-row">
                    <button class="btn-tool btn-success" type="submit">
                        保存修改
                    </button>

                    <a class="btn-tool btn-secondary"
                       href="{{ url_for('admin.admin_records', pin=pin) }}">
                        返回
                    </a>
                </div>

            </form>

        </div>

    </div>

    </body>
    </html>
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


@admin_bp.route("/admin_report", methods=["GET", "POST"])
def admin_report():
    pin = (
        request.form.get("admin_pin")
        or request.args.get("pin")
        or ""
    ).strip()

    t = get_text()

    if pin != ADMIN_PIN:
        flash(t["admin_pin_wrong"], "bad")
        return redirect(url_for("index"))

    return render_template_string("""
<!doctype html>
<html lang="{{ t['html_lang'] }}">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{{ t["admin_title"] }}</title>

<link rel="stylesheet"
      href="{{ url_for('static', filename='css/toolbox.css') }}">

<style>
.admin-page {
    max-width: 760px;
}

.admin-tool-stack {
    display: grid;
    gap: 14px;
}

.admin-page .section-title {
    font-size: 26px;
    margin-bottom: 16px;
}

.admin-page .btn-tool {
    min-height: 60px;
    font-size: 22px;
}

.admin-desc {
    font-size: 17px;
    color: #666;
    margin-top: 8px;
    line-height: 1.5;
}
</style>
</head>

<body>

<div class="page admin-page">

    <h1 class="page-title">🔐 {{ t["admin_title"] }}</h1>
    <p class="page-subtitle">
        {{ t.get("admin_subtitle", "签到管理、数据中心与管理员工具") }}
    </p>

    <div class="card">
        <div class="section-title">📋 {{ t.get("today_attendance_manage", "今日签到管理") }}</div>

        <div class="admin-tool-stack">
            <a class="btn-tool btn-danger"
               href="{{ url_for('admin.admin_today_open', pin=pin) }}">
                {{ t.get("today_open_records", "今日进行中（未签退）") }}
            </a>

            <a class="btn-tool btn-primary"
               href="{{ url_for('admin.admin_today_records', pin=pin) }}">
                {{ t.get("view_today_records", "查看今日记录") }}
            </a>
        </div>
    </div>

    <div class="card">
        <div class="section-title">📊 {{ t.get("data_center", "数据中心") }}</div>

        <div class="admin-tool-stack">
            <a class="btn-tool btn-success"
               href="{{ url_for('admin.data_center', pin=pin) }}">
                {{ t.get("enter_data_center", "进入数据中心") }}
            </a>
        </div>

        <div class="admin-desc">
            {{ t.get("data_center_desc", "下载签到、白话共修及未来各系统资料。") }}
        </div>
    </div>

    <div class="card">
        <div class="section-title">🛠 {{ t["admin_tools"] }}</div>

        <div class="admin-tool-stack">
            <a class="btn-tool btn-warning"
               href="{{ url_for('admin.admin_add_record', pin=pin) }}">
                {{ t["admin_add_record"] }}
            </a>

            <a class="btn-tool btn-purple"
               href="{{ url_for('admin.admin_records', pin=pin) }}">
                {{ t["admin_records"] }}
            </a>
        </div>
    </div>

    <div class="card">
        <a class="btn-tool btn-secondary"
           href="/">
            {{ t["back_home"] }}
        </a>
    </div>

</div>

</body>
</html>
""", t=t, pin=pin)


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
    <!doctype html>
    <html lang="zh">
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>今日记录管理</title>

        <link rel="stylesheet"
              href="{{ url_for('static', filename='css/toolbox.css') }}">
    </head>

    <body>

    <div class="page">

        <div class="card">

            <h1 class="page-title">✏️ 今日记录管理</h1>
            <div class="page-subtitle">
                管理今日签到记录
            </div>

            <div class="btn-row">
                <a class="btn-tool btn-secondary"
                   href="{{ url_for('admin.admin_report', pin=pin) }}">
                    返回管理员
                </a>
            </div>

            {% if rows %}

            <div class="table-responsive">
                <table class="record-table">
                    <thead>
                        <tr>
                            <th>姓名</th>
                            <th>岗位</th>
                            <th>开始</th>
                            <th>结束</th>
                            <th>时数</th>
                            <th>操作</th>
                        </tr>
                    </thead>

                    <tbody>
                        {% for r in rows %}
                        <tr>
                            <td>{{ r.name }}</td>
                            <td>{{ r.role }}</td>
                            <td>{{ r.start_time }}</td>
                            <td>{{ r.end_time }}</td>
                            <td>{{ r.hours }}</td>
                            <td>
                                <div class="btn-row">
                                    <a class="btn-tool btn-warning"
                                       href="{{ url_for('admin.admin_edit_record', record_id=r.id, pin=pin) }}">
                                        修改
                                    </a>

                                    <a class="btn-tool btn-danger"
                                       href="{{ url_for('admin.admin_delete_record', record_id=r.id, pin=pin) }}"
                                       onclick="return confirm('确定删除这笔记录吗？');">
                                        删除
                                    </a>
                                </div>
                            </td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>

            {% else %}

            <div class="empty-state">
                今日还没有签到记录。
            </div>

            {% endif %}

        </div>

    </div>

    </body>
    </html>
    """, rows=rows, pin=pin)


@admin_bp.route("/data_center")
def data_center():
    pin = request.args.get("pin", "").strip()
    t = get_text()
    today = date.today().isoformat()
    month_start = date.today().replace(day=1).isoformat()

    start_date = request.args.get("start_date", month_start).strip()
    end_date = request.args.get("end_date", today).strip()

    if pin != ADMIN_PIN:
        flash(t["admin_pin_wrong"], "bad")
        return redirect(url_for("index"))
    
    attendance_count_row = db_query("""
        select count(*) as total
        from attendance
        where date between %s and %s
    """, (start_date, end_date), fetchone=True)

    reading_count_row = db_query("""
        select count(*) as total
        from reading
        where date between %s and %s
    """, (start_date, end_date), fetchone=True)

    attendance_count = attendance_count_row["total"] if attendance_count_row else 0
    reading_count = reading_count_row["total"] if reading_count_row else 0

    return render_template_string("""
<!doctype html>
<html lang="{{ t['html_lang'] }}">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>📊 {{ t.get("data_center", "数据中心") }}</title>

<link rel="stylesheet" href="{{ url_for('static', filename='css/toolbox.css') }}">

<style>
.data-page{ max-width:820px; }
.data-grid{ display:grid; gap:16px; }

.quick-row{
    display:grid;
    grid-template-columns:1fr 1fr 1fr;
    gap:10px;
}

.form-row{
    display:grid;
    grid-template-columns:1fr 1fr;
    gap:14px;
}

.download-card{
    display:grid;
    gap:12px;
}

.small-desc{
    color:#666;
    font-size:16px;
    line-height:1.5;
}

.disabled-box{
    background:#f6f7f9;
    color:#888;
    padding:16px;
    border-radius:16px;
    text-align:center;
    font-size:17px;
}

.mini-btn-grid{
    display:grid;
    grid-template-columns:1fr 1fr;
    gap:12px;
}

.mini-btn-grid .btn-tool{
    min-height:50px;
    font-size:18px;
}

a.btn-tool,
.download-link{
    cursor:pointer !important;
    text-decoration:none !important;
}

a.btn-tool:hover,
.download-link:hover{
    transform:translateY(-1px);
    filter:brightness(0.98);
}

button:disabled{
    cursor:not-allowed !important;
    opacity:.65;
}

@media(max-width:600px){
    .mini-btn-grid,
    .quick-row,
    .form-row{
        grid-template-columns:1fr;
    }
}
</style>
</head>

<body>
<div class="page data-page">

    <div class="card">
        <h1 class="page-title">📊 {{ t.get("data_center", "数据中心") }}</h1>
        <p class="page-subtitle">
            {{ t.get("data_center_desc", "先选择日期范围，再选择要下载的系统资料。") }}
        </p>
    </div>

    <div class="card">
        <div class="section-title">📅 {{ t.get("download_date_range", "下载日期范围") }}</div>

        <div class="quick-row">
            <button type="button" class="btn-tool btn-primary" onclick="setToday()">
                📌 {{ t.get("download_today", "今天") }}
            </button>

            <button type="button" class="btn-tool btn-success" onclick="setThisWeek()">
                🗓 {{ t.get("download_this_week", "本星期") }}
            </button>

            <button type="button" class="btn-tool btn-warning" onclick="setThisMonth()">
                📆 {{ t.get("download_this_month", "本月") }}
            </button>
        </div>

        <div style="height:16px;"></div>

        <div class="form-row">
            <div class="form-group">
                <label class="form-label">{{ t.get("start_date", "开始日期") }}</label>
                <input class="form-input" type="date" id="start_date">
            </div>

            <div class="form-group">
                <label class="form-label">{{ t.get("end_date", "结束日期") }}</label>
                <input class="form-input" type="date" id="end_date">
            </div>
        </div>

        <div class="alert alert-info">
            {{ t.get("date_range_tip", "选择日期后，下面的下载按钮会自动带入这个日期范围。") }}
        </div>
                                  
        <div class="alert alert-success" id="range_text">
            当前下载范围：-
        </div>
    </div>
                                  
    <div class="card">

        <div class="section-title">
            📈 当前资料概览
        </div>

        <div class="summary-grid">

            <div class="summary-box">
                <div class="summary-title">
                    👥 签到记录
                </div>

                <div class="summary-value">
                    {{ attendance_count }}
                </div>
            </div>

            <div class="summary-box">
                <div class="summary-title">
                    📖 共修
                </div>

                <div class="summary-value">
                    {{ reading_count }}
                </div>
            </div>

            <div class="summary-box">
                <div class="summary-title">
                    📅 排班
                </div>

                <div class="summary-value">
                    --
                </div>
            </div>

            <div class="summary-box">
                <div class="summary-title">
                    💰 财政
                </div>

                <div class="summary-value">
                    --
                </div>
            </div>

        </div>

    </div>

    <div class="card">
        <div class="section-title">📥 {{ t.get("select_download_item", "选择下载项目") }}</div>

        <div class="data-grid">

            <div class="download-card">
                <h2 class="section-title">✅ {{ t.get("signin_system", "签到系统") }}</h2>
                <div class="small-desc">
                    {{ t.get("signin_data_desc", "下载签到记录。") }}
                </div>

                <a class="btn-tool btn-warning download-link"
                   data-base="{{ url_for('admin.download_data') }}">
                    📊 {{ t.get("download_old_signin_excel", "下载旧签到 Excel") }}
                </a>

                <a class="btn-tool btn-secondary download-link"
                   data-base="{{ url_for('admin.download_att_logs') }}">
                    📋 {{ t.get("download_new_attendance_logs", "下载新签到日志") }}
                </a>
            </div>

            <hr>

            <div class="download-card">
                <h2 class="section-title">{{ t.get("reading_system", "白话佛法共修") }}</h2>
                <div class="small-desc">
                    {{ t.get("reading_data_desc", "下载白话佛法共修记录。") }}
                </div>

                <a class="btn-tool btn-primary download-link"
                   data-base="/download_reading">
                    {{ t.get("download_reading_excel", "下载共修记录 Excel") }}
                </a>
            </div>

            <hr>

            <div class="download-card">
                <h2 class="section-title">📅 {{ t.get("schedule_system", "排班系统") }}</h2>
                <div class="small-desc">
                    {{ t.get("schedule_data_desc", "下载报名资料、正式排班、排班签到与缺人工统计。") }}
                </div>

                <div class="mini-btn-grid">
                    <a class="btn-tool btn-primary download-link"
                    data-base="{{ url_for('admin.download_schedule_signups') }}">
                        📝 报名资料
                    </a>

                    <a class="btn-tool btn-success download-link"
                    data-base="{{ url_for('admin.download_schedule_assignments') }}">
                        📋 正式排班
                    </a>
                    
                </div>
            </div>

            <hr>
            
        </div>
    </div>

    <div class="card">
        <a class="btn-tool btn-secondary"
           href="{{ url_for('admin.admin_report') }}?pin={{ pin }}">
            {{ t.get("back_admin_home", "⬅ 返回管理员首页") }}
        </a>
    </div>

</div>

<script>
function formatDate(d){
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return `${y}-${m}-${day}`;
}

function setToday(){
    const today = new Date();
    document.getElementById("start_date").value = formatDate(today);
    document.getElementById("end_date").value = formatDate(today);
    updateLinks();
}

function setThisWeek(){
    const today = new Date();
    const day = today.getDay();
    const diff = day === 0 ? 6 : day - 1;

    const monday = new Date(today);
    monday.setDate(today.getDate() - diff);

    document.getElementById("start_date").value = formatDate(monday);
    document.getElementById("end_date").value = formatDate(today);
    updateLinks();
}

function setThisMonth(){
    const today = new Date();
    const firstDay = new Date(today.getFullYear(), today.getMonth(), 1);

    document.getElementById("start_date").value = formatDate(firstDay);
    document.getElementById("end_date").value = formatDate(today);
    updateLinks();
}

function updateLinks(){
    const startDate = document.getElementById("start_date").value;
    const endDate = document.getElementById("end_date").value;

    const rangeText = document.getElementById("range_text");

    if(startDate && endDate){
        rangeText.innerText = `当前下载范围：${startDate} 至 ${endDate}`;
    }else{
        rangeText.innerText = "当前下载范围：未选择";
    }

    document.querySelectorAll(".download-link").forEach(function(link){
        const base = link.dataset.base;

        if(startDate && endDate){
            link.href = `${base}?start_date=${startDate}&end_date=${endDate}`;
        }else{
            link.href = base;
        }
    });
}

document.getElementById("start_date").addEventListener("change", updateLinks);
document.getElementById("end_date").addEventListener("change", updateLinks);

setToday();
</script>

</body>
</html>
""",
    t=t,
    pin=pin,
    start_date=start_date,
    end_date=end_date,
    attendance_count=attendance_count,
    reading_count=reading_count
)

ADMIN_HOME_HTML = """
<!doctype html>
<html lang="{{ t.get('html_lang', 'zh') }}">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">

<link rel="manifest" href="/admin-manifest.json">

<title>{{ t.get("admin_center_title", "观音堂管理员中心") }}</title>

<link rel="stylesheet"
      href="{{ url_for('static', filename='css/toolbox.css') }}">

<style>
.admin-page{
    max-width:820px;
}

.admin-grid{
    display:grid;
    grid-template-columns:1fr 1fr;
    gap:18px;
}

.admin-card-desc{
    color:#666;
    font-size:18px;
    line-height:1.6;
    margin:8px 0 16px;
}

.admin-page .btn-tool{
    min-height:58px;
    font-size:22px;
}

@media (max-width:700px){
    .admin-grid{
        grid-template-columns:1fr;
    }
}
</style>
</head>

<body>

<div class="page admin-page">

    <div class="card" style="text-align:center;">
        <h1 class="page-title">
            🙏 {{ t.get("admin_center_title", "观音堂管理员中心") }}
        </h1>

        <div class="page-subtitle">
            Temple Administration Center
        </div>
    </div>

    <div class="admin-grid">

        <div class="card">
            <h2 class="section-title">📋 签到系统</h2>
            <div class="admin-card-desc">
                今日签到、签退、补录、修改及删除签到记录。
            </div>
            <a class="btn-tool btn-purple"
               href="{{ url_for('admin.admin_report', pin=pin) }}">
                进入签到系统
            </a>
        </div>

        <div class="card">
            <h2 class="section-title">📅 排班系统</h2>
            <div class="admin-card-desc">
                义工报名、排班、发布值班表及查看安排。
            </div>
            <a class="btn-tool btn-primary"
               href="/schedule/admin?pin={{ pin }}">
                进入排班系统
            </a>
        </div>

        <div class="card">
            <h2 class="section-title">💰 财政系统</h2>
            <div class="admin-card-desc">
                月费收款、财政记录、银行过账及 Dashboard。
            </div>
            <a class="btn-tool btn-success"
               href="/finance/">
                进入财政系统
            </a>
        </div>

        <div class="card">
            <h2 class="section-title">📿 月费管理</h2>
            <div class="admin-card-desc">
                月费资料管理、查询佛友已供养月份。
            </div>
            <a class="btn-tool btn-warning"
               href="/member/admin">
                进入月费管理
            </a>
        </div>

        <div class="card">
            <h2 class="section-title">📚 藏经阁系统</h2>
            <div class="admin-card-desc">
                法宝查询、入库、出库、库存及藏经阁管理。
            </div>
            <a class="btn-tool btn-success"
               href="/library/">
                进入藏经阁系统
            </a>
        </div>

        <div class="card">
            <h2 class="section-title">📖 白话佛法共修</h2>
            <div class="admin-card-desc">
                记录每日共修、查看今日共修记录及统计。
            </div>
            <a class="btn-tool btn-primary"
               href="/reading">
                进入共修记录
            </a>
        </div>

        <div class="card">
            <h2 class="section-title">📊 数据中心</h2>
            <div class="admin-card-desc">
                统一下载签到、共修、排班、藏经阁、财政资料。
            </div>
            <a class="btn-tool btn-purple"
               href="{{ url_for('admin.data_center', pin=pin) }}">
                进入数据中心
            </a>
        </div>

        <div class="card">
            <h2 class="section-title">⚙ 系统设置</h2>
            <div class="admin-card-desc">
                今日签到码、系统设定、语言及备份。
            </div>
            <a class="btn-tool btn-secondary"
               href="#">
                开发中
            </a>
        </div>

    </div>

    <div class="card">
        <a class="btn-tool btn-secondary"
           href="/">
            返回首页
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
        t=get_text(),
        today_code=""
    )

@admin_bp.route("/download_att_logs")
def download_att_logs():

    try:
        start_date = request.args.get("start_date", "").strip()
        end_date = request.args.get("end_date", "").strip()

        use_date_filter = bool(start_date and end_date)

        if use_date_filter:
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
                where attendance_date between %s and %s
                order by attendance_date, name, check_in_time
            """, (start_date, end_date), fetchall=True)

            file_date_text = f"{start_date}_to_{end_date}"

        else:
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

            file_date_text = "all"

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
            download_name=f"att_logs_{file_date_text}.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        return f"下载签到日志失败：{e}"
    
@admin_bp.route("/download_schedule_signups")
def download_schedule_signups():

    try:
        start_date = request.args.get("start_date", "").strip()
        end_date = request.args.get("end_date", "").strip()

        use_date_filter = bool(start_date and end_date)

        if use_date_filter:

            rows = db_query("""
                select
                    signup_date as "日期",
                    volunteer_id as "编号",
                    name as "姓名",
                    role as "岗位",
                    start_time as "开始时间",
                    end_time as "结束时间",
                    status as "状态",
                    remarks as "备注"
                from volunteer_schedule_signups
                where signup_date between %s and %s
                order by signup_date,name,start_time
            """,(start_date,end_date),fetchall=True)

            filename = f"schedule_signups_{start_date}_to_{end_date}.xlsx"

        else:

            rows = db_query("""
                select
                    signup_date as "日期",
                    volunteer_id as "编号",
                    name as "姓名",
                    role as "岗位",
                    start_time as "开始时间",
                    end_time as "结束时间",
                    status as "状态",
                    remarks as "备注"
                from volunteer_schedule_signups
                order by signup_date,name,start_time
            """,fetchall=True)

            filename = "schedule_signups_all.xlsx"

        df = pd.DataFrame(rows)

        output = io.BytesIO()

        with pd.ExcelWriter(output,engine="openpyxl") as writer:
            df.to_excel(
                writer,
                index=False,
                sheet_name="Signups"
            )

        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        return f"下载报名资料失败：{e}"
    
@admin_bp.route("/download_schedule_assignments")
def download_schedule_assignments():

    try:
        start_date = request.args.get("start_date", "").strip()
        end_date = request.args.get("end_date", "").strip()

        use_date_filter = bool(start_date and end_date)

        if use_date_filter:

            rows = db_query("""
                select
                    assignment_date as "日期",
                    volunteer_id as "编号",
                    name as "姓名",
                    shift_label as "班别",
                    assigned_place as "岗位",
                    start_time as "开始时间",
                    end_time as "结束时间",
                    status as "状态",
                    remarks as "备注"
                from volunteer_schedule_assignments
                where assignment_date between %s and %s
                order by assignment_date,shift_label,name,start_time
            """,(start_date,end_date),fetchall=True)

            filename = f"schedule_assignments_{start_date}_to_{end_date}.xlsx"

        else:

            rows = db_query("""
                select
                    assignment_date as "日期",
                    volunteer_id as "编号",
                    name as "姓名",
                    shift_label as "班别",
                    assigned_place as "岗位",
                    start_time as "开始时间",
                    end_time as "结束时间",
                    status as "状态",
                    remarks as "备注"
                from volunteer_schedule_assignments
                order by assignment_date,shift_label,name,start_time
            """,fetchall=True)

            filename = "schedule_assignments_all.xlsx"

        df = pd.DataFrame(rows)

        output = io.BytesIO()

        with pd.ExcelWriter(output,engine="openpyxl") as writer:
            df.to_excel(
                writer,
                index=False,
                sheet_name="Assignments"
            )

        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        return f"下载正式排班失败：{e}"