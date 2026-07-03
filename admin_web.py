# admin_web.py

import io
import pandas as pd


from datetime import datetime
from io import BytesIO
from db import db_query
from utils import get_text
from openpyxl import Workbook
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
        old_att_rows = db_query("""
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

        att_logs = db_query("""
            select
                attendance_date as "日期",
                volunteer_id as "编号",
                name as "姓名",
                actual_role as "岗位",
                actual_place as "地点",
                to_char(check_in_time, 'HH12:MIam') as "签到时间",
                to_char(check_out_time, 'HH12:MIam') as "签退时间",
                case when walk_in then '是' else '' end as "现场签到",
                remarks as "备注"
            from volunteer_attendance_logs
            order by attendance_date, name, check_in_time
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

        old_df = pd.DataFrame(old_att_rows)
        logs_df = pd.DataFrame(att_logs)

        if not logs_df.empty:
            new_att_df = pd.DataFrame({
                "日期": logs_df.get("日期", ""),
                "编号": logs_df.get("编号", ""),
                "姓名": logs_df.get("姓名", ""),
                "报名": logs_df["现场签到"].apply(lambda x: 0 if str(x).strip() == "是" else 1),
                "签到": 1,
                "岗位": logs_df.get("岗位", ""),
                "开始时间": logs_df.get("签到时间", ""),
                "结束时间": logs_df.get("签退时间", ""),
                "时数": "",
                "备注": logs_df.get("备注", ""),
            })
        else:
            new_att_df = pd.DataFrame(columns=[
                "日期", "编号", "姓名", "报名", "签到",
                "岗位", "开始时间", "结束时间", "时数", "备注"
            ])

        attendance_df = pd.concat(
            [old_df, new_att_df],
            ignore_index=True
        )

        if not attendance_df.empty:
            attendance_df = attendance_df.drop_duplicates(
                subset=["日期", "编号", "姓名", "岗位", "开始时间"],
                keep="last"
            )

            attendance_df = attendance_df.sort_values(
                by=["日期", "姓名", "开始时间"],
                ascending=[True, True, True]
            )

        output = io.BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:

            attendance_df.to_excel(
                writer,
                index=False,
                sheet_name="attendance"
            )

            logs_df.to_excel(
                writer,
                index=False,
                sheet_name="att_logs"
            )

            pd.DataFrame(vol_rows).to_excel(
                writer,
                index=False,
                sheet_name="volunteers"
            )

            pd.DataFrame(reading_rows).to_excel(
                writer,
                index=False,
                sheet_name="reading"
            )

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


@admin_bp.route("/admin_report", methods=["GET", "POST"])
def admin_report():
    pin = (
        request.form.get("admin_pin")
        or request.args.get("pin")
        or ""
    ).strip()

    print("输入 pin =", repr(pin))
    print("系统 ADMIN_PIN =", repr(ADMIN_PIN))

    t = get_text()

    if pin != ADMIN_PIN:
        flash(t["admin_pin_wrong"], "bad")
        return redirect(url_for("index"))

    code = get_display_today_code()

    return render_template_string("""
<!doctype html>
<html lang="zh">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>管理员工具</title>

<link rel="stylesheet"
      href="{{ url_for('static', filename='css/toolbox.css') }}">
<style>
.admin-page{
    max-width:760px;
}

.today-code-box{
    font-size:56px;
    font-weight:bold;
    color:#dc2626;
    text-align:center;
    padding:22px;
    background:#fff7d6;
    border-radius:18px;
    border:1px solid #fde68a;
}

.admin-tool-stack{
    display:grid;
    gap:14px;
}

.admin-page .section-title{
    font-size:26px;
    margin-bottom:16px;
}

.admin-page .btn-tool{
    min-height:58px;
    font-size:22px;
}
</style>
</head>

<body>

<div class="page admin-page">

    <h1 class="page-title">🔐 {{ t["admin_title"] }}</h1>
    <p class="page-subtitle">签到管理、今日记录与管理员工具</p>

    <div class="card">
        <div class="section-title">{{ t["today_code_big"] }}</div>

        <div class="today-code-box">
            {{ code }}
        </div>

        <div class="alert alert-warning">
            ⚠ 请只写在观音堂现场，不要发群
        </div>
    </div>

    <div class="card">
        <div class="section-title">📋 今日签到管理</div>

        <div class="admin-tool-stack">
            <a class="btn-tool btn-danger"
               href="/admin_today_open?pin={{ pin }}">
                今日进行中（未签退）
            </a>

            <a class="btn-tool btn-primary"
               href="/admin_today_records?pin={{ pin }}">
                查看今日记录
            </a>
        </div>
    </div>

    <div class="card">
        <div class="section-title">🛠 管理工具</div>

        <div class="admin-tool-stack">
            <a class="btn-tool btn-success"
               href="/download_data">
                下载签到日志 Excel
            </a>

            <a class="btn-tool btn-warning"
               href="/admin_add_record?pin={{ pin }}">
                补录签到
            </a>

            <a class="btn-tool btn-purple"
               href="/admin_records?pin={{ pin }}">
                修改 / 删除今日记录
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
""", t=t, code=code, pin=pin)

ADMIN_HOME_HTML = """
<!doctype html>
<html lang="zh">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">

<link rel="manifest" href="/manifest.json">

<title>观音堂管理员入口</title>

<link rel="stylesheet"
      href="{{ url_for('static', filename='css/toolbox.css') }}">

</head>

<body>

<div class="page">

    <div class="card">

        <h1 class="page-title">
            🙏 观音堂管理员入口
        </h1>

        <div class="page-subtitle">
            Temple Administration Center
        </div>

    </div>

    <div class="card">

        <h2 class="section-title">
            📅 今日签到码
        </h2>

        <div class="alert alert-warning"
             style="text-align:center;
                    font-size:60px;
                    font-weight:bold;
                    letter-spacing:6px;">
            {{ today_code }}
        </div>

    </div>

    <div class="card">

        <h2 class="section-title">
            📋 义工报名系统
        </h2>

        <div class="btn-row">
            <a class="btn-tool btn-primary"
               href="/volunteer">
                进入系统
            </a>
        </div>

    </div>

    <div class="card">

        <h2 class="section-title">
            💰 月费管理员系统
        </h2>

        <div class="btn-row">
            <a class="btn-tool btn-success"
               href="/member">
                进入系统
            </a>
        </div>

    </div>

    <div class="card">

        <h2 class="section-title">
            ✅ 义工签到系统
        </h2>

        <div class="btn-row">
            <a class="btn-tool btn-purple"
               href="/">
                进入系统
            </a>
        </div>

    </div>

    <div class="card">

        <h2 class="section-title">
            📥 数据下载
        </h2>

        <div class="btn-row">

            <a class="btn-tool btn-warning"
               href="{{ url_for('admin.download_data') }}">
                下载旧签到 Excel
            </a>

            <a class="btn-tool btn-secondary"
               href="{{ url_for('admin.download_att_logs') }}">
                下载新签到日志
            </a>

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

@admin_bp.route("/download_att_logs")
def download_att_logs():

    try:
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
            download_name="att_logs.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        return f"下载签到日志失败：{e}"