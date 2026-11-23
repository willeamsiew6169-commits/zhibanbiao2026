# admin_web.py

import io
import pandas as pd

from io import BytesIO
from signin_policy import get_signin_policies, set_signin_policy
from db import db_query, get_conn
from utils import get_text
from openpyxl import Workbook
from datetime import date, datetime
from utils import now_date_str, get_text
from utils import get_display_today_code
from psycopg2.extras import RealDictCursor
from attendance_service import auto_signout_unfinished_today, _assignment_ids_from_remark
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
    """管理员补录历史出席；保留原 URL 和管理员 PIN 校验。"""
    from attendance_service import ROLES
    from utils import calc_hours
    from datetime import date as date_type

    pin = request.args.get("pin") or request.form.get("pin")
    if pin != ADMIN_PIN:
        return "无权限", 403

    volunteers = db_query(
        "select id, name from volunteers where status='在册' order by id",
        fetchall=True
    ) or []
    chosen_date = request.form.get("date") or now_date_str()
    chosen_id = request.form.get("id", "").strip()
    records = []
    if chosen_id and chosen_date:
        records = db_query("""
            select id, role, start_time, end_time, hours, remark
            from attendance where date=%s and volunteer_id=%s order by id
        """, (chosen_date, chosen_id), fetchall=True) or []

    if request.method == "POST" and request.form.get("action") == "save":
        role = request.form.get("role", "").strip()
        start = request.form.get("start", "").strip()
        end = request.form.get("end", "").strip()
        reason = request.form.get("reason", "").strip()
        try:
            duty_date = date_type.fromisoformat(chosen_date)
            if duty_date > date_type.fromisoformat(now_date_str()):
                raise ValueError("不能补录未来日期")
            if role not in ROLES or not chosen_id or not reason:
                raise ValueError("请选择义工、岗位并填写补录原因")
            if not start or not end:
                raise ValueError("请填写实际开始和结束时间")
            from datetime import datetime as dt
            def parse_clock(s):
                for fmt in ("%H:%M", "%I:%M%p", "%I:%M %p"):
                    try: return dt.strptime(s.strip(), fmt).strftime("%H:%M")
                    except ValueError: pass
                raise ValueError("时间格式不正确")
            start24, end24 = parse_clock(start), parse_clock(end)
            if end24 <= start24:
                raise ValueError("结束时间必须晚于开始时间")
            hours = round((int(end24[:2])*60+int(end24[3:]) - int(start24[:2])*60-int(start24[3:]))/60, 2)
            if hours <= 0 or hours > 18:
                raise ValueError("实际时数不合理，请核对")
            with get_conn() as conn:
                with conn.cursor(cursor_factory=RealDictCursor) as cur:
                    cur.execute("select id, name from volunteers where id=%s and status='在册'", (chosen_id,))
                    volunteer = cur.fetchone()
                    if not volunteer:
                        raise ValueError("找不到在册义工，请重新选择")
                    cur.execute("""select * from attendance where date=%s and volunteer_id=%s
                                   and role=%s order by id for update""", (chosen_date, chosen_id, role))
                    existing = cur.fetchall() or []
                    if len(existing) > 1:
                        raise ValueError("当天此岗位有多笔记录，请先到历史记录核对，不能自动补录")
                    if existing and existing[0].get('end_time'):
                        raise ValueError("当天此岗位已有完成记录，请到历史记录修改，不重复新增")
                    note = "管理员补录｜原因：" + reason
                    if existing:
                        old = existing[0]
                        note = str(old.get('remark') or '') + '｜' + note + '｜补签退'
                        cur.execute("""update attendance set start_time=%s,end_time=%s,hours=%s,remark=%s
                                       where id=%s""", (start24,end24,hours,note,old['id']))
                    else:
                        cur.execute("""insert into attendance
                                       (date,volunteer_id,name,signup,signin,role,start_time,end_time,hours,remark)
                                       values (%s,%s,%s,1,1,%s,%s,%s,%s,%s)""",
                                    (chosen_date,chosen_id,volunteer['name'],role,start24,end24,hours,note))
                    # 值班可有多个合法排班分段日志；不能因为多笔日志而撤销补录。
                    # 已有日志属于既有签到/排班，不擅自覆盖或另造重复日志。
                    cur.execute("""select id, check_in_time, check_out_time
                                   from volunteer_attendance_logs
                                   where volunteer_id=%s and attendance_date=%s and actual_role=%s
                                   order by id for update""", (chosen_id, chosen_date, role))
                    logs = cur.fetchall() or []
                    if not logs:
                        # 仅完全没有对应岗位日志时，新增一笔真实历史时间的补录日志。
                        cur.execute("""insert into volunteer_attendance_logs
                                       (assignment_id,volunteer_id,name,attendance_date,actual_role,actual_place,
                                        walk_in,remarks,check_in_time,check_out_time)
                                       values (null,%s,%s,%s,%s,%s,true,%s,
                                               %s::date + %s::time, %s::date + %s::time)""",
                                    (chosen_id,volunteer['name'],chosen_date,role,role,note,
                                     chosen_date,start24,chosen_date,end24))
                    elif existing and len(logs) == 1 and logs[0]['check_out_time'] is None:
                        # 只有明确一笔未签退日志且 attendance 原本未签退时才补日志签退。
                        cur.execute("""update volunteer_attendance_logs
                                       set check_out_time=%s::date + %s::time,
                                           remarks=concat_ws('｜', nullif(remarks,''), %s)
                                       where id=%s and check_out_time is null""",
                                    (chosen_date,end24,note,logs[0]['id']))
                    # 多笔值班分段日志或已签退日志：保留原状，不误改、不重复新增。
            flash("补录成功；若该月份已生成报表，请重新生成月报。", "ok")
            return redirect(url_for("admin.admin_add_record", pin=pin, date=chosen_date, id=chosen_id))
        except (ValueError, TypeError) as exc:
            flash(str(exc), "bad")
        except Exception:
            flash("保存失败，数据库未完成写入；请检查日志。", "bad")

    chosen_date = request.args.get("date") or chosen_date
    chosen_id = request.args.get("id") or chosen_id
    if chosen_id:
        records = db_query("""select id,role,start_time,end_time,hours from attendance
                              where date=%s and volunteer_id=%s order by id""",
                           (chosen_date,chosen_id),fetchall=True) or []
    return render_template_string("""
<!doctype html><html lang="zh"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>补录签到</title><link rel="stylesheet" href="{{ url_for('static', filename='css/toolbox.css') }}">
<style>.backfill-input{min-height:52px;font-size:19px}.backfill-grid{display:grid;grid-template-columns:1fr 1fr;gap:12px}
@media(max-width:600px){.backfill-grid{grid-template-columns:1fr}}</style></head><body><div class="page"><div class="card">
<h1 class="page-title">📝 管理员补录签到</h1>
<nav style="display:flex;gap:8px;flex-wrap:wrap;margin:12px 0 18px">
<a style="padding:10px;border-radius:8px;background:#e8efff;text-decoration:none" href="{{ url_for('admin.admin_report',pin=pin) }}">← 返回管理员</a>
<a style="padding:10px;border-radius:8px;background:#e8efff;text-decoration:none" href="{{ url_for('admin.manage_history',pin=pin,kind='attendance',day=chosen_date) }}">签到记录</a>
<span aria-current="page" style="padding:10px;border-radius:8px;background:#c56b16;color:white">📝 补录签到</span>
<a style="padding:10px;border-radius:8px;background:#e8efff;text-decoration:none" href="{{ url_for('admin.manage_history',pin=pin,kind='reading',day=chosen_date) }}">白话佛法</a>
</nav><p>可补录过去任何日期；同岗位已有记录时不重复新增。</p>
{% with messages=get_flashed_messages(with_categories=true) %}{% for category,msg in messages %}
<div class="alert {{ 'alert-success' if category=='ok' else 'alert-danger' }}">{{ msg }}</div>
{% endfor %}{% endwith %}
<form method="get" id="lookupForm"><input type="hidden" name="pin" value="{{ pin }}">
<div class="form-group"><label class="form-label">日期</label><input class="form-input backfill-input" type="date" name="date" value="{{ chosen_date }}" max="{{ today }}" required></div>
<div class="form-group"><label class="form-label">义工编号／姓名</label>
<select class="form-input backfill-input" name="id" required onchange="document.getElementById('lookupForm').submit()">
<option value="">请选择义工</option>{% for v in volunteers %}<option value="{{ v.id }}" {{ 'selected' if v.id==chosen_id else '' }}>{{ v.name }}（{{ v.id }}）</option>{% endfor %}
</select></div><button class="btn-tool btn-primary" type="submit">🔎 查询当天记录</button></form>
{% if chosen_id %}<h2>📋 当天已有记录</h2>
{% if records %}{% for r in records %}<div class="alert alert-success">{{ r.role }}：{{ r.start_time or '未签到' }} ～ {{ r.end_time or '未签退' }}（{{ r.hours or 0 }}小时）</div>{% endfor %}
{% else %}<p>当天没有签到记录。</p>{% endif %}
<form method="post"><input type="hidden" name="pin" value="{{ pin }}"><input type="hidden" name="action" value="save"><input type="hidden" name="date" value="{{ chosen_date }}"><input type="hidden" name="id" value="{{ chosen_id }}">
<div class="form-group"><label class="form-label">补录岗位</label><select class="form-input backfill-input" name="role" required>
{% for r in roles %}<option value="{{ r }}">{{ '整理佛台' if r=='佛台' else r }}</option>{% endfor %}</select></div>
<div class="backfill-grid"><div class="form-group"><label class="form-label">实际开始</label><input class="form-input backfill-input" type="time" name="start" required></div>
<div class="form-group"><label class="form-label">实际结束</label><input class="form-input backfill-input" type="time" name="end" required></div></div>
<div class="form-group"><label class="form-label">补录原因</label><select class="form-input backfill-input" name="reason" required>
<option value="忘记签到">忘记签到</option><option value="忘记签退">忘记签退</option><option value="不会操作">不会操作</option><option value="负责人核实补录">负责人核实补录</option></select></div>
<button class="btn-tool btn-success" type="submit" onclick="return confirm('确认按填写时间补录这笔实际出席？')">✅ 确认补录</button></form>{% endif %}
<p><a class="btn-tool btn-secondary" href="{{ url_for('admin.manage_history',pin=pin,kind='attendance',day=chosen_date) }}">返回签到记录管理</a></p>
</div></div></body></html>
""", pin=pin, today=now_date_str(), chosen_date=chosen_date, chosen_id=chosen_id,
       volunteers=volunteers, records=records, roles=ROLES)

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
    """导出实际签到；attendance.id 是本地重复导入的稳定主键。"""
    if (request.args.get("pin") or "").strip() != ADMIN_PIN:
        return "无权限 / Unauthorized", 403

    try:
        start_date = request.args.get("start_date", "").strip()
        end_date = request.args.get("end_date", "").strip()
        if bool(start_date) != bool(end_date):
            return "请同时填写开始日期和结束日期", 400
        if start_date:
            try:
                begin = date.fromisoformat(start_date)
                finish = date.fromisoformat(end_date)
            except ValueError:
                return "日期格式必须是 YYYY-MM-DD", 400
            if begin > finish:
                return "开始日期不能晚于结束日期", 400

        columns = ["记录 ID", "日期", "编号", "姓名", "报名", "签到", "岗位",
                   "开始时间", "结束时间", "时数", "备注"]
        sql = """SELECT id AS "记录 ID", date AS "日期",
                        volunteer_id AS "编号", name AS "姓名",
                        signup AS "报名", signin AS "签到", role AS "岗位",
                        start_time AS "开始时间", end_time AS "结束时间",
                        hours AS "时数", remark AS "备注"
                 FROM attendance"""
        params = ()
        if start_date:
            sql += " WHERE date BETWEEN %s AND %s"
            params = (start_date, end_date)
        sql += " ORDER BY date, name, start_time, id"
        rows = db_query(sql, params, fetchall=True) or []
        df = pd.DataFrame(rows, columns=columns)  # 没记录也保留标题行

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="attendance")
        output.seek(0)
        date_text = f"{start_date}_to_{end_date}" if start_date else "all"
        return send_file(
            output, as_attachment=True,
            download_name=f"actual_attendance_{date_text}.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        return f"下载实际签到失败：{e}", 500

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

    policies = get_signin_policies()
    # 日历模块异常或尚未建表时，不影响管理员其他工具。
    try:
        calendar_campaign = db_query(
            "SELECT year, is_open FROM calendar_campaigns ORDER BY year DESC LIMIT 1",
            fetchone=True,
        )
    except Exception:
        calendar_campaign = None

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
.admin-home-float {
    position: fixed;
    right: 18px;
    bottom: 22px;
    z-index: 1000;
    display: inline-block;
    padding: 12px 18px;
    border-radius: 14px;
    background: #6b7280;
    color: white;
    text-decoration: none;
    font-size: 18px;
    font-weight: bold;
    box-shadow: 0 3px 12px #0002;
}
</style>
</head>

<body>

<div class="page admin-page">

    <h1 class="page-title">🔐 {{ t["admin_title"] }}</h1>
    <p class="page-subtitle">
        {{ "Attendance and record management" if t["html_lang"].startswith("en") else "签到与记录管理" }}
    </p>

    <div class="card">
        <div class="section-title">📋 {{ t.get("today_attendance_manage", "今日签到管理") }}</div>

        <div class="admin-tool-stack">
            <a class="btn-tool btn-danger"
               href="{{ url_for('admin.admin_today_open', pin=pin) }}">
                {{ t.get("today_open_records", "今日进行中（未签退）") }}
            </a>

            <a class="btn-tool btn-primary"
               href="{{ url_for('admin.manage_history', pin=pin, kind='attendance') }}">📋 记录管理（签到／白话佛法）</a>
        </div>
    </div>

    <div class="card">
        <div class="section-title">🔐 报名签到限制管理 / Check-in Rules</div>
        <p class="admin-desc">仅控制值班签到；佛台、卫生及膳食仍按各自资格判断。<br>
        Duty only; other roles retain their own eligibility rules.</p>
        {% for code, label, detail in [
          ('duty', '📋 值班签到模式 / Duty check-in mode', '测试模式允许无排班签到；正式模式必须有当天有效值班排班。 / Test mode allows walk-in duty; official mode requires a valid duty assignment.')
        ] %}
        <form method="post" action="{{ url_for('admin.update_signin_policy') }}"
              style="padding:12px 0;border-top:1px solid #eee;">
            <input type="hidden" name="role" value="{{ code }}">
            <input type="hidden" name="admin_pin" value="{{ pin }}">
            <div style="font-weight:700;font-size:19px">{{ label }}</div>
            <div class="admin-desc">{{ detail }}</div>
            <div style="margin:8px 0;font-weight:700;color:{{ '#198754' if policies[code] else '#b02a37' }}">
                {{ '🟢 正式模式 / Official' if policies[code] else '🔴 测试模式 / Test' }}
            </div>
            <button type="submit" name="enabled" value="{{ '0' if policies[code] else '1' }}"
                    class="btn-tool {{ 'btn-danger' if policies[code] else 'btn-success' }}">
                {{ '切换测试模式 / Test mode' if policies[code] else '切换正式模式 / Official mode' }}
            </button>
        </form>
        {% endfor %}
    </div>

    <div class="card">
        <div class="section-title">📅 日历领取管理 / Calendar</div>
        {% if calendar_campaign %}
        <p class="admin-desc">{{ calendar_campaign.year }}年：
          <strong>{{ '🟢 领取开放中' if calendar_campaign.is_open else '🔴 已关闭' }}</strong>
        </p>
        <form method="post" action="/calendar/admin/quick-toggle"
              onsubmit="return confirm({{ ('确定关闭日历领取？签到首页将隐藏日历卡片。' if calendar_campaign.is_open else '确定重新开放日历领取？')|tojson }});">
            <input type="hidden" name="admin_pin" value="{{ pin }}">
            <input type="hidden" name="year" value="{{ calendar_campaign.year }}">
            <button type="submit" class="btn-tool {{ 'btn-danger' if calendar_campaign.is_open else 'btn-success' }}">
              {{ '🔒 关闭日历领取' if calendar_campaign.is_open else '🟢 重新开放日历领取' }}
            </button>
        </form>
        {% else %}
        <p class="admin-desc">尚未建立日历年度。</p>
        {% endif %}
        <a class="btn-tool btn-primary" href="/calendar" style="margin-top:12px;">📋 进入日历管理（库存／名单／新年度）</a>
    </div>

    <a class="admin-home-float" href="/">🏠 {{ t["back_home"] }}</a>

</div>

</body>
</html>
""", t=t, pin=pin, policies=policies, calendar_campaign=calendar_campaign)


@admin_bp.route("/admin/signin-policy", methods=["POST"])
def update_signin_policy():
    pin = (request.form.get("admin_pin") or "").strip()
    if pin != ADMIN_PIN:
        flash("管理员 PIN 错误 / Invalid admin PIN", "bad")
        return redirect(url_for("index"))
    role = (request.form.get("role") or "").strip()
    if role not in ("duty", "cleaning", "meal"):
        flash("未知岗位 / Invalid role", "bad")
        return redirect(url_for("admin.admin_report", pin=pin))
    set_signin_policy(role, request.form.get("enabled") == "1")
    flash("签到限制设置已保存 / Check-in rule saved", "ok")
    return redirect(url_for("admin.admin_report", pin=pin))


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
                    📊 下载实际签到 Excel
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
        const url = new URL(base, window.location.origin);
        url.searchParams.set("pin", {{ pin | tojson }});
        if(startDate && endDate){
            url.searchParams.set("start_date", startDate);
            url.searchParams.set("end_date", endDate);
        }
        link.href = url.pathname + url.search;
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

<link rel="manifest" href="/admin-manifest.json?v=5">

<link rel="icon"
      type="image/png"
      sizes="192x192"
      href="/static/icon_192.png?v=5">

<link rel="apple-touch-icon"
      href="/static/icon_512.png?v=5">

<meta name="theme-color" content="#7a0000">

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
            <h2 class="section-title">👥 义工个人档案</h2>
            <div class="admin-card-desc">
                查看义工资料、照片及个人出席履历。
            </div>
            <a class="btn-tool btn-primary" href="/people/">
                进入义工个人档案
            </a>
        </div>

        <div class="card">
            <h2 class="section-title">🎓 佛学班管理</h2>
            <div class="admin-card-desc">
                学生管理、课程记录、点名及报表。
            </div>
            <a class="btn-tool btn-success" href="/class/admin">
                进入佛学班管理
            </a>
        </div>

        <div class="card">
            <h2 class="section-title">📊 年月报表系统</h2>
            <div class="admin-card-desc">
                生成义工月报、年报及白话佛法统计 Excel。
            </div>
            <a class="btn-tool btn-primary" href="/reports/">
                进入年月报表系统
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
    if (request.args.get("pin") or "").strip() != ADMIN_PIN:
        return "无权限 / Unauthorized", 403

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

# 历史记录管理：独立管理员入口，不修改义工签到与 /reading 页面。
def _history_day(raw):
    try:
        chosen = date.fromisoformat(str(raw or now_date_str()))
    except (TypeError, ValueError):
        return None
    if chosen > datetime.now(MY_TZ).date():
        return None
    return chosen.isoformat()


def _history_authorized():
    return (request.values.get("pin") or "").strip() == ADMIN_PIN


def _history_redirect(kind, day, pin):
    return redirect(url_for("admin.manage_history", kind=kind, day=day, pin=pin))


def _history_display_time(value):
    """仅格式化历史签到列表的显示时间，不修改原始记录。"""
    text = str(value or "").strip()
    for fmt in ("%H:%M", "%H:%M:%S", "%I:%M%p", "%I:%M %p"):
        try:
            clock = datetime.strptime(text.upper(), fmt)
            return f"{clock.hour % 12 or 12}:{clock.minute:02d}{'am' if clock.hour < 12 else 'pm'}"
        except ValueError:
            pass
    return text


@admin_bp.route("/admin/history")
def manage_history():
    if not _history_authorized():
        return "无权限", 403
    pin = request.args["pin"]
    kind = request.args.get("kind", "attendance")
    if kind not in ("attendance", "reading"):
        return "无效类型", 400
    day = _history_day(request.args.get("day"))
    if not day:
        return "日期无效或不能选择未来日期", 400
    rows = db_query(
        "SELECT * FROM attendance WHERE date=%s ORDER BY id DESC" if kind == "attendance"
        else "SELECT * FROM reading WHERE date=%s ORDER BY id DESC",
        (day,), fetchall=True) or []
    candidates = []
    if kind == "reading":
        candidates = db_query("""
            SELECT DISTINCT name FROM attendance
            WHERE date=%s AND signin=1 AND name IS NOT NULL AND TRIM(name)<>''
            ORDER BY name
        """, (day,), fetchall=True) or []
    recorded = {str(r.get("name") or "").strip() for r in rows}
    return render_template_string(HISTORY_PAGE, pin=pin, kind=kind, day=day,
                                  rows=rows, candidates=candidates, recorded=recorded,
                                  display_time=_history_display_time, today=now_date_str(),
                                  weekday="星期" + "一二三四五六日"[date.fromisoformat(str(day)).weekday()])


@admin_bp.route("/admin/history/reading/add", methods=["POST"])
def history_reading_add():
    if not _history_authorized():
        return "无权限", 403
    pin = request.form["pin"]
    day = _history_day(request.form.get("day"))
    if not day:
        return "日期无效", 400
    topic = request.form.get("topic", "").strip()
    time_text = request.form.get("time", "").strip()
    names = {n.strip(): "义工" for n in request.form.getlist("names") if n.strip()}
    for raw in request.form.get("extra_names", "").replace("，", ",").replace("、", ",").replace("\n", ",").split(","):
        name = raw.strip()
        if name and name not in names:
            names[name] = "佛友"
    if not topic or not names or not time_text:
        flash("请填写姓名、主题及共修时间", "bad")
        return _history_redirect("reading", day, pin)
    # 防止客户端伪造义工名单；手工添加的姓名归类为佛友。
    eligible = db_query("SELECT DISTINCT name FROM attendance WHERE date=%s AND signin=1", (day,), fetchall=True) or []
    eligible_names = {str(r.get("name") or "").strip() for r in eligible}
    for name in names:
        if name not in eligible_names:
            names[name] = "佛友"
    from db import get_conn
    with get_conn() as conn:
        with conn.cursor() as cur:
            for name, identity in names.items():
                cur.execute("SELECT pg_advisory_xact_lock(hashtext(%s),hashtext(%s))", (day, name))
                cur.execute("SELECT 1 FROM reading WHERE date=%s AND name=%s LIMIT 1", (day, name))
                if cur.fetchone():
                    continue
                cur.execute("INSERT INTO reading (date,name,identity,topic,session,time) VALUES (%s,%s,%s,%s,%s,%s)",
                            (day, name, identity, topic, "", time_text))
    flash("补点名已保存（已有记录者自动跳过）", "ok")
    return _history_redirect("reading", day, pin)


@admin_bp.route("/admin/history/<kind>/<int:record_id>/edit", methods=["GET", "POST"])
def history_edit(kind, record_id):
    if not _history_authorized():
        return "无权限", 403
    if kind not in ("attendance", "reading"):
        return "无效类型", 400
    pin = request.values["pin"]
    row = db_query(f"SELECT * FROM {kind} WHERE id=%s", (record_id,), fetchone=True)
    if not row:
        return "找不到记录", 404
    day = str(row["date"])
    if request.method == "POST":
        if kind == "reading":
            name = request.form.get("name", "").strip()
            identity = request.form.get("identity", "").strip()
            topic = request.form.get("topic", "").strip()
            time_text = request.form.get("time", "").strip()
            if not all((name, identity, topic, time_text)):
                return "姓名、身份、主题、时间不能为空", 400
            # 保持同一天同名只能一笔的原规则。
            conflict = db_query("SELECT 1 FROM reading WHERE date=%s AND name=%s AND id<>%s LIMIT 1",
                                (day, name, record_id), fetchone=True)
            if conflict:
                return "该姓名在这一天已经有白话佛法记录", 400
            db_query("UPDATE reading SET name=%s, identity=%s, topic=%s, time=%s WHERE id=%s",
                     (name, identity, topic, time_text, record_id))
        else:
            role = request.form.get("role", "").strip()
            start = request.form.get("start_time", "").strip()
            end = request.form.get("end_time", "").strip()
            remark = request.form.get("remark", "").strip()
            if role not in ("值班", "卫生", "佛台", "供台", "供花", "供果", "膳食", "佛学班"):
                return "岗位无效", 400
            try:
                hours = calc_hours(start, end) if start and end else None
            except Exception:
                return "时间格式不正确", 400
            db_query("UPDATE attendance SET role=%s,start_time=%s,end_time=%s,hours=%s,remark=%s WHERE id=%s",
                     (role, start, end, hours, remark, record_id))
        flash("记录已修改", "ok")
        return _history_redirect(kind, day, pin)
    return render_template_string(HISTORY_EDIT_PAGE, row=row, kind=kind, pin=pin, day=day)


@admin_bp.route("/admin/history/<kind>/<int:record_id>/delete", methods=["POST"])
def history_delete(kind, record_id):
    if not _history_authorized():
        return "无权限", 403
    if kind not in ("attendance", "reading"):
        return "无效类型", 400
    pin = request.form["pin"]
    if kind == "reading":
        row = db_query("SELECT date FROM reading WHERE id=%s", (record_id,), fetchone=True)
        if not row:
            return "找不到记录", 404
        day = str(row["date"])
        db_query("DELETE FROM reading WHERE id=%s", (record_id,))
        flash("记录已删除", "ok")
        return _history_redirect(kind, day, pin)

    # 签到会同时写 attendance 和 volunteer_attendance_logs。
    # 两表主键互不对应；只允许确定唯一关联时在同一事务删除。
    from psycopg2.extras import RealDictCursor
    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("SELECT * FROM attendance WHERE id=%s FOR UPDATE", (record_id,))
            row = cur.fetchone()
            if not row:
                return "找不到记录", 404
            day = str(row["date"])
            volunteer_id = str(row.get("volunteer_id") or "")
            role = str(row.get("role") or "")
            aids = _assignment_ids_from_remark(row.get("remark"))
            if aids:
                cur.execute("""
                    SELECT id, assignment_id FROM volunteer_attendance_logs
                    WHERE volunteer_id=%s AND attendance_date=%s
                      AND assignment_id=ANY(%s) FOR UPDATE
                """, (volunteer_id, day, aids))
                logs = cur.fetchall()
                # 连续值班可能对应多段日志，必须每段恰好一笔。
                if len(logs) != len(set(aids)) or {x['assignment_id'] for x in logs} != set(aids):
                    return "无法唯一对应排班签到日志，未删除；请先核对数据。", 409
            else:
                cur.execute("""
                    SELECT id FROM volunteer_attendance_logs
                    WHERE volunteer_id=%s AND attendance_date=%s
                      AND actual_role=%s AND assignment_id IS NULL
                    FOR UPDATE
                """, (volunteer_id, day, role))
                logs = cur.fetchall()
                # 同一日同岗位多次签到时没有关联键，不能猜测删除哪一笔。
                if len(logs) > 1:
                    return "同日同岗位有多笔日志，无法确定对应关系；未删除。", 409
            if logs:
                cur.execute("DELETE FROM volunteer_attendance_logs WHERE id=ANY(%s)",
                            ([x['id'] for x in logs],))
            cur.execute("DELETE FROM attendance WHERE id=%s", (record_id,))
    flash("签到记录及可确认对应的日志已同步删除", "ok")
    return _history_redirect(kind, day, pin)


HISTORY_PAGE = """
<!doctype html><html lang="zh"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>记录管理</title><link rel="stylesheet" href="{{ url_for('static',filename='css/toolbox.css') }}">
<style>.history-wrap{max-width:1050px;margin:auto}.history-table{width:100%;border-collapse:collapse}.history-table td,.history-table th{padding:12px;border-bottom:1px solid #ddd;text-align:left}.history-actions{display:flex;gap:8px;flex-wrap:wrap}.history-actions form{margin:0}.history-check{display:block;padding:9px;border-bottom:1px solid #eee}.history-check.done{color:#888;background:#f4f4f4}.history-check input{margin-right:10px}.history-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:8px}.history-scroll{overflow-x:auto}.history-input{max-width:100%;padding:10px;margin:5px 0}.history-link{display:inline-block;padding:10px;margin:4px;background:#e8efff;border-radius:8px;text-decoration:none}
.reading-history .card{border:1px solid #dce8df;border-radius:18px;padding:clamp(16px,3vw,28px)}
.reading-history .page-title{color:#185c42;font-size:clamp(25px,4vw,36px)}
.reading-history .history-link{background:#e8f3ed;color:#185c42;font-weight:700}
.reading-history .history-link.active{background:#185c42;color:white}
.reading-history .history-table th{background:#e8f3ed;color:#185c42;white-space:nowrap}
.reading-history .history-table td{border-color:#e4ebe6}
.reading-history .history-actions{align-items:center;flex-wrap:nowrap}
.reading-history .history-actions .history-link,.reading-history .history-actions button{font-size:16px;min-height:44px;margin:0;white-space:nowrap}
.reading-history .history-check{display:flex;align-items:center;gap:8px;flex-wrap:wrap;background:#f0f8f2;border:1px solid #d6e8db;border-radius:12px;padding:14px;min-height:58px;cursor:pointer}
.reading-history .history-check input{width:22px;height:22px;accent-color:#198754;margin:0}
.reading-history .history-check.done{background:#f1f2f3;color:#777;cursor:default}
.reading-history .history-check:has(input:checked){background:#d9f2e1;border-color:#198754}
.reading-history .history-status{font-size:14px;color:#66756a}
.reading-history .history-backfill h2{color:#185c42;margin-top:0}
.reading-history .history-backfill .form-input{min-height:50px;font-size:18px;width:100%;box-sizing:border-box}
.reading-history .history-backfill .btn-success{min-height:52px;font-size:20px}
@media(max-width:600px){.reading-history .history-grid{grid-template-columns:1fr}.reading-history .history-table{min-width:620px}}
</style>
</head><body><div class="page history-wrap {{ 'reading-history' if kind=='reading' else '' }}"><div class="card"><h1 class="page-title">{{ '📋 义工签到记录管理' if kind=='attendance' else '📖 白话佛法记录管理' }}</h1>
<a class="history-link" href="{{ url_for('admin.admin_report',pin=pin) }}">← 返回管理员</a>
<a class="history-link" href="{{ url_for('admin.manage_history',pin=pin,kind='attendance',day=day) }}">签到记录</a>
<a class="history-link" href="{{ url_for('admin.admin_add_record',pin=pin,date=day) }}">📝 补录签到</a>
<a class="history-link {{ 'active' if kind=='reading' else '' }}" href="{{ url_for('admin.manage_history',pin=pin,kind='reading',day=day) }}">白话佛法</a>
{% with msgs=get_flashed_messages(with_categories=true) %}{% for category,msg in msgs %}<div class="alert {{ 'alert-success' if category=='ok' else 'alert-danger' }}">{{msg}}</div>{% endfor %}{% endwith %}
<form method="get"><input type="hidden" name="pin" value="{{pin}}"><input type="hidden" name="kind" value="{{kind}}"><label for="history-day" style="font-size:20px;font-weight:700">📅 选择日期：</label><input id="history-day" class="history-input" style="font-size:24px;min-height:56px;width:260px;box-sizing:border-box;border:2px solid #2563eb;border-radius:10px;cursor:pointer" type="date" name="day" value="{{day}}" max="{{today}}" required onchange="if(this.value &amp;&amp; this.checkValidity()) this.form.requestSubmit()"></form>
<p>所选日期：<b>{{day}}（{{weekday}}）</b>　共 {{rows|length}} 笔</p>
<div class="history-scroll"><table class="history-table"><thead><tr><th>姓名</th>{% if kind=='attendance' %}<th>岗位</th><th>签到</th><th>签退</th><th>时数</th>{% else %}<th>身份</th><th>主题</th><th>时间</th>{% endif %}<th>操作</th></tr></thead><tbody>
{% for r in rows %}<tr><td>{{r.name}}</td>{% if kind=='attendance' %}<td>{{r.role}}</td><td>{{display_time(r.start_time) or '-'}}</td><td>{{display_time(r.end_time) or '未签退'}}</td><td>{{r.hours if r.hours is not none else '-'}}</td>{% else %}<td>{{r.identity}}</td><td>{{r.topic}}</td><td>{{display_time(r.time) or "-"}}</td>{% endif %}<td><div class="history-actions"><a class="history-link" href="{{url_for('admin.history_edit',kind=kind,record_id=r.id,pin=pin)}}">✏️ 修改</a><form method="post" action="{{url_for('admin.history_delete',kind=kind,record_id=r.id)}}" onsubmit="return confirm('确定删除这笔记录吗？删除后无法从此页面恢复。')"><input type="hidden" name="pin" value="{{pin}}"><button class="btn-tool btn-danger" style="width:auto;padding:10px" type="submit">🗑️ 删除</button></form></div></td></tr>{% else %}<tr><td colspan="6">这一天暂无记录。</td></tr>{% endfor %}
</tbody></table></div></div>
{% if kind=='reading' %}<div class="card history-backfill"><h2>➕ 补录白话佛法点名</h2><p>以下为 {{day}} 实际签到的义工。已点名者显示灰色，不可重复勾选。</p><form method="post" action="{{url_for('admin.history_reading_add')}}"><input type="hidden" name="pin" value="{{pin}}"><input type="hidden" name="day" value="{{day}}"><div class="history-grid">{% for c in candidates %}<label class="history-check {{'done' if c.name in recorded else ''}}"><input type="checkbox" name="names" value="{{c.name}}" {{'disabled' if c.name in recorded else ''}}><strong>{{c.name}}</strong><span class="history-status">{{'✅ 已点名' if c.name in recorded else '🔴 未点名'}}</span></label>{% else %}<p>当天没有实际签到的义工。</p>{% endfor %}</div><p><label>其他佛友姓名（可用逗号分隔多人）<br><input class="form-input" name="extra_names" placeholder="例如：张三，李四"></label></p><p><label>共修主题<br><input class="form-input" name="topic" required></label></p><p><label>共修时间<br><input class="form-input" name="time" placeholder="例如：10:30 AM" required></label></p><button class="btn-tool btn-success" type="submit">保存补点名</button></form></div>{% endif %}
</div></body></html>
"""

HISTORY_EDIT_PAGE = """
<!doctype html><html lang="zh"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>修改记录</title><link rel="stylesheet" href="{{url_for('static',filename='css/toolbox.css')}}"></head><body><div class="page"><div class="card"><h1>✏️ 修改{{'签到' if kind=='attendance' else '白话佛法'}}记录</h1><p>{{row.name}} · {{day}}</p><form method="post"><input type="hidden" name="pin" value="{{pin}}">{% if kind=='attendance' %}<label>岗位</label><select class="form-input" name="role">{% for role in ['值班','卫生','佛台','供台','供花','供果','膳食','佛学班'] %}<option value="{{role}}" {{'selected' if role==row.role else ''}}>{{role}}</option>{% endfor %}</select><label>签到时间</label><input class="form-input" name="start_time" value="{{row.start_time or ''}}"><label>签退时间</label><input class="form-input" name="end_time" value="{{row.end_time or ''}}"><label>备注</label><input class="form-input" name="remark" value="{{row.remark or ''}}">{% else %}<label>姓名</label><input class="form-input" name="name" value="{{row.name or ''}}" required><label>身份</label><select class="form-input" name="identity"><option value="义工" {{'selected' if row.identity=='义工' else ''}}>义工</option><option value="佛友" {{'selected' if row.identity=='佛友' else ''}}>佛友</option></select><label>共修主题</label><input class="form-input" name="topic" value="{{row.topic or ''}}" required><label>共修时间</label><input class="form-input" name="time" value="{{row.time or ''}}" required>{% endif %}<div class="btn-row" style="margin-top:20px"><button class="btn-tool btn-success" type="submit">保存修改</button><a class="btn-tool btn-secondary" href="{{url_for('admin.manage_history',kind=kind,day=day,pin=pin)}}">取消／返回</a></div></form></div></div></body></html>
"""
