# volunteer_routes.py

import calendar

from flask import request, redirect, url_for, render_template_string, redirect, session, jsonify
from psycopg2.extras import RealDictCursor

from db import get_conn
from schedule.blueprint import schedule_bp
from datetime import datetime, timedelta, date, timezone
from utils import apply_branch_prefix
from schedule.helpers import (
    find_volunteer_by_keyword,
    build_monthly_signup_text,
    get_daily_buddha_quote,
)

from lunar_rules import get_special_day_info
from schedule.constants import TIME_OPTIONS
from schedule.services.quote_service import get_daily_dharma
from schedule.services.settings_service import is_schedule_setting_on

from schedule.builders.schedule_builder import sync_schedule_after_signup_change
from schedule.builders.time_utils import time_to_minutes, malaysia_today, malaysia_now
from schedule.services.publish_service import is_schedule_published
from schedule.services.whatsapp_service import build_whatsapp_from_assigned
from schedule.services.shortage_service import build_signup_shortage_notice
from schedule.volunteer_templates import (
    VOLUNTEER_SIGNUP_HTML,
    VOLUNTEER_PREBOOK_HTML,
)

def find_next_meal_signup_date(base_date, days_ahead=7):
    for i in range(days_ahead + 1):
        check_date = base_date + timedelta(days=i)
        info = get_special_day_info(check_date)

        if info.get("template_type") in ["lunar_1_15", "buddhist_festival"]:
            return check_date, info

    return None, None

@schedule_bp.route("/volunteer")
def volunteer_home():

    now = malaysia_now()
    today = now.date()

    multi_day_signup_open = is_schedule_setting_on("multi_day_signup_open")
    meal_signup_open = is_schedule_setting_on("meal_signup_open")

    if today.month == 12:
        prebook_year = today.year + 1
        prebook_month = 1
    else:
        prebook_year = today.year
        prebook_month = today.month + 1

    if now.hour >= 18:
        default_date = (now + timedelta(days=1)).strftime("%Y-%m-%d")
        schedule_label = "明日值班报名情况"
    else:
        default_date = now.strftime("%Y-%m-%d")
        schedule_label = "今日值班报名情况"

    meal_button_show = False
    meal_status_text = ""
    meal_date = ""
    meal_button_date = ""
    meal_count = 0

    base_date = datetime.strptime(default_date, "%Y-%m-%d").date()

    for i in range(0, 8):   # 今天到未来7天
        check_date = base_date + timedelta(days=i)
        info = get_special_day_info(check_date)

        if info.get("template_type") in ["lunar_1_15", "buddhist_festival"]:
            meal_button_show = True
            meal_date = check_date.strftime("%Y-%m-%d")
            break

    if meal_button_show:
        with get_conn() as conn:
            with conn.cursor(cursor_factory=RealDictCursor) as cur:
                cur.execute("""
                    select count(*) as cnt
                    from volunteer_schedule_signups
                    where signup_date = %s
                    and role = '膳食'
                    and coalesce(status, 'pending') <> 'cancelled'
                """, (meal_date,))

                total_meal = cur.fetchone()["cnt"]

        meal_count = min(total_meal, 9)

        meal_button_date = check_date.strftime("%d/%m").lstrip("0").replace("/0", "/")

        meal_status_text = "已满" if meal_count >= 9 else f"{meal_count}/9"
        
    return render_template_string(
        VOLUNTEER_SIGNUP_HTML,
        default_date=default_date,
        schedule_label=schedule_label,
        times=TIME_OPTIONS,
        signup_date=default_date,
        prebook_year=prebook_year,
        prebook_month=prebook_month,
        meal_button_show=meal_button_show,
        meal_status_text=meal_status_text,
        meal_date=meal_date,
        meal_button_date=meal_button_date,
        meal_count=meal_count,
        multi_day_signup_open=multi_day_signup_open,
        meal_signup_open=meal_signup_open,
    )

@schedule_bp.route("/volunteer/signup", methods=["POST"])
def volunteer_signup():

    keyword = request.form.get("keyword", "").strip()
    signup_date = request.form.get("signup_date", "").strip()
    role = request.form.get("role", "").strip()
    start_time = request.form.get("start_time", "").strip()
    end_time = request.form.get("end_time", "").strip()
    branch = request.form.get("branch", "CHE").strip().upper()

    keyword = apply_branch_prefix(keyword, branch)
    daily_quote = get_daily_buddha_quote()

    if keyword.isdigit() and branch == "STW":
        keyword = f"STW-{keyword}"

    matches = find_volunteer_by_keyword(keyword)

    if not matches:
        return "❌ 找不到义工，请检查编号 / 姓名<br><a href='/volunteer'>返回</a>"

    if len(matches) > 1:
        return "❌ 找到多个同名义工，请用义工编号报名<br><a href='/volunteer'>返回</a>"

    vol = matches[0]
    vol_id = str(vol["id"])
    name = str(vol["name"])

    meal_role = None

    # =========================
    # 处理岗位时间
    # =========================

    if role == "值班":

        s_min = time_to_minutes(start_time)
        e_min = time_to_minutes(end_time)

        if s_min is None or e_min is None:
            return "❌ 时间格式错误，请重新选择<br><a href='/volunteer'>返回</a>"

        if e_min <= s_min:
            return "❌ 结束时间必须比开始时间迟<br><a href='/volunteer'>返回</a>"

    elif role == "卫生":

        with get_conn() as conn:
            with conn.cursor(cursor_factory=RealDictCursor) as cur:
                cur.execute("""
                    select count(*) as cnt
                    from volunteer_schedule_signups
                    where signup_date = %s
                    and role = '卫生'
                    and coalesce(status, 'pending') <> 'cancelled'
                """, (signup_date,))

                cleaning_count = int(cur.fetchone()["cnt"] or 0)

        if cleaning_count >= 3:
            return """
            <h2>❌ 卫生岗位已满</h2>
            <p>感恩您的发心，今天卫生已有 3 位义工报名。</p>
            <p>请改报其他日期或联系负责人。</p>
            <a href="/volunteer">返回义工报名</a>
            """

        date_obj = datetime.strptime(signup_date, "%Y-%m-%d").date()
        special_info = get_special_day_info(date_obj)

        if special_info.get("template_type") == "buddhist_festival":
            start_time = "6:00am"
            end_time = "8:00am"
        else:
            start_time = "8:00am"
            end_time = "10:00am"

    elif role == "供台":

        date_obj = datetime.strptime(signup_date, "%Y-%m-%d").date()
        special_info = get_special_day_info(date_obj)

        if special_info.get("template_type") not in ["lunar_1_15", "buddhist_festival"]:
            return """
            <h1>❌ 这一天不需要供台报名</h1>
            <p>供台通常只开放在初一、十五或佛诞大日子。</p>
            <p>请检查日期是否选错。</p>
            <a href="/volunteer">返回重新报名</a>
            """

        start_time = "6:00am"
        end_time = "8:00am"

    elif role == "膳食":

        selected_date = datetime.strptime(signup_date, "%Y-%m-%d").date()
        meal_date, special_info = find_next_meal_signup_date(selected_date, days_ahead=7)

        if not meal_date:
            return """
            <h1>❌ 近期没有开放膳食组报名</h1>
            <p>系统在接下来 7 天内找不到初一、十五或佛诞大日子。</p>
            <p>请检查日期是否选错。</p>
            <a href="/volunteer">返回重新报名</a>
            """

        signup_date = meal_date.strftime("%Y-%m-%d")
        start_time = "8:00am"
        end_time = "2:00pm"

    else:
        return "❌ 岗位错误，请重新选择<br><a href='/volunteer'>返回</a>"

    # =========================
    # 新增 / 更新报名
    # =========================

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:

            cur.execute("""
                select id
                from volunteer_schedule_signups
                where volunteer_id = %s
                and signup_date = %s
                and role = %s
                and coalesce(status, 'pending') <> 'cancelled'
                limit 1
            """, (vol_id, signup_date, role))

            existing = cur.fetchone()

            if role == "膳食":
                cur.execute("""
                    select count(*) as cnt
                    from volunteer_schedule_signups
                    where signup_date = %s
                    and role = '膳食'
                    and coalesce(status, 'pending') <> 'cancelled'
                """, (signup_date,))

                meal_count = int(cur.fetchone()["cnt"] or 0)

                if existing:
                    meal_role = None
                elif meal_count < 9:
                    meal_role = "派餐义工"
                else:
                    meal_role = "候补义工"

            if existing:

                signup_id = existing["id"]

                cur.execute("""
                    update volunteer_schedule_signups
                    set start_time = %s,
                        end_time = %s,
                        status = 'pending',
                        assigned_place = null,
                        remarks = '义工网页更新报名',
                        meal_role = %s
                    where id = %s
                """, (
                    start_time,
                    end_time,
                    meal_role,
                    signup_id
                ))

                result_title = "✅ 已更新报名"

            else:

                cur.execute("""
                    insert into volunteer_schedule_signups
                    (
                        volunteer_id,
                        name,
                        signup_date,
                        role,
                        start_time,
                        end_time,
                        status,
                        remarks,
                        meal_role
                    )
                    values
                    (
                        %s, %s, %s, %s, %s, %s,
                        'pending',
                        '义工报名',
                        %s
                    )
                    returning id
                """, (
                    vol_id,
                    name,
                    signup_date,
                    role,
                    start_time,
                    end_time,
                    meal_role
                ))

                signup_id = cur.fetchone()["id"]

                result_title = "🎉 报名成功"

            conn.commit()

    # =========================
    # V2 同步：只走统一函数
    # =========================

    sync_schedule_after_signup_change(
        signup_id,
        action="upsert",
        changed_by="volunteer"
    )

    # =========================
    # 判断发布状态与安排结果
    # =========================

    published = is_schedule_published(signup_date)
    
    assignment_rows = []

    if published:

        with get_conn() as conn:
            with conn.cursor(cursor_factory=RealDictCursor) as cur:

                cur.execute("""
                    select
                        assigned_place,
                        shift_label,
                        start_time,
                        end_time
                    from volunteer_schedule_assignments
                    where signup_id = %s
                    order by start_time, id
                """, (signup_id,))

                assignment_rows = cur.fetchall()
    if published and assignment_rows:

        assignment_lines = ""

        for a in assignment_rows:
            shift = a.get("shift_label") or ""
            place = a.get("assigned_place") or ""
            s = a.get("start_time") or ""
            e = a.get("end_time") or ""

            assignment_lines += f"""
            <div class="assignment-line">
                ✅ {shift} {place}<br>
                <span>{s} ~ {e}</span>
            </div>
            """

        status_html = f"""
        <div class="alert alert-success">
            🟢 状态：已加入正式值班表<br><br>
            {assignment_lines}
            <br>
            请依照以上岗位值班。若负责人之后有更新，请以最新正式值班表为准。
        </div>
        """

    elif published and not assignment_rows:

        status_html = """
        <div class="alert alert-warning">
            🟡 状态：报名成功，等待系统安排<br><br>
            系统暂时还没有找到适合岗位。<br>
            负责人会查看后再处理，请以最新正式值班表为准。
        </div>
        """

    else:

        status_html = """
        <div class="alert alert-warning">
            🟡 状态：等待负责人安排<br><br>
            ⚠️ 当前属于报名阶段。<br>
            最终岗位安排请以负责人公布的正式值班表为准。<br>
            请多留意义工群信息，感恩大家护持观音堂 🙏
        </div>
        """

    # =========================
    # 膳食报名成功页
    # =========================

    if role == "膳食":

        with get_conn() as conn:
            with conn.cursor(cursor_factory=RealDictCursor) as cur:
                cur.execute("""
                    select id, name
                    from volunteer_schedule_signups
                    where signup_date = %s
                    and role = '膳食'
                    and coalesce(status, 'pending') <> 'cancelled'
                    order by created_at, id
                """, (signup_date,))

                all_rows = cur.fetchall()

        my_no = None

        for i, row in enumerate(all_rows, start=1):
            if row["id"] == signup_id:
                my_no = i
                break

        if my_no and my_no <= 9:
            my_meal_role = "派餐义工"
        else:
            my_meal_role = "候补义工"

        meal_rows = all_rows[:9]
        backup_rows = all_rows[9:]

        lines = ""

        for i in range(1, 10):
            person = meal_rows[i - 1]["name"] if i <= len(meal_rows) else ""
            lines += f"<div class='name-line'>{i}）{person}</div>"

        backup_lines = ""

        for i, row in enumerate(backup_rows, start=10):
            backup_lines += f"<div class='name-line'>{i}）{row['name']}</div>"

        return render_template_string("""
        <!doctype html>
        <html>
        <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">

        <style>
        body {
            font-family:"Microsoft YaHei", Arial;
            background:#f7f2e8;
            padding:20px;
            font-size:22px;
        }

        .card {
            max-width:760px;
            margin:auto;
            background:white;
            border-radius:20px;
            padding:28px;
            box-shadow:0 4px 18px rgba(0,0,0,0.12);
        }

        h1 {
            font-size:34px;
            margin-bottom:10px;
        }

        .big-green {
            font-size:30px;
            color:#168a3a;
            font-weight:bold;
            line-height:1.5;
        }

        .info {
            background:#fff7dc;
            border-radius:16px;
            padding:18px;
            margin:18px 0;
            line-height:1.7;
        }

        .section {
            margin-top:24px;
            padding-top:18px;
            border-top:2px solid #ddd;
        }

        .name-line {
            font-size:24px;
            padding:8px 0;
        }

        .btn {
            display:block;
            margin-top:18px;
            background:#1f9d55;
            color:white;
            text-align:center;
            padding:14px;
            border-radius:14px;
            text-decoration:none;
            font-size:22px;
            font-weight:bold;
        }
        </style>
        </head>

        <body>

        <div class="card">

            <h1>🍱 素食结缘报名成功</h1>

            <div class="info">
                <div><b>日期：</b>{{ signup_date }}</div>
                <div><b>开放堂食：</b>11:00am</div>
            </div>

            <div class="big-green">
                您是第 {{ my_no }} 位膳食组义工<br>
                您的岗位：{{ my_meal_role }}
            </div>

            <div class="section">
                <h2>🍱 派餐义工（最多9位）</h2>
                {{ lines|safe }}
            </div>

            <div class="section">
                <h2>👥 候补义工</h2>
                {{ backup_lines|safe if backup_lines else "<p>暂时还没有候补义工</p>"|safe }}
            </div>

            <div class="section">
                <p>✅ 派餐义工最迟 <b>10:45am</b> 开始站岗</p>
                <p>✅ 所有报名义工需要在 <b>9:30am</b> 前报到</p>
                <p>✅ 活动结束后协助清理场地及餐具</p>
            </div>

            <a class="btn" href="/volunteer">返回义工首页</a>

        </div>

        </body>
        </html>
        """,
        signup_date=signup_date,
        my_no=my_no,
        my_meal_role=my_meal_role,
        lines=lines,
        backup_lines=backup_lines)

    # =========================
    # 普通报名成功页
    # =========================

    return render_template_string("""
    <!doctype html>
    <html>
    <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet"
          href="{{ url_for('static', filename='css/toolbox.css') }}">

    <style>
    body {
        font-family:"Microsoft YaHei";
        background:#f5f5f5;
        padding:20px;
    }

    .box {
        max-width:760px;
        margin:30px auto;
        background:white;
        border-radius:20px;
        padding:30px;
        text-align:center;
    }

    .success {
        background:#e8f5e9;
        border:2px solid #4CAF50;
        border-radius:15px;
        padding:25px;
        margin-bottom:25px;
    }

    .success h1 {
        color:#2e7d32;
        margin-top:0;
        font-size:38px;
    }

    .info {
        font-size:25px;
        line-height:1.9;
    }

    .quote-box {
        margin-top:20px;
        padding:18px;
        background:#fff5f5;
        border:2px solid #efc8c8;
        border-radius:18px;
        text-align:center;
    }

    .quote-title {
        font-size:21px;
        font-weight:bold;
        color:#9b6b00;
    }

    .quote-content {
        font-size:23px;
        line-height:1.8;
        color:#b84a4a;
    }

    .notice {
        border-radius:16px;
        padding:22px;
        margin:22px 0;
        font-size:23px;
        line-height:1.8;
    }

    .yellow-notice {
        background:#fff8e1;
        color:#8d6e00;
    }

    .green-notice {
        background:#e8f5e9;
        color:#1b6e34;
        border:2px solid #7acb8a;
    }

    .assignment-line {
        background:white;
        border-radius:12px;
        padding:12px;
        margin:10px 0;
        font-weight:bold;
    }

    .assignment-line span {
        font-weight:normal;
        color:#555;
    }
    
    @media (max-width:700px) {
        body {
            padding:12px;
        }

        .box {
            padding:20px;
            margin:15px auto;
        }

        .success h1 {
            font-size:32px;
        }

        .info,
        .notice {
            font-size:21px;
        }

        .btn {
            font-size:20px;
            padding:13px;
        }
    }
    </style>
    </head>

    <body>

    <div class="box">

        <div class="success">

            <h1>{{ result_title }}</h1>

            <div class="info">

                🙏 感恩发心护持观音堂<br><br>

                义工：{{ name }}<br>
                日期：{{ signup_date }}<br>
                岗位：{{ role }}

                {% if role == "值班" %}
                <br>
                时间：{{ start_time }} ~ {{ end_time }}
                {% endif %}

            </div>

        </div>

        <div class="quote-card">

            <div class="quote-title">
                🌸 每日佛言佛语
            </div>

            <div class="quote-content">
                {{ daily_quote }}
            </div>

        </div>

        {{ status_html|safe }}

        <div class="btn-row">

            <a class="btn-tool btn-blue btn-full"
            href="/volunteer/my_schedule_search">
                📋 查看我的报名
            </a>

            <a class="btn-tool btn-orange btn-full"
            href="/volunteer/day_schedule?date={{ signup_date }}">
                📅 查看当天值班表
            </a>

            <a class="btn-tool btn-green btn-full"
            href="/volunteer">
                ➕ 继续报名
            </a>

        </div>

    </div>

    </body>
    </html>
    """,
    result_title=result_title,
    name=name,
    signup_date=signup_date,
    role=role,
    start_time=start_time,
    end_time=end_time,
    daily_quote=daily_quote,
    status_html=status_html)


@schedule_bp.route("/volunteer/query_volunteer")
def query_volunteer_api():

    keyword = request.args.get("keyword", "").strip()
    branch = request.args.get("branch", "CHE").strip()

    if not keyword:
        return jsonify({
            "ok": False,
            "message": "请输入编号 / 姓名 / 电话"
        })
    
    matches = find_volunteer_by_keyword(keyword)

    if not matches:
        return jsonify({
            "ok": False,
            "message": "找不到义工"
        })

    if len(matches) > 1:
        return jsonify({
            "ok": False,
            "message": "找到多个义工，请输入完整编号"
        })

    v = matches[0]

    return jsonify({
        "ok": True,
        "volunteer_id": v.get("id") or v.get("volunteer_id"),
        "name": v.get("name")
    })


@schedule_bp.route("/volunteer/meal_status")
def volunteer_meal_status():

    signup_date = request.args.get("date") or malaysia_today().strftime("%Y-%m-%d")

    date_obj = datetime.strptime(signup_date, "%Y-%m-%d").date()
    special_info = get_special_day_info(date_obj)

    lunar_text = special_info.get("lunar_text") or special_info.get("lunar_date") or "农历初一 / 十五"
    festival_name = special_info.get("festival_name") or special_info.get("name") or ""

    page_title = f"🍱 {lunar_text} · 素食结缘"
    festival_line = f"🙏 {festival_name}" if festival_name else ""

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
            select id,name
            from volunteer_schedule_signups
            where signup_date=%s
            and role='膳食'
            and coalesce(status,'pending')<>'cancelled'
            order by created_at,id
            """,(signup_date,))

            rows = cur.fetchall()
            print("MEAL STATUS ROWS:", rows)
            

    meal_names = [r["name"] for r in rows[:9]]
    backup_names = [r["name"] for r in rows[9:]]

    lines = ""
    for i in range(1, 10):
        person = meal_names[i - 1] if i <= len(meal_names) else ""
        lines += f"<div class='name-line'>{i}）{person}</div>"

    backup_lines = ""
    for i, person in enumerate(backup_names, start=10):
        backup_lines += f"<div class='name-line'>{i}）{person}</div>"

    meal_count = len(meal_names)
    backup_count = len(backup_names)

    return render_template_string("""
    <style>
    body {
        font-family: Arial, "Microsoft YaHei", sans-serif;
        background:#f7f2e8;
        padding:20px;
        font-size:24px;
    }

    .card {
        max-width:820px;
        margin:auto;
        background:white;
        border-radius:22px;
        padding:30px;
        box-shadow:0 4px 18px rgba(0,0,0,0.12);
    }

    h1 {
        font-size:36px;
        color:#8b5a2b;
        margin-top:0;
        line-height:1.4;
    }

    .info-box {
        background:#fff8e6;
        border:2px solid #f3d9a5;
        border-radius:18px;
        padding:18px;
        margin:20px 0;
        font-size:24px;
        line-height:1.8;
        color:#5c3b1e;
    }

    .status-box {
        background:#eaf8ef;
        border:2px solid #b8e0c4;
        border-radius:18px;
        padding:18px;
        margin:20px 0;
        font-size:28px;
        font-weight:bold;
        color:#167a35;
        text-align:center;
    }

    .section {
        margin-top:24px;
        padding-top:18px;
        border-top:2px solid #eee;
    }

    h2 {
        font-size:30px;
        color:#5c3b1e;
    }

    .name-line {
        font-size:28px;
        padding:10px 0;
        border-bottom:1px dashed #eee;
    }

    .empty {
        color:#999;
        font-size:24px;
        padding:10px 0;
    }

    .note {
        background:#fff3f0;
        border:2px solid #f5c2bd;
        border-radius:16px;
        padding:16px;
        margin-top:20px;
        font-size:22px;
        line-height:1.7;
        color:#7a2e24;
    }

    .btn {
        display:block;
        margin-top:24px;
        background:#1f9d55;
        color:white;
        text-align:center;
        padding:18px;
        border-radius:14px;
        text-decoration:none;
        font-size:26px;
        font-weight:bold;
    }

    @media (max-width:700px) {
        body {
            padding:10px;
        }

        .card {
            padding:22px;
        }

        h1 {
            font-size:30px;
        }

        .name-line {
            font-size:25px;
        }
    }
    </style>

    <div class="card">

        <h1>{{ page_title }}</h1>

        <div class="info-box">
            <div><b>📅 日期：</b>{{ signup_date }}</div>
            <div><b>🌙 农历：</b>{{ lunar_text }}</div>

            {% if festival_line %}
            <div><b>🙏 佛诞：</b>{{ festival_line }}</div>
            {% endif %}

            <div><b>🍱 开放堂食：</b>11:00am</div>
        </div>

        <div class="status-box">
            派餐义工：{{ meal_count }} / 9 位
        </div>

        <div class="section">
            <h2>🍱 派餐义工（最多9位）</h2>
            {{ lines|safe }}
        </div>

        <div class="section">
            <h2>👥 候补义工</h2>

            {% if backup_lines %}
                {{ backup_lines|safe }}
            {% else %}
                <div class="empty">暂时还没有候补义工</div>
            {% endif %}
        </div>

        <div class="note">
            ⚠️ 派餐义工最多 9 位。<br>
            第 10 位开始会列为候补义工，如现场需要，膳食组负责人会再安排。<br><br>
            ✅ 派餐义工最迟 10:45am 开始站岗。<br>
            ✅ 所有报名义工需要在 9:30am 前报到。<br>
            ✅ 活动结束后协助清理场地及餐具。
        </div>

        <a class="btn" href="/volunteer">返回义工首页</a>

    </div>
    """,
    page_title=page_title,
    signup_date=signup_date,
    lunar_text=lunar_text,
    festival_line=festival_line,
    meal_count=meal_count,
    backup_count=backup_count,
    lines=lines,
    backup_lines=backup_lines)


@schedule_bp.route("/schedule/meal_leader/add", methods=["POST"])
def add_meal_leader():

    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    signup_date = request.form.get("date")
    keyword = request.form.get("keyword", "").strip()

    matches = find_volunteer_by_keyword(keyword)

    if not matches:
        return "❌ 找不到义工<br><a href='/schedule'>返回</a>"

    if len(matches) > 1:
        return "❌ 找到多个义工，请用义工编号报名<br><a href='/schedule'>返回</a>"

    vol = matches[0]
    vol_id = str(vol["id"])
    name = vol["name"]

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:

            cur.execute("""
                insert into volunteer_schedule_signups
                (volunteer_id, name, signup_date, role, start_time, end_time, status, remarks, meal_role)
                values (%s, %s, %s, '膳食', '8:00am', '2:00pm', 'pending', '负责人加入膳食组长', '组长')
            """, (
                vol_id,
                name,
                signup_date
            ))

        conn.commit()

    return redirect(f"/schedule/notice_center?date={signup_date}")


@schedule_bp.route("/volunteer/my_schedule")
def volunteer_my_schedule():
    keyword = request.args.get("keyword", "").strip()

    if not keyword:
        return "❌ 请输入义工编号 / 姓名 / 电话<br><a href='/volunteer'>返回</a>"

    matches = find_volunteer_by_keyword(keyword)

    if not matches:
        return "❌ 找不到义工<br><a href='/volunteer'>返回</a>"

    if len(matches) > 1:
        return "❌ 找到多个同名义工，请用义工编号查询<br><a href='/volunteer'>返回</a>"

    vol = matches[0]
    vol_id = str(vol["id"])
    name = str(vol["name"])

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                select *
                from volunteer_schedule_signups
                where volunteer_id = %s
                and coalesce(status, 'pending') <> 'cancelled'
                order by signup_date, start_time
            """, (vol_id,))
            rows = cur.fetchall()

    html = f"""
    <!doctype html>
    <html>
    <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <style>
    body {{
        font-family:"Microsoft YaHei";
        background:#f5f5f5;
        padding:18px;
        font-size:26px;
    }}

    .box {{
        max-width:800px;
        margin:auto;
    }}

    .card {{
        background:white;
        border-radius:18px;
        padding:24px;
        margin:18px 0;
        box-shadow:0 3px 12px rgba(0,0,0,0.08);
        font-size:28px;
        line-height:1.8;
    }}

    .status {{
        background:#fff3cd;
        color:#8a5a00;
        padding:10px 14px;
        border-radius:10px;
        display:inline-block;
        font-weight:bold;
    }}

    .big-btn {{
        width:100%;
        border:0;
        border-radius:15px;
        padding:22px;
        margin-top:14px;
        font-size:30px;
        font-weight:bold;
        color:white;
    }}

    .edit-btn {{
        background:#4CAF50;
    }}

    .cancel-btn {{
        background:#d9534f;
    }}

    .back-btn {{
        display:block;
        background:#607D8B;
        color:white;
        text-align:center;
        text-decoration:none;
        padding:18px;
        border-radius:15px;
        font-size:28px;
        font-weight:bold;
        margin:18px 0;
    }}
    </style>
    </head>
    <body>
    <div class="box">

    <h1>📋 我的报名记录</h1>
    <h2>{name}</h2>

    <a class="back-btn" href="/volunteer">⬅ 返回报名</a>
    """

    if not rows:
        html += """
        <div class="card">
            暂时没有报名记录。
        </div>
        """
    else:
        for r in rows:
            assigned_place = r.get("assigned_place") or "尚未安排"
            status = str(r.get("status") or "pending")

            if status == "pending":
                status_text = "🟡 等待负责人安排"
            elif status == "assigned":
                status_text = "🟢 已安排"
            else:
                status_text = status

            html += f"""
            <div class="card">
                📅 <b>{r["signup_date"]}</b><br>
                岗位：<b>{r["role"]}</b><br>
                时间：{r["start_time"]} ~ {r["end_time"]}<br>
                系统安排：{assigned_place}<br>
                状态：<span class="status">{status_text}</span>
            """

            signup_date_obj = r["signup_date"]

            if str(signup_date_obj) != str(malaysia_today()):
                html += f"""
                <form method="get" action="/volunteer/edit_signup/{r['id']}">
                    <button class="big-btn edit-btn" type="submit">
                        ✏️ 修改这次报名
                    </button>
                </form>

                <form method="post"
                    action="/volunteer/cancel/{r['id']}"
                    onsubmit="return confirm('确定不参加这次报名了吗？');">
                    <button class="big-btn cancel-btn" type="submit">
                        ❌ 不参加了，取消报名
                    </button>
                </form>
                """
            else:
                html += """
                <p style="color:#b36b00; font-size:26px;">
                    🚫 今天值班如需更改或取消，请务必通知负责人。
                </p>
                """

            html += """
            </div>
            """

    html += """
    </div>
    </body>
    </html>
    """

    return html


@schedule_bp.route("/volunteer/guide")
def volunteer_guide():
    return render_template_string("""
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>义工须知</title>

<link rel="stylesheet"
      href="{{ url_for('static', filename='css/toolbox.css') }}">
<style>
.guide-wrap{
    max-width:900px;
    margin:auto;
    padding:24px;
}

.guide-content{
    font-size:26px;
    line-height:1.8;
}

.guide-content h2{
    margin-top:30px;
    font-size:32px;
}
</style>
</head>

<body>

<div class="guide-wrap">

    <div class="card">

        <h1 class="page-title">📖 义工须知</h1>

        <p class="page-subtitle">
            感恩您发心护持观音堂，请先了解报名与值班规则。
        </p>

        <div class="guide-content">

            <h2>🙏 感恩发心护持道场</h2>
            <p>
                感恩您发心参与观音堂义工服务。每一位义工都是护持道场的重要力量，
                愿大家互相配合，共同成就道场。
            </p>

            <h2>📝 关于报名</h2>
            <p>
                报名后，系统会根据当天需要自动安排岗位。负责人也可能因现场情况调整岗位或时间。
                报名成功并不代表最终岗位。
            </p>

            <h2>📢 正式值班表</h2>
            <p>
                每天约 <b>10:00pm</b> 发布正式值班表。请以最新正式值班表为准。
                发布后如有调整，请以负责人通知为准。
            </p>

            <h2>✏️ 修改或取消报名</h2>
            <p>
                正式发布前，可自行进入系统修改或取消报名。正式发布后，请联络负责人处理。
            </p>

            <h2>🚫 值班当天无法出席</h2>
            <p>
                若当天因突发情况无法值班，请尽快通知负责人，以便安排其他义工补位。
                请不要直接缺席，以免影响当天道场运作。
            </p>

            <h2>⏰ 请准时签到</h2>
            <p>
                到达观音堂后，请先签到。签到后请依照系统显示的岗位值班。
                如负责人现场有调整，请以负责人安排为准。
            </p>

            <h2>🤝 值班期间</h2>
            <p>
                请礼貌待人，配合负责人安排。如有任何疑问，可随时向负责人请教。
                请穿着整齐仪容：有义工服者请穿义工服；尚未领取义工服者，请穿红色上衣及黑色长裤。
            </p>

            <h2>❤️ 感恩大家</h2>
            <p>
                愿大家发心护持道场，广结善缘，福慧增长。
                🙏 感恩每一位义工。
            </p>

        </div>

        <div class="btn-row">
            <a class="btn-tool btn-gray btn-full" href="/volunteer">
                ⬅ 返回义工首页
            </a>
        </div>

    </div>

</div>

</body>
</html>
""")


@schedule_bp.route("/volunteer/edit_signup/<int:signup_id>", methods=["GET"])
def volunteer_edit_signup(signup_id):

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                select *
                from volunteer_schedule_signups
                where id = %s
                and coalesce(status, 'pending') <> 'cancelled'
            """, (signup_id,))
            row = cur.fetchone()

    if not row:
        return "❌ 找不到这笔报名，或已取消<br><a href='/volunteer'>返回</a>"

    return render_template_string("""
    <!doctype html>
    <html>
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>修改报名</title>
        <link rel="stylesheet"
              href="{{ url_for('static', filename='css/toolbox.css') }}">
    </head>

    <body>

    <div class="page-wrap">

        <div class="card">

            <div class="result-icon" style="text-align:center;">✏️</div>
            <div class="result-title" style="text-align:center;">修改报名</div>

            <div class="name" style="text-align:center;font-size:32px;font-weight:bold;margin-bottom:20px;">
                {{ row.name }}
            </div>

            <form method="post"
                action="/volunteer/edit_signup/{{ row.id }}"
                onsubmit="return confirm('确定修改这次报名？');">

                <div class="form-grid">

                    <div class="form-group">
                        <label class="form-label">📅 日期</label>

                        <input
                            class="form-input"
                            type="date"
                            name="signup_date"
                            value="{{ row.signup_date }}"
                            required>
                    </div>

                    <div class="form-group">
                        <label class="form-label">👷 岗位</label>

                        <select
                            class="form-select"
                            name="role">

                            <option value="值班"
                                {% if row.role=="值班" %}selected{% endif %}>
                                值班
                            </option>

                            <option value="卫生"
                                {% if row.role=="卫生" %}selected{% endif %}>
                                卫生
                            </option>

                            <option value="供台"
                                {% if row.role=="供台" %}selected{% endif %}>
                                供台
                            </option>

                        </select>
                    </div>

                    <div class="form-group">
                        <label class="form-label">🕙 开始时间</label>

                        <input
                            class="form-input"
                            name="start_time"
                            value="{{ row.start_time }}">
                    </div>

                    <div class="form-group">
                        <label class="form-label">🕒 结束时间</label>

                        <input
                            class="form-input"
                            name="end_time"
                            value="{{ row.end_time }}">
                    </div>

                </div>

                <div class="btn-row">

                    <button
                        class="btn-tool btn-green btn-full"
                        type="submit">

                        ✅ 确认修改报名

                    </button>

                    <a
                        class="btn-tool btn-gray btn-full"
                        href="/volunteer/my_schedule?keyword={{ row.volunteer_id }}">

                        ⬅ 返回我的报名

                    </a>

                </div>

            </form>

        </div>

    </div>

    </body>
    </html>
    """, row=row)


@schedule_bp.route("/volunteer/edit_signup/<int:signup_id>", methods=["POST"])
def volunteer_edit_signup_post(signup_id):

    new_date = request.form.get("signup_date", "").strip()
    new_role = request.form.get("role", "").strip()
    new_start = request.form.get("start_time", "").strip()
    new_end = request.form.get("end_time", "").strip()

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:

            cur.execute("""
                select *
                from volunteer_schedule_signups
                where id = %s
                and coalesce(status, 'pending') <> 'cancelled'
            """, (signup_id,))
            old = cur.fetchone()

            if not old:

                return render_template_string("""
                <!DOCTYPE html>
                <html>
                <head>
                    <meta charset="utf-8">
                    <title>找不到报名记录</title>
                    <link rel="stylesheet"
                          href="{{ url_for('static', filename='css/toolbox.css') }}">
                </head>
                <body>

                <div class="page-wrap">

                    <div class="card result-card">

                        <div class="result-icon">❌</div>

                        <div class="result-title">
                            找不到报名记录
                        </div>

                        <p>
                            这笔报名不存在，或已经取消，无法修改。
                        </p>

                        <div class="btn-row">
                            <a class="btn-tool btn-gray" href="/volunteer">
                                返回首页
                            </a>
                        </div>

                    </div>

                </div>

                </body>
                </html>
                """)
            
            today = malaysia_today()

            if str(old["signup_date"]) == str(today):

                return render_template_string("""
                <!DOCTYPE html>
                <html>
                <head>
                    <meta charset="utf-8">
                    <title>今天不能修改报名</title>
                    <link rel="stylesheet"
                          href="{{ url_for('static', filename='css/toolbox.css') }}">
                </head>
                <body>

                <div class="page-wrap">

                    <div class="card result-card">

                        <div class="result-icon">🚫</div>

                        <div class="result-title">
                            今天不能修改报名
                        </div>

                        <div class="info-box">

                            <div class="info-row">
                                <span class="info-label">姓名</span>
                                <span class="info-value">{{ old.name }}</span>
                            </div>

                            <div class="info-row">
                                <span class="info-label">日期</span>
                                <span class="info-value">{{ old.signup_date }}</span>
                            </div>

                            <div class="info-row">
                                <span class="info-label">岗位</span>
                                <span class="info-value">{{ old.role }}</span>
                            </div>

                        </div>

                        <p class="text-red" style="font-weight:bold;">
                            今天已经是值班当天，如需修改，请通知负责人处理。
                        </p>

                        <div class="btn-row">
                            <a class="btn-tool btn-gray"
                            href="/volunteer/my_schedule?keyword={{ old.volunteer_id }}">
                                返回我的报名
                            </a>
                        </div>

                    </div>

                </div>

                </body>
                </html>
                """, old=old)

            cur.execute("""
                update volunteer_schedule_signups
                set status = 'cancelled'
                where id = %s
            """, (signup_id,))

            cur.execute("""
                insert into volunteer_schedule_signups
                    (volunteer_id, name, signup_date, role, start_time, end_time, status, remarks)
                values
                    (%s, %s, %s, %s, %s, %s, 'pending', %s)
                returning id
            """, (
                old["volunteer_id"],
                old["name"],
                new_date,
                new_role,
                new_start,
                new_end,
                f"由报名 #{signup_id} 修改而来"
            ))

            new_signup = cur.fetchone()
            new_signup_id = new_signup["id"]

        conn.commit()

    sync_schedule_after_signup_change(
        signup_id,
        action="cancel",
        changed_by="volunteer"
    )

    sync_schedule_after_signup_change(
        new_signup_id,
        action="upsert",
        changed_by="volunteer"
    )
    
    return redirect(f"/volunteer/my_schedule?keyword={old['volunteer_id']}")

@schedule_bp.route(
    "/volunteer/cancel/<int:signup_id>",
    methods=["POST"]
)
def volunteer_cancel_signup(signup_id):

    today = malaysia_today()

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:

            cur.execute("""
                select
                    id,
                    volunteer_id,
                    name,
                    signup_date,
                    role,
                    coalesce(status, 'pending') as status
                from volunteer_schedule_signups
                where id = %s
            """, (signup_id,))

            row = cur.fetchone()

            if not row:

                return render_template_string("""
                <!DOCTYPE html>
                <html>
                <head>
                    <meta charset="utf-8">
                    <title>找不到报名记录</title>
                    <link rel="stylesheet"
                          href="{{ url_for('static', filename='css/toolbox.css') }}">
                </head>
                <body>

                <div class="page-wrap">

                    <div class="card result-card">

                        <div class="result-icon">❌</div>

                        <div class="result-title">
                            找不到报名记录
                        </div>

                        <p>
                            这笔报名记录不存在，可能已经删除或取消。
                        </p>

                        <div class="btn-row">

                            <a class="btn-tool btn-gray"
                            href="/volunteer">
                                返回首页
                            </a>

                        </div>

                    </div>

                </div>

                </body>
                </html>
                """)

            signup_date = row["signup_date"]

            if str(signup_date) == str(today):

                return render_template_string("""
                <!DOCTYPE html>
                <html>
                <head>
                    <meta charset="utf-8">
                    <title>今天不能取消报名</title>
                    <link rel="stylesheet"
                          href="{{ url_for('static', filename='css/toolbox.css') }}">
                </head>
                <body>

                <div class="page-wrap">

                    <div class="card result-card">

                        <div class="result-icon">🚫</div>

                        <div class="result-title">
                            今天不能取消报名
                        </div>

                        <div class="info-box">

                            <div class="info-row">
                                <span class="info-label">姓名</span>
                                <span class="info-value">{{ row.name }}</span>
                            </div>

                            <div class="info-row">
                                <span class="info-label">日期</span>
                                <span class="info-value">{{ row.signup_date }}</span>
                            </div>

                            <div class="info-row">
                                <span class="info-label">岗位</span>
                                <span class="info-value">{{ row.role }}</span>
                            </div>

                        </div>

                        <p style="color:#dc2626;font-weight:bold;">
                            今天已经是值班当天，如需取消，请通知负责人处理。
                        </p>

                        <div class="btn-row" style="justify-content:center;">

                            <a class="btn-tool btn-gray"
                            href="/volunteer/my_schedule?keyword={{ row.volunteer_id }}">
                                返回我的报名
                            </a>

                        </div>

                    </div>

                </div>

                </body>
                </html>
                """, row=row)

            cur.execute("""
                update volunteer_schedule_signups
                set
                    status = 'cancelled',
                    assigned_place = null,
                    remarks = '义工自行取消报名'
                where id = %s
            """, (signup_id,))

            conn.commit()

    sync_schedule_after_signup_change(
        signup_id,
        action="cancel",
        changed_by="volunteer"
    )

    return render_template_string("""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <title>已取消报名</title>
        <link rel="stylesheet"
              href="{{ url_for('static', filename='css/toolbox.css') }}">
    </head>
    <body>

    <div class="page-wrap">

        <div class="card result-card">

            <div class="result-icon">✅</div>
            <div class="result-title">已取消报名</div>

            <div class="info-box">

                <div class="info-row">
                    <span class="info-label">姓名</span>
                    <span class="info-value">{{ row.name }}</span>
                </div>

                <div class="info-row">
                    <span class="info-label">日期</span>
                    <span class="info-value">{{ row.signup_date }}</span>
                </div>

                <div class="info-row">
                    <span class="info-label">岗位</span>
                    <span class="info-value">{{ row.role }}</span>
                </div>

                <div class="info-row">
                    <span class="info-label">状态</span>
                    <span class="info-value">已取消</span>
                </div>

            </div>

            <p>你的报名已经取消，系统会自动同步正式值班表。</p>

            <div class="btn-row" style="justify-content:center;">
                <a class="btn-tool btn-blue"
                href="/volunteer/my_schedule?keyword={{ row.volunteer_id }}">
                    返回我的报名
                </a>
            </div>

        </div>

    </div>

    </body>
    </html>
    """, row=row)

@schedule_bp.route("/volunteer/today_schedule")
def volunteer_today_schedule():

    now = malaysia_now()

    if now.hour >= 18:
        target_date = (now + timedelta(days=1)).strftime("%Y-%m-%d")
    else:
        target_date = now.strftime("%Y-%m-%d")

    return redirect(url_for(
        "schedule.volunteer_day_schedule",
        date=target_date
    ))

@schedule_bp.route("/volunteer/day_schedule")
def volunteer_day_schedule():
    signup_date = request.args.get("date", "").strip()

    if not signup_date:
        return "❌ 没有日期<br><a href='/volunteer'>返回</a>"

    try:
        output = build_whatsapp_from_assigned(signup_date)
        notice_html = build_signup_shortage_notice(signup_date)
    except Exception as e:
        output = f"❌ 暂时无法生成值班表：{e}"

    return render_template_string("""
    <!doctype html>
    <html>
    <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>当天值班表</title>
    <style>
    body { font-family:"Microsoft YaHei"; background:#f5f5f5; padding:20px; }
    .box { background:white; max-width:900px; margin:auto; padding:25px; border-radius:15px; }
    textarea { width:100%; height:650px; font-size:20px; padding:15px; box-sizing:border-box; }
    a, button { font-size:22px; padding:10px 18px; margin:8px; }
    </style>
    </head>
    <body>
    <div class="box">
    <h1>📋 当天值班表</h1>
    {{ notice_html|safe }}
    <a href="/volunteer">继续报名</a>
    <br><br>
    <textarea readonly>{{ output }}</textarea>
    </div>
    </body>
    </html>
    """, output=output, notice_html=notice_html)

@schedule_bp.route("/volunteer/monthly_signup_list")
def volunteer_monthly_signup_list():
    today = malaysia_today()

    year = int(request.args.get("year", today.year))
    month = int(request.args.get("month", today.month))

    text = build_monthly_signup_text(year, month)

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                select
                    count(*) as total_signup,
                    count(distinct signup_date) as total_days
                from volunteer_schedule_signups
                where extract(year from signup_date) = %s
                  and extract(month from signup_date) = %s
                  and coalesce(status, 'pending') <> 'cancelled'
            """, (year, month))
            summary = cur.fetchone()

            cur.execute("""
                select role, count(*) as cnt
                from volunteer_schedule_signups
                where extract(year from signup_date) = %s
                  and extract(month from signup_date) = %s
                  and coalesce(status, 'pending') <> 'cancelled'
                group by role
            """, (year, month))
            role_rows = cur.fetchall()

    total_signup = summary["total_signup"] or 0
    total_days = summary["total_days"] or 0

    role_counts = {r["role"]: r["cnt"] for r in role_rows}
    duty_count = role_counts.get("值班", 0)
    cleaning_count = role_counts.get("卫生", 0)
    offering_count = role_counts.get("供台", 0)
    meal_count = role_counts.get("膳食", 0)

    return render_template_string("""
    <style>
    body {
        font-family:"Microsoft YaHei", Arial;
        background:#f8f3ea;
        padding:20px;
        font-size:22px;
    }

    .page {
        max-width:1200px;
        margin:20px auto;
        background:white;
        padding:35px;
        border-radius:20px;
        box-shadow:0 4px 15px rgba(0,0,0,0.08);
    }

    h1 {
        color:#8b5a2b;
        font-size:40px;
        margin-top:0;
        margin-bottom:25px;
    }

    .dashboard-header {
        display:flex;
        gap:25px;
        margin-bottom:25px;
    }

    .filter-panel,
    .summary-panel {
        background:#fff8e6;
        border:2px solid #f3d9a5;
        border-radius:18px;
        padding:25px;
        box-sizing:border-box;
    }

    .filter-panel {
        flex:1;
    }

    .summary-panel {
        width:430px;
    }

    .row {
        display:flex;
        align-items:center;
        gap:15px;
        margin-bottom:18px;
    }

    .row label {
        width:80px;
        font-size:28px;
        font-weight:bold;
        color:#5c3b1e;
    }

    .row select {
        width:180px;
        font-size:24px;
        padding:10px;
        border-radius:10px;
        border:1px solid #d6c7b0;
    }

    .search-btn {
        margin-top:8px;
        width:280px;
        height:65px;
        background:#4CAF50;
        color:white;
        border:none;
        border-radius:12px;
        font-size:28px;
        font-weight:bold;
        cursor:pointer;
    }

    .top-summary {
        display:flex;
        justify-content:space-around;
        margin-bottom:16px;
    }

    .summary-item {
        text-align:center;
    }

    .summary-icon {
        font-size:30px;
        margin-bottom:4px;
    }

    .summary-num {
        font-size:50px;
        color:#1f6feb;
        font-weight:bold;
        line-height:1.1;
    }

    .summary-label {
        font-size:20px;
        color:#666;
    }

    .role-grid {
        display:grid;
        grid-template-columns:1fr 1fr;
        gap:14px 22px;
        margin-top:16px;
        font-size:23px;
        color:#5c3b1e;
    }

    .role-grid div {
        display:flex;
        justify-content:space-between;
        gap:12px;
    }

    .role-grid b {
        color:#1f6feb;
        font-size:24px;
    }

    textarea {
        width:100%;
        height:520px;
        font-size:22px;
        line-height:1.8;
        padding:22px;
        border-radius:18px;
        border:2px solid #ddd;
        box-sizing:border-box;
        background:#fffdf8;
        color:#222;
    }

    .action-row {
        display:flex;
        gap:18px;
        margin-top:22px;
        flex-wrap:wrap;
    }

    .big-btn {
        width:300px;
        padding:18px;
        border-radius:14px;
        border:0;
        color:white;
        font-size:26px;
        font-weight:bold;
        text-align:center;
        text-decoration:none;
        box-sizing:border-box;
    }

    .copy-btn { background:#2baa4b; }
    .home-btn { background:#666; }

    @media (max-width:900px) {
        .dashboard-header {
            flex-direction:column;
        }

        .summary-panel {
            width:auto;
        }

        .big-btn {
            width:100%;
        }

        h1 {
            font-size:32px;
        }
    }
    </style>

    <div class="page">

        <h1>📖 {{ year }}年{{ month }}月预报名名单</h1>

        <div class="dashboard-header">

            <div class="filter-panel">
                <form method="get">
                    <div class="row">
                        <label>年份：</label>
                        <select name="year">
                            {% for y in range(year - 1, year + 2) %}
                            <option value="{{ y }}" {% if y == year %}selected{% endif %}>
                                {{ y }}
                            </option>
                            {% endfor %}
                        </select>
                    </div>

                    <div class="row">
                        <label>月份：</label>
                        <select name="month">
                            {% for m in range(1, 13) %}
                            <option value="{{ m }}" {% if m == month %}selected{% endif %}>
                                {{ m }}月
                            </option>
                            {% endfor %}
                        </select>
                    </div>

                    <button type="submit" class="search-btn">
                        🔍 查看
                    </button>
                </form>
            </div>

            <div class="summary-panel">
                <div class="top-summary">
                    <div class="summary-item">
                        <div class="summary-icon">👥</div>
                        <div class="summary-num">{{ total_signup }}</div>
                        <div class="summary-label">总人次</div>
                    </div>

                    <div class="summary-item">
                        <div class="summary-icon">📅</div>
                        <div class="summary-num">{{ total_days }}</div>
                        <div class="summary-label">有报名天数</div>
                    </div>
                </div>

                <hr>

                <div class="role-grid">
                    <div><span>👤 值班</span><b>{{ duty_count }}</b></div>
                    <div><span>🧹 卫生</span><b>{{ cleaning_count }}</b></div>
                    <div><span>🙏 供台</span><b>{{ offering_count }}</b></div>
                    <div><span>🍱 膳食</span><b>{{ meal_count }}</b></div>
                </div>
            </div>

        </div>

        <textarea id="signupText">{{ text }}</textarea>

        <div class="action-row">
            <button class="big-btn copy-btn" onclick="copySignupText()">
                📋 复制 WhatsApp 格式
            </button>

            <a class="big-btn home-btn" href="/volunteer">
                🏠 返回首页
            </a>
        </div>

    </div>

    <script>
    function copySignupText() {
        const text = document.getElementById("signupText");
        text.select();
        text.setSelectionRange(0, 999999);
        document.execCommand("copy");
        alert("✅ 已复制，可以贴去 WhatsApp");
    }
    </script>
    """,
    year=year,
    month=month,
    text=text,
    total_signup=total_signup,
    total_days=total_days,
    duty_count=duty_count,
    cleaning_count=cleaning_count,
    offering_count=offering_count,
    meal_count=meal_count)

@schedule_bp.route("/volunteer/prebook", methods=["GET", "POST"])
def volunteer_prebook():

    if request.method == "GET":
        today = malaysia_today()

        year = int(request.args.get("year", today.year))
        month = int(request.args.get("month", today.month))

        keyword = request.args.get("keyword", "").strip()
        branch = request.args.get("branch", "CHE").strip().upper()

        special_days = {}
        days_in_month = calendar.monthrange(year, month)[1]

        for d in range(1, days_in_month + 1):
            date_obj = date(year, month, d)
            info = get_special_day_info(date_obj)

            if info["template_type"] == "lunar_1_15":
                special_days[d] = "lunar"
            elif info["template_type"] == "buddhist_festival":
                special_days[d] = "festival"

        return render_template_string(
            VOLUNTEER_PREBOOK_HTML,
            default_year=year,
            default_month=month,
            times=TIME_OPTIONS,
            keyword=keyword,
            branch=branch,
            special_days=special_days,
        )

    # POST 提交报名
    keyword = request.form.get("keyword", "").strip()
    branch = request.form.get("branch", "CHE").strip().upper()
    keyword = apply_branch_prefix(keyword, branch)

    year = int(request.form.get("year"))
    month = int(request.form.get("month"))
    days = request.form.getlist("days")
    role = request.form.get("role", "").strip()
    start_time = request.form.get("start_time", "").strip()
    end_time = request.form.get("end_time", "").strip()

    if not days:
        return "❌ 请选择至少一个日期<br><a href='/volunteer/prebook'>返回</a>"

    matches = find_volunteer_by_keyword(keyword)

    if not matches:
        return "❌ 找不到义工，请检查编号 / 姓名<br><a href='/volunteer/prebook'>返回</a>"

    if len(matches) > 1:
        return "❌ 找到多个同名义工，请用义工编号报名<br><a href='/volunteer/prebook'>返回</a>"

    vol = matches[0]
    vol_id = str(vol["id"])
    name = str(vol["name"])

    if role == "值班":
        s_min = time_to_minutes(start_time)
        e_min = time_to_minutes(end_time)

        if s_min is None or e_min is None:
            return "❌ 时间格式错误，请重新选择<br><a href='/volunteer/prebook'>返回</a>"

        if e_min <= s_min:
            return "❌ 结束时间必须比开始时间迟<br><a href='/volunteer/prebook'>返回</a>"
    else:
        start_time = None
        end_time = None

    inserted = 0
    updated = 0
    skipped = 0
    new_signup_items = []

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:

            for d in days:
                try:
                    signup_date = date(year, month, int(d))
                except ValueError:
                    skipped += 1
                    continue

                cur.execute("""
                    select id
                    from volunteer_schedule_signups
                    where volunteer_id = %s
                    and signup_date = %s
                    and role = %s
                    and coalesce(status, 'pending') <> 'cancelled'
                    limit 1
                """, (vol_id, signup_date, role))

                existing = cur.fetchone()

                if existing:
                    cur.execute("""
                        update volunteer_schedule_signups
                        set start_time = %s,
                            end_time = %s,
                            status = 'pending',
                            assigned_place = null,
                            remarks = '义工多日报名更新'
                        where id = %s
                    """, (
                        start_time,
                        end_time,
                        existing["id"]
                    ))

                    updated += 1

                    new_signup_items.append((existing["id"], signup_date))
                else:
                    cur.execute("""
                        insert into volunteer_schedule_signups
                        (volunteer_id, name, signup_date, role, start_time, end_time, status, remarks)
                        values (%s, %s, %s, %s, %s, %s, 'pending', '义工多日报名')
                        returning id
                    """, (
                        vol_id,
                        name,
                        signup_date,
                        role,
                        start_time,
                        end_time
                    ))

                    new_signup_id = cur.fetchone()["id"]
                    new_signup_items.append((new_signup_id, signup_date))
                    inserted += 1

            conn.commit()

    for signup_id, signup_date in new_signup_items:

        try:
            synced = sync_schedule_after_signup_change(
                signup_id,
                action="upsert",
                changed_by="volunteer_multi_signup"
            )

            print(
                f"✅ {signup_date} 同步完成，共新增 {synced} 笔 assignment"
            )

        except Exception as e:
            print(f"多日报名同步失败：{signup_date}：{e}")

    return f"""
    <h1>✅ 多日报名完成</h1>
    <p>义工：{name}</p>
    <p>岗位：{role}</p>
    <p>新增：{inserted} 笔</p>
    <p>更新：{updated} 笔</p>
    <p>跳过无效日期：{skipped} 笔</p>
    <a href="/volunteer/prebook">继续多日报名</a><br>
    <a href="/volunteer">返回首页</a>
    """

@schedule_bp.route("/volunteer/my_schedule_search")
def volunteer_my_schedule_search():
    return render_template_string("""
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>我的报名</title>

<link rel="stylesheet"
      href="{{ url_for('static', filename='css/toolbox.css') }}">

<style>
.search-wrap{
    max-width:760px;
    margin:auto;
    padding:24px;
}

.branch-search-row{
    display:grid;
    grid-template-columns:110px 1fr;
    gap:12px;
    align-items:center;
}

.branch-toggle-btn{
    height:66px;
    font-size:26px;
    font-weight:bold;
    color:white;
    border:none;
    border-radius:16px;
    cursor:pointer;
}

@media (max-width:700px){
    .branch-search-row{
        grid-template-columns:1fr;
    }
}
</style>

<script>
window.addEventListener("DOMContentLoaded", function(){
    document.getElementById("keyword").focus();
});
</script>

</head>

<body>

<div class="search-wrap">

    <div class="card">

        <h1 class="page-title">🔍 我的报名</h1>

        <p class="page-subtitle">
            查看已报名日期、修改报名、取消报名或查看正式岗位。
        </p>

        <div class="alert alert-info">
            💡 可输入义工编号、电话或姓名查询。
        </div>

        <form method="get" action="/volunteer/my_schedule">

            <div class="form-group">

                <label class="form-label">
                    义工编号 / 电话 / 姓名
                </label>

                <div class="branch-search-row">

                    <button
                        type="button"
                        id="branch_btn"
                        onclick="toggleBranch()"
                        class="branch-toggle-btn"
                        style="background:#28a745;">
                        CHE
                    </button>

                    <input
                        type="hidden"
                        id="branch"
                        name="branch"
                        value="CHE">

                    <input
                        class="form-input"
                        name="keyword"
                        id="keyword"
                        required
                        placeholder="例如：108 / 姓名 / 电话">

                </div>

            </div>

            <div class="btn-row">

                <button
                    class="btn-tool btn-blue btn-full"
                    type="submit">
                    📋 查询我的报名
                </button>

                <a
                    class="btn-tool btn-gray btn-full"
                    href="/volunteer">
                    ⬅ 返回义工首页
                </a>

            </div>

        </form>

    </div>

</div>

<script>
function toggleBranch(){

    const btn = document.getElementById("branch_btn");
    const branch = document.getElementById("branch");

    if(branch.value === "CHE"){
        branch.value = "STW";
        btn.innerText = "STW";
        btn.style.background = "#dc3545";
    }else{
        branch.value = "CHE";
        btn.innerText = "CHE";
        btn.style.background = "#28a745";
    }
}
</script>

</body>
</html>
""")

@schedule_bp.route("/volunteer/day_info")
def volunteer_day_info():
    date_str = request.args.get("date", "").strip()

    try:
        date_obj = datetime.strptime(date_str, "%Y-%m-%d").date()
    except:
        try:
            date_obj = datetime.strptime(date_str, "%Y/%m/%d").date()
        except:
            return {
                "ok": False,
                "template_type": "normal",
                "duty_start": "10:00am"
            }

    info = get_special_day_info(date_obj)

    template_type = info.get("template_type", "normal")

    if template_type == "lunar_1_15":
        duty_start = "8:00am"
    elif template_type == "buddhist_festival":
        duty_start = "8:00am"
    else:
        duty_start = "10:00am"

    return {
        "ok": True,
        "template_type": template_type,
        "duty_start": duty_start,
        "lunar_text": info.get("lunar_text", ""),
        "special_names": info.get("special_names", [])
    }
