# volunteer_routes.py

import calendar

from flask import request, redirect, url_for, render_template_string, jsonify
from psycopg2.extras import RealDictCursor

from db import get_conn
from schedule.blueprint import schedule_bp
from datetime import datetime, timedelta, date, timezone

from schedule.helpers import (
    find_volunteer_by_keyword,
    build_monthly_signup_text,
)

from lunar_rules import get_special_day_info
from schedule.constants import TIME_OPTIONS
from schedule.builders.schedule_builder import patch_schedule_for_date
from schedule.builders.time_utils import time_to_minutes, malaysia_today, malaysia_now
from schedule.services.publish_service import is_schedule_published
from schedule.services.whatsapp_service import build_whatsapp_from_assigned
from schedule.services.shortage_service import build_signup_shortage_notice
from schedule.volunteer_templates import (
    VOLUNTEER_SIGNUP_HTML,
    VOLUNTEER_PREBOOK_HTML,
)

@schedule_bp.route("/volunteer")
def volunteer_home():

    now = malaysia_now()
    today = now.date()

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
        
    return render_template_string(
        VOLUNTEER_SIGNUP_HTML,
        default_date=default_date,
        schedule_label=schedule_label,
        times=TIME_OPTIONS,
        prebook_year=prebook_year,
        prebook_month=prebook_month,
    )

@schedule_bp.route("/volunteer/signup", methods=["POST"])
def volunteer_signup():
    keyword = request.form.get("keyword", "").strip()
    signup_date = request.form.get("signup_date", "").strip()
    role = request.form.get("role", "").strip()
    start_time = request.form.get("start_time", "").strip()
    end_time = request.form.get("end_time", "").strip()

    matches = find_volunteer_by_keyword(keyword)

    if not matches:
        return "❌ 找不到义工，请检查编号 / 姓名<br><a href='/volunteer'>返回</a>"

    if len(matches) > 1:
        return "❌ 找到多个同名义工，请用义工编号报名<br><a href='/volunteer'>返回</a>"

    vol = matches[0]
    vol_id = str(vol["id"])
    name = str(vol["name"])

    if role == "值班":
        s_min = time_to_minutes(start_time)
        e_min = time_to_minutes(end_time)

        if s_min is None or e_min is None:
            return "❌ 时间格式错误，请重新选择<br><a href='/volunteer'>返回</a>"

        if e_min <= s_min:
            return "❌ 结束时间必须比开始时间迟<br><a href='/volunteer'>返回</a>"

    elif role == "卫生":
        date_obj = datetime.strptime(signup_date, "%Y-%m-%d").date()
        special_info = get_special_day_info(date_obj)

        if special_info["template_type"] == "buddhist_festival":
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

            if existing:
                cur.execute("""
                    update volunteer_schedule_signups
                    set start_time = %s,
                        end_time = %s,
                        status = 'pending',
                        assigned_place = null,
                        remarks = '义工网页更新报名'
                    where id = %s
                """, (
                    start_time,
                    end_time,
                    existing["id"]
                ))

                conn.commit()

                return f"""
                <h1>✅ 已更新报名</h1>
                <p>{name}</p>
                <p>{signup_date}</p>
                <p>{role}：{start_time} ~ {end_time}</p>
                <p>系统已用新的资料覆盖旧报名。</p>
                <a href="/volunteer/day_schedule?date={signup_date}">查看当天值班表</a><br>
                <a href="/volunteer">继续报名</a>
                """

            cur.execute("""
                insert into volunteer_schedule_signups
                (volunteer_id, name, signup_date, role, start_time, end_time, status, remarks)
                values (%s, %s, %s, %s, %s, %s, 'pending', '义工报名')
                returning id
            """, (
                vol_id,
                name,
                signup_date,
                role,
                start_time,
                end_time
            ))

            signup_id = cur.fetchone()["id"]

            conn.commit()

            if is_schedule_published(signup_date):

                try:

                    if hasattr(signup_date, "strftime"):
                        date_str = signup_date.strftime("%Y-%m-%d")
                    else:
                        date_str = str(signup_date)

                    inserted = patch_schedule_for_date(date_str, only_signup_id=signup_id)

                    print(f"✅ 自动补排完成，共新增 {inserted} 笔")

                except Exception as e:

                    print("自动补排失败：", e)

    return render_template_string("""
    <!doctype html>
    <html>
    <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">

    <style>
    body{
        font-family:"Microsoft YaHei";
        background:#f5f5f5;
        padding:20px;
    }

    .box{
        max-width:700px;
        margin:auto;
        background:white;
        border-radius:20px;
        padding:30px;
        text-align:center;
    }

    .success{
        background:#e8f5e9;
        border:2px solid #4CAF50;
        border-radius:15px;
        padding:25px;
        margin-bottom:25px;
    }

    .success h1{
        color:#2e7d32;
        margin-top:0;
        font-size:42px;
    }

    .info{
        font-size:26px;
        line-height:1.9;
    }

    .notice{
        background:#fff8e1;
        border-radius:12px;
        padding:18px;
        margin:20px 0;
        color:#8d6e00;
        font-size:22px;
        line-height:1.7;
    }

    .btn{
        display:block;
        text-decoration:none;
        color:white;
        padding:18px;
        margin:12px 0;
        border-radius:12px;
        font-size:24px;
        font-weight:bold;
    }

    .blue{
        background:#2196F3;
    }

    .orange{
        background:#FF9800;
    }

    .green{
        background:#4CAF50;
    }

    .gray{
        background:#607D8B;
    }
    </style>
    </head>

    <body>

    <div class="box">

    <div class="success">

    <h1>🎉 报名成功</h1>

    <div class="info">

    🙏 感恩发心护持观音堂<br><br>

    义工：{{ name }}<br>

    日期：{{ signup_date }}<br>

    岗位：{{ role }}

    {% if role == "值班" %}
    <br>
    时间：{{ start_time }} ~ {{ end_time }}
    {% endif %}

    <br><br>

    状态：等待负责人安排

    </div>

    </div>

    <div class="notice">

    ⚠️ 当前属于报名阶段。<br>

    最终岗位安排请以负责人公布的正式值班表为准。<br>

    请多留意义工群信息，感恩大家护持观音堂 🙏

    </div>

    <a class="btn blue"
    href="/volunteer/my_schedule_search">
    📋 查看我的报名
    </a>

    <a class="btn orange"
    href="/volunteer/day_schedule?date={{ signup_date }}">
    📅 查看当天值班表
    </a>

    <a class="btn green"
    href="/volunteer">
    ➕ 继续报名
    </a>
    
    </div>

    </body>
    </html>
    """,
    name=name,
    signup_date=signup_date,
    role=role,
    start_time=start_time,
    end_time=end_time
    )

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
    <h1>我的报名记录</h1>
    <h2>{name}</h2>
    <a href="/volunteer">返回报名</a>
    <hr>
    """

    if not rows:
        html += "<p>暂时没有报名记录。</p>"
    else:
        for r in rows:
            assigned_place = r.get("assigned_place") or "尚未安排"
            status = str(r.get("status") or "pending")

            html += f"""
            <p style="font-size:22px;">
            📅 {r["signup_date"]}<br>
            岗位：{r["role"]}<br>
            时间：{r["start_time"]} ~ {r["end_time"]}<br>
            系统安排：{assigned_place}<br>
            状态：{status}
            </p>
            """

            if status == "pending":
                html += f"""
                <form method="post"
                      action="/volunteer/cancel/{r['id']}"
                      onsubmit="return confirm('确定取消报名？');">

                    <button type="submit">
                    ❌ 取消报名
                    </button>

                </form>
                """
            else:
                html += """
                <p style="color:#b36b00; font-size:20px;">
                ⚠️ 已进入正式排班，如需取消请联系负责人。
                </p>
                """

            html += "<hr>"

    return html

@schedule_bp.route(
    "/volunteer/cancel/<int:signup_id>",
    methods=["POST"]
)
def volunteer_cancel_signup(signup_id):

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:

            cur.execute("""
                select
                    id,
                    name,
                    signup_date,
                    role,
                    coalesce(status, 'pending') as status
                from volunteer_schedule_signups
                where id = %s
            """, (signup_id,))

            row = cur.fetchone()

            if not row:
                return """
                <h1>❌ 找不到报名记录</h1>
                <a href="/volunteer">返回</a>
                """

            if row["status"] == "assigned":
                return f"""
                <h1>❌ 已安排值班</h1>

                <p>
                {row["name"]}
                </p>

                <p>
                {row["signup_date"]}
                </p>

                <p>
                {row["role"]}
                </p>

                <p style="color:red;">
                已经安排值班，请联系负责人取消。
                </p>

                <a href="/volunteer">
                返回
                </a>
                """

            cur.execute("""
                update volunteer_schedule_signups
                set
                    status = 'cancelled',
                    assigned_place = null,
                    remarks = '义工自行取消报名'
                where id = %s
            """, (signup_id,))

            conn.commit()

    return """
    <h1>✅ 已取消报名</h1>

    <a href="/volunteer">
    返回义工报名
    </a>
    """

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

    year = request.args.get("year", today.year)
    month = request.args.get("month", today.month)

    year = int(year)
    month = int(month)

    text = build_monthly_signup_text(year, month)

    return render_template_string("""
    <h1>📖 {{ year }}年{{ month }}月预报名名单</h1>

    <form method="get">
        <label>年份：</label>
        <select name="year" style="width:110px;">
            {% for y in range(year - 1, year + 2) %}
                <option value="{{ y }}" {% if y == year %}selected{% endif %}>
                    {{ y }}
                </option>
            {% endfor %}
        </select>

        <label>月份：</label>
        <select name="month" style="width:90px;">
            {% for m in range(1, 13) %}
                <option value="{{ m }}" {% if m == month %}selected{% endif %}>
                    {{ m }}月
                </option>
            {% endfor %}
        </select>

        <button type="submit">查看</button>
    </form>

    <br>

    <textarea id="signupText" style="width:100%; height:500px; font-size:18px;">{{ text }}</textarea>

    <br><br>

    <div style="
        display:flex;
        gap:15px;
        margin-top:15px;
    ">

        <button
            onclick="copySignupText()"
            style="width:280px;"
        >
            📋 复制 WhatsApp 格式
        </button>
                                
        <a href="/volunteer"
        style="
            display:flex;
            align-items:center;
            justify-content:center;
            width:220px;
            text-decoration:none;
            background:#666;
            color:white;
            border-radius:10px;
            font-size:24px;
            font-weight:bold;
        ">
            🏠 返回首页
        </a>

    </div>

    <script>
    function copySignupText() {
        const text = document.getElementById("signupText");
        text.select();
        text.setSelectionRange(0, 999999);
        document.execCommand("copy");
        alert("已复制，可以贴去 WhatsApp");
    }
    </script>
    """, year=year, month=month, text=text)

@schedule_bp.route("/volunteer/prebook", methods=["GET", "POST"])
def volunteer_prebook():

    if request.method == "GET":
        today = malaysia_today()

        year = int(request.args.get("year", today.year))
        month = int(request.args.get("month", today.month))

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
            special_days=special_days
        )

    keyword = request.form.get("keyword", "").strip()
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

                    new_signup_items.append((new_signup_id, signup_date))
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

        if not is_schedule_published(signup_date):
            continue

        try:
            if hasattr(signup_date, "strftime"):
                date_str = signup_date.strftime("%Y-%m-%d")
            else:
                date_str = str(signup_date)

            inserted = patch_schedule_for_date(date_str, only_signup_id=signup_id)

            print(
                f"✅ {signup_date} 自动补排完成，共新增 {inserted} 笔"
            )

        except Exception as e:
            print(f"自动补排失败：{signup_date}：{e}")

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
<style>
body { font-family:"Microsoft YaHei"; background:#f5f5f5; padding:20px; font-size:24px; }
.box { background:white; max-width:700px; margin:auto; padding:25px; border-radius:15px; }
input, button { font-size:28px; padding:14px; width:100%; box-sizing:border-box; margin:10px 0; }
a { font-size:22px; }
</style>
</head>
<body>
<div class="box">

<h1>🔍 我的报名</h1>

<form method="get" action="/volunteer/my_schedule">
    <label>义工编号 / 电话 / 姓名</label>
    <input name="keyword" required placeholder="例如 CHE-108 / 108 / 姓名">
    <button type="submit">查询我的报名</button>
</form>

<br>
<a href="/volunteer">⬅ 返回首页</a>

</div>
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
