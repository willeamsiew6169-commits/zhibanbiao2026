# schedule_web.py

import os
import calendar
import psycopg2
import pandas as pd
import schedule.routes

from opencc import OpenCC
from dotenv import load_dotenv
from db import get_db, get_conn
from supabase import create_client
from openpyxl import load_workbook
from psycopg2.extras import RealDictCursor
from sqlalchemy import create_engine, text
from schedule.blueprint import schedule_bp
from datetime import datetime, timedelta, date, timezone
from schedule.services.admin_dashboard_service import load_admin_dashboard_data
from lunar_rules import get_special_day_info, get_next_day_remove_info
from flask import Blueprint, request, session, redirect, url_for, render_template_string, flash
from schedule.builders.schedule_builder import (
    run_schedule_for_date,
    patch_schedule_for_date,
    parse_signup_line_multi,
    normalize_time_text,
    build_buddhist_festival_message,
    get_special_day_info,
    get_next_day_remove_info,
    build_lunar_1_15_message,
    build_normal_message,
    get_duty_targets,
)

from schedule.helpers import (
    to_simple,
    normalize_vol_id_for_search,
    time_to_minutes,
    find_volunteer_by_keyword,
    build_monthly_signup_text
)

from schedule.services.settings_service import (
    get_schedule_setting,
    get_schedule_settings,
    save_schedule_setting,
)

from schedule.services.publish_service import (
    is_schedule_published,
    publish_schedule_for_date,
    unpublish_schedule_for_date,
)

from schedule.services.assignment_service import (
    load_assigned_places_for_date,
    load_display_records,
    load_schedule_admin_dashboard_data,
)

from schedule.services.shortage_service import (
    build_shortage_notice_from_assignments,
    build_shortage_summary_for_admin,
    build_signup_shortage_notice
)

from schedule.services.supply_service import (
    load_supply_signups_for_date,
    load_upcoming_supply_signup_alerts,
    load_day_flags,
)

from schedule.constants import (
    PREBOOK_FILE,
    TIME_OPTIONS,
    ROLE_OPTIONS,
)

from schedule.services.whatsapp_service import (
    build_whatsapp_from_assigned
)

load_dotenv()

cc = OpenCC('t2s')  # 繁 → 简

schedule_bp = Blueprint("schedule", __name__)

DATABASE_URL = os.environ.get("DATABASE_URL")
engine = create_engine(DATABASE_URL, pool_pre_ping=True) if DATABASE_URL else None

DATABASE_URL = os.environ.get("DATABASE_URL")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DATA_DIR = os.path.join(BASE_DIR, "data")

PREBOOK_FILE = os.path.join(
    DATA_DIR,
    "prebook_schedule.xlsx"
)

SCHEDULE_PIN = "1234"

schedule_records = []

def get_supabase_client():
    return create_client(
        os.getenv("SUPABASE_URL"),
        os.getenv("SUPABASE_SERVICE_ROLE_KEY")
    )


def display_report_name(filename):
    name = filename.replace(".xlsx", "")

    if name == "2026_year_report":
        return "2026 年报"

    if "_month_report" in name:
        parts = name.split("_")
        year = parts[0]
        month = parts[1]
        return f"{year}-{month} 月报"

    return filename


def get_pending_signup_counts():
    from datetime import date, timedelta
    from db import db_query

    today = date.today()
    tomorrow = today + timedelta(days=1)
    month_start = today.replace(day=1)

    row = db_query("""
        select
            count(*) filter (
                where signup_date = %s
            ) as today_count,

            count(*) filter (
                where signup_date = %s
            ) as tomorrow_count,

            count(*) filter (
                where signup_date >= %s
            ) as month_count

        from volunteer_schedule_signups
        where coalesce(status, 'pending') = 'pending'
        and signup_date >= %s
    """, (
        today,
        tomorrow,
        month_start,
        month_start
    ), fetchone=True)

    return {
        "today": row["today_count"] or 0,
        "tomorrow": row["tomorrow_count"] or 0,
        "month": row["month_count"] or 0,
    }


def search_volunteer_names(keyword, volunteers):
    keyword_simple = to_simple(keyword)

    if not keyword_simple:
        return []

    results = []

    for v in volunteers:
        name = str(v.get("name") or v.get("姓名") or "").strip()
        phone = str(v.get("phone") or v.get("电话") or v.get("电话号码") or "").strip()

        name_simple = to_simple(name)

        if keyword_simple in name_simple or keyword_simple in phone:
            results.append(v)

    return results

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
    
from opencc import OpenCC
cc = OpenCC("t2s")


@schedule_bp.route("/schedule/logout")
def schedule_logout():

    session.pop("schedule_login", None)

    return redirect("/volunteer")


@schedule_bp.route(
    "/schedule/signup_cancel/<int:signup_id>",
    methods=["POST"]
)
def schedule_signup_cancel(signup_id):

    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    with get_conn() as conn:
        with conn.cursor() as cur:

            cur.execute("""
                update volunteer_schedule_signups
                set
                    status = 'cancelled',
                    assigned_place = null,
                    remarks = '负责人取消报名'
                where id = %s
            """, (signup_id,))

            conn.commit()

    return redirect("/schedule/signups")

    
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

    if role in ["卫生", "供台", "整理佛台"]:
        return None, None

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

    
def get_signup_summary_for_date(date_str):
    from db import db_query

    row = db_query("""
        select
            count(*) as total,
            count(*) filter (
                where coalesce(status, 'pending') = 'pending'
            ) as pending,
            count(*) filter (
                where status = 'assigned'
            ) as assigned,
            count(*) filter (
                where status = 'cancelled'
            ) as cancelled
        from volunteer_schedule_signups
        where signup_date = %s
    """, (date_str,), fetchone=True)

    return {
        "total": row["total"] or 0,
        "pending": row["pending"] or 0,
        "assigned": row["assigned"] or 0,
        "cancelled": row["cancelled"] or 0,
    }
    

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


@schedule_bp.route("/schedule/copy_whatsapp", methods=["POST"])
def schedule_copy_whatsapp():
    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    date_str = request.form.get("date", "").strip()

    if not date_str:
        return "❌ 没有日期<br><a href='/schedule?mode=day'>返回</a>"

    output = build_whatsapp_from_assigned(date_str)

    return render_template_string(DAY_OUTPUT_HTML, output=output)


def is_big_day(date_str):
    from datetime import datetime

    date_obj = datetime.strptime(date_str, "%Y-%m-%d").date()
    info = get_special_day_info(date_obj)

    return info.get("template_type") in [
        "lunar_1_15",
        "buddhist_festival"
    ]


@schedule_bp.route(
    "/schedule/signup_restore/<int:signup_id>",
    methods=["POST"]
)
def schedule_signup_restore(signup_id):
    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                update volunteer_schedule_signups
                set
                    status = 'pending',
                    assigned_place = null,
                    remarks = '负责人恢复报名'
                where id = %s
            """, (signup_id,))

            conn.commit()

    return redirect("/schedule/signups")


def build_month_prebook_message(year, month):
    import calendar
    from datetime import date

    first_day = date(year, month, 1)
    last_day = date(year, month, calendar.monthrange(year, month)[1])

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                select
                    signup_date,
                    name,
                    role,
                    start_time,
                    end_time
                from volunteer_schedule_signups
                where signup_date between %s and %s
                and coalesce(status, 'pending') <> 'cancelled'
                order by signup_date, role, start_time, name
            """, (first_day, last_day))
            rows = cur.fetchall()

    by_day = {}
    for r in rows:
        d = r["signup_date"]
        by_day.setdefault(d, []).append(r)

    weekday_map = {
        0: "星期一",
        1: "星期二",
        2: "星期三",
        3: "星期四",
        4: "星期五",
        5: "星期六",
        6: "星期日",
    }

    lines = []
    lines.append(f"{year}年{month}月份预报名名单")
    lines.append("")
    lines.append("以下为目前已收到的预报名资料，最终岗位安排以负责人公布值班表为准。")
    lines.append("")

    days_in_month = calendar.monthrange(year, month)[1]

    for day in range(1, days_in_month + 1):
        d = date(year, month, day)
        day_rows = by_day.get(d, [])

        if not day_rows:
            continue

        cleaning = []
        supply = []
        duty = []
        others = []

        for r in day_rows:
            name = str(r["name"]).strip()
            role = str(r["role"]).strip()
            start = r.get("start_time") or ""
            end = r.get("end_time") or ""

            if role == "卫生":
                cleaning.append(name)

            elif role == "供台":
                supply.append(name)

            elif role == "值班":
                if start and end:
                    duty.append(f"{name} {start}~{end}")
                else:
                    duty.append(name)

            else:
                others.append(f"{name}（{role}）")

        lines.append(f"{day}/{month}/{year} {weekday_map[d.weekday()]}")

        if cleaning:
            lines.append("卫生：" + "、".join(cleaning))

        if supply:
            lines.append("供台：" + "、".join(supply))

        if duty:
            lines.append("值班：")
            lines.extend(duty)

        if others:
            lines.append("其他：")
            lines.extend(others)

        lines.append("")

    return "\n".join(lines)


@schedule_bp.route("/schedule/edit_assigned/<int:assignment_id>", methods=["GET", "POST"])
def edit_assigned_place(assignment_id):

    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    shift_options = ["绿班", "橙班", "黄班"]
    place_options = [
        "观音堂",
        "活动中心",
        "佛堂卫生",
        "二楼卫生",
        "楼梯卫生",
        "设师父供台",
        "整理佛台",
    ]

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:

            if request.method == "POST":
                shift_label = request.form.get("shift_label", "").strip()
                assigned_place = request.form.get("assigned_place", "").strip()
                start_time = request.form.get("start_time", "").strip()
                end_time = request.form.get("end_time", "").strip()

                cur.execute("""
                    update volunteer_schedule_assignments
                    set
                        shift_label = %s,
                        assigned_place = %s,
                        start_time = %s,
                        end_time = %s,
                        remarks = '负责人调整排班'
                    where id = %s
                """, (
                    shift_label or None,
                    assigned_place,
                    start_time or None,
                    end_time or None,
                    assignment_id
                ))

                conn.commit()

                return redirect("/schedule/signups")

            cur.execute("""
                select *
                from volunteer_schedule_assignments
                where id = %s
            """, (assignment_id,))

            row = cur.fetchone()

    if not row:
        return "❌ 找不到排班记录<br><a href='/schedule/signups'>返回</a>"

    return render_template_string("""
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>调整排班</title>
<style>
body { font-family:"Microsoft YaHei", Arial; background:#f5f5f5; padding:20px; }
.box { background:white; max-width:700px; margin:auto; padding:25px; border-radius:15px; }
input, select, button { font-size:24px; padding:10px; margin:8px 0; width:100%; box-sizing:border-box; }
p { font-size:22px; }
</style>
</head>
<body>
<div class="box">

<h1>✏️ 调整排班</h1>

<p><b>{{ row.name }}</b>（{{ row.volunteer_id or "" }}）</p>
<p>日期：{{ row.assignment_date }}</p>
<p>岗位：{{ row.role }}</p>

<form method="post">

    <label>班别</label>
    <select name="shift_label">
        <option value="">无</option>
        {% for s in shift_options %}
        <option value="{{ s }}" {% if row.shift_label == s %}selected{% endif %}>
            {{ s }}
        </option>
        {% endfor %}
    </select>

    <label>地点</label>
    <select name="assigned_place" required>
        {% for p in place_options %}
        <option value="{{ p }}" {% if row.assigned_place == p %}selected{% endif %}>
            {{ p }}
        </option>
        {% endfor %}
    </select>

    <label>开始时间</label>
    <input name="start_time" value="{{ row.start_time or '' }}" placeholder="例如 11:00am">

    <label>结束时间</label>
    <input name="end_time" value="{{ row.end_time or '' }}" placeholder="例如 2:00pm">

    <button type="submit">💾 保存调整</button>

</form>

<br>
<a href="/schedule/signups">⬅ 返回报名管理</a>

</div>
</body>
</html>
""",
        row=row,
        shift_options=shift_options,
        place_options=place_options
    )


def auto_patch_new_signup_to_assignment(signup_id):
    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:

            cur.execute("""
                select *
                from volunteer_schedule_signups
                where id = %s
                and coalesce(status, 'pending') <> 'cancelled'
            """, (signup_id,))
            s = cur.fetchone()

            if not s:
                return

            cur.execute("""
                select id
                from volunteer_schedule_assignments
                where signup_id = %s
                and coalesce(status, 'assigned') <> 'cancelled'
                limit 1
            """, (signup_id,))
            exists = cur.fetchone()

            if exists:
                return

            role = s["role"]
            signup_date = s["signup_date"]
            name = s["name"]
            volunteer_id = s["volunteer_id"]
            start_time = s["start_time"]
            end_time = s["end_time"]

            if role == "供台":
                return

            if role == "卫生":
                assigned_place = "佛堂卫生"
                shift_label = "卫生"

            else:
                s_min = time_to_minutes(start_time)
                e_min = time_to_minutes(end_time)

                if s_min is None or e_min is None:
                    return

                if s_min < 14 * 60:
                    shift_label = "橙"
                else:
                    shift_label = "黄"

                cur.execute("""
                    select assigned_place, count(*) as cnt
                    from volunteer_schedule_assignments
                    where assignment_date = %s
                    and shift_label = %s
                    and assigned_place in ('观音堂', '活动中心')
                    and coalesce(status, 'assigned') <> 'cancelled'
                    group by assigned_place
                """, (signup_date, shift_label))

                count_rows = cur.fetchall()
                counts = {"观音堂": 0, "活动中心": 0}

                for r in count_rows:
                    counts[r["assigned_place"]] = r["cnt"]

                if counts["观音堂"] <= counts["活动中心"]:
                    assigned_place = "观音堂"
                else:
                    assigned_place = "活动中心"

            cur.execute("""
                insert into volunteer_schedule_assignments
                (signup_id, volunteer_id, name, assignment_date, role,
                 shift_label, assigned_place, start_time, end_time, status, remarks)
                values (%s, %s, %s, %s, %s,
                        %s, %s, %s, %s, 'assigned', '报名后自动补排')
            """, (
                s["id"],
                volunteer_id,
                name,
                signup_date,
                role,
                shift_label,
                assigned_place,
                start_time,
                end_time
            ))

            cur.execute("""
                update volunteer_schedule_signups
                set status = 'assigned'
                where id = %s
            """, (s["id"],))

            conn.commit()


def build_shortage_notice(date_str):

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:

            cur.execute("""
                select role
                from volunteer_schedule_signups
                where signup_date = %s
                and coalesce(status,'pending') <> 'cancelled'
            """, (date_str,))

            rows = cur.fetchall()

    cleaning_count = 0
    duty_count = 0

    for r in rows:

        role = str(r.get("role") or "").strip()

        if role == "卫生":
            cleaning_count += 1

        elif role == "值班":
            duty_count += 1

    notices = []

    if cleaning_count < 2:
        notices.append(
            f"卫生可以多加 {2-cleaning_count} 位义工。"
        )

    if duty_count < 6:
        notices.append(
            "全日值班可以多加 2~4 位义工。"
        )

    if not notices:
        return f"""
师兄们，大家好！

{date_str} 的报名人数暂时充足。

感恩大家发心护持观音堂。
🙏🙏🙏
"""

    body = "\n".join(notices)

    return f"""
师兄们，大家好！

{body}

感恩大家🙏🙏🙏
"""


@schedule_bp.route("/schedule", methods=["GET", "POST"])
@schedule_bp.route("/schedule/admin", methods=["GET", "POST"])
def schedule_admin():
    import time

    if not session.get("schedule_login"):
        if request.method == "POST":
            pin = request.form.get("pin", "").strip()

            if pin == SCHEDULE_PIN:
                session.permanent = True
                session["schedule_login"] = True
                return redirect(url_for("schedule.schedule_admin"))

            return "❌ PIN 错误<br><a href='/schedule'>返回</a>"

        return LOGIN_HTML

    t0 = time.time()

    now = datetime.now()

    switch_time = get_schedule_setting(
        "default_day_switch_time",
        "18:00"
    )

    try:
        switch_hour, switch_minute = map(int, switch_time.split(":"))
    except:
        switch_hour, switch_minute = 18, 0

    change_time = now.replace(
        hour=switch_hour,
        minute=switch_minute,
        second=0,
        microsecond=0
    )

    if now >= change_time:
        default_schedule_date = (now + timedelta(days=1)).strftime("%Y-%m-%d")
    else:
        default_schedule_date = now.strftime("%Y-%m-%d")

    mode = request.args.get("mode", "")

    override_date = request.args.get("override_date") or default_schedule_date
    dashboard = load_admin_dashboard_data(
        mode=mode,
        override_date=override_date,
    )

    buddha_names = load_buddha_name_options()
    fixed_buddha_today = get_fixed_buddha_for_date(override_date)

    print("schedule_admin total:", round(time.time() - t0, 2))

    return render_template_string(
        SCHEDULE_HTML,
        mode=mode,
        times=TIME_OPTIONS,
        roles=ROLE_OPTIONS,
        tomorrow=default_schedule_date,
        buddha_names=buddha_names,
        override_date=override_date,
        fixed_buddha_today=fixed_buddha_today,
        records=dashboard["records"],
        pending_counts=dashboard["pending_counts"],
        day_summary=dashboard["day_summary"],
        is_today_or_past=dashboard["is_today_or_past"],
        day_flags=dashboard["day_flags"],
        special_day_info=dashboard["special_day_info"],
        remove_info=dashboard["remove_info"],
        shortage_summary=dashboard["shortage_summary"],
        supply_signups=dashboard["supply_signups"],
        supply_alerts=dashboard["supply_alerts"],
        is_published=dashboard["is_published"],
        default_year=dashboard["default_year"],
        default_month=dashboard["default_month"],
    )


# ============================================================
# ⚠️ LEGACY：旧版负责人录入入口
#
# 这支 route 属于旧流程：
# - 使用 schedule_records
# - 使用 prebook_schedule.xlsx
# - 使用 save_prebook_record()
# - 使用旧版 monthly_prebook_message
#
# 新流程已经改为：
# signups → assignments → publish → patch
#
# 以后不要再把新功能写进这里。
# 如需新增功能，请另开新的 route / service。
# ============================================================
@schedule_bp.route("/schedule/add", methods=["POST"])
def schedule_add():
    print("⚠️ LEGACY route used: /schedule/add")
    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    mode = request.form.get("mode", "").strip()

    vol_keyword = request.form.get("vol_id", "").strip()
    roles = request.form.getlist("roles")
    start_time = request.form.get("start_time", "").strip()
    end_time = request.form.get("end_time", "").strip()

    action = request.form.get("action", "add")
    single_date = request.form.get("single_date", "").strip()
    year = request.form.get("year", "").strip()
    month = request.form.get("month", "").strip()
    days = request.form.getlist("days")

    if mode == "prebook" and action == "generate_monthly":
        return """
        <h1>⚠️ 旧版月预报名功能已停用</h1>
        <p>请使用新的「多日报名」和「查看预报名名单」。</p>
        <a href="/schedule?mode=prebook">返回</a>
        """
        year = request.form.get("year", "").strip()
        month = request.form.get("month", "").strip()

        try:
            output = generate_monthly_prebook_message(int(year), int(month))
        except Exception as e:
            output = f"❌ 生成失败：{e}"

        return render_template_string(
            MONTHLY_PREBOOK_HTML,
            output=output
        )

    if mode == "day" and action == "parse_raw":
        if not single_date:
            single_date = request.form.get("date", "").strip()

        if not single_date:
            return "❌ 请先选择日期<br><a href='/schedule?mode=day'>返回</a>"

        raw_signup = request.form.get("raw_signup", "").strip()

        if not raw_signup:
            return "❌ 请先贴上 WhatsApp 报名内容<br><a href='/schedule?mode=day'>返回</a>"

        lines = raw_signup.splitlines()

        for line in lines:
            line = normalize_time_text(line.strip())
            if not line:
                continue

            parsed_list = parse_signup_line_multi(line)

            for p in parsed_list:
                name = str(p.get("姓名", "")).strip()
                role = str(p.get("岗位", "值班")).strip()

                job_start = p.get("开始时间") or "10:00am"
                job_end = p.get("结束时间") or "2:00pm"

                record = {
                    "日期": single_date,
                    "编号": "",
                    "姓名": name,
                    "岗位": role,
                    "开始时间": job_start,
                    "结束时间": job_end,
                    "备注": "WhatsApp解析",
                }

                schedule_records.append(record)
                save_prebook_record(record)

        return redirect(url_for("schedule.schedule_admin", mode="day"))

    if not vol_keyword:
        return "❌ 请填写义工编号<br><a href='/schedule/admin'>返回</a>"

    if not roles:
        return "❌ 请至少选择一个岗位<br><a href='/schedule/admin'>返回</a>"

    matches = find_volunteer_by_keyword(vol_keyword)

    if not matches:
        return "❌ 找不到义工<br><a href='/schedule'>返回</a>"

    if len(matches) > 1:
        return render_template_string("""
        <h3>找到多个义工，请选择：</h3>

        {% for v in matches %}
            <form method="post" action="/schedule/prebook_add">
                <input type="hidden" name="vol_id" value="{{ v.id }}">
                <input type="hidden" name="mode" value="{{ mode }}">
                <input type="hidden" name="action" value="{{ action }}">

                <input type="hidden" name="year" value="{{ year }}">
                <input type="hidden" name="month" value="{{ month }}">

                {% for d in days %}
                    <input type="hidden" name="days" value="{{ d }}">
                {% endfor %}

                {% for r in roles %}
                    <input type="hidden" name="roles" value="{{ r }}">
                {% endfor %}

                <input type="hidden" name="start_time" value="{{ start_time }}">
                <input type="hidden" name="end_time" value="{{ end_time }}">
                <input type="hidden" name="single_date" value="{{ single_date }}">

                <button type="submit">{{ v.name }} ({{ v.id }})</button>
            </form>
        {% endfor %}
        """,
        matches=matches,
        mode=mode,
        action=action,
        roles=roles,
        start_time=start_time,
        end_time=end_time,
        single_date=single_date,
        year=year,
        month=month,
        days=days
        )

    vol = matches[0]
    vol_id = str(vol["id"])
    name = str(vol["name"])

    if mode == "day":
        if not single_date:
            single_date = request.form.get("date", "").strip()

        if not single_date:
            return "❌ 请先选择日期<br><a href='/schedule?mode=day'>返回</a>"

        date_list = [single_date]

    elif mode == "prebook":
        if not year or not month:
            return "❌ 请选择年份和月份<br><a href='/schedule?mode=prebook'>返回</a>"

        if action == "generate_monthly":
            return """
            <h1>⚠️ 旧版月预报名功能已停用</h1>
            <p>请使用新的「多日报名」和「查看预报名名单」。</p>
            <a href="/schedule?mode=prebook">返回</a>
            """
            try:
                output = generate_monthly_prebook_message(
                    int(year),
                    int(month)
                )
            except Exception as e:
                output = f"❌ 生成失败：{e}"

            return render_template_string(
                MONTHLY_PREBOOK_HTML,
                output=output
            )

        if not days:
            return "❌ 请至少选择一天<br><a href='/schedule?mode=prebook'>返回</a>"

        month_full = f"{int(year)}-{int(month):02d}"
        date_list = [
            f"{month_full}-{int(day):02d}"
            for day in days
        ]

    else:
        return "❌ 模式错误<br><a href='/schedule'>返回</a>"

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

    return redirect(url_for("schedule.schedule_admin", mode=mode))


@schedule_bp.route("/schedule/prebook_add", methods=["POST"])
def schedule_prebook_add():

    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    keyword = request.form.get("vol_id", "").strip()
    single_date = request.form.get("single_date", "").strip()

    roles = request.form.getlist("roles")

    start_time = request.form.get("start_time", "").strip()
    end_time = request.form.get("end_time", "").strip()

    if not keyword:
        return "❌ 请输入义工编号 / 姓名<br><a href='/schedule?mode=day'>返回</a>"

    if not single_date:
        return "❌ 缺少报名日期<br><a href='/schedule?mode=day'>返回</a>"

    if not roles:
        return "❌ 请选择岗位<br><a href='/schedule?mode=day'>返回</a>"

    try:
        signup_date = datetime.strptime(single_date, "%Y-%m-%d").date()
    except ValueError:
        return "❌ 日期格式错误<br><a href='/schedule?mode=day'>返回</a>"

    matches = find_volunteer_by_keyword(keyword)

    if not matches:
        return "❌ 找不到义工<br><a href='/schedule?mode=day'>返回</a>"

    if len(matches) > 1:
        return "❌ 找到多个同名义工，请用义工编号查询<br><a href='/schedule?mode=day'>返回</a>"

    vol = matches[0]
    volunteer_id = str(vol["id"])
    name = str(vol["name"])

    inserted = 0
    skipped = 0

    with get_db() as conn:
        with conn.cursor() as cur:

            for role in roles:

                role = str(role).strip()

                if role == "值班":
                    role_start = start_time
                    role_end = end_time

                    if not role_start or not role_end:
                        skipped += 1
                        continue

                    s_min = time_to_minutes(role_start)
                    e_min = time_to_minutes(role_end)

                    if s_min is None or e_min is None or e_min <= s_min:
                        skipped += 1
                        continue

                else:
                    role_start = None
                    role_end = None

                cur.execute("""
                    select id
                    from volunteer_schedule_signups
                    where volunteer_id = %s
                    and signup_date = %s
                    and role = %s
                    and coalesce(status, 'pending') <> 'cancelled'
                    limit 1
                """, (
                    volunteer_id,
                    signup_date,
                    role
                ))

                exists = cur.fetchone()

                if exists:
                    skipped += 1
                    continue

                cur.execute("""
                    insert into volunteer_schedule_signups (
                        volunteer_id,
                        name,
                        signup_date,
                        role,
                        start_time,
                        end_time,
                        status,
                        remarks,
                        created_at
                    )
                    values (
                        %s, %s, %s, %s, %s, %s,
                        'pending',
                        '负责人代报名',
                        now()
                    )
                """, (
                    volunteer_id,
                    name,
                    signup_date,
                    role,
                    role_start,
                    role_end
                ))

                inserted += 1

        conn.commit()

    return f"""
    <h1>✅ 负责人代报名已加入</h1>
    <p>日期：{signup_date}</p>
    <p>义工：{name}</p>
    <p>成功加入：{inserted} 笔</p>
    <p>跳过重复/无效：{skipped} 笔</p>
    <a href="/schedule?mode=day&override_date={signup_date}">返回当天安排</a><br>
    <a href="/schedule/admin">返回后台</a>
    """
    
@schedule_bp.route("/schedule/generate_day", methods=["POST"])
def schedule_generate_day():
    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    date = (
        request.form.get("date")
        or request.form.get("single_date")
        or request.args.get("date")
        or ""
    ).strip()

    if not date:
        return "❌ 没有收到日期，请返回选择日期<br><a href='/schedule?mode=day'>返回</a>"

    target_date = datetime.strptime(date, "%Y-%m-%d").date()

    if target_date < datetime.today().date():
        return """
        <h1>❌ 当天排班已锁定</h1>
        <p>今天或过去日期不能重新自动排班，只能使用「补入待安排」。</p>
        <a href="/schedule?mode=day">返回</a>
        """

    try:
        output = run_schedule_for_date(date)
    except Exception as e:
        output = f"❌ 生成失败：{e}"

    return render_template_string(DAY_OUTPUT_HTML, output=output)


@schedule_bp.route("/schedule/monthly_prebook")
def schedule_monthly_prebook():

    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    year = int(request.args.get("year"))
    month = int(request.args.get("month"))

    output = build_month_prebook_message(year, month)

    return render_template_string(
        MONTHLY_PREBOOK_HTML,
        output=output
    )


@schedule_bp.route("/schedule/clear", methods=["POST"])
def schedule_clear():
    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    schedule_records.clear()
    return redirect(url_for("schedule.schedule_admin"))

@schedule_bp.route("/schedule/delete/<int:index>", methods=["POST"])
def schedule_delete(index):
    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    mode = request.form.get("mode", "day")

    if 0 <= index < len(schedule_records):
        record = schedule_records[index]
        delete_prebook_record(record)
        schedule_records.pop(index)

    return redirect(url_for("schedule.schedule_admin", mode=mode))
      

def load_pending_signups_for_date(date_str):
    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                select
                    id,
                    volunteer_id,
                    name,
                    role,
                    start_time,
                    end_time
                from volunteer_schedule_signups
                where signup_date = %s
                and coalesce(status, 'pending') = 'pending'
                order by role, start_time, name
            """, (date_str,))

            return cur.fetchall()
        

def get_recent_cleaning_places(volunteer_id, before_date, limit=5):
    if not volunteer_id:
        return []

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                select assigned_place
                from volunteer_schedule_assignments
                where volunteer_id = %s
                and assignment_date < %s
                and role = '卫生'
                and assigned_place in ('佛堂卫生', '二楼卫生', '楼梯卫生')
                and coalesce(status, 'assigned') <> 'cancelled'
                order by assignment_date desc, id desc
                limit %s
            """, (volunteer_id, before_date, limit))

            rows = cur.fetchall()

    return [r["assigned_place"] for r in rows if r.get("assigned_place")]


def fill_pending_volunteers(date_str):
    assigned_rows = load_assigned_places_for_date(date_str)
    pending_rows = load_pending_signups_for_date(date_str)

    if not pending_rows:
        return "✅ 没有待安排报名"

    place_count = {
        "佛堂卫生": 0,
        "二楼卫生": 0,
        "楼梯卫生": 0,
        "设师父供台": 0,
        "绿观音堂": 0,
        "绿活动中心": 0,
        "橙观音堂": 0,
        "橙活动中心": 0,
        "黄观音堂": 0,
        "黄活动中心": 0,
    }

    for r in assigned_rows:
        shift_label = r.get("shift_label") or ""
        place = r.get("assigned_place") or ""

        key = f"{shift_label}{place}" if shift_label else place

        if key in place_count:
            place_count[key] += 1
        elif place in place_count:
            place_count[place] += 1

    with get_conn() as conn:
        with conn.cursor() as cur:

            for r in pending_rows:
                signup_id = r["id"]
                volunteer_id = r.get("volunteer_id")
                name = r["name"]
                role = r["role"]
                start_time = r.get("start_time")
                end_time = r.get("end_time")

                assigned_place = None
                shift_label = None

                if role == "卫生":
                    cleaning_places = ["佛堂卫生", "二楼卫生", "楼梯卫生"]

                    recent_places = get_recent_cleaning_places(
                        volunteer_id,
                        date_str,
                        limit=5
                    )

                    def cleaning_score(place):
                        score = place_count.get(place, 0) * 10

                        if recent_places:
                            if place == recent_places[0]:
                                score += 100

                            if place in recent_places[:3]:
                                score += 30

                            if place in recent_places:
                                score += 10

                        return score

                    assigned_place = min(
                        cleaning_places,
                        key=cleaning_score
                    )

                elif role == "供台":
                    assigned_place = "设师父供台"

                elif role == "值班":
                    duty_places = [
                        ("绿班", "观音堂", "绿观音堂"),
                        ("绿班", "活动中心", "绿活动中心"),
                        ("橙班", "观音堂", "橙观音堂"),
                        ("橙班", "活动中心", "橙活动中心"),
                        ("黄班", "观音堂", "黄观音堂"),
                        ("黄班", "活动中心", "黄活动中心"),
                    ]

                    shift_label, assigned_place, key = min(
                        duty_places,
                        key=lambda x: place_count[x[2]]
                    )

                else:
                    assigned_place = role

                key = f"{shift_label.replace('班', '')}{assigned_place}" if shift_label else assigned_place
                if key in place_count:
                    place_count[key] += 1

                # 写入真正排班表
                cur.execute("""
                    insert into volunteer_schedule_assignments (
                        signup_id,
                        volunteer_id,
                        name,
                        assignment_date,
                        role,
                        shift_label,
                        assigned_place,
                        start_time,
                        end_time,
                        status,
                        remarks,
                        created_at
                    )
                    values (
                        %s, %s, %s, %s, %s, %s, %s, %s, %s, 'assigned', %s, now()
                    )
                """, (
                    signup_id,
                    volunteer_id,
                    name,
                    date_str,
                    role,
                    shift_label,
                    assigned_place,
                    start_time,
                    end_time,
                    "负责人补入待安排"
                ))

                # 报名表只改状态
                cur.execute("""
                    update volunteer_schedule_signups
                    set
                        status = 'assigned',
                        remarks = '负责人补入待安排'
                    where id = %s
                """, (signup_id,))

            conn.commit()

    return f"✅ 已补入 {len(pending_rows)} 位待安排义工"


from datetime import date, datetime, timedelta
from psycopg2.extras import RealDictCursor


@schedule_bp.route("/schedule/reports")
def schedule_reports():

    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    ym = request.args.get("ym", date.today().strftime("%Y-%m"))

    start_date = f"{ym}-01"

    y, m = map(int, ym.split("-"))
    if m == 12:
        next_month = f"{y+1}-01-01"
    else:
        next_month = f"{y}-{m+1:02d}-01"

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:

            # 本月签到人数
            cur.execute("""
                select count(distinct volunteer_id) as total_people
                from volunteer_attendance_logs
                where check_in_time >= %s
                and check_in_time < %s
            """, (start_date, next_month))
            total_people = cur.fetchone()["total_people"]

            # 本月签到次数
            cur.execute("""
                select count(*) as total_checkins
                from volunteer_attendance_logs
                where check_in_time >= %s
                and check_in_time < %s
            """, (start_date, next_month))
            total_checkins = cur.fetchone()["total_checkins"]

            # 义工排行榜
            cur.execute("""
                select
                    name,
                    count(*) as checkin_count
                from volunteer_attendance_logs
                where check_in_time >= %s
                and check_in_time < %s
                group by name
                order by checkin_count desc, name
                limit 30
            """, (start_date, next_month))
            ranking = cur.fetchall()

            # 岗位统计
            cur.execute("""
                select
                    assigned_place,
                    count(*) as total
                from volunteer_schedule_assignments
                where assignment_date >= %s
                and assignment_date < %s
                and coalesce(status, '') <> 'cancelled'
                group by assigned_place
                order by total desc
            """, (start_date, next_month))
            place_stats = cur.fetchall()

            # 已排班未签到
            cur.execute("""
                select
                    a.assignment_date,
                    a.name,
                    a.role,
                    a.assigned_place
                from volunteer_schedule_assignments a
                left join volunteer_attendance_logs l
                    on l.volunteer_id = a.volunteer_id
                    and date(l.check_in_time) = a.assignment_date
                where a.assignment_date >= %s
                and a.assignment_date < %s
                and coalesce(a.status, '') <> 'cancelled'
                and l.id is null
                order by a.assignment_date, a.name
            """, (start_date, next_month))
            absent_rows = cur.fetchall()

    return render_template_string("""
    <h1>📊 排班报表中心</h1>

    <form method="get">
        <label>月份：</label>
        <input type="month" name="ym" value="{{ ym }}">
        <button type="submit">查看</button>
    </form>

    <hr>

    <h2>📌 本月概况</h2>
    <p>签到义工人数：<b>{{ total_people }}</b></p>
    <p>签到总次数：<b>{{ total_checkins }}</b></p>

    <hr>

    <h2>🏆 义工排行榜</h2>
    <table border="1" cellpadding="8">
        <tr>
            <th>姓名</th>
            <th>签到次数</th>
        </tr>
        {% for r in ranking %}
        <tr>
            <td>{{ r.name }}</td>
            <td>{{ r.checkin_count }}</td>
        </tr>
        {% endfor %}
    </table>

    <hr>

    <h2>🛕 岗位统计</h2>
    <table border="1" cellpadding="8">
        <tr>
            <th>岗位</th>
            <th>次数</th>
        </tr>
        {% for r in place_stats %}
        <tr>
            <td>{{ r.assigned_place or "未安排" }}</td>
            <td>{{ r.total }}</td>
        </tr>
        {% endfor %}
    </table>

    <hr>

    <h2>⚠️ 已排班未签到</h2>
    <table border="1" cellpadding="8">
        <tr>
            <th>日期</th>
            <th>姓名</th>
            <th>角色</th>
            <th>岗位</th>
        </tr>
        {% for r in absent_rows %}
        <tr>
            <td>{{ r.assignment_date }}</td>
            <td>{{ r.name }}</td>
            <td>{{ r.role }}</td>
            <td>{{ r.assigned_place }}</td>
        </tr>
        {% endfor %}
    </table>

    <br>
    <a href="/schedule/admin">返回负责人页面</a>
    """,
    ym=ym,
    total_people=total_people,
    total_checkins=total_checkins,
    ranking=ranking,
    place_stats=place_stats,
    absent_rows=absent_rows
    )


@schedule_bp.route("/schedule/save_day_flags", methods=["POST"])
def save_day_flags():
    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    date_str = request.form.get("date", "").strip()

    setup_person_1 = request.form.get("setup_person_1", "").strip()
    setup_person_2 = request.form.get("setup_person_2", "").strip()

    remove_person_1 = request.form.get("remove_person_1", "").strip()
    remove_person_2 = request.form.get("remove_person_2", "").strip()
    remove_extra_person = request.form.get("remove_extra_person", "").strip()

    extra_buddha_person = request.form.get("extra_buddha_person", "").strip()

    setup_people = "\n".join(
        x for x in [
            setup_person_1,
            setup_person_2
        ] if x
    )

    remove_people = "\n".join(
        x for x in [
            remove_person_1,
            remove_person_2,
            remove_extra_person
        ] if x
    )

    # 兼容旧字段：有填人员 = 有安排
    need_setup = bool(setup_people)
    need_remove = bool(remove_people)

    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                insert into schedule_day_flags (
                    flag_date,
                    need_setup_master_table,
                    need_remove_master_table,
                    setup_people,
                    remove_people,
                    extra_buddha_person,
                    updated_at
                )
                values (%s, %s, %s, %s, %s, %s, now())
                on conflict (flag_date)
                do update set
                    need_setup_master_table = excluded.need_setup_master_table,
                    need_remove_master_table = excluded.need_remove_master_table,
                    setup_people = excluded.setup_people,
                    remove_people = excluded.remove_people,
                    extra_buddha_person = excluded.extra_buddha_person,
                    updated_at = now()
            """, (
                date_str,
                need_setup,
                need_remove,
                setup_people,
                remove_people,
                extra_buddha_person
            ))

            conn.commit()

    return redirect(url_for(
        "schedule.schedule_admin",
        mode="day",
        override_date=date_str
    ))


@schedule_bp.route("/schedule/fill_pending", methods=["POST"])
def schedule_fill_pending():
    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    date_str = request.form.get("date", "").strip()

    if not date_str:
        return "❌ 没有日期<br><a href='/schedule/admin'>返回</a>"

    msg = fill_pending_volunteers(date_str)

    return f"""
    <h1>{msg}</h1>
    <a href="/schedule?mode=day&override_date={date_str}">返回当天值班安排</a><br>
    <a href="/schedule/signups?date={date_str}">查看报名管理</a>
    """


@schedule_bp.route("/schedule/parse_raw", methods=["POST"])
def schedule_parse_raw():

    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    single_date = request.form.get("single_date", "").strip()
    raw_signup = request.form.get("raw_signup", "").strip()

    if not single_date:
        return "❌ 没有日期"

    if not raw_signup:
        return "❌ 没有报名内容"

    try:
        signup_date = datetime.strptime(
            single_date,
            "%Y-%m-%d"
        ).date()
    except:
        return "❌ 日期格式错误"

    lines = raw_signup.splitlines()

    added = 0
    skipped = 0

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:

            for line in lines:

                line = line.strip()

                if not line:
                    continue

                parts = line.split()

                if len(parts) < 2:
                    skipped += 1
                    continue

                name = parts[0]

                time_text = parts[-1]

                time_text = (
                    time_text
                    .replace("～", "-")
                    .replace("~", "-")
                    .replace("—", "-")
                )

                start_time = "3:00pm"
                end_time = "6:00pm"

                if "-" in time_text:

                    try:
                        s, e = time_text.split("-", 1)

                        start_time = s.strip()

                        if (
                            "am" not in e.lower()
                            and
                            "pm" not in e.lower()
                        ):
                            if "pm" in s.lower():
                                e += "pm"
                            elif "am" in s.lower():
                                e += "am"

                        end_time = e.strip()

                    except:
                        pass

                matches = find_volunteer_by_keyword(name)

                if not matches:
                    skipped += 1
                    continue

                if len(matches) > 1:
                    skipped += 1
                    continue

                vol = matches[0]

                volunteer_id = str(vol["id"])
                real_name = str(vol["name"])

                cur.execute("""
                    select id
                    from volunteer_schedule_signups
                    where volunteer_id = %s
                    and signup_date = %s
                    and role = '值班'
                    and coalesce(status,'pending') <> 'cancelled'
                    limit 1
                """, (
                    volunteer_id,
                    signup_date
                ))

                exists = cur.fetchone()

                if exists:
                    skipped += 1
                    continue

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
                        remarks
                    )
                    values
                    (
                        %s,
                        %s,
                        %s,
                        '值班',
                        %s,
                        %s,
                        'pending',
                        'WhatsApp解析'
                    )
                """, (
                    volunteer_id,
                    real_name,
                    signup_date,
                    start_time,
                    end_time
                ))

                added += 1

        conn.commit()

    return f"""
    <h1>✅ WhatsApp 报名导入完成</h1>

    <p>日期：{single_date}</p>
    <p>成功加入：{added} 位</p>
    <p>跳过：{skipped} 位</p>

    <br>

    <a href="/schedule/admin?mode=day&override_date={single_date}">
        返回当天安排
    </a>
    """


@schedule_bp.route("/schedule/publish", methods=["POST"])
def publish_schedule():

    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    target_date = request.form["date"]

    publish_schedule_for_date(target_date)

    return redirect(
        url_for(
            "schedule.schedule_admin",
            mode="day",
            override_date=target_date,
            published="1"
        )
    )


@schedule_bp.route("/schedule/override", methods=["POST"])
def schedule_override():
    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    date = request.form.get("date", "").strip()

    original_names = request.form.getlist("original_name")
    replacement_names = request.form.getlist("replacement_name")

    final_names = []

    for original, replacement in zip(original_names, replacement_names):
        original = str(original).strip()
        replacement = str(replacement).strip()

        if replacement == "__REMOVE__":
            continue

        if replacement:
            final_names.append(replacement)
        elif original:
            final_names.append(original)

    save_buddha_override(date, final_names)

    return redirect(url_for("schedule.schedule_admin", mode="day", override_date=date))


@schedule_bp.route("/schedule/signup_edit/<int:signup_id>", methods=["GET"])
def schedule_signup_edit(signup_id):
    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                select *
                from volunteer_schedule_signups
                where id = %s
            """, (signup_id,))
            row = cur.fetchone()

    if not row:
        return "❌ 找不到报名记录<br><a href='/schedule/signups'>返回</a>"

    return render_template_string("""
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>修改报名</title>
<style>
body { font-family:"Microsoft YaHei", Arial; background:#f5f5f5; padding:20px; }
.box { background:white; max-width:700px; margin:auto; padding:25px; border-radius:15px; }
input, select, button { font-size:22px; padding:8px; margin:8px 0; }
label { display:block; font-size:22px; margin-top:12px; }
</style>
</head>
<body>
<div class="box">

<h1>✏️ 修改报名</h1>

<p><b>{{ row.name }}</b>（{{ row.volunteer_id }}）</p>

<form method="post" action="/schedule/signup_edit/{{ row.id }}">

    <label>日期</label>
    <input type="date" name="signup_date" value="{{ row.signup_date }}" required>

    <label>岗位</label>
    <select name="role" required>
        <option value="值班">值班</option>
        <option value="卫生">卫生</option>
        <option value="设师父供台">设师父供台</option>
        <option value="收师父供台">收师父供台</option>
    </select>

    <label>开始时间</label>
    <select name="start_time">
        {% for t in times %}
        <option value="{{ t }}" {% if row.start_time == t %}selected{% endif %}>{{ t }}</option>
        {% endfor %}
    </select>

    <label>结束时间</label>
    <select name="end_time">
        {% for t in times %}
        <option value="{{ t }}" {% if row.end_time == t %}selected{% endif %}>{{ t }}</option>
        {% endfor %}
    </select>

    <br><br>

    <button type="submit">💾 保存修改</button>
</form>

<br>
<a href="/schedule/signups">⬅ 返回报名管理</a>

</div>
</body>
</html>
""", row=row, roles=ROLE_OPTIONS, times=TIME_OPTIONS)


@schedule_bp.route("/schedule/generate_shortage_notice", methods=["POST"])
def generate_shortage_notice():

    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    date_str = request.form.get("date")

    msg = build_shortage_notice_from_assignments(date_str)

    return render_template_string("""
    <h1>📢 缺人工通知</h1>

    <textarea
        style="width:100%;height:400px;font-size:22px;"
    >{{ msg }}</textarea>

    <br><br>

    <button onclick="navigator.clipboard.writeText(document.querySelector('textarea').value)">
        📋 复制通知
    </button>

    <br><br>

    <a href="/schedule?mode=day&override_date={{ date_str }}">
        返回
    </a>
    """, msg=msg, date_str=date_str)


@schedule_bp.route("/schedule/view_assigned", methods=["POST"])
def schedule_view_assigned():

    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    date_str = request.form.get("date", "").strip()

    rows = load_assigned_places_for_date(date_str)

    cleaning_groups = {
        "佛堂卫生": [],
        "二楼卫生": [],
        "楼梯卫生": [],
    }

    offering_group = []

    duty_groups = {
        "绿": {
            "观音堂": [],
            "活动中心": [],
        },
        "橙": {
            "观音堂": [],
            "活动中心": [],
        },
        "黄": {
            "观音堂": [],
            "活动中心": [],
        },
    }

    other_groups = {}

    for r in rows:
        item = {
            "id": r["id"],
            "name": r["name"],
            "role": r.get("role"),
            "shift_label": r.get("shift_label"),
            "assigned_place": r.get("assigned_place"),
            "start_time": r.get("start_time"),
            "end_time": r.get("end_time"),
        }

        place = r.get("assigned_place") or "未安排"
        role = r.get("role")
        shift = r.get("shift_label")

        if role == "卫生" and place in cleaning_groups:
            cleaning_groups[place].append(item)

        elif role == "供台" or place == "设师父供台":
            offering_group.append(item)

        elif role == "值班":
            shift_key = str(shift or "").replace("班", "")

            if shift_key in duty_groups and place in duty_groups[shift_key]:
                duty_groups[shift_key][place].append(item)
            else:
                other_groups.setdefault(place, []).append(item)

        else:
            other_groups.setdefault(place, []).append(item)

    return render_template_string("""
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>排班结果</title>
<style>
body {
    font-family: "Microsoft YaHei", Arial;
    background:#f5f5f5;
    padding:20px;
}
.box {
    background:white;
    max-width:900px;
    margin:auto;
    padding:25px;
    border-radius:15px;
}
h1, h2, h3 {
    margin-top:18px;
}
.section {
    border:1px solid #ccc;
    padding:15px;
    margin:15px 0;
    border-radius:12px;
    background:#fafafa;
}
.person {
    font-size:22px;
    padding:8px 0;
}
button {
    font-size:22px;
    padding:10px;
    margin:6px 0;
}
a {
    font-size:20px;
}
</style>
</head>
<body>
<div class="box">

<h1>📅 {{ date_str }} 排班结果</h1>

<a href="/schedule?mode=day&override_date={{ date_str }}">⬅ 返回当天值班安排</a>

<hr>

<h2>🧹 卫生</h2>

{% for place, names in cleaning_groups.items() %}
    <div class="section">
        <h3>{{ place }}</h3>

        {% if names %}
            {% for item in names %}
                <div class="person">
                    {{ item.name }}
                    <a href="/schedule/edit_assigned/{{ item.id }}">✏️调整</a>
                </div>
            {% endfor %}
        {% else %}
            <p>暂无安排</p>
        {% endif %}
    </div>
{% endfor %}

<h2>🙏 供台</h2>

<div class="section">
{% if offering_group %}
    {% for item in offering_group %}
        <div class="person">
            {{ item.name }}
            <a href="/schedule/edit_assigned/{{ item.id }}">✏️调整</a>
        </div>
    {% endfor %}
{% else %}
    <p>暂无安排</p>
{% endif %}
</div>

<h2>🏠 值班</h2>

{% for shift, places in duty_groups.items() %}

    {% if shift == "绿" %}
        <h3>🟢 绿班</h3>
    {% elif shift == "橙" %}
        <h3>🟠 橙班</h3>
    {% elif shift == "黄" %}
        <h3>🟡 黄班</h3>
    {% else %}
        <h3>{{ shift }}</h3>
    {% endif %}

    {% for place, names in places.items() %}
        <div class="section">
            <h3>{{ place }}</h3>

            {% if names %}
                {% for item in names %}
                    <div class="person">
                        {{ item.name }}
                        {% if item.start_time and item.end_time %}
                            （{{ item.start_time }}~{{ item.end_time }}）
                        {% endif %}
                        <a href="/schedule/edit_assigned/{{ item.id }}">✏️调整</a>
                    </div>
                {% endfor %}
            {% else %}
                <p>暂无安排</p>
            {% endif %}
        </div>
    {% endfor %}

{% endfor %}

{% if other_groups %}
    <h2>其他</h2>

    {% for place, names in other_groups.items() %}
        <div class="section">
            <h3>{{ place }}</h3>
            {% for item in names %}
                <div class="person">
                    {{ item.name }}
                    {% if item.start_time and item.end_time %}
                        （{{ item.start_time }}~{{ item.end_time }}）
                    {% endif %}
                    <a href="/schedule/edit_assigned/{{ item.id }}">✏️调整</a>
                </div>
            {% endfor %}
        </div>
    {% endfor %}
{% endif %}

<hr>

<form method="post" action="/schedule/copy_whatsapp">
    <input type="hidden" name="date" value="{{ date_str }}">
    <button type="submit">📋 生成 WhatsApp 文字</button>
</form>

</div>
</body>
</html>
""",
        date_str=date_str,
        cleaning_groups=cleaning_groups,
        offering_group=offering_group,
        duty_groups=duty_groups,
        other_groups=other_groups,
    )


@schedule_bp.route("/reports")
def public_reports():

    supabase = get_supabase_client()

    files = supabase.storage.from_("reports").list()

    report_files = []

    for f in files:
        filename = f.get("name")

        if not filename or not filename.endswith(".xlsx"):
            continue

        signed = supabase.storage.from_("reports").create_signed_url(
            filename,
            60 * 60
        )

        report_files.append({
            "filename": filename,
            "display_name": display_report_name(filename),
            "url": signed["signedURL"]
        })

    report_files.sort(
        key=lambda x: x["filename"],
        reverse=True
    )

    return render_template_string("""
    <!doctype html>
    <html>
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>文件中心</title>

        <style>
            body {
                font-family: "Microsoft YaHei", Arial;
                background:#f3f6fb;
                padding:20px;
            }

            .box {
                max-width:900px;
                margin:auto;
                background:white;
                padding:25px;
                border-radius:18px;
                box-shadow:0 4px 14px rgba(0,0,0,0.08);
            }

            h1 {
                font-size:34px;
            }

            .file-row {
                display:flex;
                justify-content:space-between;
                align-items:center;
                gap:15px;
                padding:18px;
                border-bottom:1px solid #e5e7eb;
                font-size:24px;
            }

            .download-btn {
                background:#dbeafe;
                padding:12px 18px;
                border-radius:12px;
                text-decoration:none;
                color:#111827;
                font-weight:bold;
                white-space:nowrap;
            }

            .download-btn:hover {
                filter:brightness(0.95);
            }

            @media (max-width:700px) {
                .file-row {
                    flex-direction:column;
                    align-items:flex-start;
                }

                .download-btn {
                    width:100%;
                    text-align:center;
                }
            }
        </style>
    </head>

    <body>
    <div class="box">

        <h1>📚 观音堂资料中心</h1>

        <p style="font-size:20px;color:#666;">
            提供报表、年报及相关资料下载。
        </p>

        {% if report_files %}
            {% for f in report_files %}
            <div class="file-row">
                <div>📄 {{ f.display_name }}</div>
                <a class="download-btn" href="{{ f.url }}">
                    📥 下载
                </a>
            </div>
            {% endfor %}
        {% else %}
            <p style="font-size:22px;">目前还没有上传报表。</p>
        {% endif %}

        <br>
        <a href="/schedule/admin">返回负责人页面</a>

    </div>
    </body>
    </html>
    """,
    report_files=report_files)


@schedule_bp.route("/schedule/signup_edit/<int:signup_id>", methods=["POST"])
def schedule_signup_edit_save(signup_id):
    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    signup_date = request.form.get("signup_date", "").strip()
    role = request.form.get("role", "").strip()
    start_time = request.form.get("start_time", "").strip()
    end_time = request.form.get("end_time", "").strip()

    if role == "值班":
        if not start_time or not end_time:
            return "❌ 值班必须选择开始和结束时间<br><a href='/volunteer/signup'>返回</a>"

        if end_time <= start_time:
            return "❌ 结束时间必须比开始时间迟<br><a href='/volunteer/signup'>返回</a>"
    else:
        start_time = None
        end_time = None

    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                update volunteer_schedule_signups
                set
                    signup_date = %s,
                    role = %s,
                    start_time = %s,
                    end_time = %s,
                    status = 'pending',
                    assigned_place = null,
                    remarks = '负责人修改报名'
                where id = %s
            """, (
                signup_date,
                role,
                start_time,
                end_time,
                signup_id
            ))

            conn.commit()

    return redirect("/schedule/signups")


@schedule_bp.route("/schedule/signup_place/<int:signup_id>", methods=["GET"])
def schedule_signup_place(signup_id):
    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:

            cur.execute("""
                select *
                from volunteer_schedule_signups
                where id = %s
            """, (signup_id,))
            signup = cur.fetchone()

            if not signup:
                return "❌ 找不到报名记录<br><a href='/schedule/signups'>返回</a>"

            cur.execute("""
                select id
                from volunteer_schedule_assignments
                where signup_id = %s
                and coalesce(status, 'assigned') <> 'cancelled'
                order by start_time, id
            """, (signup_id,))
            assignments = cur.fetchall()

    if len(assignments) == 1:
        assignment_id = assignments[0]["id"]
        return redirect(f"/schedule/edit_assigned/{assignment_id}")

    if len(assignments) > 1:
        return f"""
        <h1>⚠️ 这位义工有多段安排</h1>
        <p>{signup['name']} 在 {signup['signup_date']} 有多段排班。</p>
        <p>请到当天排班结果页面逐段调整。</p>

        <form method="post" action="/schedule/view_assigned">
            <input type="hidden" name="date" value="{signup['signup_date']}">
            <button type="submit">查看当天排班结果</button>
        </form>

        <br>
        <a href="/schedule/signups">返回报名管理</a>
        """

    return f"""
    <h1>❌ 还没有真正排班记录</h1>
    <p>{signup['name']} 目前在 assignments 里没有安排。</p>
    <p>请先自动排班，或到排班结果页面补入。</p>
    <a href="/schedule/signups">返回报名管理</a>
    """


@schedule_bp.route("/schedule/copy_month_prebook", methods=["POST"])
def copy_month_prebook():
    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    year = int(request.form.get("year"))
    month = int(request.form.get("month"))

    msg = build_month_prebook_message(year, month)

    return render_template_string("""
    <h1>📋 {{ year }}年{{ month }}月预报名名单</h1>

    <textarea
        id="msg"
        style="width:100%;height:600px;font-size:22px;"
    >{{ msg }}</textarea>

    <br><br>

    <button
        onclick="navigator.clipboard.writeText(document.getElementById('msg').value)"
        style="font-size:24px;padding:15px 25px;"
    >
        📋 复制
    </button>

    <br><br>

    <a href="/schedule?mode=prebook">返回月预报名</a>
    """, msg=msg, year=year, month=month)


@schedule_bp.route("/schedule/attendance_status")
def schedule_attendance_status():
    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    date_str = request.args.get("date", date.today().strftime("%Y-%m-%d"))

    def group_by_person(rows, source="assignment"):
        grouped = {}

        for r in rows:
            key = r.get("volunteer_id") or r.get("name")

            if not key:
                continue

            if key not in grouped:
                grouped[key] = {
                    "name": r.get("name"),
                    "volunteer_id": r.get("volunteer_id"),
                    "roles": [],
                    "places": [],
                }

            role = r.get("role") or r.get("actual_role") or ""
            if role and role not in grouped[key]["roles"]:
                grouped[key]["roles"].append(role)

            if source == "log":
                place = r.get("actual_place") or ""
                shift = ""
            else:
                place = r.get("assigned_place") or ""
                shift = r.get("shift_label") or ""

            place_text = f"{shift} {place}".strip()

            if place_text and place_text not in grouped[key]["places"]:
                grouped[key]["places"].append(place_text)

        result = []

        for g in grouped.values():
            result.append({
                "name": g["name"],
                "volunteer_id": g["volunteer_id"],
                "role_text": " / ".join(g["roles"]),
                "place_text": " / ".join(g["places"]),
            })

        return result

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:

            cur.execute("""
                select *
                from volunteer_schedule_assignments
                where assignment_date = %s
                and coalesce(status, 'assigned') <> 'cancelled'
                order by role, shift_label, assigned_place, start_time, name
            """, (date_str,))
            assignments = cur.fetchall()

            cur.execute("""
                select *
                from volunteer_attendance_logs
                where attendance_date = %s
                order by check_in_time, name
            """, (date_str,))
            logs = cur.fetchall()

    signed_assignment_ids = {
        r["assignment_id"]
        for r in logs
        if r.get("assignment_id")
    }

    checked_in_raw = []
    not_checked_in_raw = []
    walk_ins_raw = []
    not_checked_out_raw = []

    for a in assignments:
        if a["id"] in signed_assignment_ids:
            checked_in_raw.append(a)
        else:
            not_checked_in_raw.append(a)

    for l in logs:
        if l.get("walk_in"):
            walk_ins_raw.append(l)

        if not l.get("check_out_time"):
            not_checked_out_raw.append(l)

    checked_in = group_by_person(checked_in_raw, source="assignment")
    not_checked_in = group_by_person(not_checked_in_raw, source="assignment")
    walk_ins = group_by_person(walk_ins_raw, source="log")
    not_checked_out = group_by_person(not_checked_out_raw, source="log")

    return render_template_string("""
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>签到情况</title>
<style>
body {
    font-family:"Microsoft YaHei";
    background:#f5f5f5;
    padding:20px;
}
.box {
    background:white;
    max-width:1000px;
    margin:auto;
    padding:25px;
    border-radius:15px;
}
.section {
    border:1px solid #ddd;
    padding:15px;
    margin:15px 0;
    border-radius:10px;
}
li {
    font-size:22px;
    margin:8px 0;
}
input, button {
    font-size:22px;
    padding:8px;
}
.summary {
    display:flex;
    flex-wrap:wrap;
    gap:12px;
    margin:18px 0;
}
.card {
    flex:1;
    min-width:180px;
    padding:15px;
    border-radius:12px;
    font-size:24px;
    font-weight:bold;
    text-align:center;
}
.green { background:#e8fff0; }
.red { background:#ffe8e8; }
.blue { background:#e8f2ff; }
.yellow { background:#fff8d6; }
</style>
</head>

<body>
<div class="box">

<h1>📋 {{ date_str }} 签到情况</h1>

<form method="get">
    <input type="date" name="date" value="{{ date_str }}">
    <button type="submit">查询</button>
</form>

<a href="/schedule/admin">⬅ 返回负责人首页</a>

<div class="summary">
    <div class="card green">✅ 已签到<br>{{ checked_in|length }}</div>
    <div class="card red">❌ 未签到<br>{{ not_checked_in|length }}</div>
    <div class="card blue">🚶 临时报到<br>{{ walk_ins|length }}</div>
    <div class="card yellow">⏳ 未签退<br>{{ not_checked_out|length }}</div>
</div>

<div class="section">
<h2>✅ 已签到：{{ checked_in|length }}</h2>
<ul>
{% for r in checked_in %}
<li>{{ r.name }}｜{{ r.role_text }}｜{{ r.place_text }}</li>
{% endfor %}
</ul>
</div>

<div class="section">
<h2>❌ 已排班未签到：{{ not_checked_in|length }}</h2>
<ul>
{% for r in not_checked_in %}
<li>{{ r.name }}｜{{ r.role_text }}｜{{ r.place_text }}</li>
{% endfor %}
</ul>
</div>

<div class="section">
<h2>🚶 临时报到：{{ walk_ins|length }}</h2>
<ul>
{% for r in walk_ins %}
<li>{{ r.name }}｜{{ r.role_text }}｜{{ r.place_text }}</li>
{% endfor %}
</ul>
</div>

<div class="section">
<h2>⏳ 未签退：{{ not_checked_out|length }}</h2>
<ul>
{% for r in not_checked_out %}
<li>{{ r.name }}｜{{ r.role_text }}｜{{ r.place_text }}</li>
{% endfor %}
</ul>
</div>

</div>
</body>
</html>
""",
        date_str=date_str,
        checked_in=checked_in,
        not_checked_in=not_checked_in,
        walk_ins=walk_ins,
        not_checked_out=not_checked_out,
    )


@schedule_bp.route("/schedule/settings", methods=["GET", "POST"])
def schedule_settings():

    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    if request.method == "POST":
        keys = [
            "default_day_switch_time",
            "target_orange_guanyintang",
            "target_orange_activity",
            "target_yellow_guanyintang",
            "target_yellow_activity",
            "target_cleaning",
            "supply_alert_days",
        ]

        for key in keys:
            value = request.form.get(key, "").strip()
            save_schedule_setting(key, value)

        return redirect(url_for("schedule.schedule_settings", saved="1"))

    settings = get_schedule_settings()
    saved = request.args.get("saved") == "1"

    return render_template_string("""
    <h1>⚙️ 排班设置</h1>

    {% if saved %}
        <p style="color:green; font-weight:bold;">✅ 设置已保存</p>
    {% endif %}

    <form method="post" style="font-size:18px; max-width:500px;">

        <h2>📅 默认日期切换</h2>

        <label>几点后默认看明天？</label><br>
        <input name="default_day_switch_time"
               value="{{ settings.default_day_switch_time }}"
               placeholder="18:00"
               style="font-size:20px; padding:8px; width:150px;">
        <p style="color:#666;">例：18:00 = 晚上6点后负责人页面默认切到明天</p>

        <hr>

        <h2>👥 缺义工目标人数</h2>

        <label>橙班 - 观音堂</label><br>
        <input name="target_orange_guanyintang" value="{{ settings.target_orange_guanyintang }}" style="font-size:20px; padding:8px;"><br><br>

        <label>橙班 - 活动中心</label><br>
        <input name="target_orange_activity" value="{{ settings.target_orange_activity }}" style="font-size:20px; padding:8px;"><br><br>

        <label>黄班 - 观音堂</label><br>
        <input name="target_yellow_guanyintang" value="{{ settings.target_yellow_guanyintang }}" style="font-size:20px; padding:8px;"><br><br>

        <label>黄班 - 活动中心</label><br>
        <input name="target_yellow_activity" value="{{ settings.target_yellow_activity }}" style="font-size:20px; padding:8px;"><br><br>

        <label>卫生人数</label><br>
        <input name="target_cleaning" value="{{ settings.target_cleaning }}" style="font-size:20px; padding:8px;"><br><br>

        <hr>

        <h2>🪔 供台提醒</h2>

        <label>提醒未来几天的大日子？</label><br>
        <input name="supply_alert_days" value="{{ settings.supply_alert_days }}" style="font-size:20px; padding:8px;"><br><br>

        <button type="submit" style="font-size:22px; padding:12px 20px;">
            💾 保存设置
        </button>

    </form>

    <br>
    <a href="{{ url_for('schedule.schedule_admin') }}">⬅ 返回负责人首页</a>
    """, settings=settings, saved=saved)


@schedule_bp.route("/schedule/signups")
def schedule_signups_manage():
    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    status = request.args.get("status", "all")
    date_filter = request.args.get("date", "").strip()
    role_filter = request.args.get("role", "all")
    show_history = request.args.get("show_history") == "1"

    sql = """
        select
            id,
            volunteer_id,
            name,
            signup_date,
            role,
            start_time,
            end_time,
            coalesce(status, 'pending') as status,
            assigned_place,
            remarks
        from volunteer_schedule_signups
        where 1=1
    """

    if not show_history:
        sql += """
            and signup_date >= current_date
            and coalesce(status,'pending') <> 'cancelled'
        """

    params = []

    if status != "all":
        sql += " and coalesce(status, 'pending') = %s"
        params.append(status)

    if date_filter:
        sql += " and signup_date = %s"
        params.append(date_filter)

    if role_filter != "all":
        sql += " and role = %s"
        params.append(role_filter)

    sql += " order by signup_date, role, start_time, name"

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute(sql, params)
            rows = cur.fetchall()

            signup_ids = [r["id"] for r in rows]
            assignment_map = {}

            if signup_ids:
                cur.execute("""
                    select
                        signup_id,
                        shift_label,
                        assigned_place,
                        start_time,
                        end_time
                    from volunteer_schedule_assignments
                    where signup_id = any(%s)
                    and coalesce(status, 'assigned') <> 'cancelled'
                    order by start_time, shift_label, assigned_place
                """, (signup_ids,))

                assignment_rows = cur.fetchall()

                for a in assignment_rows:
                    sid = a["signup_id"]

                    shift = a.get("shift_label") or ""
                    place = a.get("assigned_place") or ""
                    start = a.get("start_time") or ""
                    end = a.get("end_time") or ""

                    text = ""

                    if shift:
                        text += shift

                    if place:
                        if text:
                            text += " "
                        text += place

                    if start and end:
                        text += f"<br><small>{start} ~ {end}</small>"

                    assignment_map.setdefault(sid, []).append(text)

            for r in rows:
                r["final_assignment"] = assignment_map.get(r["id"], [])

    summary = {
        "total": len(rows),
        "pending": sum(1 for r in rows if r["status"] == "pending"),
        "assigned": sum(1 for r in rows if r["status"] == "assigned"),
        "cancelled": sum(1 for r in rows if r["status"] == "cancelled"),
    }

    today = date.today().strftime("%Y-%m-%d")
    tomorrow = (date.today() + timedelta(days=1)).strftime("%Y-%m-%d")

    role_summary = {}

    for r in rows:
        role = r["role"] or "未填写"
        role_summary[role] = role_summary.get(role, 0) + 1

    return render_template_string("""
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>报名管理</title>

<style>
body {
    font-family: "Microsoft YaHei", Arial;
    background:#f5f5f5;
    padding:20px;
}

.box {
    background:white;
    max-width:1200px;
    margin:auto;
    padding:25px;
    border-radius:15px;
}

button, input, select {
    font-size:20px;
    padding:8px;
    margin:5px;
}

table {
    width:100%;
    border-collapse:collapse;
    font-size:20px;
}

th, td {
    border:1px solid #999;
    padding:8px;
    text-align:center;
}

.table-wrap {
    overflow-x:auto;
}

.pending {
    background:#fff8d6;
}

.assigned {
    background:#e8fff0;
}

.cancelled {
    background:#eee;
    color:#777;
}

.summary-box {
    background:#eef7ff;
    padding:12px;
    border-radius:10px;
    font-size:22px;
    margin:12px 0;
}

.role-summary-box {
    background:#fff8d6;
    padding:12px;
    border-radius:10px;
    font-size:22px;
    margin:12px 0;
}

.filter-box {
    background:#fafafa;
    border:1px solid #ddd;
    padding:15px;
    border-radius:12px;
    margin:15px 0;
    font-size:20px;
}

.op-form {
    margin:5px 0;
}

.op-btn {
    width:120px;
}

.floating-back {
    position: fixed;
    right: 25px;
    bottom: 25px;
    background: #2196F3;
    color: white;
    padding: 16px 24px;
    border-radius: 50px;
    text-decoration: none;
    font-size: 22px;
    font-weight: bold;
    box-shadow: 0 4px 12px rgba(0,0,0,0.25);
    z-index: 9999;
}

.floating-back:hover {
    background: #0b7dda;
}
</style>
</head>

<body>

<div class="box">

<h1>📋 义工报名管理</h1>

<a href="/schedule/admin">⬅ 返回负责人首页</a>

<hr>

<div class="summary-box">
    当前查询：<b>{{ summary.total }}</b> 条　
    未安排：<b>{{ summary.pending }}</b>　
    已安排：<b>{{ summary.assigned }}</b>　
    已取消：<b>{{ summary.cancelled }}</b>
</div>

<div class="role-summary-box">
    <b>岗位统计：</b>
    {% if role_summary %}
        {% for role, count in role_summary.items() %}
            {{ role }}：<b>{{ count }}</b> 人　
        {% endfor %}
    {% else %}
        没有记录
    {% endif %}
</div>

<div class="filter-box">
<form method="get">

    日期：
    <input type="date" name="date" value="{{ date_filter }}">

    <a href="/schedule/signups?date={{ today }}&status={{ status }}&role={{ role_filter }}">
        <button type="button">今天</button>
    </a>

    <a href="/schedule/signups?date={{ tomorrow }}&status={{ status }}&role={{ role_filter }}">
        <button type="button">明天</button>
    </a>

    <a href="/schedule/signups">
        <button type="button">全部</button>
    </a>

    <br><br>

    状态：
    <select name="status" style="width:160px;">

    <option value="all"
    {% if status=="all" %}selected{% endif %}>
    全部状态
    </option>

    <option value="pending"
    {% if status=="pending" %}selected{% endif %}>
    未安排
    </option>

    <option value="assigned"
    {% if status=="assigned" %}selected{% endif %}>
    已安排
    </option>

    <option value="cancelled"
    {% if status=="cancelled" %}selected{% endif %}>
    已取消
    </option>

    </select>

    岗位：
    <select name="role" style="width:160px;">

    <option value="all"
    {% if role_filter=="all" %}selected{% endif %}>
    全部岗位
    </option>

    {% for role in roles %}
    <option value="{{ role }}"
    {% if role_filter==role %}selected{% endif %}>
    {{ role }}
    </option>
    {% endfor %}

    </select>
        <option value="all" {% if role_filter == "all" %}selected{% endif %}>全部</option>
        {% for role in roles %}
            <option value="{{ role }}" {% if role_filter == role %}selected{% endif %}>
                {{ role }}
            </option>
        {% endfor %}
    </select>

    <label style="font-size:18px;">
        <input
            type="checkbox"
            name="show_history"
            value="1"
            {% if show_history %}checked{% endif %}
        >
        显示过去记录
    </label>

    <button type="submit">
        🔍 查询
    </button>

</form>
</div>

<hr>

<div class="table-wrap">
<table>
<tr>
    <th>日期</th>
    <th>编号</th>
    <th>姓名</th>
    <th>岗位</th>
    <th>时间</th>
    <th>状态</th>
    <th>系统安排</th>
    <th>备注</th>
    <th>操作</th>
</tr>

{% for r in rows %}
<tr class="{{ r.status }}">
    <td>{{ r.signup_date }}</td>
    <td>{{ r.volunteer_id }}</td>
    <td>{{ r.name }}</td>
    <td>{{ r.role }}</td>
    <td>
        {% if r.start_time or r.end_time %}
            {{ r.start_time or "" }} ~ {{ r.end_time or "" }}
        {% else %}
            -
        {% endif %}
    </td>
    <td>
        {% if r.status == "pending" %}
            未安排
        {% elif r.status == "assigned" %}
            已安排
        {% elif r.status == "cancelled" %}
            已取消
        {% else %}
            {{ r.status }}
        {% endif %}
    </td>
    <td>
        {% if r.final_assignment %}
            {% for item in r.final_assignment %}
                <div>{{ item|safe }}</div>
            {% endfor %}
        {% else %}
            -
        {% endif %}
    </td>
    <td>{{ r.remarks or "" }}</td>

    <td>

        <form class="op-form" method="get" action="/schedule/signup_edit/{{ r.id }}">
            <button class="op-btn" type="submit">✏️ 修改</button>
        </form>

        {% if r.status == "assigned" %}
            <form class="op-form" method="get" action="/schedule/signup_place/{{ r.id }}">
                <button class="op-btn" type="submit">🔧 调整</button>
            </form>
        {% endif %}

        {% if r.status == "pending" %}
            <form class="op-form"
                  method="post"
                  action="/schedule/signup_cancel/{{ r.id }}"
                  onsubmit="return confirm('确定取消报名？');">
                <button class="op-btn" type="submit">❌ 取消</button>
            </form>

        {% elif r.status == "cancelled" %}
            <form class="op-form"
                  method="post"
                  action="/schedule/signup_restore/{{ r.id }}"
                  onsubmit="return confirm('确定恢复报名？');">
                <button class="op-btn" type="submit">↩️ 恢复</button>
            </form>

        {% elif r.status == "assigned" %}
            <div style="font-size:18px; color:#666;">已安排</div>

        {% else %}
            -
        {% endif %}

    </td>
</tr>
{% endfor %}
</table>
</div>

{% if not rows %}
<p style="font-size:22px;">没有记录。</p>
{% endif %}

</div>
                                  
<a href="/schedule/admin" class="floating-back">
    ⬅ 返回首页
</a>
                                  
</body>
</html>
""",
        rows=rows,
        status=status,
        date_filter=date_filter,
        role_filter=role_filter,
        roles=ROLE_OPTIONS,
        summary=summary,
        role_summary=role_summary,
        today=today,
        tomorrow=tomorrow
    )


LOGIN_HTML = """
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>负责人排班系统</title>
<style>
body { font-family: "Microsoft YaHei", Arial; background:#f5f5f5; padding:20px; font-size:24px; }
.box { background:white; max-width:1200px; margin:auto; padding:30px; border-radius:15px; text-align:center; }
input, button { font-size:30px; padding:18px; margin:10px; width:90%; border-radius:12px; }
</style>
</head>
<body>
<div class="box">
<h1>📅 负责人排班系统</h1>
<form method="post">
    <input
        type="password"
        name="pin"
        placeholder="请输入负责人PIN"
        inputmode="numeric"
        autocomplete="new-password"
        autocorrect="off"
        autocapitalize="off"
        spellcheck="false"
        readonly
        onfocus="this.removeAttribute('readonly');"
    >
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
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>负责人排班系统</title>

<style>
body {
    font-family: "Microsoft YaHei", Arial, sans-serif;
    background: #f3f6fb;
    margin: 0;
    padding: 18px;
    color: #1f2937;
}

.box {
    background: transparent;
    max-width: 1180px;
    margin: auto;
}

.top-bar {
    background: white;
    padding: 20px 24px;
    border-radius: 18px;
    display:flex;
    justify-content:space-between;
    align-items:center;
    gap:15px;
    flex-wrap:wrap;
    box-shadow: 0 4px 14px rgba(0,0,0,0.06);
}

.top-bar h1 {
    margin: 0;
    font-size: 30px;
}

.card {
    background: white;
    border-radius: 18px;
    padding: 22px;
    margin-top: 18px;
    box-shadow: 0 4px 14px rgba(0,0,0,0.06);
}

.info-box {
    background: #eaf5ff;
    border-left: 6px solid #3b82f6;
    padding: 18px;
    border-radius: 16px;
    font-size: 23px;
}

.warn-box {
    background: #fff4f4;
    border-left: 6px solid #ef4444;
    padding: 18px;
    border-radius: 16px;
    font-size: 22px;
}

.main-menu,
.quick-actions {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 14px;
}

.big-btn,
.quick-btn,
.action-btn {
    width: 100%;
    border: 0;
    border-radius: 16px;
    background: #f1f5f9;
    padding: 18px 16px;
    font-size: 24px;
    font-weight: 700;
    cursor: pointer;
    color: #111827;
    text-align: center;
    box-sizing: border-box;
    text-decoration: none;
    display: block;
}

.big-btn {
    font-size: 28px;
    padding: 24px;
    background: #eef6ff;
}

.quick-btn.primary {
    background: #dbeafe;
}

.quick-btn.warn {
    background: #fee2e2;
}

.quick-btn.whatsapp {
    background: #dcfce7;
}

.big-btn:hover,
.quick-btn:hover,
.action-btn:hover {
    filter: brightness(0.96);
}

.section-title {
    margin: 0 0 16px 0;
    font-size: 25px;
}

.action-row {
    display:flex;
    gap:15px;
    flex-wrap:wrap;
    align-items:center;
}

.date-input,
input,
select,
textarea {
    font-size: 22px;
    padding: 12px;
    border: 1px solid #cbd5e1;
    border-radius: 12px;
    box-sizing: border-box;
}

.date-input {
    font-size: 26px;
}

textarea {
    width:100%;
}

.two-col {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 18px;
}

.sub-card {
    background: #f8fafc;
    border-radius: 16px;
    padding: 18px;
}

.role-box label {
    display: inline-block;
    background: white;
    border: 1px solid #cbd5e1;
    border-radius: 12px;
    padding: 8px 12px;
    margin: 6px;
    font-size: 22px;
}

.special-box {
    margin-top:25px;
    padding:25px;
    background:#f8fafc;
    border:2px solid #dbeafe;
    border-radius:18px;
}

.publish-btn{
    background:#2e7d32;
    color:#fff;
    font-size:30px;
    font-weight:bold;
}

.auto-btn{
    background:#1976d2;
    color:#fff;
}

.view-btn{
    background:#7b1fa2;
    color:#fff;
}

.whatsapp-btn{
    background:#25D366;
    color:#fff;
}

.shortage-btn{
    background:#ef6c00;
    color:#fff;
}

.attendance-btn{
    background:#0097a7;
    color:#fff;
}

.report-btn{
    background:#546e7a;
    color:#fff;
}

.publish-btn{
    background:#2e7d32;
    color:white;
}

.published-btn{
    background:#9e9e9e;
    color:white;
}

.published-btn{
    background:#9e9e9e;
    color:white;
    opacity:0.9;
}

button {
    font-family: inherit;
}

a {
    text-decoration: none;
}

@media (max-width: 800px) {
    .main-menu,
    .quick-actions,
    .two-col {
        grid-template-columns: 1fr;
    }

    .top-bar h1 {
        font-size: 25px;
    }

    .big-btn,
    .quick-btn,
    .action-btn {
        font-size: 22px;
    }
}
</style>

<script>
document.addEventListener("DOMContentLoaded", function () {

    function syncDates() {
        const mainDate = document.getElementById("main_date");

        const rawDate = document.getElementById("raw_single_date");
        const genDate = document.getElementById("generate_day_date");
        const fillDate = document.getElementById("fill_pending_date");
        const viewAssignedDate = document.getElementById("view_assigned_date");
        const overrideDate = document.getElementById("override_date");
        const prebookDate = document.getElementById("prebook_single_date");
        const publishDate = document.getElementById("publish_date");
        const noticeDate = document.getElementById("notice_date");

        if (!mainDate) return;

        if (rawDate) rawDate.value = mainDate.value;
        if (genDate) genDate.value = mainDate.value;
        if (fillDate) fillDate.value = mainDate.value;
        if (viewAssignedDate) viewAssignedDate.value = mainDate.value;
        if (overrideDate) overrideDate.value = mainDate.value;
        if (prebookDate) prebookDate.value = mainDate.value;
        if (publishDate) publishDate.value = mainDate.value;
        if (noticeDate) noticeDate.value = mainDate.value;
    }

    syncDates();

    const mainDate = document.getElementById("main_date");
    if (mainDate) {
        mainDate.addEventListener("change", syncDates);
    }
});
</script>

<script>
function showSpecial(id) {
    const boxes = document.querySelectorAll(".special-box");

    boxes.forEach(function(box) {
        box.style.display = "none";
    });

    const target = document.getElementById(id);

    if (target) {
        target.style.display = "block";
    }
}
</script>

</head>

<body>
<div class="box">

<div class="top-bar">
    <h1>📅 负责人排班系统</h1>

    <a href="/schedule/logout">
        <button type="button">🚪 退出登录</button>
    </a>
</div>

<div class="card">
    <div class="main-menu">
        <a href="/schedule/signups">
            <button type="button" class="big-btn">📋 报名管理</button>
        </a>

        <a href="/schedule?mode=day">
            <button type="button" class="big-btn">📅 当天安排</button>
        </a>

        <a href="/volunteer">
            <button type="button" class="big-btn">👥 义工报名</button>
        </a>

        <a href="/schedule/settings">
            <button type="button" class="big-btn">⚙️ 排班设置</button>
        </a>

        <a href="/reports">
            <button type="button" class="big-btn">📚 资料中心</button>
        </a>
    </div>
</div>

<div class="warn-box">
    <h3 style="margin-top:0;">🔴 未安排提醒</h3>
    今日：<b>{{ pending_counts.today }}</b> 人　
    明日：<b>{{ pending_counts.tomorrow }}</b> 人　
    本月：<b>{{ pending_counts.month }}</b> 人
</div>

<div class="alert-box" style="background:#fff8e6;border-left:6px solid #f0b429;padding:18px;margin:18px 0;border-radius:12px;">
    <h2>🌸 供台报名提醒</h2>

    {% for item in supply_alerts %}
        <div style="margin-bottom:14px;">
            <b>
                {{ item.date.strftime("%m/%d") }}
                {% if item.type == "lunar_1_15" %}
                    初一 / 十五
                {% elif item.type == "buddhist_festival" %}
                    佛诞日
                {% endif %}
            </b><br>

            {% if item.names %}
                已报名：{{ "、".join(item.names) }}
            {% else %}
                <span style="color:#c0392b;">暂时没人报名</span>
            {% endif %}
        </div>
    {% endfor %}
</div>

<hr>

{% if mode == "day" %}

<h2>📅 当天值班安排</h2>

<div class="card">
    <h3>📅 日期选择</h3>

    <div class="info-box">
        <b>🗓 日期资料</b><br>

        阳历：{{ special_day_info.solar_date }}<br>
        农历：{{ special_day_info.lunar_text }}<br>

        {% if special_day_info.is_special %}
            <span style="color:red;font-weight:bold;">
                🔴 {{ special_day_info.special_names | join("、") }}
            </span><br>

            模板类型：{{ special_day_info.template_text }}<br>

            {% if special_day_info.setup_shifu %}
                🛕 需要设师父供台<br>
            {% endif %}

            {% if special_day_info.remove_next_day %}
                📦 明日需要收供台<br>
            {% endif %}
        {% else %}
            🟢 平时日
        {% endif %}

        {% if remove_info.need_remove_today_after_12 %}
            <br>⚠️ 今日中午后需要收昨日供台
        {% endif %}
    </div>

    <div class="action-row">
        <input
            type="date"
            id="main_date"
            class="date-input"
            value="{{ override_date }}"
            required
        >

        <a href="/schedule?mode=day&override_date={{ override_date }}">
            <button type="button" class="action-btn">📅 当前日期：{{ override_date }}</button>
        </a>
    </div>
</div>

<div class="info-box">
    <b>📊 当天概况</b><br>
    已报名：<b>{{ day_summary.total }}</b> 人　
    待安排：<b>{{ day_summary.pending }}</b> 人　
    已安排：<b>{{ day_summary.assigned }}</b> 人　
    已取消：<b>{{ day_summary.cancelled }}</b> 人
</div>

<div class="info-box">
    <b>📢 缺人工提醒</b><br>

    {% if shortage_summary %}

        {% for line in shortage_summary %}
            {{ line }}<br>
        {% endfor %}

    {% else %}

        🟢 目前没有缺人工岗位

    {% endif %}
</div>

<div class="section quick-panel">
    <h3 class="section-title">⚡ 快捷操作</h3>

    <div class="quick-actions">

        <form method="post" action="/schedule/publish" style="grid-column:1 / -1;">
            <input type="hidden" name="date" id="publish_date" value="{{ override_date }}">
            <button
                type="submit"
                class="quick-btn {% if is_published %}published-btn{% else %}publish-btn{% endif %}"
            >
                {% if is_published %}
                    ✅ 已发布（重新发布）
                {% else %}
                    🟢 发布正式值班表
                {% endif %}
            </button>
        </form>

        {% if not is_today_or_past %}
        <form method="post" action="/schedule/generate_day">
            <input type="hidden" name="date" id="generate_day_date" value="{{ override_date }}">
            <button type="submit" class="quick-btn auto-btn">
                ⚡ 自动排班
            </button>
        </form>
        {% else %}
        <form method="post" action="/schedule/fill_pending">
            <input type="hidden" name="date" id="fill_pending_date" value="{{ override_date }}">
            <button type="submit" class="quick-btn auto-btn">
                ➕ 补入待安排
            </button>
        </form>
        {% endif %}

        <form method="post" action="/schedule/view_assigned">
            <input type="hidden" name="date" id="view_assigned_date" value="{{ override_date }}">
            <button type="submit" class="quick-btn view-btn">
                👀 查看排班
            </button>
        </form>

        <form method="post" action="/schedule/copy_whatsapp">
            <input type="hidden" name="date" id="copy_whatsapp_date" value="{{ override_date }}">
            <button type="submit" class="quick-btn whatsapp-btn">
                📋 WhatsApp
            </button>
        </form>

        <form method="post" action="/schedule/generate_shortage_notice">
            <input type="hidden" name="date" id="notice_date" value="{{ override_date }}">
            <button type="submit" class="quick-btn shortage-btn">
                📢 缺义工通知
            </button>
        </form>

        <a href="/schedule/attendance_status" class="quick-btn attendance-btn">
            📋 签到情况
        </a>

        <a href="/reports" class="quick-btn report-btn">
            📚 资料中心
        </a>

    </div>
</div>

<div class="card">
    <h3>👥 报名录入</h3>

    <div class="two-col">

        <div class="sub-card">
            <h4>➕ 负责人代报名</h4>

            <form method="post" action="/schedule/prebook_add">
                <input type="hidden" name="mode" value="day">
                <input type="hidden" name="single_date" id="prebook_single_date" value="{{ override_date }}">

                <p>义工编号 / 姓名</p>
                <input name="vol_id" placeholder="输入义工编号 / 姓名" required>

                <p>岗位</p>
                <div class="role-box">
                {% for role in roles %}
                    <label>
                        <input type="checkbox" name="roles" value="{{ role }}"> {{ role }}
                    </label>
                {% endfor %}
                </div>

                <p>时间（只给值班用）</p>

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

                <br><br>

                <button type="submit" class="action-btn">➕ 代报名</button>
            </form>
        </div>

        <div class="sub-card">
            <h4>📥 WhatsApp 报名解析</h4>

            <form method="post" action="/schedule/parse_raw">
                <input type="hidden" name="single_date" id="raw_single_date" value="{{ override_date }}">

                <textarea
                    name="raw_signup"
                    rows="10"
                    style="font-size:22px;"
                    placeholder="输入姓名 / 时间，例如：&#10;张三 10:00am-2:00pm&#10;李四 8:00am-12:00pm"
                ></textarea>

                <br><br>

                <button type="submit" class="action-btn">📥 解析加入报名</button>
            </form>
        </div>

    </div>
</div>

    <div class="card">
        <h3>🛕 特殊设置</h3>

        <div class="quick-actions">
            <button type="button" class="quick-btn" onclick="showSpecial('master-table')">
                🛕 师父供台
            </button>

            <button type="button" class="quick-btn" onclick="showSpecial('extra-buddha')">
                🌸 初一补人
            </button>

            <button type="button" class="quick-btn" onclick="showSpecial('buddha-leave')">
                🔁 佛台请假
            </button>
        </div>

        <!-- 1. 师父供台 -->
        <div id="master-table" class="special-box" style="display:none;">
            <h3>🛕 师父供台设置</h3>

            {% if supply_signups %}
            <div style="
                background:#fff8e1;
                border:2px solid #f3d36b;
                border-radius:14px;
                padding:15px;
                margin:15px 0;
                font-size:22px;
            ">
                <b>🌸 已报名供台义工：</b><br>
                {% for p in supply_signups %}
                    {{ p.name }}{% if not loop.last %}、{% endif %}
                {% endfor %}
            </div>
            {% endif %}

            <form method="post" action="/schedule/save_day_flags">
                <input type="hidden" name="date" value="{{ override_date }}">

                <div class="info-box">
                    系统会根据初一、十五及佛诞日自动判断是否需要设供台。<br>
                    若前一天应收供台但未安排人员，系统会继续提醒收供台。
                </div>

                <h4>🪷 设师父供台人员</h4>
                <p style="color:#666;font-size:20px;">通常 1 至 2 位佛台组人员即可。</p>

                <select name="setup_person1">
                    <option value="">请选择</option>

                    {% for p in supply_signups %}
                        <option value="{{ p.name }}"
                            {% if loop.index == 1 %}selected{% endif %}>
                            {{ p.name }}
                        </option>
                    {% endfor %}

                    {% for name in volunteer_names %}
                        <option value="{{ name }}">
                            {{ name }}
                        </option>
                    {% endfor %}
                </select>

                <select name="setup_person_2">
                    <option value="">请选择</option>

                    {% for p in supply_signups %}
                        <option value="{{ p.name }}"
                            {% if loop.index == 2 %}selected{% endif %}>
                            {{ p.name }}
                        </option>
                    {% endfor %}

                    {% for n in buddha_names %}
                        <option value="{{ n }}" {% if day_flags.setup_person_2 == n %}selected{% endif %}>
                            {{ n }}
                        </option>
                    {% endfor %}
                </select>

                <hr>

                <h4>📦 收师父供台人员</h4>
                <p style="color:#666;font-size:20px;">
                    通常 1 至 2 位佛台组人员。若没有安排人员，系统视为继续供奉。
                </p>

                <select name="remove_person_1">
                    <option value="">请选择</option>
                    {% for n in buddha_names %}
                    <option value="{{ n }}" {% if day_flags.remove_person_1 == n %}selected{% endif %}>
                        {{ n }}
                    </option>
                    {% endfor %}
                </select>

                <select name="remove_person_2">
                    <option value="">请选择</option>
                    {% for n in buddha_names %}
                    <option value="{{ n }}" {% if day_flags.remove_person_2 == n %}selected{% endif %}>
                        {{ n }}
                    </option>
                    {% endfor %}
                </select>

                <p>普通义工帮忙收供台：</p>
                <input
                    name="remove_extra_person"
                    value="{{ day_flags.remove_extra_person }}"
                    placeholder="例如：张三，可空"
                    style="width:100%;font-size:22px;"
                >

                <br><br>

                <button type="submit" class="action-btn">💾 保存供台人员</button>
            </form>
        </div>

        <!-- 2. 初一补人 -->
        <div id="extra-buddha" class="special-box" style="display:none;">
            <h3>🌸 初一整理佛台补人</h3>

            <form method="post" action="/schedule/save_day_flags">
                <input type="hidden" name="date" value="{{ override_date }}">

                <p style="font-size:20px;color:#666;">
                    仅初一需要。可增加第 3 位整理佛台义工。
                </p>

                <p>额外第 3 位整理佛台义工：</p>

                <select name="extra_buddha_person" style="width:100%;font-size:22px;">
                    <option value="">不需要增加</option>
                    {% for n in buddha_names %}
                    <option value="{{ n }}" {% if day_flags.extra_buddha_person == n %}selected{% endif %}>
                        {{ n }}
                    </option>
                    {% endfor %}
                </select>

                <br><br>

                <button type="submit" class="action-btn">💾 保存初一补人</button>
            </form>
        </div>

        <!-- 3. 佛台请假 -->
        <div id="buddha-leave" class="special-box" style="display:none;">
            <h3>🔁 佛台请假 / 换人</h3>

            <p>日期：{{ override_date }}</p>

            <form method="post" action="/schedule/override">
                <input type="hidden" name="date" id="override_date" value="{{ override_date }}">

                {% if fixed_buddha_today %}
                    {% for old_name in fixed_buddha_today %}
                        <div style="font-size:22px; margin:12px 0;">
                            {{ old_name }}
                            <input type="hidden" name="original_name" value="{{ old_name }}">

                            换成：
                            <select name="replacement_name">
                                <option value="">不换，保留原本</option>
                                <option value="__REMOVE__">请假，不找替补</option>

                                {% for n in buddha_names %}
                                <option value="{{ n }}">{{ n }}</option>
                                {% endfor %}
                            </select>
                        </div>
                    {% endfor %}

                    <button type="submit" class="action-btn">💾 保存佛台请假 / 换人</button>
                {% else %}
                    <p style="font-size:22px;">这一天没有固定佛台。</p>
                {% endif %}
            </form>
        </div>
    </div>
    
</div>

{% else %}

<h2 style="text-align:center;">请选择要做的功能</h2>

{% endif %}

</div>
</body>
</html>
"""


DAY_OUTPUT_HTML = """
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
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
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>月预报名表</title>

<style>
body {
    font-family: "Microsoft YaHei", Arial;
    background:#f5f5f5;
    padding:20px;
}

.box {
    background:white;
    max-width:900px;
    margin:auto;
    padding:25px;
    border-radius:15px;
}

textarea {
    width:100%;
    height:650px;
    font-size:20px;
    padding:15px;
    box-sizing:border-box;
}

a, button {
    font-size:22px;
    padding:10px 18px;
    margin:8px;
    cursor:pointer;
}
</style>

<script>
function copyOutput(btn) {

    const text =
        document.getElementById("output").value;

    navigator.clipboard.writeText(text);

    btn.innerText = "✅ 已复制";

    setTimeout(() => {
        btn.innerText = "📋 一键复制";
    }, 2000);
}
</script>

</head>

<body>

<div class="box">

<h1>📢 月预报名表</h1>

<a href="/schedule?mode=prebook">⬅ 返回月预报名</a>

<button onclick="copyOutput(this)">
📋 一键复制
</button>

<br><br>

<textarea id="output" readonly>{{ output }}</textarea>

</div>

</body>
</html>
"""
