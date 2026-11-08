# schedule_web.py

import os
import calendar
import time
from functools import lru_cache
import psycopg2
import pandas as pd

from opencc import OpenCC
from dotenv import load_dotenv
from db import get_db, get_conn, db_query
from supabase import create_client
from openpyxl import load_workbook
from psycopg2.extras import RealDictCursor
from sqlalchemy import create_engine, text
from utils import apply_branch_prefix
from schedule.blueprint import schedule_bp
from schedule.services.whatsapp_agent_api import register_whatsapp_agent_api
import schedule.routes

from schedule.services.quote_service import get_daily_dharma
from schedule.builders.time_utils import malaysia_today, malaysia_now
from datetime import datetime, timedelta, date, timezone
from zoneinfo import ZoneInfo
from schedule.builders.schedule_builder import sync_schedule_after_signup_change
from schedule.services.admin_dashboard_service import load_admin_dashboard_data
from lunar_rules import get_special_day_info, get_next_day_remove_info
from flask import request, session, redirect, url_for, render_template_string, flash, jsonify
from schedule.services.food_offering_routes import register_food_offering_routes
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
    set_schedule_setting,
    is_schedule_setting_on,
    is_meal_signup_open,
    set_meal_signup_override,
)

from schedule.services.publish_service import (
    clear_schedule_need_republish,
    is_schedule_published,
    mark_schedule_need_republish,
    publish_schedule_for_date,
    unpublish_schedule_for_date,
    get_schedule_republish_info,
)

from schedule.services.assignment_service import (
    clear_prebook_cache,
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

DATABASE_URL = os.environ.get("DATABASE_URL")
engine = create_engine(
    DATABASE_URL,
    pool_pre_ping=True,
    pool_recycle=300,
) if DATABASE_URL else None

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DATA_DIR = os.path.join(BASE_DIR, "data")

PREBOOK_FILE = os.path.join(
    DATA_DIR,
    "prebook_schedule.xlsx"
)

SCHEDULE_PIN = "1234"

schedule_records = []

# 膳食岗位路由
register_food_offering_routes(
    schedule_bp,
    is_schedule_setting_on=is_schedule_setting_on,
)

# WhatsApp 自动通知 API（第二阶段）
register_whatsapp_agent_api(schedule_bp)

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
    
    today = malaysia_today()

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

    sync_schedule_after_signup_change(
        signup_id,
        action="cancel",
        changed_by="admin"
    )

    return redirect("/schedule/signups")

    
@lru_cache(maxsize=4)
def _read_fixed_buddha_sheet(path, mtime_ns, size):
    """Cache the fixed Buddha roster until the Excel file changes."""
    df = pd.read_excel(path, sheet_name="佛台固定")
    df.columns = df.columns.astype(str).str.strip()
    return df


def _fixed_buddha_dataframe(path):
    stat = os.stat(path)
    return _read_fixed_buddha_sheet(path, stat.st_mtime_ns, stat.st_size)


def load_buddha_name_options():
    file = os.path.join(BASE_DIR, "data", "fixed_schedule.xlsx")

    if not os.path.exists(file):
        return []

    try:
        df = _fixed_buddha_dataframe(file)

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

def get_base_fixed_buddha_for_date(date_str):
    file = os.path.join(BASE_DIR, "data", "fixed_schedule.xlsx")

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

        df = _fixed_buddha_dataframe(file)

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

def get_fixed_buddha_for_date(date_str):
    """当天有效佛台名单：云端覆盖优先；无覆盖时读取固定轮值。"""
    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                SELECT names FROM schedule_buddha_overrides
                WHERE duty_date = %s
            """, (date_str,))
            row = cur.fetchone()
    if row is not None:
        # [] 表示当天全部请假，不能回退到固定名单。
        return list(row['names'] or [])
    return get_base_fixed_buddha_for_date(date_str)


def delete_buddha_override(date_str):
    """恢复指定日期的固定轮值。"""
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute('DELETE FROM schedule_buddha_overrides WHERE duty_date=%s', (date_str,))
        conn.commit()


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
    """按日期保存云端佛台覆盖名单，空名单也必须保存。"""
    names = list(dict.fromkeys(str(n).strip() for n in final_names if str(n).strip()))
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                INSERT INTO schedule_buddha_overrides (duty_date, names, updated_at)
                VALUES (%s, %s::jsonb, NOW())
                ON CONFLICT (duty_date) DO UPDATE
                SET names=EXCLUDED.names, updated_at=NOW()
            """, (date_str, __import__('json').dumps(names, ensure_ascii=False)))
        conn.commit()


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

        clear_prebook_cache()

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

    clear_schedule_need_republish(date_str)

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

    sync_schedule_after_signup_change(
        signup_id,
        action="upsert",
        changed_by="admin"
    )

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
                    select assignment_date
                    from volunteer_schedule_assignments
                    where id = %s
                """, (assignment_id,))

                assignment_info = cur.fetchone()

                if not assignment_info:
                    return "❌ 找不到排班记录<br><a href='/schedule/signups'>返回</a>"

                assignment_date = assignment_info["assignment_date"]

                cur.execute("""
                    update volunteer_schedule_assignments
                    set
                        shift_label = %s,
                        assigned_place = %s,
                        start_time = %s,
                        end_time = %s,
                        remarks = '负责人调整排班',
                        assignment_source = 'admin',
                        locked_by_admin = true
                    where id = %s
                """, (
                    shift_label or None,
                    assigned_place,
                    start_time or None,
                    end_time or None,
                    assignment_id
                ))

                conn.commit()

                mark_schedule_need_republish(assignment_date)

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

{% if row.locked_by_admin %}
    <div style="background:#fff3cd;border:2px solid #f0c36d;border-radius:12px;padding:14px;font-size:22px;margin:15px 0;">
        🔒 已由负责人锁定<br>
        <small>系统不会自动覆盖这笔安排</small>
    </div>
{% else %}
    <div style="background:#e8f5e9;border:2px solid #81c784;border-radius:12px;padding:14px;font-size:22px;margin:15px 0;">
        🤖 系统自动安排<br>
        <small>保存调整后将变成负责人锁定</small>
    </div>
{% endif %}

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
                # 卫生必须统一交给 builder 重排全部卫生岗位，不能固定插入佛堂。
                # 先结束当前数据库事务，再调用 builder（使用独立连接）。
                conn.commit()
                sync_schedule_after_signup_change(signup_id, action="upsert", changed_by="system")
                return

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

def build_signup_notice_whatsapp(date_str):

    date_obj = datetime.strptime(date_str, "%Y-%m-%d").date()

    weekday_map = {
        0: "星期一",
        1: "星期二",
        2: "星期三",
        3: "星期四",
        4: "星期五",
        5: "星期六",
        6: "星期日",
    }

    date_text = f"{date_obj.day}/{date_obj.month}/{date_obj.year} ({weekday_map[date_obj.weekday()]})"

    # 依通知所选日期判断模板，而非依服务器实际日期。
    day_info = get_special_day_info(date_obj)
    big_day = day_info.get("template_type") in ("lunar_1_15", "buddhist_festival")
    # 优先显示完整农历月份，例如「农历十月初一」。
    lunar_label = str(day_info.get("lunar_text") or "").strip()
    festival_names = [str(name).strip() for name in (day_info.get("special_names") or []) if str(name).strip()]
    # 农历与节日名称同时显示，不能因为有农历文字就省略节日名称。
    day_label = "\n".join(part for part in [
        f"（{lunar_label}）" if lunar_label else "",
        *[f"🙏 {name}" for name in festival_names],
    ] if part)
    fixed_buddha = get_fixed_buddha_for_date(date_str)

    # 以下文案必须以通知指定的 date_str 判断，不能使用服务器当天日期。
    special_heading = f"\n{day_label}" if day_label else ""
    altar_start, altar_end = ("6:00am", "8:00am") if big_day else ("8:00am", "10:00am")
    cleaning_start, cleaning_end = altar_start, altar_end
    green_section = (
        "8:00am~10:00am \n🟢 观音堂: \n{hall}\n🟢 活动中心: \n{center}\n\n"
        if big_day else ""
    )

    groups = {
        "佛堂卫生": [],
        "二楼卫生": [],
        "楼梯卫生": [],
        "绿观音堂": [],
        "绿活动中心": [],
        "橙观音堂": [],
        "橙活动中心": [],
        "黄观音堂": [],
        "黄活动中心": [],
    }

    # 19:00 初步表：读取全部有效报名，不依赖 22:00 才产生的正式排班。
    # 只在内存中预览地点，绝不写入 assignments / 发布标志。
    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                SELECT id, volunteer_id, name, role, start_time, end_time
                FROM volunteer_schedule_signups
                WHERE signup_date = %s
                  AND COALESCE(status, 'pending') <> 'cancelled'
                  AND role IN ('卫生', '值班')
                ORDER BY start_time NULLS LAST, id
            """, (date_str,))
            signups = cur.fetchall()
            cur.execute("""
                SELECT signup_id, volunteer_id, name, role, assigned_place,
                       shift_label, start_time, end_time
                FROM volunteer_schedule_assignments
                WHERE assignment_date = %s
                  AND COALESCE(status, 'assigned') <> 'cancelled'
                ORDER BY start_time NULLS LAST, name
            """, (date_str,))
            assigned_rows = cur.fetchall()
            # 供台可能由负责人直接指定（schedule_day_flags），不一定有报名。
            cur.execute("""
                SELECT setup_people FROM schedule_day_flags WHERE flag_date = %s
            """, (date_str,))
            setup_row = cur.fetchone()
            cur.execute("""
                SELECT name FROM volunteer_schedule_signups
                WHERE signup_date = %s AND role = '供台'
                  AND COALESCE(status, 'pending') <> 'cancelled'
                ORDER BY id
            """, (date_str,))
            supply_signup_rows = cur.fetchall()

    # 整理佛台：保留固定轮值，并补上已保存的初一第3位及正式排班人员。
    # 只影响本报名通知的显示，不更改任何数据库排班。
    for row in assigned_rows:
        if row.get("role") == "整理佛台" or row.get("assigned_place") == "整理佛台":
            name = str(row.get("name") or "").strip()
            if name and name not in fixed_buddha:
                fixed_buddha.append(name)
    if big_day and "初一" in lunar_label:
        with get_conn() as conn:
            with conn.cursor(cursor_factory=RealDictCursor) as cur:
                cur.execute("SELECT extra_buddha_person FROM schedule_day_flags WHERE flag_date = %s", (date_str,))
                extra_row = cur.fetchone()
        extra_name = str((extra_row or {}).get("extra_buddha_person") or "").strip()
        if extra_name and extra_name not in fixed_buddha:
            fixed_buddha.append(extra_name)

    setup_people = [name.strip() for name in str((setup_row or {}).get('setup_people') or '').splitlines() if name.strip()]
    supply_people = list(dict.fromkeys(setup_people + [str(r.get('name') or '').strip() for r in supply_signup_rows if r.get('name')]))

    # 卫生：已发布排班可能给同一人多条地点记录，通知必须逐条对应，
    # 不可用单一字典值覆盖前面的地点（例如佛堂+楼梯被误写成只有楼梯）。
    cleaning_places = ('佛堂卫生', '二楼卫生', '楼梯卫生')
    assigned_by_signup = {}
    assigned_by_person = {}
    for row in assigned_rows:
        if row.get('role') != '卫生' or row.get('assigned_place') not in cleaning_places:
            continue
        place = row['assigned_place']
        if row.get('signup_id') is not None:
            assigned_by_signup.setdefault(row['signup_id'], [])
            if place not in assigned_by_signup[row['signup_id']]:
                assigned_by_signup[row['signup_id']].append(place)
        person_key = str(row.get('volunteer_id') or row.get('name') or '')
        assigned_by_person.setdefault(person_key, [])
        if place not in assigned_by_person[person_key]:
            assigned_by_person[person_key].append(place)

    def _minutes(value):
        """支持数据库中的 11:00am、6:00pm、18:00 及 time 对象。"""
        if value is None:
            return None
        if hasattr(value, 'hour'):
            return value.hour * 60 + value.minute
        import re
        value = str(value).strip().lower().replace(' ', '')
        match = re.fullmatch(r'(\d{1,2}):(\d{2})(am|pm)?', value)
        if not match:
            return None
        hour, minute = int(match.group(1)), int(match.group(2))
        meridiem = match.group(3)
        if minute > 59:
            return None
        if meridiem:
            if not 1 <= hour <= 12:
                return None
            hour = hour % 12 + (12 if meridiem == 'pm' else 0)
        elif hour > 23:
            return None
        return hour * 60 + minute

    def _clock(minutes):
        hour, minute = divmod(minutes, 60)
        return f"{hour % 12 or 12}:{minute:02d}{'am' if hour < 12 else 'pm'}"

    duty_signups = []
    for signup in signups:
        role = signup.get('role')
        if role == '卫生':
            person_key = str(signup.get('volunteer_id') or signup.get('name') or '')
            places = (assigned_by_signup.get(signup.get('id'))
                      or assigned_by_person.get(person_key) or [])
            if not places:
                places = [min(cleaning_places, key=lambda k: len(groups[k]))]
            for place in places:
                # 防止重复报名/重复 assignment 在同一地点显示重复姓名。
                if not any(str(r.get('volunteer_id') or r.get('name') or '') == person_key
                           for r in groups[place]):
                    groups[place].append(signup)
        elif role == '值班':
            duty_signups.append(signup)

    # 值班：按上午／下午拆开，地点仅供 19:00 初步通知参考。
    # 22:00 仍由原自动排班器正式计算，可能重新调整。
    duty_counts = {key: 0 for key in ('绿观音堂', '绿活动中心', '橙观音堂', '橙活动中心', '黄观音堂', '黄活动中心')}
    for signup in duty_signups:
        start = _minutes(signup.get('start_time'))
        end = _minutes(signup.get('end_time'))
        if start is None or end is None or end <= start:
            continue
        shifts = ([(480, 600, '绿')] if big_day else []) + [(600, 840, '橙'), (840, 1080, '黄')]
        for part_start, part_end, prefix in shifts:
            seg_start, seg_end = max(start, part_start), min(end, part_end)
            if seg_start >= seg_end:
                continue
            place_a, place_b = prefix + '观音堂', prefix + '活动中心'
            key = place_a if duty_counts[place_a] <= duty_counts[place_b] else place_b
            duty_counts[key] += 1
            groups[key].append({
                'name': signup.get('name'),
                'start_time': _clock(seg_start),
                'end_time': _clock(seg_end),
            })

    def format_people(rows):
        if not rows:
            return ""

        lines = []
        for r in rows:
            name = r.get("name") or ""
            start = r.get("start_time") or ""
            end = r.get("end_time") or ""

            lines.append(str(name))

            if start and end:
                lines.append(f"{start}~{end}")

        return "\n".join(lines)

    def format_duty_people(rows):
        """仅用于值班通知：同地点且起止时间相同的姓名合并一行。"""
        if not rows:
            return ""
        from collections import OrderedDict
        grouped = OrderedDict()
        for row in rows:
            start = str(row.get("start_time") or "").strip()
            end = str(row.get("end_time") or "").strip()
            name = str(row.get("name") or "").strip()
            if not name:
                continue
            key = (start, end)
            names = grouped.setdefault(key, [])
            if name not in names:
                names.append(name)
        lines = []
        for (start, end), names in grouped.items():
            lines.append("　".join(names))
            if start and end:
                lines.append(f"{start}~{end}")
        return "\n".join(lines)

    buddha_text = "　".join(fixed_buddha) if fixed_buddha else ""
    if big_day:
        green_section = green_section.format(
            hall=format_duty_people(groups["绿观音堂"]),
            center=format_duty_people(groups["绿活动中心"]),
        )

    greeting_time = "八点正请安" if big_day else "十点正请安"

    text = f"""师兄们大家好！

{date_text}{special_heading}

{greeting_time}

{altar_start}~{altar_end}  或      
{altar_start}~完成佛台工作 
整理佛台: 
{buddha_text}

{('设师父供台：' + chr(10) + ('　'.join(supply_people) if supply_people else '尚未安排') + chr(10) + chr(10)) if big_day else ''}
{cleaning_start}~{cleaning_end} 或 
{cleaning_start}~完成卫生工作 
佛堂卫生: {"　".join(str(r.get("name") or "").strip() for r in groups["佛堂卫生"] if str(r.get("name") or "").strip())}
二楼卫生: {"　".join(str(r.get("name") or "").strip() for r in groups["二楼卫生"] if str(r.get("name") or "").strip())}
楼梯卫生: {"　".join(str(r.get("name") or "").strip() for r in groups["楼梯卫生"] if str(r.get("name") or "").strip())}
每日卫生义工请注意：清理完卫生之后，请把卫生用具包括吸尘机放回原位（活动中心store里面的小房间）

{green_section}10:00am~2:00pm 
🟠 观音堂: 
{format_duty_people(groups["橙观音堂"])}
🟠 活动中心: 
{format_duty_people(groups["橙活动中心"])}

2:00pm~6:00pm 
🟡 观音堂: 
{format_duty_people(groups["黄观音堂"])}

🟡 活动中心: 
{format_duty_people(groups["黄活动中心"])}


观音堂早晚香 由值班义工带领上香。

观音堂续香
佛友可以续香（黑色无烟香），但是要跟着请安词，必须燃烧完了一支香才续香。如有多位佛友要续香，请先让给先到达观音堂的佛友，轮流续香。

值班义工请注意
1）下雨天记得关上烧送小房子的窗口。
2）在离开观音堂之前，请确保把所有的窗口关上。
3）观音堂第一架的冷气坚决不能调。第二架和第三架可以轮流。

另外，请大家多留意义工群信息，以便大家能够团结一致的护持好观音堂。佛子齐心，普度众生。

义工报名请点击以下链接：
https://gyt-checkin-sg.onrender.com/volunteer

非常感恩大家！
大家功德无量！
🙏🙏🙏
"""

    return text


def load_admin_signup_details(date_str):
    """负责人日期面板：显示所有报名（含取消），地点以正式 assignments 为准。"""
    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                SELECT id, volunteer_id, name, role, start_time, end_time,
                       COALESCE(status, 'pending') AS status, assigned_place
                FROM volunteer_schedule_signups
                WHERE signup_date = %s
                ORDER BY CASE COALESCE(status, 'pending')
                           WHEN 'pending' THEN 0
                           WHEN 'assigned' THEN 1
                           WHEN 'cancelled' THEN 2
                           ELSE 3 END,
                         start_time NULLS LAST, name, role
            """, (date_str,))
            signups = cur.fetchall()
            if not signups:
                return signups

            cur.execute("""
                SELECT signup_id, volunteer_id, role, assigned_place,
                       start_time, end_time
                FROM volunteer_schedule_assignments
                WHERE assignment_date = %s
                  AND COALESCE(status, 'assigned') <> 'cancelled'
                ORDER BY start_time NULLS LAST, end_time NULLS LAST, assigned_place
            """, (date_str,))
            assignments = cur.fetchall()

    by_signup = {}
    by_person_role = {}
    for item in assignments:
        sid = item.get('signup_id')
        if sid is not None:
            by_signup.setdefault(str(sid), []).append(item)
        key = (str(item.get('volunteer_id') or '').strip(),
               str(item.get('role') or '').strip())
        by_person_role.setdefault(key, []).append(item)

    for signup in signups:
        # 已取消的报名不能显示其他有效排班，避免误导负责人。
        if signup.get('status') == 'cancelled':
            continue
        key = (str(signup.get('volunteer_id') or '').strip(),
               str(signup.get('role') or '').strip())
        items = by_signup.get(str(signup['id'])) or by_person_role.get(key, [])
        if items:
            places = []
            seen = set()
            for item in items:
                place = str(item.get('assigned_place') or '').strip()
                if not place:
                    continue
                start = str(item.get('start_time') or '').strip()
                end = str(item.get('end_time') or '').strip()
                detail = f'{place}（{start}～{end}）' if start and end else place
                if detail not in seen:
                    places.append(detail)
                    seen.add(detail)
            if places:
                signup['assignment_places'] = places
    return signups


@schedule_bp.route("/schedule/api/admin-signup-version", methods=["GET"])
def admin_signup_version():
    """轻量检查报名和排班变化；负责人页面只在数据改变时刷新。"""
    if not session.get("schedule_login"):
        return jsonify({"error": "unauthorized"}), 401
    day = (request.args.get("date") or "").strip()
    try:
        day = date.fromisoformat(day).isoformat()
    except (ValueError, TypeError):
        return jsonify({"error": "invalid date"}), 400
    import hashlib
    import json
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                SELECT id, volunteer_id, name, role, start_time, end_time,
                       status, assigned_place, remarks
                FROM volunteer_schedule_signups
                WHERE signup_date = %s ORDER BY id
            """, (day,))
            signups = cur.fetchall()
            cur.execute("""
                SELECT volunteer_id, role, shift_label, assigned_place,
                       start_time, end_time, status, locked_by_admin
                FROM volunteer_schedule_assignments
                WHERE assignment_date = %s
                ORDER BY volunteer_id, role, shift_label, assigned_place, start_time, end_time
            """, (day,))
            assignments = cur.fetchall()
            cur.execute("""
                SELECT is_published, need_republish, last_schedule_change
                FROM schedule_day_flags WHERE flag_date = %s
            """, (day,))
            flags = cur.fetchall()
    payload = json.dumps([signups, assignments, flags], default=str, ensure_ascii=False)
    version = hashlib.sha256(payload.encode("utf-8")).hexdigest()
    return jsonify({"version": version})


@schedule_bp.route("/schedule", methods=["GET", "POST"])
@schedule_bp.route("/schedule/admin", methods=["GET", "POST"])
def schedule_admin():
    if not session.get("schedule_login"):
        if request.method == "POST":
            pin = request.form.get("pin", "").strip()

            if pin == SCHEDULE_PIN:
                session.permanent = True
                session["schedule_login"] = True
                return redirect(url_for("schedule.schedule_admin"))

            return "❌ PIN 错误<br><a href='/schedule'>返回</a>"

        return render_template_string(LOGIN_HTML)


    now = malaysia_now()

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
    _dashboard_started = time.perf_counter()
    dashboard = load_admin_dashboard_data(
        mode=mode,
        override_date=override_date,
    )
    _dashboard_loaded = time.perf_counter()

    signup_details = load_admin_signup_details(override_date) if mode in ("day", "signup", "special", "controls") else []
    buddha_names = load_buddha_name_options()
    fixed_buddha_today = get_fixed_buddha_for_date(override_date)
    # 佛台请假日期独立于负责人主页日期；允许提前安排未来30天。
    buddha_leave_dates = [(malaysia_today() + timedelta(days=i)).isoformat() for i in range(31)] if mode == "special" else []
    requested_leave_date = request.args.get("buddha_leave_date", "")
    selected_buddha_leave_date = (requested_leave_date if requested_leave_date in buddha_leave_dates
                                  else (override_date if override_date in buddha_leave_dates
                                        else (buddha_leave_dates[0] if buddha_leave_dates else override_date)))
    buddha_leave_names = (get_fixed_buddha_for_date(selected_buddha_leave_date)
                          if mode == "special" else [])
    # 初一补人独立选日期，不影响负责人页面及佛台请假使用的日期。
    extra_buddha_dates = []
    for offset in (range(91) if mode == "special" else ()):
        candidate = malaysia_today() + timedelta(days=offset)
        info = get_special_day_info(candidate) or {}
        if info.get("template_type") == "lunar_1_15" and "初一" in str(info.get("lunar_text") or ""):
            extra_buddha_dates.append(candidate.isoformat())
    requested_extra_date = request.args.get("extra_buddha_date", "")
    selected_extra_buddha_date = (requested_extra_date if requested_extra_date in extra_buddha_dates
                                  else (extra_buddha_dates[0] if extra_buddha_dates else ""))
    extra_buddha_fixed_names = (get_fixed_buddha_for_date(selected_extra_buddha_date)
                                if selected_extra_buddha_date else [])
    extra_buddha_saved = ""
    if selected_extra_buddha_date:
        with get_conn() as conn:
            with conn.cursor() as cur:
                cur.execute("SELECT extra_buddha_person FROM schedule_day_flags WHERE flag_date = %s",
                            (selected_extra_buddha_date,))
                saved_row = cur.fetchone()
                extra_buddha_saved = (saved_row.get("extra_buddha_person") or "") if saved_row else ""
    _roster_loaded = time.perf_counter()

    if os.environ.get("SCHEDULE_DASHBOARD_TIMING") == "1":
        print(
            "[schedule-admin] dashboard={:.3f}s roster={:.3f}s".format(
                _dashboard_loaded - _dashboard_started,
                _roster_loaded - _dashboard_loaded,
            ),
            flush=True,
        )


    master_dates = _master_table_dates()
    flags = dashboard["day_flags"] or {}
    selected_master_date = override_date if override_date in {v["date"] for v in master_dates} else (master_dates[0]["date"] if master_dates else "")
    # 日历选中的设供台日期可能不同于负责人主页日期；必须单独读取。
    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("SELECT setup_people FROM schedule_day_flags WHERE flag_date = %s", (selected_master_date,))
            setup_row = cur.fetchone()
    setup_saved = [v.strip() for v in str((setup_row or {}).get("setup_people") or "").splitlines() if v.strip()]

    # 首页提醒：合并义工报名和负责人已安排人员；不改动报名表。
    supply_alerts_display = []
    alert_items = dashboard.get("supply_alerts") or []
    alert_dates = []
    for item in alert_items:
        raw_day = item.get("date")
        try:
            alert_dates.append(raw_day if isinstance(raw_day, date) else date.fromisoformat(str(raw_day)[:10]))
        except (ValueError, TypeError):
            alert_dates.append(None)
    scheduled_by_date = {}
    valid_dates = [d for d in alert_dates if d is not None]
    if valid_dates:
        with get_conn() as conn:
            with conn.cursor(cursor_factory=RealDictCursor) as cur:
                cur.execute("""
                    SELECT flag_date, setup_people FROM schedule_day_flags
                    WHERE flag_date = ANY(%s)
                """, (valid_dates,))
                for row in cur.fetchall():
                    scheduled_by_date[str(row["flag_date"])] = [
                        name.strip() for name in str(row.get("setup_people") or "").splitlines() if name.strip()
                    ]
    for item, day in zip(alert_items, alert_dates):
        item_display = dict(item)
        # 只修正供台报名提醒的大日子名称，保留原有报名及负责人安排资料。
        festival_label = ""
        if day:
            info = get_special_day_info(day)
            lunar_text = str(info.get("lunar_text") or "")
            names = info.get("special_names") or []
            if isinstance(names, str):
                names = [names]
            names = [str(n).strip() for n in names if str(n).strip()]
            if item.get("type") == "lunar_1_15":
                festival_label = "初一" if "初一" in lunar_text else ("十五" if "十五" in lunar_text else "初一 / 十五")
            elif item.get("type") == "buddhist_festival":
                festival_label = "恩师卢军宏涅槃日" if "十月初六" in lunar_text or "十月初六" in "、".join(names) else ("、".join(names) if names else "佛诞日")
        item_display["festival_label"] = festival_label
        item_display["scheduled_names"] = scheduled_by_date.get(day.isoformat(), []) if day else []
        supply_alerts_display.append(item_display)
    # 收供台日期独立于设供台日期，允许未来90天内的普通日。
    remove_date_arg = request.args.get("remove_date", "").strip()
    try:
        remove_day = date.fromisoformat(remove_date_arg)
    except ValueError:
        remove_day = malaysia_today()
    if not malaysia_today() <= remove_day <= malaysia_today() + timedelta(days=90):
        remove_day = malaysia_today()
    selected_remove_date = remove_day.isoformat()
    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("SELECT remove_people FROM schedule_day_flags WHERE flag_date = %s", (selected_remove_date,))
            remove_row = cur.fetchone()
    remove_saved = [v.strip() for v in str((remove_row or {}).get("remove_people") or "").splitlines() if v.strip()]


    return render_template_string(
        SCHEDULE_HTML,
        mode=mode,
        times=TIME_OPTIONS,
        roles=ROLE_OPTIONS,
        tomorrow=default_schedule_date,
        buddha_names=buddha_names,
        master_dates=master_dates,
        selected_master_date=selected_master_date,
        selected_remove_date=selected_remove_date,
        remove_max_date=(malaysia_today() + timedelta(days=90)).isoformat(),
        remove_min_date=malaysia_today().isoformat(),
        setup_saved=setup_saved,
        remove_saved=remove_saved,
        override_date=override_date,
        fixed_buddha_today=fixed_buddha_today,
        buddha_leave_dates=buddha_leave_dates,
        selected_buddha_leave_date=selected_buddha_leave_date,
        buddha_leave_names=buddha_leave_names,
        extra_buddha_dates=extra_buddha_dates,
        selected_extra_buddha_date=selected_extra_buddha_date,
        extra_buddha_fixed_names=extra_buddha_fixed_names,
        extra_buddha_saved=extra_buddha_saved,
        multi_day_signup_open=dashboard["multi_day_signup_open"],
        meal_signup_open=is_meal_signup_open(),
        food_offering_open=dashboard["food_offering_open"],
        records=dashboard["records"],
        pending_counts=dashboard["pending_counts"],
        day_summary=dashboard["day_summary"],
        signup_details=signup_details,
        is_today_or_past=dashboard["is_today_or_past"],
        day_flags=dashboard["day_flags"],
        special_day_info=dashboard["special_day_info"],
        remove_info=dashboard["remove_info"],
        shortage_summary=dashboard["shortage_summary"],
        supply_signups=dashboard["supply_signups"],
        supply_alerts=supply_alerts_display,
        is_published=dashboard["is_published"],
        republish_info=dashboard["republish_info"],
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

    vol_keyword = request.form.get("keyword", "").strip()

    branch = request.form.get("branch", "CHE").strip().upper()
    vol_keyword = apply_branch_prefix(vol_keyword, branch)
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


@schedule_bp.route("/schedule/toggle_setting", methods=["POST"])
def toggle_schedule_setting():

    print("======== TOGGLE_SETTING ENTERED =======", flush=True)

    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    key = request.form.get("key", "").strip()
    force = request.form.get("force", "").strip()
    return_to = request.form.get("return_to") or url_for("schedule.schedule_admin")

    allowed_keys = {
        "multi_day_signup_open",
        "meal_signup_open",
        "food_offering_open",
    }

    if key not in allowed_keys:
        return "❌ 不允许的设置"

    if key == "meal_signup_open":
        if force in ("true", "false"):
            set_meal_signup_override(force == "true", updated_by="admin")
        else:
            set_meal_signup_override(not is_meal_signup_open(), updated_by="admin")
        return redirect(return_to)

    if force in ("true", "false"):
        set_schedule_setting(key, force, updated_by="admin")

        print("key =", key)
        print("force =", force)
        print(
            "saved =",
            get_schedule_setting(
                "food_offering_open",
                "false"
            )
        )
    else:
        current = get_schedule_setting(key, "false")
        new_value = "false" if current == "true" else "true"
        set_schedule_setting(key, new_value, updated_by="admin")

    return redirect(return_to)


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

    with get_conn() as conn:
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

                # ============================================================
                # 检查重复报名
                #
                # 值班：
                #   同一天允许多轮值班
                #   只有「开始时间 + 结束时间」完全相同才算重复
                #
                # 其他岗位：
                #   同一天同岗位只允许一笔
                # ============================================================

                if role == "值班":

                    cur.execute("""
                        select id
                        from volunteer_schedule_signups
                        where volunteer_id = %s
                        and signup_date = %s
                        and role = '值班'
                        and start_time = %s
                        and end_time = %s
                        and coalesce(status, 'pending') <> 'cancelled'
                        limit 1
                    """, (
                        volunteer_id,
                        signup_date,
                        role_start,
                        role_end
                    ))

                else:

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
    
@schedule_bp.route("/schedule/dry_run_2200", methods=["GET"])
def schedule_dry_run_2200():
    # 负责人专用，只读预演
    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    target = (request.args.get("date") or "").strip()
    try:
        target = datetime.strptime(target, "%Y-%m-%d").date().isoformat()
    except ValueError:
        return "日期格式错误，请返回负责人首页重新选择日期。", 400

    report = None
    error = None
    try:
        from schedule.services.whatsapp_publish_2200 import dry_run_2200
        report = dry_run_2200(target)
    except Exception:
        import logging
        logging.getLogger(__name__).exception("22:00 read-only dry-run failed for %s", target)
        error = "预演计算失败，请检查 Render 日志（没有执行发布或发送）。"

    return render_template_string('''<!doctype html><html lang="zh"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>22:00 自动排班安全预演</title>
<style>body{font-family:Arial,"Microsoft YaHei",sans-serif;background:#f2f6f8;color:#172a39;margin:0;padding:20px}
main{max-width:920px;margin:auto;background:white;padding:24px;border-radius:16px;box-shadow:0 4px 20px #0001}
.badge{display:inline-block;background:#d9f4e8;color:#075a35;padding:8px 12px;border-radius:9px;font-weight:bold}
.warn{background:#fff3d8;padding:12px;border-radius:9px;margin:14px 0}table{width:100%;border-collapse:collapse}th,td{padding:10px;border-bottom:1px solid #e1e8ed;text-align:left}
th{background:#eef4f7}a{display:inline-block;margin:12px 0;color:#0755b6}ul{line-height:1.9}
</style></head><body><main>
<h1>🧪 22:00 自动排班安全预演</h1><div class="badge">只读测试 · 不保存 · 不发布 · 不发送 WhatsApp</div>
<p><b>预演日期：</b>{{ target }}</p>
<a href="/schedule?mode=day&amp;override_date={{ target }}">← 返回负责人页面</a>
{% if error %}<div class="warn">❌ {{ error }}</div>{% elif report %}
<p><b>发布状态：</b>{{ '已发布' if report.published else '未发布' }}　｜　
<b>现有排班记录：</b>{{ report.existing_assignment_count }}　｜　
<b>理论排班记录：</b>{{ report.theoretical_assignment_count }}　｜　
<b>值班报名：</b>{{ report.duty_signup_count }}</p>
<div class="warn">⚠️ 当前预演只核对<strong>值班报名</strong>的时间覆盖；不代表卫生、佛台及其他岗位已通过完整验证。<br>
“未安排”指数据库当前尚无对应排班，不代表正式自动排班一定失败；请参考右侧理论排班。</div>
<h2>值班报名检查</h2>
{% if report.coverage_checks %}<div style="overflow-x:auto"><table><thead><tr><th>义工</th><th>报名时间</th><th>当前排班状态</th><th>理论安排</th></tr></thead><tbody>
{% for c in report.coverage_checks %}<tr><td>{{ c.name }}</td><td>{{ c.signup_time|join(' ~ ') }}</td>
<td>{% if c.status == 'covered' %}✅ 已覆盖{% elif c.status == 'missing' %}⚪ 当前未安排{% elif c.status == 'invalid_signup_time' %}❌ 报名时间异常{% else %}⚠️ 当前排班未完全覆盖{% endif %}</td>
<td>{% if c.theoretical %}{% for a in c.theoretical %}{{ a.assigned_place }} {{ a.start_time }}–{{ a.end_time }}<br>{% endfor %}{% else %}—{% endif %}</td></tr>{% endfor %}
</tbody></table></div>{% else %}<p>此日期没有有效值班报名。</p>{% endif %}
<h2>预演注意事项</h2><ul>{% for w in report.warnings %}<li>{{ w }}</li>{% endfor %}</ul>
{% endif %}<a href="/schedule?mode=day&amp;override_date={{ target }}">← 返回负责人页面</a>
</main></body></html>''', target=target, report=report, error=error)


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

    if is_schedule_published(date):
        return """
        <h1>🔒 已发布，不能重新自动排班</h1>
        <p>这一天已经发布正式值班表。</p>
        <p>发布后新报名义工会自动补排，不需要重新自动排班。</p>
        <a href="/schedule?mode=day">返回负责人首页</a>
        """

    target_date = datetime.strptime(date, "%Y-%m-%d").date()

    if target_date < malaysia_today():
        return """
        <h1>❌ 当天排班已锁定</h1>
        <p>过去日期不能重新自动排班，只能使用「补入待安排」。</p>
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


@schedule_bp.route("/schedule/reports")
def schedule_reports():

    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    today = malaysia_today()

    ym = request.args.get("ym", today.strftime("%Y-%m"))

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


def _master_table_dates(days=90):
    """Use the existing lunar rules, not a separately maintained festival calendar."""
    today = malaysia_today()
    dates = []
    for offset in range(days + 1):
        day = today + timedelta(days=offset)
        info = get_special_day_info(day) or {}
        if info.get("template_type") not in ("lunar_1_15", "buddhist_festival"):
            continue
        title = (info.get("festival_name") or info.get("name") or
                 info.get("title") or info.get("lunar_text") or "大日子")
        dates.append({"date": day.isoformat(), "label": f"{day:%Y-%m-%d} · {title}"})
    return dates


@schedule_bp.route("/schedule/save_day_flags", methods=["POST"])
def save_day_flags():
    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    date_str = request.form.get("date", "").strip()
    form_type = request.form.get("form_type", "legacy").strip()
    try:
        selected = datetime.strptime(date_str, "%Y-%m-%d").date()
    except ValueError:
        return "❌ 日期格式错误", 400

    if form_type == "remove_table" and not (malaysia_today() <= selected <= malaysia_today() + timedelta(days=90)):
        return "❌ 收供台日期必须在未来90天内", 400

    if form_type == "extra_buddha":
        info = get_special_day_info(selected) or {}
        if not (malaysia_today() <= selected <= malaysia_today() + timedelta(days=90)
                and info.get("template_type") == "lunar_1_15"
                and "初一" in str(info.get("lunar_text") or "")):
            return "❌ 初一补人只能选择未来90天内的农历初一", 400

    if form_type == "master_table":
        allowed = {item["date"] for item in _master_table_dates()}
        if date_str not in allowed:
            return "❌ 只能选择未来90天内的系统大日子", 400

    setup_people = "\n".join(dict.fromkeys(x for x in (
        request.form.get("setup_person_1", "").strip(),
        request.form.get("setup_person_2", "").strip()
    ) if x))
    remove_people = "\n".join(dict.fromkeys(x for x in (
        request.form.get("remove_person_1", "").strip(),
        request.form.get("remove_person_2", "").strip(),
        request.form.get("remove_extra_person", "").strip()
    ) if x))
    extra_buddha_person = request.form.get("extra_buddha_person", "").strip()

    with get_conn() as conn:
        with conn.cursor() as cur:
            if form_type == "master_table":
                # Preserve existing extra-buddha override when saving master-table names.
                cur.execute("""
                    INSERT INTO schedule_day_flags
                        (flag_date, need_setup_master_table, need_remove_master_table,
                         setup_people, remove_people, updated_at)
                    VALUES (%s, %s, %s, %s, %s, NOW())
                    ON CONFLICT (flag_date) DO UPDATE SET
                        need_setup_master_table=EXCLUDED.need_setup_master_table,
                        setup_people=EXCLUDED.setup_people,
                        updated_at=NOW()
                """, (date_str, bool(setup_people), False,
                      setup_people, ""))
            elif form_type == "remove_table":
                # 没有收供台人员 = 继续供奉，不自动撤台。
                cur.execute("""
                    INSERT INTO schedule_day_flags
                        (flag_date, need_remove_master_table, remove_people, updated_at)
                    VALUES (%s, %s, %s, NOW())
                    ON CONFLICT (flag_date) DO UPDATE SET
                        need_remove_master_table=EXCLUDED.need_remove_master_table,
                        remove_people=EXCLUDED.remove_people,
                        updated_at=NOW()
                """, (date_str, bool(remove_people), remove_people))
            elif form_type == "extra_buddha":
                # Do not wipe master-table names when saving the extra Buddha helper.
                cur.execute("""
                    INSERT INTO schedule_day_flags (flag_date, extra_buddha_person, updated_at)
                    VALUES (%s, %s, NOW())
                    ON CONFLICT (flag_date) DO UPDATE SET
                        extra_buddha_person=EXCLUDED.extra_buddha_person,
                        updated_at=NOW()
                """, (date_str, extra_buddha_person))
                # 已发布的初一：只同步本按钮管理的第3位佛台义工，不重排其他人员。
                cur.execute("""
                    SELECT is_published FROM schedule_day_flags
                    WHERE flag_date=%s FOR UPDATE
                """, (date_str,))
                published_row = cur.fetchone()
                if published_row and published_row.get("is_published"):
                    # 仅清除本功能以前插入的记录；原有固定佛台及其他岗位不动。
                    cur.execute("""
                        DELETE FROM volunteer_schedule_assignments
                        WHERE assignment_date=%s AND role='整理佛台'
                          AND remarks='初一额外第3位整理佛台义工'
                    """, (date_str,))
                    if extra_buddha_person:
                        cur.execute("""
                            SELECT 1 FROM volunteer_schedule_assignments
                            WHERE assignment_date=%s AND role='整理佛台'
                              AND name=%s AND COALESCE(status,'assigned') <> 'cancelled'
                            LIMIT 1
                        """, (date_str, extra_buddha_person))
                        if not cur.fetchone():
                            cur.execute("""
                                INSERT INTO volunteer_schedule_assignments
                                    (name, assignment_date, role, shift_label,
                                     assigned_place, status, remarks, created_at)
                                VALUES (%s,%s,'整理佛台','整理佛台',
                                        '整理佛台','assigned','初一额外第3位整理佛台义工',NOW())
                            """, (extra_buddha_person, date_str))
            else:
                # Compatibility with older forms.
                cur.execute("""
                    INSERT INTO schedule_day_flags
                        (flag_date, need_setup_master_table, need_remove_master_table,
                         setup_people, remove_people, extra_buddha_person, updated_at)
                    VALUES (%s, %s, %s, %s, %s, %s, NOW())
                    ON CONFLICT (flag_date) DO UPDATE SET
                        need_setup_master_table=EXCLUDED.need_setup_master_table,
                        need_remove_master_table=EXCLUDED.need_remove_master_table,
                        setup_people=EXCLUDED.setup_people,
                        remove_people=EXCLUDED.remove_people,
                        extra_buddha_person=EXCLUDED.extra_buddha_person,
                        updated_at=NOW()
                """, (date_str, bool(setup_people), bool(remove_people),
                      setup_people, remove_people, extra_buddha_person))
        conn.commit()

    return redirect(url_for("schedule.schedule_admin", mode="special",
                            override_date=(date_str if form_type != "remove_table" else request.form.get("setup_date", "") or date_str),
                            remove_date=(date_str if form_type == "remove_table" else ""),
                            panel=("master-table" if form_type in ("master_table", "remove_table") else "extra-buddha"),
                            extra_buddha_date=(date_str if form_type == "extra_buddha" else "")))


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


def normalize_whatsapp_keyword(text):
    text = str(text or "").strip().upper().replace(" ", "")

    if text.startswith("STW-"):
        return text

    if text.startswith("STW") and text[3:].isdigit():
        return f"STW-{int(text[3:])}"

    if text.startswith("CHE-"):
        return text

    if text.startswith("CHE") and text[3:].isdigit():
        return f"CHE-{int(text[3:])}"

    return text


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

                name = normalize_whatsapp_keyword(parts[0])

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

                # 同一位义工同一天可以报名多轮值班。
                # 只有日期、岗位、开始时间、结束时间全部相同，才视为重复报名。
                cur.execute("""
                    select id
                    from volunteer_schedule_signups
                    where volunteer_id = %s
                    and signup_date = %s
                    and role = '值班'
                    and start_time = %s
                    and end_time = %s
                    and coalesce(status,'pending') <> 'cancelled'
                    limit 1
                """, (
                    volunteer_id,
                    signup_date,
                    start_time,
                    end_time
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

    try:
        date_obj = datetime.strptime(date, "%Y-%m-%d").date()
        if not original_names or len(original_names) != len(replacement_names):
            raise ValueError("没有有效的佛台人员资料")
        save_buddha_override(date_obj.isoformat(), final_names)
    except Exception as exc:
        flash(f"保存佛台请假失败：{exc}")

    panel = request.form.get("return_panel", "")
    if panel in ("extra-buddha", "buddha-leave"):
        kwargs = {"mode": "special", "override_date": request.form.get("return_override_date") or date, "panel": panel}
        kwargs["extra_buddha_date" if panel == "extra-buddha" else "buddha_leave_date"] = date
        return redirect(url_for("schedule.schedule_admin", **kwargs))
    return redirect(url_for("schedule.schedule_admin", mode="day", override_date=date))


@schedule_bp.route("/schedule/override/reset", methods=["POST"])
def schedule_override_reset():
    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))
    date_str = request.form.get("date", "").strip()
    try:
        date.fromisoformat(date_str)
        delete_buddha_override(date_str)
    except Exception as exc:
        flash(f"恢复固定佛台失败：{exc}")
    panel = request.form.get("return_panel", "")
    if panel in ("extra-buddha", "buddha-leave"):
        kwargs = {"mode": "special", "override_date": request.form.get("return_override_date") or date_str, "panel": panel}
        kwargs["extra_buddha_date" if panel == "extra-buddha" else "buddha_leave_date"] = date_str
        return redirect(url_for("schedule.schedule_admin", **kwargs))
    return redirect(url_for("schedule.schedule_admin", mode="day", override_date=date_str))


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


@schedule_bp.route(
    "/schedule/generate_shortage_notice",
    methods=["GET", "POST"]
)
def generate_shortage_notice():

    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    date_str = (
        request.form.get("date")
        or request.args.get("date")
    )

    return_section = request.values.get("return_section", "")
    if return_section not in ("management", "today"):
        return_section = ""

    msg = build_shortage_notice_from_assignments(date_str)

    return render_template_string("""
    <!doctype html>
    <html lang="zh">
    <head>
    <meta charset="utf-8">
    <title>缺义工通知</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">

    <link rel="stylesheet"
        href="{{ url_for('static', filename='css/toolbox.css') }}">

    <style>
    .notice-textarea{
        width:100%;
        min-height:420px;
        font-size:20px;
        padding:16px;
        border:1px solid #d1d5db;
        border-radius:16px;
        resize:vertical;
        box-sizing:border-box;
    }
    .notice-actions{
        display:grid;
        grid-template-columns:1fr 1fr;
        gap:14px;
        margin-top:20px;
    }
    @media(max-width:700px){
        .notice-actions{
            grid-template-columns:1fr;
        }
    }
    </style>

    </head>
    <body>

    <div class="page">

        <h1 class="page-title">
            📢 缺义工通知
        </h1>

        <p class="page-subtitle">
            WhatsApp 缺义工通知，可直接复制发送
        </p>

        <div class="card">

            <div class="section-title">
                📄 通知内容
            </div>

            <textarea
                id="notice_text"
                class="notice-textarea"
            >{{ msg }}</textarea>

            <div class="notice-actions">

                <button
                    class="btn-tool btn-primary"
                    onclick="copyNotice()">

                    📋 复制通知

                </button>

                <a
                    class="btn-tool btn-secondary"
                    href="{{ url_for('schedule.schedule_notice_center', date=date_str, return_section=return_section) }}">

                    ← 返回通知中心

                </a>

            </div>

        </div>

    </div>

    <script>
    function copyNotice(){

        const ta = document.getElementById("notice_text");

        if(!ta){
            alert("❌ 找不到通知内容");
            return;
        }

        ta.focus();
        ta.select();
        ta.setSelectionRange(0, 999999);

        const ok = document.execCommand("copy");

        if(ok){
            alert("✅ 已复制通知");
        }else{
            alert("❌ 复制失败，请手动全选复制");
        }
    }
    </script>

    </body>
    </html>
    """,
    msg=msg,
    date_str=date_str,
    return_section=return_section
    )


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

    <link rel="stylesheet"
          href="{{ url_for('static', filename='css/toolbox.css') }}">

    <style>
    .assigned-wrap{
        max-width:1100px;
        margin:auto;
        padding:24px;
    }

    .assignment-section{
        background:#f8fafc;
        border:1px solid #e5e7eb;
        border-radius:18px;
        padding:20px;
        margin:16px 0;
    }

    .assignment-place-title{
        font-size:30px;
        font-weight:bold;
        margin-bottom:14px;
    }

    .person-row{
        display:flex;
        justify-content:space-between;
        align-items:center;
        gap:16px;
        border-bottom:1px solid #e5e7eb;
        padding:14px 0;
        font-size:28px;
    }

    .person-row:last-child{
        border-bottom:none;
    }

    .person-name{
        font-weight:bold;
    }

    .person-time{
        color:#6b7280;
        font-size:24px;
    }

    .edit-link{
        font-size:22px;
        padding:8px 14px;
        border-radius:12px;
        background:#dbeafe;
        color:#1e40af;
        font-weight:bold;
        white-space:nowrap;
    }

    .shift-title{
        font-size:34px;
        margin-top:26px;
    }

    .empty-line{
        font-size:24px;
        color:#6b7280;
    }
    </style>
    </head>

    <body>

    <div class="assigned-wrap">

        <div class="card">

            <h1 class="page-title">
                📅 {{ date_str }} 排班结果
            </h1>

            <p class="page-subtitle">
                查看正式安排，也可进入个别义工调整岗位。
            </p>

            <div class="btn-row">

                <a class="btn-tool btn-gray"
                href="/schedule?mode=day&override_date={{ date_str }}">
                    ⬅ 返回当天值班安排
                </a>

                <form method="post" action="/schedule/copy_whatsapp">
                    <input type="hidden" name="date" value="{{ date_str }}">
                    <button class="btn-tool btn-green" type="submit">
                        📋 生成 WhatsApp 文字
                    </button>
                </form>

            </div>

        </div>

        <div class="card">

            <h2 class="section-title">🧹 卫生</h2>

            {% for place, names in cleaning_groups.items() %}
                <div class="assignment-section">

                    <div class="assignment-place-title">
                        {{ place }}
                    </div>

                    {% if names %}
                        {% for item in names %}
                            <div class="person-row">
                                <div>
                                    <span class="person-name">{{ item.name }}</span>
                                </div>

                                <a class="edit-link"
                                href="/schedule/edit_assigned/{{ item.id }}">
                                    ✏️ 调整
                                </a>
                            </div>
                        {% endfor %}
                    {% else %}
                        <div class="empty-line">暂无安排</div>
                    {% endif %}

                </div>
            {% endfor %}

        </div>

        <div class="card">

            <h2 class="section-title">🙏 供台</h2>

            <div class="assignment-section">

                {% if offering_group %}
                    {% for item in offering_group %}
                        <div class="person-row">
                            <div>
                                <span class="person-name">{{ item.name }}</span>
                            </div>

                            <a class="edit-link"
                            href="/schedule/edit_assigned/{{ item.id }}">
                                ✏️ 调整
                            </a>
                        </div>
                    {% endfor %}
                {% else %}
                    <div class="empty-line">暂无安排</div>
                {% endif %}

            </div>

        </div>

        <div class="card">

            <h2 class="section-title">🏠 值班</h2>

            {% for shift, places in duty_groups.items() %}

                {% if shift == "绿" %}
                    <h3 class="shift-title">🟢 绿班</h3>
                {% elif shift == "橙" %}
                    <h3 class="shift-title">🟠 橙班</h3>
                {% elif shift == "黄" %}
                    <h3 class="shift-title">🟡 黄班</h3>
                {% else %}
                    <h3 class="shift-title">{{ shift }}</h3>
                {% endif %}

                {% for place, names in places.items() %}
                    <div class="assignment-section">

                        <div class="assignment-place-title">
                            {{ place }}
                        </div>

                        {% if names %}
                            {% for item in names %}
                                <div class="person-row">

                                    <div>
                                        <span class="person-name">
                                            {{ item.name }}
                                        </span>

                                        {% if item.start_time and item.end_time %}
                                            <span class="person-time">
                                                （{{ item.start_time }} ~ {{ item.end_time }}）
                                            </span>
                                        {% endif %}
                                    </div>

                                    <a class="edit-link"
                                    href="/schedule/edit_assigned/{{ item.id }}">
                                        ✏️ 调整
                                    </a>

                                </div>
                            {% endfor %}
                        {% else %}
                            <div class="empty-line">暂无安排</div>
                        {% endif %}

                    </div>
                {% endfor %}

            {% endfor %}

        </div>

        {% if other_groups %}
        <div class="card">

            <h2 class="section-title">其他</h2>

            {% for place, names in other_groups.items() %}
                <div class="assignment-section">

                    <div class="assignment-place-title">
                        {{ place }}
                    </div>

                    {% for item in names %}
                        <div class="person-row">

                            <div>
                                <span class="person-name">
                                    {{ item.name }}
                                </span>

                                {% if item.start_time and item.end_time %}
                                    <span class="person-time">
                                        （{{ item.start_time }} ~ {{ item.end_time }}）
                                    </span>
                                {% endif %}
                            </div>

                            <a class="edit-link"
                            href="/schedule/edit_assigned/{{ item.id }}">
                                ✏️ 调整
                            </a>

                        </div>
                    {% endfor %}

                </div>
            {% endfor %}

        </div>
        {% endif %}

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

    <link rel="stylesheet"
          href="{{ url_for('static', filename='css/toolbox.css') }}">

    <style>
    .reports-wrap{
        max-width:900px;
        margin:auto;
        padding:24px;
    }

    .file-row{
        display:flex;
        justify-content:space-between;
        align-items:center;
        gap:18px;
        padding:18px 0;
        border-bottom:1px solid #e5e7eb;
        font-size:26px;
    }

    .file-row:last-child{
        border-bottom:none;
    }

    .file-name{
        font-weight:bold;
    }

    .download-btn{
        min-width:150px;
    }

    @media (max-width:700px){
        .file-row{
            flex-direction:column;
            align-items:stretch;
        }

        .download-btn{
            width:100%;
        }
    }
    </style>
    </head>

    <body>

    <div class="reports-wrap">

        <div class="card">

            <h1 class="page-title">📚 观音堂资料中心</h1>

            <p class="page-subtitle">
                提供月报、年报及相关资料下载。
            </p>

            {% if report_files %}

                {% for f in report_files %}
                <div class="file-row">

                    <div class="file-name">
                        📄 {{ f.display_name }}
                    </div>

                    <a class="btn-tool btn-blue download-btn"
                    href="{{ f.url }}">
                        📥 下载
                    </a>

                </div>
                {% endfor %}

            {% else %}

                <div class="empty-state">
                    <div class="empty-icon">📭</div>
                    <div class="empty-title">
                        目前还没有上传报表
                    </div>
                    <div class="empty-text">
                        上传月报或年报后，会显示在这里。
                    </div>
                </div>

            {% endif %}

            <div class="btn-row">
                <a class="btn-tool btn-gray btn-full"
                href="/schedule/admin">
                    ⬅ 返回负责人页面
                </a>
            </div>

        </div>

    </div>

    </body>
    </html>
    """, report_files=report_files)


# 手动查看正式发布后的新增报名；不触发 WhatsApp 自动发送。
@schedule_bp.route("/schedule/late_signup_notice", methods=["GET"])
def schedule_late_signup_notice():
    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    date_str = (request.args.get("date") or "").strip()
    try:
        target_date = date.fromisoformat(date_str)
    except ValueError:
        return "日期格式不正确", 400

    def _late_notice_aware(value):
        if value.tzinfo is None:
            return value.replace(tzinfo=timezone.utc)
        return value

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                SELECT is_published, published_at, late_notice_sent_at, now() AS page_snapshot_at
                FROM schedule_day_flags WHERE flag_date = %s
            """, (date_str,))
            flag = cur.fetchone() or {}
            signups = []
            assignments = []
            if flag.get("is_published") and flag.get("published_at"):
                cur.execute("""
                    SELECT id, volunteer_id, name, role, start_time, end_time
                    FROM volunteer_schedule_signups
                    WHERE signup_date = %s
                      AND created_at > %s
                      AND created_at <= %s
                      AND COALESCE(status, 'pending') <> 'cancelled'
                    ORDER BY created_at, id
                """, (date_str, max(
                    # PostgreSQL may return timestamp without timezone as naive datetime.
                    # Treat legacy naive DB timestamps as UTC before comparing with MYT 22:00.
                    _late_notice_aware(flag["published_at"]),
                    _late_notice_aware(flag.get("late_notice_sent_at") or flag["published_at"]),
                    datetime.combine(target_date - timedelta(days=1), datetime.min.time(),
                                     tzinfo=ZoneInfo("Asia/Kuala_Lumpur")) + timedelta(hours=22)
                ), flag["page_snapshot_at"]))
                signups = cur.fetchall()
                if signups:
                    ids = [row["id"] for row in signups]
                    cur.execute("""
                        SELECT signup_id, role, assigned_place, start_time, end_time
                        FROM volunteer_schedule_assignments
                        WHERE assignment_date = %s
                          AND signup_id = ANY(%s)
                          AND COALESCE(status, 'assigned') <> 'cancelled'
                        ORDER BY start_time NULLS LAST, end_time NULLS LAST
                    """, (date_str, ids))
                    assignments = cur.fetchall()

    by_signup = {}
    for item in assignments:
        by_signup.setdefault(item["signup_id"], []).append(item)

    def time_text(value):
        if value is None:
            return "?"
        if hasattr(value, "hour"):
            hour, minute = value.hour, value.minute
        else:
            parts = str(value).split(":")
            try:
                hour, minute = int(parts[0]), int(parts[1])
            except (ValueError, IndexError):
                return str(value)
        return f"{(hour - 1) % 12 + 1}:{minute:02d}{'am' if hour < 12 else 'pm'}"

    groups = {}
    seen = set()
    for signup in signups:
        rows = by_signup.get(signup["id"], [])
        if not rows:
            continue
        for row in rows:
            place = str(row.get("assigned_place") or row.get("role") or signup["role"]).strip()
            name = str(signup.get("name") or signup.get("volunteer_id") or "").strip()
            start = time_text(row.get("start_time"))
            end = time_text(row.get("end_time"))
            key = (signup["id"], place, start, end)
            if key in seen:
                continue
            seen.add(key)
            groups.setdefault(place, []).append(f"{name}\n{start}~{end}")

    weekday = "一二三四五六日"[target_date.weekday()]
    lines = ["📣 观音堂新增义工报名通知", "", f"日期：{target_date.day}/{target_date.month}/{target_date.year}（星期{weekday}）", ""]
    if not flag.get("is_published"):
        lines.append("当天值班表尚未发布，暂无发布后新增报名。")
    elif not flag.get("published_at"):
        lines.append("缺少首次发布时间，无法准确判断新增报名。")
    elif not signups:
        lines.append("目前没有新增报名。")
    else:
        lines.append("以下义工已新增报名：")
        for place, people in groups.items():
            lines.extend(["", f"📍 {place}：", *people])
        lines.extend(["", "感恩以上师兄发心护持观音堂！", "🙏🙏🙏"])
    if signups and not groups:
        lines = ["目前没有已完成自动补排的新增报名。"]
    message = "\n".join(lines)
    return render_template_string("""
<!doctype html><html lang="zh"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>新报名通知（手动）</title>
<style>body{font-family:Arial,'Microsoft YaHei',sans-serif;background:#f5f4fa;margin:0;padding:20px}
main{max-width:720px;margin:auto;background:white;padding:22px;border-radius:14px}
h2{color:#6b35b5}textarea{width:100%;box-sizing:border-box;min-height:420px;font-size:17px;line-height:1.6;padding:12px}
button,a{display:inline-block;margin:12px 10px 0 0;padding:12px 18px;border:0;border-radius:9px;text-decoration:none;font-size:16px}
button{background:#8039e8;color:white;cursor:pointer}a{background:#e7e5ed;color:#333}</style></head>
<body><main><h2>📣 新报名通知（手动）</h2>
<p>仅显示首次正式发布后、上次确认发送以来的新报名及实际排班；不会自动发送。</p>
<textarea id="late-notice" readonly>{{ message }}</textarea><br>
{% if has_notice %}
<button type="button" onclick="copyNotice()">📋 复制通知</button>
<form method="post" action="{{ url_for('schedule.schedule_late_signup_notice_confirm') }}" onsubmit="return confirm('确认已经把本次通知发送到 WhatsApp 群？\n确认后本次名单将不再显示，但不会删除报名或排班。');">
<input type="hidden" name="date" value="{{ date_str }}">
<input type="hidden" name="snapshot_at" value="{{ snapshot_at }}">
<button type="submit" style="background:#17864a">✅ 已发送，清空本次名单</button>
</form>
{% endif %}
<a href="{{ url_for('schedule.schedule_notice_center', date=date_str) }}">← 返回通知中心</a>
<script>async function copyNotice(){const box=document.getElementById('late-notice');
try{await navigator.clipboard.writeText(box.value);alert('已复制，可以贴到 WhatsApp 群组');}
catch(e){box.select();document.execCommand('copy');alert('已尝试复制，请检查剪贴板');}}</script>
</main></body></html>
    """, message=message, date_str=date_str, has_notice=bool(groups),
        snapshot_at=flag.get("page_snapshot_at").isoformat() if flag.get("page_snapshot_at") else "")


@schedule_bp.route("/schedule/late_signup_notice/confirm", methods=["POST"])
def schedule_late_signup_notice_confirm():
    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))
    date_str = (request.form.get("date") or "").strip()
    snapshot_text = (request.form.get("snapshot_at") or "").strip()
    try:
        date.fromisoformat(date_str)
        snapshot_at = datetime.fromisoformat(snapshot_text)
        if snapshot_at.tzinfo is None:
            raise ValueError("missing timezone")
    except ValueError:
        return "日期或通知快照格式不正确", 400
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                UPDATE schedule_day_flags
                SET late_notice_sent_at = GREATEST(
                    COALESCE(late_notice_sent_at, published_at),
                    LEAST(%s, now())
                )
                WHERE flag_date = %s AND is_published = true AND published_at IS NOT NULL
            """, (snapshot_at, date_str))
        conn.commit()
    return redirect(url_for("schedule.schedule_late_signup_notice", date=date_str))


@schedule_bp.route("/schedule/notice_center")
def schedule_notice_center():
    
    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    now = malaysia_now()

    if now.hour >= 18:
        default_date = (now + timedelta(days=1)).strftime("%Y-%m-%d")
    else:
        default_date = now.strftime("%Y-%m-%d")

    date_str = request.args.get("date") or default_date
    republish_info = get_schedule_republish_info(date_str)
    return_section = request.args.get("return_section", "")
    if return_section not in ("management", "today"):
        return_section = ""

    # 与「新报名通知（手动）」页面使用同一截止时间和已发送水位。
    # 仅统计仍有效、已安排岗位的新增义工；同一人多个地点只计一位。
    late_notice_count = 0
    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                SELECT is_published, published_at, late_notice_sent_at
                FROM schedule_day_flags WHERE flag_date = %s
            """, (date_str,))
            late_flag = cur.fetchone() or {}
            if late_flag.get("is_published") and late_flag.get("published_at"):
                def _aware_notice_time(value):
                    return value.replace(tzinfo=timezone.utc) if value.tzinfo is None else value

                target_date = date.fromisoformat(date_str)
                cutoff = max(
                    _aware_notice_time(late_flag["published_at"]),
                    _aware_notice_time(late_flag.get("late_notice_sent_at") or late_flag["published_at"]),
                    datetime.combine(target_date - timedelta(days=1), datetime.min.time(),
                                     tzinfo=ZoneInfo("Asia/Kuala_Lumpur")) + timedelta(hours=22),
                )
                cur.execute("""
                    SELECT COUNT(DISTINCT COALESCE(NULLIF(s.volunteer_id, ''), s.id::text)) AS total
                    FROM volunteer_schedule_signups s
                    WHERE s.signup_date = %s
                      AND s.created_at > %s
                      AND COALESCE(s.status, 'pending') <> 'cancelled'
                      AND EXISTS (
                          SELECT 1 FROM volunteer_schedule_assignments a
                          WHERE a.assignment_date = %s AND a.signup_id = s.id
                            AND COALESCE(a.status, 'assigned') <> 'cancelled'
                      )
                """, (date_str, cutoff, date_str))
                late_notice_count = int((cur.fetchone() or {}).get("total") or 0)

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
    font-size:26px;
}
.box{
    max-width:900px;
    margin:auto;
    background:white;
    border-radius:20px;
    padding:30px;
}
.notice-btn{
    display:flex;
    flex-direction:column;
    justify-content:center;
    align-items:center;

    text-align:center;

    font-size:34px;
    font-weight:bold;

    height:110px;
}

.notice-btn small{
    margin-top:8px;
    font-size:20px;
    font-weight:normal;
    opacity:0.9;
}
.notice-btn{
    display:block;
    width:100%;
    text-align:center;
    text-decoration:none;
    color:white;
    border-radius:18px;
    padding:28px;
    margin:18px 0;
    font-size:32px;
    font-weight:bold;
    box-sizing:border-box;
}
.blue{background:#2196F3;}
.green{background:#25D366;}
.orange{background:#FF9800;}
.gray{background:#607D8B;}
.purple{background:#7136b7;}
</style>
</head>
<body>
<div class="box">

<h1>📢 通知中心</h1>

<p style="font-size:22px;color:#666;">
负责人每天只需依照以下顺序发送 WhatsApp 信息。
</p>
                                  
{% if republish_info is defined and republish_info.need_republish %}
<div style="
background:#fff8dc;
border:3px solid orange;
border-radius:15px;
padding:18px;
margin-bottom:20px;
font-size:24px;
line-height:1.8;
">

🟠 <b>正式值班表已有更新</b><br>

自上次发布正式值班表后，已有义工报名、修改报名、取消报名或负责人调整排班。

<br><br>

<b>请重新发送最新正式值班表 WhatsApp。</b>

<br><br>

<form method="post" action="/schedule/copy_whatsapp">
    <input type="hidden" name="date" value="{{ date_str }}">

    <button class="notice-btn green" type="submit">
        📋 复制最新正式值班表 WhatsApp<br>
        <small>📤 可随时重新发送最新版</small>
    </button>
</form>

{% if republish_info.last_schedule_change_text %}
<br>
<small>
    🕒 最后更新：{{ republish_info.last_schedule_change_text }}
</small>
{% endif %}

</div>
{% endif %}

<a class="notice-btn blue"
   href="/schedule/signup_notice?date={{ date_str }}">

    📢 报名通知<br>
    <small>🕖 晚上 7:00 发群</small>

</a>

<a class="notice-btn purple"
   href="/schedule/late_signup_notice?date={{ date_str }}">
    📣 新报名通知（手动）<br>
    {% if late_notice_count > 0 %}
    <small style="color:#fff5a6;font-weight:bold;">🔴 有 {{ late_notice_count }} 位新报名 · 待发送</small><br>
    <small>📋 点击查看并复制通知</small>
    {% else %}
    <small>✅ 目前没有新报名</small>
    {% endif %}
</a>

<a class="notice-btn orange"
   href="{{ url_for('schedule.generate_shortage_notice', date=date_str, return_section=return_section) }}">

    📣 缺义工通知<br>
    <small>📌 缺人时发送</small>

</a>

<a class="notice-btn gray" href="{{ '/schedule?override_date=' ~ date_str ~ '#management-functions' if return_section == 'management' else ('/schedule?mode=day&amp;override_date=' ~ date_str ~ '#today-operations' if return_section == 'today' else '/schedule?override_date=' ~ date_str) }}">
   ⬅ 返回负责人首页
</a>

</div>
</body>
</html>
""", date_str=date_str, republish_info=republish_info,
        late_notice_count=late_notice_count, return_section=return_section)


@schedule_bp.route("/schedule/signup_notice")
def schedule_signup_notice():

    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    now = malaysia_now()

    if now.hour >= 18:
        default_date = (now + timedelta(days=1)).strftime("%Y-%m-%d")
    else:
        default_date = now.strftime("%Y-%m-%d")

    # 优先使用负责人页面所选日期；不能每次都强制改为今日/明日。
    selected_date = (request.args.get("date") or "").strip()
    try:
        datetime.strptime(selected_date, "%Y-%m-%d")
        date_str = selected_date
    except ValueError:
        date_str = default_date

    text = build_signup_notice_whatsapp(date_str)

    return render_template_string("""
    <!doctype html>
    <html lang="zh">
    <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>今晚义工报名</title>

    <link rel="stylesheet"
        href="{{ url_for('static', filename='css/toolbox.css') }}">

    <style>

    .notice-textarea{
        width:100%;
        min-height:520px;
        font-size:19px;
        line-height:1.6;
        padding:18px;
        border:1px solid #d1d5db;
        border-radius:16px;
        box-sizing:border-box;
        resize:vertical;
    }

    .notice-actions{
        display:grid;
        grid-template-columns:1fr 1fr;
        gap:14px;
        margin-top:20px;
    }

    @media(max-width:700px){

        .notice-actions{
            grid-template-columns:1fr;
        }

    }

    </style>

    </head>
    <body>

    <div class="page">

        <h1 class="page-title">
            📢 今晚义工报名
        </h1>

        <p class="page-subtitle">
            WhatsApp 报名通知，可直接复制发送
        </p>

        <div class="card">

            <div class="section-title">
                📄 通知内容
            </div>

            <textarea
                id="noticeText"
                class="notice-textarea"
                readonly>{{ text }}</textarea>

            <div class="notice-actions">

                <button
                    type="button"
                    class="btn-tool btn-primary"
                    onclick="copyText()">

                    📋 复制 WhatsApp

                </button>

                <a
                    class="btn-tool btn-secondary"
                    href="/schedule/notice_center?date={{ date_str }}">

                    ← 返回通知中心

                </a>

            </div>

        </div>

    </div>

    <script>

    function copyText(){

        const ta = document.getElementById("noticeText");

        ta.focus();
        ta.select();
        ta.setSelectionRange(0,999999);

        const ok = document.execCommand("copy");

        if(ok){
            alert("✅ 已复制，可以贴去 WhatsApp 群");
        }else{
            alert("❌ 复制失败，请手动复制");
        }

    }

    </script>

    </body>
    </html>

    """, text=text, date_str=date_str)


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

    sync_schedule_after_signup_change(
        signup_id,
        action="upsert",
        changed_by="admin"
    )

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

    today = malaysia_today()

    date_str = request.args.get(
        "date",
        today.strftime("%Y-%m-%d")
    )

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

            # 签到情况只展示实际岗位地点；橙班／黄班仅供后台排班使用。
            place_text = str(place).strip()
            if role == "卫生" and place_text == "卫生":
                place_text = "卫生"

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

    # 固定佛台由 Excel 轮值和负责人当天覆盖名单决定，不一定有普通排班 ID。
    # 旧签到记录即使 walk_in=True，仍按当天有效名单在页面上归入正式签到。
    fixed_altar_names = {
        "".join(cc.convert(str(name or "")).split()).casefold()
        for name in get_fixed_buddha_for_date(date_str)
    }

    def is_fixed_altar_checkin(log):
        if str(log.get("actual_role") or "").strip() != "佛台":
            return False
        name = "".join(cc.convert(str(log.get("name") or "")).split()).casefold()
        return bool(name) and name in fixed_altar_names

    fixed_altar_logs = []
    for l in logs:
        if is_fixed_altar_checkin(l):
            fixed_altar_logs.append(l)
        elif l.get("walk_in"):
            walk_ins_raw.append(l)

        if not l.get("check_out_time"):
            not_checked_out_raw.append(l)

    checked_in = group_by_person(checked_in_raw, source="assignment")
    for person in group_by_person(fixed_altar_logs, source="log"):
        key = person.get("volunteer_id") or person.get("name")
        existing = next((p for p in checked_in if (p.get("volunteer_id") or p.get("name")) == key), None)
        if existing:
            for field, value in (("role_text", "佛台"), ("place_text", "佛台")):
                parts = [x.strip() for x in existing[field].split(" / ") if x.strip()]
                if value not in parts:
                    parts.append(value)
                    existing[field] = " / ".join(parts)
        else:
            checked_in.append(person)

    not_checked_in = group_by_person(not_checked_in_raw, source="assignment")
    walk_ins = group_by_person(walk_ins_raw, source="log")
    not_checked_out = group_by_person(not_checked_out_raw, source="log")

    return render_template_string("""
    <!doctype html>
    <html>
    <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>签到情况</title>

    <link rel="stylesheet"
          href="{{ url_for('static', filename='css/toolbox.css') }}">

    <style>
    .attendance-wrap{
        max-width:1100px;
        margin:auto;
        padding:24px;
    }

    .attendance-list{
        list-style:none;
        padding:0;
        margin:0;
    }

    .attendance-list li{
        font-size:26px;
        padding:14px 0;
        border-bottom:1px solid #e5e7eb;
    }

    .attendance-list li:last-child{
        border-bottom:none;
    }

    .date-query-row{
        display:grid;
        grid-template-columns:260px 180px;
        gap:14px;
        align-items:end;
    }
                                  
    .date-query-row .form-group{
        margin-bottom:0;
    }

    .date-query-row .btn-tool{
        height:68px;
        min-height:68px;
        font-size:26px;
        align-self:end;
    }

    @media (max-width:700px){
        .date-query-row{
            grid-template-columns:1fr;
        }
    }
    </style>
    </head>

    <body>

    <div class="attendance-wrap">

        <div class="card">

            <h1 class="page-title">
                📋 签到情况
            </h1>

            <p class="page-subtitle">
                日期：{{ date_str }}
            </p>

            <form method="get">

                <div class="date-query-row">

                    <div class="form-group">
                        <label class="form-label">查询日期</label>

                        <input
                            class="form-input"
                            type="date"
                            name="date"
                            value="{{ date_str }}">
                    </div>

                    <button
                        class="btn-tool btn-blue"
                        type="submit">
                          查询
                    </button>

                </div>

            </form>

            <div class="btn-row">
                <a class="btn-tool btn-gray"
                href="/schedule/admin">
                    ⬅ 返回负责人首页
                </a>
            </div>

        </div>

        <div class="summary-grid">

            <div class="summary-box">
                <div class="summary-title">✅ 已签到</div>
                <div class="summary-value">{{ checked_in|length }}</div>
            </div>

            <div class="summary-box">
                <div class="summary-title">❌ 未签到</div>
                <div class="summary-value">{{ not_checked_in|length }}</div>
            </div>

            <div class="summary-box">
                <div class="summary-title">🚶 临时报到</div>
                <div class="summary-value">{{ walk_ins|length }}</div>
            </div>

            <div class="summary-box">
                <div class="summary-title">⏳ 未签退</div>
                <div class="summary-value">{{ not_checked_out|length }}</div>
            </div>

        </div>

        <div class="card">
            <h2 class="section-title">✅ 已签到：{{ checked_in|length }}</h2>

            <ul class="attendance-list">
            {% for r in checked_in %}
                <li>{{ r.name }}｜{{ r.role_text }}｜{{ r.place_text }}</li>
            {% else %}
                <li class="text-gray">暂无记录</li>
            {% endfor %}
            </ul>
        </div>

        <div class="card">
            <h2 class="section-title">❌ 已排班未签到：{{ not_checked_in|length }}</h2>

            <ul class="attendance-list">
            {% for r in not_checked_in %}
                <li>{{ r.name }}｜{{ r.role_text }}｜{{ r.place_text }}</li>
            {% else %}
                <li class="text-gray">暂无记录</li>
            {% endfor %}
            </ul>
        </div>

        <div class="card">
            <h2 class="section-title">🚶 临时报到：{{ walk_ins|length }}</h2>

            <ul class="attendance-list">
            {% for r in walk_ins %}
                <li>{{ r.name }}｜{{ r.role_text }}｜{{ r.place_text }}</li>
            {% else %}
                <li class="text-gray">暂无记录</li>
            {% endfor %}
            </ul>
        </div>

        <div class="card">
            <h2 class="section-title">⏳ 未签退：{{ not_checked_out|length }}</h2>

            <ul class="attendance-list">
            {% for r in not_checked_out %}
                <li>{{ r.name }}｜{{ r.role_text }}｜{{ r.place_text }}</li>
            {% else %}
                <li class="text-gray">暂无记录</li>
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
    daily_quote = get_daily_dharma()

    return render_template_string("""
    <!doctype html>
    <html>
    <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>排班设置</title>

    <link rel="stylesheet"
          href="{{ url_for('static', filename='css/toolbox.css') }}">

    <style>
    .settings-wrap{
        max-width:900px;
        margin:auto;
        padding:24px;
    }

    .settings-grid{
        display:grid;
        grid-template-columns:1fr 1fr;
        gap:18px;
    }

    @media(max-width:700px){
        .settings-grid{
            grid-template-columns:1fr;
        }
    }
    </style>
    </head>

    <body>

    <div class="settings-wrap">

        <div class="card">

            <h1 class="page-title">⚙️ 排班系统设置</h1>

            <p class="page-subtitle">
                管理负责人页面默认日期、缺义工目标人数与供台提醒。
            </p>

            <div class="quote-card" style="margin-bottom:22px;">
                <div class="quote-title">
                    🌸 今日佛言佛语
                </div>

                <div class="quote-content">
                    {{ daily_quote }}
                </div>
            </div>

            {% if saved %}
            <div class="alert alert-success">
                ✅ 设置已保存
            </div>
            {% endif %}

            <form method="post">

                <h2 class="section-title">📅 默认工作日期</h2>

                <div class="form-group">
                    <label class="form-label">
                        几点后默认看明天？
                    </label>

                    <input
                        class="form-input"
                        name="default_day_switch_time"
                        value="{{ settings.default_day_switch_time }}"
                        placeholder="18:00">

                    <div class="form-help">
                        例：18:00 = 晚上 6 点后，负责人首页默认切换到明天。
                    </div>
                </div>

                <div class="divider"></div>

                <h2 class="section-title">👥 缺义工目标人数</h2>

                <div class="settings-grid">

                    <div class="form-group">
                        <label class="form-label">橙班 - 观音堂</label>
                        <input
                            class="form-input"
                            name="target_orange_guanyintang"
                            value="{{ settings.target_orange_guanyintang }}">
                    </div>

                    <div class="form-group">
                        <label class="form-label">橙班 - 活动中心</label>
                        <input
                            class="form-input"
                            name="target_orange_activity"
                            value="{{ settings.target_orange_activity }}">
                    </div>

                    <div class="form-group">
                        <label class="form-label">黄班 - 观音堂</label>
                        <input
                            class="form-input"
                            name="target_yellow_guanyintang"
                            value="{{ settings.target_yellow_guanyintang }}">
                    </div>

                    <div class="form-group">
                        <label class="form-label">黄班 - 活动中心</label>
                        <input
                            class="form-input"
                            name="target_yellow_activity"
                            value="{{ settings.target_yellow_activity }}">
                    </div>

                    <div class="form-group">
                        <label class="form-label">卫生人数</label>
                        <input
                            class="form-input"
                            name="target_cleaning"
                            value="{{ settings.target_cleaning }}">
                    </div>

                </div>

                <div class="divider"></div>

                <h2 class="section-title">🪔 供台提醒</h2>

                <div class="form-group">
                    <label class="form-label">
                        提醒未来几天的大日子？
                    </label>

                    <input
                        class="form-input"
                        name="supply_alert_days"
                        value="{{ settings.supply_alert_days }}">

                    <div class="form-help">
                        系统会提前显示未来几天需要设供台的大日子。
                    </div>
                </div>

                <div class="btn-row">

                    <button
                        class="btn-tool btn-green btn-full"
                        type="submit">
                        保存排班设置
                    </button>

                    <a
                        class="btn-tool btn-gray btn-full"
                        href="{{ url_for('schedule.schedule_admin') }}">
                        返回负责人首页
                    </a>

                </div>

            </form>

        </div>

    </div>

    </body>
    </html>
    """, settings=settings, saved=saved, daily_quote=daily_quote)


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

    today_date = malaysia_today()

    today = today_date.strftime("%Y-%m-%d")
    tomorrow = (today_date + timedelta(days=1)).strftime("%Y-%m-%d")

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

    <link rel="stylesheet"
          href="{{ url_for('static', filename='css/toolbox.css') }}">

    <style>
    .signup-admin-wrap{
        max-width:1280px;
        margin:auto;
        padding:24px;
    }

    .filter-grid{
        display:grid;
        grid-template-columns:240px 1fr 1fr 180px;
        gap:16px;
        align-items:end;
    }

    .quick-date-row{
        display:flex;
        flex-wrap:wrap;
        gap:12px;
        margin-top:16px;
        align-items:center;
    }

    .role-stat-row{
        display:flex;
        flex-wrap:wrap;
        gap:12px;
        margin-top:12px;
    }

    .role-stat-pill{
        background:#fff7ed;
        color:#9a3412;
        border:1px solid #fed7aa;
        border-radius:999px;
        padding:10px 18px;
        font-size:22px;
        font-weight:bold;
    }

    .status-pending{
        background:#fff8d6;
    }

    .status-assigned{
        background:#e8fff0;
    }

    .status-cancelled{
        background:#f3f4f6;
        color:#777;
    }

    .signup-table{
        width:100%;
        border-collapse:collapse;
        font-size:22px;
        background:white;
    }

    .signup-table th{
        background:#f8fafc;
        font-size:22px;
        padding:14px 10px;
        border-bottom:2px solid #e5e7eb;
        white-space:nowrap;
    }

    .signup-table td{
        padding:14px 10px;
        border-bottom:1px solid #e5e7eb;
        text-align:center;
        vertical-align:top;
    }

    .signup-table tr:hover{
        filter:brightness(.98);
    }

    .assignment-text{
        font-size:20px;
        line-height:1.6;
    }

    .remark-text{
        font-size:20px;
        color:#6b7280;
        max-width:200px;
    }

    .op-stack{
        display:flex;
        flex-direction:column;
        gap:8px;
        align-items:center;
    }

    .op-stack form{
        margin:0;
    }

    .op-btn{
        min-height:42px;
        padding:0 14px;
        font-size:18px;
        border-radius:12px;
        width:110px;
    }

    .floating-back{
        position:fixed;
        right:25px;
        bottom:25px;
        z-index:9999;
        min-width:170px;
    }

    .history-check{
        display:flex;
        gap:8px;
        align-items:center;
        font-size:22px;
        font-weight:bold;
        color:#374151;
    }

    .history-check input{
        width:22px;
        height:22px;
    }

    @media (max-width:900px){
        .filter-grid{
            grid-template-columns:1fr;
        }

        .signup-table{
            font-size:20px;
        }

        .signup-table th,
        .signup-table td{
            padding:10px 8px;
        }

        .floating-back{
            position:static;
            margin:20px auto;
            width:100%;
        }
    }
    </style>
    </head>

    <body>

    <div class="signup-admin-wrap">

        <div class="card">

            <h1 class="page-title">📋 义工报名管理</h1>

            <p class="page-subtitle">
                查看、筛选、修改、取消或恢复义工报名。
            </p>

            <div class="btn-row">
                <a class="btn-tool btn-gray"
                href="/schedule/admin">
                    ⬅ 返回负责人首页
                </a>
            </div>

        </div>

        <div class="summary-grid">

            <div class="summary-box">
                <div class="summary-title">当前查询</div>
                <div class="summary-value">{{ summary.total }}</div>
            </div>

            <div class="summary-box">
                <div class="summary-title">🟡 未安排</div>
                <div class="summary-value">{{ summary.pending }}</div>
            </div>

            <div class="summary-box">
                <div class="summary-title">🟢 已安排</div>
                <div class="summary-value">{{ summary.assigned }}</div>
            </div>

            <div class="summary-box">
                <div class="summary-title">⚪ 已取消</div>
                <div class="summary-value">{{ summary.cancelled }}</div>
            </div>

        </div>

        <div class="card">

            <h2 class="section-title">📊 岗位统计</h2>

            {% if role_summary %}
                <div class="role-stat-row">
                {% for role, count in role_summary.items() %}
                    <div class="role-stat-pill">
                        {{ role }}：{{ count }} 人
                    </div>
                {% endfor %}
                </div>
            {% else %}
                <div class="empty-state">
                    <div class="empty-icon">📭</div>
                    <div class="empty-title">没有岗位统计</div>
                </div>
            {% endif %}

        </div>

        <div class="card">

            <h2 class="section-title">🔍 筛选查询</h2>

            <form method="get">

                <div class="filter-grid">

                    <div class="form-group">
                        <label class="form-label">日期</label>
                        <input
                            class="form-input"
                            type="date"
                            name="date"
                            value="{{ date_filter }}">
                    </div>

                    <div class="form-group">
                        <label class="form-label">状态</label>
                        <select class="form-select" name="status">
                            <option value="all" {% if status=="all" %}selected{% endif %}>全部状态</option>
                            <option value="pending" {% if status=="pending" %}selected{% endif %}>未安排</option>
                            <option value="assigned" {% if status=="assigned" %}selected{% endif %}>已安排</option>
                            <option value="cancelled" {% if status=="cancelled" %}selected{% endif %}>已取消</option>
                        </select>
                    </div>

                    <div class="form-group">
                        <label class="form-label">岗位</label>
                        <select class="form-select" name="role">
                            <option value="all" {% if role_filter=="all" %}selected{% endif %}>全部岗位</option>

                            {% for role in roles %}
                            <option value="{{ role }}" {% if role_filter==role %}selected{% endif %}>
                                {{ role }}
                            </option>
                            {% endfor %}
                        </select>
                    </div>

                    <div class="form-group">
                        <label class="form-label">&nbsp;</label>
                        <button class="btn-tool btn-blue btn-full" type="submit">
                            🔍 查询
                        </button>
                    </div>

                </div>

                <div class="quick-date-row">

                    <a class="btn-tool btn-orange"
                    href="/schedule/signups?date={{ today }}&status={{ status }}&role={{ role_filter }}">
                        今天
                    </a>

                    <a class="btn-tool btn-purple"
                    href="/schedule/signups?date={{ tomorrow }}&status={{ status }}&role={{ role_filter }}">
                        明天
                    </a>

                    <a class="btn-tool btn-gray"
                    href="/schedule/signups">
                        全部
                    </a>

                    <label class="history-check">
                        <input
                            type="checkbox"
                            name="show_history"
                            value="1"
                            {% if show_history %}checked{% endif %}>
                        显示过去记录
                    </label>

                </div>

            </form>

        </div>

        <div class="card">

            <h2 class="section-title">📄 报名记录</h2>

            <div class="table-wrap">

                <table class="signup-table">

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
                    <tr class="status-{{ r.status }}">
                        <td>{{ r.signup_date }}</td>
                        <td>{{ r.volunteer_id }}</td>
                        <td><b>{{ r.name }}</b></td>
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
                                <span class="badge badge-orange">未安排</span>
                            {% elif r.status == "assigned" %}
                                <span class="badge badge-green">已安排</span>
                            {% elif r.status == "cancelled" %}
                                <span class="badge badge-gray">已取消</span>
                            {% else %}
                                <span class="badge badge-blue">{{ r.status }}</span>
                            {% endif %}
                        </td>

                        <td class="assignment-text">
                            {% if r.final_assignment %}
                                {% for item in r.final_assignment %}
                                    <div>{{ item|safe }}</div>
                                {% endfor %}
                            {% else %}
                                -
                            {% endif %}
                        </td>

                        <td class="remark-text">
                            {{ r.remarks or "" }}
                        </td>

                        <td>
                            <div class="op-stack">

                                <form method="get" action="/schedule/signup_edit/{{ r.id }}">
                                    <button class="btn-tool btn-blue op-btn" type="submit">
                                        ✏️ 修改
                                    </button>
                                </form>

                                {% if r.status == "assigned" %}
                                <form method="get" action="/schedule/signup_place/{{ r.id }}">
                                    <button class="btn-tool btn-purple op-btn" type="submit">
                                        🔧 调整
                                    </button>
                                </form>
                                {% endif %}

                                {% if r.status != "cancelled" %}
                                <form
                                    method="post"
                                    action="/schedule/signup_cancel/{{ r.id }}"
                                    onsubmit="return confirm('确定取消这位义工的报名？');">
                                    <button class="btn-tool btn-red op-btn" type="submit">
                                        ❌ 取消
                                    </button>
                                </form>

                                {% else %}

                                <form
                                    method="post"
                                    action="/schedule/signup_restore/{{ r.id }}"
                                    onsubmit="return confirm('确定恢复报名？');">
                                    <button class="btn-tool btn-green op-btn" type="submit">
                                        ↩️ 恢复
                                    </button>
                                </form>

                                {% endif %}

                            </div>
                        </td>
                    </tr>
                    {% endfor %}

                </table>

            </div>

            {% if not rows %}
                <div class="empty-state">
                    <div class="empty-icon">📭</div>
                    <div class="empty-title">没有记录</div>
                    <div class="empty-text">请调整筛选条件后再查询。</div>
                </div>
            {% endif %}

        </div>

        <a href="/schedule/admin"
        class="btn-tool btn-blue floating-back">
            ⬅ 返回首页
        </a>

    </div>

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
        tomorrow=tomorrow,
        show_history=show_history,
    )


LOGIN_HTML = """
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>负责人排班系统</title>

<link rel="stylesheet"
      href="{{ url_for('static', filename='css/toolbox.css') }}">

<style>
.login-wrap{
    max-width:620px;
    margin:60px auto;
    padding:20px;
}

.login-card{
    text-align:center;
}

.login-icon{
    font-size:64px;
    margin-bottom:10px;
}

.login-title{
    font-size:52px;
    font-weight:bold;
    color:#1e293b;
}

.login-subtitle{
    margin:15px 0 35px;
    color:#666;
    font-size:22px;
    line-height:1.8;
}

.login-input{
    margin-top:10px;
}

.login-input input{
    text-align:center;
    letter-spacing:6px;
    font-size:34px;
}

@media (max-width:700px){

    .login-title{
        font-size:40px;
    }

    .login-subtitle{
        font-size:20px;
    }

}
</style>

<script>
window.addEventListener("DOMContentLoaded",function(){

    document.querySelector("input[name='pin']").focus();

});
</script>

</head>

<body>

<div class="login-wrap">

    <div class="card login-card">

        <div class="login-icon">
            📅
        </div>

        <div class="login-title">
            负责人排班系统
        </div>

        <div class="login-subtitle">
            请输入负责人 PIN<br>
            进入排班管理系统。
        </div>

        <form method="post">

            <div class="form-group">

                <label class="form-label">
                    负责人 PIN
                </label>

                <div class="login-input">

                    <input
                        class="form-input"
                        type="password"
                        name="pin"
                        placeholder="请输入负责人 PIN"
                        inputmode="numeric"
                        autocomplete="new-password"
                        autocorrect="off"
                        autocapitalize="off"
                        spellcheck="false"
                        readonly
                        onfocus="this.removeAttribute('readonly');">

                </div>

            </div>

            <div class="btn-row">

                <button
                    class="btn-tool btn-green btn-full"
                    type="submit">

                    🔓 进入负责人系统

                </button>

            </div>

        </form>

    </div>

</div>

</body>
</html>
"""




# 固定岗位人员管理：独立页面，不影响现有排班、佛台及通知功能。
_FIXED_ROLE_OPTIONS = ("供花", "供果", "佛学班")

_FIXED_ROLES_HTML = """
<!doctype html><html lang="zh"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>固定岗位人员管理</title>
<style>
body{font-family:"Microsoft YaHei",Arial,sans-serif;background:#f5f7fa;margin:0;padding:20px;color:#243047}
.wrap{max-width:900px;margin:auto}.card{background:white;padding:22px;margin:16px 0;border-radius:14px;box-shadow:0 2px 10px #0001}
h1{font-size:26px}h2{font-size:21px}input,select,button{font-size:17px;padding:12px;border-radius:8px;border:1px solid #ccd3df}
button{background:#187c71;color:white;border:0;cursor:pointer}.danger{background:#ba3e43;padding:7px 12px}
.row{display:flex;flex-wrap:wrap;gap:12px;align-items:center}table{width:100%;border-collapse:collapse}td,th{padding:11px;border-bottom:1px solid #e5e9ef;text-align:left}
.notice{padding:12px;background:#e6f4ed;border-radius:8px}.error{background:#ffe7e7}a{color:#187c71}
</style></head><body><div class="wrap">
<a href="/schedule/admin">← 返回负责人首页</a>
<h1>👥 固定岗位人员管理</h1>
<p>只管理供花、供果及佛学班固定成员；不会修改排班或签到记录。</p>
{% with messages = get_flashed_messages(with_categories=true) %}{% for category,msg in messages %}
<div class="notice {{ 'error' if category == 'error' else '' }}">{{ msg }}</div>
{% endfor %}{% endwith %}
<div class="card"><h2>➕ 加入固定成员</h2>
<form method="post" class="row"><input type="hidden" name="action" value="add">
<label>岗位 <select name="role" required>{% for role in roles %}<option value="{{ role }}">{{ role }}</option>{% endfor %}</select></label>
<label>义工编号 <input name="volunteer_id" placeholder="CHE-208 / STW-xxx / 8xx" required autocomplete="off"></label>
<button type="submit">查询并加入</button></form>
<p style="color:#667">以 Supabase volunteers 的正式编号为准；若编号无匹配，将拒绝加入。</p></div>
{% for role in roles %}
<div class="card"><h2>{{ '🌸' if role == '供花' else ('🍎' if role == '供果' else '📖') }} {{ role }}（{{ members[role]|length }} 人）</h2>
{% if members[role] %}<table><thead><tr><th>编号</th><th>姓名</th><th>操作</th></tr></thead><tbody>
{% for member in members[role] %}<tr><td>{{ member.volunteer_id }}</td><td>{{ member.name }}</td>
<td><form method="post" onsubmit="return confirm('确定从{{ role }}移除 {{ member.name }}？');">
<input type="hidden" name="action" value="remove"><input type="hidden" name="role" value="{{ role }}">
<input type="hidden" name="volunteer_id" value="{{ member.volunteer_id }}"><button class="danger">移除</button></form></td></tr>{% endfor %}
</tbody></table>{% else %}<p>尚无固定成员。</p>{% endif %}</div>
{% endfor %}</div></body></html>
"""


@schedule_bp.route("/schedule/fixed-roles", methods=["GET", "POST"])
def schedule_fixed_roles():
    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    if request.method == "POST":
        action = request.form.get("action", "").strip()
        role = request.form.get("role", "").strip()
        raw_id = request.form.get("volunteer_id", "").strip().upper()
        if action not in ("add", "remove") or role not in _FIXED_ROLE_OPTIONS or not raw_id:
            flash("操作资料不正确。", "error")
            return redirect(url_for("schedule.schedule_fixed_roles"))
        try:
            with get_conn() as conn:
                with conn.cursor(cursor_factory=RealDictCursor) as cur:
                    if action == "add":
                        # 数字 CHE 编号与 8xx 编号均支持；禁止模糊匹配姓名。
                        candidates = [raw_id]
                        if raw_id.isdigit() and not raw_id.startswith("8"):
                            candidates.append("CHE-" + raw_id)
                        cur.execute("""
                            SELECT id, name FROM volunteers
                            WHERE id = ANY(%s) AND status = '在册'
                            ORDER BY CASE WHEN id=%s THEN 0 ELSE 1 END LIMIT 1
                        """, (candidates, raw_id))
                        volunteer = cur.fetchone()
                        if not volunteer:
                            flash("找不到在册义工编号，请检查后重试。", "error")
                            return redirect(url_for("schedule.schedule_fixed_roles"))
                        cur.execute("""
                            INSERT INTO volunteer_fixed_roles (volunteer_id, role)
                            VALUES (%s, %s) ON CONFLICT (volunteer_id, role) DO NOTHING
                        """, (volunteer["id"], role))
                        flash(f"已保存：{volunteer['name']}（{volunteer['id']}）→ {role}", "ok")
                    else:
                        cur.execute("""
                            DELETE FROM volunteer_fixed_roles
                            WHERE volunteer_id=%s AND role=%s
                        """, (raw_id, role))
                        flash("已移除固定岗位成员。", "ok")
        except Exception:
            flash("保存失败：请先确认 Supabase 已建立 volunteer_fixed_roles 数据表。", "error")
        return redirect(url_for("schedule.schedule_fixed_roles"))

    try:
        rows = db_query("""
            SELECT f.volunteer_id, f.role, COALESCE(v.name, '义工资料已不存在') AS name
            FROM volunteer_fixed_roles f
            LEFT JOIN volunteers v ON v.id=f.volunteer_id
            WHERE f.role = ANY(%s)
            ORDER BY f.role, v.name, f.volunteer_id
        """, (list(_FIXED_ROLE_OPTIONS),), fetchall=True) or []
    except Exception:
        rows = []
        flash("读取名单失败：请先执行 Supabase 建表 SQL。", "error")
    members = {role: [] for role in _FIXED_ROLE_OPTIONS}
    for row in rows:
        members[row["role"]].append(row)
    return render_template_string(_FIXED_ROLES_HTML, roles=_FIXED_ROLE_OPTIONS, members=members)

SCHEDULE_HTML = """
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<link rel="manifest"
      href="/schedule-manifest.json?v=5">

<link rel="icon"
      type="image/png"
      sizes="192x192"
      href="/static/schedule_icon_192.png?v=5">

<link rel="apple-touch-icon"
      href="/static/schedule_icon_512.png?v=5">

<meta name="theme-color"
      content="#43a047">

<title>负责人排班系统</title>

<link rel="stylesheet"
      href="{{ url_for('static', filename='css/toolbox.css') }}">

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
function openMasterDay(day) {
    const valid = Array.from(document.querySelectorAll('#master_day_select option')).some(o => o.value === day);
    if (!valid) { alert('请选择列表中的大日子'); return; }
    window.location.href = '/schedule?mode=special&override_date=' + encodeURIComponent(day) + '&panel=master-table';
}
document.addEventListener('DOMContentLoaded', function () {
    const select = document.getElementById('master_day_select');
    const calendar = document.getElementById('master_calendar');
    if (select) select.addEventListener('change', function () { openMasterDay(this.value); });
    if (calendar) calendar.addEventListener('change', function () {
        if (select && Array.from(select.options).some(o => o.value === this.value)) openMasterDay(this.value);
        else { alert('请选择未来90天内的大日子'); calendar.value = select ? select.value : ''; }
    });
    const removeCalendar = document.getElementById('remove_day_calendar');
    if (removeCalendar) removeCalendar.addEventListener('change', function () {
        const params = new URLSearchParams(window.location.search);
        params.set('mode', 'special');
        params.set('panel', 'master-table');
        params.set('remove_date', this.value);
        window.location.href = '/schedule?' + params.toString();
    });
    const panel = new URLSearchParams(window.location.search).get('panel');
    if (panel === 'master-table' && select && select.value !== new URLSearchParams(window.location.search).get('override_date')) { openMasterDay(select.value); return; }
    if (panel === 'master-table' || panel === 'extra-buddha' || panel === 'buddha-leave') showSpecial(panel);
});
function showSpecial(id) {
    if (id === 'master-table') {
        const select = document.getElementById('master_day_select');
        const shown = new URLSearchParams(window.location.search).get('override_date');
        if (select && select.value !== shown) { openMasterDay(select.value); return; }
    }
    const boxes = document.querySelectorAll(".special-box");

    boxes.forEach(function(box) {
        box.style.display = "none";
    });

    const target = document.getElementById(id);

    if (target) {
        target.style.display = "block";
        setTimeout(function () {
            target.scrollIntoView({ behavior: "smooth", block: "start" });
        }, 80);
    }
}
</script>

<style>
.login-wrap{
    max-width:620px;
    margin:60px auto;
    padding:20px;
}

.login-card{
    text-align:center;
}

.login-icon{
    font-size:64px;
    margin-bottom:10px;
}

.login-title{
    font-size:52px;
    font-weight:bold;
    color:#1e293b;
}

.login-subtitle{
    margin:15px 0 35px;
    color:#666;
    font-size:22px;
    line-height:1.8;
}

.login-input{
    margin-top:10px;
}

.login-input input{
    text-align:center;
    letter-spacing:6px;
    font-size:34px;
}

html {
    scroll-behavior: smooth;
}

@media (max-width:700px){
    .login-title{
        font-size:40px;
    }

    .login-subtitle{
        font-size:20px;
    }
}
/* 每天必用入口：醒目金黄色，避免与其他管理按钮混淆 */
.daily-entry-highlight { display:block; text-align:center; background:linear-gradient(135deg,#ffbf24,#f17800) !important; color:#172235 !important; border:3px solid #e18a00 !important; font-weight:900 !important; font-size:clamp(21px,3vw,30px) !important; padding:20px 14px !important; border-radius:15px !important; box-shadow:0 5px 0 #b76a00,0 9px 18px rgba(173,99,0,.23) !important; text-decoration:none; }
.daily-entry-highlight:hover { filter:brightness(1.06); transform:translateY(-2px); }
/* 滚动时固定在右下角，手机和电脑均可使用 */
.floating-home { position:fixed; right:20px; bottom:22px; z-index:9999; background:#245de6; color:white !important; padding:11px 16px; border-radius:999px; font-weight:800; font-size:15px; box-shadow:0 4px 15px rgba(0,0,0,.28); text-decoration:none; border:2px solid white; }
.floating-home:hover { background:#1247b9; }
@media(max-width:600px) { .floating-home { right:12px; bottom:12px; font-size:14px; padding:10px 13px; } }
</style>

</head>

<body>
{% if mode in ("day", "signup", "special", "controls") %}<a class="floating-home" href="/schedule" aria-label="返回负责人首页">🏠 返回首页</a>{% endif %}
<div class="box">

<div class="top-bar">
    <h1>📅 负责人排班系统</h1>
    {% if mode in ("day", "signup", "special", "controls") %}
    <a class="btn-tool btn-blue" style="width:auto" href="/schedule">🏠 负责人首页</a>
    {% endif %}

    <a class="btn-tool btn-gray" style="width:auto;min-width:160px;" href="/schedule/logout">
        🚪 退出登录
    </a>
</div>

{% if mode not in ("day", "signup", "special", "controls") %}
<div class="card">
 <h2 class="section-title">📊 今日报名概况 · {{ override_date }}</h2>
 <div class="action-grid">
  <div class="date-card">📋 已报名：<b>{{ day_summary.total }}</b></div>
  <div class="date-card">🟠 待安排：<b>{{ day_summary.pending }}</b></div>
  <div class="date-card">🟢 已安排：<b>{{ day_summary.assigned }}</b></div>
  <div class="date-card">⚪ 已取消：<b>{{ day_summary.cancelled }}</b></div>
 </div>
 <a class="btn-tool daily-entry-highlight" href="/schedule?mode=day&override_date={{ override_date }}">⭐ 进入今日安排 → <span style="font-size:.7em">（每日必用）</span></a>
</div>

<div class="card" id="management-functions">
    <h2 class="section-title">🌸 管理功能</h2>
    <div class="main-menu">
      <a class="btn-tool btn-green" style="background:#159947;border-color:#159947;color:#fff" href="/schedule?mode=signup&override_date={{ override_date }}">👥 报名管理（代报名 / WhatsApp 解析）</a>
      <a class="btn-tool btn-orange" style="background:#ef7118;border-color:#ef7118;color:#fff" href="/schedule/notice_center?date={{ override_date }}&amp;return_section=management">📢 通知中心</a>
      <a class="btn-tool btn-purple" style="background:#8039e8;border-color:#8039e8;color:#fff" href="/schedule/signups">📋 报名记录</a>
      <a class="btn-tool btn-gray" style="background:#d64a84;border-color:#d64a84;color:#fff" href="/schedule?mode=special&override_date={{ override_date }}">🌸 供台及佛台特别安排</a>
      <a class="btn-tool btn-gray" style="background:#078b9b;border-color:#078b9b;color:#fff" href="/schedule?mode=controls&override_date={{ override_date }}">🔧 报名开放控制</a>
      <a class="btn-tool btn-green" style="background:#b47a17;border-color:#b47a17;color:#fff" href="/reports">📚 资料中心</a>
      <a class="btn-tool btn-green" style="background:#237d76;border-color:#237d76;color:#fff" href="/schedule/fixed-roles">👥 固定岗位人员管理</a>
      <a class="btn-tool btn-gray" style="background:#596879;border-color:#596879;color:#fff" href="/schedule/settings">⚙️ 排班设置</a>
      <a class="btn-tool btn-gray" style="background:#146e86;border-color:#146e86;color:#fff" href="/schedule/whatsapp-automation">📲 WhatsApp 自动通知</a>
      <a class="btn-tool btn-blue" style="background:#275de4;border-color:#275de4;color:#fff" href="/volunteer">👥 打开义工报名页面</a>
    </div>
</div>
{% endif %}

{% if mode in ("day", "signup", "special", "controls") %}
{% if mode == "day" %}

<div class="card today-panel" id="today-schedule">
    <h2 class="section-title">📅 今日概况</h2>
    <div class="dashboard-note">负责人打开后，先确认日期、发布状态、报名人数和提醒。</div>

    <div class="date-card">
        <b>🗓 日期资料</b><br>
        阳历：{{ special_day_info.solar_date }}<br>
        农历：{{ special_day_info.lunar_text }}<br>

        {% if special_day_info.is_special %}
            <span class="text-red" style="font-weight:bold;">
                🔴 {{ special_day_info.special_names | join("、") }}
            </span><br>
            模板：{{ special_day_info.template_text }}<br>

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

    <div class="date-status-row compact-date-row">

        <div class="date-picker-card compact-date-picker">

            <div class="status-title">
                📅 选择工作日期
            </div>

            <input
                type="date"
                id="main_date"
                class="date-input"
                value="{{ override_date }}"
                onchange="goToScheduleDate(this.value)"
            >

        </div>

    </div>

    <div class="kpi-grid">
        <div class="kpi-card">
            <div class="kpi-title">👥 已报名</div>
            <div class="kpi-value">{{ day_summary.total }}</div>
        </div>

        <div class="kpi-card">
            <div class="kpi-title">🟡 待安排</div>
            <div class="kpi-value">{{ day_summary.pending }}</div>
        </div>

        <div class="kpi-card">
            <div class="kpi-title">🟢 已安排</div>
            <div class="kpi-value">{{ day_summary.assigned }}</div>
        </div>

        <div class="kpi-card">
            <div class="kpi-title">⚪ 已取消</div>
            <div class="kpi-value">{{ day_summary.cancelled }}</div>
        </div>
    </div>

    <div class="info-box" style="margin-top:14px;">
        <b>👥 当日义工报名详情（{{ override_date }}）</b>
        <div style="font-size:13px;color:#64748b;margin:5px 0 10px;">显示所有报名记录，包括待安排、已安排及已取消。</div>
        {% if signup_details %}
        <div style="overflow-x:auto;">
          <table style="width:100%;border-collapse:collapse;font-size:14px;min-width:590px;">
            <thead><tr style="background:#edf3f9;text-align:left;">
              <th style="padding:9px;border-bottom:1px solid #dce5ee;">义工姓名</th>
              <th style="padding:9px;border-bottom:1px solid #dce5ee;">报名岗位</th>
              <th style="padding:9px;border-bottom:1px solid #dce5ee;">报名时段</th>
              <th style="padding:9px;border-bottom:1px solid #dce5ee;">安排地点</th>
              <th style="padding:9px;border-bottom:1px solid #dce5ee;">状态</th>
            </tr></thead>
            <tbody>
            {% for person in signup_details %}
              <tr style="border-bottom:1px solid #edf1f5;{% if person.status == 'cancelled' %}opacity:.55;{% endif %}">
                <td style="padding:9px;">{{ person.name or person.volunteer_id or '—' }}</td>
                <td style="padding:9px;">{{ person.role or '—' }}</td>
                <td style="padding:9px;white-space:nowrap;">{{ person.start_time|string if person.start_time else '—' }} ～ {{ person.end_time|string if person.end_time else '—' }}</td>
                <td style="padding:9px;">
                  {% if person.assignment_places %}
                    <div style="display:flex;flex-wrap:wrap;align-items:center;gap:6px;">
                      {% for place in person.assignment_places %}
                        <span style="display:inline-block;white-space:nowrap;padding:5px 9px;border-radius:7px;font-size:13px;line-height:1.4;{% if '观音堂' in place %}background:#ffedd5;color:#9a3412;{% elif '活动中心' in place %}background:#fef9c3;color:#854d0e;{% else %}background:#e8f3ff;color:#1e40af;{% endif %}">{{ place }}</span>
                      {% endfor %}
                    </div>
                  {% else %}{{ person.assigned_place or '—' }}{% endif %}
                </td>
                <td style="padding:9px;white-space:nowrap;">
                  {% if person.status == 'assigned' %}<span style="color:#16803c;font-weight:700;">🟢 已安排</span>
                  {% elif person.status == 'cancelled' %}<span style="color:#64748b;">⚪ 已取消</span>
                  {% elif person.status == 'pending' %}<span style="color:#a66a00;font-weight:700;">🟡 待安排</span>
                  {% else %}{{ person.status }}{% endif %}
                </td>
              </tr>
            {% endfor %}
            </tbody>
          </table>
        </div>
        {% else %}<div style="color:#64748b;">这一天暂时没有报名记录。</div>{% endif %}
    </div>

    <div class="info-box">
        <b>📢 缺义工提醒</b><br>

        {% if shortage_summary %}
            {% for line in shortage_summary %}
                {{ line }}<br>
            {% endfor %}
        {% else %}
            🟢 目前没有缺义工岗位
        {% endif %}
    </div>
</div>

<div class="card">
    <h2 class="section-title">🔔 今日提醒</h2>
    <div class="reminder-stack">
<div class="alert alert-error">
    <h3 style="margin-top:0;">🔴 未安排提醒</h3>
    今日：<b>{{ pending_counts.today }}</b> 人　
    明日：<b>{{ pending_counts.tomorrow }}</b> 人　
    本月：<b>{{ pending_counts.month }}</b> 人
</div>



        {% if republish_info and republish_info.need_republish %}
        <div class="alert alert-warning">
            🟠 正式值班表已有更新<br>
            最后更新时间：{{ republish_info.last_schedule_change }}<br>
            请到通知中心复制最新正式值班表 WhatsApp。
        </div>
        {% endif %}
<div class="alert alert-warning">
    <h2>🌸 供台报名提醒</h2>

    {% for item in supply_alerts %}
        <div style="margin-bottom:14px;">
            <b>
                {{ item.date.strftime("%m/%d") }}
                {{ item.festival_label }}
            </b><br>

            {% if item.scheduled_names %}
                <span style="color:#176b39;">🟢 已安排设供台：{{ "、".join(item.scheduled_names) }}</span><br>
            {% endif %}
            {% if item.names %}
                义工自行报名：{{ "、".join(item.names) }}
            {% elif not item.scheduled_names %}
                <span style="color:#c0392b;">尚未安排设供台人员，也没有义工报名</span>
            {% endif %}
        </div>
    {% endfor %}
</div>


    </div>
</div>

<div class="card" id="today-operations">
    <h2 class="section-title">📝 今日操作｜负责人工作流程</h2>
    <div class="dashboard-note" style="margin-bottom:14px;line-height:1.8;">
        请按① → ② → ③ → ④操作。以下按钮针对目前选择的排班日期；发送 WhatsApp 前请先核对日期及名单。
    </div>

    <div class="action-grid">
        <div>
            <div style="font-weight:700;margin-bottom:7px;">① 晚上7:00｜发出报名通知</div>
            <a href="/schedule/notice_center?date={{ override_date }}&amp;return_section=today" class="btn-tool btn-orange">
                📢 报名通知／通知中心
                <span class="helper-text">进入通知中心，复制明日的报名通知，发送到 WhatsApp 群组；也可查看缺义工通知。</span>
            </a>
        </div>

        <div>
            <div style="font-weight:700;margin-bottom:7px;">② 晚上10:00｜安排报名义工</div>
            {% if not is_published %}
                {% if not is_today_or_past %}
                <form method="post" action="/schedule/generate_day">
                    <input type="hidden" name="date" id="generate_day_date" value="{{ override_date }}">
                    <button type="submit" class="btn-tool btn-blue">
                        🔄 重新自动排班
                        <span class="helper-text">报名截止后点击；系统依报名资料安排岗位、地点及时间。发布前可再次执行。</span>
                    </button>
                </form>
                {% else %}
                <form method="post" action="/schedule/fill_pending">
                    <input type="hidden" name="date" id="fill_pending_date" value="{{ override_date }}">
                    <button type="submit" class="btn-tool btn-blue">
                        ➕ 补入待安排
                        <span class="helper-text">今天或过去日期：把尚未安排的报名义工补入空缺，不重新生成整张排班表。</span>
                    </button>
                </form>
                {% endif %}
            {% else %}
                <div class="btn-tool btn-gray" style="cursor:default;box-sizing:border-box;">
                    ✅ 已发布，自动补排生效
                    <span class="helper-text">正式发布后，新报名由系统自动补排，不必再按「重新自动排班」。</span>
                </div>
            {% endif %}
        </div>

        <div>
            <div style="font-weight:700;margin-bottom:7px;">③ 排班完成后｜检查及调整</div>
            <form method="post" action="/schedule/view_assigned">
                <input type="hidden" name="date" id="view_assigned_date" value="{{ override_date }}">
                <button type="submit" class="btn-tool btn-purple">
                    👀 查看排班
                    <span class="helper-text">检查义工姓名、值班时间、岗位和地点；如有需要，进入后手动调整。</span>
                </button>
            </form>
        </div>

        <div>
            <div style="font-weight:700;margin-bottom:7px;">④ 检查无误后｜正式发布</div>
            <form method="post" action="/schedule/publish">
                <input type="hidden" name="date" id="publish_date" value="{{ override_date }}">
                <button type="submit" class="btn-tool {% if is_published %}btn-gray{% else %}btn-green{% endif %}">
                    {% if is_published %}
                        ✅ 已发布
                        <span class="helper-text">值班表已生效。若有新报名，系统会自动补排；可继续查看最新安排。</span>
                    {% else %}
                        🟢 发布正式值班表
                        <span class="helper-text">确认名单正确后点击发布，再到通知中心复制正式值班表发送 WhatsApp。</span>
                    {% endif %}
                </button>
            </form>
        </div>
    </div>

    <div class="dashboard-note" style="margin-top:14px;line-height:1.8;">
        💡 <b>发布后有新义工报名：</b>系统会自动补入空缺岗位。负责人可按「查看排班」核对最新名单；不需要重新自动排班或重新发布。<br>
        📌 <b>提醒：</b>进入通知中心或查看排班，不代表 WhatsApp 已发送或名单已确认；请负责人自行核对。
    </div>

    {% if is_published %}
    <div style="margin-top:12px;">
        <a href="/schedule/attendance_status" class="btn-tool" style="background:#078b9b;border-color:#078b9b;color:#fff;">
            📋 签到情况
            <span class="helper-text">发布后可查看义工签到纪录及当天签到情况。</span>
        </a>
    </div>
    {% endif %}
</div>


{% endif %}
{% if mode == "signup" %}
<div class="card">
    <h2 class="section-title">📋 报名录入</h2>
    <div class="dashboard-note">负责人只需要在这里加入报名：可以人工输入，也可以贴 WhatsApp 文字解析。</div>

    <div class="signup-unified">

        <div class="signup-block">
            <h3>➕ 负责人代报名</h3>

            <form method="post" action="/schedule/prebook_add">
                <input type="hidden" name="mode" value="day">
                <input type="hidden" name="single_date" id="prebook_single_date" value="{{ override_date }}">

                <div class="form-row-large">
                    <label>义工编号 / 姓名</label>

                    <div class="branch-search-row">
                        <button
                            type="button"
                            id="admin_branch_btn"
                            onclick="toggleAdminBranch()"
                            class="branch-toggle-btn">
                            CHE
                        </button>

                        <input type="hidden" id="admin_branch" name="branch" value="CHE">

                        <input
                            name="keyword"
                            id="admin_keyword"
                            placeholder="例如：108 / 姓名">
                    </div>

                    <div id="volunteer_lookup_result" class="lookup-text"></div>
                </div>

                {% for role in roles %}
                <label class="role-btn">

                    <input
                        type="checkbox"
                        name="roles"
                        value="{{ role }}">

                    {% if role == "值班" %}
                        👷 值班
                    {% elif role == "卫生" %}
                        🧹 卫生
                    {% elif role == "供台" %}
                        🛕 供台
                    {% else %}
                        {{ role }}
                    {% endif %}

                </label>
                {% endfor %}

                <div class="form-row-large">
                    <p>时间（只给值班用）</p>

                    <div class="time-grid">
                        <div>
                            <label>开始</label>
                            <select name="start_time" style="width:100%;">
                            {% for t in times %}
                                <option value="{{ t }}">{{ t }}</option>
                            {% endfor %}
                            </select>
                        </div>

                        <div>
                            <label>结束</label>
                            <select name="end_time" style="width:100%;">
                            {% for t in times %}
                                <option value="{{ t }}">{{ t }}</option>
                            {% endfor %}
                            </select>
                        </div>
                    </div>
                </div>

                <button type="submit" class="btn-tool btn-blue">➕ 加入报名</button>
            </form>
        </div>

        <div class="signup-divider"></div>

        <div class="signup-block">
            <h3>📥 WhatsApp 报名解析</h3>

            <form method="post" action="/schedule/parse_raw">
                <input type="hidden" name="single_date" id="raw_single_date" value="{{ override_date }}">

                <textarea
                    name="raw_signup"
                    rows="8"
                    placeholder="输入姓名 / 时间，例如：&#10;张三 10:00am-2:00pm&#10;李四 8:00am-12:00pm"></textarea>

                <br><br>

                <button type="submit" class="btn-tool btn-blue">📥 解析加入报名</button>
            </form>
        </div>

    </div>
</div>



{% endif %}
{% if mode == "special" %}
    <div class="card">
        <h2 class="section-title">⚙️ 特殊设置</h2>

        <div class="quick-actions">
            <button type="button" class="btn-tool btn-gray" onclick="showSpecial('master-table')">
                🛕 师父供台
                <span class="helper-text">点击后自动跳到设置区</span>
            </button>

            <button type="button" class="btn-tool btn-gray" onclick="showSpecial('extra-buddha')">
                🌸 初一补人
            </button>

            <button type="button" class="btn-tool btn-gray" onclick="showSpecial('buddha-leave')">
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

            <div class="info-box" style="margin:12px 0;">
                <b>📅 提前选择未来大日子（90天内）</b>
                <p style="font-size:16px;">选好日期后会自动载入该日已保存的供台人员。</p>
                {% if master_dates %}
                <select id="master_day_select" aria-label="选择大日子" style="width:100%;font-size:19px;">
                    {% for item in master_dates %}
                    <option value="{{ item.date }}" {% if item.date == selected_master_date %}selected{% endif %}>{{ item.label }}</option>
                    {% endfor %}
                </select>
                <label for="master_calendar">或打开日历选择：</label>
                <input type="date" id="master_calendar" value="{{ selected_master_date }}"
                       min="{{ master_dates[0].date }}" max="{{ master_dates[-1].date }}"
                       style="width:100%;font-size:19px;padding:9px;">
                <small>日历只接受上方列表中的大日子；普通日期不会保存。</small>
                {% else %}
                <p>未来90天内没有找到大日子，请检查农历规则。</p>
                {% endif %}
            </div>

            <form method="post" action="/schedule/save_day_flags">
                <input type="hidden" name="form_type" value="master_table">
                <input type="hidden" name="date" value="{{ selected_master_date }}">

                <div class="info-box">
                    系统会根据初一、十五及佛诞日自动判断是否需要设供台。<br>
                    若前一天应收供台但未安排人员，系统会继续提醒收供台。
                </div>

                <h4>🪷 设师父供台人员</h4>
                <p style="color:#666;font-size:20px;">通常 1 至 2 位佛台组人员即可。</p>

                <select name="setup_person_1">

                    <option value="">请选择</option>

                    {% for p in supply_signups %}
                        <option value="{{ p.name }}"
                            {% if setup_saved and setup_saved[0] == p.name %}selected{% endif %}>
                            {{ p.name }}
                        </option>
                    {% endfor %}

                    {% for n in buddha_names %}
                        <option value="{{ n }}"
                            {% if setup_saved and setup_saved[0] == n %}selected{% endif %}>
                            {{ n }}
                        </option>
                    {% endfor %}

                </select>

                <select name="setup_person_2">
                    <option value="">请选择</option>

                    {% for p in supply_signups %}
                        <option value="{{ p.name }}"
                            {% if setup_saved|length > 1 and setup_saved[1] == p.name %}selected{% endif %}>
                            {{ p.name }}
                        </option>
                    {% endfor %}

                    {% for n in buddha_names %}
                        <option value="{{ n }}" {% if setup_saved|length > 1 and setup_saved[1] == n %}selected{% endif %}>
                            {{ n }}
                        </option>
                    {% endfor %}
                </select>

                <hr>

                <button type="submit" class="btn-tool btn-blue">💾 保存设供台人员</button>
            </form>

            <form method="post" action="/schedule/save_day_flags" style="margin-top:24px;">
                <input type="hidden" name="form_type" value="remove_table">
                <input type="hidden" name="setup_date" value="{{ selected_master_date }}">
                <h4>📦 收师父供台人员</h4>
                <label for="remove_day_calendar"><b>📅 选择收供台日期（未来90天内，普通日也可以）</b></label>
                <input id="remove_day_calendar" type="date" name="date" value="{{ selected_remove_date }}"
                       min="{{ remove_min_date }}" max="{{ remove_max_date }}" required
                       style="width:100%;font-size:19px;padding:9px;margin:10px 0;">
                <p style="font-size:15px;color:#666;">更换日期会载入该日已保存的收供台人员。</p>
                <p style="color:#666;font-size:20px;">
                    通常 1 至 2 位佛台组人员。若没有安排人员，系统视为继续供奉。
                </p>

                <select name="remove_person_1">
                    <option value="">请选择</option>
                    {% for n in buddha_names %}
                    <option value="{{ n }}" {% if remove_saved and remove_saved[0] == n %}selected{% endif %}>
                        {{ n }}
                    </option>
                    {% endfor %}
                </select>

                <select name="remove_person_2">
                    <option value="">请选择</option>
                    {% for n in buddha_names %}
                    <option value="{{ n }}" {% if remove_saved|length > 1 and remove_saved[1] == n %}selected{% endif %}>
                        {{ n }}
                    </option>
                    {% endfor %}
                </select>

                <p>普通义工帮忙收供台：</p>
                <input
                    name="remove_extra_person"
                    value="{{ remove_saved[2] if remove_saved|length > 2 else '' }}"
                    placeholder="例如：张三，可空"
                    style="width:100%;font-size:22px;"
                >

                <br><br>

                <button type="submit" class="btn-tool btn-blue">💾 保存收供台人员</button>
            </form>
        </div>

        <!-- 2. 初一补人 -->
        <div id="extra-buddha" class="special-box" style="display:none;">
            <h3>🌸 初一整理佛台补人</h3>

            {% if extra_buddha_dates %}
            <form method="get" action="/schedule">
                <input type="hidden" name="mode" value="special">
                <input type="hidden" name="override_date" value="{{ override_date }}">
                <input type="hidden" name="panel" value="extra-buddha">
                <label for="extra_buddha_date">选择初一日期：</label>
                <select id="extra_buddha_date" name="extra_buddha_date" onchange="this.form.submit()" style="width:100%;font-size:22px;">
                    {% for d in extra_buddha_dates %}
                    <option value="{{ d }}" {% if d == selected_extra_buddha_date %}selected{% endif %}>{{ d }}（初一）</option>
                    {% endfor %}
                </select>
            </form>
            <div class="info-box">
                <b>🛕 当天负责整理佛台义工</b><br><br>

                {% if extra_buddha_fixed_names %}
                    <form method="post" action="/schedule/override">
                        <input type="hidden" name="date" value="{{ selected_extra_buddha_date }}">
                        <input type="hidden" name="return_panel" value="extra-buddha">
                        <input type="hidden" name="return_override_date" value="{{ override_date }}">
                        {% for n in extra_buddha_fixed_names %}
                        <div style="margin:10px 0;display:flex;align-items:center;gap:12px;flex-wrap:wrap;">
                            <span>🟢 {{ n }}</span>
                            <input type="hidden" name="original_name" value="{{ n }}">
                            <select name="replacement_name" style="font-size:20px;max-width:100%;">
                                <option value="">不变，保留原本</option>
                                <option value="__REMOVE__">请假，不找替补</option>
                                {% for candidate in buddha_names %}
                                {% if candidate != n %}<option value="{{ candidate }}">换人：{{ candidate }}</option>{% endif %}
                                {% endfor %}
                            </select>
                        </div>
                        {% endfor %}
                        <button type="submit" class="btn-tool btn-blue">💾 保存请假 / 换人</button>
                    </form>
                    <form method="post" action="/schedule/override/reset" style="margin-top:10px;">
                        <input type="hidden" name="date" value="{{ selected_extra_buddha_date }}">
                        <input type="hidden" name="return_panel" value="extra-buddha">
                        <input type="hidden" name="return_override_date" value="{{ override_date }}">
                        <button type="submit" class="btn-tool btn-red">↩ 恢复当天固定佛台名单</button>
                    </form>
                {% else %}
                    ⚪ 该日没有固定整理佛台义工
                    <form method="post" action="/schedule/override/reset" style="margin-top:10px;">
                        <input type="hidden" name="date" value="{{ selected_extra_buddha_date }}">
                        <input type="hidden" name="return_panel" value="extra-buddha">
                        <input type="hidden" name="return_override_date" value="{{ override_date }}">
                        <button type="submit" class="btn-tool btn-red">↩ 恢复当天固定佛台名单</button>
                    </form>
                {% endif %}
            </div>

            <form method="post" action="/schedule/save_day_flags">
                <input type="hidden" name="form_type" value="extra_buddha">
                <input type="hidden" name="date" value="{{ selected_extra_buddha_date }}">

                <p style="font-size:20px;color:#666;">
                    仅初一需要。可增加第 3 位整理佛台义工。
                </p>

                <p>额外第 3 位整理佛台义工：</p>

                <select name="extra_buddha_person" style="width:100%;font-size:22px;">
                    <option value="">不需要增加</option>
                    {% for n in buddha_names %}
                    <option value="{{ n }}" {% if extra_buddha_saved == n %}selected{% endif %}>
                        {{ n }}
                    </option>
                    {% endfor %}
                </select>

                <br><br>

                <button type="submit" class="btn-tool btn-blue">💾 保存初一补人</button>
            </form>
            {% else %}
            <p>未来90天内没有可选择的初一日期。</p>
            {% endif %}
        </div>

        <!-- 3. 佛台请假 -->
        <div id="buddha-leave" class="special-box" style="display:none;">
            <h3>🔁 佛台请假 / 换人</h3>

            <form method="get" action="/schedule" style="margin-bottom:18px;">
                <input type="hidden" name="mode" value="special">
                <input type="hidden" name="override_date" value="{{ override_date }}">
                <input type="hidden" name="panel" value="buddha-leave">
                <label for="buddha_leave_date">选择请假 / 换人日期：</label>
                <input type="date" id="buddha_leave_date" name="buddha_leave_date"
                       value="{{ selected_buddha_leave_date }}"
                       min="{{ buddha_leave_dates[0] }}" max="{{ buddha_leave_dates[-1] }}"
                       onchange="this.form.submit()"
                       style="width:100%;font-size:22px;box-sizing:border-box;">
            </form>

            <form method="post" action="/schedule/override">
                <input type="hidden" name="date" value="{{ selected_buddha_leave_date }}">
                <input type="hidden" name="return_panel" value="buddha-leave">
                <input type="hidden" name="return_override_date" value="{{ override_date }}">

                {% if buddha_leave_names %}
                    {% for old_name in buddha_leave_names %}
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

                    <button type="submit" class="btn-tool btn-blue">💾 保存佛台请假 / 换人</button>
                {% else %}
                    <p style="font-size:22px;">这一天没有固定佛台。</p>
                {% endif %}
            </form>
            <form method="post" action="/schedule/override/reset" style="margin-top:14px;">
                <input type="hidden" name="date" value="{{ selected_buddha_leave_date }}">
                <input type="hidden" name="return_panel" value="buddha-leave">
                <input type="hidden" name="return_override_date" value="{{ override_date }}">
                <button type="submit" class="btn-tool btn-red">↩ 恢复当天固定佛台名单</button>
            </form>
        </div>
    </div>

{% endif %}
{% if mode == "controls" %}
<div class="card">

    <h2 class="section-title">⚙️ 报名开放控制</h2>

    <table class="setting-table">

        <!-- 多日报名 -->

        <tr>
            <td style="padding:12px 0;">
                <b>📅 下个月多日报名</b><br>

                {% if multi_day_signup_open %}
                    <span style="color:green;font-weight:bold;">
                        🟢 已开放
                    </span>
                {% else %}
                    <span style="color:#999;font-weight:bold;">
                        ⚪ 已关闭
                    </span>
                {% endif %}
            </td>

            <td style="text-align:right;white-space:nowrap;">

                <form method="post"
                      action="/schedule/toggle_setting"
                      class="inline-form">

                    <input type="hidden"
                           name="key"
                           value="multi_day_signup_open">

                    <input type="hidden"
                           name="force"
                           value="true">

                    <button class="btn-tool btn-green mini-btn">
                        开启
                    </button>

                </form>

                <form method="post"
                      action="/schedule/toggle_setting"
                      class="inline-form">

                    <input type="hidden"
                           name="key"
                           value="multi_day_signup_open">

                    <input type="hidden"
                           name="force"
                           value="false">

                    <button class="btn-tool btn-red mini-btn">
                        关闭
                    </button>

                </form>

            </td>
            
        </tr>

        <tr>
            <td colspan="2">
                <hr>
            </td>
        </tr>

        <!-- 膳食 -->

        <tr>

            <td style="padding:12px 0;">

                <b>🍱 膳食组报名</b><br>

                {% if meal_signup_open %}
                    <span style="color:green;font-weight:bold;">
                        🟢 已开放
                    </span>
                {% else %}
                    <span style="color:#999;font-weight:bold;">
                        ⚪ 已关闭
                    </span>
                {% endif %}

            </td>

            <td style="text-align:right;white-space:nowrap;">

                <form method="post"
                      action="/schedule/toggle_setting"
                      class="inline-form">

                    <input type="hidden"
                           name="key"
                           value="meal_signup_open">

                    <input type="hidden"
                           name="force"
                           value="true">

                    <button class="btn-tool btn-green mini-btn">
                        开启
                    </button>

                </form>

                <form method="post"
                      action="/schedule/toggle_setting"
                      class="inline-form">

                    <input type="hidden"
                           name="key"
                           value="meal_signup_open">

                    <input type="hidden"
                           name="force"
                           value="false">

                    <button class="btn-tool btn-red mini-btn">
                        关闭
                    </button>

                </form>

            </td>

        </tr>

        <tr>
            <td colspan="2">
                <hr>
            </td>
        </tr>

        <!-- 素食结缘 -->

        <tr>

            <td style="padding:12px 0;">

                <b>🥗 素食结缘报名</b><br>

                {% if food_offering_open %}
                    <span style="color:green;font-weight:bold;">
                        🟢 已开放
                    </span>
                {% else %}
                    <span style="color:#999;font-weight:bold;">
                        ⚪ 已关闭
                    </span>
                {% endif %}

            </td>

            <td style="text-align:right;white-space:nowrap;">

                <form method="post"
                    action="/schedule/toggle_setting"
                    class="inline-form">

                    <input type="hidden"
                        name="key"
                        value="food_offering_open">

                    <input type="hidden"
                        name="force"
                        value="true">

                    <button class="btn-tool btn-green mini-btn">
                        开启
                    </button>

                </form>

                <form method="post"
                    action="/schedule/toggle_setting"
                    class="inline-form">

                    <input type="hidden"
                        name="key"
                        value="food_offering_open">

                    <input type="hidden"
                        name="force"
                        value="false">

                    <button class="btn-tool btn-red mini-btn">
                        关闭
                    </button>

                </form>

            </td>

        </tr>

    </table>

</div>



{% endif %}

{% endif %}

</div>
<script>
function toggleAdminBranch(){
    const btn = document.getElementById("admin_branch_btn");
    const branch = document.getElementById("admin_branch");

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

window.addEventListener("load", function () {
    if (window.location.hash === "#today-schedule") {
        setTimeout(function () {
            var target = document.getElementById("today-schedule");
            if (target) {
                target.scrollIntoView({
                    behavior: "smooth",
                    block: "start"
                });
            }
        }, 300);
    }
});

</script>
<script>
function goToScheduleDate(dateValue) {
    if (!dateValue) return;

    window.location.href =
        "/schedule?mode=day&override_date=" + encodeURIComponent(dateValue);
}
</script>
<script>
// 每20秒轻量检查；报名/取消/自动补排变化后自动更新负责人数据。
// 正在输入或修改表单时暂停刷新，避免丢失未保存内容。
(function () {
    const dateValue = {{ override_date|tojson }};
    const apiUrl = '/schedule/api/admin-signup-version?date=' + encodeURIComponent(dateValue);
    let previousVersion = null;
    let busy = false;
    let editing = false;
    document.addEventListener('input', function (event) {
        if (event.target.closest('input,textarea,select,[contenteditable]')) editing = true;
    }, true);
    async function checkSignupUpdates() {
        if (busy || document.hidden || editing) return;
        busy = true;
        try {
            const response = await fetch(apiUrl, {cache:'no-store', credentials:'same-origin'});
            if (!response.ok) return;
            const data = await response.json();
            if (previousVersion !== null && data.version !== previousVersion) {
                const active = document.activeElement;
                if (active && active.matches('input,textarea,select,[contenteditable]')) return;
                window.location.reload();
                return;
            }
            previousVersion = data.version;
        } catch (error) {
            // 网络暂时断线时保留当前页面，下次继续检查。
        } finally { busy = false; }
    }
    checkSignupUpdates();
    setInterval(checkSignupUpdates, 20000);
})();
</script>
</body>
</html>
"""


DAY_OUTPUT_HTML = """
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>正式值班表</title>

<link rel="stylesheet"
      href="{{ url_for('static', filename='css/toolbox.css') }}">

<style>

.page-wrap{
    max-width:1100px;
    margin:auto;
    padding:25px;
}

.output-box{
    min-height:650px;
}

.output-box textarea{
    width:100%;
    height:650px;

    font-size:26px;
    line-height:1.7;

    padding:18px;

    border-radius:16px;
    border:2px solid #d6d6d6;

    box-sizing:border-box;

    resize:vertical;
}

</style>

<script>
function copySchedule(){

    const ta = document.getElementById("output");

    if(!ta){
        alert("❌ 找不到值班表内容");
        return;
    }

    ta.focus();
    ta.select();
    ta.setSelectionRange(0, 999999);

    const ok = document.execCommand("copy");

    if(ok){
        alert("✅ 已复制，可以贴去 WhatsApp 群");
    }else{
        alert("❌ 复制失败，请手动全选复制");
    }
}
</script>

</head>

<body>

<div class="page-wrap">

<div class="card">

    <h1 class="page-title">
        📋 正式值班表
    </h1>

    <p class="page-subtitle">
        可直接复制到 WhatsApp 发布。
    </p>

    <div class="btn-row">

        <a
            class="btn-tool btn-gray"
            href="/schedule?mode=day">

            ⬅ 返回负责人首页

        </a>

        <button
            type="button"
            class="btn-tool btn-green"
            onclick="copySchedule()">

            📋 一键复制

        </button>

    </div>

    <div class="output-box">

        <textarea
            id="output"
            readonly>{{ output }}</textarea>

    </div>

</div>

</div>

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

<link rel="stylesheet"
      href="{{ url_for('static', filename='css/toolbox.css') }}">

<style>
.page-wrap{
    max-width:1100px;
    margin:auto;
    padding:25px;
}

.output-box textarea{
    width:100%;
    height:650px;
    font-size:26px;
    line-height:1.7;
    padding:18px;
    border-radius:16px;
    border:2px solid #d6d6d6;
    box-sizing:border-box;
    resize:vertical;
}
</style>

<script>
function copyOutput(btn){
    const text = document.getElementById("output").value;
    navigator.clipboard.writeText(text);

    btn.innerText = "✅ 已复制";

    setTimeout(function(){
        btn.innerText = "📋 一键复制";
    }, 2000);
}
</script>

</head>

<body>

<div class="page-wrap">

<div class="card">

    <h1 class="page-title">📢 月预报名表</h1>

    <p class="page-subtitle">
        可直接复制到 WhatsApp 群发布。
    </p>

    <div class="btn-row">

        <a class="btn-tool btn-gray"
           href="/schedule?mode=prebook">
            ⬅ 返回月预报名
        </a>

        <button class="btn-tool btn-green"
                onclick="copyOutput(this)">
            📋 一键复制
        </button>

    </div>

    <div class="output-box">
        <textarea id="output" readonly>{{ output }}</textarea>
    </div>

</div>

</div>

</body>
</html>
"""

# ============================================================
# WhatsApp 自动通知控制台（第一阶段：测试群组限定）
# ============================================================
@schedule_bp.route("/schedule/whatsapp-automation", methods=["GET", "POST"])
def whatsapp_automation_admin():
    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    from datetime import timedelta
    from schedule.services.whatsapp_queue import enqueue_whatsapp_test, recent_whatsapp_jobs
    flags = (
        ("wa_auto_enabled", "自动发送总开关 / Master switch"),
        ("wa_test_mode", "测试模式 / Test mode"),
        ("wa_agent_auto_send", "🤖 Windows Agent 自动发送（关闭则人工确认 SEND）/ Agent auto send"),
        ("wa_notice_1900", "🕖 19:00 报名通知 / Signup notice"),
        ("wa_roster_2200", "🕙 22:00 正式值班表 / Full roster"),
        ("wa_late_signup", "🌅 早上新增义工汇总通知（普通日 07:00／大日子 05:30）/ Morning late-signup summary"),
        ("wa_auto_publish_2200", "🕙 22:00 自动排班发布 / Auto publish"),
    )
    message = ""
    if request.method == "POST":
        action = request.form.get("action", "save")
        if action == "save":
            import re
            new_time = request.form.get("wa_notice_time", "19:00").strip()
            if not re.fullmatch(r"(?:[01]\d|2[0-3]):[0-5]\d", new_time):
                message = "❌ 时间格式错误，请使用 HH:MM / Invalid time"
            else:
                # 测试阶段：只有测试模式开启时才允许启用 22:00 功能。
                test_mode_requested = request.form.get("wa_test_mode") == "on"
                restricted_keys = ("wa_roster_2200", "wa_auto_publish_2200", "wa_late_signup")
                rejected = (not test_mode_requested and any(
                    request.form.get(key) == "on" for key in restricted_keys
                ))
                if rejected:
                    message = "❌ 请先开启测试模式，才能启用 22:00 或早上通知 / Enable test mode first"
                else:
                    for key, _ in flags:
                        enabled = request.form.get(key) == "on"
                        if key in restricted_keys and not test_mode_requested:
                            enabled = False
                        save_schedule_setting(key, "true" if enabled else "false")
                    save_schedule_setting("wa_notice_time", new_time)
                    message = "✅ 设置已保存 / Settings saved"
        elif action in ("test", "test_signup"):
            if not is_schedule_setting_on("wa_test_mode"):
                message = "请先开启并保存测试模式 / Enable and save test mode first"
            else:
                try:
                    if action == "test_signup":
                        target = (malaysia_now().date() + timedelta(days=1)).isoformat()
                        body = build_signup_notice_whatsapp(target)
                        if not isinstance(body, str) or not body.strip():
                            raise ValueError("Empty signup notice")
                        body = "🧪【测试：明日报名通知 / TEST SIGNUP NOTICE】\n" + body
                    else:
                        body = "🧪 观音堂 WhatsApp 自动通知测试 / Test message"
                    job_id = enqueue_whatsapp_test(body, "admin_manual_test")
                    message = f"✅ 已加入测试群组队列 / Queued for test group: #{job_id}。若 Agent 正在运行，可能立即发送。"
                except Exception:
                    import logging
                    logging.exception("WhatsApp 测试入队失败")
                    message = "❌ 测试入队失败，请查看服务器日志 / Queue error"
        else:
            message = "不支持的操作 / Unsupported action"

    current = {key: is_schedule_setting_on(key) for key, _ in flags}
    notice_time = str(get_schedule_setting("wa_notice_time", "19:00") or "19:00")
    import re
    if not re.fullmatch(r"(?:[01]\d|2[0-3]):[0-5]\d", notice_time):
        notice_time = "19:00"
    try:
        jobs = recent_whatsapp_jobs(20)
        queue_error = ""
    except Exception:
        jobs = []
        queue_error = "无法读取发送记录，请检查 whatsapp_notification_jobs 表 / Queue unavailable"

    return render_template_string('''<!doctype html><html lang="zh"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>WhatsApp 自动通知 / Automation</title>
<style>body{font-family:system-ui,'Microsoft YaHei',sans-serif;background:#f3f7fa;margin:0;padding:20px;color:#183047}
main{max-width:850px;margin:auto}.card{background:white;border-radius:14px;padding:22px;margin:16px 0;box-shadow:0 2px 12px #dce5eb}
h1{font-size:25px}label{display:flex;justify-content:space-between;align-items:center;padding:14px 0;border-bottom:1px solid #eee;gap:16px}
input[type=checkbox]{width:25px;height:25px;accent-color:#14876a}input[type=time]{font-size:18px;padding:8px;border-radius:8px;border:1px solid #aaa}
button,.btn{display:inline-block;border:0;border-radius:9px;padding:12px 18px;color:white;background:#146e86;cursor:pointer;text-decoration:none;margin:8px 5px 4px 0}
.secondary{background:#657786}.warning{background:#fff5df;padding:12px;border-radius:8px}.success{background:#e3f8ee;padding:10px;border-radius:8px}
table{width:100%;border-collapse:collapse;font-size:13px}td,th{padding:8px;border-bottom:1px solid #ddd;text-align:left;word-break:break-word}</style></head>
<body><main><h1>📲 WhatsApp 自动通知 / Automation</h1>
<p class="warning">🔒 目前只允许发送至测试群组。Windows Agent 运行时，测试消息可能实际发送至测试群组。<br>
TEST GROUP ONLY. Test messages may be sent when the Windows Agent is running.</p>
{% if message %}<p class="success">{{ message }}</p>{% endif %}
<div class="card"><form method="post"><input type="hidden" name="action" value="save">
{% for key,label in flags %}
<label><span>{{ label }}</span><input type="checkbox" name="{{ key }}" {% if current[key] %}checked{% endif %}></label>
{% endfor %}
<p class="warning">⚠️ 22:00 与早上通知开关仅限测试模式；开启自动发布可能修改共用数据库中的真实排班。请先保持两个 22:00 开关关闭。<br>22:00 switches require test mode. Auto publish may modify live shared schedule data.</p>
<p class="warning">🌅 早上通知只汇总昨晚正式值班表之后新增的值班／卫生义工；需要 22:00 快照及 Windows Agent 定时检查。<br>Morning summary requires the 22:00 snapshot and a scheduled Windows Agent tick.</p>
<label><span>🕖 报名通知时间（马来西亚）<br>Signup notice time (MYT)</span><input type="time" name="wa_notice_time" value="{{ notice_time }}" required></label>
<p class="warning">定时任务需要每分钟运行一次检查脚本；设置时间不会自动修改 Windows Task Scheduler。<br>Run the Windows tick every minute; changing this time does not reconfigure Task Scheduler.</p>
<button type="submit">💾 保存设置 / Save settings</button></form></div>
<div class="card"><h2>🧪 测试队列 / Test queue</h2>
<p>下列按钮会建立真实的测试群组待发送任务。请确认测试群组及 Agent 配置正确。</p>
<form method="post" style="display:inline"><input type="hidden" name="action" value="test"><button type="submit">📤 普通测试 / Basic test</button></form>
<form method="post" style="display:inline"><input type="hidden" name="action" value="test_signup"><button type="submit">📣 测试明日报名通知 / Test signup notice</button></form>
<p class="warning">22:00 接口已接入；请先验证数据库隔离及完整排班流程，再开启自动发布。</p>
{% if queue_error %}<p class="warning">{{ queue_error }}</p>{% endif %}</div>
<div class="card"><details><summary style="cursor:pointer;font-size:21px;font-weight:700;padding:8px 0;user-select:none">📜 最近记录 / Recent jobs（点击展开 / Click to expand）</summary><div style="overflow:auto;max-height:380px;margin-top:12px"><table><tr><th>ID</th><th>类别 / Type</th><th>状态 / Status</th><th>创建 / Created</th></tr>
{% for j in jobs %}<tr><td>{{j.id}}</td><td>{{j.message_type}}</td><td>{{j.status}}</td><td>{{j.created_at}}</td></tr>{% else %}<tr><td colspan="4">暂无记录 / No records</td></tr>{% endfor %}</table></div></details></div>
<a class="btn secondary" href="{{url_for('schedule.schedule_admin')}}">← 返回负责人首页 / Back</a></main></body></html>''',flags=flags,current=current,notice_time=notice_time,jobs=jobs,message=message,queue_error=queue_error)
