from schedule.services.schedule_management_queries import load_management_signup_view
# schedule_web.py
from schedule.services.schedule_view_queries import (
    build_month_prebook_message,
    build_shortage_notice,
)

from schedule.services.schedule_calendar_queries import (
    get_pending_signup_counts,
    get_pending_signup_details,
    is_big_day,
    _master_table_dates,
)
from schedule.services.schedule_signup_queries import (
    get_signup_summary_for_date,
    load_admin_signup_details,
    display_supply_role,
)
from schedule.services.schedule_pending_queries import (
    load_pending_signups_for_date,
    get_recent_cleaning_places,
)
from schedule.services.schedule_text_helpers import (
    display_report_name,
    search_volunteer_names,
    get_default_time_by_role,
    normalize_whatsapp_keyword,
)
import os
from schedule.schedule_page_templates import (
    _MANAGEMENT_MENU_HTML, LOGIN_HTML, _FIXED_ROLES_HTML,
    SCHEDULE_HTML, DAY_OUTPUT_HTML, MONTHLY_PREBOOK_HTML,
)
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
            WHERE id IN ({placeholders}) OR member_id IN ({placeholders})
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

def load_weekly_buddha_names(weekday):
    """Supabase 星期固定佛台名单；空名单也是有效设置，不回退 Excel。"""
    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                SELECT v.name
                FROM schedule_fixed_buddha_weekly f
                JOIN volunteers v ON v.id::text = f.volunteer_id::text
                WHERE f.weekday = %s
                ORDER BY f.position
            """, (weekday,))
            return [str(row['name']).strip() for row in cur.fetchall()]


def get_base_fixed_buddha_for_date(date_str):
    # 云端轮值优先；仅数据库异常时使用旧 Excel 作为临时备用。
    try:
        weekday_number = pd.to_datetime(date_str).weekday()
        return load_weekly_buddha_names(weekday_number)
    except Exception as e:
        print("⚠️ Supabase 佛台固定名单读取失败，暂用 Excel：", repr(e))

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


def _sync_buddha_assignments(cur, date_str, names):
    """仅同步该日期的整理佛台名单；不触碰其他岗位、报名或签到记录。"""
    cur.execute("SELECT pg_advisory_xact_lock(hashtext(%s))", ("buddha:" + date_str,))
    cur.execute("""
        SELECT id, name, remarks FROM volunteer_schedule_assignments
        WHERE assignment_date=%s AND role='整理佛台'
          AND COALESCE(status,'assigned') <> 'cancelled'
        ORDER BY id FOR UPDATE
    """, (date_str,))
    rows = cur.fetchall()
    cur.execute("SELECT is_published, extra_buddha_person FROM schedule_day_flags WHERE flag_date=%s", (date_str,))
    flag = cur.fetchone() or {}
    # 未生成排班时仅保存覆盖名单；正式排班生成时原有 builder 会读取它。
    if not rows and not flag.get('is_published'):
        return False

    extra = str(flag.get('extra_buddha_person') or '').strip()
    desired = list(dict.fromkeys([str(n).strip() for n in names if str(n).strip()] + ([extra] if extra else [])))
    existing = set()
    changed = False
    for row in rows:
        name = str(row.get('name') or '').strip()
        if name not in desired or name in existing:
            cur.execute("DELETE FROM volunteer_schedule_assignments WHERE id=%s", (row['id'],))
            changed = True
        else:
            existing.add(name)
    for name in desired:
        if name in existing:
            continue
        cur.execute("SELECT id FROM volunteers WHERE name=%s AND status='在册' ORDER BY id LIMIT 1", (name,))
        person = cur.fetchone()
        cur.execute("""
            INSERT INTO volunteer_schedule_assignments
                (volunteer_id, name, assignment_date, role, shift_label,
                 assigned_place, status, remarks, created_at)
            VALUES (%s,%s,%s,'整理佛台','整理佛台','整理佛台',
                    'assigned','负责人佛台特别安排',NOW())
        """, ((person or {}).get('id'), name, date_str))
        changed = True
    return changed and bool(flag.get('is_published'))


def delete_buddha_override(date_str):
    """恢复指定日期的固定轮值，并同步已生成的佛台排班。"""
    names = get_base_fixed_buddha_for_date(date_str)
    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute('DELETE FROM schedule_buddha_overrides WHERE duty_date=%s', (date_str,))
            republish = _sync_buddha_assignments(cur, date_str, names)
        conn.commit()
    if republish:
        mark_schedule_need_republish(date_str)



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

    
    

def save_buddha_override(date_str, final_names):
    """保存云端佛台覆盖名单，并在同一事务同步已生成的佛台排班。"""
    names = list(dict.fromkeys(str(n).strip() for n in final_names if str(n).strip()))
    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                INSERT INTO schedule_buddha_overrides (duty_date, names, updated_at)
                VALUES (%s, %s::jsonb, NOW())
                ON CONFLICT (duty_date) DO UPDATE
                SET names=EXCLUDED.names, updated_at=NOW()
            """, (date_str, __import__('json').dumps(names, ensure_ascii=False)))
            republish = _sync_buddha_assignments(cur, date_str, names)
        conn.commit()
    if republish:
        mark_schedule_need_republish(date_str)



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




@schedule_bp.route("/schedule/edit_assigned/<int:assignment_id>", methods=["GET", "POST"])
def edit_assigned_place(assignment_id):

    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    # 从排班结果进入时，保存后回到同一天的排班结果。
    return_date = request.values.get("return_date", "").strip()
    try:
        return_date = datetime.strptime(return_date, "%Y-%m-%d").strftime("%Y-%m-%d") if return_date else ""
    except ValueError:
        return_date = ""

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

            if request.method == "POST" and request.form.get("action") == "delete":
                # 只删除选中的 assignment，不误删同一报名拆出的其他时段。
                cur.execute("""
                    select id, signup_id, assignment_date
                    from volunteer_schedule_assignments where id=%s for update
                """, (assignment_id,))
                target = cur.fetchone()
                if not target:
                    return "❌ 找不到排班记录", 404
                cur.execute("delete from volunteer_schedule_assignments where id=%s", (assignment_id,))
                # 最后一笔被删除时同步取消报名，防止夜间缺漏修复重新补入。
                if target["signup_id"] is not None:
                    cur.execute("""
                        select 1 from volunteer_schedule_assignments
                        where signup_id=%s and role='值班'
                          and coalesce(status,'assigned') <> 'cancelled' limit 1
                    """, (target["signup_id"],))
                    if not cur.fetchone():
                        cur.execute("""
                            update volunteer_schedule_signups
                            set status='cancelled', assigned_place=null
                            where id=%s and role='值班'
                        """, (target["signup_id"],))
                conn.commit()
                mark_schedule_need_republish(target["assignment_date"])
                if return_date and return_date == str(target["assignment_date"]):
                    return redirect(url_for("schedule.schedule_view_assigned", date=return_date))
                return redirect("/schedule/signups")

            if request.method == "POST":
                assigned_place = request.form.get("assigned_place", "").strip()
                start_time = request.form.get("start_time", "").strip()
                end_time = request.form.get("end_time", "").strip()
                signup_start = request.form.get("signup_start_time", "").strip()
                signup_end = request.form.get("signup_end_time", "").strip()

                def _minutes(value):
                    raw = str(value or "").lower().replace(" ", "")
                    for fmt in ("%I:%M%p", "%I%p", "%H:%M", "%H:%M:%S"):
                        try:
                            t = datetime.strptime(raw, fmt)
                            return t.hour * 60 + t.minute
                        except ValueError:
                            pass
                    return None

                def _clock(value):
                    h, m = divmod(value, 60)
                    return f"{h % 12 or 12}:{m:02d}{'am' if h < 12 else 'pm'}"

                def _shift(value):
                    return "绿" if value < 600 else ("橙" if value < 840 else "黄")


                if assigned_place not in place_options:
                    return "❌ 地点无效", 400
                a_start, a_end = _minutes(start_time), _minutes(end_time)
                if a_start is None or a_end is None or a_start >= a_end:
                    return "❌ 排班时段无效", 400

                cur.execute("SELECT assignment_date FROM volunteer_schedule_assignments WHERE id=%s", (assignment_id,))
                date_row = cur.fetchone()
                if not date_row:
                    return "❌ 找不到排班记录", 404
                cur.execute("SELECT pg_advisory_xact_lock(hashtext(%s))", ("duty:" + str(date_row["assignment_date"]),))
                cur.execute("""
                    SELECT * FROM volunteer_schedule_assignments
                    WHERE id=%s FOR UPDATE
                """, (assignment_id,))
                assignment_info = cur.fetchone()
                if not assignment_info:
                    return "❌ 找不到排班记录", 404
                assignment_date = assignment_info["assignment_date"]
                signup_id = assignment_info.get("signup_id")
                role = assignment_info.get("role")
                old_signup_start = old_signup_end = None
                siblings = []
                if signup_id is not None:
                    s_start, s_end = _minutes(signup_start), _minutes(signup_end)
                    if s_start is None or s_end is None or s_start >= s_end:
                        return "❌ 报名时段无效", 400
                    if a_start < s_start or a_end > s_end:
                        if role == "值班" and (a_start, a_end) == (_minutes(assignment_info["start_time"]), _minutes(assignment_info["end_time"])):
                            a_start, a_end = max(a_start, s_start), min(a_end, s_end)
                            if a_start >= a_end:
                                return "❌ 当前排班完全超出新报名，请先调整当前排班时间", 400
                            start_time, end_time = _clock(a_start), _clock(a_end)
                        else:
                            return "❌ 实际排班必须在报名时间范围内，请同时调整排班时间", 400
                    cur.execute("""
                        SELECT start_time, end_time FROM volunteer_schedule_signups
                        WHERE id=%s AND COALESCE(status,'pending') <> 'cancelled' FOR UPDATE
                    """, (signup_id,))
                    signup_row_for_update = cur.fetchone()
                    if not signup_row_for_update:
                        return "❌ 关联报名不存在或已取消", 400
                    old_signup_start = _minutes(signup_row_for_update["start_time"])
                    old_signup_end = _minutes(signup_row_for_update["end_time"])
                    cur.execute("""
                        SELECT * FROM volunteer_schedule_assignments
                        WHERE signup_id=%s AND id<>%s
                          AND COALESCE(status,'assigned') <> 'cancelled'
                        ORDER BY id FOR UPDATE
                    """, (signup_id, assignment_id))
                    siblings = cur.fetchall()
                    # 已锁定的其他排班不能被静默缩短或删除。
                    for sibling in siblings:
                        x, y = _minutes(sibling.get('start_time')), _minutes(sibling.get('end_time'))
                        if x is None or y is None:
                            return "❌ 其他排班时间无法识别", 400
                        if role == "值班" and max(x, s_start, a_start) < min(y, s_end, a_end):
                            return "❌ 与同一义工另一笔排班重叠，请先调整该时段", 400
                        if sibling.get('locked_by_admin') and (x < s_start or y > s_end):
                            return "❌ 其他时段已由负责人锁定，请先单独调整该时段", 400

                # 无需手选班别：根据实际排班开始时间自动确定。
                cur.execute("""
                    UPDATE volunteer_schedule_assignments
                    SET shift_label=%s, assigned_place=%s, start_time=%s, end_time=%s,
                        remarks='负责人调整排班', assignment_source='admin', locked_by_admin=true
                    WHERE id=%s
                """, (_shift(a_start) if role == "值班" else assignment_info.get("shift_label"), assigned_place, start_time, end_time, assignment_id))

                if signup_id is not None:
                    cur.execute("""
                        UPDATE volunteer_schedule_signups SET start_time=%s, end_time=%s WHERE id=%s
                    """, (signup_start, signup_end, signup_id))
                    if role == '值班':
                        # 缩短报名：仅修剪同一报名的未锁定时段；其他义工完全不动。
                        for sibling in siblings:
                            x, y = _minutes(sibling.get('start_time')), _minutes(sibling.get('end_time'))
                            if x is None or y is None or sibling.get('locked_by_admin'):
                                continue
                            nx, ny = max(x, s_start), min(y, s_end)
                            if nx >= ny:
                                cur.execute("DELETE FROM volunteer_schedule_assignments WHERE id=%s", (sibling['id'],))
                            elif nx != x or ny != y:
                                cur.execute("""
                                    UPDATE volunteer_schedule_assignments
                                    SET start_time=%s, end_time=%s, shift_label=%s WHERE id=%s
                                """, (_clock(nx), _clock(ny), _shift(nx), sibling['id']))

                        # 延长报名：只补旧报名边界以外的新增时段，不重排既有记录。
                        if old_signup_start is not None and old_signup_end is not None:
                            extensions = []
                            if s_start < old_signup_start:
                                extensions.append((s_start, min(old_signup_start, s_end)))
                            if s_end > old_signup_end:
                                extensions.append((max(old_signup_end, s_start), s_end))
                            occupied = [(a_start, a_end)]
                            for sibling in siblings:
                                x, y = _minutes(sibling.get('start_time')), _minutes(sibling.get('end_time'))
                                if x is not None and y is not None:
                                    occupied.append((max(x, s_start), min(y, s_end)))
                            if extensions:
                                from schedule.builders.duty_v2 import plan_for_date
                                from schedule.builders.schedule_builder import _dry_run_2200_gaps
                                from schedule.services.smart_assignment_engine import load_place_history_cache
                                cur.execute("""
                                    SELECT volunteer_id, assigned_place, start_time, end_time
                                    FROM volunteer_schedule_assignments
                                    WHERE assignment_date=%s AND role='值班'
                                      AND COALESCE(status,'assigned') <> 'cancelled'
                                """, (assignment_date,))
                                existing = list(cur.fetchall())
                                history = load_place_history_cache(target_date=str(assignment_date))
                                for ext_start, ext_end in extensions:
                                    for gx, gy in _dry_run_2200_gaps(ext_start, ext_end, occupied):
                                        person = dict(signup_id=signup_id, volunteer_id=assignment_info['volunteer_id'],
                                                      name=assignment_info['name'], start_time=_clock(gx), end_time=_clock(gy))
                                        planned = plan_for_date([person], target_date=assignment_date,
                                                                existing=existing, history=history, patch_mode=True)
                                        intervals = [(_minutes(item['start_time']), _minutes(item['end_time'])) for _, item in planned]
                                        if not planned or _dry_run_2200_gaps(gx, gy, intervals):
                                            raise ValueError('新增时段未完整补排，取消保存')
                                        for key, item in planned:
                                            cur.execute("""
                                                INSERT INTO volunteer_schedule_assignments
                                                (signup_id, volunteer_id, name, assignment_date, role,
                                                 shift_label, assigned_place, start_time, end_time,
                                                 status, remarks, assignment_source, locked_by_admin)
                                                VALUES (%s,%s,%s,%s,'值班',%s,%s,%s,%s,'assigned',
                                                        '负责人延长报名后自动补排','system',false)
                                            """, (signup_id, person['volunteer_id'], person['name'], assignment_date,
                                                  key[0], key[1:], item['start_time'], item['end_time']))
                                            existing.append(dict(item, assigned_place=key[1:]))
                                        occupied.extend(intervals)

                conn.commit()
                mark_schedule_need_republish(assignment_date)
                if return_date and return_date == str(assignment_date):
                    return redirect(url_for("schedule.schedule_view_assigned", date=return_date))
                return redirect("/schedule/signups")

            cur.execute("""
                select *
                from volunteer_schedule_assignments
                where id = %s
            """, (assignment_id,))

            row = cur.fetchone()
            signup_row = None
            if row and row.get("signup_id") is not None:
                cur.execute("""
                    select start_time, end_time from volunteer_schedule_signups
                    where id=%s
                """, (row["signup_id"],))
                signup_row = cur.fetchone()

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
.time-picker { display:flex; align-items:center; gap:8px; margin:8px 0; }
.time-picker input[type=number] { width:90px; min-width:0; margin:0; text-align:center; }
.time-picker .time-minute { width:100px; }
.time-picker select { width:125px; margin:0; }
.time-picker span { font-size:26px; }
@media(max-width:480px) { .time-picker input[type=number] { width:70px; } .time-picker .time-minute { width:80px; } .time-picker select { width:100px; } }

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
    <input type="hidden" name="return_date" value="{{ return_date }}">

    {% if signup_row %}
    <div style="background:#eef7ff;padding:14px;border-radius:12px;margin:12px 0;">
        <h3 style="margin:0 0 10px;">📝 修改报名时段</h3>
        <label>报名开始时间</label>
        <div class="time-picker" data-time-picker>
  <input type="hidden" name="signup_start_time" value="{{ signup_row.start_time or '' }}">
  <input type="number" class="time-hour" min="1" max="12" step="1" inputmode="numeric" aria-label="小时" required>
  <span>:</span>
  <input type="number" class="time-minute" min="0" max="59" step="1" inputmode="numeric" aria-label="分钟" required>
  <select class="time-period" aria-label="上午或下午"><option value="am">AM</option><option value="pm">PM</option></select>
</div>
        <label>报名结束时间</label>
        <div class="time-picker" data-time-picker>
  <input type="hidden" name="signup_end_time" value="{{ signup_row.end_time or '' }}">
  <input type="number" class="time-hour" min="1" max="12" step="1" inputmode="numeric" aria-label="小时" required>
  <span>:</span>
  <input type="number" class="time-minute" min="0" max="59" step="1" inputmode="numeric" aria-label="分钟" required>
  <select class="time-period" aria-label="上午或下午"><option value="am">AM</option><option value="pm">PM</option></select>
</div>
        <small>只修改这位义工的原始报名，不会自动重排其他义工。</small>
    </div>
    {% else %}
    <p style="color:#805b15;">此排班没有关联报名记录，只能调整排班。</p>
    {% endif %}

    <label>地点</label>
    <select name="assigned_place" required>
        {% for p in place_options %}
        <option value="{{ p }}" {% if row.assigned_place == p %}selected{% endif %}>
            {{ p }}
        </option>
        {% endfor %}
    </select>

    <label>开始时间</label>
    <div class="time-picker" data-time-picker>
  <input type="hidden" name="start_time" value="{{ row.start_time or '' }}">
  <input type="number" class="time-hour" min="1" max="12" step="1" inputmode="numeric" aria-label="小时" required>
  <span>:</span>
  <input type="number" class="time-minute" min="0" max="59" step="1" inputmode="numeric" aria-label="分钟" required>
  <select class="time-period" aria-label="上午或下午"><option value="am">AM</option><option value="pm">PM</option></select>
</div>

    <label>结束时间</label>
    <div class="time-picker" data-time-picker>
  <input type="hidden" name="end_time" value="{{ row.end_time or '' }}">
  <input type="number" class="time-hour" min="1" max="12" step="1" inputmode="numeric" aria-label="小时" required>
  <span>:</span>
  <input type="number" class="time-minute" min="0" max="59" step="1" inputmode="numeric" aria-label="分钟" required>
  <select class="time-period" aria-label="上午或下午"><option value="am">AM</option><option value="pm">PM</option></select>
</div>

    <button type="submit">💾 保存调整</button>

</form>

<form method="post" onsubmit="return confirm('确定删除这笔排班？\\n如果这是该报名最后一笔值班排班，系统也会取消该报名，避免自动补排回来。');">
    <input type="hidden" name="return_date" value="{{ return_date }}">
    <input type="hidden" name="action" value="delete">
    <button type="submit" style="background:#fff0f0;color:#a51e1e;border:2px solid #d99;border-radius:10px;">🗑️ 删除这笔排班</button>
</form>

<br>
{% if return_date %}
<form method="get" action="/schedule/view_assigned"><input type="hidden" name="date" value="{{ return_date }}"><button type="submit">⬅ 返回排班结果</button></form>
{% else %}
<a href="/schedule/signups">⬅ 返回报名管理</a>
{% endif %}

</div>
<script>
(function () {
  const pickers = Array.from(document.querySelectorAll('[data-time-picker]'));
  for (const picker of pickers) {
    const hidden = picker.querySelector('input[type="hidden"]');
    const hour = picker.querySelector('.time-hour');
    const minute = picker.querySelector('.time-minute');
    const period = picker.querySelector('.time-period');
    const raw = (hidden.value || '').trim();
    const m = raw.match(/^(\d{1,2})(?::(\d{1,2}))?(am|pm)$/i);
    const m24 = raw.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
    if (m) {
      hour.value = Number(m[1]);
      minute.value = Number(m[2] || 0);
      period.value = m[3].toLowerCase();
    } else if (m24) {
      const h = Number(m24[1]);
      hour.value = h % 12 || 12;
      minute.value = Number(m24[2]);
      period.value = h < 12 ? 'am' : 'pm';
    }
    const form = picker.closest('form');
    form.addEventListener('submit', function (event) {
      const h = Number(hour.value), min = Number(minute.value);
      if (!Number.isInteger(h) || h < 1 || h > 12 || !Number.isInteger(min) || min < 0 || min > 59) {
        event.preventDefault();
        alert('请选择正确的小时（1–12）和分钟（0–59）');
        return;
      }
      hidden.value = h + ':' + String(min).padStart(2, '0') + period.value;
    });
  }
})();
</script>
</body>
</html>
""",
        row=row,
        signup_row=signup_row,
        place_options=place_options,
        return_date=return_date
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



def build_signup_notice_whatsapp(date_str):
    # 19:00 手动预览与自动通知统一：只提醒明天报名，不显示初步排班。
    target_day = date.fromisoformat(date_str)
    weekday = '一二三四五六日'[target_day.weekday()]
    return (f'明天义工报名（{target_day:%d/%m/%Y}，星期{weekday}）\n\n'
            '义工报名请点击以下链接：\n'
            'https://gyt-checkin-sg.onrender.com/volunteer\n\n'
            '感恩大家, 功德无量！🙏🙏🙏')


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
        "8:00am~10:00am \n🟢 观音堂: \n{hall}\n\n🟢 活动中心: \n{center}\n\n\n"
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
                SELECT setup_people, remove_people FROM schedule_day_flags WHERE flag_date = %s
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
    remove_people = [name.strip() for name in str((setup_row or {}).get('remove_people') or '').splitlines() if name.strip()]
    # 普通日如有实际供台安排也应显示；无收供台安排则不显示空栏目。
    setup_notice = ('设师父供台：\n' + ('　'.join(supply_people) if supply_people else '尚未安排') + '\n\n') if (day_info.get('setup_shifu') or supply_people) else ''
    remove_notice = ('12:00pm~2:00pm\n师父供台:\n' + '　'.join(remove_people) + '\n\n12pm的香结束之后，请下供桌。\n\n') if remove_people else ''

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
        """仅合并通知显示：同一人同地点的连续/重叠时段；保留同时间多人同行。"""
        if not rows:
            return ""
        from collections import OrderedDict

        # 此函数每次只处理一个地点，因此不跨地点合并。
        by_name = OrderedDict()
        unparsed = []
        for row in rows:
            name = str(row.get("name") or "").strip()
            start_text = str(row.get("start_time") or "").strip()
            end_text = str(row.get("end_time") or "").strip()
            if not name:
                continue
            start, end = _minutes(start_text), _minutes(end_text)
            if start is None or end is None or end <= start:
                unparsed.append((name, start_text, end_text))
                continue
            by_name.setdefault(name, []).append((start, end))

        merged_rows = []
        for name, intervals in by_name.items():
            intervals.sort()
            merged = []
            for start, end in intervals:
                if merged and start <= merged[-1][1]:
                    merged[-1] = (merged[-1][0], max(merged[-1][1], end))
                else:
                    merged.append((start, end))
            for start, end in merged:
                merged_rows.append((name, _clock(start), _clock(end)))
        merged_rows.extend(unparsed)

        grouped = OrderedDict()
        for name, start, end in merged_rows:
            names = grouped.setdefault((start, end), [])
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

    # 19:00 报名通知：仅调整文字排版，不改变报名来源、分组或发送流程。
    date_only = f"{date_obj.day}/{date_obj.month}/{date_obj.year}"
    weekday_only = f"({weekday_map[date_obj.weekday()]})"
    lunar_and_festival = ("\n\n" + "\n".join(
        [f"（{lunar_label}）"] if lunar_label else []
    )) if lunar_label else ""
    if festival_names:
        lunar_and_festival += "\n\n" + "\n".join(f"🙏 {name}" for name in festival_names)

    text = f"""师兄们大家好！

报名通知（初步名单）

{date_only}

{weekday_only}{lunar_and_festival}

{greeting_time}

［整理佛台］

{altar_start}~{altar_end}  或
{altar_start}~完成佛台工作

{buddha_text}

{setup_notice}{remove_notice}［打扫卫生］

{cleaning_start}~{cleaning_end} 或
{cleaning_start}~完成卫生工作

佛堂卫生: {"　".join(str(r.get("name") or "").strip() for r in groups["佛堂卫生"] if str(r.get("name") or "").strip())}

二楼卫生: {"　".join(str(r.get("name") or "").strip() for r in groups["二楼卫生"] if str(r.get("name") or "").strip())}

楼梯卫生: {"　".join(str(r.get("name") or "").strip() for r in groups["楼梯卫生"] if str(r.get("name") or "").strip())}

每日卫生义工请注意：清理完卫生之后，请把卫生用具包括吸尘机放回原位（活动中心store里面的小房间）

［值班］

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

- 观音堂早晚香 由值班义工带领上香。

- 观音堂续香 -
佛友可以续香（黑色无烟香），但是要跟着请安词，必须燃烧完了一支香才续香。如有多位佛友要续香，请先让给先到达观音堂的佛友，轮流续香。

！！值班义工请注意：
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
    # 历史日期仅查询正式历史出席；不进入现有报名/排班 dashboard。
    try:
        selected_day = date.fromisoformat(override_date)
    except ValueError:
        return "日期格式无效", 400
    weekday_label = "星期" + "一二三四五六日"[selected_day.weekday()]
    if mode == "day" and selected_day < date(2026, 11, 15):
        if request.method != "GET":
            return "历史出席仅供查询", 405
        with get_conn() as conn:
            with conn.cursor(cursor_factory=RealDictCursor) as cur:
                cur.execute("""
                    SELECT volunteer_id, volunteer_name, role, start_time, end_time, hours
                    FROM public.volunteer_attendance_history
                    WHERE attendance_date = %s
                    ORDER BY start_time NULLS LAST, volunteer_name, volunteer_id, role
                """, (selected_day,))
                history_rows = cur.fetchall()
        history_people = len({str(r["volunteer_id"]) for r in history_rows if r["volunteer_id"]})
        history_hours = sum(r["hours"] for r in history_rows if r["hours"] is not None)
        return render_template_string(
            SCHEDULE_HTML, mode=mode, override_date=override_date,
            weekday_label=weekday_label, historical_view=True,
            special_day_info=get_special_day_info(selected_day),
            history_rows=history_rows, history_people=history_people,
            history_hours=history_hours,
            history_unknown_hours=sum(r["hours"] is None for r in history_rows),
        )
    actual_rows = []
    actual_error = False
    if mode == "day" and selected_day <= malaysia_today():
        try:
            from reports_service import load_live_attendance
            actual_rows = load_live_attendance(selected_day, selected_day + timedelta(days=1))
        except Exception as exc:
            print("实际出席查询失败:", repr(exc))
            actual_error = True
    _dashboard_started = time.perf_counter()
    dashboard = load_admin_dashboard_data(
        mode=mode,
        override_date=override_date,
    )
    _dashboard_loaded = time.perf_counter()

    buddha_names = load_buddha_name_options()
    fixed_buddha_today = get_fixed_buddha_for_date(override_date)
    signup_details = load_admin_signup_details(
        override_date, fixed_buddha=fixed_buddha_today if mode == "day" else None,
        is_published=dashboard["is_published"], big_day=is_big_day(override_date) if mode == "day" else False
    ) if mode in ("day", "signup", "special", "controls") else []
    # 佛台请假日期独立于负责人主页日期；允许提前安排未来30天。
    buddha_leave_dates = [(malaysia_today() + timedelta(days=i)).isoformat() for i in range(31)] if mode == "special" else []
    requested_leave_date = request.args.get("buddha_leave_date", "")
    # 佛台请假默认日期独立于负责人页面 override_date：10:00 起默认明天。
    buddha_leave_default_date = (now.date() + timedelta(days=1) if now.hour >= 10 else now.date()).isoformat()
    selected_buddha_leave_date = (requested_leave_date if requested_leave_date in buddha_leave_dates
                                  else (buddha_leave_default_date if buddha_leave_default_date in buddha_leave_dates
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
        item_display["date_weekday"] = (day.strftime("%m/%d") + "（" + "一二三四五六日"[day.weekday()] + "）") if day else ""
        item_display["scheduled_names"] = scheduled_by_date.get(day.isoformat(), []) if day else []
        supply_alerts_display.append(item_display)
    # 收供台日期独立于设供台日期，允许未来90天内的普通日。
    remove_date_arg = request.args.get("remove_date", "").strip()
    try:
        remove_day = date.fromisoformat(remove_date_arg)
    except ValueError:
        remove_day = date.fromisoformat(default_schedule_date)
    if not malaysia_today() <= remove_day <= malaysia_today() + timedelta(days=90):
        remove_day = date.fromisoformat(default_schedule_date)
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
        weekday_label=weekday_label,
        actual_visible=mode == "day" and selected_day <= malaysia_today(),
        actual_rows=actual_rows,
        actual_error=actual_error,
        actual_people=len({str(r['volunteer_id']) for r in actual_rows if r['volunteer_id']}),
        actual_hours=sum(float(r['hours']) for r in actual_rows if r['hours'] is not None and r.get('end_time')),
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
        pending_tomorrow_label=(malaysia_today() + timedelta(days=1)).strftime("%Y-%m-%d") + "（" + "一二三四五六日"[(malaysia_today() + timedelta(days=1)).weekday()] + "）",
        pending_details=get_pending_signup_details() if mode == "day" else [],
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
    return_to = url_for("schedule.schedule_admin", mode="controls", override_date=request.form.get("override_date") or request.args.get("override_date") or malaysia_today().isoformat())

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
    """一键安全排班并发布；已存在的安排绝不整表覆盖。"""
    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    from markupsafe import escape
    from schedule.builders.schedule_builder import (
        repair_unassigned_duty_for_date,
        rebuild_published_cleaning_assignments,
        _wa6_minutes,
        _validate_full_duty_before_publish,
    )

    target_date = (request.form.get("date") or "").strip()
    try:
        selected = date.fromisoformat(target_date)
    except ValueError:
        return "❌ 日期格式错误", 400

    def failure(reason):
        return (f"<h2>❌ 未发布：{escape(str(reason))}</h2>"
                f"<p>原有排班和发布状态保持不变；请检查后重试。</p>"
                f"<a href='/schedule?mode=day&amp;override_date={target_date}'>返回当天安排</a>"), 409

    try:
        # 已发布/已有排班时禁止调用 run_schedule_for_date，因它会删除整日旧安排。
        published = is_schedule_published(target_date)
        existing = db_query("""
            SELECT count(*) AS n FROM volunteer_schedule_assignments
            WHERE assignment_date=%s
        """, (target_date,), fetchone=True)
        has_assignments = int((existing or {}).get("n") or 0) > 0

        if not published and not has_assignments:
            if selected < malaysia_today():
                return failure("过去日期不能重新自动排班")
            result = run_schedule_for_date(target_date)
            if not isinstance(result, str) or result.lstrip().startswith("❌"):
                return failure(result or "自动排班失败")
        else:
            # 只补没有 assignment 的值班报名；保留全部现有排班及管理员调整。
            repair = repair_unassigned_duty_for_date(target_date)
            if repair["failed"]:
                return failure(f"值班补排失败，报名编号：{repair['failed']}")

        # 统一按实际 assignments 验证，不相信报名表的 status 字段。
        signups = db_query("""
            SELECT id, volunteer_id, name, role, start_time, end_time
            FROM volunteer_schedule_signups
            WHERE signup_date=%s AND role IN ('值班','卫生')
              AND COALESCE(status,'pending') <> 'cancelled'
        """, (target_date,), fetchall=True) or []
        assignments = db_query("""
            SELECT signup_id, role, assigned_place, start_time, end_time,
                   assignment_source, locked_by_admin
            FROM volunteer_schedule_assignments
            WHERE assignment_date=%s
              AND COALESCE(status,'assigned') <> 'cancelled'
        """, (target_date,), fetchall=True) or []
        if not published:
            try:
                _validate_full_duty_before_publish(signups, assignments, existing=assignments)
            except ValueError as exc:
                return failure(exc)
        by_signup = {}
        for row in assignments:
            if row.get("signup_id") is not None:
                by_signup.setdefault(int(row["signup_id"]), []).append(row)

        problems = []
        for signup in signups:
            sid = int(signup["id"])
            matches = [r for r in by_signup.get(sid, []) if r.get("role") == signup["role"]]
            if not matches:
                problems.append(f"{signup['name']}（{signup['role']}）未安排")
                continue
            if signup["role"] == "值班":
                try:
                    start = _wa6_minutes(signup["start_time"])
                    end = _wa6_minutes(signup["end_time"])
                    intervals = sorted((_wa6_minutes(r["start_time"]),
                                        _wa6_minutes(r["end_time"])) for r in matches)
                    cursor = start
                    for a, b in intervals:
                        if a > cursor:
                            break
                        cursor = max(cursor, b)
                    if start >= end or cursor < end:
                        problems.append(f"{signup['name']}值班时间未完整覆盖")
                except (ValueError, TypeError):
                    problems.append(f"{signup['name']}值班时间格式异常")
        if problems:
            return failure("；".join(problems[:12]))

        # 已发布不重复触发快照和 WhatsApp 入队；首次发布沿用原有服务。
        if not published:
            publish_schedule_for_date(target_date)
        return redirect(url_for("schedule.schedule_admin", mode="day",
                                override_date=target_date, published="1"))
    except Exception:
        import logging
        logging.getLogger(__name__).exception("一键排班并发布失败：%s", target_date)
        return failure("系统处理异常，请查看 Render 日志")


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


@schedule_bp.route("/schedule/view_assigned", methods=["GET", "POST"])
def schedule_view_assigned():

    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    date_str = request.values.get("date", "").strip()
    try:
        date_str = datetime.strptime(date_str, "%Y-%m-%d").strftime("%Y-%m-%d")
    except ValueError:
        return "❌ 日期格式错误", 400

    rows = load_assigned_places_for_date(date_str)

    cleaning_groups = {
        "佛堂卫生": [],
        "二楼卫生": [],
        "楼梯卫生": [],
    }

    offering_groups = {"设师父供台": [], "收师父供台": []}

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
            "volunteer_id": r.get("volunteer_id"),
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

        elif display_supply_role(role, place):
            offering_groups[display_supply_role(role, place)].append(item)

        elif role == "值班":
            shift_key = str(shift or "").replace("班", "")

            if shift_key in duty_groups and place in duty_groups[shift_key]:
                duty_groups[shift_key][place].append(item)
            else:
                other_groups.setdefault(place, []).append(item)

        else:
            other_groups.setdefault(place, []).append(item)

    flags = load_day_flags(date_str)
    for role, field in (("设师父供台", "setup_people"), ("收师父供台", "remove_people")):
        group = offering_groups[role]
        seen = {str(item.get("name") or "").strip() for item in group}
        for name in str(flags.get(field) or "").splitlines():
            name = name.strip()
            if name and name not in seen:
                group.append({"name": name, "id": None})
                seen.add(name)
    setup_required = bool(flags.get("need_setup_master_table") or get_special_day_info(date.fromisoformat(date_str)).get("setup_shifu"))
    remove_required = bool(flags.get("need_remove_master_table"))
    other_groups_display = {place: names for place, names in other_groups.items() if place != "整理佛台"}

    # 仅合并查看排班页面的连续时段，不修改原始 assignments。
    # 保留每段原始 ID，负责人仍可分别调整，避免只修改其中一段却误以为全部已改。
    def _time_minutes(value):
        import re
        if value is None:
            return None
        if hasattr(value, "hour"):
            return value.hour * 60 + value.minute
        match = re.fullmatch(r"(\d{1,2}):(\d{2})\s*(am|pm)?", str(value).strip().lower())
        if not match:
            return None
        hour, minute, suffix = int(match[1]), int(match[2]), match[3]
        if minute > 59 or (suffix and not 1 <= hour <= 12) or (not suffix and hour > 23):
            return None
        if suffix:
            hour = hour % 12 + (12 if suffix == "pm" else 0)
        return hour * 60 + minute

    for places in duty_groups.values():
        for place, names in places.items():
            by_person = {}
            for item in names:
                # 仅同一编号（缺编号时用姓名）及同一地点/班别才考虑合并。
                key = str(item.get("volunteer_id") or item["name"]).strip()
                by_person.setdefault(key, []).append(item)
            display_rows = []
            for person_rows in by_person.values():
                valid, invalid = [], []
                for item in person_rows:
                    start = _time_minutes(item.get("start_time"))
                    end = _time_minutes(item.get("end_time"))
                    if start is None or end is None or start >= end:
                        invalid.append(item)
                    else:
                        valid.append((start, end, item))
                valid.sort(key=lambda entry: (entry[0], entry[1]))
                merged = []
                for start, end, item in valid:
                    if merged and start <= merged[-1][1]:
                        group = merged[-1]
                        if end > group[1]:
                            group[1] = end
                            group[3] = item["end_time"]
                        group[4].append(item["id"])
                    else:
                        merged.append([start, end, item["start_time"], item["end_time"], [item["id"]], item["name"]])
                for start, end, start_text, end_text, ids, name in merged:
                    display_rows.append({"name": name, "start_time": start_text,
                                         "end_time": end_text, "edit_ids": ids,
                                         "sort_time": start})
                for item in invalid:
                    display_rows.append({**item, "edit_ids": [item["id"]],
                                         "sort_time": float("inf")})
            display_rows.sort(key=lambda item: item["sort_time"])
            places[place] = display_rows

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
    .assigned-floating-back {
        position: fixed;
        right: 22px;
        bottom: 22px;
        z-index: 9999;
        background: #2563eb;
        color: white !important;
        padding: 13px 20px;
        border-radius: 999px;
        font-size: 20px;
        font-weight: bold;
        text-decoration: none;
        box-shadow: 0 4px 14px rgba(0,0,0,.2);
    }
    /* 排班结果：仅更换页面底色 */
    body { background-color: #F5F1FC !important; }

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
    /* 仅调整查看排班的响应式排列，不改变数据和按钮 */
    .assigned-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:16px;align-items:start;}
    .assigned-grid > .card,.assigned-grid > .assignment-section{min-width:0;margin:0;}
    .assigned-grid .person-row{flex-wrap:wrap;}
    @media(max-width:767px){.assigned-grid{grid-template-columns:minmax(0,1fr);}}
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

        <div class="assigned-grid">
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
                                href="/schedule/edit_assigned/{{ item.id }}?return_date={{ date_str }}">
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

            <h2 class="section-title">🙏 整理佛台</h2>

            <div class="assignment-section">
                {% for item in other_groups.get("整理佛台", []) %}
                    <div class="person-row">
                        <div><span class="person-name">{{ item.name }}</span></div>
                        <a class="edit-link"
                           href="/schedule/edit_assigned/{{ item.id }}?return_date={{ date_str }}">✏️ 调整</a>
                    </div>
                {% endfor %}
                {% if not other_groups.get("整理佛台") %}
                    <div class="empty-line">暂无安排</div>
                {% endif %}
            </div>

            {% for role, people in offering_groups.items() %}
            {% if people or (role == '设师父供台' and setup_required) or (role == '收师父供台' and remove_required) %}
            <h2 class="section-title">🙏 {{ role }}</h2>
            <div class="assignment-section">
                <div class="person-time">{{ '6:00am ~ 8:00am' if role == '设师父供台' else '12:00pm ~ 2:00pm' }}</div>
                {% for item in people %}
                <div class="person-row">
                    <div><span class="person-name">{{ item.name }}</span></div>
                    {% if item.id %}<a class="edit-link" href="/schedule/edit_assigned/{{ item.id }}?return_date={{ date_str }}">✏️ 调整</a>
                    {% else %}<a class="edit-link" href="/schedule?mode=special&amp;override_date={{ date_str }}&amp;panel=master-table{% if role == '收师父供台' %}&amp;remove_date={{ date_str }}{% endif %}">供台设置</a>{% endif %}
                </div>
                {% else %}<div class="empty-line" style="color:#b91c1c;">尚未安排</div>{% endfor %}
            </div>
            {% endif %}
            {% endfor %}

        </div>

        </div>

        <div class="card">

            <h2 class="section-title">🏠 值班</h2>

            {% for shift, places in duty_groups.items() %}
            {% if shift != "绿" or places["观音堂"] or places["活动中心"] %}

                {% if shift == "绿" %}
                    <h3 class="shift-title">🟢 绿班</h3>
                {% elif shift == "橙" %}
                    <h3 class="shift-title">🟠 橙班</h3>
                {% elif shift == "黄" %}
                    <h3 class="shift-title">🟡 黄班</h3>
                {% else %}
                    <h3 class="shift-title">{{ shift }}</h3>
                {% endif %}

                <div class="assigned-grid">
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

                                    <div>
                                        {% for edit_id in item.edit_ids %}
                                            <a class="edit-link"
                                            href="/schedule/edit_assigned/{{ edit_id }}?return_date={{ date_str }}">
                                                ✏️ 调整{% if item.edit_ids|length > 1 %}{{ loop.index }}{% endif %}
                                            </a>
                                        {% endfor %}
                                    </div>

                                </div>
                            {% endfor %}
                        {% else %}
                            <div class="empty-line">暂无安排</div>
                        {% endif %}

                    </div>
                {% endfor %}
                </div>
            {% endif %}
            {% endfor %}

        </div>

        {% if other_groups_display %}
        <div class="card">

            <h2 class="section-title">其他</h2>

            {% for place, names in other_groups_display.items() %}
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
                            href="/schedule/edit_assigned/{{ item.id }}?return_date={{ date_str }}">
                                ✏️ 调整
                            </a>

                        </div>
                    {% endfor %}

                </div>
            {% endfor %}

        </div>
        {% endif %}

    </div>

    <a class="assigned-floating-back" href="/schedule?mode=day&amp;override_date={{ date_str }}">⬅ 返回</a>
    </body>
    </html>
    """,
        date_str=date_str,
        cleaning_groups=cleaning_groups,
        offering_groups=offering_groups,
        setup_required=setup_required,
        remove_required=remove_required,
        duty_groups=duty_groups,
        other_groups=other_groups,
        other_groups_display=other_groups_display,
    )


@schedule_bp.route("/reports")
def public_reports():
    now = malaysia_now()
    authorized = session.get('reports_admin_authorized') is True
    history = {}
    other_files = []
    storage_error = False
    import re
    try:
        bucket = get_supabase_client().storage.from_("reports")
        files = bucket.list()
        for f in files:
            filename = f.get("name")
            if not filename or not filename.endswith(".xlsx"):
                continue
            monthly = re.fullmatch(r'(\d{4})_(\d{2})_month_report\.xlsx', filename)
            yearly = re.fullmatch(r'(\d{4})_year_report\.xlsx', filename)
            if monthly and (int(monthly[1]), int(monthly[2])) >= (2026, 11):
                continue
            if yearly and int(yearly[1]) >= 2026:
                continue
            item = {"filename": filename, "display_name": display_report_name(filename),
                    "url": bucket.create_signed_url(filename, 3600)["signedURL"]}
            if monthly or yearly:
                history.setdefault((monthly or yearly)[1], []).append(item)
            else:
                other_files.append(item)
    except Exception:
        from flask import current_app
        current_app.logger.exception('资料中心历史文件读取失败')
        storage_error = True
    for items in history.values():
        items.sort(key=lambda item: item['filename'], reverse=True)
    periods = [(y, m) for y in range(2026, now.year + 1) for m in range(1, 13)
               if (2026, 11) <= (y, m) <= (now.year, now.month)]
    return render_template_string("""
    <!doctype html><html lang="zh"><head><meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>观音堂资料中心</title>
    <link rel="stylesheet" href="{{ url_for('static', filename='css/toolbox.css') }}">
    <style>
    body{background:#f7f6f4;font-family:Microsoft YaHei,Arial,sans-serif}
    .reports-wrap{max-width:850px;margin:auto;padding:18px 14px 80px}
    .report-card{background:white;border:1px solid #e5e7eb;border-radius:16px;padding:22px;margin:16px 0}
    .report-card h2{margin:0 0 12px;font-size:24px;color:#752d3d}
    .report-card p{font-size:16px;color:#666;line-height:1.6}
    .report-card select{width:100%;font-size:20px;padding:12px;border:1px solid #ccc;border-radius:10px;margin:8px 0 14px;background:white}
    .report-card .btn-tool{display:block;width:100%;box-sizing:border-box;text-align:center;font-size:20px}
    .file-row{display:flex;align-items:center;justify-content:space-between;gap:12px;padding:12px 0;border-bottom:1px solid #eee;font-size:18px}
    .file-row .btn-tool{width:auto;flex-shrink:0;font-size:16px}
    details{padding:12px 0;border-bottom:1px solid #eee}summary{cursor:pointer;font-size:22px;font-weight:bold}
    .back-link{position:fixed;bottom:16px;right:16px;background:#555;color:white;padding:10px 18px;border-radius:12px;text-decoration:none}
    </style></head><body><main class="reports-wrap">
    <h1 class="page-title">📚 观音堂资料中心</h1>
    <section class="report-card"><h2>📊 实时月报</h2><p>2026年11月起，下载时从云端生成最新报表。</p>
    <form action="{{ url_for('reports.download',kind='month') if authorized else url_for('reports.home') }}" method="get">
    <select id="period" aria-label="月报年月" onchange="var p=this.value.split('-');this.form.year.value=p[0];this.form.month.value=p[1];">
    {% for y,m in periods|reverse %}<option value="{{y}}-{{m}}">{{y}}年{{m}}月</option>{% endfor %}</select>
    <input type="hidden" name="year" value="{{now.year}}"><input type="hidden" name="month" value="{{now.month}}">
    <button class="btn-tool btn-blue" type="submit">{{ '📥 下载月报 Excel' if authorized else '🔐 验证后下载月报' }}</button></form></section>
    <section class="report-card"><h2>📈 实时年报</h2><p>2026年起，下载时重新生成全年出席报表。</p>
    <form action="{{ url_for('reports.download',kind='year') if authorized else url_for('reports.home') }}" method="get">
    <select name="year" aria-label="年报年份">{% for y in range(now.year,2025,-1) %}<option value="{{y}}">{{y}}年</option>{% endfor %}</select>
    <button class="btn-tool btn-blue" type="submit">{{ '📥 下载年报 Excel' if authorized else '🔐 验证后下载年报' }}</button></form></section>
    <section class="report-card"><h2>🗂 历史报表</h2><p>保留原先上传的历史 Excel 文件。</p>
    {% if storage_error %}<p>历史文件暂时无法完整读取，请稍后重试。</p>{% endif %}
    {% for year,items in history %}<details><summary>{{year}} 年（{{items|length}} 份）</summary>
    {% for f in items %}<div class="file-row"><b>{{f.display_name}}</b><a class="btn-tool btn-blue" href="{{f.url}}">下载</a></div>{% endfor %}</details>
    {% else %}<p>暂无可显示的历史报表。</p>{% endfor %}</section>
    <section class="report-card"><h2>📁 其他资料</h2>
    {% for f in other_files %}<div class="file-row"><b>{{f.display_name}}</b><a class="btn-tool btn-blue" href="{{f.url}}">下载</a></div>{% else %}<p>暂无其他 Excel 资料。</p>{% endfor %}</section>
    <a class="back-link" href="/schedule/admin">⬅ 返回</a></main></body></html>
    """, now=now, periods=periods, authorized=authorized,
        history=sorted(history.items(), reverse=True), other_files=other_files,
        storage_error=storage_error)


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
        lines.extend(["", "感恩以上师兄发心护持观音堂！", "🙏🙏🙏", "", "义工报名链接：", "https://gyt-checkin-sg.onrender.com/volunteer"])
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


def _notice_center_roster_ready(date_str):
    """22:00 (MYT) 以后且该日期正式发布成功，才显示正式值班表。"""
    try:
        target_day = date.fromisoformat(date_str)
    except (ValueError, TypeError):
        return False
    deadline = datetime.combine(target_day - timedelta(days=1),
                                datetime.min.time(),
                                tzinfo=ZoneInfo('Asia/Kuala_Lumpur')) + timedelta(hours=22)
    return malaysia_now() >= deadline and is_schedule_published(date_str)


@schedule_bp.route('/schedule/notice_center/roster_status')
def notice_center_roster_status():
    if not session.get('schedule_login'):
        return jsonify({'error': 'unauthorized'}), 401
    date_str = request.args.get('date', '')
    return jsonify({'ready': _notice_center_roster_ready(date_str)})


@schedule_bp.route('/schedule/notice_center/official_roster')
def notice_center_official_roster():
    if not session.get('schedule_login'):
        return redirect(url_for('schedule.schedule_admin'))
    date_str = request.args.get('date', '')
    if not _notice_center_roster_ready(date_str):
        return redirect(url_for('schedule.schedule_notice_center', date=date_str))
    # 仅展示/复制；不触发重新排班、发布、重发或清除更新提醒。
    output = build_whatsapp_from_assigned(date_str)
    return render_template_string(DAY_OUTPUT_HTML, output=output)


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
    official_roster_ready = _notice_center_roster_ready(date_str)
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

<a id="mainNoticeButton" class="notice-btn {{ 'green' if official_roster_ready else 'blue' }}"
   href="{{ '/schedule/notice_center/official_roster?date=' ~ date_str if official_roster_ready else '/schedule/signup_notice?date=' ~ date_str }}">
    <span id="mainNoticeTitle">{{ '📋 正式值班表' if official_roster_ready else '📢 报名通知' }}</span><br>
    <small id="mainNoticeTime">{{ '🕙 晚上 10:00 发群' if official_roster_ready else '🕖 晚上 7:00 发群' }}</small>
</a>
<script>
(function () {
  const dateStr = {{ date_str | tojson }};
  const button = document.getElementById('mainNoticeButton');
  async function refreshNoticeButton() {
    try {
      const response = await fetch('/schedule/notice_center/roster_status?date=' + encodeURIComponent(dateStr),
                                   {cache: 'no-store'});
      if (!response.ok) return;
      const data = await response.json();
      if (data.ready) {
        button.classList.remove('blue');
        button.classList.add('green');
        button.href = '/schedule/notice_center/official_roster?date=' + encodeURIComponent(dateStr);
        document.getElementById('mainNoticeTitle').textContent = '📋 正式值班表';
        document.getElementById('mainNoticeTime').textContent = '🕙 晚上 10:00 发群';
      }
    } catch (e) { /* 网络异常时保持原状态，不误切换 */ }
  }
  setInterval(refreshNoticeButton, 30000);
})();
</script>

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
        late_notice_count=late_notice_count, return_section=return_section,
        official_roster_ready=official_roster_ready)


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

    # 固定佛台由 Excel 轮值和负责人当天覆盖名单决定，不一定有普通排班 ID。
    # 旧签到记录即使 walk_in=True，仍按当天有效名单在页面上归入正式签到。
    fixed_altar_names = {
        "".join(cc.convert(str(name or "")).split()).casefold()
        for name in get_fixed_buddha_for_date(date_str)
    }

    def is_fixed_altar_checkin(log):
        if str(log.get("actual_role") or "").strip() not in ("佛台", "整理佛台"):
            return False
        name = "".join(cc.convert(str(log.get("name") or "")).split()).casefold()
        return bool(name) and name in fixed_altar_names

    # 固定佛台排班可能无编号；沿用已确认的固定名单姓名匹配。
    altar_logs = [r for r in logs if r.get("check_in_time") and is_fixed_altar_checkin(r)]
    altar_names = {"".join(cc.convert(str(r.get("name") or "")).split()).casefold() for r in altar_logs}
    matched_altar_logs = set()
    for a in assignments:
        name = "".join(cc.convert(str(a.get("name") or "")).split()).casefold()
        altar_matched = (
            str(a.get("role") or "").strip() in ("佛台", "整理佛台")
            and bool(name) and name in altar_names
        )
        if altar_matched:
            matched_altar_logs.add(name)
        if a["id"] in signed_assignment_ids or altar_matched:
            checked_in_raw.append(a)
        else:
            not_checked_in_raw.append(a)

    fixed_altar_logs = []
    for l in logs:
        if is_fixed_altar_checkin(l):
            name = "".join(cc.convert(str(l.get("name") or "")).split()).casefold()
            if name not in matched_altar_logs:
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
                href="/schedule?mode=day&amp;override_date={{ date_str }}">
                    ⬅ 返回当天安排
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


# 负责人首页合并入口：仅提供导航，不修改原有功能或保存逻辑。






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
                        href="/schedule/management/settings">
                        返回系统设置
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

    rows, summary, role_summary, today, tomorrow = load_management_signup_view(
        status, date_filter, role_filter, show_history
    )

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
    /* 报名管理：仅更换页面底色 */
    body { background-color: #F0F8F1 !important; }

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
        right:18px;
        bottom:18px;
        z-index:9999;
        display:inline-flex;
        align-items:center;
        justify-content:center;
        width:auto !important;
        min-width:0 !important;
        padding:10px 15px !important;
        font-size:15px !important;
        border-radius:999px;
        box-shadow:0 4px 12px rgba(0,0,0,.2);
        text-decoration:none;
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
            right:12px;
            bottom:12px;
            width:auto !important;
            margin:0;
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

        <a href="/schedule#management-functions"
        class="btn-tool btn-blue floating-back">
            ⬅ 返回
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






# 固定岗位人员管理：独立页面，不影响现有排班、佛台及通知功能。
_FIXED_ROLE_OPTIONS = ("供花", "供果", "佛学班", "膳食")



@schedule_bp.route("/schedule/fixed-roles/search", methods=["GET"])
def schedule_fixed_roles_search():
    # 仅供已登录负责人搜索在册义工；不修改任何固定岗位名单。
    if not session.get("schedule_login"):
        return jsonify({"results": []}), 401
    keyword = request.args.get("q", "").strip()[:60]
    branch = request.args.get("branch", "CHE").strip().upper()
    if not keyword or branch not in ("CHE", "STW", "8XX"):
        return jsonify({"results": []})
    try:
        # 姓名允许部分匹配；编号仅匹配所选类别，最终写入仍由原 POST 校验。
        if branch == "8XX":
            condition = "v.id ~ '^8[0-9]+$'"
        else:
            condition = "v.id LIKE %s"
        params = [] if branch == "8XX" else [branch + "-%"]
        if keyword.isdigit():
            params.append((keyword + "%") if branch == "8XX" else (branch + "-" + keyword + "%"))
            search_condition = "v.id LIKE %s"
        else:
            params.append("%" + keyword + "%")
            search_condition = "v.name ILIKE %s"
        rows = db_query("SELECT v.id, v.name FROM volunteers v WHERE v.status='在册' AND "
                        + condition + " AND " + search_condition
                        + " ORDER BY v.name, v.id LIMIT 15", tuple(params), fetchall=True) or []
        return jsonify({"results": [{"id": str(r["id"]), "name": str(r["name"])} for r in rows]})
    except Exception:
        return jsonify({"results": [], "error": "查询暂时不可用"}), 503


@schedule_bp.route("/schedule/fixed-roles", methods=["GET", "POST"])
def schedule_fixed_roles():
    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    if request.method == "POST":
        action = request.form.get("action", "").strip()
        role = request.form.get("role", "").strip()
        raw_id = request.form.get("volunteer_id", "").strip().upper()
        # 佛台每星期固定轮值：独立数据表，不影响原有三种固定岗位。
        if action in ("buddha_add", "buddha_remove"):
            weekday_text = request.form.get("weekday", "").strip()
            if weekday_text not in tuple(str(i) for i in range(7)) or not raw_id:
                flash("星期或义工编号不正确。", "error")
                return redirect(url_for("schedule.schedule_fixed_roles"))
            weekday = int(weekday_text)
            try:
                with get_conn() as conn:
                    with conn.cursor(cursor_factory=RealDictCursor) as cur:
                        if action == "buddha_add":
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
                            # 同一星期锁定写入，避免并发加入导致 position 重复。
                            cur.execute("SELECT pg_advisory_xact_lock(382716, %s)", (weekday,))
                            cur.execute("""
                                SELECT volunteer_id, position FROM schedule_fixed_buddha_weekly
                                WHERE weekday=%s ORDER BY position
                            """, (weekday,))
                            existing = cur.fetchall()
                            if any(str(row["volunteer_id"]) == str(volunteer["id"]) for row in existing):
                                flash("这位义工已在该星期的佛台固定名单中。", "ok")
                            elif len(existing) >= 3:
                                flash("每个星期最多安排 3 位固定佛台义工。", "error")
                            else:
                                used = {int(row["position"]) for row in existing}
                                position = next(n for n in range(1, 4) if n not in used)
                                cur.execute("""
                                    INSERT INTO schedule_fixed_buddha_weekly
                                    (weekday, volunteer_id, position) VALUES (%s, %s, %s)
                                """, (weekday, volunteer["id"], position))
                                flash(f"已加入星期{'一二三四五六日'[weekday]}：{volunteer['name']}", "ok")
                        else:
                            cur.execute("""
                                DELETE FROM schedule_fixed_buddha_weekly
                                WHERE weekday=%s AND volunteer_id=%s
                            """, (weekday, raw_id))
                            flash("已移除该星期的佛台固定人员。" if cur.rowcount else "该星期没有找到对应人员。", "ok")
                    conn.commit()
            except Exception:
                flash("佛台固定轮值保存失败，请检查数据库表及字段。", "error")
            return redirect(url_for("schedule.schedule_fixed_roles"))

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
    buddha_members = {i: [] for i in range(7)}
    try:
        buddha_rows = db_query("""
            SELECT f.weekday, f.position, f.volunteer_id,
                   COALESCE(v.name, '义工资料已不存在') AS name
            FROM schedule_fixed_buddha_weekly f
            LEFT JOIN volunteers v ON v.id=f.volunteer_id
            ORDER BY f.weekday, f.position
        """, fetchall=True) or []
        for row in buddha_rows:
            buddha_members[int(row["weekday"])].append(row)
    except Exception:
        flash("读取佛台固定轮值失败，请检查 Supabase 数据表。", "error")
    return render_template_string(_FIXED_ROLES_HTML, roles=_FIXED_ROLE_OPTIONS,
                                  members=members, buddha_members=buddha_members)





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
<a class="btn secondary" href="/schedule/management/settings">← 返回系统设置 / Back</a></main></body></html>''',flags=flags,current=current,notice_time=notice_time,jobs=jobs,message=message,queue_error=queue_error)


# 第五阶段：注册独立管理菜单路由（保留原 Blueprint 和 endpoint 名称）
from schedule.routes import management_menu_routes  # noqa: F401

# 第六阶段：注册独立负责人辅助路由（原 endpoint 名称不变）
from schedule.routes import admin_utility_routes  # noqa: F401

from schedule.services.flower_fruit_routes import register_flower_fruit_routes
register_flower_fruit_routes(schedule_bp)
