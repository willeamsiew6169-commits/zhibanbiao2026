# schedule_web.py

import os
import psycopg2
import pandas as pd

from db import get_db
from opencc import OpenCC
from openpyxl import load_workbook
from datetime import datetime, timedelta, date, timezone
from psycopg2.extras import RealDictCursor
from sqlalchemy import create_engine, text

from monthly_prebook_message import generate_monthly_prebook_message
from flask import Blueprint, request, session, redirect, url_for, render_template_string
from schedule_builder import run_schedule_for_date, parse_signup_line_multi, normalize_time_text, build_buddhist_festival_message, get_special_day_info, get_next_day_remove_info, build_lunar_1_15_message, build_normal_message


cc = OpenCC('t2s')  # 繁 → 简

schedule_bp = Blueprint("schedule", __name__)

DATABASE_URL = os.environ.get("DATABASE_URL")
engine = create_engine(DATABASE_URL, pool_pre_ping=True) if DATABASE_URL else None

DATABASE_URL = os.environ.get("DATABASE_URL")

def get_conn():
    return psycopg2.connect(DATABASE_URL) 

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PREBOOK_FILE = os.path.join(BASE_DIR, "prebook_schedule.xlsx")

SCHEDULE_PIN = "1234"

TIME_OPTIONS = [
    "6:00am", "6:30am", "7:00am", "7:30am",
    "8:00am", "8:30am", "9:00am", "9:30am",
    "10:00am", "10:30am", "11:00am", "11:30am",
    "12:00pm", "12:30pm", "1:00pm", "1:30pm",
    "2:00pm", "2:30pm", "3:00pm", "3:30pm",
    "4:00pm", "4:30pm", "5:00pm", "5:30pm", "6:00pm"
]

ROLE_OPTIONS = [
    "值班",
    "卫生",
]

schedule_records = []

TARGETS = {
    "橙观音堂": 2,
    "橙活动中心": 2,
    "黄观音堂": 2,
    "黄活动中心": 2,
    "卫生": 2,
}

def build_shortage_notice_from_assignments(date_str):

    rows = load_assigned_places_for_date(date_str)

    counts = {
        "橙观音堂": 0,
        "橙活动中心": 0,
        "黄观音堂": 0,
        "黄活动中心": 0,
        "卫生": 0,
    }

    for r in rows:

        role = r.get("role")
        shift = r.get("shift_label")
        place = r.get("assigned_place")

        if role == "卫生":
            counts["卫生"] += 1

        elif role == "值班":

            if shift and place:

                key = f"{shift.replace('班','')}{place}"

                if key in counts:
                    counts[key] += 1

    shortages = []

    for key, target in TARGETS.items():

        current = counts.get(key, 0)

        if current < target:

            shortages.append(
                f"🔴 {key}：缺 {target-current} 位"
            )

        elif current == target:

            shortages.append(
                f"🟢 {key}：已满"
            )

        else:

            shortages.append(
                f"🟢 {key}：充足"
            )

    msg = "师兄们，大家好！\n\n"
    msg += "明天义工岗位情况：\n\n"

    msg += "\n".join(shortages)

    msg += "\n\n欢迎大家随缘发心护持观音堂。\n\n感恩大家 🙏🙏🙏"

    return msg


def to_simple(text):
    if not text:
        return ""
    return cc.convert(str(text).strip())


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


def load_schedule_admin_dashboard_data(override_date):
    from datetime import date, timedelta
    from psycopg2.extras import RealDictCursor

    today = date.today()
    tomorrow = today + timedelta(days=1)
    month_start = today.replace(day=1)

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:

            # 未安排统计
            cur.execute("""
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
            """, (today, tomorrow, month_start, month_start))

            pending_row = cur.fetchone()

            pending_counts = {
                "today": pending_row["today_count"] or 0,
                "tomorrow": pending_row["tomorrow_count"] or 0,
                "month": pending_row["month_count"] or 0,
            }

            # 当天概况
            cur.execute("""
                select
                    count(*) as total,
                    count(*) filter (
                        where coalesce(status, 'pending') = 'pending'
                    ) as pending,
                    count(*) filter (
                        where coalesce(status, '') = 'assigned'
                    ) as assigned,
                    count(*) filter (
                        where coalesce(status, '') = 'cancelled'
                    ) as cancelled
                from volunteer_schedule_signups
                where signup_date = %s
            """, (override_date,))

            day_summary = cur.fetchone()

            # 供台设置
            cur.execute("""
                select
                    flag_date,
                    coalesce(need_setup_master_table, false) as need_setup_master_table,
                    coalesce(need_remove_master_table, false) as need_remove_master_table,
                    coalesce(setup_people, '') as setup_people,
                    coalesce(remove_people, '') as remove_people,
                    remarks
                from schedule_day_flags
                where flag_date = %s
            """, (override_date,))

            day_flags = cur.fetchone()

    if not day_flags:
        day_flags = {
            "need_setup_master_table": False,
            "need_remove_master_table": False,
            "setup_people": "",
            "remove_people": "",
            "remarks": "",
        }

    return pending_counts, day_summary, day_flags


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


def time_to_minutes(t):
    from datetime import datetime

    if not t:
        return None

    s = str(t).strip().lower().replace(" ", "")

    if "am" in s or "pm" in s:
        if ":" not in s:
            s = s.replace("am", ":00am").replace("pm", ":00pm")

        try:
            dt = datetime.strptime(s, "%I:%M%p")
            return dt.hour * 60 + dt.minute
        except:
            return None

    return None


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

def to_simple(text):
    if not text:
        return ""
    return cc.convert(str(text).strip())

def normalize_vol_id_for_search(s):
    s = str(s or "").strip().upper()
    if not s:
        return ""

    if s.isdigit():
        if s.startswith("0"):
            return "STW-" + s.lstrip("0")
        return "CHE-" + s

    return s


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


def find_volunteer_by_keyword(keyword):
    keyword = str(keyword or "").strip()
    if not keyword:
        return []

    key_simple = to_simple(keyword)
    keyword_id = normalize_vol_id_for_search(keyword)

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                select id, name
                from volunteers
                where status = '在册'
            """)
            volunteers = cur.fetchall()

    # 1. 编号完全符合 → 直接返回这个人
    id_matches = []
    for v in volunteers:
        vid = str(v["id"] or "").strip()
        if vid == keyword_id or vid == keyword.upper():
            id_matches.append(v)

    if id_matches:
        return id_matches

    # 2. 姓名简繁体模糊匹配 → 返回全部
    name_matches = []
    for v in volunteers:
        name = str(v["name"] or "").strip()
        if key_simple in to_simple(name):
            name_matches.append(v)

    return name_matches
    
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

    if role in ["卫生", "佛台"]:
        return "8:00am", "10:00am"

    if role == "供台":
        return "6:00am", "8:00am"

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

def load_display_records(mode, target_date=None, year=None, month=None):
    if not os.path.exists(PREBOOK_FILE):
        return []

    try:
        df = pd.read_excel(PREBOOK_FILE, sheet_name="预报名")
        df.columns = df.columns.astype(str).str.strip()

        if df.empty:
            return []

        df["日期"] = pd.to_datetime(df["日期"], errors="coerce")
        df = df.dropna(subset=["日期"])

        if mode == "day" and target_date:
            target = pd.to_datetime(target_date).date()
            df = df[df["日期"].dt.date == target]

        elif mode == "prebook" and year and month:
            df = df[
                (df["日期"].dt.year == int(year)) &
                (df["日期"].dt.month == int(month))
            ]

        else:
            return []

        df["日期"] = df["日期"].dt.strftime("%Y-%m-%d")

        return df.fillna("").to_dict("records")

    except Exception as e:
        print("load_display_records error:", e)
        return []
    

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


def build_whatsapp_from_assigned(date_str):
    from datetime import datetime

    rows = load_assigned_places_for_date(date_str)

    if not rows:
        return f"📅 {date_str}\n\n暂时没有已安排的值班资料。"

    date_obj = datetime.strptime(date_str, "%Y-%m-%d")
    special_info = get_special_day_info(date_obj)
    remove_info = get_next_day_remove_info(date_obj)

    arranged = {
        "整理佛台": [],
        "佛堂卫生": [],
        "二楼卫生": [],
        "楼梯卫生": [],
        "设师父供台": [],
        "绿观音堂": [],
        "绿活动中心": [],
        "橙观音堂": [],
        "橙活动中心": [],
        "黄观音堂": [],
        "黄活动中心": [],
    }

    for r in rows:
        name = r.get("name")
        role = r.get("role")
        place = r.get("assigned_place")
        shift = r.get("shift_label")
        start = r.get("start_time")
        end = r.get("end_time")

        if not name:
            continue

        if role == "整理佛台" or place == "整理佛台":
            arranged["整理佛台"].append(name)

        elif role == "卫生" and place in ["佛堂卫生", "二楼卫生", "楼梯卫生"]:
            arranged[place].append(name)

        elif role == "供台" or place == "设师父供台":
            arranged["设师父供台"].append(name)

        elif role == "值班":
            shift_key = str(shift or "").replace("班", "")

            is_lunar_day = special_info["template_type"] == "lunar_1_15"

            # 普通日不要绿班
            if shift_key == "绿" and not is_lunar_day:
                shift_key = "黄"

            if shift_key in ["绿", "橙", "黄"] and place in ["观音堂", "活动中心"]:
                key = f"{shift_key}{place}"
                arranged[key].append((name, start, end))

    
    if special_info["template_type"] == "normal":
        return build_normal_message(date_obj, arranged, special_info, remove_info)

    elif special_info["template_type"] == "lunar_1_15":
        return build_lunar_1_15_message(date_obj, arranged, special_info, remove_info)

    else:
        return build_buddhist_festival_message(date_obj, arranged, special_info, remove_info)


@schedule_bp.route("/schedule/copy_whatsapp", methods=["POST"])
def schedule_copy_whatsapp():
    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    date_str = request.form.get("date", "").strip()

    if not date_str:
        return "❌ 没有日期<br><a href='/schedule?mode=day'>返回</a>"

    output = build_whatsapp_from_assigned(date_str)

    return render_template_string(DAY_OUTPUT_HTML, output=output)


@schedule_bp.route("/volunteer")
def volunteer_home():

    MY_TZ = timezone(timedelta(hours=8))
    now = datetime.now(MY_TZ)

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
        times=TIME_OPTIONS
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
                values (%s, %s, %s, %s, %s, %s, 'pending', '义工网页报名')
            """, (
                vol_id,
                name,
                signup_date,
                role,
                start_time,
                end_time
            ))

            conn.commit()

    return redirect(url_for("schedule.volunteer_day_schedule", date=signup_date))

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


def build_signup_shortage_notice(date_str):
    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                select role, start_time, end_time
                from volunteer_schedule_signups
                where signup_date = %s
                and coalesce(status, 'pending') <> 'cancelled'
            """, (date_str,))
            rows = cur.fetchall()

    cleaning_count = 0
    morning_count = 0
    afternoon_count = 0
    full_day_count = 0

    for r in rows:
        role = r.get("role")
        start = r.get("start_time") or ""
        end = r.get("end_time") or ""

        if role == "卫生":
            cleaning_count += 1

        if role == "值班":
            s = time_to_minutes(start)
            e = time_to_minutes(end)

            if s is None or e is None:
                continue

            if s <= 10 * 60 and e >= 14 * 60:
                morning_count += 1

            if s <= 14 * 60 and e >= 18 * 60:
                afternoon_count += 1

            if s <= 10 * 60 and e >= 18 * 60:
                full_day_count += 1

    notices = []

    if cleaning_count < 2:
        need = 2 - cleaning_count
        notices.append(f"🧹 卫生 可以多加 {need} 位义工。")

    total_duty = morning_count + afternoon_count

    if total_duty < 6:
        notices.append("🏠 全日值班 可以多加 2~4 位义工。")

    if morning_count < 3:
        notices.append("⏰ 10:00am~2:00pm 值班人数还不足。")

    if afternoon_count < 3:
        notices.append("⏰ 2:00pm~6:00pm 值班人数还不足。")

    if not notices:
        return """
<div style="
    background:#e8fff0;
    border:1px solid #9be7aa;
    color:#1b5e20;
    padding:15px;
    border-radius:12px;
    font-size:22px;
    margin-bottom:18px;
">
✅ 明日义工报名人数暂时足够，感恩大家发心护持。
</div>
"""

    target_date = datetime.strptime(date_str, "%Y-%m-%d").date()
    today = date.today()

    if target_date == today:
        title = "📢 今日仍需要义工："
    elif target_date == today + timedelta(days=1):
        title = "📢 明日仍需要义工："
    else:
        title = f"📢 {date_str} 仍需要义工："

    html = f"""
    <div style="background:#fff3cd; padding:18px; border-radius:12px; font-size:24px; color:#7a5a00;">
    <b>{title}</b><br>
    """

    for n in notices:
        html += f"{n}<br>"

    html += """
<br>
感恩大家🙏🙏🙏
</div>
"""

    return html


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


def load_day_flags(date_str):
    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                select
                    flag_date,
                    coalesce(need_setup_master_table, false) as need_setup_master_table,
                    coalesce(need_remove_master_table, false) as need_remove_master_table,
                    coalesce(setup_people, '') as setup_people,
                    coalesce(remove_people, '') as remove_people,
                    remarks
                from schedule_day_flags
                where flag_date = %s
            """, (date_str,))
            row = cur.fetchone()

    if not row:
        return {
            "need_setup_master_table": False,
            "need_remove_master_table": False,
            "setup_people": "",
            "remove_people": "",
            "remarks": "",
        }

    return row


@schedule_bp.route("/schedule", methods=["GET", "POST"])
@schedule_bp.route("/schedule/admin", methods=["GET", "POST"])
def schedule_admin():
    from datetime import datetime, date
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

    change_time = now.replace(hour=18, minute=0, second=0, microsecond=0)

    if now >= change_time:
        default_schedule_date = (now + timedelta(days=1)).strftime("%Y-%m-%d")
    else:
        default_schedule_date = now.strftime("%Y-%m-%d")

    mode = request.args.get("mode", "")

    override_date = request.args.get("override_date") or default_schedule_date

    t1 = time.time()
    buddha_names = load_buddha_name_options()
    print("load_buddha_name_options:", round(time.time() - t1, 2))

    t2 = time.time()
    fixed_buddha_today = get_fixed_buddha_for_date(override_date)
    print("get_fixed_buddha_for_date:", round(time.time() - t2, 2))

    t3 = time.time()
    pending_counts, day_summary, day_flags = load_schedule_admin_dashboard_data(override_date)
    print("load_schedule_admin_dashboard_data:", round(time.time() - t3, 2))

    selected_date_obj = datetime.strptime(override_date, "%Y-%m-%d").date()
    is_today_or_past = selected_date_obj < date.today()

    today = date.today()

    if today.month == 12:
        next_year = today.year + 1
        next_month = 1
    else:
        next_year = today.year
        next_month = today.month + 1

    t6 = time.time()
    display_records = load_display_records(
        mode=mode,
        target_date=override_date if mode == "day" else None,
        year=next_year if mode == "prebook" else None,
        month=next_month if mode == "prebook" else None,
    )
    print("load_display_records:", round(time.time() - t6, 2))

    print("schedule_admin total:", round(time.time() - t0, 2))

    return render_template_string(
        SCHEDULE_HTML,
        mode=mode,
        times=TIME_OPTIONS,
        roles=ROLE_OPTIONS,
        records=display_records,
        tomorrow=default_schedule_date,
        buddha_names=buddha_names,
        override_date=override_date,
        fixed_buddha_today=fixed_buddha_today,
        default_year=next_year,
        default_month=next_month,
        pending_counts=pending_counts,
        day_summary=day_summary,
        is_today_or_past=is_today_or_past,
        day_flags=day_flags,
    )


@schedule_bp.route("/volunteer/today_schedule")
def volunteer_today_schedule():

    MY_TZ = timezone(timedelta(hours=8))
    now = datetime.now(MY_TZ)

    if now.hour >= 18:
        target_date = (now + timedelta(days=1)).strftime("%Y-%m-%d")
    else:
        target_date = now.strftime("%Y-%m-%d")

    return redirect(url_for(
        "schedule.volunteer_day_schedule",
        date=target_date
    ))


@schedule_bp.route("/schedule/add", methods=["POST"])
def schedule_add():
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

    year = int(request.form.get("year"))
    month = int(request.form.get("month"))
    keyword = request.form.get("vol_id", "").strip()

    days = request.form.getlist("days")
    roles = request.form.getlist("roles")

    start_time = request.form.get("start_time", "").strip()
    end_time = request.form.get("end_time", "").strip()

    if not keyword:
        return "❌ 请输入义工编号 / 姓名<br><a href='/schedule?mode=prebook'>返回</a>"

    if not days:
        return "❌ 请选择日期<br><a href='/schedule?mode=prebook'>返回</a>"

    if not roles:
        return "❌ 请选择岗位<br><a href='/schedule?mode=prebook'>返回</a>"

    matches = find_volunteer_by_keyword(keyword)

    if not matches:
        return "❌ 找不到义工<br><a href='/schedule?mode=prebook'>返回</a>"

    if len(matches) > 1:
        return "❌ 找到多个同名义工，请用义工编号查询<br><a href='/schedule?mode=prebook'>返回</a>"

    vol = matches[0]
    volunteer_id = str(vol["id"])
    name = str(vol["name"])

    inserted = 0
    skipped = 0

    with get_db() as conn:
        with conn.cursor() as cur:

            for d in days:
                day = int(d)

                try:
                    signup_date = date(year, month, day)
                except ValueError:
                    skipped += 1
                    continue

                for role in roles:

                    role_start = start_time if role == "值班" else None
                    role_end = end_time if role == "值班" else None

                    if role == "值班" and (not role_start or not role_end):
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
                            '管理员月预报名',
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
    <h1>✅ 月预报名已加入</h1>
    <p>义工：{name}</p>
    <p>成功加入：{inserted} 笔</p>
    <p>跳过重复/无效：{skipped} 笔</p>
    <a href="/schedule?mode=prebook">继续加入</a><br>
    <a href="/schedule/admin">返回后台</a>
    """    


@schedule_bp.route("/volunteer/prebook", methods=["GET", "POST"])
def volunteer_prebook():

    if request.method == "GET":
        today = date.today()

        return render_template_string(
            VOLUNTEER_PREBOOK_HTML,
            default_year=today.year,
            default_month=today.month,
            times=TIME_OPTIONS
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

                else:
                    cur.execute("""
                        insert into volunteer_schedule_signups
                        (volunteer_id, name, signup_date, role, start_time, end_time, status, remarks)
                        values (%s, %s, %s, %s, %s, %s, 'pending', '义工多日报名')
                    """, (
                        vol_id,
                        name,
                        signup_date,
                        role,
                        start_time,
                        end_time
                    ))
                    inserted += 1

            conn.commit()

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


@schedule_bp.route("/schedule/monthly_prebook", methods=["POST"])
def schedule_monthly_prebook():
    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    year = request.form.get("year", "").strip()
    month = request.form.get("month", "").strip()

    try:
        output = generate_monthly_prebook_message(int(year), int(month))
    except Exception as e:
        output = f"❌ 生成失败：{e}"

    return render_template_string(MONTHLY_PREBOOK_HTML, output=output)


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

    return redirect(url_for("schedule.schedule", mode=mode))


def load_assigned_places_for_date(date_str):
    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                select
                    id,
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
                    remarks
                from volunteer_schedule_assignments
                where assignment_date = %s
                and coalesce(status, 'assigned') <> 'cancelled'
                order by
                    case
                        when role = '卫生' then 1
                        when role = '供台' then 2
                        when role = '值班' then 3
                        else 9
                    end,
                    case
                        when shift_label = '绿班' then 1
                        when shift_label = '橙班' then 2
                        when shift_label = '黄班' then 3
                        else 9
                    end,
                    assigned_place,
                    start_time,
                    name
            """, (date_str,))

            return cur.fetchall()
        

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

    need_setup = request.form.get("need_setup_master_table") == "1"
    need_remove = request.form.get("need_remove_master_table") == "1"

    setup_people = request.form.get("setup_people", "").strip()
    remove_people = request.form.get("remove_people", "").strip()

    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                insert into schedule_day_flags (
                    flag_date,
                    need_setup_master_table,
                    need_remove_master_table,
                    setup_people,
                    remove_people,
                    updated_at
                )
                values (%s, %s, %s, %s, %s, now())
                on conflict (flag_date)
                do update set
                    need_setup_master_table = excluded.need_setup_master_table,
                    need_remove_master_table = excluded.need_remove_master_table,
                    setup_people = excluded.setup_people,
                    remove_people = excluded.remove_people,
                    updated_at = now()
            """, (
                date_str,
                need_setup,
                need_remove,
                setup_people,
                remove_people
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

    lines = raw_signup.splitlines()

    added = 0

    for line in lines:

        line = line.strip()

        if not line:
            continue

        parts = line.split()

        if len(parts) < 2:
            continue

        name = parts[0]

        time_text = parts[-1]

        start_time = "3:00pm"
        end_time = "6:00pm"

        if "-" in time_text:
            try:
                s, e = time_text.split("-", 1)

                start_time = s.strip()

                if "pm" not in e.lower() and "am" not in e.lower():
                    if "pm" in s.lower():
                        e += "pm"
                    elif "am" in s.lower():
                        e += "am"

                end_time = e.strip()

            except:
                pass

        record = {
            "日期": single_date,
            "编号": "",
            "姓名": name,
            "岗位": "值班",
            "开始时间": start_time,
            "结束时间": end_time,
            "备注": "WhatsApp",
        }

        schedule_records.append(record)

        save_prebook_record(record)

        added += 1

    print("🔥 parse_raw route 有进来")

    return redirect(url_for("schedule.schedule", mode="day"))


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

    return redirect(url_for("schedule.schedule", mode="day", override_date=date))


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

        // 新增
        const noticeDate = document.getElementById("notice_date");

        if (!mainDate) return;

        if (rawDate) rawDate.value = mainDate.value;
        if (genDate) genDate.value = mainDate.value;
        if (fillDate) fillDate.value = mainDate.value;
        if (viewAssignedDate) viewAssignedDate.value = mainDate.value;
        if (overrideDate) overrideDate.value = mainDate.value;
        if (prebookDate) prebookDate.value = mainDate.value;

        // 新增
        if (noticeDate) noticeDate.value = mainDate.value;
    }

    syncDates();

    const mainDate = document.getElementById("main_date");
    if (mainDate) {
        mainDate.addEventListener("change", syncDates);
    }
});
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
    </div>
</div>

<div class="warn-box">
    <h3 style="margin-top:0;">🔴 未安排提醒</h3>
    今日：<b>{{ pending_counts.today }}</b> 人　
    明日：<b>{{ pending_counts.tomorrow }}</b> 人　
    本月：<b>{{ pending_counts.month }}</b> 人
</div>

<hr>

{% if mode == "day" %}

<h2>📅 当天值班安排</h2>

<div class="card">
    <h3>📅 日期选择</h3>

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

<div class="section quick-panel">
    <h3 class="section-title">⚡ 快捷操作</h3>

    <div class="quick-actions">

        {% if not is_today_or_past %}
        <form method="post" action="/schedule/generate_day">
            <input type="hidden" name="date" id="generate_day_date" value="{{ override_date }}">
            <button type="submit" class="quick-btn primary">⚡ 自动排班</button>
        </form>
        {% else %}
        <form method="post" action="/schedule/fill_pending">
            <input type="hidden" name="date" id="fill_pending_date" value="{{ override_date }}">
            <button type="submit" class="quick-btn warn">➕ 补入待安排</button>
        </form>
        {% endif %}

        <form method="post" action="/schedule/view_assigned">
            <input type="hidden" name="date" id="view_assigned_date" value="{{ override_date }}">
            <button type="submit" class="quick-btn">👀 查看排班</button>
        </form>

        <form method="post" action="/schedule/copy_whatsapp">
            <input type="hidden" name="date" id="copy_whatsapp_date" value="{{ override_date }}">
            <button type="submit" class="quick-btn whatsapp">📋 WhatsApp</button>
        </form>

        <form method="post" action="/schedule/generate_shortage_notice">
            <input type="hidden" name="date" id="notice_date" value="{{ override_date }}">
            <button type="submit" class="quick-btn warn">
                📢 生成缺人工通知
            </button>
        </form>

        <a href="/schedule/attendance_status">
            <button type="button" class="quick-btn">📋 签到情况</button>
        </a>
        
        <a href="/schedule/reports" class="quick-btn">
            📊 报表中心
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

    <div class="two-col">

        <div class="sub-card">
            <h4>师父供台设置</h4>

            <form method="post" action="/schedule/save_day_flags">

                <input type="hidden" name="date" value="{{ override_date }}">

                <label style="font-size:22px;">
                    <input
                        type="checkbox"
                        name="need_setup_master_table"
                        value="1"
                        {% if day_flags.need_setup_master_table %}checked{% endif %}
                    >
                    明日需要设师父供台
                </label>

                <p>设供台人员：</p>
                <textarea
                    name="setup_people"
                    rows="3"
                    style="width:100%; font-size:22px;"
                    placeholder="例如：张三&#10;李四"
                >{{ day_flags.setup_people }}</textarea>

                <br><br>

                <label style="font-size:22px;">
                    <input
                        type="checkbox"
                        name="need_remove_master_table"
                        value="1"
                        {% if day_flags.need_remove_master_table %}checked{% endif %}
                    >
                    明日需要收师父供台
                </label>

                <p>收供台人员：</p>
                <textarea
                    name="remove_people"
                    rows="3"
                    style="width:100%; font-size:22px;"
                    placeholder="例如：张三"
                >{{ day_flags.remove_people }}</textarea>

                <br><br>

                <button type="submit" class="action-btn">💾 保存供台设置</button>
            </form>
        </div>

        <div class="sub-card">
            <h4>佛台请假 / 换人</h4>

            <p>日期：{{ override_date }}</p>

            <form method="post" action="/schedule/override">
                <input type="hidden" name="date" id="override_date" value="{{ override_date }}">

                {% if fixed_buddha_today %}
                    {% for old_name in fixed_buddha_today %}
                        <div style="font-size:22px; margin:10px 0;">
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

VOLUNTEER_SIGNUP_HTML = """
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>义工报名系统</title>

<style>
body {
    font-family:"Microsoft YaHei", Arial;
    background:#f5f5f5;
    padding:20px;
    font-size:24px;
}

.box {
    background:white;
    max-width:1200px;
    margin:auto;
    padding:25px;
    border-radius:15px;
}

.top-actions {
    display:flex;
    justify-content:center;
    gap:20px;
    flex-wrap:wrap;
    margin-bottom:18px;
}

.top-btn {
    display:flex;
    align-items:center;
    justify-content:center;
    width:280px;
    min-height:100px;
    border-radius:15px;
    color:white;
    text-decoration:none;
    font-size:28px;
    font-weight:bold;
    text-align:center;
    padding:10px;
    box-sizing:border-box;
}

.green { background:#4CAF50; }
.blue { background:#2196F3; }
.orange { background:#FF9800; }

.notice {
    background:#fff3cd;
    color:#856404;
    border:1px solid #ffeeba;
    padding:15px 18px;
    border-radius:10px;
    margin:0 auto 25px auto;
    max-width:900px;
    font-size:22px;
    line-height:1.6;
}

.admin {
    text-align:right;
    font-size:18px;
    margin-bottom:20px;
}

input, select, button {
    font-size:24px;
    padding:12px;
    margin:8px 0 18px 0;
    width:100%;
    box-sizing:border-box;
}

button {
    cursor:pointer;
    background:#4CAF50;
    color:white;
    border:0;
    border-radius:10px;
    font-weight:bold;
    padding:16px;
}
</style>
</head>

<body>
<div class="box">

<div class="top-actions">
    <a class="top-btn green" href="/volunteer/today_schedule">
        📋 {{ schedule_label }}
    </a>

    <a class="top-btn blue" href="/volunteer/prebook">
        📅 多日报名
    </a>

    <a class="top-btn orange" href="/volunteer/my_schedule_search">
        🔍 我的报名
    </a>
</div>

<div class="notice">
    ⚠️ 当前为报名状态，最终岗位安排请以负责人公布的正式值班表为准。
</div>

<div class="admin">
    <a href="/schedule/admin">🔐 管理员入口</a>
</div>

<h1>义工报名系统</h1>

<form method="post" action="/volunteer/signup">

<label>义工编号 / 电话 / 姓名</label>
<input name="keyword" required placeholder="例如 CHE-108 / 108 / 姓名">

<label>日期</label>
<input type="date" name="signup_date" value="{{ default_date }}" required>

<label>岗位</label>
<select name="role" required>
    <option value="值班">值班</option>
    <option value="卫生">卫生</option>
    <option value="供台">供台</option>
</select>

<label>开始时间</label>
<select name="start_time">
{% for t in times %}
<option value="{{ t }}">{{ t }}</option>
{% endfor %}
</select>

<label>结束时间</label>
<select name="end_time">
{% for t in times %}
<option value="{{ t }}">{{ t }}</option>
{% endfor %}
</select>

<button type="submit">提交报名</button>

</form>

</div>
</body>
</html>
"""

VOLUNTEER_PREBOOK_HTML = """
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>多日报名</title>

<style>
body {
    font-family:"Microsoft YaHei", Arial;
    background:#f5f5f5;
    padding:20px;
    font-size:24px;
}

.box {
    background:white;
    max-width:750px;
    margin:auto;
    padding:25px;
    border-radius:15px;
}

input, select, button {
    font-size:24px;
    padding:12px;
    margin:8px 0;
    width:100%;
    box-sizing:border-box;
}

.calendar {
    display:grid;
    grid-template-columns:repeat(7, 1fr);
    gap:10px;
    margin:15px 0;
}

.week-title {
    text-align:center;
    font-weight:bold;
    font-size:20px;
    color:#555;
}

.day-card {
    display:block;
    text-align:center;
    border:2px solid #ccc;
    border-radius:12px;
    padding:14px 0;
    font-size:24px;
    background:white;
    cursor:pointer;
}

.day-card input {
    display:none;
}

.day-card span {
    display:block;
}

.day-card.checked {
    background:#2196F3;
    color:white;
    border-color:#2196F3;
    font-weight:bold;
}

.empty-day {
    height:55px;
}

a {
    font-size:22px;
}
</style>

<script>
function updateCalendar() {
    const year = parseInt(document.getElementById("year").value);
    const month = parseInt(document.getElementById("month").value);
    const calendarDays = document.getElementById("calendar-days");

    calendarDays.innerHTML = "";

    if (!year || !month) return;

    const firstDay = new Date(year, month - 1, 1).getDay();
    const lastDate = new Date(year, month, 0).getDate();

    for (let i = 0; i < firstDay; i++) {
        const empty = document.createElement("div");
        empty.className = "empty-day";
        calendarDays.appendChild(empty);
    }

    for (let d = 1; d <= lastDate; d++) {
        const label = document.createElement("label");
        label.className = "day-card";

        const input = document.createElement("input");
        input.type = "checkbox";
        input.name = "days";
        input.value = d;

        const span = document.createElement("span");
        span.innerText = d;

        input.addEventListener("change", function () {
            if (input.checked) {
                label.classList.add("checked");
            } else {
                label.classList.remove("checked");
            }
        });

        label.appendChild(input);
        label.appendChild(span);
        calendarDays.appendChild(label);
    }
}

document.addEventListener("DOMContentLoaded", function () {
    updateCalendar();

    document.getElementById("year").addEventListener("change", updateCalendar);
    document.getElementById("month").addEventListener("change", updateCalendar);
});
</script>

</head>

<body>
<div class="box">

<h1>📅 多日报名</h1>

<form method="post" action="/volunteer/prebook">

<label>义工编号 / 电话 / 姓名</label>
<input name="keyword" required placeholder="例如 CHE-108 / 108 / 姓名">

<label>年份</label>
<input id="year" name="year" value="{{ default_year }}" required>

<label>月份</label>
<select id="month" name="month">
{% for m in range(1, 13) %}
<option value="{{ m }}" {% if m == default_month %}selected{% endif %}>{{ m }}月</option>
{% endfor %}
</select>

<h3>选择日期</h3>

<div class="calendar">
    <div class="week-title">日</div>
    <div class="week-title">一</div>
    <div class="week-title">二</div>
    <div class="week-title">三</div>
    <div class="week-title">四</div>
    <div class="week-title">五</div>
    <div class="week-title">六</div>
</div>

<div id="calendar-days" class="calendar"></div>

<label>岗位</label>
<select name="role" required>
    <option value="值班">值班</option>
    <option value="卫生">卫生</option>
    <option value="供台">供台</option>
</select>

<label>开始时间（值班才需要）</label>
<select name="start_time">
{% for t in times %}
<option value="{{ t }}">{{ t }}</option>
{% endfor %}
</select>

<label>结束时间（值班才需要）</label>
<select name="end_time">
{% for t in times %}
<option value="{{ t }}">{{ t }}</option>
{% endfor %}
</select>

<button type="submit">提交多日报名</button>

</form>

<br>
<a href="/volunteer">返回</a>

</div>
</body>
</html>
"""