# schedule_builder.py

import os
import re
import pandas as pd
import os
import traceback
import psycopg2

from datetime import datetime
from db import get_db, db_query, get_conn
from psycopg2.extras import RealDictCursor
from sqlalchemy import create_engine, text
from lunar_rules import get_special_day_info, get_next_day_remove_info
from schedule.builders.time_utils import (
    parse_min,
    min_to_ampm,
    fix_time_format,
    normalize_time_text,
    choose_split_time,
)

from schedule.builders.flatten_builder import (
    flatten_arranged_for_db
)

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

DATA_DIR = os.path.join(BASE_DIR, "data")

FIXED_FILE = os.path.join(
    DATA_DIR,
    "fixed_schedule.xlsx"
)
PREBOOK_FILE = os.path.join(BASE_DIR, "prebook_schedule.xlsx")
OUTPUT_FILE = os.path.join(BASE_DIR, "schedule_output.txt")
DATABASE_URL = os.environ.get("DATABASE_URL")
engine = create_engine(DATABASE_URL, pool_pre_ping=True) if DATABASE_URL else None
VOLUNTEER_SIGNUP_URL = "https://gyt-checkin.onrender.com/volunteer"


DUTY_TARGETS = {
    "normal": {
        "卫生": 2,
        "绿观音堂": 1,
        "绿活动中心": 1,
        "橙观音堂": 1,
        "橙活动中心": 1,
        "黄观音堂": 1,
        "黄活动中心": 1,
    },

    "lunar_1_15": {
        "卫生": 3,
        "绿观音堂": 1,
        "绿活动中心": 1,
        "橙观音堂": 2,
        "橙活动中心": 2,
        "黄观音堂": 2,
        "黄活动中心": 2,
    },

    "buddhist_festival": {
        "卫生": 3,
        "绿观音堂": 2,
        "绿活动中心": 2,
        "橙观音堂": 2,
        "橙活动中心": 2,
        "黄观音堂": 2,
        "黄活动中心": 2,
    },
}


def get_duty_targets(template_type="normal"):
    return DUTY_TARGETS.get(
        template_type,
        DUTY_TARGETS["normal"]
    )

def get_shortage_targets(template_type):
    return DUTY_TARGETS.get(
        template_type,
        DUTY_TARGETS["normal"]
    )


def build_master_table_section(arranged):
    setup_master = format_people_inline(
        arranged.get("设师父供台", [])
    )

    remove_master = format_people_inline(
        arranged.get("收师父供台", [])
    )

    section = ""

    if setup_master:
        section += f"""
6:00am~8:00am 或 
6:00am~完成供台工作
设师父供台:
{setup_master}
"""

    if remove_master:
        section += f"""
12:00pm~2:00pm
收师父供台:
{remove_master}
12pm的香结束之后，请下供桌。
"""

    return section


def patch_schedule_for_date(date_str, only_signup_id=None):

    try:

        date_obj, arranged, assigned_rows, special_info, remove_info = \
            calculate_schedule_for_date(
                date_str,
                patch_mode=True,
            )

        print("========== PATCH DEBUG ==========")
        print("日期:", date_str)
        print("理论排班数量:", len(assigned_rows))

        for r in assigned_rows:
            print(
                r["signup_id"],
                r["role"],
                r["assigned_place"],
                r.get("shift_label"),
                r.get("start_time"),
                r.get("end_time")
            )
        print("=================================")

        inserted = sync_patch_assignments(
            date_str,
            assigned_rows,
            only_signup_id=only_signup_id
        )

        print(f"✅ Patch 完成，共新增 {inserted} 笔 assignment")

        return inserted

    except Exception as e:

        print("❌ patch_schedule_for_date 失败：", e)
        traceback.print_exc()

        return 0
    
def sync_patch_assignments(date_str, assigned_rows, only_signup_id=None):

    rows = db_query("""
        select
            signup_id,
            shift_label,
            assigned_place,
            start_time,
            end_time
        from volunteer_schedule_assignments
        where assignment_date=%s
    """, (date_str,), fetchall=True)

    existing_keys = set()

    for r in rows:
        if r["signup_id"] is None:
            continue

        existing_keys.add((
            int(r["signup_id"]),
            r.get("shift_label"),
            r.get("assigned_place"),
            str(r.get("start_time") or ""),
            str(r.get("end_time") or ""),
        ))

    print("已有 assignment key：", existing_keys)

    new_rows = []

    for r in assigned_rows:

        if r.get("role") != "值班":
            continue

        signup_id = r.get("signup_id")

        if signup_id is None:
            continue

        signup_id = int(signup_id)

        if only_signup_id is not None and signup_id != int(only_signup_id):
            continue

        key = (
            int(signup_id),
            r.get("shift_label"),
            r.get("assigned_place"),
            str(r.get("start_time") or ""),
            str(r.get("end_time") or ""),
        )

        if key in existing_keys:
            continue

        new_rows.append(r)
        existing_keys.add(key)

    print("需要新增：", len(new_rows))

    for r in new_rows:
        print(
            "新增 assignment：",
            r["signup_id"],
            r["name"],
            r["role"],
            r["assigned_place"],
            r.get("start_time"),
            r.get("end_time"),
        )

        db_query("""
            insert into volunteer_schedule_assignments
            (
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
            )
            values
            (
                %s, %s, %s, %s,
                %s, %s, %s,
                %s, %s,
                'assigned',
                '发布后自动补排'
            )
        """, (
            int(r["signup_id"]),
            r.get("volunteer_id"),
            r.get("name"),
            date_str,
            r.get("role"),
            r.get("shift_label"),
            r.get("assigned_place"),
            r.get("start_time"),
            r.get("end_time"),
        ))

        db_query("""
            update volunteer_schedule_signups
            set status='assigned'
            where id=%s
        """, (int(r["signup_id"]),))

    return len(new_rows)

def auto_assign_new_duty(result):
    """
    发布后自动补排：
    - 4小时或以上：拆成两段，前半段一个地方，后半段换地方
    - 少过4小时：不拆，整段安排去人数少的地方
    """

    duty_keys = [
        "绿观音堂", "绿活动中心",
        "橙观音堂", "橙活动中心",
        "黄观音堂", "黄活动中心",
    ]

    def get_start(item):
        return item.get("start_time") or item.get("开始时间")

    def get_end(item):
        return item.get("end_time") or item.get("结束时间")

    

    

    def shift_of(start_min):
        if start_min < 10 * 60:
            return "绿"
        if start_min < 14 * 60:
            return "橙"
        return "黄"

    def make_key(shift, place):
        return f"{shift}{place}"

    def make_item(item, start_txt, end_txt):
        new_item = dict(item)
        new_item["start_time"] = start_txt
        new_item["end_time"] = end_txt
        new_item["开始时间"] = start_txt
        new_item["结束时间"] = end_txt
        new_item["role"] = "值班"
        new_item["岗位"] = "值班"
        return new_item

    def place_count(shift, place):
        return len(result.get(make_key(shift, place), []))

    def choose_less_place(shift):
        gyt = place_count(shift, "观音堂")
        act = place_count(shift, "活动中心")

        if gyt <= act:
            return "观音堂"
        return "活动中心"

    def opposite_place(place):
        return "活动中心" if place == "观音堂" else "观音堂"

    duty_pool = []

    for key in duty_keys:
        for item in result.get(key, []):
            s = parse_min(get_start(item))
            e = parse_min(get_end(item))

            if s is not None and e is not None and e > s:
                duty_pool.append(item)

        result[key] = []

    duty_pool.sort(
        key=lambda item: (
            parse_min(get_start(item)) or 9999,
            parse_min(get_end(item)) or 9999,
            item.get("name") or item.get("姓名") or ""
        )
    )

    for item in duty_pool:
        s = parse_min(get_start(item))
        e = parse_min(get_end(item))

        if s is None or e is None or e <= s:
            continue

        duration = e - s

        if duration >= 4 * 60:
            mid = choose_split_time(s, e)

            first_shift = shift_of(s)
            first_place = choose_less_place(first_shift)

            second_shift = shift_of(mid)
            second_place = opposite_place(first_place)

            result[make_key(first_shift, first_place)].append(
                make_item(item, min_to_ampm(s), min_to_ampm(mid))
            )

            result[make_key(second_shift, second_place)].append(
                make_item(item, min_to_ampm(mid), min_to_ampm(e))
            )

        else:
            shift = shift_of(s)
            place = choose_less_place(shift)

            result[make_key(shift, place)].append(
                make_item(item, min_to_ampm(s), min_to_ampm(e))
            )

    return result


def load_prebook_input(target_date_str):
    if not os.path.exists(PREBOOK_FILE):
        return pd.DataFrame(columns=["姓名", "岗位", "开始时间", "结束时间", "优先岗位", "备注"])

    df = pd.read_excel(PREBOOK_FILE, sheet_name="预报名")
    df.columns = df.columns.astype(str).str.strip()

    need_cols = ["姓名", "岗位", "日期", "开始时间", "结束时间"]
    missing = [c for c in need_cols if c not in df.columns]
    if missing:
        raise ValueError(f"prebook_schedule.xlsx 缺少栏位：{missing}")

    df["日期"] = pd.to_datetime(df["日期"], errors="coerce").dt.strftime("%Y-%m-%d")
    df = df[df["日期"] == target_date_str].copy()

    if df.empty:
        return pd.DataFrame(columns=["姓名", "岗位", "开始时间", "结束时间", "优先岗位", "备注"])

    df["姓名"] = df["姓名"].astype(str).str.strip()
    df["岗位"] = df["岗位"].astype(str).str.strip().map(normalize_job_name)
    df["开始时间"] = df["开始时间"].astype(str).str.strip()
    df["结束时间"] = df["结束时间"].astype(str).str.strip()
    df["优先岗位"] = ""
    if "备注" not in df.columns:
        df["备注"] = ""
    else:
        df["备注"] = df["备注"].astype(str).str.strip()

    return df[["姓名", "岗位", "开始时间", "结束时间", "优先岗位", "备注"]].copy()

def get_latest_date_from_prebook():
    if not os.path.exists(PREBOOK_FILE):
        return None

    df = pd.read_excel(PREBOOK_FILE, sheet_name="预报名")
    df.columns = df.columns.astype(str).str.strip()

    if "日期" not in df.columns:
        return None

    df["日期"] = pd.to_datetime(df["日期"], errors="coerce")
    df = df.dropna(subset=["日期"])

    if df.empty:
        return None

    latest_date = df["日期"].max()
    return latest_date.to_pydatetime()

def merge_signup_and_prebook(signup_df, prebook_df):
    if signup_df is None or signup_df.empty:
        return prebook_df.copy()

    if prebook_df is None or prebook_df.empty:
        return signup_df.copy()

    merged = pd.concat([signup_df, prebook_df], ignore_index=True)

    merged = merged.drop_duplicates(
        subset=["姓名", "岗位", "开始时间", "结束时间"],
        keep="first"
    ).reset_index(drop=True)

    return merged

# ===== 星期映射 =====
WEEKDAY_MAP = {
    0: "星期一",
    1: "星期二",
    2: "星期三",
    3: "星期四",
    4: "星期五",
    5: "星期六",
    6: "星期日",
}

DEFAULT_TIME = {
    "佛台": ("08:00", "10:00"),
    "整理佛台": ("08:00", "10:00"),
    "佛堂卫生": ("08:00", "10:00"),
    "二楼卫生": ("08:00", "10:00"),
    "楼梯卫生": ("08:00", "10:00"),
    "设师父供台": ("06:00", "08:00"),
    "师父供台": ("06:00", "08:00"),
    "观音堂": ("10:00", "14:00"),
    "活动中心": ("10:00", "14:00"),
    "绿观音堂": ("08:00", "10:00"),
    "绿活动中心": ("08:00", "10:00"),
}

SPECIAL_DEFAULT_TIME = {
    "佛台": ("06:00", "08:00"),
    "整理佛台": ("06:00", "08:00"),
    "佛堂卫生": ("06:00", "08:00"),
    "二楼卫生": ("06:00", "08:00"),
    "楼梯卫生": ("06:00", "08:00"),
    "设师父供台": ("06:00", "08:00"),
    "师父供台": ("06:00", "08:00"),
    "观音堂": ("10:00", "14:00"),
    "活动中心": ("10:00", "14:00"),
    "绿观音堂": ("08:00", "10:00"),
    "绿活动中心": ("08:00", "10:00"),
}


def save_assignment_history(date_obj, result):
    if not engine:
        print("⚠️ 没有 DATABASE_URL，不能保存 assignment_history")
        return

    records = []

    def add_records(names, job):
        for n in names:
            if n:
                records.append({
                    "date": date_obj.strftime("%Y-%m-%d"),
                    "name": str(n).strip(),
                    "role": job
                })

    for key in ["橙观音堂", "橙活动中心", "黄观音堂", "黄活动中心"]:
        for item in result.get(key, []):
            name = item[0]
            job = "观音堂" if "观音堂" in key else "活动中心"
            add_records([name], job)

    add_records(result.get("佛堂卫生", []), "佛堂卫生")
    add_records(result.get("二楼卫生", []), "二楼卫生")
    add_records(result.get("楼梯卫生", []), "楼梯卫生")
    add_records(result.get("整理佛台", []), "整理佛台")
    add_records(result.get("设师父供台", []), "设师父供台")

    if not records:
        return

    date_str = date_obj.strftime("%Y-%m-%d")

    with engine.begin() as conn:
        # 重新生成同一天时，先删旧记录，避免重复
        conn.execute(
            text("delete from assignment_history where date = :date"),
            {"date": date_str}
        )

        for r in records:
            conn.execute(
                text("""
                    insert into assignment_history (date, name, role)
                    values (:date, :name, :role)
                """),
                r
            )


def load_supabase_signups(target_date_str):
    database_url = os.environ.get("DATABASE_URL")
    if not database_url:
        return pd.DataFrame(columns=["姓名", "岗位", "开始时间", "结束时间", "优先岗位", "备注"])

    with psycopg2.connect(database_url) as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                select id, volunteer_id, name, role, start_time, end_time
                from volunteer_schedule_signups
                where signup_date = %s
                and coalesce(status, 'pending') <> 'cancelled'
                order by created_at
            """, (target_date_str,))
            rows = cur.fetchall()

    if not rows:
        return pd.DataFrame(columns=["姓名", "岗位", "开始时间", "结束时间", "优先岗位", "备注"])

    data = []
    for r in rows:
        data.append({
            "姓名": r["name"],
            "岗位": r["role"],
            "开始时间": r["start_time"],
            "结束时间": r["end_time"],
            "优先岗位": "",
            "signup_id": r["id"],
            "volunteer_id": r["volunteer_id"],
            "备注": "义工网页报名",
        })

    return pd.DataFrame(data)


def update_assigned_places(target_date, assigned_rows):
    date_str = str(target_date)

    with get_conn() as conn:
        with conn.cursor() as cur:

            cur.execute("""
                delete from volunteer_schedule_assignments
                where assignment_date = %s
            """, (date_str,))

            for row in assigned_rows:
                signup_id = row.get("signup_id") or row.get("id")
                volunteer_id = row.get("volunteer_id")
                name = row.get("name") or row.get("姓名")
                role = row.get("role") or row.get("岗位")
                shift_label = row.get("shift_label")
                assigned_place = row.get("assigned_place")
                start_time = row.get("start_time")
                end_time = row.get("end_time")
                remarks = row.get("remarks") or row.get("备注")

                if not name or not role or not assigned_place:
                    continue

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
                    remarks
                ))

                if signup_id:
                    cur.execute("""
                        update volunteer_schedule_signups
                        set status = 'assigned'
                        where id = %s
                        and coalesce(status, 'pending') <> 'cancelled'
                    """, (signup_id,))
                else:
                    cur.execute("""
                        update volunteer_schedule_signups
                        set status = 'assigned'
                        where signup_date = %s
                        and name = %s
                        and role = %s
                        and coalesce(status, 'pending') <> 'cancelled'
                    """, (
                        date_str,
                        name,
                        role
                    ))

        conn.commit()

def load_last_assignment():
    if not engine:
        return {}

    try:
        with engine.connect() as conn:
            rows = conn.execute(text("""
                select distinct on (name)
                    name, role
                from assignment_history
                order by name, date desc, id desc
            """)).mappings().all()

        return {r["name"]: r["role"] for r in rows}

    except Exception as e:
        print("load_last_assignment error:", e)
        return {}
    
def load_yesterday_hygiene_assignment(date_obj):
    if not engine:
        return {}

    yesterday = (pd.to_datetime(date_obj) - pd.Timedelta(days=1)).strftime("%Y-%m-%d")

    try:
        with engine.connect() as conn:
            rows = conn.execute(text("""
                select distinct on (name)
                    name, role
                from assignment_history
                where date = :date
                  and role in ('佛堂卫生', '二楼卫生', '楼梯卫生')
                order by name, id desc
            """), {"date": yesterday}).mappings().all()

        return {r["name"]: r["role"] for r in rows}

    except Exception as e:
        print("load_yesterday_hygiene_assignment error:", e)
        return {}

def assign_hygiene_no_same_as_yesterday(hygiene_names, date_obj):
    jobs = ["佛堂卫生", "二楼卫生", "楼梯卫生"]
    no_fotang_names = {"葉荔銖", "葉茘銖"}

    yesterday_map = load_yesterday_hygiene_assignment(date_obj)

    result = {
        "佛堂卫生": [],
        "二楼卫生": [],
        "楼梯卫生": []
    }

    used = set()

    for job in jobs:
        candidates = []

        for name in hygiene_names:
            name = str(name).strip()

            if not name or name in used:
                continue

            if job == "佛堂卫生" and name in no_fotang_names:
                continue

            # 核心规则：不要做昨天同一个岗位
            if yesterday_map.get(name) != job:
                candidates.append(name)

        # 如果没人可选，才放宽允许重复昨天岗位
        if not candidates:
            for name in hygiene_names:
                name = str(name).strip()

                if not name or name in used:
                    continue

                if job == "佛堂卫生" and name in no_fotang_names:
                    continue

                candidates.append(name)

        if candidates:
            chosen = candidates[0]
            result[job].append(chosen)
            used.add(chosen)

    return result

def parse_time_str(t):
    if pd.isna(t):
        return None
    s = str(t).strip()
    if not s:
        return None
    return s


def get_weekday_name(date_obj):
    return WEEKDAY_MAP[date_obj.weekday()]


def load_fixed_schedule():
    xls = pd.ExcelFile(FIXED_FILE)
    sheets = xls.sheet_names

    print("检测到 fixed_schedule.xlsx 工作表：", sheets)

    if "佛台固定" in sheets:
        buddhist_df = pd.read_excel(FIXED_FILE, sheet_name="佛台固定")
        buddhist_df.columns = buddhist_df.columns.astype(str).str.strip()
        print("佛台固定栏位：", list(buddhist_df.columns))
    else:
        raise ValueError("❌ 必须有【佛台固定】sheet")

    if "卫生固定" in sheets:
        cleaning_df = pd.read_excel(FIXED_FILE, sheet_name="卫生固定")
        cleaning_df.columns = cleaning_df.columns.astype(str).str.strip()
        print("卫生固定栏位：", list(cleaning_df.columns))
    else:
        print("⚠️ 没有【卫生固定】，将全部由报名决定")
        cleaning_df = pd.DataFrame(columns=["星期", "佛堂卫生", "二楼卫生", "楼梯卫生1", "楼梯卫生2"])

    return buddhist_df, cleaning_df

def fix_time_range(text):
    text = fix_time_format(text)

    # 1030-1230pm → 10:30am~12:30pm
    text = re.sub(
        r"(\d{3,4})(am|pm)?\s*[-~]\s*(\d{3,4})(am|pm)",
        lambda m: f"{fix_time_format(m.group(1)+ (m.group(2) or 'am'))}~{fix_time_format(m.group(3)+m.group(4))}",
        text,
        flags=re.IGNORECASE
    )

    return text


def fix_time_range_v2(text):
    """
    修复：
    10 - 2 pm → 10am~2pm
    10-2pm → 10am~2pm
    10 - 2pm → 10am~2pm
    """

    text = str(text)

    def repl(m):
        h1 = int(m.group(1))
        h2 = int(m.group(2))
        ap2 = m.group(3).lower()

        # 推断第一个时间
        ap1 = "am" if h1 < h2 else ap2

        return f"{h1}{ap1}~{h2}{ap2}"

    text = re.sub(
        r"\b(\d{1,2})\s*[-~]\s*(\d{1,2})\s*(am|pm)\b",
        repl,
        text,
        flags=re.IGNORECASE
    )

    return text



def normalize_time_range(text):
    import re

    if not text:
        return None, None

    s = str(text).lower().strip()
    s = s.replace("～", "~").replace("-", "~")
    s = s.replace(".", ":")
    s = re.sub(r"\s+", " ", s)

    # 12:00pm~3:00pm
    m = re.search(r"(\d{1,2}:\d{2}(?:am|pm))\s*~\s*(\d{1,2}:\d{2}(?:am|pm))", s)
    if m:
        return m.group(1), m.group(2)

    # 4~6pm
    m = re.search(r"(\d{1,2})\s*~\s*(\d{1,2})(am|pm)", s)
    if m:
        start_hour = int(m.group(1))
        end_hour = int(m.group(2))
        suffix = m.group(3)

        # 特别规则：
        # 10-12pm / 9-12pm / 11-12pm 这种，前一个通常应视为 am
        if suffix == "pm" and start_hour < end_hour and end_hour == 12:
            return f"{start_hour}:00am", f"{end_hour}:00pm"

        return f"{start_hour}:00{suffix}", f"{end_hour}:00{suffix}"

    # 12pm 5pm
    m = re.findall(r"(\d{1,2})(am|pm)", s)
    if len(m) >= 2:
        return f"{m[0][0]}:00{m[0][1]}", f"{m[1][0]}:00{m[1][1]}"

    return None, None

def get_fixed_people(date_obj, buddhist_df, cleaning_df):
    weekday = get_weekday_name(date_obj)

    result = {
        "整理佛台": [],
        "佛堂卫生": [],
        "二楼卫生": [],
        "楼梯卫生": [],
    }

    if not buddhist_df.empty:
        buddhist_df = buddhist_df.copy()
        buddhist_df.columns = buddhist_df.columns.astype(str).str.strip()

        if "星期" not in buddhist_df.columns:
            raise ValueError(f"❌【佛台固定】缺少栏位：星期。当前栏位：{list(buddhist_df.columns)}")

        buddhist_row = buddhist_df[buddhist_df["星期"].astype(str).str.strip() == weekday]

        print("DEBUG weekday =", weekday)
        print("DEBUG buddhist_row =", buddhist_row)

        if not buddhist_row.empty:
            row = buddhist_row.iloc[0]
            name_cols = [c for c in buddhist_df.columns if str(c).startswith("姓名")]

            for col in name_cols:
                if pd.notna(row[col]) and str(row[col]).strip():
                    result["整理佛台"].append(str(row[col]).strip())

    if not cleaning_df.empty:
        cleaning_df = cleaning_df.copy()
        cleaning_df.columns = cleaning_df.columns.astype(str).str.strip()

        if "星期" in cleaning_df.columns:
            cleaning_row = cleaning_df[cleaning_df["星期"].astype(str).str.strip() == weekday]

            if not cleaning_row.empty:
                row = cleaning_row.iloc[0]

                if "佛堂卫生" in cleaning_df.columns and pd.notna(row.get("佛堂卫生")) and str(row.get("佛堂卫生")).strip():
                    result["佛堂卫生"].append(str(row.get("佛堂卫生")).strip())

                if "二楼卫生" in cleaning_df.columns and pd.notna(row.get("二楼卫生")) and str(row.get("二楼卫生")).strip():
                    result["二楼卫生"].append(str(row.get("二楼卫生")).strip())

                for col in ["楼梯卫生1", "楼梯卫生2"]:
                    if col in cleaning_df.columns and pd.notna(row.get(col)) and str(row.get(col)).strip():
                        result["楼梯卫生"].append(str(row.get(col)).strip())

    return result

def load_buddha_override(date_obj):
    file = "buddha_override.xlsx"

    if not os.path.exists(file):
        return None

    df = pd.read_excel(file)

    if "日期" not in df.columns:
        return None

    df["日期"] = pd.to_datetime(df["日期"], errors="coerce")

    row = df[df["日期"] == date_obj]

    if row.empty:
        return None

    names = []

    for col in ["姓名1", "姓名2", "姓名3"]:
        if col in row.columns:
            val = str(row.iloc[0][col]).strip()
            if val and val != "nan":
                names.append(val)

    return names if names else None


def normalize_job_name(job):
    job = str(job).strip()

    mapping = {
        "佛台": "整理佛台",
        "整理佛台": "整理佛台",
        "卫生": "卫生",
        "佛堂卫生": "佛堂卫生",
        "二楼卫生": "二楼卫生",
        "楼梯卫生": "楼梯卫生",
        "设师父供台": "设师父供台",
        "师父供台": "设师父供台",
        "观音堂": "观音堂",
        "活动中心": "活动中心",
        "绿观音堂": "绿观音堂",
        "绿活动中心": "绿活动中心",
    }
    return mapping.get(job, job)

def parse_signup_line_multi(text):
    text = str(text).strip()
    if not text:
        return []

    text = text.replace("　", " ")
    text = re.sub(r"\s+", " ", text)

    # 统一常见字
    text = text.replace("衛生", "卫生")

    # ===== 全日规则 =====
    is_full_day = "全日" in text

    # 全日先去掉，避免被吃进名字
    clean_text = text.replace("全日", "").strip()

    # 抓时间
    start, end = normalize_time_range(clean_text)

    # 如果写了全日但没写时间，默认 10am~6pm
    if is_full_day and (not start or not end):
        start, end = "10:00am", "6:00pm"

    role_text = clean_text
    if start:
        role_text = role_text.replace(start, "")
    if end:
        role_text = role_text.replace(end, "")
    role_text = role_text.replace("~", " ").replace("-", " ")
    role_text = re.sub(r"\s+", " ", role_text).strip()

    role_text = role_text.replace("&", "/").replace("＆", "/").replace("和", "/")

    job_keywords = ["活动中心", "观音堂", "卫生", "佛台", "值班"]

    found = []
    for kw in job_keywords:
        start_pos = 0
        while True:
            pos = role_text.find(kw, start_pos)
            if pos == -1:
                break
            found.append((pos, kw))
            start_pos = pos + len(kw)

    found.sort(key=lambda x: x[0])

    if not found:
        m = re.match(r"^([A-Za-z\u4e00-\u9fff·•.]+)", role_text)
        if not m:
            return []
        name = m.group(1).rstrip(".").strip()
        return [{
            "姓名": name,
            "岗位": "值班",
            "开始时间": start,
            "结束时间": end,
            "优先岗位": "",
            "备注": "",
        }]

    first_pos = found[0][0]
    name = role_text[:first_pos].strip().rstrip(".")

    if not name:
        return []

    records = []

    for _, kw in found:
        if kw == "卫生":
            job = "卫生"
            job_start, job_end = None, None
        elif kw == "观音堂":
            job = "观音堂"
            job_start, job_end = start, end
        elif kw == "活动中心":
            job = "活动中心"
            job_start, job_end = start, end
        elif kw == "佛台":
            job = "整理佛台"
            job_start, job_end = start, end
        elif kw == "值班":
            job = "值班"
            job_start, job_end = start, end
        else:
            continue

        records.append({
            "姓名": name,
            "岗位": job,
            "开始时间": job_start,
            "结束时间": job_end,
            "优先岗位": "",
            "备注": "",
        })

    dedup = []
    seen = set()
    for r in records:
        key = (r["姓名"], r["岗位"], r["开始时间"], r["结束时间"])
        if key not in seen:
            seen.add(key)
            dedup.append(r)

    return dedup

def parse_signup_line(text):
    text = str(text).strip()
    if not text:
        return None

    text = text.replace("　", " ")
    text = re.sub(r"\s+", " ", text)

    # 先定义岗位关键词
    job_keywords = ["活动中心", "观音堂", "卫生", "佛台", "值班"]

    found_job = None
    found_pos = None

    for kw in job_keywords:
        pos = text.find(kw)
        if pos >= 0:
            if found_pos is None or pos < found_pos:
                found_job = kw
                found_pos = pos

    # 时间先抓出来
    start, end = normalize_time_range(text)

    if found_job is not None:
        before_job = text[:found_pos].strip()
        after_job = text[found_pos + len(found_job):].strip()

        # 名字通常在岗位前面；如果岗位前面还有时间，要去掉
        name = before_job

        # 去掉名字后面可能混进去的时间
        name = re.sub(r"\d{1,2}[:.]?\d{0,2}\s*(?:am|pm)?\s*[-~]?\s*\d{0,2}[:.]?\d{0,2}\s*(?:am|pm)?", "", name, flags=re.I).strip()
        name = name.rstrip(".").strip()

        rest = f"{found_job} {after_job}".strip()
    else:
        # 没岗位关键词时，抓开头名字
        m = re.match(r"^([A-Za-z\u4e00-\u9fff·•.]+)", text)
        if not m:
            return None

        name = m.group(1).rstrip(".").strip()
        rest = text[len(m.group(0)):].strip()
        found_job = None

    # 最终岗位映射
    if found_job == "卫生":
        job = "卫生"
    elif found_job == "观音堂":
        job = "观音堂"
    elif found_job == "活动中心":
        job = "活动中心"
    elif found_job == "佛台":
        job = "整理佛台"
    elif found_job == "值班":
        job = "值班"
    else:
        # 没写岗位时，默认值班
        job = "值班"

    return {
        "姓名": name,
        "岗位": job,
        "开始时间": start,
        "结束时间": end,
        "优先岗位": "",
        "备注": "",
    }

def auto_assign_cleaning_roles(result, signup_df):
    cleaning_people = []

    for _, r in signup_df.iterrows():
        if str(r.get("岗位", "")).strip() == "卫生":
            name = str(r.get("姓名", "")).strip()
            if name and name not in cleaning_people:
                cleaning_people.append(name)

    result["佛堂卫生"] = []
    result["二楼卫生"] = []
    result["楼梯卫生"] = []

    if not cleaning_people:
        return result

    # 叶荔铢不排佛堂卫生
    special_name = "葉荔銖"

    # ===== 1人：三处同一人 =====
    if len(cleaning_people) == 1:
        name = cleaning_people[0]

        if name == special_name:
            result["二楼卫生"] = [name]
            result["楼梯卫生"] = [name]
        else:
            result["佛堂卫生"] = [name]
            result["二楼卫生"] = [name]
            result["楼梯卫生"] = [name]

        return result

    # ===== 2人：佛堂1人，二楼1人，楼梯两人共同 =====
    if len(cleaning_people) == 2:
        p1, p2 = cleaning_people

        if special_name in cleaning_people:
            other = p2 if p1 == special_name else p1

            result["佛堂卫生"] = [other]
            result["二楼卫生"] = [special_name]
            result["楼梯卫生"] = [other, special_name]

        else:
            result["佛堂卫生"] = [p1]
            result["二楼卫生"] = [p2]
            result["楼梯卫生"] = [p1, p2]

        return result

    # ===== 3人以上：最多取前3人分配 =====
    people = cleaning_people[:3]

    if special_name in people:
        others = [p for p in people if p != special_name]

        result["二楼卫生"] = [special_name]

        if others:
            result["佛堂卫生"] = [others[0]]

        if len(others) >= 2:
            result["楼梯卫生"] = [others[1]]
        elif others:
            result["楼梯卫生"] = [others[0]]
        else:
            result["楼梯卫生"] = [special_name]

    else:
        result["佛堂卫生"] = [people[0]]
        result["二楼卫生"] = [people[1]]
        result["楼梯卫生"] = [people[2]]

    return result

def normalize_signup(df, is_special_day=False):
    rows = []

    for _, r in df.iterrows():
        name = str(r.get("姓名", "")).strip()
        wishes = str(r.get("岗位意向", "")).strip()
        start = parse_time_str(r.get("开始时间", ""))
        end = parse_time_str(r.get("结束时间", ""))
        priority = str(r.get("优先岗位", "")).strip()
        note = str(r.get("备注", "")).strip()

        if not name or not wishes:
            continue

        jobs = [x.strip() for x in wishes.split(",") if x.strip()]
        for raw_job in jobs:
            job = normalize_job_name(raw_job)

            job_start = start
            job_end = end

            if not job_start or not job_end:
                if is_special_day and job in SPECIAL_DEFAULT_TIME:
                    job_start, job_end = SPECIAL_DEFAULT_TIME[job]
                elif job in DEFAULT_TIME:
                    job_start, job_end = DEFAULT_TIME[job]

            rows.append({
                "姓名": name,
                "岗位": job,
                "开始时间": job_start,
                "结束时间": job_end,
                "优先岗位": priority,
                "备注": note,
            })

    return pd.DataFrame(rows)


def classify_shift_slot(start_time):
    s = parse_time_str(start_time)
    if not s:
        return "橙"

    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            dt = datetime.strptime(s, fmt)
            hm = dt.hour * 60 + dt.minute
            if 8 * 60 <= hm < 10 * 60:
                return "绿"
            elif 10 * 60 <= hm < 14 * 60:
                return "橙"
            else:
                return "黄"
        except ValueError:
            pass

    s2 = s.lower().replace(" ", "")
    if "am" in s2 or "pm" in s2:
        for fmt in ("%I:%M%p", "%I.%M%p"):
            try:
                dt = datetime.strptime(s2, fmt)
                hm = dt.hour * 60 + dt.minute
                if 8 * 60 <= hm < 10 * 60:
                    return "绿"
                elif 10 * 60 <= hm < 14 * 60:
                    return "橙"
                else:
                    return "黄"
            except ValueError:
                pass

    return "橙"


def get_person_name(p):
    if isinstance(p, dict):
        return p.get("name") or p.get("姓名")
    if isinstance(p, (list, tuple)):
        return p[0] if p else None
    return p


def assign_jobs(fixed_people, signup_df, special_info, patch_mode=False,):
    result = {
        "整理佛台": fixed_people["整理佛台"][:],
        "佛堂卫生": fixed_people["佛堂卫生"][:],
        "二楼卫生": fixed_people["二楼卫生"][:],
        "楼梯卫生": fixed_people["楼梯卫生"][:],
        "设师父供台": [],
        "绿观音堂": [],
        "绿活动中心": [],
        "橙观音堂": [],
        "橙活动中心": [],
        "黄观音堂": [],
        "黄活动中心": [],
    }

    cleaning_signups = []

    for _, r in signup_df.iterrows():
        name = r["姓名"]
        job = r["岗位"]
        start = r["开始时间"]
        end = r["结束时间"]

        signup_id = r.get("signup_id")
        volunteer_id = r.get("volunteer_id")

        person = {
            "name": name,
            "姓名": name,
            "start_time": start,
            "end_time": end,
            "开始时间": start,
            "结束时间": end,
            "signup_id": signup_id,
            "volunteer_id": volunteer_id,
        }

        if job == "整理佛台":
            person["role"] = "整理佛台"
            person["岗位"] = "整理佛台"
            result["整理佛台"].append(person)

        elif job in ["卫生", "佛堂卫生", "二楼卫生", "楼梯卫生"]:
            person["role"] = "卫生"
            person["岗位"] = "卫生"
            cleaning_signups.append(person)

        elif job in ["设师父供台", "供台"]:
            person["role"] = "供台"
            person["岗位"] = "供台"
            result["设师父供台"].append(person)

        elif job == "绿观音堂":
            person["role"] = "值班"
            person["岗位"] = "值班"
            result["绿观音堂"].append(person)

        elif job == "绿活动中心":
            person["role"] = "值班"
            person["岗位"] = "值班"
            result["绿活动中心"].append(person)

        elif job in ["观音堂", "活动中心", "值班"]:
            if not start or not end:
                start = "10:00am"
                end = "2:00pm"

            person["start_time"] = start
            person["end_time"] = end
            person["开始时间"] = start
            person["结束时间"] = end
            person["role"] = "值班"
            person["岗位"] = "值班"

            slot = classify_shift_slot(start)

            if job == "值班":
                key = f"{slot}观音堂"
            else:
                key = f"{slot}{job}"

            result[key].append(person)

    # ✅ 合并固定卫生 + 报名卫生
    all_cleaners = []
    all_cleaners.extend(result["佛堂卫生"])
    all_cleaners.extend(result["二楼卫生"])
    all_cleaners.extend(result["楼梯卫生"])
    all_cleaners.extend(cleaning_signups)

    # ✅ 卫生去重
    unique_cleaners = []
    seen = set()

    for p in all_cleaners:
        if isinstance(p, dict):
            name = p.get("name") or p.get("姓名")
            item = dict(p)
        else:
            name = str(p).strip()
            item = {
                "name": name,
                "姓名": name,
            }

        if not name:
            continue

        if name in seen:
            continue

        seen.add(name)

        item["role"] = "卫生"
        item["岗位"] = "卫生"
        unique_cleaners.append(item)

    # ✅ 重新分配卫生
    result["佛堂卫生"] = []
    result["二楼卫生"] = []
    result["楼梯卫生"] = []

    cleaning_count = len(unique_cleaners)

    if cleaning_count == 1:
        p = dict(unique_cleaners[0])
        result["佛堂卫生"].append(dict(p))
        result["二楼卫生"].append(dict(p))
        result["楼梯卫生"].append(dict(p))

    elif cleaning_count == 2:
        p1 = dict(unique_cleaners[0])
        p2 = dict(unique_cleaners[1])

        result["佛堂卫生"].append(dict(p1))
        result["二楼卫生"].append(dict(p2))
        result["楼梯卫生"].append(dict(p1))
        result["楼梯卫生"].append(dict(p2))

    elif cleaning_count >= 3:
        cleaning_places = ["佛堂卫生", "二楼卫生", "楼梯卫生"]

        for i, p in enumerate(unique_cleaners):
            place = cleaning_places[i % len(cleaning_places)]
            result[place].append(dict(p))

    if patch_mode:
        result = auto_assign_new_duty(result)
    else:
        result = auto_assign_duty(result)

    return result

    
def auto_assign_duty(result):
    def parse_min(t):
        if not t:
            return None

        s = str(t).strip().lower().replace(" ", "")

        if re.match(r"^\d{1,2}(am|pm)$", s):
            s = s.replace("am", ":00am").replace("pm", ":00pm")

        for fmt in ("%H:%M", "%H:%M:%S", "%I:%M%p"):
            try:
                dt = datetime.strptime(s, fmt)
                return dt.hour * 60 + dt.minute
            except:
                pass

        return None

    def min_to_ampm(m):
        hour = m // 60
        minute = m % 60
        suffix = "am" if hour < 12 else "pm"
        hour12 = hour % 12 or 12
        return f"{hour12}:{minute:02d}{suffix}"

    def get_name(item):
        if isinstance(item, dict):
            return item.get("name") or item.get("姓名")
        if isinstance(item, (list, tuple)):
            return item[0] if item else None
        return str(item) if item else None

    def get_start(item):
        if isinstance(item, dict):
            return item.get("start_time") or item.get("开始时间")
        if isinstance(item, (list, tuple)):
            return item[1] if len(item) > 1 else None
        return None

    def get_end(item):
        if isinstance(item, dict):
            return item.get("end_time") or item.get("结束时间")
        if isinstance(item, (list, tuple)):
            return item[2] if len(item) > 2 else None
        return None

    def get_volunteer_id(item):
        if isinstance(item, dict):
            return str(
                item.get("volunteer_id")
                or item.get("编号")
                or ""
            ).strip()
        return ""

    def make_item(item, start_txt, end_txt):
        if isinstance(item, dict):
            new_item = dict(item)
        else:
            name = get_name(item)
            new_item = {
                "name": name,
                "姓名": name,
                "signup_id": None,
                "volunteer_id": None,
            }

        new_item["start_time"] = start_txt
        new_item["end_time"] = end_txt
        new_item["开始时间"] = start_txt
        new_item["结束时间"] = end_txt
        new_item["role"] = "值班"
        new_item["岗位"] = "值班"
        return new_item

    def shift_of(start_min):
        if start_min < 10 * 60:
            return "绿"
        if start_min < 14 * 60:
            return "橙"
        return "黄"

    def make_key(shift, place):
        return f"{shift}{place}"

    def place_load(key):
        total = 0
        for item in result.get(key, []):
            s = parse_min(get_start(item))
            e = parse_min(get_end(item))
            if s is not None and e is not None and e > s:
                total += e - s
        return total

    def choose_less_place(shift):
        gyt_key = make_key(shift, "观音堂")
        act_key = make_key(shift, "活动中心")

        if len(result.get(gyt_key, [])) == 0:
            return "观音堂"

        if len(result.get(act_key, [])) == 0:
            return "活动中心"

        if place_load(gyt_key) <= place_load(act_key):
            return "观音堂"

        return "活动中心"

    def opposite_place(place):
        return "活动中心" if place == "观音堂" else "观音堂"

    def choose_place_for_item(shift, item, preferred_place=None):
        volunteer_id = get_volunteer_id(item)

        if volunteer_id == "CHE-238":
            if shift in ["绿", "橙"]:
                return "活动中心"
            if shift == "黄":
                return "观音堂"

        if preferred_place:
            return preferred_place

        return choose_less_place(shift)

    duty_keys = [
        "绿观音堂", "绿活动中心",
        "橙观音堂", "橙活动中心",
        "黄观音堂", "黄活动中心",
    ]

    duty_pool = []

    for key in duty_keys:
        for item in result.get(key, []):
            s = parse_min(get_start(item))
            e = parse_min(get_end(item))

            if s is None or e is None or e <= s:
                continue

            duty_pool.append({
                "name": get_name(item),
                "start": s,
                "end": e,
                "duration": e - s,
                "raw": item,
            })

    for key in duty_keys:
        result[key] = []

    duty_pool.sort(
        key=lambda x: (
            x["start"],
            -x["duration"],
            x["name"] or ""
        )
    )

    for p in duty_pool:
        s = p["start"]
        e = p["end"]
        duration = p["duration"]
        raw = p["raw"]

        if duration >= 4 * 60:
            mid = s + duration // 2

            first_shift = shift_of(s)
            first_place = choose_place_for_item(first_shift, raw)

            second_shift = shift_of(mid)
            second_place = opposite_place(first_place)

            first_item = make_item(raw, min_to_ampm(s), min_to_ampm(mid))
            second_item = make_item(raw, min_to_ampm(mid), min_to_ampm(e))

            result[make_key(first_shift, first_place)].append(first_item)
            result[make_key(second_shift, second_place)].append(second_item)

        else:
            shift = shift_of(s)
            place = choose_place_for_item(shift, raw)

            item = make_item(raw, min_to_ampm(s), min_to_ampm(e))
            result[make_key(shift, place)].append(item)

    return result

def format_people_inline(names, sep="  "):
    if not names:
        return ""

    result = []

    for item in names:
        if isinstance(item, dict):
            name = item.get("name") or item.get("姓名")
        elif isinstance(item, (list, tuple)):
            name = item[0] if item else ""
        else:
            name = item

        if name:
            result.append(str(name))

    return sep.join(result)


def to_ampm_text(t):
    if pd.isna(t):
        return ""

    s = str(t).strip()
    if not s:
        return ""

    s_lower = s.lower().replace(" ", "")

    if "am" in s_lower or "pm" in s_lower:
        return s_lower

    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            dt = datetime.strptime(s, fmt)
            hour = dt.hour
            minute = dt.minute
            suffix = "am" if hour < 12 else "pm"
            hour12 = hour % 12
            if hour12 == 0:
                hour12 = 12
            return f"{hour12}:{minute:02d}{suffix}"
        except ValueError:
            pass

    return s


def format_shift_block(items):
    if not items:
        return ""

    lines = []

    for item in items:
        if isinstance(item, dict):
            name = item.get("name") or item.get("姓名") or ""
            start = item.get("start_time") or item.get("开始时间")
            end = item.get("end_time") or item.get("结束时间")

        elif isinstance(item, (list, tuple)):
            name = item[0] if len(item) > 0 else ""
            start = item[1] if len(item) > 1 else ""
            end = item[2] if len(item) > 2 else ""

        else:
            name = str(item)
            start = ""
            end = ""

        if not name:
            continue

        lines.append(str(name))

        if start and end:
            start_txt = to_ampm_text(start)
            end_txt = to_ampm_text(end)
            lines.append(f"{start_txt}~{end_txt}")

    return "\n".join(lines)


def get_display_name(p):
    if isinstance(p, dict):
        return p.get("name") or p.get("姓名") or ""
    if isinstance(p, (list, tuple)):
        return p[0] if p else ""
    return str(p or "")


def build_normal_message(date_obj, arranged, special_info, remove_info):
    weekday = get_weekday_name(date_obj)
    date_text = f"{date_obj.day}/{date_obj.month}/{date_obj.year}"

    setup_people = arranged.get("设师父供台", [])
    remove_people = arranged.get("收师父供台", [])

    setup_master = format_people_inline(setup_people)
    remove_master = format_people_inline(remove_people)

    master_section = ""

    # 自动判断需要设供台
    if special_info.get("setup_shifu"):

        if not setup_master:
            setup_master = "待安排"

        master_section += f"""
    6:00am~8:00am
    设师父供台:
    {setup_master}
    """

    # 自动判断需要收供台
    if remove_info.get("need_remove_today_after_12"):

        if not remove_master:
            remove_master = "待安排"

        master_section += f"""
    12:00pm~2:00pm
    收师父供台:
    {remove_master}
    12pm的香结束之后，请下供桌。
    """

    green_section = ""

    if arranged.get("绿观音堂") or arranged.get("绿活动中心"):
        green_section = f"""
🟢 绿班
🟢 观音堂:
{format_shift_block(arranged.get("绿观音堂", []))}
🟢 活动中心:
{format_shift_block(arranged.get("绿活动中心", []))}
"""

    msg = f"""师兄们大家好！

{date_text} ({weekday}) 

十点正请安

8:00am~10:00am  或      
8:00am~完成佛台工作 
整理佛台: 
{format_people_inline(arranged.get("整理佛台", []))}

{master_section}
8:00am~10:00am 或 
8:00am~完成卫生工作 
佛堂卫生: {format_people_inline(arranged.get("佛堂卫生", []))}
二楼卫生: {format_people_inline(arranged.get("二楼卫生", []))}
楼梯卫生: {format_people_inline(arranged.get("楼梯卫生", []), sep="/")}
每日卫生义工请注意：清理完卫生之后，请把卫生用具包括吸尘机放回原位（活动中心store里面的小房间）

{green_section}

10:00am~2:00pm 
🟠 观音堂: 
{format_shift_block(arranged.get("橙观音堂", []))}
🟠 活动中心: 
{format_shift_block(arranged.get("橙活动中心", []))}

2:00pm~6:00pm 
🟡 观音堂: 
{format_shift_block(arranged.get("黄观音堂", []))}
🟡 活动中心: 
{format_shift_block(arranged.get("黄活动中心", []))}

观音堂早晚香 由值班义工带领上香。

观音堂续香
佛友可以续香（黑色无烟香），但是要跟着请安词，必须燃烧完了一支香才续香。如有多位佛友要续香，请先让给先到达观音堂的佛友，轮流续香。

值班义工请注意
1）下雨天记得关上烧送小房子的窗口。
2）在离开观音堂之前，请确保把所有的窗口关上。
3）观音堂第一架的冷气坚决不能调。第二架和第三架可以轮流。

另外，请大家多留意义工群信息，以便大家能够团结一致的护持好观音堂。佛子齐心，普度众生。

义工报名请点击以下链接：
{VOLUNTEER_SIGNUP_URL}

非常感恩大家！
大家功德无量！
🙏🙏🙏
"""
    return msg


def build_lunar_1_15_message(date_obj, arranged, special_info, remove_info):
    weekday = get_weekday_name(date_obj)
    date_text = f"{date_obj.day}/{date_obj.month}/{date_obj.year}"
    lunar_text = special_info.get("lunar_text", "")

    msg = f"""师兄们，大家好！

{date_text} ({weekday}) 

{lunar_text}

八点正请安 

6:00am~8:00am 或 
6:00am~完成佛台工作 
整理佛台:  
{format_people_inline(arranged["整理佛台"])}

{build_master_table_section(arranged)}

6:00am~8:00am 或 
6:00am~完成卫生工作 
佛堂卫生: {format_people_inline(arranged["佛堂卫生"])}
二楼卫生: {format_people_inline(arranged["二楼卫生"])}
楼梯卫生: {format_people_inline(arranged["楼梯卫生"], sep="/")}
每日卫生义工请注意：清理完卫生之后，请把卫生用具包括吸尘机放回原位（活动中心store里面的小房间）

🟢 8:00am~10:00am
每个岗位1位师兄
观音堂: {format_people_inline(arranged["绿观音堂"])}
活动中心: {format_people_inline(arranged["绿活动中心"])}

10:00am~2:00pm 
🟠 观音堂: 
{format_shift_block(arranged["橙观音堂"])}
🟠 活动中心: 
{format_shift_block(arranged["橙活动中心"])}

2:00pm~6:00pm 
🟡 观音堂: 
{format_shift_block(arranged["黄观音堂"])}
🟡 活动中心: 
{format_shift_block(arranged["黄活动中心"])}

观音堂早晚香
由值班义工带领上香。

观音堂续香
佛友可以续香（黑色无烟香），但是要跟着请安词，必须燃烧完了一支香才续香。如有多位佛友要续香，请先让给先到达观音堂的佛友，轮流续香。

值班义工请注意
1）下雨天记得关上烧送小房子的窗口。
2）在离开观音堂之前，请确保把所有的窗口关上。
3）观音堂第一架的冷气坚决不能调。第二架和第三架可以轮流。

另外，请大家多留意义工群信息，以便大家能够团结一致的护持好观音堂。佛子齐心，普度众生。

义工报名请点击以下链接：
{VOLUNTEER_SIGNUP_URL}

非常感恩大家！
大家功德无量！
🙏🙏🙏
"""
    return msg


def build_buddhist_festival_message(date_obj, arranged, special_info, remove_info):
    weekday = get_weekday_name(date_obj)
    date_text = f"{date_obj.day}/{date_obj.month}/{date_obj.year}"
    lunar_text = special_info.get("lunar_text", "")
    festival_text = "\n".join(special_info.get("special_names", []))

    remove_notice = ""
    if remove_info["need_remove_today_after_12"]:
        remove_notice = "今日中午12点后请记得撤供台。\n\n"

    msg = f"""师兄们，大家好！

{date_text} ({weekday}) 

{lunar_text}

{festival_text}

八点正请安 

6:00am~8:00am 或 
6:00am~完成佛台工作
整理佛台: 
{format_people_inline(arranged["整理佛台"])}

{build_master_table_section(arranged)}

6:00am~8:00am 或 
6:00am~完成卫生工作 
佛堂卫生: {format_people_inline(arranged["佛堂卫生"])}
二楼卫生: {format_people_inline(arranged["二楼卫生"])}
楼梯卫生: {format_people_inline(arranged["楼梯卫生"], sep="/")}
每日卫生义工请注意：清理完卫生之后，请把卫生用具包括吸尘机放回原位（活动中心store里面的小房间）

8:00am~10:00am
每个岗位1位师兄
观音堂: {format_people_inline(arranged["绿观音堂"])}
活动中心: {format_people_inline(arranged["绿活动中心"])}

10:00am~2:00pm 
🔴 观音堂: 
{format_shift_block(arranged["橙观音堂"])}
🔴 活动中心: 
{format_shift_block(arranged["橙活动中心"])}

2:00pm~6:00pm 
🟡 观音堂: 
{format_shift_block(arranged["黄观音堂"])}
🟡 活动中心: 
{format_shift_block(arranged["黄活动中心"])}

观音堂早晚香
由值班义工带领上香。

观音堂续香
佛友可以续香（黑色无烟香），但是要跟着请安词，必须燃烧完了一支香才续香。如有多位佛友要续香，请先让给先到达观音堂的佛友，轮流续香。

值班义工请注意
1）下雨天记得关上烧送小房子的窗口。
2）在离开观音堂之前，请确保把所有的窗口关上。
3）观音堂第一架的冷气坚决不能调。第二架和第三架可以轮流。

另外，请大家多留意义工群信息，以便大家能够团结一致的护持好观音堂。佛子齐心，普度众生。

义工报名请点击以下链接：
{VOLUNTEER_SIGNUP_URL}

{remove_notice}非常感恩大家！
大家功德无量！
🙏🙏🙏
"""
    return msg


def main():
    try:
        date_obj = get_latest_date_from_prebook()
        if date_obj is None:
            print("❌ prebook_schedule.xlsx 没有有效日期")
            return

        target_date_str = date_obj.strftime("%Y-%m-%d")
        print(f"自动使用日期：{target_date_str}")

    except Exception as e:
        print(f"❌ 读取日期失败：{e}")
        return

    try:
        message = run_schedule_for_date(target_date_str)

        with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
            f.write(message)

        print("已生成排班文案：", OUTPUT_FILE)
        print("=" * 50)
        print(message)

    except Exception as e:
        print(f"❌ 生成排班失败：{e}")



def calculate_schedule_for_date(date_str, patch_mode=False,):
    from datetime import datetime

    date_obj = datetime.strptime(date_str, "%Y-%m-%d")

    buddhist_df, cleaning_df = load_fixed_schedule()
    fixed_people = get_fixed_people(date_obj, buddhist_df, cleaning_df)
    override_names = load_buddha_override(date_obj)

    print("DEBUG fixed_people =", fixed_people)

    if override_names:
        original = fixed_people.get("整理佛台", [])

        final = []
        used = set()

        for i in range(3):
            if i < len(override_names) and override_names[i]:
                name = override_names[i]
                if name not in used:
                    final.append(name)
                    used.add(name)
            elif i < len(original):
                name = original[i]
                if name not in used:
                    final.append(name)
                    used.add(name)

        fixed_people["整理佛台"] = final

    special_info = get_special_day_info(date_obj)
    remove_info = get_next_day_remove_info(date_obj)

    signup_df = load_prebook_input(date_str)
    web_df = load_supabase_signups(date_str)

    if web_df is not None and not web_df.empty:
        signup_df = pd.concat(
            [signup_df, web_df],
            ignore_index=True
        )

    arranged = assign_jobs(fixed_people, signup_df, special_info, patch_mode=patch_mode,)
    assigned_rows = flatten_arranged_for_db(arranged)

    return date_obj, arranged, assigned_rows, special_info, remove_info


        


def run_schedule_for_date(date_str):
    try:
        date_obj, arranged, assigned_rows, special_info, remove_info = calculate_schedule_for_date(date_str)

        print("========== DEBUG ==========")
        print("排班结果数量:", len(assigned_rows))
        print(assigned_rows[:5])
        print("===========================")

        try:
            update_assigned_places(date_str, assigned_rows)
        except Exception as e:
            print("⚠️ 写入 Supabase assignments 失败：", e)
            traceback.print_exc()

        if special_info["template_type"] == "normal":
            message = build_normal_message(date_obj, arranged, special_info, remove_info)

        elif special_info["template_type"] == "lunar_1_15":
            message = build_lunar_1_15_message(date_obj, arranged, special_info, remove_info)

        else:
            message = build_buddhist_festival_message(date_obj, arranged, special_info, remove_info)

        return message

    except Exception as e:
        traceback.print_exc()
        return f"❌ 排班失败：{e}"

if __name__ == "__main__":
    main()

