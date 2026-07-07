# attendance_service.py

import datetime

from db import db_query, get_conn
from psycopg2.extras import RealDictCursor
from schedule.builders.time_utils import malaysia_today, malaysia_now, time_to_minutes
from utils import (
    now_date_str,
    now_time_str,
    get_lang,
    parse_time_to_datetime,
)


ENABLE_SIGNIN_TIME_LIMIT = False      # 是否限制签到时间
SIGNIN_EARLY_MINUTES = 30             # 可提前几分钟签到
ENABLE_AUTO_SIGNOUT = False           # 自动签退
AUTO_SIGNOUT_TIME = "19:00"           # 系统几点执行自动签退
AUTO_SIGNOUT_DISPLAY = "18:30"        # 报表显示的签退时间
REQUIRE_ASSIGNMENT_FOR_SIGNIN = False   # 测试时 False，正式上线改 True
# ===== 签到系统设定 =====
TODAY_CODE_ENABLED = False


def format_date_value(v):
    if v is None:
        return ""

    try:
        return v.strftime("%Y-%m-%d")
    except Exception:
        return str(v)


def load_attendance_rows() -> list[dict]:
    rows = db_query("""
        select *
        from attendance
        order by id desc
        limit 200
    """, fetchall=True)

    result = []
    for r in rows:
        result.append({
            "日期": r.get("date"),
            "编号": r.get("volunteer_id"),
            "姓名": r.get("name"),
            "报名": r.get("signup"),
            "签到": r.get("signin"),
            "岗位": r.get("role"),
            "card_no": r.get("card_no", ""),
            "开始时间": r.get("start_time"),
            "结束时间": r.get("end_time") or "",
            "时数": r.get("hours") or "",
            "备注": r.get("remark") or "",
            "_row": r.get("id"),
        })

    return result

def get_today_records(limit: int | None = None) -> list[dict]:
    today = now_date_str()
    out = []
    for item in load_attendance_rows():
        if format_date_value(item.get("日期")) == today:
            out.append(item)
    if limit is not None:
        return out[-limit:]
    return out

def get_today_open_records() -> list[dict]:
    today = now_date_str()

    rows = db_query("""
        select *
        from attendance
        where date = %s
          and (end_time is null or end_time = '')
        order by id desc
    """, (today,), fetchall=True)

    result = []
    for r in rows:
        result.append({
            "日期": r.get("date"),
            "编号": r.get("volunteer_id"),
            "姓名": r.get("name"),
            "报名": r.get("signup"),
            "签到": r.get("signin"),
            "岗位": r.get("role"),
            "开始时间": r.get("start_time"),
            "结束时间": r.get("end_time") or "",
            "时数": r.get("hours") or "",
            "备注": r.get("remark") or "",
            "_row": r.get("id"),
        })

    return result

def role_label(role: str, lang: str | None = None) -> str:
    lang = lang or get_lang()
    return ROLE_TEXT.get(role, {}).get(lang, role)

# 系统内部岗位永远用中文；英文只用于网页显示
ROLES = ["值班", "卫生", "佛台", "供台", "供花", "供果", "膳食", "佛学班"]

ROLE_TEXT = {
    "值班": {"zh": "值班", "en": "Duty"},
    "卫生": {"zh": "卫生", "en": "Cleaning"},
    "佛台": {"zh": "佛台", "en": "Altar"},
    "供台": {"zh": "供台", "en": "Offering Table"},
    "供花": {"zh": "供花", "en": "Flowers"},
    "供果": {"zh": "供果", "en": "Fruit Offering"},
    "膳食": {"zh": "膳食组", "en": "Meal Team"},
    "佛学班": {"zh": "佛学班", "en": "Buddhist Class"},
}


# =========================
# 5) 签到 / 签退 / 修改
# =========================
def sign_in(volunteer_id: str, pin: str, role: str, card_no: str = "") -> tuple[bool, str]:

    raw_id = str(volunteer_id or "").strip()
    role = str(role or "").strip()
    card_no = str(card_no or "").strip()

    if not raw_id:
        return False, "请输入编号。"

    if not pin:
        return False, "请输入 PIN。"

    volunteer = find_volunteer(raw_id)

    if not volunteer:
        return False, f"找不到编号：{raw_id}"

    if volunteer.get("状态") not in ["在册", "", None]:
        return False, f"此义工状态不是在册：{volunteer.get('状态')}"

    if not verify_pin_for_volunteer(volunteer, pin):
        return False, "PIN 不正确。"

    opened = db_query("""
        select id
        from attendance
        where date = %s
          and volunteer_id = %s
          and (end_time is null or end_time = '')
        order by id desc
        limit 1
    """, (now_date_str(), volunteer["编号"]), fetchone=True)

    if opened:
        return False, f"{volunteer['姓名']} 今天已经签到，还没签退。请先签退。"

    assignments = load_today_assignments_for_volunteer(volunteer, raw_id)

    if assignments:
        job_lines = [
            format_assignment_for_attendance(a)
            for a in assignments
        ]

        role_text = " / ".join(job_lines)

        db_query("""
            insert into attendance
            (date, volunteer_id, name, signup, signin, role, start_time, end_time, hours, card_no, remark)
            values (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """, (
            now_date_str(),
            volunteer["编号"],
            volunteer["姓名"],
            1,
            1,
            role_text,
            now_time_str(),
            "",
            None,
            card_no,
            "按排班签到"
        ))

        for a in assignments:
            db_query("""
                insert into volunteer_attendance_logs (
                    assignment_id,
                    volunteer_id,
                    name,
                    attendance_date,
                    actual_role,
                    actual_place,
                    walk_in,
                    remarks
                )
                values (%s, %s, %s, %s, %s, %s, false, '按排班签到')
            """, (
                a["id"],
                a["volunteer_id"],
                a["name"],
                a["assignment_date"],
                a["role"],
                a["assigned_place"],
            ))

        jobs_display = "\n".join([f"・{x}" for x in job_lines])

        return True, f"""{volunteer['姓名']} 已签到

今天岗位：
{jobs_display}

完成后请记得签退。"""

    # 没有排班，才需要义工选择岗位
    if role not in ROLES:
        return False, "今天没有排班记录，请选择实际协助岗位。"

    db_query("""
        insert into attendance
        (date, volunteer_id, name, signup, signin, role, start_time, end_time, hours, card_no, remark)
        values (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
    """, (
        now_date_str(),
        volunteer["编号"],
        volunteer["姓名"],
        0,
        1,
        role,
        now_time_str(),
        "",
        None,
        card_no,
        f"现场签到：{role}"
    ))

    db_query("""
        insert into volunteer_attendance_logs (
            assignment_id,
            volunteer_id,
            name,
            attendance_date,
            actual_role,
            actual_place,
            walk_in,
            remarks
        )
        values (null, %s, %s, %s, %s, %s, true, %s)
    """, (
        volunteer["编号"],
        volunteer["姓名"],
        now_date_str(),
        role,
        role,
        f"现场签到：{role}"
    ))

    return True, f"""{volunteer['姓名']} 已签到

    今天没有正式排班记录。

    已登记：
    {role}

    完成后请记得签退。"""

def sign_out(volunteer_id: str, pin: str) -> tuple[bool, str]:
    raw_id = str(volunteer_id or "").strip()

    if not raw_id:
        return False, "请输入编号。"
    if not pin:
        return False, "请输入 PIN。"

    volunteer = find_volunteer(raw_id)
    if not volunteer:
        return False, f"找不到编号：{raw_id}"

    if not verify_pin_for_volunteer(volunteer, pin):
        return False, "PIN 不正确。"

    row = db_query("""
        select *
        from attendance
        where date = %s
          and volunteer_id = %s
          and (end_time is null or end_time = '')
        order by id desc
        limit 1
    """, (now_date_str(), volunteer["编号"]), fetchone=True)

    if not row:
        return False, f"{volunteer['姓名']} 今天没有未签退记录。"

    end_time = now_time_str()

    try:
        start_dt = parse_time_to_datetime(row["start_time"])
        end_dt = parse_time_to_datetime(end_time)
        hours = round((end_dt - start_dt).total_seconds() / 3600, 2)
        if hours < 0:
            hours = 0
    except Exception:
        hours = None

    db_query("""
        update attendance
        set end_time = %s,
            hours = %s
        where id = %s
    """, (end_time, hours, row["id"]))

    db_query("""
        update volunteer_attendance_logs
        set check_out_time = now()
        where volunteer_id = %s
        and attendance_date = %s
        and check_out_time is null
    """, (
        volunteer["编号"],
        now_date_str(),
    ))

    return True, f"{volunteer['姓名']} 已签退。"

def find_volunteer(volunteer_id: str):
    raw_id = str(volunteer_id or "").strip()

    if not raw_id:
        return None

    ids = [raw_id]

    # CHE-108 / STW-160 → 同时尝试数字部分
    if "-" in raw_id:
        branch, num = raw_id.split("-", 1)
        branch = branch.strip().upper()
        num = num.strip()

        ids.append(num)

        # STW-160 也尝试 0160
        if branch == "STW" and num.isdigit():
            ids.append("0" + num)

    else:
        # 108 → 尝试 CHE-108
        ids.append(f"CHE-{raw_id}")

        # 0160 → 尝试 STW-160
        if raw_id.startswith("0") and raw_id[1:].isdigit():
            ids.append(f"STW-{raw_id[1:]}")

    # 去重
    ids = list(dict.fromkeys(ids))

    placeholders = ",".join(["%s"] * len(ids))

    result = db_query(f"""
        select
            id as "编号",
            name as "姓名",
            status as "状态",
            phone as "电话号码",
            pin as "PIN",
            branch as "分会"
        from volunteers
        where id in ({placeholders})
        limit 1
    """, tuple(ids), fetchone=True)

    return result

def verify_pin_for_volunteer(volunteer, pin):
    input_pin = str(pin).strip()

    # 1️⃣ 先用数据库 PIN（如果有）
    real_pin = volunteer.get("PIN")

    if real_pin:
        return input_pin == str(real_pin).strip()

    # 2️⃣ fallback → 电话后4位 or 0000
    phone = only_digits(volunteer.get("电话号码", ""))

    if not phone:
        return input_pin == "0000"   # ✅ 关键在这里

    return input_pin == phone[-4:]

def load_today_assignments_for_volunteer(volunteer, raw_id):

    ids = get_assignment_id_candidates(volunteer, raw_id)

    if not ids:
        return []

    return db_query("""
        select
            id,
            volunteer_id,
            name,
            assignment_date,
            role,
            shift_label,
            assigned_place,
            start_time,
            end_time
        from volunteer_schedule_assignments
        where volunteer_id = any(%s)
        and assignment_date = %s
        and coalesce(status,'assigned') <> 'cancelled'
    """, (
        ids,
        now_date_str()
    ), fetchall=True) or []

def format_assignment_for_attendance(a):
    role = a.get("role")
    place = a.get("assigned_place") or ""
    shift = a.get("shift_label") or ""
    start = a.get("start_time") or ""
    end = a.get("end_time") or ""

    if role == "值班":
        time_text = f"{start} ~ {end}" if start and end else ""
        return f"{time_text} {shift} {place}".strip()

    return place or role

def only_digits(value) -> str:
    return "".join(ch for ch in str(value or "") if ch.isdigit())

def get_assignment_id_candidates(volunteer, raw_id):
    ids = []

    if volunteer.get("编号"):
        ids.append(str(volunteer["编号"]).strip())

    if raw_id:
        ids.append(str(raw_id).strip())

    # 去重复
    return list(dict.fromkeys(ids))


def get_today_assignments(volunteer_id, role=None, current_time=False):

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:

            if role:
                cur.execute("""
                    select *
                    from volunteer_schedule_assignments
                    where volunteer_id = %s
                      and assignment_date = %s
                      and role = %s
                      and coalesce(status, 'assigned') <> 'cancelled'
                    order by start_time
                """, (
                    volunteer_id,
                    malaysia_today(),
                    role
                ))
            else:
                cur.execute("""
                    select *
                    from volunteer_schedule_assignments
                    where volunteer_id = %s
                      and assignment_date = %s
                      and coalesce(status, 'assigned') <> 'cancelled'
                    order by start_time
                """, (
                    volunteer_id,
                    malaysia_today()
                ))

            rows = cur.fetchall()

            print("volunteer_id =", volunteer_id)
            print("role =", role)
            print("today =", malaysia_today())
            print("rows =", rows)

            # 开关关闭：不检查时间，直接回传今天安排
            if not current_time or not ENABLE_SIGNIN_TIME_LIMIT:
                return rows

            now = malaysia_now()
            now_min = now.hour * 60 + now.minute

            valid_rows = []

            for r in rows:
                start_min = time_to_minutes(r.get("start_time"))
                end_min = time_to_minutes(r.get("end_time"))

                if start_min is None or end_min is None:
                    continue

                allowed_start = start_min - SIGNIN_EARLY_MINUTES

                if allowed_start <= now_min <= end_min:
                    valid_rows.append(r)

            print("valid_rows =", valid_rows)

            return valid_rows

    # 开关关闭：不检查时间，直接回传今天安排
    if not current_time or not ENABLE_SIGNIN_TIME_LIMIT:
        return rows

    now = malaysia_now()
    now_min = now.hour * 60 + now.minute

    valid_rows = []

    for r in rows:
        start_min = time_to_minutes(r.get("start_time"))
        end_min = time_to_minutes(r.get("end_time"))

        if start_min is None or end_min is None:
            continue

        allowed_start = start_min - SIGNIN_EARLY_MINUTES

        if allowed_start <= now_min <= end_min:
            valid_rows.append(r)

        print("volunteer_id =", volunteer_id)
        print("role =", role)
        print("today =", malaysia_today())
        print("rows =", rows)

    return valid_rows


def auto_signout_unfinished_today():
    if not ENABLE_AUTO_SIGNOUT:
        return {
            "ok": False,
            "msg": "自动签退功能未开启",
            "count": 0
        }

    today = now_date_str()

    rows = db_query("""
        select *
        from attendance
        where date = %s
          and (end_time is null or end_time = '')
    """, (today,), fetchall=True)

    count = 0

    for r in rows:
        db_query("""
            update attendance
            set end_time = %s,
                remarks = coalesce(remarks, '') || '｜系统自动签退'
            where id = %s
        """, (
            AUTO_SIGNOUT_TIME,
            r["id"]
        ))
        count += 1

    return {
        "ok": True,
        "msg": f"自动签退完成：{count} 位",
        "count": count
    }