# attendance_service.py

import datetime

from db import db_query, get_conn
from signin_policy import get_signin_policies
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
ENABLE_AUTO_SIGNOUT = True           # 自动签退
AUTO_SIGNOUT_TIME = "19:00"           # 系统几点执行自动签退
AUTO_SIGNOUT_DISPLAY = "18:30"        # 报表显示的签退时间
REQUIRE_ASSIGNMENT_FOR_SIGNIN = False   # 值班：正式排班限制（试用期关闭）
REQUIRE_CLEANING_SIGNUP_FOR_SIGNIN = False  # 卫生：当天有效报名即可，不必正式排班
REQUIRE_MEAL_SIGNUP_FOR_SIGNIN = False      # 膳食：预留独立限制，正式启用前关闭


def has_active_signup_for_role(volunteer, raw_id, role):
    """只认本人、当天、对应岗位的未取消报名；不把其他岗位报名算进去。"""
    ids = get_assignment_id_candidates(volunteer, raw_id)
    if not ids:
        return False
    row = db_query("""
        select 1
        from volunteer_schedule_signups
        where volunteer_id = any(%s)
          and signup_date = %s
          and role = %s
          and coalesce(status, 'pending') <> 'cancelled'
        limit 1
    """, (ids, now_date_str(), role), fetchone=True)
    return bool(row)



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
def same_role_contiguous_group(assignments, selected):
    """Combine only directly adjacent duty segments; never combine different roles or gaps."""
    if selected.get("role") != "值班":
        return [selected]
    ordered = sorted(
        (a for a in assignments if a.get("role") == "值班"),
        key=lambda a: (time_to_minutes(a.get("start_time")) if time_to_minutes(a.get("start_time")) is not None else 9999, int(a.get("id") or 0))
    )
    idx = next((i for i, a in enumerate(ordered) if int(a["id"]) == int(selected["id"])), None)
    if idx is None:
        return [selected]
    left = right = idx
    while left > 0:
        prev_end = time_to_minutes(ordered[left - 1].get("end_time"))
        curr_start = time_to_minutes(ordered[left].get("start_time"))
        if prev_end is None or curr_start is None or prev_end != curr_start:
            break
        left -= 1
    while right + 1 < len(ordered):
        curr_end = time_to_minutes(ordered[right].get("end_time"))
        next_start = time_to_minutes(ordered[right + 1].get("start_time"))
        if curr_end is None or next_start is None or curr_end != next_start:
            break
        right += 1
    return ordered[left:right + 1]


def sign_in(volunteer_id: str, pin: str, role: str, card_no: str = "", assignment_id=None) -> tuple[bool, str]:
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

    if role not in ROLES:
        return False, "请选择正确岗位。"

    # 同一时间只允许一笔进行中的签到。
    opened = db_query("""
        select id from attendance
        where date=%s and volunteer_id=%s
          and (end_time is null or end_time='')
        order by id desc limit 1
    """, (now_date_str(), volunteer["编号"]), fetchone=True)
    if opened:
        return False, f"{volunteer['姓名']} 已经签到，还没签退。请先签退。"

    assignments = load_today_assignments_for_volunteer(volunteer, raw_id)

    # 每次签到读取 Supabase 最新开关；不依赖重启或 Python 常量。
    policies = get_signin_policies()
    require_duty = policies["duty"]
    require_cleaning = policies["cleaning"]
    require_meal = policies["meal"]

    # 按岗位独立控制；卫生只需报名，不需负责人正式安排。
    # 注意：不能用值班排班冒充卫生/膳食报名。
    if role == "值班" and require_duty:
        if not any(a.get("role") == "值班" for a in assignments):
            return False, "今天没有有效的值班正式排班，不能签到。"
    elif role == "卫生" and require_cleaning:
        if not has_active_signup_for_role(volunteer, raw_id, "卫生"):
            return False, "今天没有有效的卫生报名，请先报名后再签到。"
    elif role == "膳食" and require_meal:
        if not has_active_signup_for_role(volunteer, raw_id, "膳食"):
            return False, "今天没有有效的膳食组报名，请先报名后再签到。"

    # 即使限制关闭，也不允许把其他岗位的 assignment_id 当作本岗位。
    assignments = [a for a in assignments if a.get("role") == role]

    # 有正式排班时：一轮 assignment = 一次签到。
    if assignments:
        selected = None
        if assignment_id not in (None, ""):
            try:
                aid = int(assignment_id)
            except Exception:
                aid = 0
            selected = next((a for a in assignments if int(a.get("id") or 0) == aid), None)
            if not selected:
                return False, "找不到这轮正式排班，请重新查询后再签到。"
        else:
            # 兼容旧页面：自动找第一轮尚未完成的排班。
            assignment_ids = [int(a["id"]) for a in assignments]
            done_rows = db_query("""
                select distinct assignment_id from volunteer_attendance_logs
                where assignment_id = any(%s) and volunteer_id=%s
                  and attendance_date=%s and check_out_time is not null
            """, (assignment_ids, volunteer["编号"], now_date_str()), fetchall=True) or []
            done_ids = {int(r["assignment_id"]) for r in done_rows if r.get("assignment_id") is not None}
            selected = next((a for a in assignments if int(a["id"]) not in done_ids), None)

        if not selected:
            return False, f"{volunteer['姓名']} 今天的正式排班已经全部完成。"

        group = same_role_contiguous_group(assignments, selected)
        group_ids = [int(a["id"]) for a in group]
        # Any segment already checked in means the group cannot be started again.
        existing_log = db_query("""
            select assignment_id, check_out_time from volunteer_attendance_logs
            where assignment_id = any(%s) and volunteer_id=%s and attendance_date=%s
            order by id desc limit 1
        """, (group_ids, volunteer["编号"], now_date_str()), fetchone=True)
        if existing_log:
            return False, "这段值班已经签到或完成，请勿重复签到。"

        job_text = "；".join(format_assignment_for_attendance(a) for a in group)
        # 同一事务写入实际签到及各段排班日志；避免只写入其中一张表。
        with get_conn() as conn:
            with conn.cursor() as cur:
                cur.execute("""
                    insert into attendance
                    (date, volunteer_id, name, signup, signin, role, start_time, end_time, hours, card_no, remark)
                    values (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                """, (
                    now_date_str(), volunteer["编号"], volunteer["姓名"], 1, 1,
                    role, now_time_str(), "", None, card_no,
                    f"按排班签到｜{job_text}｜assignment_ids={','.join(map(str, group_ids))}"
                ))
                from psycopg2.extras import execute_values
                execute_values(cur, """
                    insert into volunteer_attendance_logs (
                        assignment_id, volunteer_id, name, attendance_date,
                        actual_role, actual_place, walk_in, remarks
                    ) values %s
                """, [(
                    segment["id"], segment["volunteer_id"], segment["name"],
                    segment["assignment_date"], segment["role"], segment["assigned_place"],
                    False, "按排班签到（同岗位连续时段）"
                ) for segment in group])

        return True, f"{volunteer['姓名']} 已签到\n\n岗位：{role}\n安排：{job_text}\n\n同岗位连续换地点无需重新签到；完成后请签退。"

    # 没有该岗位正式排班：允许符合政策的现场签到。
    if role not in ROLES:
        return False, "今天没有排班记录，请选择实际协助岗位。"

    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                insert into attendance
                (date, volunteer_id, name, signup, signin, role, start_time, end_time, hours, card_no, remark)
                values (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
            """, (
                now_date_str(), volunteer["编号"], volunteer["姓名"], 0, 1, role,
                now_time_str(), "", None, card_no, f"现场签到：{role}"
            ))
            cur.execute("""
                insert into volunteer_attendance_logs (
                    assignment_id, volunteer_id, name, attendance_date,
                    actual_role, actual_place, walk_in, remarks
                ) values (null,%s,%s,%s,%s,%s,true,%s)
            """, (
                volunteer["编号"], volunteer["姓名"], now_date_str(), role, role,
                f"现场签到：{role}"
            ))
    return True, f"{volunteer['姓名']} 已签到\n\n今天没有正式排班记录。\n已登记：{role}\n\n完成后请记得签退。"


def sign_out(volunteer_id: str, pin: str, assignment_id=None) -> tuple[bool, str]:
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

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                select * from attendance
                where date=%s and volunteer_id=%s
                  and (end_time is null or end_time='')
                order by id desc limit 1
                for update
            """, (now_date_str(), volunteer["编号"]))
            row = cur.fetchone()
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

            # 有正式排班时，服务时数不能超过最后一段排班结束时间。
            scheduled_end, _ = _scheduled_end_minutes(cur, row)
            if scheduled_end is not None:
                actual_end = time_to_minutes(end_time)
                actual_start = time_to_minutes(row.get("start_time"))
                if actual_end is not None and actual_start is not None:
                    credited_end = min(actual_end, scheduled_end)
                    hours = round(max(0, credited_end - actual_start) / 60, 2)

            # 签退更新两张表必须同一事务提交。
            remark = str(row.get("remark") or "")
            group_marker = "assignment_ids="
            legacy_marker = "assignment_id="
            aids = []
            if group_marker in remark:
                try:
                    value = remark.split(group_marker, 1)[1].split("｜", 1)[0].strip()
                    aids = [int(x) for x in value.split(",") if x.strip()]
                except ValueError:
                    aids = []
            elif legacy_marker in remark:
                try:
                    aids = [int(remark.split(legacy_marker, 1)[1].split("｜", 1)[0].strip())]
                except ValueError:
                    aids = []

            cur.execute("update attendance set end_time=%s, hours=%s where id=%s",
                        (end_time, hours, row["id"]))
            if aids:
                cur.execute("""
                    update volunteer_attendance_logs set check_out_time=now()
                    where assignment_id = any(%s) and volunteer_id=%s
                      and attendance_date=%s and check_out_time is null
                """, (aids, volunteer["编号"], now_date_str()))
            else:
                cur.execute("""
                    update volunteer_attendance_logs set check_out_time=now()
                    where id = (
                        select id from volunteer_attendance_logs
                        where volunteer_id=%s and attendance_date=%s
                          and check_out_time is null and assignment_id is null
                        order by id desc limit 1
                    )
                """, (volunteer["编号"], now_date_str()))

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

    # 兼容仍使用800系列旧义工编号、但已有正式月费编号的人员。
    # 只在正式义工编号查询不到时查找，且必须唯一匹配，避免误认他人。
    if not result:
        member_ids = [x for x in ids if x.upper().startswith(('CHE-', 'STW-'))]
        if member_ids:
            matched = db_query("""
                select id as "编号", name as "姓名", status as "状态",
                       phone as "电话号码", pin as "PIN", branch as "分会"
                from volunteers where member_id = any(%s)
                limit 2
            """, (member_ids,), fetchall=True) or []
            if len(matched) == 1:
                return matched[0]
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


def _assignment_ids_from_remark(remark):
    """Extract the original scheduled segments from an attendance row."""
    remark = str(remark or "")
    for marker in ("assignment_ids=", "assignment_id="):
        if marker in remark:
            raw = remark.split(marker, 1)[1].split("｜", 1)[0].strip()
            try:
                return [int(x.strip()) for x in raw.split(",") if x.strip()]
            except ValueError:
                return []
    return []


def _scheduled_end_minutes(cur, row):
    """Only scheduled attendance can be capped/auto-closed."""
    ids = _assignment_ids_from_remark(row.get("remark"))
    if not ids:
        return None, []
    cur.execute("""
        select id, end_time from volunteer_schedule_assignments
        where id = any(%s) and assignment_date::text = %s
    """, (ids, str(row["date"])))
    assignments = cur.fetchall() or []
    if len(assignments) != len(set(ids)):
        return None, []
    ends = [time_to_minutes(a["end_time"]) for a in assignments]
    if any(x is None for x in ends):
        return None, []
    return max(ends), ids


def auto_signout_unfinished_today():
    """Close scheduled shifts 30 min after their end; no fake actual checkout time.

    Called on normal web requests (no background scheduler required).  The
    attendance end_time is the *credited cutoff*, clearly marked as automatic;
    log check_out_time is the actual time the system processed the closure.
    """
    if not ENABLE_AUTO_SIGNOUT:
        return {"ok": False, "msg": "自动签退功能未开启", "count": 0}

    now = malaysia_now()
    today = now.date().isoformat()  # attendance.date 是 text，不能传 Python date
    now_minutes = now.hour * 60 + now.minute
    count = 0
    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                select * from attendance
                where date=%s and (end_time is null or end_time='')
                for update skip locked
            """, (today,))
            for row in cur.fetchall() or []:
                cutoff, ids = _scheduled_end_minutes(cur, row)
                if cutoff is None or now_minutes < cutoff + 30:
                    continue
                credited_end = f"{cutoff // 60:02d}:{cutoff % 60:02d}"
                start_min = time_to_minutes(row.get("start_time"))
                hours = round(max(0, cutoff - start_min) / 60, 2) if start_min is not None else None
                remark = str(row.get("remark") or "") + "｜系统自动签退（忘记签退；服务计至排班结束）"
                cur.execute("""
                    update attendance set end_time=%s, hours=%s, remark=%s
                    where id=%s
                """, (credited_end, hours, remark, row["id"]))
                cur.execute("""
                    update volunteer_attendance_logs
                    set check_out_time=now(),
                        remarks=concat_ws('｜', nullif(remarks, ''), '系统自动签退；服务计至排班结束')
                    where assignment_id=any(%s) and volunteer_id=%s
                      and attendance_date=%s and check_out_time is null
                """, (ids, row["volunteer_id"], today))
                count += 1
    return {"ok": True, "msg": f"自动签退完成：{count} 位", "count": count}
