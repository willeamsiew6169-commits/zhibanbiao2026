# shortage_service.py

from datetime import datetime, timedelta, date, timezone

from schedule.services.assignment_service import load_assigned_places_for_date
from psycopg2.extras import RealDictCursor
from schedule.builders.time_utils import time_to_minutes
from db import get_db, get_conn
from lunar_rules import (
    get_special_day_info,
)

from schedule.builders.schedule_builder import (
    get_duty_targets,
)

def build_shortage_summary_for_admin(date_str):
    msg = build_shortage_notice_from_assignments(date_str)

    lines = []

    for line in msg.splitlines():
        if line.startswith("🔴") or line.startswith("🟡"):
            lines.append(line)

    return lines


def build_shortage_notice_from_assignments(date_str):

    date_obj = datetime.strptime(
        date_str,
        "%Y-%m-%d"
    ).date()

    special_day_info = get_special_day_info(date_obj)

    template_type = special_day_info["template_type"]

    targets = get_duty_targets(template_type)

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

    for key, target in targets.items():

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

    msg += f"📅 日期：{date_str}\n"
    msg += f"🌙 {special_day_info['lunar_text']}\n"

    if special_day_info["is_special"]:
        msg += f"🛕 {'、'.join(special_day_info['special_names'])}\n"

    msg += f"📋 模板：{special_day_info['template_text']}\n\n"

    msg += "明天义工岗位情况：\n\n"

    msg += "\n".join(shortages)

    msg += "\n\n欢迎大家随缘发心护持观音堂。\n\n感恩大家 🙏🙏🙏"

    return msg

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