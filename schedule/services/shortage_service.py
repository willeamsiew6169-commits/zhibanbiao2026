# shortage_service.py (完整版本，保留原有通知逻辑)
from datetime import datetime, timedelta, date, timezone
from schedule.services.assignment_service import load_assigned_places_for_date
from psycopg2.extras import RealDictCursor
from schedule.builders.time_utils import time_to_minutes, malaysia_today
from db import get_db, get_conn
from lunar_rules import get_special_day_info
from schedule.builders.schedule_builder import get_duty_targets

def build_shortage_summary_for_admin(date_str, conn=None):
    msg = build_shortage_notice_from_assignments(date_str, conn=conn) or ""
    lines = []
    for line in msg.splitlines():
        if line.startswith(("🧹", "🕗", "🕙", "🕑")):
            lines.append(line)
    return lines


def build_shortage_notice_from_assignments(date_str, conn=None):
    date_obj = datetime.strptime(date_str, "%Y-%m-%d").date()
    special_day_info = get_special_day_info(date_obj)
    template_type = special_day_info["template_type"]
    targets = get_duty_targets(template_type)
    rows = load_assigned_places_for_date(date_str, conn=conn)
    # 已排班按地点统计；待安排的有效报名也计入人数，但同一义工不重复计算。
    counts = {"绿观音堂": set(), "绿活动中心": set(), "橙观音堂": set(), "橙活动中心": set(),
              "黄观音堂": set(), "黄活动中心": set(), "卫生": set()}
    assigned_people = {"绿": set(), "橙": set(), "黄": set(), "卫生": set()}
    for r in rows:
        role = r.get("role")
        shift = str(r.get("shift_label") or "").replace("班", "")
        place = r.get("assigned_place")
        identity = str(r.get("volunteer_id") or r.get("name") or r.get("signup_id") or "").strip()
        if not identity:
            continue
        if role == "卫生":
            counts["卫生"].add(identity)
            assigned_people["卫生"].add(identity)
        elif role == "值班" and shift in ("绿", "橙", "黄"):
            key = f"{shift}{place}"
            if key in counts:
                counts[key].add(identity)
                assigned_people[shift].add(identity)

    def _read_signups(db_conn):
        with db_conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                SELECT volunteer_id, name, role, start_time, end_time
                FROM volunteer_schedule_signups
                WHERE signup_date = %s
                  AND COALESCE(status, 'pending') <> 'cancelled'
            """, (date_str,))
            return cur.fetchall()

    if conn is None:
        with get_conn() as db_conn:
            signups = _read_signups(db_conn)
    else:
        signups = _read_signups(conn)

    pending = {"绿": set(), "橙": set(), "黄": set(), "卫生": set()}
    periods = {"绿": (8 * 60, 10 * 60), "橙": (10 * 60, 14 * 60), "黄": (14 * 60, 18 * 60)}
    for r in signups:
        identity = str(r.get("volunteer_id") or r.get("name") or "").strip()
        if not identity:
            continue
        if r.get("role") == "卫生":
            if identity not in assigned_people["卫生"]:
                pending["卫生"].add(identity)
        elif r.get("role") == "值班":
            start = time_to_minutes(r.get("start_time"))
            end = time_to_minutes(r.get("end_time"))
            if start is None or end is None:
                continue
            for shift, (shift_start, shift_end) in periods.items():
                if start <= shift_start and end >= shift_end and identity not in assigned_people[shift]:
                    pending[shift].add(identity)

    def _need(shift, places):
        # 已安排的按各地点最低人数计算；未安排者可补其中一个缺额。
        missing = sum(max(0, targets.get(f"{shift}{place}", 0) - len(counts[f"{shift}{place}"]))
                      for place in places)
        return max(0, missing - len(pending[shift]))

    early_need = _need("绿", ("观音堂", "活动中心"))
    morning_need = _need("橙", ("观音堂", "活动中心"))
    afternoon_need = _need("黄", ("观音堂", "活动中心"))
    cleaning_need = max(0, targets["卫生"] - len(counts["卫生"]) - len(pending["卫生"]))
    today = malaysia_today()
    if date_obj == today:
        day_text = "今日"
    elif date_obj == today + timedelta(days=1):
        day_text = "明日"
    else:
        day_text = f"{date_obj.month}月{date_obj.day}日"
    date_text = f"{date_obj.month}月{date_obj.day}日"
    if early_need == 0 and morning_need == 0 and afternoon_need == 0 and cleaning_need == 0:
        return f"✅ {day_text}义工已足够，无需发送通知。"
    msg = "师兄们，大家好！🙏\n\n" + f"📅 {day_text}（{date_text}）义工尚需\n\n"
    if day_text != "今日" and cleaning_need > 0:
        msg += f"🧹 卫生义工　{cleaning_need} 位\n"
    if early_need > 0:
        msg += f"🕗 8:00am ～ 10:00am　{early_need} 位\n"
    if morning_need > 0:
        msg += f"🕙 10:00am ～ 2:00pm　{morning_need} 位\n"
    if afternoon_need > 0:
        msg += f"🕑 2:00pm ～ 6:00pm　{afternoon_need} 位\n"
    msg += "\n欢迎有空的师兄随喜报名护持。\n\n感恩大家 🙏🙏🙏"
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
    today = malaysia_today()
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
