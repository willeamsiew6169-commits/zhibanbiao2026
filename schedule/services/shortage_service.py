# shortage_service.py (完整版本，保留原有通知逻辑)
from datetime import datetime, timedelta, date, timezone
from schedule.services.assignment_service import load_assigned_places_for_date
from psycopg2.extras import RealDictCursor
from schedule.builders.time_utils import time_to_minutes, malaysia_today
from db import get_db, get_conn
from lunar_rules import get_special_day_info
from schedule.builders.schedule_builder import get_duty_targets

def _identity(row):
    return str(row.get("volunteer_id") or row.get("name") or "").strip()


def _minute(value):
    if hasattr(value, "hour"):
        return value.hour * 60 + value.minute
    return time_to_minutes(value)


def _clock(value):
    h, m = divmod(value, 60)
    return f"{h % 12 or 12}:{m:02d}{'am' if h < 12 else 'pm'}"


def _span(start, end):
    return f"{_clock(start)} ～ {_clock(end)}"


def _coverage(rows, signups, published, big_day):
    places = ("观音堂", "活动中心")
    periods = ([(480, 600)] if big_day else []) + [(600, 840), (840, 1080)]
    active_ids = {str(r['id']) for r in signups if r.get('id') is not None}
    rows = [r for r in rows if r.get('status') != 'cancelled' and
            (r.get('signup_id') is None or str(r['signup_id']) in active_ids)]
    cleaning = {_identity(r) for r in rows + signups if r.get('role') == '卫生'} - {''}
    assigned = []
    available = []
    for source, dest in ((rows, assigned), (signups, available)):
        for r in source:
            if r.get('role') != '值班' or not _identity(r):
                continue
            a, b = _minute(r.get('start_time')), _minute(r.get('end_time'))
            if a is not None and b is not None and a < b:
                dest.append((a, b, _identity(r), r.get('assigned_place')))
    output = []
    for start, end in periods:
        points = sorted({start, end} | {t for a,b,_,_ in assigned + available for t in (a,b) if start < t < end})
        segments = {}
        for a,b in zip(points, points[1:]):
            by_place = {p: {i for x,y,i,z in assigned if x <= a and y >= b and z == p} for p in places}
            if published:
                needs = [(p, max(0, 1-len(ids)), max(0, 2-len(ids))) for p,ids in by_place.items()]
            else:
                # 尚未发布：地点未确定，不把待安排者虚构成某地点的义工。
                people = set().union(*by_place.values()) | {i for x,y,i,_ in available if x <= a and y >= b}
                needs = [("地点待安排", max(0, 2-len(people)), max(0, 4-len(people)))]
            for place, minimum, ideal in needs:
                if not ideal:
                    continue
                key = (place, minimum, ideal)
                group = segments.setdefault(key, [])
                if group and group[-1][1] == a:
                    group[-1] = (group[-1][0], b)
                else:
                    group.append((a,b))
        for (place, minimum, ideal), spans in segments.items():
            output.extend((place,a,b,minimum,ideal) for a,b in spans)
    return len(cleaning), output


def _load_coverage(date_str, conn=None):
    if conn is None:
        with get_conn() as db_conn:
            return _load_coverage(date_str, db_conn)
    day = datetime.strptime(date_str, "%Y-%m-%d").date()
    big_day = get_special_day_info(day)["template_type"] != "normal"
    rows = load_assigned_places_for_date(date_str, conn=conn)
    with conn.cursor(cursor_factory=RealDictCursor) as cur:
        cur.execute("""SELECT id, volunteer_id, name, role, start_time, end_time
                       FROM volunteer_schedule_signups WHERE signup_date = %s
                       AND COALESCE(status, 'pending') <> 'cancelled'""", (date_str,))
        signups = cur.fetchall()
        cur.execute("SELECT is_published FROM schedule_day_flags WHERE flag_date = %s", (date_str,))
        flag = cur.fetchone()
    return _coverage(rows, signups, bool(flag and flag['is_published']), big_day)


def _need_text(minimum, ideal):
    return f"{minimum}～{ideal}" if 0 < minimum < ideal else str(ideal)


def _cleaning_text(count):
    return {0: "2 ~ 3", 1: "1 ~ 2", 2: "1"}.get(count, "")


def build_shortage_summary_for_admin(date_str, conn=None):
    cleaning, gaps = _load_coverage(date_str, conn)
    # 仅负责人当天提醒：已结束时段不再显示，未来日期仍显示全天。
    day = datetime.strptime(date_str, "%Y-%m-%d").date()
    now = datetime.now(timezone(timedelta(hours=8)))
    cutoff = now.hour * 60 + now.minute if day == now.date() else -1
    lines = []
    for place,a,b,minimum,ideal in sorted(gaps, key=lambda g: (not bool(g[3]), g[1], g[0])):
        if b <= cutoff:
            continue
        label = ("🔴 无人值班" if minimum else "🟠 建议增员") if place != "地点待安排" else "🟠 报名不足（地点待安排）"
        lines.append(f"{label}｜{place}｜{_span(a,b)}｜缺 {_need_text(minimum,ideal)} 位")
    if cutoff < 600 and _cleaning_text(cleaning):
        # 大日子卫生 6–8 点，普通日 8–10 点。
        cleaning_end = 480 if get_special_day_info(day)["template_type"] != "normal" else 600
        if cutoff < cleaning_end:
            lines.append(f"🧹 卫生义工　还需要 {_cleaning_text(cleaning)} 位")
    return lines


def build_shortage_notice_from_assignments(date_str, conn=None):
    cleaning, gaps = _load_coverage(date_str, conn)
    day = datetime.strptime(date_str, "%Y-%m-%d").date()
    today = malaysia_today()
    day_text = "今日" if day == today else "明日" if day == today + timedelta(days=1) else f"{day.month}月{day.day}日"
    lines = []
    if day != today and _cleaning_text(cleaning):
        lines.append(f"🧹 卫生义工　{_cleaning_text(cleaning)} 位")
    # 群通知只显示最低覆盖不足，不显示负责人内部的建议增员。
    now = datetime.now(timezone(timedelta(hours=8)))
    cutoff = now.hour * 60 + now.minute if day == today else -1
    for place in ("观音堂", "活动中心", "地点待安排"):
        merged = []
        for a, b in sorted((a, b) for p, a, b, minimum, ideal in gaps
                           if p == place and minimum > 0 and b > cutoff):
            if merged and a <= merged[-1][1]:
                merged[-1] = (merged[-1][0], max(b, merged[-1][1]))
            else:
                merged.append((a, b))
        lines.extend(f"🕒 {place} {_clock(a)} ~ {_clock(b)}" for a, b in merged)
    if not lines:
        return f"✅ {day_text}没有需要发送的缺义工通知。"
    return (f"师兄们，大家好！🙏\n\n📅 {day_text}（{day.month}月{day.day}日）义工尚需\n\n"
            + "\n".join(lines) + "\n\n欢迎有空的师兄随喜报名护持。\n\n感恩大家 🙏🙏🙏")


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
    if cleaning_count < 3:
        need_text = "2 ~ 3" if cleaning_count == 0 else ("1 ~ 2" if cleaning_count == 1 else "1")
        notices.append(f"🧹 卫生 还需要 {need_text} 位义工。")
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
