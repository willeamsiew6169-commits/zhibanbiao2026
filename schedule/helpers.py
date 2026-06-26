# helpers.py

import calendar

from opencc import OpenCC
from db import get_db, get_conn
from psycopg2.extras import RealDictCursor
from datetime import datetime, timedelta, date, timezone

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

def build_monthly_signup_text(year, month):
    year = int(year)
    month = int(month)

    first_day = date(year, month, 1)
    last_day_num = calendar.monthrange(year, month)[1]
    last_day = date(year, month, last_day_num)

    cleaning_places = ["佛堂卫生", "二楼卫生", "楼梯卫生"]
    duty_places = ["观音堂", "活动中心"]

    cleaning_count = {}
    duty_count = {}

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                select *
                from volunteer_schedule_signups
                where signup_date between %s and %s
                and coalesce(status, 'pending') <> 'cancelled'
                order by signup_date, start_time, name
            """, (first_day, last_day))
            rows = cur.fetchall()

    by_date = {}

    for r in rows:
        d = r["signup_date"]

        if isinstance(d, datetime):
            d = d.date()

        by_date.setdefault(d, {
            "卫生": [],
            "供台": [],
            "值班": [],
        })

        role = r.get("role") or ""
        name = r.get("name") or ""
        start_time = r.get("start_time") or ""
        end_time = r.get("end_time") or ""

        if not name:
            continue

        if role == "卫生":
            idx = cleaning_count.get(name, 0) % len(cleaning_places)
            place = cleaning_places[idx]
            cleaning_count[name] = cleaning_count.get(name, 0) + 1

            text = f"{place}：{name}"

            if text not in by_date[d]["卫生"]:
                by_date[d]["卫生"].append(text)

        elif role == "供台":
            if name not in by_date[d]["供台"]:
                by_date[d]["供台"].append(name)

        else:
            idx = duty_count.get(name, 0) % len(duty_places)
            place = duty_places[idx]
            duty_count[name] = duty_count.get(name, 0) + 1

            text = f"{place}：{name}"

            if start_time and end_time:
                text += f"（{start_time}-{end_time}）"

            if text not in by_date[d]["值班"]:
                by_date[d]["值班"].append(text)

    lines = []
    lines.append(f"{year}年{month}月份预报名名单")
    lines.append("")

    weekday_cn = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]

    for day in range(1, last_day_num + 1):
        d = date(year, month, day)
        weekday = weekday_cn[d.weekday()]

        lines.append(f"{day}/{month}/{year}  {weekday}")

        data = by_date.get(d, {
            "卫生": [],
            "供台": [],
            "值班": [],
        })

        lines.append("卫生：" + "、".join(data["卫生"]))

        if data["供台"]:
            lines.append("供台：" + "、".join(data["供台"]))

        lines.append("值班：" + "、".join(data["值班"]))
        lines.append("")

    return "\n".join(lines)
