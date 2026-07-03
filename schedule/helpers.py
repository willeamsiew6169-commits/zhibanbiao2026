# helpers.py

import os
import calendar
import pandas as pd

from opencc import OpenCC
from db import get_db, get_conn
from psycopg2.extras import RealDictCursor
from schedule.builders.time_utils import malaysia_now
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

def clean_phone(phone):
    """只保留数字"""
    return "".join(ch for ch in str(phone or "") if ch.isdigit())


def find_volunteer_by_keyword(keyword):

    keyword = str(keyword or "").strip()

    if not keyword:
        return []

    key_simple = to_simple(keyword)
    keyword_upper = keyword.upper()
    keyword_phone = clean_phone(keyword)

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:

            cur.execute("""
                select
                    id,
                    name,
                    phone,
                    branch,
                    status
                from volunteers
                where status='在册'
            """)

            volunteers = cur.fetchall()

    # =========
    # 1. 编号 / 电话 完全符合
    # =========

    exact_matches = []

    keyword_num = None

    if keyword.isdigit():
        keyword_num = int(keyword)

    for v in volunteers:

        vid = str(v["id"] or "").strip().upper()

        phone = clean_phone(v.get("phone"))

        # 输入完整编号
        if vid == keyword_upper:
            exact_matches.append(v)
            continue

        # 输入数字，例如 69
        if keyword_num is not None:

            if (
                vid == f"CHE-{keyword_num}"
                or vid == f"STW-{keyword_num}"
            ):
                exact_matches.append(v)
                continue

        # 电话
        if keyword_phone and phone == keyword_phone:
            exact_matches.append(v)
            continue

    if exact_matches:
        return exact_matches

    # =========
    # 2. 姓名（简繁体）
    # =========

    name_matches = []

    for v in volunteers:

        name = str(v["name"] or "").strip()

        if key_simple in to_simple(name):
            name_matches.append(v)

    return name_matches

def get_daily_buddha_quote():
    """
    从 daily_quotes.xlsx 读取每日佛言佛语。
    马来西亚时间每天 18:00 后自动切换到下一句。
    """

    try:
        file_path = os.path.join(
            os.path.dirname(__file__),
            "daily_quotes.xlsx"
        )

        df = pd.read_excel(file_path)

        df = df[df["active"] == True]

        if df.empty:
            return "发心护持道场，就是培福培慧。"

        now = malaysia_now()

        if now.hour >= 18:
            quote_date = now.date() + timedelta(days=1)
        else:
            quote_date = now.date()

        idx = quote_date.toordinal() % len(df)

        return str(df.iloc[idx]["content"]).strip()

    except Exception as e:
        print("读取每日佛言失败:", e)
        return "发心护持道场，就是培福培慧。"

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
            "膳食": [],
            "值班": {
                "观音堂": [],
                "活动中心": [],
            },
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
            text = name

            if start_time and end_time:
                text += f"（{start_time}-{end_time}）"

            if text not in by_date[d]["供台"]:
                by_date[d]["供台"].append(text)

        elif role == "膳食":
            text = name

            if start_time and end_time:
                text += f"（{start_time}-{end_time}）"

            if text not in by_date[d]["膳食"]:
                by_date[d]["膳食"].append(text)

        elif role == "值班":
            idx = duty_count.get(name, 0) % len(duty_places)
            place = duty_places[idx]
            duty_count[name] = duty_count.get(name, 0) + 1

            text = name

            if start_time and end_time:
                text += f"（{start_time}-{end_time}）"

            if text not in by_date[d]["值班"][place]:
                by_date[d]["值班"][place].append(text)

    lines = []
    lines.append(f"{year}年{month}月份预报名名单")
    lines.append("")

    weekday_cn = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]

    for day in range(1, last_day_num + 1):
        d = date(year, month, day)
        weekday = weekday_cn[d.weekday()]

        data = by_date.get(d, {
            "卫生": [],
            "供台": [],
            "膳食": [],
            "值班": {
                "观音堂": [],
                "活动中心": [],
            },
        })

        lines.append(f"{day}/{month}/{year}  {weekday}")

        lines.append("卫生：")
        if data["卫生"]:
            for i, item in enumerate(data["卫生"], start=1):
                lines.append(f"{i}）{item}")
        else:
            lines.append("")

        if data["供台"]:
            lines.append("供台：")
            for i, item in enumerate(data["供台"], start=1):
                lines.append(f"{i}）{item}")

        if data["膳食"]:
            lines.append("膳食：")
            for i, item in enumerate(data["膳食"], start=1):
                lines.append(f"{i}）{item}")

        lines.append("值班：")

        for place in duty_places:
            people = data["值班"].get(place, [])

            if people:
                lines.append(f"{place}：")

                for i, item in enumerate(people, start=1):
                    lines.append(f"{i}）{item}")

        lines.append("")

    return "\n".join(lines)