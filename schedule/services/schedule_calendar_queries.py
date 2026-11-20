# Extracted unchanged from schedule_web.py (stage 4).
from datetime import timedelta
from db import db_query
from schedule.builders.time_utils import malaysia_today
from schedule.builders.schedule_builder import get_special_day_info


def get_pending_signup_counts():
    
    today = malaysia_today()

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
        and role IS DISTINCT FROM '膳食'
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

def get_pending_signup_details():
    """与 get_pending_signup_counts 的 month_count 使用完全相同的筛选范围。"""
    month_start = malaysia_today().replace(day=1)
    return db_query("""
        SELECT id, signup_date, name, volunteer_id, role, start_time, end_time
        FROM volunteer_schedule_signups
        WHERE COALESCE(status, 'pending') = 'pending'
          AND role IS DISTINCT FROM '膳食'
          AND signup_date >= %s
        ORDER BY signup_date, start_time NULLS LAST, id
    """, (month_start,), fetchall=True) or []


def is_big_day(date_str):
    from datetime import datetime

    date_obj = datetime.strptime(date_str, "%Y-%m-%d").date()
    info = get_special_day_info(date_obj)

    return info.get("template_type") in [
        "lunar_1_15",
        "buddhist_festival"
    ]

def _master_table_dates(days=90):
    """Use the existing lunar rules, not a separately maintained festival calendar."""
    today = malaysia_today()
    dates = []
    for offset in range(days + 1):
        day = today + timedelta(days=offset)
        info = get_special_day_info(day) or {}
        if not info.get("setup_shifu"):
            continue
        title = (info.get("festival_name") or info.get("name") or
                 info.get("title") or info.get("lunar_text") or "大日子")
        dates.append({"date": day.isoformat(), "label": f"{day:%Y-%m-%d} · {title}"})
    return dates
