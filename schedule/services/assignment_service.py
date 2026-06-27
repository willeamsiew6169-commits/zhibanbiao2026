# assignment_service.py

import os
import pandas as pd

from db import get_conn
from psycopg2.extras import RealDictCursor
from schedule.constants import PREBOOK_FILE
from datetime import datetime, date, timedelta
from schedule.builders.time_utils import malaysia_today


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
    

def load_schedule_admin_dashboard_data(override_date):
    
    today = malaysia_today()
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