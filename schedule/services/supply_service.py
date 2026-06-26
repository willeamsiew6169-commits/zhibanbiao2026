# supply_service.py

from db import get_conn
from psycopg2.extras import RealDictCursor
from lunar_rules import get_special_day_info
from datetime import (
    datetime,
    timedelta,
    date
)


def load_supply_signups_for_date(date_str):
    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                select volunteer_id, name
                from volunteer_schedule_signups
                where signup_date = %s
                and role = '供台'
                and coalesce(status, 'pending') <> 'cancelled'
                order by created_at, name
            """, (date_str,))
            return cur.fetchall()
        

def load_upcoming_supply_signup_alerts(days_ahead=60, limit=2):
    today = date.today()
    end_date = today + timedelta(days=days_ahead)

    special_dates = []

    d = today
    while d <= end_date:
        info = get_special_day_info(d)

        if info["template_type"] in ["lunar_1_15", "buddhist_festival"]:
            special_dates.append({
                "date": d,
                "type": info["template_type"],
                "names": []
            })

        d += timedelta(days=1)

    if not special_dates:
        return []

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            for item in special_dates:
                cur.execute("""
                    select name
                    from volunteer_schedule_signups
                    where signup_date = %s
                    and role = '供台'
                    and coalesce(status, 'pending') <> 'cancelled'
                    order by name
                """, (item["date"],))

                rows = cur.fetchall()
                item["names"] = [r["name"] for r in rows]

    return special_dates[:limit]


def load_day_flags(date_str):
    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                select
                    flag_date,
                    coalesce(need_setup_master_table, false) as need_setup_master_table,
                    coalesce(need_remove_master_table, false) as need_remove_master_table,
                    coalesce(extra_buddha_person, '') as extra_buddha_person,
                    coalesce(setup_people, '') as setup_people,
                    coalesce(remove_people, '') as remove_people,
                    coalesce(remarks, '') as remarks
                from schedule_day_flags
                where flag_date = %s
            """, (date_str,))
            row = cur.fetchone()

    if not row:
        return {
            "need_setup_master_table": False,
            "need_remove_master_table": False,

            "setup_people": "",
            "remove_people": "",

            "setup_person_1": "",
            "setup_person_2": "",

            "remove_person_1": "",
            "remove_person_2": "",
            "remove_extra_person": "",

            "extra_buddha_person": "",
            "remarks": "",
        }

    setup_list = [
        x.strip()
        for x in str(row.get("setup_people") or "").splitlines()
        if x.strip()
    ]

    remove_list = [
        x.strip()
        for x in str(row.get("remove_people") or "").splitlines()
        if x.strip()
    ]

    row["setup_person_1"] = setup_list[0] if len(setup_list) >= 1 else ""
    row["setup_person_2"] = setup_list[1] if len(setup_list) >= 2 else ""

    row["remove_person_1"] = remove_list[0] if len(remove_list) >= 1 else ""
    row["remove_person_2"] = remove_list[1] if len(remove_list) >= 2 else ""
    row["remove_extra_person"] = remove_list[2] if len(remove_list) >= 3 else ""

    return row