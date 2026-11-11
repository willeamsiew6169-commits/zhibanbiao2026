# 负责人报名管理只读查询：从 schedule_web.py 原样迁移。
from datetime import timedelta
from psycopg2.extras import RealDictCursor
from db import get_conn
from schedule.builders.time_utils import malaysia_today


def load_management_signup_view(status, date_filter, role_filter, show_history):
    sql = """
        select
            id,
            volunteer_id,
            name,
            signup_date,
            role,
            start_time,
            end_time,
            coalesce(status, 'pending') as status,
            assigned_place,
            remarks
        from volunteer_schedule_signups
        where 1=1
    """

    if not show_history:
        sql += """
            and signup_date >= current_date
            and coalesce(status,'pending') <> 'cancelled'
        """

    params = []

    if status != "all":
        sql += " and coalesce(status, 'pending') = %s"
        params.append(status)

    if date_filter:
        sql += " and signup_date = %s"
        params.append(date_filter)

    if role_filter != "all":
        sql += " and role = %s"
        params.append(role_filter)

    sql += " order by signup_date, role, start_time, name"

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute(sql, params)
            rows = cur.fetchall()

            signup_ids = [r["id"] for r in rows]
            assignment_map = {}

            if signup_ids:
                cur.execute("""
                    select
                        signup_id,
                        shift_label,
                        assigned_place,
                        start_time,
                        end_time
                    from volunteer_schedule_assignments
                    where signup_id = any(%s)
                    and coalesce(status, 'assigned') <> 'cancelled'
                    order by start_time, shift_label, assigned_place
                """, (signup_ids,))

                assignment_rows = cur.fetchall()

                for a in assignment_rows:
                    sid = a["signup_id"]

                    shift = a.get("shift_label") or ""
                    place = a.get("assigned_place") or ""
                    start = a.get("start_time") or ""
                    end = a.get("end_time") or ""

                    text = ""

                    if shift:
                        text += shift

                    if place:
                        if text:
                            text += " "
                        text += place

                    if start and end:
                        text += f"<br><small>{start} ~ {end}</small>"

                    assignment_map.setdefault(sid, []).append(text)

            for r in rows:
                r["final_assignment"] = assignment_map.get(r["id"], [])

    summary = {
        "total": len(rows),
        "pending": sum(1 for r in rows if r["status"] == "pending"),
        "assigned": sum(1 for r in rows if r["status"] == "assigned"),
        "cancelled": sum(1 for r in rows if r["status"] == "cancelled"),
    }

    today_date = malaysia_today()

    today = today_date.strftime("%Y-%m-%d")
    tomorrow = (today_date + timedelta(days=1)).strftime("%Y-%m-%d")

    role_summary = {}

    for r in rows:
        role = r["role"] or "未填写"
        role_summary[role] = role_summary.get(role, 0) + 1

    return rows, summary, role_summary, today, tomorrow
