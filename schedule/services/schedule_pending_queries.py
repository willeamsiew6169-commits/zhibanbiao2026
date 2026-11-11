"""原 schedule_web.py 中的待安排报名及卫生轮替查询；仅原样迁移。"""

from db import get_conn
from psycopg2.extras import RealDictCursor


def load_pending_signups_for_date(date_str):
    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                select
                    id,
                    volunteer_id,
                    name,
                    role,
                    start_time,
                    end_time
                from volunteer_schedule_signups
                where signup_date = %s
                and coalesce(status, 'pending') = 'pending'
                order by role, start_time, name
            """, (date_str,))

            return cur.fetchall()


def get_recent_cleaning_places(volunteer_id, before_date, limit=5):
    if not volunteer_id:
        return []

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                select assigned_place
                from volunteer_schedule_assignments
                where volunteer_id = %s
                and assignment_date < %s
                and role = '卫生'
                and assigned_place in ('佛堂卫生', '二楼卫生', '楼梯卫生')
                and coalesce(status, 'assigned') <> 'cancelled'
                order by assignment_date desc, id desc
                limit %s
            """, (volunteer_id, before_date, limit))

            rows = cur.fetchall()

    return [r["assigned_place"] for r in rows if r.get("assigned_place")]


