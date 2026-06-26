# publish_service.py

from db import get_conn
from psycopg2.extras import RealDictCursor

def is_schedule_published(target_date):
    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                select coalesce(is_published, false) as is_published
                from schedule_day_flags
                where flag_date = %s
                limit 1
            """, (target_date,))
            row = cur.fetchone()

    return bool(row and row["is_published"])

def publish_schedule_for_date(target_date, published_by="Admin"):

    with get_conn() as conn:
        with conn.cursor() as cur:

            cur.execute("""
            insert into schedule_day_flags
            (
                flag_date,
                is_published,
                published_at,
                published_by
            )
            values
            (
                %s,
                true,
                now(),
                %s
            )

            on conflict (flag_date)

            do update set

                is_published = true,
                published_at = now(),
                published_by = excluded.published_by
            """, (
                target_date,
                published_by
            ))

            conn.commit()


def unpublish_schedule_for_date(target_date):

    with get_conn() as conn:
        with conn.cursor() as cur:

            cur.execute("""

            update schedule_day_flags

            set

                is_published = false

            where flag_date=%s

            """,(target_date,))

            conn.commit()