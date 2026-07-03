# publish_service.py

from psycopg2.extras import RealDictCursor
from db import get_conn


def is_schedule_published(target_date):
    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                select
                    coalesce(is_published, false) as is_published
                from schedule_day_flags
                where flag_date = %s
                limit 1
            """, (target_date,))

            row = cur.fetchone()

    return bool(row and row["is_published"])

def get_schedule_republish_info(target_date):

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:

            cur.execute("""
                select
                    coalesce(need_republish, false) as need_republish,
                    last_schedule_change
                from schedule_day_flags
                where flag_date = %s
                limit 1
            """, (target_date,))

            row = cur.fetchone()

    if not row:
        return {
            "need_republish": False,
            "last_schedule_change": None
        }

    return row


def mark_schedule_need_republish(target_date):

    with get_conn() as conn:
        with conn.cursor() as cur:

            cur.execute("""
                insert into schedule_day_flags
                (
                    flag_date,
                    need_republish,
                    last_schedule_change
                )
                values
                (
                    %s,
                    true,
                    now()
                )

                on conflict (flag_date)

                do update set

                    need_republish = true,
                    last_schedule_change = now()
            """, (target_date,))

            conn.commit()


def clear_schedule_need_republish(target_date):

    with get_conn() as conn:
        with conn.cursor() as cur:

            cur.execute("""
                update schedule_day_flags
                set
                    need_republish = false
                where flag_date = %s
            """, (target_date,))

            conn.commit()


def publish_schedule_for_date(target_date, published_by="Admin"):
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                insert into schedule_day_flags
                (
                    flag_date,
                    is_published,
                    published_at,
                    published_by,
                    need_republish
                )
                values
                (
                    %s,
                    true,
                    now(),
                    %s,
                    false
                )

                on conflict (flag_date)

                do update set

                    is_published = true,
                    published_at = now(),
                    published_by = excluded.published_by,
                    need_republish = false
            """, (
                target_date,
                published_by
            ))

            conn.commit()


def get_schedule_republish_info(target_date):

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:

            cur.execute("""
                select
                    coalesce(need_republish, false) as need_republish,
                    last_schedule_change
                from schedule_day_flags
                where flag_date = %s
                limit 1
            """, (target_date,))

            row = cur.fetchone()

    if not row:
        return {
            "need_republish": False,
            "last_schedule_change": None
        }

    return row


def unpublish_schedule_for_date(target_date):
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                update schedule_day_flags

                set
                    is_published = false

                where flag_date = %s
            """, (target_date,))