from psycopg2.extras import RealDictCursor

from db import get_conn


def get_schedule_publish_state(target_date):
    """
    一次读取当天发布状态与重新发布状态。
    """
    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                select
                    coalesce(is_published, false) as is_published,
                    coalesce(need_republish, false) as need_republish,
                    last_schedule_change
                from schedule_day_flags
                where flag_date = %s
                limit 1
            """, (target_date,))

            row = cur.fetchone()

    if not row:
        return {
            "is_published": False,
            "need_republish": False,
            "last_schedule_change": None,
        }

    return row


def is_schedule_published(target_date):
    state = get_schedule_publish_state(target_date)
    return bool(state["is_published"])


def get_schedule_republish_info(target_date):
    state = get_schedule_publish_state(target_date)

    return {
        "need_republish": bool(
            state["need_republish"]
        ),
        "last_schedule_change": (
            state["last_schedule_change"]
        ),
    }


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
                set need_republish = false
                where flag_date = %s
            """, (target_date,))

        conn.commit()


def publish_schedule_for_date(
    target_date,
    published_by="Admin",
):
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
                    published_at = coalesce(schedule_day_flags.published_at, excluded.published_at),
                    published_by = excluded.published_by,
                    need_republish = false
            """, (
                target_date,
                published_by,
            ))

        conn.commit()

    # 手动发布不受22:00窗口限制；已发布的排班也能接续建立基准与入队。
    # 失败只记录，不回滚已成功的负责人发布；可安全再次按发布重试。
    import logging
    try:
        from schedule.services.whatsapp_publish_2200 import finalize_published_roster
        finalize_published_roster(str(target_date))
    except Exception:
        logging.getLogger(__name__).exception(
            '负责人发布成功，但 WhatsApp 快照/正式表入队失败：%s', target_date)


def unpublish_schedule_for_date(target_date):
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                update schedule_day_flags
                set is_published = false
                where flag_date = %s
            """, (target_date,))

        conn.commit()
