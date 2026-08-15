from functools import lru_cache

from db import get_conn
from psycopg2.extras import RealDictCursor
from schedule.builders.time_utils import malaysia_now


DEFAULT_SCHEDULE_SETTINGS = {
    "default_day_switch_time": "18:00",
    "target_orange_guanyintang": "2",
    "target_orange_activity": "2",
    "target_yellow_guanyintang": "2",
    "target_yellow_activity": "2",
    "target_cleaning": "3",
    "supply_alert_days": "7",
}


@lru_cache(maxsize=1)
def get_schedule_settings():
    """
    一次读取全部排班设置并缓存。

    set/save 后会自动清除缓存，
    所以下一次读取会取得最新数据库资料。
    """
    settings = DEFAULT_SCHEDULE_SETTINGS.copy()

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                select key, value
                from schedule_settings
            """)
            rows = cur.fetchall()

    for row in rows:
        settings[row["key"]] = row["value"]

    return settings


def get_schedule_setting(key, default=""):
    return get_schedule_settings().get(key, default)


def save_schedule_setting(key, value, updated_by="admin"):
    set_schedule_setting(
        key,
        value,
        updated_by=updated_by,
    )


def set_schedule_setting(key, value, updated_by="admin"):
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                insert into schedule_settings
                (
                    key,
                    value,
                    updated_at,
                    updated_by
                )
                values
                (
                    %s,
                    %s,
                    %s,
                    %s
                )
                on conflict (key)
                do update set
                    value = excluded.value,
                    updated_at = excluded.updated_at,
                    updated_by = excluded.updated_by
            """, (
                key,
                str(value),
                malaysia_now(),
                updated_by,
            ))

        conn.commit()

    get_schedule_settings.cache_clear()


def is_schedule_setting_on(key):
    value = get_schedule_setting(key, "false")

    return str(value).strip().lower() in {
        "1",
        "true",
        "yes",
        "on",
    }
