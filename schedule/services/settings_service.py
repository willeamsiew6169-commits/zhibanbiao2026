# settings_service.py

from functools import lru_cache

from db import get_conn
from psycopg2.extras import RealDictCursor
from schedule.builders.time_utils import malaysia_now


@lru_cache(maxsize=1)
def get_schedule_settings():
    defaults = {
        "default_day_switch_time": "18:00",
        "target_orange_guanyintang": "2",
        "target_orange_activity": "2",
        "target_yellow_guanyintang": "2",
        "target_yellow_activity": "2",
        "target_cleaning": "3",
        "supply_alert_days": "7",
    }

    settings = defaults.copy()

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                select key, value
                from schedule_settings
            """)
            rows = cur.fetchall()

    for r in rows:
        settings[r["key"]] = r["value"]

    return settings


def get_schedule_setting(key, default_value=""):
    settings = get_schedule_settings()
    return settings.get(key, default_value)


def save_schedule_setting(key, value):
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                insert into schedule_settings (key, value, updated_at)
                values (%s, %s, now())
                on conflict (key)
                do update set
                    value = excluded.value,
                    updated_at = now()
            """, (key, value))

    # 清除缓存，下次重新读取数据库
    get_schedule_settings.cache_clear()

def get_schedule_setting(key, default="false"):
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                select value
                from schedule_settings
                where key = %s
            """, (key,))
            row = cur.fetchone()

    if not row:
        return default

    return row["value"]

def set_schedule_setting(key, value, updated_by="admin"):
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                insert into schedule_settings (key, value, updated_at, updated_by)
                values (%s, %s, %s, %s)
                on conflict (key)
                do update set
                    value = excluded.value,
                    updated_at = excluded.updated_at,
                    updated_by = excluded.updated_by
            """, (
                key,
                value,
                malaysia_now(),
                updated_by
            ))
        conn.commit()


def is_schedule_setting_on(key):
    return get_schedule_setting(key) == "true"