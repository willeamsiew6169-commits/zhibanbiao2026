# settings_service.py

from functools import lru_cache

from db import get_conn
from psycopg2.extras import RealDictCursor


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