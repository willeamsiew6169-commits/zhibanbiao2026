# settings_service.py

import os
import psycopg2

from db import get_db
from psycopg2.extras import RealDictCursor
from psycopg2.extras import RealDictCursor

DATABASE_URL = os.environ.get("DATABASE_URL")

def get_conn():
    return psycopg2.connect(DATABASE_URL)

def get_schedule_setting(key, default_value=""):
    with get_db() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                select value
                from schedule_settings
                where key = %s
            """, (key,))
            row = cur.fetchone()

    if row:
        return row["value"]

    return default_value


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

    with get_db() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                select key, value
                from schedule_settings
            """)
            rows = cur.fetchall()

    for r in rows:
        settings[r["key"]] = r["value"]

    return settings


def save_schedule_setting(key, value):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                insert into schedule_settings (key, value, updated_at)
                values (%s, %s, now())
                on conflict (key)
                do update set
                    value = excluded.value,
                    updated_at = now()
            """, (key, value))
        conn.commit()