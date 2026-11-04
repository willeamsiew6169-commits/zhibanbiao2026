from functools import lru_cache
from datetime import datetime, time, timedelta

from db import get_conn
from psycopg2.extras import RealDictCursor
from schedule.builders.time_utils import malaysia_now
from lunar_rules import solar_to_lunar


DEFAULT_SCHEDULE_SETTINGS = {
    "default_day_switch_time": "18:00",
    "target_orange_guanyintang": "2",
    "target_orange_activity": "2",
    "target_yellow_guanyintang": "2",
    "target_yellow_activity": "2",
    "target_cleaning": "3",
    "supply_alert_days": "7",
    "wa_auto_enabled": "false",
    "wa_test_mode": "true",
    "wa_notice_1900": "false",
    "wa_roster_2200": "false",
    "wa_late_signup": "false",
    "wa_auto_publish_2200": "false",
    # 手动临时开放仅持续到当前活动窗口结束；不永久影响下一次佛日。
    "meal_manual_event_date": "",
    "meal_manual_state": "",
}


@lru_cache(maxsize=1)
def get_schedule_settings():
    """一次读取设置并缓存；写入时清除缓存。"""
    settings = DEFAULT_SCHEDULE_SETTINGS.copy()
    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("select key, value from schedule_settings")
            rows = cur.fetchall()
    for row in rows:
        settings[row["key"]] = row["value"]
    return settings


def get_schedule_setting(key, default=""):
    return get_schedule_settings().get(key, default)


def save_schedule_setting(key, value, updated_by="admin"):
    set_schedule_setting(key, value, updated_by=updated_by)


def set_schedule_setting(key, value, updated_by="admin"):
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                insert into schedule_settings (key, value, updated_at, updated_by)
                values (%s, %s, %s, %s)
                on conflict (key)
                do update set value = excluded.value,
                              updated_at = excluded.updated_at,
                              updated_by = excluded.updated_by
            """, (key, str(value), malaysia_now(), updated_by))
        conn.commit()
    get_schedule_settings.cache_clear()


def is_meal_festival_day(day):
    """仅初一十五、观音三大日、农历十月初六；闰月同样按农历月日判断。"""
    lunar = solar_to_lunar(day)
    return lunar.day in (1, 15) or (lunar.month, lunar.day) in {
        (2, 19), (6, 19), (9, 19), (10, 6)
    }


def current_meal_event_date(now=None):
    """返回当前有效的活动日期；窗口为前一天00:00至活动当天14:00。"""
    now = now or malaysia_now()
    today = now.date()
    if is_meal_festival_day(today) and now.time() < time(14, 0):
        return today
    tomorrow = today + timedelta(days=1)
    if is_meal_festival_day(tomorrow):
        return tomorrow
    return None


def is_meal_signup_open(now=None):
    """自动窗口 + 仅本次活动有效的负责人覆盖。"""
    event_date = current_meal_event_date(now)
    if event_date is None:
        return False
    settings = get_schedule_settings()
    if settings.get("meal_manual_event_date") == event_date.isoformat():
        state = settings.get("meal_manual_state", "").strip().lower()
        if state in ("true", "false"):
            return state == "true"
    return True


def set_meal_signup_override(enabled, updated_by="admin"):
    """仅对当前活动窗口生效。非窗口期间不创建未来活动覆盖。"""
    event_date = current_meal_event_date()
    if event_date is None:
        return False
    with get_conn() as conn:
        with conn.cursor() as cur:
            for key, value in (
                ("meal_manual_event_date", event_date.isoformat()),
                ("meal_manual_state", "true" if enabled else "false"),
            ):
                cur.execute("""
                    insert into schedule_settings (key, value, updated_at, updated_by)
                    values (%s, %s, %s, %s)
                    on conflict (key)
                    do update set value = excluded.value,
                                  updated_at = excluded.updated_at,
                                  updated_by = excluded.updated_by
                """, (key, value, malaysia_now(), updated_by))
        conn.commit()
    get_schedule_settings.cache_clear()
    return True


def is_schedule_setting_on(key):
    if key == "meal_signup_open":
        return is_meal_signup_open()
    value = get_schedule_setting(key, "false")
    return str(value).strip().lower() in {"1", "true", "yes", "on"}
