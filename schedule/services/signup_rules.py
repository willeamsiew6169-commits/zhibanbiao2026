# signup_rules.py

from datetime import timedelta
from schedule.builders.time_utils import malaysia_now
from lunar_rules  import get_special_day_info


# ======================================================
# 基本开关
# ======================================================

SUPPLY_OPEN_DAYS = 5          # 供台提前几天开放
SUPPLY_REMOVE_CLOSE_HOUR = 12 # 收供台当天几点关闭


# ======================================================
# 判断是否显示供台报名
# ======================================================

def is_supply_signup_open(target_date):

    today = malaysia_now().date()

    info = get_special_day_info(target_date)

    if not info.get("is_special"):
        return False

    open_date = target_date - timedelta(days=SUPPLY_OPEN_DAYS)

    return open_date <= today <= target_date


# ======================================================
# 是否还能报名设供台
# ======================================================

def can_signup_supply_setup(target_date):

    return is_supply_signup_open(target_date)


# ======================================================
# 是否还能报名收供台
# ======================================================

def can_signup_supply_remove(target_date):

    now = malaysia_now()

    if not is_supply_signup_open(target_date):
        return False

    if now.date() == target_date and now.hour >= SUPPLY_REMOVE_CLOSE_HOUR:
        return False

    return True