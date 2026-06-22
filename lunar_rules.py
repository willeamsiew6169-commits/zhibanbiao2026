# lunar_rules.py

from datetime import timedelta

try:
    from lunardate import LunarDate
except ImportError:
    LunarDate = None


LUNAR_MONTH_NAMES = {
    1: "正月", 2: "二月", 3: "三月", 4: "四月", 5: "五月", 6: "六月",
    7: "七月", 8: "八月", 9: "九月", 10: "十月", 11: "十一月", 12: "十二月"
}

LUNAR_DAY_NAMES = {
    1: "初一", 2: "初二", 3: "初三", 4: "初四", 5: "初五",
    6: "初六", 7: "初七", 8: "初八", 9: "初九", 10: "初十",
    11: "十一", 12: "十二", 13: "十三", 14: "十四", 15: "十五",
    16: "十六", 17: "十七", 18: "十八", 19: "十九", 20: "二十",
    21: "廿一", 22: "廿二", 23: "廿三", 24: "廿四", 25: "廿五",
    26: "廿六", 27: "廿七", 28: "廿八", 29: "廿九", 30: "三十"
}

SPECIAL_DAYS = {
    ("any", 1): {
        "name": "农历初一",
        "template_type": "lunar_1_15",
        "start_time": "06:00",
        "setup_shifu": True,
        "remove_next_day": False,
    },
    ("any", 15): {
        "name": "农历十五",
        "template_type": "lunar_1_15",
        "start_time": "06:00",
        "setup_shifu": True,
        "remove_next_day": False,
    },

    (2, 8): {
        "name": "释迦摩尼佛出家日",
        "template_type": "buddhist_festival",
        "start_time": "06:00",
        "setup_shifu": True,
        "remove_next_day": True,
    },
    (2, 15): {
        "name": "释迦摩尼佛涅槃日",
        "template_type": "buddhist_festival",
        "start_time": "06:00",
        "setup_shifu": True,
        "remove_next_day": True,
    },
    (2, 19): {
        "name": "观世音菩萨诞辰日",
        "template_type": "buddhist_festival",
        "start_time": "06:00",
        "setup_shifu": True,
        "remove_next_day": True,
    },
    (4, 8): {
        "name": "释迦摩尼佛诞辰日",
        "template_type": "buddhist_festival",
        "start_time": "06:00",
        "setup_shifu": True,
        "remove_next_day": True,
    },
    (6, 19): {
        "name": "观世音菩萨成道日",
        "template_type": "buddhist_festival",
        "start_time": "06:00",
        "setup_shifu": True,
        "remove_next_day": True,
    },
    (9, 19): {
        "name": "观世音菩萨出家日",
        "template_type": "buddhist_festival",
        "start_time": "06:00",
        "setup_shifu": True,
        "remove_next_day": True,
    },
    (10, 6): {
        "name": "恩师卢军宏涅槃日",
        "template_type": "buddhist_festival",
        "start_time": "06:00",
        "setup_shifu": True,
        "remove_next_day": True,
    },
    (12, 9): {
        "name": "释迦摩尼佛成道日",
        "template_type": "buddhist_festival",
        "start_time": "06:00",
        "setup_shifu": True,
        "remove_next_day": True,
    },
}


def solar_to_lunar(date_obj):
    if LunarDate is None:
        raise ImportError("请先安装 lunardate：python -m pip install lunardate")
    return LunarDate.fromSolarDate(date_obj.year, date_obj.month, date_obj.day)


def get_lunar_text(lunar_month, lunar_day):
    month_text = LUNAR_MONTH_NAMES.get(lunar_month, f"{lunar_month}月")
    day_text = LUNAR_DAY_NAMES.get(lunar_day, str(lunar_day))
    return f"农历{month_text}{day_text}"


def get_special_day_info(date_obj):
    lunar = solar_to_lunar(date_obj)

    result = {
        "solar_date": date_obj.strftime("%Y-%m-%d"),
        "lunar_month": lunar.month,
        "lunar_day": lunar.day,
        "lunar_text": get_lunar_text(lunar.month, lunar.day),
        "is_special": False,
        "special_names": [],
        "template_type": "normal",
        "start_time": None,
        "setup_shifu": False,
        "remove_next_day": False,
    }

    result["template_text"] = {
        "normal": "平时值班模板",
        "lunar_1_15": "初一十五值班模板",
        "buddhist_festival": "佛诞大日子模板",
    }.get(
        result["template_type"],
        result["template_type"]
    )

    key_exact = (lunar.month, lunar.day)
    key_any = ("any", lunar.day)

    info = None
    if key_exact in SPECIAL_DAYS:
        info = SPECIAL_DAYS[key_exact]
    elif key_any in SPECIAL_DAYS:
        info = SPECIAL_DAYS[key_any]

    if info:
        result["is_special"] = True
        result["special_names"] = [info["name"]]
        result["template_type"] = info["template_type"]
        result["start_time"] = info["start_time"]
        result["setup_shifu"] = info["setup_shifu"]
        result["remove_next_day"] = info["remove_next_day"]

    return result


def get_next_day_remove_info(date_obj):
    yesterday = date_obj - timedelta(days=1)
    info = get_special_day_info(yesterday)

    return {
        "need_remove_today_after_12": info["remove_next_day"],
        "remove_confirmed": False,
    }