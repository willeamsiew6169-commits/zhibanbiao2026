from datetime import date, timedelta
from lunardate import LunarDate
import calendar

YEAR = 2026
MONTH = 11

# 上一次实际做供花供果的日期
PREV_WORK_DAY = date(2026, 10, 27)

TARGET_GAP = 7
MAX_GAP = 8

MANUAL_ADD = [
    # 例：
    # (23, "💧", "佛台", "")
]

MANUAL_REMOVE = [
    # 例：
    # 11
]

MANUAL_OVERRIDE = {
    # 例：
    # 16: ("👉", "佛台+供台", "")
}

WEEKDAYS = [
    "星期一",
    "星期二",
    "星期三",
    "星期四",
    "星期五",
    "星期六",
    "星期日",
]

SPECIAL_DAYS = {
    ("any", 1): {
        "name": "农历初一",
    },

    ("any", 15): {
        "name": "农历十五",
    },

    (2, 8): {
        "name": "释迦摩尼佛出家日",
    },

    (2, 15): {
        "name": "释迦摩尼佛涅槃日",
    },

    (2, 19): {
        "name": "观世音菩萨诞辰日",
    },

    (4, 8): {
        "name": "释迦摩尼佛诞辰日",
    },

    (6, 19): {
        "name": "观世音菩萨成道日",
    },

    (7, 13): {
        "name": "大势至菩萨圣诞",
    },

    # 七月三十会另外用特殊逻辑处理
    (7, 30): {
        "name": "地藏王菩萨圣诞",
    },

    (8, 22): {
        "name": "燃灯古佛诞辰日",
    },

    (9, 19): {
        "name": "观世音菩萨出家日",
    },

    (10, 6): {
        "name": "恩师卢军宏涅槃日",
    },

    (11, 17): {
        "name": "阿弥陀佛圣诞日",
    },

    (12, 9): {
        "name": "释迦摩尼佛成道日",
    },
}


# =========================================================
# 农历辅助
# =========================================================

def get_lunar(d):
    return LunarDate.fromSolarDate(
        d.year,
        d.month,
        d.day
    )


def is_last_day_of_lunar_month(d):
    """
    判断某一天是不是该农历月最后一天。
    """
    lunar_today = get_lunar(d)
    lunar_tomorrow = get_lunar(d + timedelta(days=1))

    return (
        lunar_today.month != lunar_tomorrow.month
        or lunar_today.year != lunar_tomorrow.year
    )


def get_special_name(d):
    """
    返回某公历日期所对应的大日子名称。
    """

    lunar = get_lunar(d)

    lm = lunar.month
    ld = lunar.day

    # -----------------------------------------------------
    # 特殊处理：
    # 地藏王菩萨圣诞 = 农历七月最后一天
    #
    # 如果该年七月有30日：
    #   七月三十
    #
    # 如果该年七月只有29日：
    #   自动用七月二十九
    # -----------------------------------------------------

    if lm == 7 and is_last_day_of_lunar_month(d):
        return "地藏王菩萨圣诞"

    # -----------------------------------------------------
    # 固定大日子
    # -----------------------------------------------------

    if (lm, ld) in SPECIAL_DAYS:
        return SPECIAL_DAYS[(lm, ld)]["name"]

    # -----------------------------------------------------
    # 初一 / 十五
    # -----------------------------------------------------

    if ("any", ld) in SPECIAL_DAYS:
        return SPECIAL_DAYS[("any", ld)]["name"]

    return ""


# =========================================================
# 自动找大日子的“前一天”
# =========================================================

def auto_special_days(year, month):
    last_day = calendar.monthrange(year, month)[1]
    result = {}

    start_date = date(year, month, 1)
    end_date = date(year, month, last_day) + timedelta(days=1)

    d = start_date

    while d <= end_date:
        name = get_special_name(d)

        # 十一月十五已准备供花供果，十七继续供奉。
        lunar = get_lunar(d)
        if name and (lunar.month, lunar.day) != (11, 17):
            prep_date = d - timedelta(days=1)

            if prep_date.year == year and prep_date.month == month:
                result[prep_date.day] = (
                    "👉",
                    "佛台+供台",
                    f"{name}前夕"
                )

        d += timedelta(days=1)

    # ✅ 如果连续两天都是佛台+供台，只保留前一天
    cleaned = {}

    for day in sorted(result):
        if cleaned:
            prev_day = max(cleaned)

            if day - prev_day == 1:
                # 前一天已经完整做佛台+供台
                # 当前这一天不再重复
                continue

        cleaned[day] = result[day]

    return cleaned


# =========================================================
# 普通日尽量避开星期日
# =========================================================

def adjust_normal_day_if_sunday(
    year,
    month,
    day,
    fixed_days,
    result
):
    d = date(year, month, day)

    # 不是星期日
    if d.weekday() != 6:
        return day

    last_day = calendar.monthrange(
        year,
        month
    )[1]

    # -----------------------------------------------------
    # 星期日普通日：
    # 优先调到星期一
    # -----------------------------------------------------

    monday = day + 1

    if (
        monday <= last_day
        and monday not in fixed_days
        and monday not in result
    ):
        return monday

    # -----------------------------------------------------
    # 星期一不能用，
    # 才考虑星期六
    # -----------------------------------------------------

    saturday = day - 1

    if (
        saturday >= 1
        and saturday not in fixed_days
        and saturday not in result
    ):
        return saturday

    # 实在没办法才保留星期日
    return day


# =========================================================
# 自动补普通佛台日
# =========================================================

def auto_fill_normal_days(year, month, fixed_days, prev_work_day=None):
    last_day = calendar.monthrange(year, month)[1]
    first_day = date(year, month, 1)
    result = {}
    fixed_days = set(fixed_days)

    if prev_work_day:
        start_anchor = 1 - (first_day - prev_work_day).days
        if start_anchor >= 1:
            raise ValueError("上一次工作日必须早于本月")
        if 1 - start_anchor > MAX_GAP:
            deadline = prev_work_day + timedelta(days=MAX_GAP)
            raise ValueError(
                f"跨月间隔超过{MAX_GAP}天，请先在上月补安排，"
                f"最迟 {deadline:%Y-%m-%d}，保存后重新生成本月。"
            )
    else:
        start_anchor = 0

    def add_day(preferred, lower, upper):
        # 先限制在合法区间，再避开星期日；不得突破8天上限。
        day = max(lower, min(upper, preferred))
        adjusted = adjust_normal_day_if_sunday(
            year, month, day, fixed_days, result
        )
        if lower <= adjusted <= upper:
            day = adjusted
        elif date(year, month, day).weekday() == 6:
            if lower <= day - 1 <= upper:
                day -= 1
        result[day] = ("💧", "佛台", "")
        return day

    latest = start_anchor
    for fixed in sorted(fixed_days):
        while fixed - latest > MAX_GAP:
            lower = max(1, latest + 1)
            upper = min(fixed - 1, latest + MAX_GAP)
            preferred = (latest + fixed) // 2
            if fixed - latest > 2 * MAX_GAP:
                preferred = latest + TARGET_GAP
            latest = add_day(preferred, lower, upper)
        latest = fixed

    # 月末必须给下月第一天留下合法间隔；月末不是工作日。
    while last_day + 1 - latest > MAX_GAP:
        latest = add_day(
            latest + TARGET_GAP, max(1, latest + 1),
            min(last_day, latest + MAX_GAP)
        )
    return result


# =========================================================
# 生成 WhatsApp 文案
# =========================================================

def make_text(year, month, prev_work_day=None):

    # 先加入所有固定大日子的前一天
    days = auto_special_days(
        year,
        month
    )

    # 再自动补普通日
    normal_days = auto_fill_normal_days(
        year,
        month,
        set(days.keys()), prev_work_day
    )

    days.update(
        normal_days
    )

    # -----------------------------------------------------
    # 输出
    # -----------------------------------------------------

    lines = [
        f"{month}月供花供果",
        ""
    ]

    for day in sorted(days):

        icon, job, remark = days[day]

        d = date(
            year,
            month,
            day
        )

        weekday = WEEKDAYS[
            d.weekday()
        ]

        date_str = (
            f"{day}/"
            f"{month}/"
            f"{str(year)[-2:]}"
        )

        line = (
            f"{icon}"
            f"{date_str} "
            f"{weekday} "
            f"{job}"
        )

        if remark:

            line += (
                f" （{remark}）"
            )

        lines.append(line)
        lines.append("")

    return "\n".join(
        lines
    ).strip(), days


# =========================================================
# 新增：负责人确认后才同步 Supabase（原有排班/TXT 不变）
# =========================================================

def text_from_confirmed_days(year, month, days):
    lines = [f"{month}月供花供果", ""]
    for day, (icon, job, remark) in sorted(days.items()):
        d = date(year, month, day)
        line = f"{icon}{day}/{month}/{str(year)[-2:]} {WEEKDAYS[d.weekday()]} {job}"
        if remark:
            line += f" （{remark}）"
        lines.extend((line, ""))
    return "\n".join(lines).strip()




def cross_month_days(year, month, prev_work_day, today):
    """Only propose missing future dates in the immediately preceding month."""
    first = date(year, month, 1)
    previous_month = (first - timedelta(days=1)).replace(day=1)
    result = {}
    if prev_work_day is None:
        return result
    if prev_work_day >= first:
        raise ValueError("上一次工作日必须早于本月")
    while (first - prev_work_day).days > MAX_GAP:
        day = prev_work_day + timedelta(days=MAX_GAP)
        if day.weekday() == 6:
            day -= timedelta(days=1)
        if day < today or day < previous_month:
            raise ValueError("跨月补位涉及过去日期或更早月份，请先核对已保存安排")
        result[day] = ("💧", "佛台", "")
        prev_work_day = day
    return result


def combined_text(year, month, days, cross):
    lines = [f"{month}月供花供果", ""]
    for d, (icon, job, remark) in sorted(cross.items()):
        line = f"{icon}{d.day}/{d.month}/{str(d.year)[-2:]} {WEEKDAYS[d.weekday()]} {job}"
        if remark:
            line += f" （{remark}）"
        lines.extend((line, ""))
    lines.extend(text_from_confirmed_days(year, month, days).splitlines()[2:])
    return "\n".join(lines).strip()
