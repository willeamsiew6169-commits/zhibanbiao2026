# utils.py

from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

MY_TZ = ZoneInfo("Asia/Kuala_Lumpur")

TODAY_CODE_LIST = [
    "2580", "7312", "4901", "8625", "1047",
    "3698", "5206", "9174", "6842", "0359",
    "2468", "1357", "8080", "1122", "5566",
    "7788", "9090", "3145", "6721", "4826",
]

def get_today_code():
    today = datetime.now(MY_TZ)
    day_index = today.toordinal() % len(TODAY_CODE_LIST)
    return TODAY_CODE_LIST[day_index]


def now_date_str():
    return datetime.now(MY_TZ).strftime("%Y-%m-%d")

def parse_time(value):
    s = str(value or "").strip().lower().replace(" ", "")
    if not s:
        return None

    for fmt in ["%I:%M%p", "%I%p", "%H:%M"]:
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            pass

    return None

def calc_hours(start_time, end_time):
    st = parse_time(start_time)
    et = parse_time(end_time)

    if not st or not et:
        return 0.0

    diff = (et - st).total_seconds() / 3600

    if diff < 0:
        return 0.0

    return round(diff, 2)