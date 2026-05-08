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