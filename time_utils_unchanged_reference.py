# time_utils.py

import re

from datetime import datetime, timedelta, timezone
MY_TZ = timezone(timedelta(hours=8))

def parse_min(t):
        if not t:
            return None

        s = str(t).strip().lower().replace(" ", "")

        if re.match(r"^\d{1,2}(am|pm)$", s):
            s = s.replace("am", ":00am").replace("pm", ":00pm")

        for fmt in ("%H:%M", "%H:%M:%S", "%I:%M%p"):
            try:
                dt = datetime.strptime(s, fmt)
                return dt.hour * 60 + dt.minute
            except:
                pass

        return None

def min_to_ampm(m):
        hour = m // 60
        minute = m % 60
        suffix = "am" if hour < 12 else "pm"
        hour12 = hour % 12 or 12
        return f"{hour12}:{minute:02d}{suffix}"

def normalize_time_text(text):
    text = str(text).strip()

    text = text.replace("～", "~").replace("—", "-").replace("–", "-").replace("－", "-")

    # 1030am -> 10:30am
    text = fix_time_format(text)

    # 黄丽玲 10 - 2 pm -> 黄丽玲 10am~2pm
    def repl_range(m):
        h1 = int(m.group(1))
        h2 = int(m.group(2))
        ap2 = m.group(3).lower()

        if ap2 == "pm" and h1 >= 7 and h2 <= 6:
            ap1 = "am"
        else:
            ap1 = ap2

        return f"{h1}{ap1}~{h2}{ap2}"

    text = re.sub(
        r"(?<!\d)(\d{1,2})\s*[-~]\s*(\d{1,2})\s*(am|pm)(?![a-zA-Z])",
        repl_range,
        text,
        flags=re.IGNORECASE
    )

    return text

def choose_split_time(start_min, end_min):
    """
    长班拆段：

    优先交班时间：
        2:00pm
        3:00pm
        1:00pm
        4:00pm
        12:00pm

    规则：
    1. 交班时间必须介于开始和结束之间。
    2. 前后两段都至少 2 小时。
    3. 如果没有符合条件，再取最接近中点的整点。
    """

    MIN_SEGMENT = 2 * 60  # 每段至少2小时

    candidates = [
        14 * 60,   # 2:00pm
        15 * 60,   # 3:00pm
        13 * 60,   # 1:00pm
        16 * 60,   # 4:00pm
        12 * 60,   # 12:00pm
    ]

    for split in candidates:

        if not (start_min < split < end_min):
            continue

        first = split - start_min
        second = end_min - split

        if first >= MIN_SEGMENT and second >= MIN_SEGMENT:
            return split

    # ---------- fallback ----------
    mid = (start_min + end_min) // 2

    rounded = round(mid / 60) * 60

    if (
        start_min < rounded < end_min
        and rounded - start_min >= MIN_SEGMENT
        and end_min - rounded >= MIN_SEGMENT
    ):
        return rounded

    return mid

def fix_time_format(text):
    text = str(text)

    def repl(m):
        num = m.group(1)
        ap = m.group(2).lower()

        if len(num) == 4:
            hour = int(num[:2])
            minute = int(num[2:])
        elif len(num) == 3:
            hour = int(num[:1])
            minute = int(num[1:])
        else:
            return m.group(0)

        if hour < 1 or hour > 12 or minute > 59:
            return m.group(0)

        return f"{hour}:{minute:02d}{ap}"

    return re.sub(
        r"(?<!\d)(\d{3,4})\s*(am|pm)(?![a-zA-Z])",
        repl,
        text,
        flags=re.IGNORECASE
    )

def time_to_minutes(t):
    return parse_min(t)

def malaysia_now():
    return datetime.now(MY_TZ)

def malaysia_today():
    return malaysia_now().date()