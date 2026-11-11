# 从 schedule_web.py 原样迁移的纯文本处理函数；不涉及数据库、路由或排班逻辑。
from schedule.helpers import to_simple

def display_report_name(filename):
    name = filename.replace(".xlsx", "")

    if name == "2026_year_report":
        return "2026 年报"

    if "_month_report" in name:
        parts = name.split("_")
        year = parts[0]
        month = parts[1]
        return f"{year}-{month} 月报"

    return filename


def search_volunteer_names(keyword, volunteers):
    keyword_simple = to_simple(keyword)

    if not keyword_simple:
        return []

    results = []

    for v in volunteers:
        name = str(v.get("name") or v.get("姓名") or "").strip()
        phone = str(v.get("phone") or v.get("电话") or v.get("电话号码") or "").strip()

        name_simple = to_simple(name)

        if keyword_simple in name_simple or keyword_simple in phone:
            results.append(v)

    return results


def get_default_time_by_role(role, start_time, end_time):

    role = str(role).strip()

    if role == "值班":
        return start_time, end_time

    if role in ["卫生", "供台", "整理佛台"]:
        return None, None

    return start_time, end_time


def normalize_whatsapp_keyword(text):
    text = str(text or "").strip().upper().replace(" ", "")

    if text.startswith("STW-"):
        return text

    if text.startswith("STW") and text[3:].isdigit():
        return f"STW-{int(text[3:])}"

    if text.startswith("CHE-"):
        return text

    if text.startswith("CHE") and text[3:].isdigit():
        return f"CHE-{int(text[3:])}"

    return text

