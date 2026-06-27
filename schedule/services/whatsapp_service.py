# whatsapp_service.py

from datetime import datetime

from lunar_rules import (
    get_special_day_info,
    get_next_day_remove_info,
)

from schedule.services.assignment_service import load_assigned_places_for_date
from schedule.services.supply_service import load_day_flags

from schedule.builders.schedule_builder import (
    build_normal_message,
    build_lunar_1_15_message,
    build_buddhist_festival_message,
)


def build_whatsapp_from_assigned(date_str):
    rows = load_assigned_places_for_date(date_str)

    if not rows:
        return f"📅 {date_str}\n\n暂时没有已安排的值班资料。"

    date_obj = datetime.strptime(date_str, "%Y-%m-%d").date()

    special_day_info = get_special_day_info(date_obj)
    remove_info = get_next_day_remove_info(date_obj)

    day_flags = load_day_flags(date_str)

    setup_people = [
        x.strip()
        for x in str(day_flags.get("setup_people") or "").splitlines()
        if x.strip()
    ]

    remove_people = [
        x.strip()
        for x in str(day_flags.get("remove_people") or "").splitlines()
        if x.strip()
    ]

    arranged = {
        "整理佛台": [],
        "佛堂卫生": [],
        "二楼卫生": [],
        "楼梯卫生": [],
        "设师父供台": setup_people,
        "收师父供台": remove_people,
        "绿观音堂": [],
        "绿活动中心": [],
        "橙观音堂": [],
        "橙活动中心": [],
        "黄观音堂": [],
        "黄活动中心": [],
    }

    template_type = special_day_info["template_type"]
    is_lunar_day = template_type == "lunar_1_15"

    for r in rows:
        name = r.get("name")
        role = r.get("role")
        place = r.get("assigned_place")
        shift = r.get("shift_label")
        start = r.get("start_time")
        end = r.get("end_time")

        if not name:
            continue

        if role == "整理佛台" or place == "整理佛台":
            arranged["整理佛台"].append(name)

        elif role == "卫生" and place in ["佛堂卫生", "二楼卫生", "楼梯卫生"]:
            arranged[place].append(name)

        elif place == "设师父供台":
            if name not in arranged["设师父供台"]:
                arranged["设师父供台"].append(name)

        elif place == "收师父供台":
            if name not in arranged["收师父供台"]:
                arranged["收师父供台"].append(name)

        elif role == "值班":
            shift_key = str(shift or "").replace("班", "")

            if shift_key == "绿" and not is_lunar_day:
                shift_key = "黄"

            if shift_key in ["绿", "橙", "黄"] and place in ["观音堂", "活动中心"]:
                key = f"{shift_key}{place}"
                arranged[key].append((name, start, end))

    if template_type == "normal":
        return build_normal_message(
            date_obj,
            arranged,
            special_day_info,
            remove_info
        )

    if template_type == "lunar_1_15":
        return build_lunar_1_15_message(
            date_obj,
            arranged,
            special_day_info,
            remove_info
        )

    return build_buddhist_festival_message(
        date_obj,
        arranged,
        special_day_info,
        remove_info
    )