# whatsapp_service.py

from datetime import datetime, time
from zoneinfo import ZoneInfo

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

    # 固定佛台人员不一定有 volunteer_schedule_assignments 记录。
    # 使用与负责人「佛台请假／换人」页面相同的有效名单来源：
    # schedule_buddha_overrides 优先，否则读取每周固定轮值。
    # 放在函数内导入，避免 schedule_web 与 whatsapp_service 循环导入。
    from schedule.schedule_web import get_fixed_buddha_for_date

    for name in get_fixed_buddha_for_date(date_str):
        name = str(name).strip()
        if name and name not in arranged["整理佛台"]:
            arranged["整理佛台"].append(name)

    # 初一额外第3位整理佛台义工：直接读取已保存设置，
    # 即使正式排班尚未同步，也要显示在 WhatsApp 正式值班表。
    extra_buddha_person = str(day_flags.get("extra_buddha_person") or "").strip()
    if (special_day_info.get("template_type") == "lunar_1_15"
            and "初一" in str(special_day_info.get("lunar_text") or "")
            and extra_buddha_person
            and extra_buddha_person not in arranged["整理佛台"]):
        arranged["整理佛台"].append(extra_buddha_person)

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
            if name not in arranged["整理佛台"]:
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

    # 卫生名单仅在显示时补齐；不改动数据库排班及 WhatsApp 自动发送。
    cleaning_places = ("佛堂卫生", "二楼卫生", "楼梯卫生")
    cleaners = list(dict.fromkeys(
        name for place in cleaning_places for name in arranged[place]
    ))
    if len(cleaners) == 1:
        malaysia_now = datetime.now(ZoneInfo("Asia/Kuala_Lumpur"))
        # 目标日当天上午10点起仍只有一人，才显示由同一人负责三处。
        if date_obj < malaysia_now.date() or (
            date_obj == malaysia_now.date() and malaysia_now.time() >= time(10, 0)
        ):
            for place in cleaning_places:
                if not arranged[place]:
                    arranged[place] = cleaners[:]
    elif len(cleaners) == 2 and not arranged["楼梯卫生"]:
        # 两位卫生义工时，楼梯由两人共同负责。
        arranged["楼梯卫生"] = cleaners[:]

    # 仅合并正式值班表的显示时段；不修改报名或排班数据库记录。
    # 相同班别、地点及姓名，且时间连续或重叠时才合并。
    def _minute(value):
        if isinstance(value, time):
            return value.hour * 60 + value.minute
        value = str(value or "").strip().lower().replace(" ", "")
        for fmt in ("%I:%M%p", "%I%p", "%H:%M", "%H:%M:%S"):
            try:
                parsed = datetime.strptime(value, fmt)
                return parsed.hour * 60 + parsed.minute
            except ValueError:
                pass
        return None

    for duty_key in ("绿观音堂", "绿活动中心", "橙观音堂", "橙活动中心",
                     "黄观音堂", "黄活动中心"):
        original = arranged[duty_key]
        by_name = {}
        for index, (name, start, end) in enumerate(original):
            begin, finish = _minute(start), _minute(end)
            if begin is None or finish is None or finish <= begin:
                continue
            by_name.setdefault(name, []).append((begin, finish, index))

        # 保留原始顺序和无法解析的时段，避免影响现有通知格式。
        merged_at = {}
        hidden = set()
        for name, spans in by_name.items():
            spans.sort(key=lambda item: (item[0], item[1]))
            groups = []
            for begin, finish, index in spans:
                if groups and begin <= groups[-1][1]:
                    groups[-1][1] = max(groups[-1][1], finish)
                    groups[-1][2].append(index)
                else:
                    groups.append([begin, finish, [index]])
            for begin, finish, indices in groups:
                if len(indices) < 2:
                    continue
                first = min(indices)
                # 沿用第一段的时间文字，结束时间沿用最晚结束的一段。
                first_start = min(indices, key=lambda i: _minute(original[i][1]))
                last_end = max(indices, key=lambda i: _minute(original[i][2]))
                merged_at[first] = (name, original[first_start][1], original[last_end][2])
                hidden.update(i for i in indices if i != first)

        arranged[duty_key] = [
            merged_at.get(i, item)
            for i, item in enumerate(original) if i not in hidden
        ]

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

