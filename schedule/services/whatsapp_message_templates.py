"""WhatsApp 唯一文案中心。只负责文字排版；不触碰数据库、排班、队列及发送。"""
import re

WHATSAPP_FOOTER = '- 观音堂早晚香 由值班义工带领上香\n\n- 观音堂续香 -\n佛友可以续香，但是要跟着请安词，必须燃烧完了一支香才续香。如有多位佛友要续香，请先让给先到达观音堂的佛友，轮流续香\n\n！！值班义工请注意：\n1）下雨天记得关上烧送小房子的窗口\n2）在离开观音堂之前，请确保把所有的窗口关上\n3）观音堂第一架的冷气坚决不能调/关，第二架和第三架需要轮流替换，不建议长时间只开其中一架\n4）活动中心的冷气也需要轮流替换，每一架冷气上都有标贴（ABCD)。请在特定时段只开其中两架\n\n例如：10am-2pm开冷气 A或者C，2pm-6pm开冷气 B或者D\n\n另外，请大家多留意义工群信息，以便大家能够团结一致的护持好观音堂。佛子齐心，普度众生。\n\n非常感恩大家！\n大家功德无量！\n🙏🙏🙏'

def _compact_whatsapp_message(message):
    """只压缩多余空行，不修改义工名单、岗位或注意事项。"""
    lines = [line.strip() for line in message.splitlines()]
    message = re.sub(r"\n{3,}", "\n\n", "\n".join(lines)).strip() + "\n"
    # WhatsApp 排版：同一时段的「活动中心」与「观音堂」之间空一行。
    message = re.sub(
        r"(?m)(?<!\n)\n(?=(?:🟢|🟠|🔴|🟡)?[ \t]*活动中心:)",
        "\n\n", message,
    )
    # 主要值班时段之间空两行；不改变姓名、个人时间和其他段落。
    message = re.sub(
        r"\n{2}(?=(?:8:00am~10:00am|🟢 8:00am~10:00am|10:00am~2:00pm|2:00pm~6:00pm)[ \t]*\n)",
        "\n\n\n", message,
    )
    return message


def get_display_name(p):
    if isinstance(p, dict):
        return p.get("name") or p.get("姓名") or ""
    if isinstance(p, (list, tuple)):
        return p[0] if p else ""
    return str(p or "")


def build_roster_message(date_obj, arranged, special_info, remove_info, kind):
    from schedule.builders.schedule_builder import get_weekday_name, format_people_inline, format_shift_block
    """负责人 2026-11-16 样式；只处理 WhatsApp 文字，不改变排班。"""
    date_text = f"{date_obj.day}/{date_obj.month}/{date_obj.year}"
    weekday = get_weekday_name(date_obj)
    lunar = str(special_info.get("lunar_text") or "").strip()
    if lunar and not lunar.startswith("（"):
        lunar = f"（{lunar}）"
    festival_names = special_info.get("special_names") or []
    festival = "\n".join(str(x) for x in festival_names if x)
    big_day = kind != "normal"
    start = "6:00am" if big_day else "8:00am"
    end = "8:00am" if big_day else "10:00am"
    greeting = "八点正请安" if big_day else "十点正请安"

    setup_people = format_people_inline(arranged.get("设师父供台", []))
    remove_people = format_people_inline(arranged.get("收师父供台", []))
    master = []
    if special_info.get("setup_shifu") or setup_people:
        master.extend(["6:00am~8:00am 或", "6:00am~完成供台工作", "设师父供台:", setup_people or "待安排"])
    if remove_info.get("need_remove_today_after_12") or remove_people:
        master.extend(["12:00pm~2:00pm", "收师父供台:", remove_people or "待安排", "12pm的香结束之后，请下供桌。"])
    master_section = "\n\n".join(["\n".join(master)]) if master else ""

    green_section = ""
    if arranged.get("绿观音堂") or arranged.get("绿活动中心") or big_day:
        green_section = f"""8:00am~10:00am
🟢 观音堂:
{format_shift_block(arranged.get('绿观音堂', []))}

🟢 活动中心:
{format_shift_block(arranged.get('绿活动中心', []))}"""

    sections = [
        "值班时间表", date_text, f"({weekday})",
    ]
    if lunar:
        sections.append(lunar.strip())
    if festival:
        sections.append(festival)
    sections.extend([
        "📝 义工报名请点击以下链接👇\n" + SIGNUP_URL,
        greeting,
        "［整理佛台］",
        f"{start}~{end}  或\n{start}~整理完毕",
        format_people_inline(arranged.get("整理佛台", [])),
    ])
    if master_section:
        sections.append(master_section)
    sections.extend([
        "［打扫卫生］",
        f"{start}~{end} 或\n{start}~打扫完毕",
        f"•佛堂卫生: {format_people_inline(arranged.get('佛堂卫生', []))}",
        f"•二楼卫生: {format_people_inline(arranged.get('二楼卫生', []))}",
        f"•楼梯卫生: {format_people_inline(arranged.get('楼梯卫生', []), sep='/')}",
        "每日卫生义工请注意：清理完卫生之后，请把卫生用具包括吸尘机放回原位（活动中心store里面的小房间）",
        "［值班］",
    ])
    if green_section:
        sections.append(green_section)
    orange_icon = "🔴" if kind == "festival" else "🟠"
    sections.extend([
        f"10:00am-2:00pm\n\n{orange_icon} 观音堂:\n{format_shift_block(arranged.get('橙观音堂', []))}\n\n{orange_icon} 活动中心:\n{format_shift_block(arranged.get('橙活动中心', []))}",
        f"2:00pm~6:00pm\n\n🟡 观音堂:\n{format_shift_block(arranged.get('黄观音堂', []))}\n\n🟡 活动中心:\n{format_shift_block(arranged.get('黄活动中心', []))}",
        WHATSAPP_FOOTER,
    ])
    # 日期行不插空行；岗位标题紧接时间，卫生地点连续显示。
    parts = []
    previous = ""
    for part in sections:
        if not part:
            continue
        tight = (
            previous == "值班时间表"
            or previous == date_text
            or previous == f"({weekday})"
            or previous in ("［整理佛台］", "［打扫卫生］", "［值班］")
            or part.startswith(("•二楼卫生:", "•楼梯卫生:"))
        )
        parts.append(("\n" if tight else "\n\n") + part if parts else part)
        previous = part
    return _compact_whatsapp_message("".join(parts))



SIGNUP_URL = 'https://gyt-checkin-sg.onrender.com/volunteer'

def build_morning_message(target, rows):
    from schedule.services.whatsapp_morning_summary import ROLES, _key, _time
    parts = [f'📋 {target} 新增义工安排', '']
    for role in ROLES:
        selected = [r for r in rows if r['role'] == role]
        if not selected:
            continue
        parts.append('🧹 卫生' if role == '卫生' else '🧡 值班')
        by_person = {}
        for r in selected:
            by_person.setdefault(_key(r), []).append(r)
        for entries in by_person.values():
            first = entries[0]
            parts.append(f"• {first.get('name') or first.get('volunteer_id') or '义工'}")
            for r in entries:
                parts.append(f"  {_time(r['start_time'])}–{_time(r['end_time'])}  {r.get('assigned_place') or role}")
        parts.append('')
    parts.append('🙏 感恩各位义工发心护持。')
    parts.extend(['', '📝 义工报名请点击以下链接：', SIGNUP_URL])
    return '\n'.join(parts)




def build_signup_message(*, date_only, weekday_only, lunar_and_festival, greeting_time, altar_start, altar_end, buddha_text, setup_notice, remove_notice, cleaning_start, cleaning_end, clean_hall, clean_upstairs, clean_stairs, green_section, orange_hall, orange_center, yellow_hall, yellow_center):
    """19:00 报名通知排版；所有报名、地点预览逻辑仍在 schedule_web。"""
    text = f"""师兄们大家好！

报名通知（初步名单）

{date_only}

{weekday_only}{lunar_and_festival}

📝 义工报名请点击以下链接👇
{SIGNUP_URL}

{greeting_time}

［整理佛台］

{altar_start}~{altar_end}  或
{altar_start}~完成佛台工作

{buddha_text}

{setup_notice}{remove_notice}［打扫卫生］

{cleaning_start}~{cleaning_end} 或
{cleaning_start}~完成卫生工作

佛堂卫生: {clean_hall}

二楼卫生: {clean_upstairs}

楼梯卫生: {clean_stairs}

每日卫生义工请注意：清理完卫生之后，请把卫生用具包括吸尘机放回原位（活动中心store里面的小房间）

［值班］

{green_section}10:00am~2:00pm

🟠 观音堂:
{orange_hall}

🟠 活动中心:
{orange_center}


2:00pm~6:00pm

🟡 观音堂:
{yellow_hall}

🟡 活动中心:
{yellow_center}

{WHATSAPP_FOOTER}
"""

    return text
