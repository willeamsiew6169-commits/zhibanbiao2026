# 第八阶段：原函数逐字搬迁，不改变查询与输出逻辑。
from db import get_conn
from psycopg2.extras import RealDictCursor


def build_month_prebook_message(year, month):
    import calendar
    from datetime import date

    first_day = date(year, month, 1)
    last_day = date(year, month, calendar.monthrange(year, month)[1])

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                select
                    signup_date,
                    name,
                    role,
                    start_time,
                    end_time
                from volunteer_schedule_signups
                where signup_date between %s and %s
                and coalesce(status, 'pending') <> 'cancelled'
                order by signup_date, role, start_time, name
            """, (first_day, last_day))
            rows = cur.fetchall()

    by_day = {}
    for r in rows:
        d = r["signup_date"]
        by_day.setdefault(d, []).append(r)

    weekday_map = {
        0: "星期一",
        1: "星期二",
        2: "星期三",
        3: "星期四",
        4: "星期五",
        5: "星期六",
        6: "星期日",
    }

    lines = []
    lines.append(f"{year}年{month}月份预报名名单")
    lines.append("")
    lines.append("以下为目前已收到的预报名资料，最终岗位安排以负责人公布值班表为准。")
    lines.append("")

    days_in_month = calendar.monthrange(year, month)[1]

    for day in range(1, days_in_month + 1):
        d = date(year, month, day)
        day_rows = by_day.get(d, [])

        if not day_rows:
            continue

        cleaning = []
        supply = []
        duty = []
        others = []

        for r in day_rows:
            name = str(r["name"]).strip()
            role = str(r["role"]).strip()
            start = r.get("start_time") or ""
            end = r.get("end_time") or ""

            if role == "卫生":
                cleaning.append(name)

            elif role == "供台":
                supply.append(name)

            elif role == "值班":
                if start and end:
                    duty.append(f"{name} {start}~{end}")
                else:
                    duty.append(name)

            else:
                others.append(f"{name}（{role}）")

        lines.append(f"{day}/{month}/{year} {weekday_map[d.weekday()]}")

        if cleaning:
            lines.append("卫生：" + "、".join(cleaning))

        if supply:
            lines.append("供台：" + "、".join(supply))

        if duty:
            lines.append("值班：")
            lines.extend(duty)

        if others:
            lines.append("其他：")
            lines.extend(others)

        lines.append("")

    return "\n".join(lines)


def build_shortage_notice(date_str):

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:

            cur.execute("""
                select role
                from volunteer_schedule_signups
                where signup_date = %s
                and coalesce(status,'pending') <> 'cancelled'
            """, (date_str,))

            rows = cur.fetchall()

    cleaning_count = 0
    duty_count = 0

    for r in rows:

        role = str(r.get("role") or "").strip()

        if role == "卫生":
            cleaning_count += 1

        elif role == "值班":
            duty_count += 1

    notices = []

    if cleaning_count < 2:
        notices.append(
            f"卫生可以多加 {2-cleaning_count} 位义工。"
        )

    if duty_count < 6:
        notices.append(
            "全日值班可以多加 2~4 位义工。"
        )

    if not notices:
        return f"""
师兄们，大家好！

{date_str} 的报名人数暂时充足。

感恩大家发心护持观音堂。
🙏🙏🙏
"""

    body = "\n".join(notices)

    return f"""
师兄们，大家好！

{body}

感恩大家🙏🙏🙏
"""

