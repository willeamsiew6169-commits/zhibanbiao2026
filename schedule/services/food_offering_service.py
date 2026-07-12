# schedule/services/food_offering_service.py
# 观音堂系统专用版 Food Offering V2

from datetime import datetime, time, timedelta

from db import get_conn
from psycopg2.extras import RealDictCursor

from utils import apply_branch_prefix
from lunar_rules import get_special_day_info

from schedule.builders.time_utils import (
    malaysia_now,
    malaysia_today,
)

from schedule.helpers import (
    find_volunteer_by_keyword,
)


FOOD_OFFERING_CATEGORIES = {
    "main_food": {
        "label": "主食",
        "icon": "🍚",
        "limit": 3,
        "help": "面／米粉／饭",
    },
    "vegetable": {
        "label": "菜肴",
        "icon": "🥬",
        "limit": 10,
        "help": "各类素食菜肴",
    },
    "dessert": {
        "label": "甜品",
        "icon": "🍰",
        "limit": 4,
        "help": "糕点／炸物／甜点",
    },
    "fruit": {
        "label": "水果",
        "icon": "🍉",
        "limit": 2,
        "help": "各类水果",
    },
    "drink": {
        "label": "饮料／糖水",
        "icon": "🥤",
        "limit": 1,
        "help": "饮料或糖水",
    },
}

WEEKDAY_TEXT = {
    0: "星期一",
    1: "星期二",
    2: "星期三",
    3: "星期四",
    4: "星期五",
    5: "星期六",
    6: "星期日",
}


def find_next_food_offering_date(start_date=None, days_ahead=40):

    if start_date is None:
        start_date = malaysia_today()

    for offset in range(days_ahead + 1):

        check_date = start_date + timedelta(days=offset)
        info = get_special_day_info(check_date)

        if info.get("template_type") in {
            "lunar_1_15",
            "buddhist_festival",
        }:
            return check_date, info

    return None, None


def get_food_offering_special_text(special_info):

    special_info = special_info or {}

    lunar_text = (
        special_info.get("lunar_text")
        or special_info.get("lunar_date_text")
        or special_info.get("lunar_date")
        or special_info.get("lunar")
        or ""
    )

    festival_name = (
        special_info.get("festival_name")
        or special_info.get("buddhist_name")
        or special_info.get("special_name")
        or special_info.get("name")
        or ""
    )

    return (
        str(lunar_text).strip(),
        str(festival_name).strip(),
    )


def get_food_offering_deadline(offering_date):

    deadline_date = offering_date - timedelta(days=4)
    now = malaysia_now()

    return datetime.combine(
        deadline_date,
        time(18, 0),
        tzinfo=now.tzinfo,
    )


def format_food_offering_deadline(offering_date):

    deadline = get_food_offering_deadline(
        offering_date
    )

    return (
        f"{deadline.strftime('%d/%m/%Y')} "
        f"{WEEKDAY_TEXT[deadline.weekday()]} "
        f"6:00pm"
    )


def is_food_offering_deadline_open(offering_date):

    return (
        malaysia_now()
        <= get_food_offering_deadline(offering_date)
    )


def normalize_food_text(value, max_length=100):

    return (value or "").strip()[:max_length]


def resolve_food_offering_volunteer(
    keyword,
    branch,
):

    keyword = (keyword or "").strip()
    branch = (branch or "CHE").strip().upper()

    keyword = apply_branch_prefix(
        keyword,
        branch,
    )

    if keyword.isdigit() and branch == "STW":
        keyword = f"STW-{keyword}"

    matches = find_volunteer_by_keyword(
        keyword
    )

    if not matches:
        return None, (
            "找不到义工，请检查编号、姓名或电话号码。"
        )

    if len(matches) > 1:
        return None, (
            "找到多位同名义工，请改用义工编号报名。"
        )

    volunteer = matches[0]

    return {
        "id": str(volunteer["id"]),
        "name": str(volunteer["name"]),
    }, None


def load_food_offering_records(offering_date):

    with get_conn() as conn:
        with conn.cursor(
            cursor_factory=RealDictCursor
        ) as cur:

            cur.execute("""
                select
                    id,
                    volunteer_id,
                    volunteer_name,
                    offering_date,
                    category,
                    food_name,
                    quantity_text,
                    status,
                    created_at,
                    updated_at
                from food_offering_signups
                where offering_date = %s
                  and status = 'active'
                order by
                    case category
                        when 'main_food' then 1
                        when 'vegetable' then 2
                        when 'dessert' then 3
                        when 'fruit' then 4
                        when 'drink' then 5
                        else 99
                    end,
                    created_at,
                    id
            """, (offering_date,))

            rows = cur.fetchall()

    grouped_records = {
        key: []
        for key in FOOD_OFFERING_CATEGORIES
    }

    for row in rows:

        category = row.get("category")

        if category in grouped_records:
            grouped_records[category].append(row)

    category_summary = {}

    for category, config in (
        FOOD_OFFERING_CATEGORIES.items()
    ):

        count = len(
            grouped_records[category]
        )

        limit = int(config["limit"])

        category_summary[category] = {
            "count": count,
            "limit": limit,
            "remaining": max(limit - count, 0),
            "is_full": count >= limit,
        }

    return (
        grouped_records,
        category_summary,
    )


def get_food_offering_category_count(
    cur,
    offering_date,
    category,
):

    cur.execute("""
        select count(*) as cnt
        from food_offering_signups
        where offering_date = %s
          and category = %s
          and status = 'active'
    """, (
        offering_date,
        category,
    ))

    row = cur.fetchone()

    return int(row["cnt"] or 0)


def find_duplicate_food(
    cur,
    offering_date,
    food_name,
):

    cur.execute("""
        select
            volunteer_name,
            food_name,
            quantity_text
        from food_offering_signups
        where offering_date = %s
          and status = 'active'
          and lower(trim(food_name))
              = lower(trim(%s))
        limit 1
    """, (
        offering_date,
        food_name,
    ))

    return cur.fetchone()


def save_food_offering_signup(
    *,
    volunteer_id,
    volunteer_name,
    offering_date,
    category,
    food_name,
    quantity_text,
):

    if category not in FOOD_OFFERING_CATEGORIES:

        return {
            "ok": False,
            "error": "请选择正确的结缘种类。",
        }

    food_name = normalize_food_text(
        food_name
    )

    quantity_text = normalize_food_text(
        quantity_text
    )

    if not food_name:

        return {
            "ok": False,
            "error": "请填写食物名称。",
        }

    if not quantity_text:

        return {
            "ok": False,
            "error": "请填写食物份量。",
        }

    category_limit = (
        FOOD_OFFERING_CATEGORIES[
            category
        ]["limit"]
    )

    with get_conn() as conn:
        with conn.cursor(
            cursor_factory=RealDictCursor
        ) as cur:

            current_count = (
                get_food_offering_category_count(
                    cur,
                    offering_date,
                    category,
                )
            )

            if current_count >= category_limit:

                label = (
                    FOOD_OFFERING_CATEGORIES[
                        category
                    ]["label"]
                )

                return {
                    "ok": False,
                    "error": (
                        f"{label}已经满额，"
                        f"请改选其他种类。"
                    ),
                }

            duplicate = find_duplicate_food(
                cur,
                offering_date,
                food_name,
            )

            cur.execute("""
                insert into food_offering_signups
                (
                    volunteer_id,
                    volunteer_name,
                    offering_date,
                    category,
                    food_name,
                    quantity_text,
                    status,
                    created_at,
                    updated_at
                )
                values
                (
                    %s,
                    %s,
                    %s,
                    %s,
                    %s,
                    %s,
                    'active',
                    %s,
                    %s
                )
                returning id
            """, (
                volunteer_id,
                volunteer_name,
                offering_date,
                category,
                food_name,
                quantity_text,
                malaysia_now(),
                malaysia_now(),
            ))

            signup_id = cur.fetchone()["id"]

        conn.commit()

    return {
        "ok": True,
        "signup_id": signup_id,
        "duplicate": duplicate,
    }


def cancel_food_offering_signup(
    signup_id,
):

    with get_conn() as conn:
        with conn.cursor() as cur:

            cur.execute("""
                update food_offering_signups
                set status = 'cancelled',
                    updated_at = %s
                where id = %s
                  and status = 'active'
            """, (
                malaysia_now(),
                signup_id,
            ))

            changed = cur.rowcount

        conn.commit()

    return changed > 0


def build_food_offering_whatsapp_text(
    offering_date,
    special_info,
    grouped_records,
    category_summary,
):

    lunar_text, festival_name = (
        get_food_offering_special_text(
            special_info
        )
    )

    lines = [
        "🩷 各位师兄，大家好。",
        "",
    ]

    date_line = (
        f"*{offering_date.strftime('%d/%m/%y')} "
        f"{WEEKDAY_TEXT[offering_date.weekday()]}"
    )

    if lunar_text:
        date_line += f"（{lunar_text}）"

    date_line += "*"

    lines.append(
        f"{date_line} 的素食结缘名单如下："
    )

    if festival_name:
        lines.append(
            f"🪷 {festival_name}"
        )

    lines.append("")

    for category, config in (
        FOOD_OFFERING_CATEGORIES.items()
    ):

        summary = (
            category_summary[category]
        )

        records = (
            grouped_records[category]
        )

        lines.append(
            f"👉🏻{config['label']}"
            f"（{summary['count']}/"
            f"{summary['limit']}）"
        )

        lines.append("")

        if records:

            for index, row in enumerate(
                records,
                start=1,
            ):

                food_text = row["food_name"]

                if row.get("quantity_text"):
                    food_text += (
                        f"（{row['quantity_text']}）"
                    )

                lines.append(
                    f"{index}. {food_text}　"
                    f"{row['volunteer_name']}"
                )

        else:
            lines.append("暂无报名")

        if summary["is_full"]:
            lines.append("❗已满❗")

        lines.append("")

    lines.append(
        f"*报名结缘截止日期："
        f"{format_food_offering_deadline(offering_date)}*"
    )

    lines.append("")

    lines.append(
        "🩷所有结缘的食物请在当天"
        "10:30am之前送到，"
        "以方便安排后续工作🙏🏻🩷"
    )

    return "\n".join(lines)
