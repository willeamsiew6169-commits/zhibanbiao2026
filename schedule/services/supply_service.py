# supply_service.py

from db import get_conn
from datetime import timedelta
from psycopg2.extras import RealDictCursor
from lunar_rules import get_special_day_info
from schedule.builders.time_utils import malaysia_today


# ==========================================================
# 供台报名设置
# ==========================================================

SUPPLY_SETUP_LIMIT = 2
SUPPLY_REMOVE_LIMIT = 2

SUPPLY_TASK_SETUP = "setup"
SUPPLY_TASK_REMOVE = "remove"

VALID_SUPPLY_TASKS = {
    SUPPLY_TASK_SETUP,
    SUPPLY_TASK_REMOVE,
}


# ==========================================================
# 基础工具
# ==========================================================

def normalize_supply_task(supply_task):
    """
    标准化供台工作类型。

    setup  = 设供台
    remove = 收供台
    """

    task = str(supply_task or "").strip().lower()

    if task not in VALID_SUPPLY_TASKS:
        return None

    return task


def supply_task_label(supply_task):
    """
    将数据库值转换成人类可读文字。
    """

    task = normalize_supply_task(supply_task)

    if task == SUPPLY_TASK_SETUP:
        return "设供台"

    if task == SUPPLY_TASK_REMOVE:
        return "收供台"

    return "供台"


# ==========================================================
# 读取供台报名
# ==========================================================

def load_supply_signups_for_date(date_str, supply_task=None):
    """
    读取指定日期的供台报名。

    supply_task=None：
        读取设供台及收供台全部报名。

    supply_task='setup'：
        只读取设供台。

    supply_task='remove'：
        只读取收供台。

    为兼容旧资料：
        supply_task 为 NULL 的旧供台记录暂时当作 setup。
    """

    task = normalize_supply_task(supply_task)

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:

            if task:

                cur.execute("""
                    select
                        id,
                        volunteer_id,
                        name,
                        signup_date,
                        role,
                        coalesce(supply_task, 'setup') as supply_task,
                        status,
                        created_at
                    from volunteer_schedule_signups
                    where signup_date = %s
                      and role = '供台'
                      and coalesce(supply_task, 'setup') = %s
                      and coalesce(status, 'pending') <> 'cancelled'
                    order by created_at, id
                """, (
                    date_str,
                    task,
                ))

            else:

                cur.execute("""
                    select
                        id,
                        volunteer_id,
                        name,
                        signup_date,
                        role,
                        coalesce(supply_task, 'setup') as supply_task,
                        status,
                        created_at
                    from volunteer_schedule_signups
                    where signup_date = %s
                      and role = '供台'
                      and coalesce(status, 'pending') <> 'cancelled'
                    order by
                        case
                            when coalesce(supply_task, 'setup') = 'setup'
                                then 1
                            when supply_task = 'remove'
                                then 2
                            else 3
                        end,
                        created_at,
                        id
                """, (date_str,))

            return cur.fetchall()


def load_supply_setup_signups_for_date(date_str):
    """
    读取设供台报名名单。
    """

    return load_supply_signups_for_date(
        date_str,
        SUPPLY_TASK_SETUP,
    )


def load_supply_remove_signups_for_date(date_str):
    """
    读取收供台报名名单。
    """

    return load_supply_signups_for_date(
        date_str,
        SUPPLY_TASK_REMOVE,
    )


# ==========================================================
# 统计供台报名人数
# ==========================================================

def count_supply_signups(date_str, supply_task):
    """
    统计指定日期及工作类型的供台报名人数。
    """

    task = normalize_supply_task(supply_task)

    if not task:
        return 0

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:

            cur.execute("""
                select count(*) as cnt
                from volunteer_schedule_signups
                where signup_date = %s
                  and role = '供台'
                  and coalesce(supply_task, 'setup') = %s
                  and coalesce(status, 'pending') <> 'cancelled'
            """, (
                date_str,
                task,
            ))

            row = cur.fetchone()

    return int(row["cnt"] or 0) if row else 0


def count_supply_setup(date_str):
    """
    统计设供台人数。
    """

    return count_supply_signups(
        date_str,
        SUPPLY_TASK_SETUP,
    )


def count_supply_remove(date_str):
    """
    统计收供台人数。
    """

    return count_supply_signups(
        date_str,
        SUPPLY_TASK_REMOVE,
    )


# ==========================================================
# 判断供台是否满员
# ==========================================================

def is_supply_setup_full(date_str):
    """
    设供台是否已经满员。
    """

    return count_supply_setup(date_str) >= SUPPLY_SETUP_LIMIT


def is_supply_remove_full(date_str):
    """
    收供台是否已经满员。
    """

    return count_supply_remove(date_str) >= SUPPLY_REMOVE_LIMIT


# ==========================================================
# 取得供台报名摘要
# ==========================================================

def get_supply_signup_summary(date_str):
    """
    一次取得指定日期的供台报名资料。

    回传示例：

    {
        "setup": {
            "count": 2,
            "limit": 2,
            "is_full": True,
            "rows": [...]
        },
        "remove": {
            "count": 1,
            "limit": 2,
            "is_full": False,
            "rows": [...]
        }
    }
    """

    setup_rows = load_supply_setup_signups_for_date(date_str)
    remove_rows = load_supply_remove_signups_for_date(date_str)

    setup_count = len(setup_rows)
    remove_count = len(remove_rows)

    return {
        "setup": {
            "task": SUPPLY_TASK_SETUP,
            "label": "设供台",
            "count": setup_count,
            "limit": SUPPLY_SETUP_LIMIT,
            "is_full": setup_count >= SUPPLY_SETUP_LIMIT,
            "rows": setup_rows,
            "names": [
                row.get("name")
                for row in setup_rows
                if row.get("name")
            ],
        },
        "remove": {
            "task": SUPPLY_TASK_REMOVE,
            "label": "收供台",
            "count": remove_count,
            "limit": SUPPLY_REMOVE_LIMIT,
            "is_full": remove_count >= SUPPLY_REMOVE_LIMIT,
            "rows": remove_rows,
            "names": [
                row.get("name")
                for row in remove_rows
                if row.get("name")
            ],
        },
    }


# ==========================================================
# 未来供台报名提醒
# ==========================================================

def load_upcoming_supply_signup_alerts(days_ahead=60, limit=2):
    """
    读取未来初一、十五及佛诞日的供台报名提醒。

    每个项目会包含：

    setup_names  = 设供台名单
    remove_names = 收供台名单
    names        = 全部供台名单（保留旧页面兼容）
    """

    today = malaysia_today()
    end_date = today + timedelta(days=days_ahead)

    special_dates = []

    current_date = today

    while current_date <= end_date:

        info = get_special_day_info(current_date)

        if info["template_type"] in [
            "lunar_1_15",
            "buddhist_festival",
        ]:
            special_dates.append({
                "date": current_date,
                "type": info["template_type"],
                "setup_names": [],
                "remove_names": [],
                "names": [],
            })

        current_date += timedelta(days=1)

    if not special_dates:
        return []

    special_date_set = {
        item["date"]
        for item in special_dates
    }

    signup_map = {
        item["date"]: {
            "setup": [],
            "remove": [],
        }
        for item in special_dates
    }

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:

            cur.execute("""
                select
                    signup_date,
                    name,
                    coalesce(supply_task, 'setup') as supply_task
                from volunteer_schedule_signups
                where signup_date between %s and %s
                  and role = '供台'
                  and coalesce(status, 'pending') <> 'cancelled'
                order by signup_date, created_at, id
            """, (
                today,
                end_date,
            ))

            rows = cur.fetchall()

    for row in rows:

        signup_date = row.get("signup_date")
        name = row.get("name")
        task = normalize_supply_task(row.get("supply_task"))

        if signup_date not in special_date_set:
            continue

        if not name or not task:
            continue

        signup_map[signup_date][task].append(name)

    for item in special_dates:

        date_value = item["date"]

        item["setup_names"] = signup_map[date_value]["setup"]
        item["remove_names"] = signup_map[date_value]["remove"]

        # 保留旧页面使用 item.names 的兼容性
        item["names"] = (
            item["setup_names"]
            + item["remove_names"]
        )

        item["setup_count"] = len(item["setup_names"])
        item["remove_count"] = len(item["remove_names"])

        item["setup_limit"] = SUPPLY_SETUP_LIMIT
        item["remove_limit"] = SUPPLY_REMOVE_LIMIT

        item["setup_full"] = (
            item["setup_count"] >= SUPPLY_SETUP_LIMIT
        )

        item["remove_full"] = (
            item["remove_count"] >= SUPPLY_REMOVE_LIMIT
        )

    return special_dates[:limit]


# ==========================================================
# 读取特殊日期设置
# ==========================================================

def load_day_flags(date_str):
    """
    读取指定日期的设供台、收供台及整理佛台设置。
    """

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:

            cur.execute("""
                select
                    flag_date,
                    coalesce(
                        need_setup_master_table,
                        false
                    ) as need_setup_master_table,

                    coalesce(
                        need_remove_master_table,
                        false
                    ) as need_remove_master_table,

                    coalesce(
                        extra_buddha_person,
                        ''
                    ) as extra_buddha_person,

                    coalesce(
                        setup_people,
                        ''
                    ) as setup_people,

                    coalesce(
                        remove_people,
                        ''
                    ) as remove_people,

                    coalesce(
                        remarks,
                        ''
                    ) as remarks

                from schedule_day_flags
                where flag_date = %s
            """, (date_str,))

            row = cur.fetchone()

    if not row:
        return {
            "need_setup_master_table": False,
            "need_remove_master_table": False,

            "setup_people": "",
            "remove_people": "",

            "setup_person_1": "",
            "setup_person_2": "",

            "remove_person_1": "",
            "remove_person_2": "",
            "remove_extra_person": "",

            "extra_buddha_person": "",
            "remarks": "",
        }

    setup_list = [
        value.strip()
        for value in str(
            row.get("setup_people") or ""
        ).splitlines()
        if value.strip()
    ]

    remove_list = [
        value.strip()
        for value in str(
            row.get("remove_people") or ""
        ).splitlines()
        if value.strip()
    ]

    row["setup_person_1"] = (
        setup_list[0]
        if len(setup_list) >= 1
        else ""
    )

    row["setup_person_2"] = (
        setup_list[1]
        if len(setup_list) >= 2
        else ""
    )

    row["remove_person_1"] = (
        remove_list[0]
        if len(remove_list) >= 1
        else ""
    )

    row["remove_person_2"] = (
        remove_list[1]
        if len(remove_list) >= 2
        else ""
    )

    row["remove_extra_person"] = (
        remove_list[2]
        if len(remove_list) >= 3
        else ""
    )

    return row