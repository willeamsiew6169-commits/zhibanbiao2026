# smart_assignment_engine.py

from psycopg2.extras import RealDictCursor

from db import get_conn


PLACE_ROTATION_ENABLED = True
PLACE_ROTATION_LOOKBACK = 10


DUTY_PLACES = [
    "观音堂",
    "活动中心",
]


def load_place_history_cache():
    """
    读取最近值班地点历史。
    只统计 role='值班'。
    不统计卫生、佛台、供台、整理佛台等。
    """

    history = {}

    if not PLACE_ROTATION_ENABLED:
        return history

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                select
                    volunteer_id,
                    assigned_place,
                    assignment_date,
                    id
                from volunteer_schedule_assignments
                where role = '值班'
                  and assigned_place in ('观音堂', '活动中心')
                  and volunteer_id is not null
                order by volunteer_id, assignment_date desc, id desc
            """)

            rows = cur.fetchall()

    for r in rows:
        volunteer_id = str(r.get("volunteer_id") or "").strip()
        place = r.get("assigned_place")

        if not volunteer_id or place not in DUTY_PLACES:
            continue

        if volunteer_id not in history:
            history[volunteer_id] = {
                "records": [],
                "counts": {
                    "观音堂": 0,
                    "活动中心": 0,
                }
            }

        if len(history[volunteer_id]["records"]) >= PLACE_ROTATION_LOOKBACK:
            continue

        history[volunteer_id]["records"].append(place)
        history[volunteer_id]["counts"][place] += 1

    return history


def get_rotation_preferred_place(volunteer_id, history_cache):
    """
    根据最近 N 次值班地点，建议这次优先地点。
    观音堂多 -> 活动中心
    活动中心多 -> 观音堂
    一样多 -> None
    """

    if not PLACE_ROTATION_ENABLED:
        return None

    if not volunteer_id:
        return None

    data = history_cache.get(volunteer_id)

    if not data:
        return None

    counts = data.get("counts", {})

    gyt_count = counts.get("观音堂", 0)
    act_count = counts.get("活动中心", 0)

    if gyt_count > act_count:
        return "活动中心"

    if act_count > gyt_count:
        return "观音堂"

    return None


def update_place_history_cache(volunteer_id, place, history_cache):
    """
    排班过程中，内存即时更新。
    这样同一次自动排班内也会越来越公平。
    """

    if not volunteer_id or place not in DUTY_PLACES:
        return

    if volunteer_id not in history_cache:
        history_cache[volunteer_id] = {
            "records": [],
            "counts": {
                "观音堂": 0,
                "活动中心": 0,
            }
        }

    history_cache[volunteer_id]["records"].insert(0, place)
    history_cache[volunteer_id]["counts"][place] += 1

    if len(history_cache[volunteer_id]["records"]) > PLACE_ROTATION_LOOKBACK:
        removed = history_cache[volunteer_id]["records"].pop()

        if removed in DUTY_PLACES:
            history_cache[volunteer_id]["counts"][removed] -= 1

def calculate_best_place(
    shift,
    item,
    preferred_place,
    cleaning_place_by_name,
    choose_less_place_func,
):
    """
    V5 智能地点判断：
    1. 财政 CHE-238 特殊规则
    2. 负责人/系统指定地点
    3. 佛堂卫生者优先活动中心
    4. 地点轮换
    5. 人数平衡
    """

    volunteer_id = str(
        item.get("volunteer_id")
        or item.get("编号")
        or ""
    ).strip()

    name = item.get("name") or item.get("姓名")

    # 1. 财政特殊规则
    if volunteer_id == "CHE-238":
        if shift in ["绿", "橙"]:
            return "活动中心"

        if shift == "黄":
            return "观音堂"

    # 2. 已指定地点
    if preferred_place:
        return preferred_place

    # 3. 佛堂卫生者优先活动中心
    if cleaning_place_by_name.get(name) == "佛堂卫生":
        return "活动中心"

    # 4. 地点轮换
    rotation_place = get_rotation_preferred_place(
        volunteer_id,
        item.get("_place_history_cache", {})
    )

    if rotation_place:
        return rotation_place

    # 5. 原本人数平衡
    return choose_less_place_func(shift)