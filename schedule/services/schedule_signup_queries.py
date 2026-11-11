# 第七阶段：仅迁移报名统计及负责人报名详情查询函数，保持函数体不变。
from db import get_conn
from psycopg2.extras import RealDictCursor


def get_signup_summary_for_date(date_str):
    from db import db_query

    row = db_query("""
        select
            count(*) as total,
            count(*) filter (
                where coalesce(status, 'pending') = 'pending'
            ) as pending,
            count(*) filter (
                where status = 'assigned'
            ) as assigned,
            count(*) filter (
                where status = 'cancelled'
            ) as cancelled
        from volunteer_schedule_signups
        where signup_date = %s
    """, (date_str,), fetchone=True)

    return {
        "total": row["total"] or 0,
        "pending": row["pending"] or 0,
        "assigned": row["assigned"] or 0,
        "cancelled": row["cancelled"] or 0,
    }


def load_admin_signup_details(date_str):
    """负责人日期面板：显示所有报名（含取消），地点以正式 assignments 为准。"""
    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                SELECT id, volunteer_id, name, role, start_time, end_time,
                       COALESCE(status, 'pending') AS status, assigned_place
                FROM volunteer_schedule_signups
                WHERE signup_date = %s
                ORDER BY CASE COALESCE(status, 'pending')
                           WHEN 'pending' THEN 0
                           WHEN 'assigned' THEN 1
                           WHEN 'cancelled' THEN 2
                           ELSE 3 END,
                         start_time NULLS LAST, name, role
            """, (date_str,))
            signups = cur.fetchall()
            # 负责人指定的师父供台人员不一定经过报名表，单独补充到详情显示。
            cur.execute("""
                SELECT setup_people, remove_people
                FROM schedule_day_flags WHERE flag_date = %s
            """, (date_str,))
            master_flag = cur.fetchone() or {}

            cur.execute("""
                SELECT signup_id, volunteer_id, role, assigned_place,
                       start_time, end_time
                FROM volunteer_schedule_assignments
                WHERE assignment_date = %s
                  AND COALESCE(status, 'assigned') <> 'cancelled'
                ORDER BY start_time NULLS LAST, end_time NULLS LAST, assigned_place
            """, (date_str,))
            assignments = cur.fetchall()

    by_signup = {}
    by_person_role = {}
    for item in assignments:
        sid = item.get('signup_id')
        if sid is not None:
            by_signup.setdefault(str(sid), []).append(item)
        key = (str(item.get('volunteer_id') or '').strip(),
               str(item.get('role') or '').strip())
        by_person_role.setdefault(key, []).append(item)

    for signup in signups:
        # 已取消的报名不能显示其他有效排班，避免误导负责人。
        if signup.get('status') == 'cancelled':
            continue
        key = (str(signup.get('volunteer_id') or '').strip(),
               str(signup.get('role') or '').strip())
        items = by_signup.get(str(signup['id'])) or by_person_role.get(key, [])
        if items:
            places = []
            seen = set()
            for item in items:
                place = str(item.get('assigned_place') or '').strip()
                if not place:
                    continue
                start = str(item.get('start_time') or '').strip()
                end = str(item.get('end_time') or '').strip()
                detail = f'{place}（{start}～{end}）' if start and end else place
                if detail not in seen:
                    places.append(detail)
                    seen.add(detail)
            if places:
                signup['assignment_places'] = places
    # 仅追加负责人手动指定的供台显示行；不写入报名/排班表，也不影响统计。
    existing = {(str(row.get('name') or '').strip(), str(row.get('role') or '').strip())
                for row in signups if row.get('status') != 'cancelled'}
    for role, field, start, end in (
        ('设师父供台', 'setup_people', '06:00', '08:00'),
        ('收师父供台', 'remove_people', '12:00', '14:00'),
    ):
        names = [name.strip() for name in str(master_flag.get(field) or '').splitlines()
                 if name.strip()]
        for name in dict.fromkeys(names):
            if (name, role) in existing:
                continue
            signups.append({
                'volunteer_id': None, 'name': name, 'role': role,
                'start_time': start, 'end_time': end,
                'status': 'assigned', 'assigned_place': role,
                'assignment_places': [role],
            })
    return signups
