# -*- coding: utf-8 -*-
"""年月报表第二阶段：只读数据适配层（不修改任何现有模块或数据库）。

此模块只提供数据读取与核对；尚不输出正式 Excel。
"""
from datetime import date
from decimal import Decimal

from db import db_query

CUTOVER = date(2026, 11, 15)


def _period(year, month=None):
    year = int(year)
    if not 2020 <= year <= 2100:
        raise ValueError('年份超出允许范围')
    if month is None:
        return date(year, 1, 1), date(year + 1, 1, 1)
    month = int(month)
    if not 1 <= month <= 12:
        raise ValueError('月份必须在1至12之间')
    return date(year, month, 1), (date(year + 1, 1, 1) if month == 12 else date(year, month + 1, 1))


def load_live_attendance(start, end):
    """One row per official attendance; explicitly identified farm supplements only."""
    import re
    rows = db_query("""
        SELECT id, volunteer_id, name AS volunteer_name, date::date AS day,
               role, start_time::text AS start_time, end_time::text AS end_time,
               hours, remark AS remarks, 'cloud' AS data_source
        FROM public.attendance
        WHERE date::date >= %s AND date::date < %s AND date::date >= DATE '2026-11-15'
        ORDER BY date, id
    """, (start, end), fetchall=True) or []
    logs = db_query("""
        SELECT id, assignment_id, volunteer_id, name, attendance_date AS day,
               actual_role, actual_place, check_in_time, check_out_time, remarks
        FROM public.volunteer_attendance_logs
        WHERE attendance_date >= %s AND attendance_date < %s
          AND attendance_date >= DATE '2026-11-15'
        ORDER BY id
    """, (start, end), fetchall=True) or []
    def key(row, role):
        return (row['day'], str(row['volunteer_id']), role)
    grouped = {}
    for log in logs:
        grouped.setdefault(key(log, log['actual_role']), []).append(log)
    occupied = set()
    for row in rows:
        k = key(row, row['role'])
        occupied.add(k)
        marker = re.search(r'assignment_ids?=([0-9,]+)', row.get('remarks') or '')
        candidates = grouped.get(k, [])
        if marker:
            ids = set(map(int, marker.group(1).split(',')))
            candidates = [l for l in candidates if l['assignment_id'] in ids]
            row['warning'] = '' if len({l['assignment_id'] for l in candidates}) == len(ids) else '岗位或排班日志不一致，请核对'
        else:
            candidates = [l for l in candidates if l['assignment_id'] is None][-1:]
            row['warning'] = ''
        row['place'] = '、'.join(dict.fromkeys(l['actual_place'] for l in candidates if l['actual_place']))
        if row['role'] == '值班' and row.get('hours') is not None and float(row['hours']) > 12:
            row['warning'] = (row['warning'] + '；' if row['warning'] else '') + '值班时数超过12小时，请核对'
    for log in logs:
        k = key(log, log['actual_role'])
        if (k in occupied or log['day'] != date(2026, 11, 15)
                or log['volunteer_id'] not in ('CHE-232', 'CHE-188')
                or log['actual_role'] != '佛学班' or log['actual_place'] != '农舍'
                or log['remarks'] != '历史补录：农舍佛学班，实际出席2小时'
                or not log['check_in_time'] or not log['check_out_time']):
            continue
        rows.append(dict(volunteer_id=log['volunteer_id'], volunteer_name=log['name'],
                         day=log['day'], role='佛学班', place='农舍',
                         start_time=log['check_in_time'].strftime('%H:%M'),
                         end_time=log['check_out_time'].strftime('%H:%M'),
                         hours=Decimal(str((log['check_out_time']-log['check_in_time']).total_seconds())) / 3600,
                         remarks=log['remarks'], data_source='cloud', warning=''))
        occupied.add(k)
    return rows


def load_attendance_rows(year, month=None):
    """Historical cutoff and official credited service hours."""
    start, end = _period(year, month)
    rows = db_query("""
        SELECT volunteer_id, volunteer_name, attendance_date AS day,
               role, start_time, end_time, hours, remarks, 'history' AS data_source
        FROM public.volunteer_attendance_history
        WHERE attendance_date >= %s AND attendance_date < %s
          AND attendance_date < DATE '2026-11-15'
        ORDER BY attendance_date, volunteer_id, start_time NULLS LAST
    """, (start, end), fetchall=True) or []
    if end > CUTOVER:
        rows.extend(load_live_attendance(max(start, CUTOVER), end))
    # 报表展示统一采用当前正式编号；数据库历史记录保持原样。
    # 只使用 volunteers.id -> member_id 的明确关联，不按姓名猜测身份。
    volunteers = load_volunteers()
    identity = {}
    for v in volunteers:
        internal_id = str(v.get('id') or '').strip()
        official_id = str(v.get('member_id') or '').strip() or internal_id
        if internal_id:
            identity[internal_id.upper()] = (official_id, v.get('name'))
        if official_id:
            identity[official_id.upper()] = (official_id, v.get('name'))
    for row in rows:
        original_id = str(row.get('volunteer_id') or '').strip()
        current = identity.get(original_id.upper())
        if current:
            row['volunteer_id'] = current[0]
            if current[1]:
                row['volunteer_name'] = current[1]
    return rows


def load_volunteers():
    """只读取报表核对所需的稳定基础字段，不推断历史编号。"""
    return db_query('''
        SELECT id, name, status, branch, phone, member_id,
               is_volunteer, is_member, remark
        FROM public.volunteers ORDER BY name, id
    ''', fetchall=True) or []


def load_reading_rows(year, month=None):
    start, end = _period(year, month)
    return db_query('''
        SELECT id, date, name, identity, topic, session, time
        FROM public.reading
        WHERE date >= %s AND date < %s
        ORDER BY date, id
    ''', (start.isoformat(), end.isoformat()), fetchall=True) or []


def audit_month(year, month):
    """返回只读核对摘要；不代表已通过与原版 Excel 的一致性验证。"""
    records = load_attendance_rows(year, month)
    volunteers = load_volunteers()
    reading = load_reading_rows(year, month)
    unique_days = {(r['volunteer_id'], r['day']) for r in records}
    role_days = {(r['volunteer_id'], r['day'], r['role']) for r in records}
    missing_names = [r for r in records if not r.get('volunteer_name')]
    unknown_hours = [r for r in records if r.get('hours') is None]
    return {
        'year': int(year), 'month': int(month),
        'attendance_records': len(records),
        'attendance_person_days': len(unique_days),
        'role_person_days': len(role_days),
        'volunteer_count': len(volunteers),
        'reading_records': len(reading),
        'missing_names': len(missing_names),
        'missing_hours': len(unknown_hours),
        'historical_records': sum(r['data_source'] == 'history' for r in records),
        'cloud_records': sum(r['data_source'] == 'cloud' for r in records),
    }
