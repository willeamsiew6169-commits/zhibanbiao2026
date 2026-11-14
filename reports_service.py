# -*- coding: utf-8 -*-
"""年月报表第二阶段：只读数据适配层（不修改任何现有模块或数据库）。

此模块只提供数据读取与核对；尚不输出正式 Excel。
"""
from datetime import date
from decimal import Decimal

from db import db_query

CUTOVER = date(2026, 12, 1)


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


def load_attendance_rows(year, month=None):
    """按正式切换日期读取历史/云端记录，保留每条岗位记录，不按姓名合并。"""
    start, end = _period(year, month)
    rows = db_query('''
        SELECT volunteer_id, volunteer_name, attendance_date AS day,
               role, start_time, end_time, hours,
               remarks, 'history' AS data_source
        FROM public.volunteer_attendance_history
        WHERE attendance_date >= %s AND attendance_date < %s
          AND attendance_date < DATE '2026-12-01'
        UNION ALL
        SELECT l.volunteer_id, v.name AS volunteer_name,
               l.attendance_date AS day, l.actual_role AS role,
               l.check_in_time::time AS start_time,
               l.check_out_time::time AS end_time,
               CASE WHEN l.check_in_time IS NOT NULL
                         AND l.check_out_time IS NOT NULL
                         AND l.check_out_time >= l.check_in_time
                    THEN EXTRACT(EPOCH FROM (l.check_out_time - l.check_in_time))/3600
                    ELSE NULL END AS hours,
               NULL AS remarks, 'cloud' AS data_source
        FROM public.volunteer_attendance_logs l
        LEFT JOIN public.volunteers v ON v.id = l.volunteer_id
        WHERE l.attendance_date >= %s AND l.attendance_date < %s
          AND l.attendance_date >= DATE '2026-12-01'
          AND l.check_in_time IS NOT NULL
        ORDER BY day, volunteer_id, start_time NULLS LAST
    ''', (start, end, start, end), fetchall=True) or []
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
        SELECT id, name, status, branch, phone, member_id
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
