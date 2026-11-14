# -*- coding: utf-8 -*-
"""云端月报第三步：原版报表输入适配与严格核对（只读）。

不修改数据库、原版脚本或报表样本；尚不输出正式 Excel。
月报只统计实际签到，不再要求报名未签到记录。
"""
from __future__ import annotations

from collections import Counter
from datetime import date, datetime, time
from decimal import Decimal

from reports_service import load_attendance_rows, load_volunteers, load_reading_rows


class ReportDataError(ValueError):
    """数据不满足正式报表生成条件。"""


def _day(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return date.fromisoformat(str(value)[:10])


def _clock(value):
    if value is None:
        return ''
    if isinstance(value, datetime):
        value = value.time()
    if isinstance(value, time):
        return value.strftime('%H:%M')
    return str(value).strip()


def _hours(value):
    if value is None or value == '':
        return None
    number = Decimal(str(value))
    if not number.is_finite() or not 0 <= number <= 24:
        raise ReportDataError(f'异常时数: {value!r}')
    return float(number)


def prepare_monthly_inputs(year: int, month: int):
    """返回原版统计函数可用的列结构及风险清单。

    attendance_rows 为旧版中文列名的 list[dict]，可交给 pandas.DataFrame。
    按最新确认规则，仅使用实际签到；报名列不参与统计。
    """
    year, month = int(year), int(month)
    if not 2020 <= year <= 2099 or not 1 <= month <= 12:
        raise ValueError('年月不正确')
    records = load_attendance_rows(year, month)
    volunteers = load_volunteers()
    readings = load_reading_rows(year, month)
    mapped = []
    problems = []
    for index, record in enumerate(records, 1):
        try:
            d = _day(record['day'])
            if d.year != year or d.month != month:
                raise ReportDataError('记录日期超出目标月份')
            name = str(record.get('volunteer_name') or '').strip()
            role = str(record.get('role') or '').strip()
            if not name or not role:
                raise ReportDataError('缺少姓名或岗位')
            hours = _hours(record.get('hours'))
            if hours is None:
                problems.append(f'第{index}笔缺少时数（不自动补零）')
            mapped.append({
                '日期': d, '姓名': name, '报名': 0, '签到': 1,
                '岗位': role, '开始时间': _clock(record.get('start_time')),
                '结束时间': _clock(record.get('end_time')),
                '时数': hours, '义工卡号': '',
                '备注': str(record.get('remarks') or '').strip(),
                '_volunteer_id': record.get('volunteer_id'),
                '_data_source': record.get('data_source'),
            })
        except (ValueError, KeyError, ArithmeticError) as exc:
            problems.append(f'第{index}笔无效：{exc}')
    # 报名字段只是旧版接口兼容占位，正式报表不得展示或比较该字段。
    problems.append('尚未核实历史义工编号与现行编号的对应关系')
    problems.append('尚未核实 reading 云端字段与原版白话佛法统计的完整映射')
    person_days = {(r['_volunteer_id'], r['日期']) for r in mapped}
    role_days = {(r['_volunteer_id'], r['日期'], r['岗位']) for r in mapped}
    return {
        'year': year, 'month': month,
        'attendance_rows': mapped,
        'volunteer_rows': volunteers,
        'reading_rows': readings,
        'audit': {
            'record_count': len(mapped), 'person_days': len(person_days),
            'role_person_days': len(role_days),
            'roles': dict(Counter(r['岗位'] for r in mapped)),
            'source_counts': dict(Counter(r['_data_source'] for r in mapped)),
            'problems': problems,
            'ready_for_official_export': False,
        },
    }


def generate_official_monthly_excel(*args, **kwargs):
    """安全闸门：数据未核对前不生成看似正式的错误报表。"""
    raise ReportDataError('尚未完成历史编号/共修映射与原版格式核对，正式月报下载保持关闭。')


def audit_identity_and_time(inputs):
    """额外只读核对：不猜测编号映射、不修正原始数据。"""
    from collections import defaultdict
    rows = inputs['attendance_rows']
    volunteers = inputs['volunteer_rows']
    issues = []
    by_id = defaultdict(set)
    by_name = defaultdict(set)
    known_ids = {str(v['id']) for v in volunteers if v.get('id') is not None}
    for r in rows:
        vid = r.get('_volunteer_id')
        name = r['姓名']
        if vid is None:
            issues.append(f"{r['日期']} {name}：缺少义工内部ID，无法可靠合并历史编号")
        else:
            by_id[str(vid)].add(name)
            by_name[name].add(str(vid))
            if str(vid) not in known_ids:
                issues.append(f"{r['日期']} {name}：义工ID {vid} 不在当前名单")
        if r.get('时数') is None:
            issues.append(f"{r['日期']} {name} {r['岗位']}：缺少有效服务时数")
        if not r.get('开始时间') or not r.get('结束时间'):
            issues.append(f"{r['日期']} {name} {r['岗位']}：缺少签到或签退时间")
    for vid, names in by_id.items():
        if len(names) > 1:
            issues.append(f"同一义工ID {vid} 对应多个姓名：{', '.join(sorted(names))}（需核对是否改名）")
    for name, ids in by_name.items():
        if len(ids) > 1:
            issues.append(f"姓名 {name} 对应多个内部ID：{', '.join(sorted(ids))}（需核对是否同名不同人）")
    return {
        'records_checked': len(rows),
        'volunteers_checked': len(volunteers),
        'identity_conflicts': sum(len(names) > 1 for names in by_id.values()) + sum(len(ids) > 1 for ids in by_name.values()),
        'issues': issues,
        'passed': not issues,
        'official_export_enabled': False,
    }
