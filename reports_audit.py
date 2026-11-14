# -*- coding: utf-8 -*-
"""云端月报上线前的只读核对工具。

运行：python reports_audit.py 2026 10
只查询 Supabase，不修改数据库、不生成正式报表。
"""
from __future__ import annotations

import argparse
import json
from collections import Counter, defaultdict
from datetime import date

from reports_excel import prepare_monthly_inputs, audit_identity_and_time


def check_month(year: int, month: int) -> dict:
    data = prepare_monthly_inputs(year, month)
    identity = audit_identity_and_time(data)
    rows = data['attendance_rows']
    issues = list(data['audit']['problems']) + list(identity['issues'])

    # 识别没有内部 ID 的多笔历史记录：不可仅凭姓名自动合并。
    missing_id = [r for r in rows if r.get('_volunteer_id') is None]
    by_day_role = defaultdict(list)
    for r in rows:
        by_day_role[(r['日期'], r['姓名'], r['岗位'])].append(r)
    repeated = [
        {'日期': str(k[0]), '姓名': k[1], '岗位': k[2], '记录数': len(v)}
        for k, v in by_day_role.items() if len(v) > 1
    ]
    # 重复岗位记录可能是合法分段值班；仅提醒核对，不自动删除。
    if repeated:
        issues.append(f'发现 {len(repeated)} 组同日同人同岗位多笔记录（可能为分段值班）')
    if missing_id:
        issues.append(f'{len(missing_id)} 笔缺少义工内部ID，不能可靠计算跨编号的个人统计')

    return {
        '期间': f'{year}-{month:02d}',
        '出席记录数': len(rows),
        '按内部ID统计的出席人日数_仅供核对': data['audit']['person_days'],
        '按内部ID统计的岗位人日数_仅供核对': data['audit']['role_person_days'],
        '历史记录数': data['audit']['source_counts'].get('history', 0),
        '正式云端记录数': data['audit']['source_counts'].get('cloud', 0),
        '义工名单人数': len(data['volunteer_rows']),
        '白话佛法记录数': len(data['reading_rows']),
        '岗位记录数': dict(Counter(r['岗位'] for r in rows)),
        '同日同人同岗位多笔记录': repeated[:30],
        '核对问题': list(dict.fromkeys(issues)),
        '允许正式下载': False,
    }


def main():
    parser = argparse.ArgumentParser(description='云端月报只读核对')
    parser.add_argument('year', type=int)
    parser.add_argument('month', type=int)
    args = parser.parse_args()
    try:
        result = check_month(args.year, args.month)
    except Exception as exc:
        print('核对失败：数据库查询或数据格式不符合预期。')
        print(f'{type(exc).__name__}: {exc}')
        raise SystemExit(1)
    print(json.dumps(result, ensure_ascii=False, indent=2, default=str))


if __name__ == '__main__':
    main()
