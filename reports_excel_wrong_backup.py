# -*- coding: utf-8 -*-
"""云端报表：只读核对预览，不是正式月报。

独立模块；不修改 reports_excel / reports_service / reports_web 或数据库。
在项目根目录运行：python reports_excel_preview.py 2026 10
"""
from __future__ import annotations

import sys
from collections import Counter, defaultdict
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from reports_excel import prepare_monthly_inputs, audit_identity_and_time


def generate_preview(year: int, month: int, output_path: str | None = None) -> str:
    data = prepare_monthly_inputs(year, month)
    identity_audit = audit_identity_and_time(data)
    rows = data['attendance_rows']
    volunteers = data['volunteer_rows']
    readings = data['reading_rows']

    # 仅用于核对，严格按内部 ID，不按姓名猜测合并。
    by_id = defaultdict(lambda: {'names': set(), 'days': set(), 'roles': set(), 'hours': 0.0, 'missing_hours': 0})
    for row in rows:
        vid = str(row.get('_volunteer_id') or '').strip()
        key = vid or f"[缺ID] {row['姓名']}"
        item = by_id[key]
        item['names'].add(row['姓名'])
        item['days'].add(row['日期'])
        item['roles'].add((row['日期'], row['岗位']))
        if row['时数'] is None:
            item['missing_hours'] += 1
        else:
            item['hours'] += row['时数']

    wb = Workbook()
    ws = wb.active
    ws.title = '核对摘要'
    ws.append(['项目', '数值'])
    for k, v in [
        ('报表期间', f'{year}-{month:02d}'),
        ('出席记录笔数', len(rows)),
        ('义工名单人数', len(volunteers)),
        ('出席人日（按ID）', len({(str(r.get('_volunteer_id') or ''), r['日期']) for r in rows})),
        ('岗位人日（按ID）', len({(str(r.get('_volunteer_id') or ''), r['日期'], r['岗位']) for r in rows})),
        ('白话佛法记录笔数', len(readings)),
        ('身份/时数问题数', len(identity_audit['issues'])),
        ('正式报表可用', '否：本文件仅供核对'),
    ]:
        ws.append([k, v])

    ws = wb.create_sheet('每人出席核对')
    ws.append(['义工ID', '姓名（记录中）', '出席天数', '岗位人日', '记录时数合计', '缺时数笔数'])
    for vid, item in sorted(by_id.items()):
        ws.append([vid, ' / '.join(sorted(item['names'])), len(item['days']),
                   len(item['roles']), round(item['hours'], 2), item['missing_hours']])

    ws = wb.create_sheet('岗位统计核对')
    ws.append(['岗位', '记录笔数', '岗位人日'])
    counts = Counter(r['岗位'] for r in rows)
    unique_roles = defaultdict(set)
    for r in rows:
        unique_roles[r['岗位']].add((str(r.get('_volunteer_id') or ''), r['日期']))
    for role in sorted(counts):
        ws.append([role, counts[role], len(unique_roles[role])])

    ws = wb.create_sheet('原始记录核对')
    ws.append(['日期', '义工ID', '姓名', '岗位', '开始', '结束', '时数', '来源'])
    for r in rows:
        ws.append([str(r['日期']), str(r.get('_volunteer_id') or ''), r['姓名'], r['岗位'],
                   r['开始时间'], r['结束时间'], r['时数'], r.get('_data_source')])

    ws = wb.create_sheet('问题清单')
    ws.append(['类型', '说明'])
    for p in data['audit']['problems']:
        ws.append(['资料核对', str(p)])
    for p in identity_audit['issues']:
        ws.append(['身份/时数', str(p)])

    for ws in wb:
        ws.freeze_panes = 'A2'
        ws.auto_filter.ref = ws.dimensions
        ws.row_dimensions[1].height = 27
        for cell in ws[1]:
            cell.font = Font(bold=True, color='FFFFFF')
            cell.fill = PatternFill('solid', fgColor='316D68')
            cell.alignment = Alignment(vertical='center')
        for col in ws.columns:
            letter = get_column_letter(col[0].column)
            width = max((len(str(c.value or '')) for c in list(col)[:200]), default=8)
            ws.column_dimensions[letter].width = min(max(width + 3, 12), 52)

    target = Path(output_path or f'{year}_{month:02d}_云端数据核对预览.xlsx')
    wb.save(target)
    return str(target)


if __name__ == '__main__':
    if len(sys.argv) not in (3, 4):
        raise SystemExit('用法: python reports_excel_preview.py 年份 月份 [输出文件.xlsx]')
    print(generate_preview(int(sys.argv[1]), int(sys.argv[2]), sys.argv[3] if len(sys.argv) == 4 else None))
