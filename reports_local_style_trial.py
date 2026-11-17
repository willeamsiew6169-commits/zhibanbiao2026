# -*- coding: utf-8 -*-
"""只读试生成：Supabase 数据 -> 本地 zhibanbiao 原版月报版式。
不会修改数据库或现有脚本；结果仅用于逐人核对，不是正式报表。
运行: python reports_local_style_trial.py 2026 10
"""
import sys
import tempfile
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.comments import Comment
from reports_excel import prepare_monthly_inputs, ReportDataError, _day, _clock, _hours
from reports_service import load_attendance_rows


def generate(year, month):
    year, month = int(year), int(month)
    data = prepare_monthly_inputs(year, month)
    rows = data['attendance_rows']
    if not rows:
        raise ReportDataError('没有读取到出席记录，停止生成')
    if data['audit']['record_count'] != len(rows):
        raise ReportDataError('记录数不一致')
    # 原版按姓名汇总。云端如有同名不同人，禁止生成可能混人的月报。
    by_name = {}
    for r in rows:
        name = r['姓名']
        vid = str(r.get('_volunteer_id') or '').strip()
        if not vid:
            raise ReportDataError(f'{name} 缺义工ID，停止生成')
        by_name.setdefault(name, set()).add(vid)
    collisions = [n for n, ids in by_name.items() if len(ids) > 1]
    if collisions:
        raise ReportDataError('同名多ID，原版按姓名统计有混人风险：' + ', '.join(collisions))

    # 只调用原版统计和排版函数，不调用原版 main()（main 会改写本地 Excel）。
    import reports_original_engine as original
    vol = []
    for v in data['volunteer_rows']:
        name = str(v.get('name') or '').strip()
        if not name:
            continue
        vol.append({
            '编号': str(v.get('id') or ''),
            '姓名': name, '状态': str(v.get('status') or '在册'),
            '电话号码': str(v.get('phone') or ''),
            '是否义工': '是' if v.get('is_volunteer') is True else '否' if v.get('is_volunteer') is False else '是',
            '是否月费': '是' if v.get('is_member') is True else '否',
            '月费编号': str(v.get('member_id') or '') if v.get('is_member') is True else '',
            '备注': str(v.get('remark') or ''),
        })
    vol_df = pd.DataFrame(vol)
    # 仅为「本月未值班」的最后值班日期提供此前历史签到。
    # 不改变目标月份的 rows，也不修改数据库或报表引擎。
    history_rows = []
    for previous_month in range(1, month):
        for record in load_attendance_rows(year, previous_month):
            d = _day(record['day'])
            if d.year != year or d.month != previous_month:
                raise ReportDataError('历史记录日期超出查询月份，停止生成')
            name = str(record.get('volunteer_name') or '').strip()
            if not name:
                raise ReportDataError(f'{d} 历史记录缺少姓名，停止生成')
            history_rows.append({
                '日期': d, '姓名': name, '报名': 0, '签到': 1,
                '岗位': str(record.get('role') or '').strip(),
                '开始时间': _clock(record.get('start_time')),
                '结束时间': _clock(record.get('end_time')),
                '时数': _hours(record.get('hours')),
                '义工卡号': '',
                '备注': str(record.get('remarks') or '').strip(),
            })
    att_df = pd.DataFrame(history_rows + rows)
    if len(att_df[
        (pd.to_datetime(att_df['日期']).dt.year == year) &
        (pd.to_datetime(att_df['日期']).dt.month == month)
    ]) != len(rows):
        raise ReportDataError('补充历史记录后目标月份笔数变化，停止生成')
    # 对应原版接口；只统计云端已签到资料，报名=0 是适配占位，
    # 因此报名/签到异常页不能用作正式异常结论。
    month_args = original.build_monthly_summary(vol_df, att_df, year, month)
    # 只调整「月表」的姓名展示顺序，不重新排序签到记录或改动统计。
    # 原版名单采用拼音排序；月表原函数本身沿用出席记录首次出现顺序。
    from pypinyin import lazy_pinyin
    args = list(month_args)
    original_names = list(args[3])
    args[3] = sorted(original_names, key=lambda name: ("".join(lazy_pinyin(str(name))).lower(), str(name)))
    if len(args[3]) != len(original_names) or set(args[3]) != set(original_names):
        raise ReportDataError('月表排序前后人员集合不一致，停止生成')
    month_args = tuple(args)

    # 仅本次试生成：将原版读取 reading.xlsx 的动作映射到 Supabase 查询结果。
    # 原版引擎保持原样；不写入或覆盖本地 reading.xlsx。
    reading_rows = data['reading_rows']
    reading = []
    for index, item in enumerate(reading_rows, 1):
        try:
            reading_date = _day(item['date'])
        except (KeyError, ValueError, TypeError) as exc:
            raise ReportDataError(f'第{index}笔白话佛法日期无效：{exc}') from exc
        if reading_date.year != year or reading_date.month != month:
            raise ReportDataError(f'第{index}笔白话佛法日期超出目标月份')
        name = str(item.get('name') or '').strip()
        if not name:
            raise ReportDataError(f'第{index}笔白话佛法缺少姓名')
        reading.append({
            '日期': reading_date, '姓名': name,
            '身份': str(item.get('identity') or '').strip(),
            '主题': str(item.get('topic') or '').strip(),
            '场次': str(item.get('session') or '').strip(),
            '时间': str(item.get('time') or '').strip(),
        })
    reading_df = pd.DataFrame(reading, columns=['日期', '姓名', '身份', '主题', '场次', '时间'])

    with tempfile.TemporaryDirectory(prefix='gyt_report_trial_') as tmp:
        # 原版 report_dir() 按 __file__ 写出；只重定向到临时目录。
        original.report_dir = lambda _year: tmp
        # 原版内部调用 pd.read_excel("reading.xlsx")；仅拦截这一处读取。
        # 其余 Excel 读取仍使用原始 pandas 函数，且异常时必定恢复。
        real_read_excel = pd.read_excel
        reading_read_count = 0

        def _trial_read_excel(io, *args, **kwargs):
            nonlocal reading_read_count
            if isinstance(io, (str, Path)) and str(io).replace('\\', '/').split('/')[-1] == 'reading.xlsx':
                reading_read_count += 1
                return reading_df.copy()
            return real_read_excel(io, *args, **kwargs)

        try:
            pd.read_excel = _trial_read_excel
            generated = Path(original.write_monthly_report(*month_args, year, month))
        finally:
            pd.read_excel = real_read_excel
        if reading_read_count != 1:
            raise ReportDataError(f'原版白话佛法读取次数异常：{reading_read_count}，停止生成')
        target = Path(f'{year}_{month:02d}_云端仿本地月报_仅核对.xlsx')
        # 原版引擎要求报名为数字，不能在统计前改成「—」。
        # 仅在已生成的核对版 Excel 中隐藏没有真实来源的报名占位值。
        workbook = load_workbook(generated)
        # 只修复核对版月表中的「签到没报名」误报红色。
        # 报名=0 是云端适配占位，并非真实未报名；不修改数据或原版引擎。
        monthly = workbook[f'{month}月月表']
        repaired = 0
        for row in monthly.iter_rows():
            for cell in row:
                comment = cell.comment
                if (cell.value != 1 or comment is None
                        or not comment.text.startswith('签到没报名\n岗位：')):
                    continue
                if cell.fill.patternType != 'solid' or cell.fill.fgColor.rgb[-6:].upper() != original.ANOMALY_COLORS['签到没报名']:
                    raise ReportDataError(f'{cell.coordinate} 异常颜色与预期不符，停止生成')
                roles_text = comment.text.split('岗位：', 1)[1].split('\n备注：', 1)[0].strip()
                roles = [role.strip() for role in roles_text.split(' & ') if role.strip()]
                if len(roles) == 1:
                    color = original.ROLE_COLORS.get(roles[0], 'FFFFFF')
                elif len(roles) > 1:
                    color = original.MULTI_ROLE_COLOR
                else:
                    raise ReportDataError(f'{cell.coordinate} 缺岗位资料，停止生成')
                cell.fill = PatternFill('solid', fgColor=color)
                cell.comment = Comment('签到\n岗位：' + comment.text.split('岗位：', 1)[1], comment.author)
                repaired += 1
        # 原版对阅读统计异常只打印提示；在输出前额外核对人数和总次数。
        if reading:
            if '白话佛法统计' not in workbook.sheetnames:
                raise ReportDataError('云端白话佛法有记录，但原版未生成统计页，停止生成')
            reading_ws = workbook['白话佛法统计']
            expected_counts = reading_df.groupby('姓名').size().to_dict()
            actual_counts = {}
            for reading_cells in reading_ws.iter_rows(min_row=5, min_col=1, max_col=2):
                name_cell, count_cell = reading_cells
                if name_cell.value in (None, '', 'Total'):
                    continue
                name = str(name_cell.value).strip()
                if name in actual_counts:
                    raise ReportDataError(f'白话佛法统计重复姓名：{name}')
                actual_counts[name] = count_cell.value
            if actual_counts != expected_counts:
                raise ReportDataError('云端白话佛法逐人次数与生成工作表不一致，停止生成')
        elif '白话佛法统计' in workbook.sheetnames:
            raise ReportDataError('云端无白话佛法记录，但生成了统计页，停止生成')

        detail = workbook['月明细']
        if detail.cell(1, 3).value != '报名':
            raise ReportDataError('月明细报名栏位置已变化，停止输出，避免误改其他列')
        for cells in detail.iter_rows(min_row=2, min_col=3, max_col=3):
            cells[0].value = '—'
        target = Path(f'{year}_{month:02d}_云端仿本地月报_仅核对.xlsx')
        workbook.save(target)
    print('核对版已生成：', target.resolve())
    print('目标月份记录数：', len(rows))
    print('补充本年度此前历史记录数：', len(history_rows))
    print('月表已恢复岗位颜色的签到格数：', repaired)
    print('月表姓名排序：拼音升序（仅月表，不影响统计）')
    print('Supabase 白话佛法记录数：', len(reading))
    print('Supabase 白话佛法参与人数：', reading_df['姓名'].nunique())
    print('警告：月明细报名栏仅显示—；真实报名数据未核实，异常统计仍不可作为正式结论。')
    print('警告：仅白话佛法来源已接入并核对逐人次数；其余正式月报审计仍未完成。')
    return target


if __name__ == '__main__':
    if len(sys.argv) != 3:
        raise SystemExit('用法：python reports_local_style_trial.py 2026 10')
    generate(int(sys.argv[1]), int(sys.argv[2]))
