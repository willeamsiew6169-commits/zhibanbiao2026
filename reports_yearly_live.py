# -*- coding: utf-8 -*-
"""Read-only, isolated annual Excel generation using the existing annual engine."""
from pathlib import Path
import tempfile
import pandas as pd
from reports_service import load_attendance_rows, load_reading_rows, load_volunteers
from reports_excel import ReportDataError, _day, _clock, report_hours, pending_note


def generate(year):
    year = int(year)
    if not 2020 <= year <= 2099:
        raise ReportDataError('年份不正确')
    rows = load_attendance_rows(year)
    if not rows:
        raise ReportDataError('该年度没有出席记录')
    volunteers = load_volunteers()
    reading_rows = load_reading_rows(year)
    # 从历史记录及云端读取全年实际出席；2026年11月15日起的签到也纳入统计。
    # 不修改数据库；沿用 reports_service 的历史/云端去重切换规则。
    identities = {}
    for r in rows:
        name = str(r.get('volunteer_name') or '').strip()
        vid = str(r.get('volunteer_id') or '').strip()
        if not name or not vid:
            raise ReportDataError('出席记录缺姓名或义工编号')
        identities.setdefault(name, set()).add(vid)
    conflicts = [n for n, ids in identities.items() if len(ids) > 1]
    if conflicts:
        raise ReportDataError('同名不同编号，原版年报按姓名统计可能混人：' + '、'.join(conflicts))
    attendance = []
    for r in rows:
        d = _day(r['day'])
        if d.year != year:
            raise ReportDataError('年度出席日期不一致')
        role = str(r.get('role') or '').strip()
        if not role:
            raise ReportDataError('出席记录缺岗位')
        hours = report_hours(r)
        attendance.append({'日期': d, '姓名': str(r['volunteer_name']).strip(),
                           '报名': 0, '签到': 1, '岗位': role,
                           '开始时间': _clock(r.get('start_time')),
                           '结束时间': _clock(r.get('end_time')),
                           '时数': hours, '义工卡号': str(r['volunteer_id']),
                           '备注': str(r.get('remarks') or '')})
    volunteer_data = []
    for v in volunteers:
        name = str(v.get('name') or '').strip()
        if name:
            volunteer_data.append({'编号': str(v.get('member_id') or v.get('id') or ''),
                                   '姓名': name, '状态': str(v.get('status') or '在册'),
                                   '电话号码': str(v.get('phone') or ''),
                                   '是否义工': '否' if v.get('is_volunteer') is False else '是',
                                   '是否月费': '是' if v.get('is_member') is True else '否',
                                   '月费编号': str(v.get('member_id') or ''),
                                   '备注': str(v.get('remark') or '')})
    reading = []
    for r in reading_rows:
        d = _day(r['date'])
        if d.year != year:
            raise ReportDataError('白话佛法记录日期超出目标年份')
        name = str(r.get('name') or '').strip()
        if not name:
            raise ReportDataError('白话佛法记录缺姓名')
        reading.append({'日期': d, '姓名': name, '身份': str(r.get('identity') or ''),
                        '主题': str(r.get('topic') or ''), '场次': str(r.get('session') or ''),
                        '时间': str(r.get('time') or '')})
    reading_df = pd.DataFrame(reading, columns=['日期', '姓名', '身份', '主题', '场次', '时间'])
    import reports_original_engine as original
    vol_df = pd.DataFrame(volunteer_data)
    att_df = pd.DataFrame(attendance)
    # Original engine has name-based aggregation: fail closed on any ambiguity.
    summary, roles, detail = original.build_yearly_summary(vol_df, att_df, year)
    pending = att_df['时数'].isna()
    if pending.any():
        # 原版汇总忽略未结算时数，但明细不得显示为零。
        detail['时数'] = detail['时数'].astype(object)
        detail.loc[pending[pending].index, '时数'] = '待结算'
    year_table = original.build_year_summary_table(vol_df, att_df, year)
    with tempfile.TemporaryDirectory(prefix='gyt_yearly_') as tmp:
        original.report_dir = lambda _year: tmp
        # 原版年报先检查 reading.xlsx 是否真实存在，再读取。
        # 只在临时目录提供云端数据，不修改原版引擎或正式 Excel。
        reading_file = Path(tmp) / 'reading.xlsx'
        reading_df.to_excel(reading_file, index=False)
        original_reading_file = original.READING_FILE
        try:
            original.READING_FILE = str(reading_file)
            output = Path(original.write_yearly_report(summary, roles, detail, year, year_table))
        finally:
            original.READING_FILE = original_reading_file
        if not output.is_file():
            raise ReportDataError('年报输出文件不存在')
        from openpyxl import load_workbook
        wb = load_workbook(output, read_only=True)
        try:
            if reading and '白话佛法年统计' not in wb.sheetnames:
                raise ReportDataError('年报缺少白话佛法年统计')
        finally:
            wb.close()
        if pending.any():
            wb = load_workbook(output)
            pending_note(wb, int(pending.sum()))
            wb.save(output)
            wb.close()
        result = Path.cwd() / f'{year}_year_report.xlsx'
        result.write_bytes(output.read_bytes())
    print('年报已生成', result, '出席记录', len(attendance), '白话佛法', len(reading))
    return result
