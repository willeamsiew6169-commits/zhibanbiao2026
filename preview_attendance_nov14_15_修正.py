# -*- coding: utf-8 -*-
"""只读核对 2026-11-14 历史出席与 11-15 农舍佛学班；绝不写入数据库。
依赖项目现有 import_attendance_history.py；不运行其 main()。"""
from collections import Counter
from datetime import date
from pathlib import Path
import import_attendance_history as h
from db import get_conn

ROOT = Path(__file__).resolve().parent
EXCEL = ROOT / 'attendance.xlsx'
if not EXCEL.exists():
    EXCEL = Path(r'C:\gyt_system\值班报表系统\attendance.xlsx')
D14, D15 = date(2026, 11, 14), date(2026, 11, 15)


def main():
    if not EXCEL.is_file():
        raise SystemExit('找不到 attendance.xlsx；请放在此程序同目录，或确认值班报表系统目录')
    print('实际读取：', EXCEL)
    local = {D14: [], D15: []}
    errors = []
    for line, row in h.read_rows(EXCEL, 'records'):
        try:
            day = h.parse_date(row.get('日期'))
        except Exception:
            continue
        if day not in local or not h.flag(row.get('签到')):
            continue
        try:
            item = dict(line=line, day=day, name=str(row.get('姓名') or '').strip(),
                        role=str(row.get('岗位') or '').strip(), start=h.parse_time(row.get('开始时间')),
                        end=h.parse_time(row.get('结束时间')), hours=h.parse_hours(row.get('时数')),
                        remark=str(row.get('备注') or '').strip(), id=str(row.get('义工卡号') or '').strip())
            if not item['name'] or not item['role'] or not item['start'] or not item['end']:
                raise ValueError('姓名/岗位/起止时间缺失')
            local[day].append(item)
        except Exception as exc:
            errors.append(f'Excel第{line}行：{exc}')
    print('11/14 Excel 已签到：', len(local[D14]))
    print('11/15 Excel 已签到：', len(local[D15]))
    print('11/15 佛学班记录：')
    for x in local[D15]:
        if x['role'] == '佛学班':
            print(f"  第{x['line']}行 {x['name']} {x['start']}-{x['end']} {x['hours']}小时 备注:{x['remark']}")
    if errors:
        print('资料异常：', *errors, sep='\n  ')
        raise SystemExit('发现异常；未修改数据库')
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute('''SELECT volunteer_id,volunteer_name,attendance_date,role,start_time,end_time,hours,remarks
                           FROM public.volunteer_attendance_history WHERE attendance_date=%s''', (D14,))
            history = cur.fetchall()
            print('11/14 云端历史记录：', len(history))
            print('11/14 Excel记录：')
            for x in local[D14]:
                matches = [r for r in history if h.norm_name(r[1]) == h.norm_name(x['name']) and
                           str(r[3]).strip() == x['role'] and r[4] == x['start'] and r[5] == x['end']]
                status = '云端已有' if matches else '待进一步核对'
                print(f"  {status} 第{x['line']}行 {x['name']} {x['role']} {x['start']}-{x['end']} {x['hours']}小时")
            print('11/15 云端正式签到记录（仅只读查询）：')
            cur.execute('''SELECT column_name FROM information_schema.columns
                           WHERE table_schema='public' AND table_name='volunteer_attendance_logs'
                           ORDER BY ordinal_position''')
            columns = [r['column_name'] for r in cur.fetchall()]
            print('volunteer_attendance_logs 栏位：', ', '.join(columns))
            if 'attendance_date' in columns:
                cur.execute('SELECT * FROM public.volunteer_attendance_logs WHERE attendance_date=%s', (D15,))
                rows = cur.fetchall()
                print('11/15 云端日志笔数：', len(rows))
                for row in rows:
                    data = dict(zip(columns, row))
                    if any(str(v).strip() in {'刘铧忆', '许银铃'} for v in data.values() if v is not None):
                        print('  疑似相关云端记录：', data)
    print('预览完成；没有修改任何数据库记录。')

if __name__ == '__main__':
    main()
