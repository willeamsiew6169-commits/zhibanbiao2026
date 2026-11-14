# -*- coding: utf-8 -*-
"""一次性历史出席精准同步：仅处理已核对的 2026-10-10 / 2026-11-10 差异。
先只读核对、备份，输入 yes 后事务写入；任何偏差停止并回滚。
放在 audit_attendance_sync.py 和 attendance_audit_helpers.py 同目录。
"""
from __future__ import annotations
import csv
import json
from collections import Counter, defaultdict
from datetime import datetime
from decimal import Decimal
from pathlib import Path

import attendance_audit_helpers as h
from audit_attendance_sync import fmt

ROOT = Path(__file__).resolve().parent
DIFF = ROOT / 'attendance_sync_differences.csv'
EXPECTED_EXCEL_ONLY = 32
EXPECTED_CLOUD_ONLY = 17
FIELDS = ('volunteer_id', 'volunteer_name', 'attendance_date', 'role', 'start_time', 'end_time', 'hours', 'source', 'source_row_key', 'remarks')


def safe_value(x):
    if isinstance(x, Decimal):
        return str(x)
    return x.isoformat() if hasattr(x, 'isoformat') else x


def load_diff():
    if not DIFF.is_file():
        raise RuntimeError('找不到 attendance_sync_differences.csv，请先运行 audit_attendance_sync.py')
    with DIFF.open(encoding='utf-8-sig', newline='') as f:
        rows = list(csv.DictReader(f))
    new = [r for r in rows if r['类型'] == 'Excel有/云端无']
    old = [r for r in rows if r['类型'] == '云端有/Excel无']
    if len(new) != EXPECTED_EXCEL_ONLY or len(old) != EXPECTED_CLOUD_ONLY or len(rows) != 49:
        raise RuntimeError(f'差异数量不符：Excel {len(new)}，云端 {len(old)}；请重新核对，禁止同步')
    if len({r['指纹'] for r in rows}) != len(rows):
        raise RuntimeError('差异指纹重复，停止')
    return new, old


def as_tuple(r):
    return (r['编号'], r['日期'], r['岗位'], r['开始'], r['结束'], r['时数'], r['备注'])


def match_changes(new, old):
    remaining_new = {r['指纹']: r for r in new}
    remaining_old = {r['指纹']: r for r in old}
    pairs = []
    # 只允许日期由 10/10 更正为 11/10，其他字段完全相同。
    for o in old:
        if o['日期'] != '2026-10-10' or o['姓名'] == '黄丽玲':
            continue
        candidates = [n for n in remaining_new.values()
                      if n['日期'] == '2026-11-10'
                      and (n['编号'], n['岗位'], n['开始'], n['结束'], n['时数'], n['备注'])
                      == (o['编号'], o['岗位'], o['开始'], o['结束'], o['时数'], o['备注'])]
        if len(candidates) != 1:
            raise RuntimeError(f'日期更正无法唯一配对：{o["姓名"]} {o["岗位"]} {o["开始"]}-{o["结束"]}，候选{len(candidates)}')
        n = candidates[0]
        pairs.append((o, n, '日期更正'))
        del remaining_new[n['指纹']]
        del remaining_old[o['指纹']]
    # 黄丽玲的下午时段：原 14-18，现 16-18；上午 10-13 已存在，不改动。
    huang_old = [o for o in remaining_old.values() if o['姓名'] == '黄丽玲' and o['日期'] == '2026-10-10'
                 and o['岗位'] == '值班' and o['开始'] == '14:00:00' and o['结束'] == '18:00:00' and o['时数'] == '4.00']
    huang_new = [n for n in remaining_new.values() if n['姓名'] == '黄丽玲' and n['日期'] == '2026-10-10'
                 and n['岗位'] == '值班' and n['开始'] == '16:00:00' and n['结束'] == '18:00:00' and n['时数'] == '2.00']
    if len(huang_old) != 1 or len(huang_new) != 1:
        raise RuntimeError('黄丽玲下午时间差异不符合已确认内容，停止')
    o, n = huang_old[0], huang_new[0]
    if (o['编号'], o['备注']) != (n['编号'], n['备注']):
        raise RuntimeError('黄丽玲编号或备注不一致')
    pairs.append((o, n, '下午时间更正'))
    del remaining_new[n['指纹']]
    del remaining_old[o['指纹']]
    if len(pairs) != 17 or remaining_old or len(remaining_new) != 15:
        raise RuntimeError(f'预期修改17、新增15；实际修改{len(pairs)}、新增{len(remaining_new)}、未匹配旧记录{len(remaining_old)}，停止')
    additions = list(remaining_new.values())
    meal = [n for n in additions if n['日期'] == '2026-10-10' and n['岗位'] == '值班'
            and n['开始'] == '08:00:00' and n['结束'] == '14:00:00' and n['时数'] == '6.00']
    anna = [n for n in additions if n['姓名'] == 'Anna' and n['日期'] == '2026-11-10'
            and n['岗位'] == '值班' and n['开始'] == '10:00:00' and n['结束'] == '14:00:00' and n['时数'] == '4.00']
    if len(meal) != 14 or len(anna) != 1 or len({r['编号'] for r in meal}) != 14:
        raise RuntimeError('新增15笔未能明确归类为14笔堂食+1笔Anna值班，停止')
    return pairs, additions


def db_snapshot(cur):
    cur.execute('''SELECT volunteer_id,volunteer_name,attendance_date,role,start_time,end_time,
                   hours,source,source_row_key,remarks
                   FROM public.volunteer_attendance_history WHERE source=%s''', (h.SOURCE,))
    rows = cur.fetchall()
    return {r['source_row_key']: dict(r) for r in rows}


def assert_db_matches(cloud, old, new):
    old_keys = {r['指纹'] for r in old}
    new_keys = {r['指纹'] for r in new}
    if len(cloud) != 6537:
        raise RuntimeError(f'云端历史记录数由6537变成{len(cloud)}，停止，重新对账')
    if not old_keys <= cloud.keys() or new_keys & cloud.keys():
        raise RuntimeError('云端差异指纹已发生变化，停止，重新对账')
    for r in old:
        c = cloud[r['指纹']]
        if (str(c['volunteer_id']).upper(), fmt(c['attendance_date']), c['role'],
            fmt(c['start_time']), fmt(c['end_time']), f'{c["hours"]:.2f}' if c['hours'] is not None else '',
            c['remarks'] or '') != as_tuple(r):
            raise RuntimeError(f'云端记录内容变化：{r["姓名"]} {r["日期"]} {r["岗位"]}')


def main():
    new, old = load_diff()
    pairs, additions = match_changes(new, old)
    print('=== 历史出席精准同步（预览）===')
    print('更新旧记录：17笔（16笔日期，1笔黄丽玲下午时间）')
    print('新增记录：15笔（14笔10月10日堂食，1笔Anna 11月10日值班）')
    print('黄丽玲上午10:00-13:00原记录保留不动。')
    for o, n, reason in pairs:
        print(f'修改｜{o["姓名"]} {reason}｜{o["日期"]} {o["开始"]}-{o["结束"]} → {n["日期"]} {n["开始"]}-{n["结束"]}')
    for n in additions:
        print(f'新增｜{n["姓名"]} {n["日期"]} {n["岗位"]} {n["开始"]}-{n["结束"]}')
    # 在写入前备份整张 local_excel 历史记录；文件保存在项目目录，不上传网络。
    with h.get_conn() as conn:
        with conn.cursor() as cur:
            cloud = db_snapshot(cur)
    assert_db_matches(cloud, old, new)
    backup = ROOT / f'attendance_history_backup_{datetime.now():%Y%m%d_%H%M%S}.json'
    with backup.open('x', encoding='utf-8') as f:
        json.dump({'source': h.SOURCE, 'count': len(cloud),
                   'records': [{k: safe_value(v) for k, v in r.items()} for r in cloud.values()]},
                  f, ensure_ascii=False, indent=2)
    print(f'已备份云端 {len(cloud)} 笔记录：{backup}')
    print('⚠️ 仅在以上17笔修改和15笔新增全部正确时继续。')
    if input('确认正式同步 Supabase？请输入 yes：').strip() != 'yes':
        print('已取消，云端未修改。')
        return
    # 单一事务：锁表防止并发插入；复查后执行，失败全部回滚。
    with h.get_conn() as conn:
        try:
            with conn.cursor() as cur:
                cur.execute('LOCK TABLE public.volunteer_attendance_history IN SHARE ROW EXCLUSIVE MODE')
                live = db_snapshot(cur)
                if live != cloud:
                    raise RuntimeError('确认期间云端数据有变化，已回滚；请重新对账')
                assert_db_matches(live, old, new)
                for o, n, reason in pairs:
                    cur.execute('''UPDATE public.volunteer_attendance_history
                        SET attendance_date=%s, start_time=%s, end_time=%s, hours=%s, source_row_key=%s
                        WHERE source=%s AND source_row_key=%s
                          AND volunteer_id=%s AND role=%s''',
                        (n['日期'], n['开始'] or None, n['结束'] or None, n['时数'] or None,
                         n['指纹'], h.SOURCE, o['指纹'], o['编号'], o['岗位']))
                    if cur.rowcount != 1:
                        raise RuntimeError(f'更新笔数异常：{o["姓名"]}')
                # 新增行的姓名使用云端 volunteers 正式姓名，避免名字字形不一致。
                cur.execute('SELECT id,name FROM public.volunteers')
                names = {str(r['id']).upper(): r['name'] for r in cur.fetchall()}
                for n in additions:
                    if n['编号'] not in names:
                        raise RuntimeError(f'义工编号不存在：{n["编号"]}')
                    cur.execute('''INSERT INTO public.volunteer_attendance_history
                        (volunteer_id, volunteer_name, attendance_date, role, start_time, end_time,
                         hours, source, source_row_key, remarks)
                         VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)''',
                        (n['编号'], names[n['编号']], n['日期'], n['岗位'], n['开始'] or None,
                         n['结束'] or None, n['时数'] or None, h.SOURCE, n['指纹'], n['备注']))
                    if cur.rowcount != 1:
                        raise RuntimeError(f'新增笔数异常：{n["姓名"]}')
                cur.execute('SELECT COUNT(*) AS n FROM public.volunteer_attendance_history WHERE source=%s', (h.SOURCE,))
                count = cur.fetchone()['n']
                if count != 6552:
                    raise RuntimeError(f'最终数量不是6552（实际{count}），已回滚')
            conn.commit()
        except Exception:
            conn.rollback()
            raise
    print('✅ 已完成：修改17笔、新增15笔，local_excel历史记录合计6552笔。')
    print('请再运行 audit_attendance_sync.py，预期差异为 0 / 0。')


if __name__ == '__main__':
    try:
        main()
    except Exception as exc:
        print(f'❌ 已停止：{exc}')
        raise
