# -*- coding: utf-8 -*-
"""仅补录 2026-11-14 本地真实出席到历史表。先预览，输入 yes 才写入。"""
import hashlib
import json
from collections import Counter, defaultdict
from datetime import date
from pathlib import Path
import import_attendance_history as h
from db import get_conn

DAY = date(2026, 11, 14)
EXPECTED = 52
EXCEL = Path(r'C:\gyt_system\值班报表系统\attendance.xlsx')
VOLUNTEERS = Path(r'C:\gyt_system\值班报表系统\volunteers.xlsx')


def main():
    if not EXCEL.is_file() or not VOLUNTEERS.is_file():
        raise RuntimeError('找不到值班报表系统目录内的 attendance.xlsx 或 volunteers.xlsx')
    local_names = defaultdict(set)
    for _, r in h.read_rows(VOLUNTEERS, 'volunteers'):
        vid, name = h.norm_id(r.get('编号')), str(r.get('姓名') or '').strip()
        if vid and name:
            local_names[h.norm_name(name)].add(vid)
    rows = []
    seen = Counter()
    for line, r in h.read_rows(EXCEL, 'records'):
        try:
            day = h.parse_date(r.get('日期'))
        except ValueError:
            continue
        if day != DAY or not h.flag(r.get('签到')):
            continue
        name, role = str(r.get('姓名') or '').strip(), str(r.get('岗位') or '').strip()
        if not name or role not in h.ROLES:
            raise RuntimeError(f'Excel第{line}行姓名/岗位异常')
        ids = local_names[h.norm_name(name)]
        if len(ids) != 1:
            raise RuntimeError(f'Excel第{line}行姓名无法唯一匹配编号：{name} {ids}')
        local_id = next(iter(ids))
        cloud_id = h.ID_ALIASES.get(local_id, local_id)
        start, end = h.parse_time(r.get('开始时间')), h.parse_time(r.get('结束时间'))
        hours = h.parse_hours(r.get('时数'))
        remark = str(r.get('备注') or '').strip()
        if start is None or end is None or hours is None:
            raise RuntimeError(f'Excel第{line}行时间/时数缺失')
        signature = json.dumps([local_id, day.isoformat(), role, start.isoformat(), end.isoformat(), str(hours), remark], ensure_ascii=False, separators=(',', ':'))
        seen[signature] += 1
        key = hashlib.sha256(f'{signature}|{seen[signature]}'.encode('utf-8')).hexdigest()
        rows.append(dict(line=line, vid=cloud_id, name=name, role=role, start=start, end=end, hours=hours, remark=remark, key=key))
    print('读取：', EXCEL)
    print('11月14日已签到记录：', len(rows))
    if len(rows) != EXPECTED:
        raise RuntimeError(f'预期52笔，实际{len(rows)}笔；已停止')

    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute('SELECT id,name FROM public.volunteers')
            names = {h.norm_id(r['id']): str(r['name'] or '') for r in cur.fetchall()}
            for r in rows:
                cname = names.get(r['vid'])
                if cname is None or (h.norm_name(cname) != h.norm_name(r['name']) and (r['vid'], r['name'], cname) not in h.NAME_ALIASES):
                    raise RuntimeError(f'义工编号或姓名不一致：{r["name"]} / {r["vid"]} / {cname}')
                r['cloud_name'] = cname
            cur.execute('''SELECT volunteer_id,role,start_time,end_time,hours,source_row_key
                           FROM public.volunteer_attendance_history WHERE attendance_date=%s''', (DAY,))
            existing = cur.fetchall()
            keys = {r['source_row_key'] for r in existing}
            for r in rows:
                r['exists'] = r['key'] in keys
                if not r['exists'] and any(h.norm_id(x['volunteer_id']) == r['vid'] and x['role'] == r['role'] and x['start_time'] == r['start'] and x['end_time'] == r['end'] for x in existing):
                    raise RuntimeError(f'云端已有相同姓名岗位时段但指纹不同：{r["name"]}；停止人工核对')
            pending = [r for r in rows if not r['exists']]
            print('云端11月14日历史记录：', len(existing), '；待新增：', len(pending))
            for r in pending:
                print(f'待新增｜第{r["line"]}行｜{r["vid"]} {r["cloud_name"]}｜{r["role"]}｜{r["start"]}-{r["end"]}｜{r["hours"]}小时')
            if not pending:
                print('无须新增；未修改数据库。')
                return
            if input('确认以上名单无误，输入 yes 才正式补录：').strip() != 'yes':
                print('已取消；未修改数据库。')
                return
            # 事务内复查并锁定表，防止核对后数据变化。
            cur.execute('LOCK TABLE public.volunteer_attendance_history IN SHARE ROW EXCLUSIVE MODE')
            cur.execute('''SELECT volunteer_id,role,start_time,end_time,hours,source_row_key
                           FROM public.volunteer_attendance_history WHERE attendance_date=%s''', (DAY,))
            now = cur.fetchall()
            if now != existing:
                raise RuntimeError('确认期间云端11月14日历史记录发生变化，已取消并回滚')
            for r in pending:
                cur.execute('''INSERT INTO public.volunteer_attendance_history
                  (volunteer_id,volunteer_name,attendance_date,role,start_time,end_time,hours,source,source_row_key,remarks)
                  VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)''',
                  (r['vid'], r['cloud_name'], DAY, r['role'], r['start'], r['end'], r['hours'], h.SOURCE, r['key'], r['remark']))
                if cur.rowcount != 1:
                    raise RuntimeError('写入数量异常，已回滚')
            print('事务准备提交：', len(pending), '笔')
    print('提交完成。请在 Supabase 查询11月14日历史记录数。')

if __name__ == '__main__':
    main()
