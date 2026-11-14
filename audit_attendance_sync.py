# -*- coding: utf-8 -*-
"""只读全表对账：不更新、不插入、不删除 Supabase。"""
import argparse, csv, hashlib, json
from collections import Counter, defaultdict
from pathlib import Path
from datetime import date
from decimal import Decimal
import attendance_audit_helpers as h


def fmt(v):
    if v is None: return ''
    if hasattr(v, 'isoformat'): return v.isoformat()
    return str(v)


def main():
    ap=argparse.ArgumentParser()
    ap.add_argument('--excel', default=str(h.ATTENDANCE), help='修改后的 attendance.xlsx 完整路径')
    ap.add_argument('--volunteers', default=str(h.VOLUNTEERS), help='volunteers.xlsx 完整路径')
    args=ap.parse_args()
    attendance=Path(args.excel); volunteers=Path(args.volunteers)
    if not attendance.is_file() or not volunteers.is_file():
        raise FileNotFoundError(f'找不到 Excel：{attendance} / {volunteers}')
    local_by_name=defaultdict(set)
    for _,r in h.read_rows(volunteers,'volunteers'):
        vid=h.norm_id(r.get('编号')); name=h.norm_name(r.get('姓名'))
        if vid and name: local_by_name[name].add(vid)
    with h.get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute('SELECT id,name FROM public.volunteers')
            cloud_by_id={h.norm_id(r['id']):str(r['name'] or '').strip() for r in cur.fetchall()}
            cur.execute('''SELECT volunteer_id, volunteer_name, attendance_date, role,
                           start_time,end_time,hours,remarks,source_row_key
                           FROM public.volunteer_attendance_history WHERE source=%s''',(h.SOURCE,))
            cloud=cur.fetchall()
    local=[]; issues=[]; seen=Counter()
    for no,r in h.read_rows(attendance,'records'):
        if not h.flag(r.get('签到')): continue
        try:
            day=h.parse_date(r.get('日期'))
            if day>h.CUTOFF: continue
            name=str(r.get('姓名') or '').strip(); role=str(r.get('岗位') or '').strip()
            ids=local_by_name[h.norm_name(name)]
            if len(ids)!=1 or role not in h.ROLES: raise ValueError('姓名编号不唯一或岗位无效')
            local_id=next(iter(ids)); cloud_id=h.ID_ALIASES.get(local_id,local_id)
            if cloud_id not in cloud_by_id: raise ValueError(f'云端编号不存在：{cloud_id}')
            cloud_name=cloud_by_id[cloud_id]
            if h.norm_name(name)!=h.norm_name(cloud_name) and (cloud_id,name,cloud_name) not in h.NAME_ALIASES:
                raise ValueError(f'云端姓名不符：{cloud_name}')
            start=h.parse_time(r.get('开始时间')); end=h.parse_time(r.get('结束时间'))
            hours=h.parse_hours(r.get('时数')); remarks=str(r.get('备注') or '').strip()
            sig=json.dumps([local_id,day.isoformat(),role,fmt(start) or None,fmt(end) or None,str(hours) if hours is not None else None,remarks],ensure_ascii=False,separators=(',',':'))
            seen[sig]+=1
            key=hashlib.sha256(f'{sig}|{seen[sig]}'.encode()).hexdigest()
            local.append(dict(key=key,excel_row=no,volunteer_id=cloud_id,name=name,day=day,role=role,start=start,end=end,hours=hours,remarks=remarks))
        except (ValueError,TypeError) as e: issues.append(f'Excel第{no}行：{e}')
    if issues:
        print('❌ 发现资料异常，停止对账：'); print('\n'.join(issues[:50])); return
    by_key={r['source_row_key']:r for r in cloud}
    local_keys={r['key'] for r in local}
    missing=[r for r in local if r['key'] not in by_key]
    extra=[r for r in cloud if r['source_row_key'] not in local_keys]
    # 仅作为人工检查线索，绝不自动将两条记录认定为同一笔
    extra_by_person_role=defaultdict(list)
    for r in extra: extra_by_person_role[(h.norm_id(r['volunteer_id']),r['role'])].append(r)
    out=Path(__file__).resolve().parent/'attendance_sync_differences.csv'
    cols=['类型','Excel行','编号','姓名','日期','岗位','开始','结束','时数','备注','可能对应旧记录（仅供核对）','指纹']
    with out.open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.writer(f);w.writerow(cols)
        for r in missing:
            candidates=extra_by_person_role[(r['volunteer_id'],r['role'])]
            candidates=sorted(candidates,key=lambda c:abs((c['attendance_date']-r['day']).days))[:3]
            hint='; '.join(f"{fmt(c['attendance_date'])} {fmt(c['start_time'])}-{fmt(c['end_time'])} {fmt(c['hours'])}h" for c in candidates)
            w.writerow(['Excel有/云端无',r['excel_row'],r['volunteer_id'],r['name'],fmt(r['day']),r['role'],fmt(r['start']),fmt(r['end']),fmt(r['hours']),r['remarks'],hint,r['key']])
        for r in extra:
            w.writerow(['云端有/Excel无','',r['volunteer_id'],r['volunteer_name'],fmt(r['attendance_date']),r['role'],fmt(r['start_time']),fmt(r['end_time']),fmt(r['hours']),r['remarks'],'',r['source_row_key']])
    print('=== 历史出席全表只读对账 ===')
    print('Excel路径：',attendance)
    print('Excel有效记录：',len(local),'云端 local_excel 记录：',len(cloud))
    print('Excel有/云端无：',len(missing),'云端有/Excel无：',len(extra))
    print('差异报告：',out)
    print('⚠️ 这只是内容差异，不能直接视为新增/删除；没有执行任何数据库写入。')
    print('⚠️ 其他 source 的云端记录不在本次对账范围。')

if __name__=='__main__':
    main()
