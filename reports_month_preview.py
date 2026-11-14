# -*- coding: utf-8 -*-
"""独立云端月报核对版：只读 Supabase，绝不修改旧报表/数据库/网页。
运行：python reports_month_preview.py 2026 10
输出：cloud_report_preview/2026_10_月报_核对版.xlsx
注意：这是核对版，不是旧版式完全复刻，不启用网页正式下载。
"""
import argparse
import calendar
from collections import defaultdict, Counter
from datetime import date, datetime
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from reports_excel import prepare_monthly_inputs, ReportDataError
from reports_service import load_attendance_rows

ROLES = ['值班','卫生','佛台','供台','供花','供果','膳食','佛学班']
COLORS = {'值班':'D9EAF7','卫生':'FFF2CC','佛台':'E2F0D9','供台':'CCE5FF','供花':'FCE4D6','供果':'D9D2E9','膳食':'EAD1DC','佛学班':'F4CCCC'}

def tmin(s):
    if not s: return None
    s=str(s).strip()
    try:
        a=s.split(':'); h=int(a[0]); m=int(a[1]); sec=int(float(a[2])) if len(a)>2 else 0
        if 0<=h<=23 and 0<=m<=59 and 0<=sec<=59: return h*60+m+sec/60
    except (ValueError, IndexError): pass
    return None

def merged_hours(ranges):
    if not ranges:return 0.0
    ranges=sorted(ranges)
    out=[list(ranges[0])]
    for a,b in ranges[1:]:
        if a<=out[-1][1]:out[-1][1]=max(out[-1][1],b)
        else:out.append([a,b])
    return round(sum(b-a for a,b in out)/60,2)

def label_person(r):
    vid=r.get('_volunteer_id')
    if vid is None: raise ReportDataError(f"缺少内部ID：{r['日期']} {r['姓名']}")
    return str(vid)

def sheet(wb,name,headers,rows):
    ws=wb.create_sheet(name)
    ws.append(headers)
    for row in rows:ws.append(list(row))
    ws.freeze_panes='C2' if name.endswith('月表') else 'A2'
    ws.auto_filter.ref=ws.dimensions
    for c in ws[1]:
        c.fill=PatternFill('solid',fgColor='D9EAD3')
        c.font=Font(name='Microsoft YaHei',bold=True)
        c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    for col in ws.columns:
        letter=get_column_letter(col[0].column)
        mx=max((sum(2 if ord(ch)>127 else 1 for ch in str(c.value or '')) for c in list(col)[:250]),default=10)
        ws.column_dimensions[letter].width=min(38,max(12,mx+3))
    return ws

def generate(year,month):
    data=prepare_monthly_inputs(year,month)
    rows=data['attendance_rows']
    if not rows:raise ReportDataError('目标月份没有实际出席记录，停止生成空白报表')
    issues=[]
    byperson=defaultdict(list)
    for r in rows:
        key=label_person(r)
        if r['时数'] is None:issues.append(f"缺少时数：{r['日期']} {r['姓名']}")
        a,b=tmin(r['开始时间']),tmin(r['结束时间'])
        if a is None or b is None or b<a:issues.append(f"时间无效：{r['日期']} {r['姓名']} {r['开始时间']}-{r['结束时间']}")
        byperson[key].append(r)
    if issues:raise ReportDataError('发现时间/时数问题，停止生成：\n'+'\n'.join(issues[:25]))
    names={}
    for v in data['volunteer_rows']:
        if v.get('id') is not None:names[str(v['id'])]=str(v.get('name') or '').strip()
    # 不依姓名合并；同一内部ID显示现行姓名，旧姓名仅留在明细
    for key,items in byperson.items():
        if key not in names:raise ReportDataError(f'内部ID {key} 不在当前义工名单，停止生成')
    wb=Workbook(); wb.remove(wb.active)
    days=calendar.monthrange(year,month)[1]
    headers=['编号(内部ID)','姓名']+[str(d) for d in range(1,days+1)]+['出席天数','岗位记录数','总服务时数','实际出席时数']
    summary=[]
    for key,items in sorted(byperson.items(),key=lambda x:(names[x[0]],x[0])):
        daily=defaultdict(set); ranges=defaultdict(list)
        for r in items:
            d=r['日期']; daily[d.day].add(r['岗位']);ranges[d].append((tmin(r['开始时间']),tmin(r['结束时间'])))
        raw=sum(r['时数'] for r in items)
        real=sum(merged_hours(rs) for rs in ranges.values())
        summary.append([key,names[key]]+[' & '.join(sorted(daily.get(d,[]),key=lambda s: ROLES.index(s) if s in ROLES else 99)) for d in range(1,days+1)]+[len(daily),len(items),round(raw,2),round(real,2)])
    ws=sheet(wb,f'{month}月月表',headers,summary)
    for row in ws.iter_rows(min_row=2,min_col=3,max_col=2+days):
        for c in row:
            if c.value and c.value in COLORS:c.fill=PatternFill('solid',fgColor=COLORS[c.value])
    detail=[]
    for r in sorted(rows,key=lambda x:(x['日期'],x['姓名'],x['开始时间'])):
        detail.append([str(r['日期']),label_person(r),r['姓名'],r['岗位'],r['开始时间'],r['结束时间'],r['时数'],r['备注'],r.get('_data_source')])
    sheet(wb,'月明细',['日期','内部ID','姓名(记录原名)','岗位','开始时间','结束时间','时数','备注','来源'],detail)
    role_days=defaultdict(set); role_hours=defaultdict(float); role_records=Counter()
    for r in rows:
        role=r['岗位'];role_days[role].add((label_person(r),r['日期']));role_hours[role]+=r['时数'];role_records[role]+=1
    sheet(wb,'当月岗位统计',['岗位','出席人日(同人同日同岗位一次)','岗位记录笔数','岗位服务时数'],[[role,len(role_days[role]),role_records[role],round(role_hours[role],2)] for role in ROLES+sorted(set(role_days)-set(ROLES))])
    # reading 直接列出原始云端字段，不擅自推断次数、主题与义工ID关联
    reading=[]
    for r in data['reading_rows']:
        reading.append([str(r.get('date') or ''),str(r.get('name') or ''),str(r.get('identity') or ''),str(r.get('topic') or ''),str(r.get('session') or ''),str(r.get('time') or ''),str(r.get('id') or '')])
    sheet(wb,'白话佛法统计',['日期','姓名','身份','主题','场次','时间','记录ID（原始资料核对）'],reading)
    # 未值班仅以当前 status 判断；无法确认历史月份的在册状态，明确标注
    worked=set(byperson)
    notworked=[]
    for v in data['volunteer_rows']:
        vid=v.get('id'); status=str(v.get('status') or '').strip()
        if vid is not None and str(vid) not in worked and status in ('在册','active','Active'):
            notworked.append([str(vid),str(v.get('name') or ''),status,str(v.get('branch') or ''),'以当前在册状态筛选；历史状态待核实'])
    sheet(wb,'本月未值班',['内部ID','姓名','当前状态','分会','核对说明'],notworked)
    audit=sheet(wb,'核对说明',['核对项目','结果'],[
        ['报表性质','仅供核对，尚非正式原版格式'],['数据来源','Supabase 只读'],['月份',f'{year}-{month:02d}'],
        ['出席记录笔数',len(rows)],['出席义工人数',len(byperson)],['出席人日数',sum(r[-4] for r in summary)],
        ['岗位人日数',sum(map(len,role_days.values()))],['历史/正式切换','2026-12-01'],
        ['报名未签到','不参与统计'],['白话佛法','仅原始云端记录，统计映射待核实'],
        ['未值班名单','以当前在册名单筛选；未核对当月历史在册状态'],
        ['编号','使用数据库内部ID，正式义工编号映射尚待核对'],
        ['原版格式','尚未逐格复刻；请勿当作正式月报'],
    ])
    out=Path(__file__).resolve().parent/'cloud_report_preview';out.mkdir(exist_ok=True)
    path=out/f'{year}_{month:02d}_月报_核对版.xlsx';wb.save(path)
    print('只读核对版已生成：',path)
    print('出席记录：',len(rows),'义工人数：',len(byperson),'出席人日：',sum(r[-4] for r in summary))
    print('注意：网页正式下载仍关闭；本文件仅供核对。')

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('year',type=int);p.add_argument('month',type=int);a=p.parse_args()
    generate(a.year,a.month)
