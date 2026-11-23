'''弘法义工出席月报／年报（只读），授权组别过滤在 SQL 层完成。'''
from collections import defaultdict
from datetime import date, datetime
from io import BytesIO
from zoneinfo import ZoneInfo
from flask import request, abort, redirect, url_for, send_file
from db import db_query

TZ = ZoneInfo('Asia/Kuala_Lumpur')


def register_attendance_reports(bp, actor, render):
    def context():
        who = actor()
        if not who:
            return None
        teams = db_query('SELECT id,name FROM hongfa_teams WHERE active=true ORDER BY id', fetchall=True) or []
        allowed = teams if who.get('secretary') else [t for t in teams if t['id'] in (who.get('team_ids') or [])]
        if not allowed:
            abort(403)
        try:
            year = int(request.args.get('year') or datetime.now(TZ).year)
            month = int(request.args.get('month') or 0)
            team_id = int(request.args.get('team_id') or 0)
        except (TypeError, ValueError):
            abort(400)
        if not 2020 <= year <= 2100 or not 0 <= month <= 12:
            abort(400)
        allowed_ids = [int(t['id']) for t in allowed]
        if team_id and team_id not in allowed_ids:
            abort(403)
        selected = [t for t in allowed if not team_id or int(t['id']) == team_id]
        first = date(year, month or 1, 1)
        last = date(year + 1, 1, 1) if month == 0 else date(year + (month == 12), month % 12 + 1, 1)
        return dict(year=year, month=month, team_id=team_id, teams=allowed, selected=selected,
                    ids=[int(t['id']) for t in selected], first=first, last=last)

    def query(c):
        # 始终在 SQL 中约束组别；不取电话、PIN 等无关敏感字段。
        events = db_query('''SELECT e.id,e.event_date,e.team_id,t.name AS team_name,
            COALESCE(l.name,'未设地点') AS location_name
            FROM hongfa_events e JOIN hongfa_teams t ON t.id=e.team_id
            LEFT JOIN hongfa_locations l ON l.id=e.location_id
            WHERE e.team_id=ANY(%s) AND e.event_date >= %s AND e.event_date < %s
              AND e.status IN ('published','completed')
            ORDER BY e.event_date,e.id''', (c['ids'],c['first'],c['last']), fetchall=True) or []
        rows = db_query('''SELECT a.id,a.event_id,a.volunteer_id,a.guest_name,a.check_in_at,a.check_out_at,
            a.checkout_method,e.event_date,e.team_id,t.name AS team_name,
            COALESCE(l.name,'未设地点') AS location_name
            FROM hongfa_attendance a JOIN hongfa_events e ON e.id=a.event_id
            JOIN hongfa_teams t ON t.id=e.team_id
            LEFT JOIN hongfa_locations l ON l.id=e.location_id
            WHERE e.team_id=ANY(%s) AND e.event_date >= %s AND e.event_date < %s
            ORDER BY e.event_date,e.id,a.check_in_at,a.id''',
            (c['ids'],c['first'],c['last']), fetchall=True) or []
        return events, rows

    def local(value):
        if not value:
            return None
        return value.astimezone(TZ) if value.tzinfo else value.replace(tzinfo=TZ)

    def duration(r):
        if not r['check_out_at']:
            return None
        seconds = (r['check_out_at'] - r['check_in_at']).total_seconds()
        return round(max(0,seconds)/3600,2)

    def build(c, events, rows):
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment
        from openpyxl.utils import get_column_letter
        wb = Workbook()
        detail = wb.active
        detail.title = '出席明细'
        detail.append(['日期','星期','弘法组','地点','义工编号','姓名／新人','签到时间','签退时间','服务时数','签退方式','活动编号'])
        weekdays = '一二三四五六日'
        for r in rows:
            detail.append([r['event_date'], '星期'+weekdays[r['event_date'].weekday()],r['team_name'],r['location_name'],
                           r['volunteer_id'] or '', r['guest_name'] or '',
                           local(r['check_in_at']).strftime('%H:%M:%S') if r['check_in_at'] else '',
                           local(r['check_out_at']).strftime('%H:%M:%S') if r['check_out_at'] else '未签退',
                           duration(r),r['checkout_method'] or '',r['event_id']])
            detail.cell(detail.max_row,1).number_format='yyyy-mm-dd'
        per_person = {}
        for r in rows:
            key=(r['team_id'],r['volunteer_id'] or '新人:'+str(r['guest_name'] or r['id']))
            if key not in per_person:
                per_person[key]=dict(team=r['team_name'],number=r['volunteer_id'] or '',name=r['guest_name'] or '',count=0,hours=0,last=None,unfinished=0)
            v=per_person[key];v['count']+=1;v['last']=max(v['last'] or r['event_date'],r['event_date'])
            h=duration(r)
            if h is None:v['unfinished']+=1
            else:v['hours']+=h
        people=wb.create_sheet('义工出席统计')
        people.append(['弘法组','义工编号','姓名／新人','出席次数','已签退服务时数','未签退次数','最后出席日期'])
        for v in sorted(per_person.values(),key=lambda x:(x['team'],x['number'],x['name'])):
            people.append([v['team'],v['number'],v['name'],v['count'],round(v['hours'],2),v['unfinished'],v['last']])
            people.cell(people.max_row,7).number_format='yyyy-mm-dd'
        totals=defaultdict(lambda:dict(events=0,attendances=0,people=set(),hours=0,unfinished=0))
        for e in events:totals[e['team_name']]['events']+=1
        for r in rows:
            v=totals[r['team_name']];v['attendances']+=1
            v['people'].add(r['volunteer_id'] or '新人:'+str(r['guest_name'] or r['id']))
            h=duration(r)
            if h is None:v['unfinished']+=1
            else:v['hours']+=h
        summary=wb.create_sheet('各组弘法统计')
        summary.append(['弘法组','活动次数','出席人次','不重复人数','已签退服务时数','未签退人次'])
        for t in c['selected']:
            v=totals[t['name']]
            summary.append([t['name'],v['events'],v['attendances'],len(v['people']),round(v['hours'],2),v['unfinished']])
        from .hongfa_excel_style import polish
        for ws in wb:
            polish(ws,freeze='F2' if ws.title=='出席明细' else ('D2' if ws.title=='义工出席统计' else 'B2'))
        stream=BytesIO();wb.save(stream);stream.seek(0)
        return stream

    @bp.route('/reports/attendance')
    def attendance_reports():
        c=context()
        if c is None:return redirect(url_for('hongfa.login'))
        events,rows=query(c)
        totals=defaultdict(lambda:dict(events=0,attendance=0,hours=0,unfinished=0))
        for e in events:totals[e['team_name']]['events']+=1
        for r in rows:
            v=totals[r['team_name']];v['attendance']+=1
            h=duration(r)
            if h is None:v['unfinished']+=1
            else:v['hours']+=h
        for v in totals.values():v['hours']=round(v['hours'],2)
        return render('''<div class="card"><h2>👥 弘法义工出席报表</h2>
          <p><a href="{{url_for('hongfa.report')}}">⬅ 返回弘法报表</a></p>
          <form method="get"><label>年份</label><input name="year" type="number" min="2020" max="2100" value="{{c.year}}" required>
          <label>月份</label><select name="month"><option value="0">全年</option>{% for m in range(1,13) %}<option value="{{m}}" {% if c.month==m %}selected{% endif %}>{{m}}月</option>{% endfor %}</select>
          <label>弘法组</label><select name="team_id"><option value="0">全部获授权组别</option>{% for t in c.teams %}<option value="{{t.id}}" {% if c.team_id==t.id %}selected{% endif %}>{{t.name}}</option>{% endfor %}</select>
          <button>🔎 查询</button></form>
          <p><a class="btn" href="{{url_for('hongfa.attendance_reports_excel',year=c.year,month=c.month,team_id=c.team_id)}}">📥 下载 Excel（{{'全年' if not c.month else c.month|string+'月'}}）</a></p>
          <p>活动 {{events|length}} 场｜出席 {{rows|length}} 人次（服务时数只统计已签退记录）</p>
          {% for t in c.selected %}{% set v=totals.get(t.name) %}<p><b>{{t.name}}</b>｜活动 {{v.events if v else 0}} 场｜出席 {{v.attendance if v else 0}} 人次｜已签退 {{v.hours if v else 0}} 小时｜未签退 {{v.unfinished if v else 0}} 人次</p>{% endfor %}
          <p class="hf-muted">Excel 包含出席明细、义工出席统计、各组弘法统计；未签退时数不估算。义工姓名如未存入弘法出席记录，报表显示义工编号。</p>
          </div>''',c=c,events=events,rows=rows,totals=totals)

    @bp.route('/reports/attendance.xlsx')
    def attendance_reports_excel():
        c=context()
        if c is None:return redirect(url_for('hongfa.login'))
        events,rows=query(c)
        stream=build(c,events,rows)
        period=f'{c["year"]}' + (f'-{c["month"]:02d}' if c['month'] else '-全年')
        return send_file(stream,as_attachment=True,download_name=f'弘法义工出席报表_{period}.xlsx',
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
