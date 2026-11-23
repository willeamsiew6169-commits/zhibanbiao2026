"""弘法书籍月报／年报：只读数据库，按登录人员的组别授权。"""
from collections import defaultdict
from datetime import date, datetime
from io import BytesIO
from zoneinfo import ZoneInfo
from flask import request, abort, redirect, url_for, send_file
from db import db_query

TZ = ZoneInfo('Asia/Kuala_Lumpur')


def register_book_reports(bp, actor, render):
    def load():
        who = actor()
        if not who:
            return None, None
        all_teams = db_query('SELECT id,name FROM hongfa_teams WHERE active=true ORDER BY id', fetchall=True) or []
        permitted = all_teams if who.get('secretary') else [t for t in all_teams if t['id'] in who.get('team_ids', [])]
        if not permitted:
            abort(403)
        try:
            year = int(request.args.get('year') or datetime.now(TZ).year)
            month = int(request.args.get('month') or 0)
            team_id = int(request.args.get('team_id') or 0)
        except ValueError:
            abort(400)
        if year < 2020 or year > 2100 or month not in range(13):
            abort(400)
        ids = [int(t['id']) for t in permitted]
        if team_id and team_id not in ids:
            abort(403)
        selected = [t for t in permitted if not team_id or int(t['id']) == team_id]
        start = date(year, month or 1, 1)
        end = date(year + 1, 1, 1) if not month else date(year + (month == 12), month % 12 + 1, 1)
        return who, dict(year=year, month=month, team_id=team_id, teams=permitted, selected=selected, ids=[int(t['id']) for t in selected], start=start, end=end)

    def query(data):
        ids, start, end = data['ids'], data['start'], data['end']
        # 只读实际结缘流水。日期、地点来自绑定的弘法活动。
        details = db_query('''SELECT m.id,m.team_id,t.name AS team_name,e.id AS event_id,e.event_date,
            COALESCE(l.name,'未设地点') AS location_name,i.id AS item_id,i.title,
            -m.quantity AS quantity,m.created_by,m.created_at
            FROM hongfa_stock_movements m
            JOIN hongfa_events e ON e.id=m.event_id
            JOIN hongfa_teams t ON t.id=m.team_id
            LEFT JOIN hongfa_locations l ON l.id=e.location_id
            JOIN hongfa_items i ON i.id=m.item_id
            WHERE m.movement_type='distribution' AND m.team_id=ANY(%s)
              AND e.team_id=m.team_id AND e.event_date >= %s AND e.event_date < %s
            ORDER BY e.event_date,e.id,i.title,m.id''', (ids,start,end),fetchall=True) or []
        # 库存流水按登记时间归属月份/年份；点货调整不冒充补货。
        movements = db_query('''SELECT m.team_id,t.name AS team_name,i.id AS item_id,i.title,
            m.quantity,m.movement_type,m.created_at
            FROM hongfa_stock_movements m
            JOIN hongfa_teams t ON t.id=m.team_id
            JOIN hongfa_items i ON i.id=m.item_id
            WHERE m.team_id=ANY(%s) AND m.created_at < %s
            ORDER BY m.created_at,m.id''', (ids,end),fetchall=True) or []
        return details, movements

    def summary(data, movements):
        result = {}
        for m in movements:
            key=(m['team_id'],m['item_id'])
            if key not in result:
                result[key]=dict(team=m['team_name'],title=m['title'],opening=0,receipt=0,distribution=0,adjustment=0,closing=0)
            r=result[key]; qty=int(m['quantity']); typ=m['movement_type']
            if m['created_at'].date() < data['start']:
                r['opening']+=qty
            else:
                if typ=='receipt': r['receipt']+=qty
                elif typ=='distribution': r['distribution']+=-qty
                else: r['adjustment']+=qty
            r['closing']+=qty
        return sorted(result.values(),key=lambda r:(r['team'],r['title']))

    @bp.route('/books/report')
    def book_report():
        who, data = load()
        if who is None: return redirect(url_for('hongfa.login'))
        details, movements = query(data)
        totals=summary(data,movements)
        return render('''<div class="card"><h2>📊 弘法书籍报表</h2>
        <p><a href="/hongfa/book-menu">⬅ 返回书籍管理</a></p>
        <form method="get"><label>年份</label><select name="year">{% for y in years %}<option value="{{y}}" {% if y==d.year %}selected{% endif %}>{{y}}</option>{% endfor %}</select>
        <label>月份</label><select name="month"><option value="0" {% if not d.month %}selected{% endif %}>全年</option>{% for m in range(1,13) %}<option value="{{m}}" {% if d.month==m %}selected{% endif %}>{{m}}月</option>{% endfor %}</select>
        <label>弘法组</label><select name="team_id"><option value="0">全部授权组别</option>{% for t in d.teams %}<option value="{{t.id}}" {% if d.team_id==t.id %}selected{% endif %}>{{t.name}}</option>{% endfor %}</select>
        <button>🔎 查看报表</button></form>
        <p><a class="btn" href="{{download_url}}" style="font-size:20px">📥 下载 Excel（{{'全年' if not d.month else d.month|string+'月'}}）</a></p>
        <h3>📖 结缘明细（{{details|length}}笔）</h3>
        <div style="overflow-x:auto"><table><tr><th>日期</th><th>组别</th><th>地点</th><th>书籍</th><th>数量</th><th>登记人编号</th></tr>
        {% for r in details %}<tr><td>{{r.event_date}}</td><td>{{r.team_name}}</td><td>{{r.location_name}}</td><td>{{r.title}}</td><td>{{r.quantity}}</td><td>{{r.created_by}}</td></tr>{% else %}<tr><td colspan="6">暂无结缘记录</td></tr>{% endfor %}</table></div>
        <h3>📦 库存统计</h3><div style="overflow-x:auto"><table><tr><th>组别</th><th>书籍</th><th>期初</th><th>领取/补货</th><th>结缘</th><th>点货调整</th><th>期末</th></tr>
        {% for r in totals %}<tr><td>{{r.team}}</td><td>{{r.title}}</td><td>{{r.opening}}</td><td>{{r.receipt}}</td><td>{{r.distribution}}</td><td>{{r.adjustment}}</td><td>{{r.closing}}</td></tr>{% else %}<tr><td colspan="7">暂无库存记录</td></tr>{% endfor %}</table></div></div>''', d=data, years=range(datetime.now(TZ).year+1,2023,-1), details=details, totals=totals,
            download_url=url_for('hongfa.book_report_download',year=data['year'],month=data['month'],team_id=data['team_id']))

    @bp.route('/books/report/download')
    def book_report_download():
        who,data=load()
        if who is None: return redirect(url_for('hongfa.login'))
        details,movements=query(data)
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment
        from openpyxl.utils import get_column_letter
        wb=Workbook(); wb.remove(wb.active)
        from .hongfa_excel_style import polish
        def sheet(name,headers,rows):
            ws=wb.create_sheet(name)
            ws.append(headers)
            for row in rows: ws.append(list(row))
            polish(ws,freeze='B2')
            return ws
        sheet('结缘明细',['日期','星期','弘法组','地点','书籍','结缘数量','登记人编号','登记时间'],[
            (r['event_date'].isoformat(),'一二三四五六日'[r['event_date'].weekday()],r['team_name'],r['location_name'],r['title'],int(r['quantity']),str(r['created_by']),str(r['created_at'])) for r in details])
        for team in data['selected']:
            group=[r for r in details if r['team_id']==team['id']]
            events=sorted({(r['event_date'],r['event_id'],r['location_name']) for r in group})
            books=sorted({r['title'] for r in group})
            matrix=defaultdict(int)
            for r in group: matrix[(r['title'],r['event_id'])]+=int(r['quantity'])
            ws=wb.create_sheet((str(team['name'])+'结缘')[:31])
            ws.append(['书籍名称']+[x[0].strftime('%d/%m/%Y') for x in events]+['合计'])
            ws.append(['弘法地点']+[x[2] for x in events]+[''])
            for title in books:
                values=[matrix[(title,x[1])] for x in events]
                ws.append([title]+values+[sum(values)])
            # 第一行日期、第二行地点，固定两行及书名列。
            # 自动筛选使用日期标题行；地点行保留在表格内。
            polish(ws,header_rows=2,freeze='B3',fixed_widths={'A':34,**{get_column_letter(j):18 for j in range(2,len(events)+3)}})
        sheet('库存月年汇总',['弘法组','书籍','期初库存','领取补货','结缘扣减','点货调整','期末库存'],[
            (r['team'],r['title'],r['opening'],r['receipt'],r['distribution'],r['adjustment'],r['closing']) for r in summary(data,movements)])
        output=BytesIO();wb.save(output);output.seek(0)
        period=f"{data['year']}-{data['month']:02d}" if data['month'] else str(data['year'])
        return send_file(output,as_attachment=True,download_name=f'hongfa_books_{period}.xlsx',mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
