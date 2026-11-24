import calendar
import json
import secrets
from datetime import date, timedelta
from zoneinfo import ZoneInfo
from datetime import datetime
from flask import request, session, redirect, render_template_string, abort, Response
from db import get_conn
from .flower_fruit_generator import make_text, text_from_confirmed_days, WEEKDAYS, cross_month_days, combined_text, MAX_GAP

HTML = '''<!doctype html><html lang="zh"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>供花供果排班</title><style>
*{box-sizing:border-box}body{font:18px Arial,sans-serif;background:#f5f1fc;color:#243047;padding:12px}main{max-width:900px;margin:auto;background:white;padding:20px;border-radius:16px}button,input,select,textarea{font:inherit;padding:10px;border:1px solid #ccd3df;border-radius:8px;max-width:100%}button{background:#187c71;color:white;cursor:pointer;margin:5px}label{display:inline-block;margin:6px}article{padding:12px;background:#f4f6fa;margin:10px 0;border-radius:10px}textarea{width:100%;min-height:180px}a{color:#245de6}.warn{color:#a43b16} .danger{background:#a43b16} </style><main><a href="/schedule/management/fixed">← 返回佛台及固定岗位管理</a><h1>🌸 供花供果排班</h1><form method="get"><label>年份 <input type="number" name="year" min="1901" max="2098" value="{{y}}" required></label><label>月份 <select name="month">{% for m in range(1,13) %}<option value="{{m}}" {% if m==mo %}selected{% endif %}>{{m}}月</option>{% endfor %}</select></label><button>查看已保存安排</button></form>
<p>{{message}}</p><p>云端本月已有 {{saved|length}} 个工作日。上一次日期来自已保存计划，并非实际签到，请核对。</p><form method="post"><input type="hidden" name="csrf" value="{{csrf}}"><input type="hidden" name="year" value="{{y}}"><input type="hidden" name="month" value="{{mo}}"><label>上一次供花供果日期 <input type="date" name="prev" value="{{prev or ''}}" max="{{maxprev}}"></label><button name="action" value="generate">⚡ 自动生成</button></form>
<form method="post" id="editor"><input type="hidden" name="csrf" value="{{csrf}}"><input type="hidden" name="year" value="{{y}}"><input type="hidden" name="month" value="{{mo}}"><input type="hidden" name="prev" value="{{prev or ''}}"><input type="hidden" name="baseline" value="{{baseline}}"><input type="hidden" name="cross" value="{{cross_key}}">{% for d,item in cross.items()|sort %}<article><b>跨月补入：{{d}} {{item[1]}}</b>（确认后写入上月，不覆盖原有安排）</article>{% endfor %}<div id="rows">{% for d,item in days.items()|sort %}<article><label>日期 <input name="day" type="number" min="1" max="{{last}}" value="{{d}}" required></label><span>{{weekday(d)}}</span><label>工作 <select name="job"><option {% if item[1]=='佛台' %}selected{% endif %}>佛台</option><option {% if item[1]=='佛台+供台' %}selected{% endif %}>佛台+供台</option></select></label><label>备注 <input name="remark" maxlength="200" value="{{item[2]}}"></label><button type="button" class="danger" onclick="this.closest('article').remove()">删除</button></article>{% endfor %}</div><button type="button" onclick="addRow()">➕ 新增日期</button><button name="action" value="preview">更新预览</button><button name="action" value="download">下载 TXT</button><p class="warn">保存会替换选定月份的安排，并仅新增上方列出的跨月补位；不会修改其他月份原有记录或签到记录。</p><label><input type="checkbox" name="confirm" value="yes"> 确认保存本月安排及跨月补位（包括本月删除的日期）</label><button name="action" value="save" onclick="return confirm('确定保存本月安排及跨月补位？')">💾 确认保存</button></form><h2>WhatsApp 文字</h2><p>修改后请先更新预览。</p><textarea id="output" readonly>{{output}}</textarea><button onclick="let t=document.getElementById('output');t.select();try{if(document.execCommand('copy')){this.textContent='已复制';return}}catch(e){} alert('请手动复制已选中的文字')">📋 复制 WhatsApp</button></main><script>function addRow(){let a=document.createElement('article');a.innerHTML=`<label>日期 <input name="day" type="number" min="1" max="{{last}}" required></label><label>工作 <select name="job"><option>佛台</option><option>佛台+供台</option></select></label><label>备注 <input name="remark" maxlength="200"></label><button type="button" class="danger" onclick="this.closest('article').remove()">删除</button>`;document.getElementById('rows').append(a)}</script></html>'''

def bounds(y,m):
    return date(y,m,1), date(y+1,1,1) if m==12 else date(y,m+1,1)

def load(cur,start,end):
    cur.execute('SELECT work_date,job,remark,icon FROM public.flower_fruit_work_days WHERE work_date >= %s AND work_date < %s ORDER BY work_date',(start,end))
    return {r['work_date'].day:(r['icon'] or '',r['job'],r['remark'] or '') for r in cur.fetchall()}

def fingerprint(days):
    import hashlib
    return hashlib.sha256(json.dumps(days,sort_keys=True,ensure_ascii=False).encode()).hexdigest()

def plan_fingerprint(days, prev):
    return fingerprint({'days': days, 'prev': str(prev or '')})


def cross_key_for(cross):
    return ','.join(d.isoformat() for d in sorted(cross))


def read_prev(cur, start):
    cur.execute('SELECT MAX(work_date) AS prev FROM public.flower_fruit_work_days WHERE work_date < %s',(start,))
    return cur.fetchone()['prev']


def validate_gaps(start, end, days, prev, cross):
    points = ([prev] if prev else []) + sorted(cross) + [start.replace(day=d) for d in sorted(days)] + [end]
    if any((b-a).days > MAX_GAP for a,b in zip(points,points[1:])):
        raise ValueError('安排间隔超过8天，请调整日期后再保存')


def parse_rows(form,last):
    ds,js,rs=(form.getlist(k) for k in ('day','job','remark'))
    if not len(ds)==len(js)==len(rs) or len(ds)>last: raise ValueError('日期资料不正确')
    result={}
    for d,j,r in zip(ds,js,rs):
        d=int(d)
        if not 1<=d<=last or d in result or j not in ('佛台','佛台+供台') or len(r)>200: raise ValueError('日期重复或内容无效')
        result[d]=('💧' if j=='佛台' else '👉',j,r.strip())
    return result

def register_flower_fruit_routes(bp):
    @bp.route('/schedule/flower-fruit',methods=['GET','POST'])
    def flower_fruit_schedule():
        if not session.get('schedule_login'): return redirect('/schedule/admin')
        today=datetime.now(ZoneInfo('Asia/Kuala_Lumpur')).date()
        try:
            y=int(request.values.get('year',today.year)); mo=int(request.values.get('month',today.month))
            if not 1901<=y<=2098: raise ValueError()
            start,end=bounds(y,mo); last=calendar.monthrange(y,mo)[1]
        except ValueError: abort(400,'年月无效')
        csrf=session.setdefault('flower_fruit_csrf',secrets.token_urlsafe(32))
        if request.method=='POST' and not secrets.compare_digest(csrf,request.form.get('csrf','')): abort(403)
        with get_conn() as conn:
            with conn.cursor() as cur:
                saved=load(cur,start,end)
                prev=read_prev(cur,start)
        db_prev=prev
        days=saved; baseline=plan_fingerprint(saved,db_prev); cross={}
        message='已保存安排' if saved else '本月尚无安排'
        if request.method=='POST':
            action=request.form.get('action')
            try:
                prev=date.fromisoformat(request.form['prev']) if request.form.get('prev') else None
                if prev and prev>=start: raise ValueError('上一次日期必须早于本月')
                if action=='generate':
                    if prev != db_prev:
                        raise ValueError('上一次日期与云端计划不一致，请先核对上月安排')
                    cross=cross_month_days(y,mo,db_prev,today)
                    _,days=make_text(y,mo,max(cross) if cross else prev)
                    validate_gaps(start,end,days,db_prev,cross)
                    message='自动生成成功，尚未保存。'
                else:
                    days=parse_rows(request.form,last)
                    baseline=request.form.get('baseline','')
                    if baseline != plan_fingerprint(saved,db_prev) or prev != db_prev:
                        raise ValueError('云端安排已变化，请重新打开本月再修改')
                    cross=cross_month_days(y,mo,db_prev,today) if request.form.get('cross') else {}
                    if request.form.get('cross','') != cross_key_for(cross):
                        raise ValueError('跨月安排已变化，请重新生成')
                    validate_gaps(start,end,days,db_prev,cross)
                    message='待确认预览，尚未保存。'
                    if action=='download':
                        return Response(combined_text(y,mo,days,cross),mimetype='text/plain',headers={'Content-Disposition':f'attachment; filename="{y}_{mo:02d}_flower_fruit.txt"'})
                    if action=='save':
                        if request.form.get('confirm')!='yes': raise ValueError('请先勾选确认保存')
                        if start<today.replace(day=1): raise ValueError('历史月份仅供查看，不允许覆盖')
                        with get_conn() as conn:
                            with conn.cursor() as cur:
                                # Always lock previous and selected month in chronological order.
                                previous=start-timedelta(days=1)
                                for yy,mm in ((previous.year,previous.month),(y,mo)):
                                    cur.execute('SELECT pg_advisory_xact_lock(%s,%s)',(yy,mm))
                                current=load(cur,start,end)
                                current_prev=read_prev(cur,start)
                                if plan_fingerprint(current,current_prev)!=baseline:
                                    raise ValueError('云端安排已变化，请重新打开本月再修改')
                                for d in set(current)|set(days):
                                    if date(y,mo,d)<today and current.get(d)!=days.get(d): raise ValueError('不能修改已经过去的日期')
                                for d,(icon,job,remark) in sorted(cross.items()):
                                    if d<today: raise ValueError('不能补入已经过去的日期')
                                    cur.execute('INSERT INTO public.flower_fruit_work_days (work_date,job,remark,icon) VALUES (%s,%s,%s,%s)',(d,job,remark,icon))
                                for d in set(current)|set(days):
                                    if current.get(d)==days.get(d): continue
                                    cur.execute('DELETE FROM public.flower_fruit_work_days WHERE work_date=%s',(date(y,mo,d),))
                                    if d in days:
                                        icon,job,remark=days[d]
                                        cur.execute('INSERT INTO public.flower_fruit_work_days (work_date,job,remark,icon) VALUES (%s,%s,%s,%s)',(date(y,mo,d),job,remark,icon))
                        saved=days; prev=max(cross) if cross else db_prev
                        baseline=plan_fingerprint(saved,prev)
                        message='✅ 保存成功（含跨月补位）' if cross else '✅ 保存成功'
                        output=combined_text(y,mo,days,cross)
                        cross={}
                    elif action!='preview': abort(400)
            except ValueError as e:
                message='⚠️ '+str(e)
                days=saved; prev=db_prev; cross={}; baseline=plan_fingerprint(saved,db_prev)
        return render_template_string(HTML,y=y,mo=mo,last=last,days=days,saved=saved,prev=prev,maxprev=start-timedelta(days=1),csrf=csrf,baseline=baseline,message=message,output=locals().get('output',combined_text(y,mo,days,cross)),cross=cross,cross_key=cross_key_for(cross),weekday=lambda d:WEEKDAYS[date(y,mo,d).weekday()])
