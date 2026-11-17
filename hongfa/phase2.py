"""弘法活动安排批量扩展：组长提交、秘书审核、地点、周表、现场码设置。
注意：本批不开放实际签到及库存扣减。
"""
from datetime import date, datetime, timedelta
from functools import wraps
from urllib.parse import urlparse
from zoneinfo import ZoneInfo
import hashlib
import hmac
import secrets
from flask import abort, redirect, request, session, url_for, current_app
from db import db_query

TZ = ZoneInfo('Asia/Kuala_Lumpur')


def register_phase2(bp, actor, render, csrf_token, valid_csrf):
    def guard(secretary=False):
        def deco(fn):
            @wraps(fn)
            def wrapped(*args, **kwargs):
                who = actor()
                if not who:
                    return redirect(url_for('hongfa.login'))
                if secretary and not who['secretary']:
                    abort(403)
                if request.method == 'POST' and not valid_csrf():
                    abort(400)
                return fn(who, *args, **kwargs)
            return wrapped
        return deco

    def page(body, **kw):
        nav = '''<style>
        .hf-nav-buttons{display:flex;flex-wrap:wrap;gap:10px;margin:12px 0 20px}
        .hf-nav-buttons a{display:inline-flex;align-items:center;justify-content:center;
          min-height:50px;padding:10px 18px;box-sizing:border-box;border-radius:10px;
          background:#16854e;color:#fff!important;text-decoration:none!important;
          font-size:19px;font-weight:700;line-height:1.35;text-align:center}
        .hf-nav-buttons a:hover,.hf-nav-buttons a:focus{background:#10683d}
        .hf-code-button{display:inline-flex;align-items:center;justify-content:center;
          min-height:55px;padding:12px 20px;margin-top:12px;box-sizing:border-box;
          border-radius:10px;background:#16854e;color:#fff!important;
          text-decoration:none!important;font-size:20px;font-weight:700;line-height:1.4;
          text-align:center}
        .hf-code-button:hover,.hf-code-button:focus{background:#10683d}
        @media(max-width:520px){.hf-code-button{display:flex;width:100%;font-size:19px}}
        @media(max-width:520px){.hf-nav-buttons{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:9px}
          .hf-nav-buttons a{font-size:17px;padding:12px 8px}}
        </style><nav class="hf-nav-buttons" aria-label="弘法功能导航">
        <a href="/hongfa/dashboard">🏠 管理首页</a>
        <a href="/hongfa/plans">📅 活动安排</a>
        <a href="/hongfa/week">📋 每周时间表</a>
        <a href="/hongfa/locations">📍 地点</a></nav>'''
        return render(nav + body, csrf=csrf_token(), **kw)

    def event_for(who, event_id, leader_only=False):
        e = db_query('''SELECT e.*,t.name AS team_name,l.name AS location_name
             FROM hongfa_events e JOIN hongfa_teams t ON t.id=e.team_id
             LEFT JOIN hongfa_locations l ON l.id=e.location_id WHERE e.id=%s''',
             (event_id,), fetchone=True)
        if not e:
            abort(404)
        if leader_only:
            if e['team_id'] not in who['team_ids']:
                abort(403)
        elif not who['secretary'] and e['team_id'] not in who['team_ids']:
            abort(403)
        return e

    @bp.route('/plans', methods=['GET', 'POST'])
    @guard()
    def plans(who):
        error = ''
        if request.method == 'POST':
            try:
                team_id = int(request.form['team_id'])
                if not who['secretary'] and team_id not in who['team_ids']:
                    abort(403)
                day = date.fromisoformat(request.form['event_date'])
                if day < datetime.now(TZ).date():
                    raise ValueError('不能提交过去的日期。')
                rest = request.form.get('rest') == '1'
                loc_id = None if rest else int(request.form['location_id'])
                start = None if rest else request.form['starts_at']
                end = None if rest else request.form['ends_at']
                if not rest and (not start or not end or start >= end):
                    raise ValueError('开始时间必须早于结束时间。')
                if loc_id is not None and not db_query('SELECT id FROM hongfa_locations WHERE id=%s AND active=true', (loc_id,), fetchone=True):
                    raise ValueError('地点无效。')
                # 每组每天最多一份计划，避免重复提交。
                if db_query('SELECT id FROM hongfa_events WHERE team_id=%s AND event_date=%s', (team_id, day), fetchone=True):
                    raise ValueError('该组当天已有安排，请勿重复提交。')
                db_query('''INSERT INTO hongfa_events
                    (team_id,location_id,event_date,starts_at,ends_at,status,note,created_by)
                    VALUES (%s,%s,%s,%s,%s,'draft',%s,%s)''',
                    (team_id,loc_id,day,start,end,(request.form.get('note') or '')[:500],who['id']))
                return redirect(url_for('hongfa.plans'))
            except (ValueError, KeyError, TypeError) as ex:
                error = str(ex)
        teams = db_query('SELECT id,name FROM hongfa_teams WHERE active=true ORDER BY id', fetchall=True) or []
        locations = db_query('SELECT id,name FROM hongfa_locations WHERE active=true ORDER BY name', fetchall=True) or []
        if who['secretary']:
            rows = db_query('''SELECT e.id,e.event_date,e.starts_at,e.ends_at,e.status,e.note,t.name AS team_name,
              l.name AS location_name FROM hongfa_events e JOIN hongfa_teams t ON t.id=e.team_id
              LEFT JOIN hongfa_locations l ON l.id=e.location_id
              WHERE e.event_date >= CURRENT_DATE ORDER BY e.event_date ASC,t.id LIMIT 150''', fetchall=True) or []
        else:
            rows = db_query('''SELECT e.id,e.event_date,e.starts_at,e.ends_at,e.status,e.note,t.name AS team_name,
              l.name AS location_name FROM hongfa_events e JOIN hongfa_teams t ON t.id=e.team_id
              LEFT JOIN hongfa_locations l ON l.id=e.location_id
              WHERE e.team_id=ANY(%s) AND e.event_date >= CURRENT_DATE
              ORDER BY e.event_date ASC,t.id LIMIT 150''', (who['team_ids'],), fetchall=True) or []
        return page('''<div class="card"><h2>📅 弘法活动安排</h2>
          {% if error %}<p class="warn">{{error}}</p>{% endif %}
          {% if who.team_ids or who.secretary %}<h3>{{'秘书代录弘法安排' if who.secretary else '负责人提交（秘书审核后才公开）'}}</h3>
          <form method="post"><input type="hidden" name="csrf_token" value="{{csrf}}">
          <style>
          .hf-rest-toggle{display:inline-flex;align-items:center;gap:12px;cursor:pointer;
            border:2px solid #dc3545;border-radius:12px;background:#fff4f4;
            padding:13px 18px;color:#b4232e;font-size:20px;font-weight:800;
            min-height:54px;box-sizing:border-box}
          .hf-rest-toggle input{width:23px!important;height:23px!important;margin:0!important;accent-color:#c62828}
          .hf-rest-toggle:has(input:checked){background:#ffe0e0;border-color:#b4232e}
          </style>
          <div class="hf-form-grid">
            <div class="hf-field"><label for="hf-team">👥 弘法组别</label><select id="hf-team" name="team_id" required>
            {% if who.secretary %}<option value="" selected disabled>请选择组别</option>{% endif %}{% for t in teams if who.secretary or t.id in who.team_ids %}<option value="{{t.id}}" {% if not who.secretary and t.id == who.team_ids[0] %}selected{% endif %}>{{t.name}}</option>{% endfor %}</select></div>
            <div class="hf-field"><label for="hf-day">📅 弘法日期</label><input id="hf-day" type="date" name="event_date" value="{{today.isoformat()}}" required></div>
            <div class="hf-field hf-wide"><label for="hf-location">📍 弘法地点</label><select id="hf-location" name="location_id"><option value="">请选择地点</option>
            {% for l in locations %}<option value="{{l.id}}">{{l.name}}</option>{% endfor %}</select></div>
            <div class="hf-field"><label for="hf-start">⏰ 开始时间</label><input id="hf-start" type="time" name="starts_at" value="07:00"></div>
            <div class="hf-field"><label for="hf-end">⏰ 结束时间</label><input id="hf-end" type="time" name="ends_at" value="09:00"></div>
            <div class="hf-field hf-wide"><label for="hf-note">📝 备注（选填）</label><input id="hf-note" name="note" maxlength="500" placeholder="特别说明"></div>
            <div class="hf-field hf-wide"><label class="hf-rest-toggle" for="hf-rest"><input id="hf-rest" type="checkbox" name="rest" value="1"> <span>🔴 当天休息 / Off</span></label></div>
          </div>
          <p class="hf-muted">默认菜市弘法 7:00am–9:00am；特殊活动可自行修改时间。</p>
          <div class="hf-actions"><button>{{'保存安排（待发布）' if who.secretary else '提交审核'}}</button></div>
          </form>
          <script>(function(){var c=document.getElementById('hf-rest');var ids=['hf-location','hf-start','hf-end'];c.addEventListener('change',function(){ids.forEach(function(id){document.getElementById(id).disabled=c.checked;});});})();</script>{% endif %}
          <h3>现有安排</h3>
          {% for e in rows %}<div class="hf-event"><b>📅 {{e.event_date}}｜{{e.team_name}}</b><br>
          📍 {{e.location_name or '休息 / Off'}}<br>
          {% if e.starts_at %}⏰ {{e.starts_at}}–{{e.ends_at}}<br>{% endif %}
          <span class="hf-muted">状态：{{'已发布' if e.status == 'published' else ('待审核' if e.status == 'draft' else e.status)}}</span><br>
          {% if who.secretary and e.status == 'draft' %}
          <form method="post" action="{{url_for('hongfa.review',event_id=e.id)}}" style="display:inline">
          <input type="hidden" name="csrf_token" value="{{csrf}}">
          <button name="decision" value="published">批准发布</button>
          <button name="decision" value="rejected" style="background:#a33">退回</button></form>{% endif %}
          {% if e.status == 'published' and e.location_name and e.event_date >= today and not who.secretary %}
          <a class="hf-code-button" href="{{url_for('hongfa.event_code',event_id=e.id)}}">🔐 生成／查看签到码</a>{% endif %}
          </div>{% else %}<p>暂无安排。</p>{% endfor %}</div>''',
          who=who, error=error, teams=teams, locations=locations, rows=rows, today=datetime.now(TZ).date())

    @bp.route('/plans/<int:event_id>/review', methods=['POST'])
    @guard(secretary=True)
    def review(who, event_id):
        e = event_for(who, event_id)
        if e['status'] != 'draft':
            abort(409)
        decision = request.form.get('decision')
        if decision not in ('published','rejected'):
            abort(400)
        db_query('UPDATE hongfa_events SET status=%s WHERE id=%s AND status=%s', (decision,event_id,'draft'))
        return redirect(url_for('hongfa.plans'))

    @bp.route('/locations', methods=['GET', 'POST'])
    @guard()
    def locations(who):
        error = ''
        if request.method == 'POST':
            if not (who['secretary'] or who['team_ids']):
                abort(403)
            name = (request.form.get('name') or '').strip()[:120]
            maps = (request.form.get('maps_url') or '').strip()[:1000]
            parsed = urlparse(maps) if maps else None
            if not name:
                error = '请输入地点名称。'
            elif maps and (parsed.scheme != 'https' or not parsed.netloc):
                error = '地图链接必须是 HTTPS。'
            elif db_query('SELECT id FROM hongfa_locations WHERE LOWER(name)=LOWER(%s) AND active=true LIMIT 1', (name,), fetchone=True):
                error = '此地点已经存在，请直接选择已有地点。'
            else:
                db_query('INSERT INTO hongfa_locations(name,maps_url) VALUES (%s,%s)', (name,maps or None))
                return redirect(url_for('hongfa.locations'))
        rows = db_query('SELECT id,name,maps_url FROM hongfa_locations WHERE active=true ORDER BY name',fetchall=True) or []
        return page('''<div class="card"><h2>📍 弘法地点</h2>
          {% if error %}<p class="warn">{{error}}</p>{% endif %}
          {% if who.secretary or who.team_ids %}<form method="post"><input type="hidden" name="csrf_token" value="{{csrf}}">
          <label>地点名称</label><input name="name" required maxlength="120">
          <label>Google Maps HTTPS 链接（可留空）</label><input name="maps_url" maxlength="1000">
          <button>新增地点</button></form>{% endif %}
          {% for l in rows %}<p>{{l.name}} {% if l.maps_url %}<a href="{{l.maps_url}}" target="_blank" rel="noopener noreferrer">导航</a>{% endif %}</p>
          {% else %}<p>暂无地点，可在上方新增。</p>{% endfor %}</div>''', who=who, error=error, rows=rows)

    @bp.route('/week')
    @guard()
    def week(who):
        raw = request.args.get('start','')
        try:
            start = date.fromisoformat(raw) if raw else datetime.now(TZ).date()
        except ValueError:
            abort(400)
        start -= timedelta(days=start.weekday())
        end = start + timedelta(days=7)
        teams = db_query('SELECT id,name FROM hongfa_teams WHERE active=true ORDER BY id',fetchall=True) or []
        if not who['secretary']:
            teams = [t for t in teams if t['id'] in who['team_ids']]
        rows = db_query('''SELECT e.team_id,e.event_date,e.starts_at,e.ends_at,e.status,
          l.name AS location_name FROM hongfa_events e LEFT JOIN hongfa_locations l ON l.id=e.location_id
          WHERE e.event_date >= %s AND e.event_date < %s AND e.status='published'
          ORDER BY e.event_date,e.team_id''',(start,end),fetchall=True) or []
        lookup = {(r['team_id'],r['event_date']):r for r in rows}
        # 周表仅列出已发布的活动；保留原有组别可见权限。
        grid=[]
        for offset in range(7):
            d = start + timedelta(days=offset)
            for t in teams:
                e = lookup.get((t['id'], d))
                if not e:
                    continue
                label = (e['location_name'] or '休息 / Off')
                if e['starts_at'] and e['ends_at']:
                    label += '　' + str(e['starts_at'])[:5] + '–' + str(e['ends_at'])[:5]
                grid.append((d, t['name'], label))

        # WhatsApp 通知保持观音堂原版排版；只读取已发布活动。
        names = {'俊达组': '俊达师兄', 'Alex组': 'Alex师兄', '诗霞组': '诗霞师兄'}
        order = {'Alex组': 0, '俊达组': 1, '诗霞组': 2}

        def readable_time(value):
            if value is None:
                return ''
            parts = str(value).split(':')
            hour, minute = int(parts[0]), int(parts[1])
            suffix = 'am' if hour < 12 else 'pm'
            h12 = hour % 12 or 12
            return f'{h12}{suffix}' if minute == 0 else f'{h12}:{minute:02d}{suffix}'

        def time_label(event):
            return f"{readable_time(event['starts_at'])} – {readable_time(event['ends_at'])}"

        lines = ['师兄们，大家好。', '', '以下是下个星期的弘法地点，欢迎师兄来助缘:', '']
        weekdays = '一二三四五六日'
        for offset in range(7):
            d = start + timedelta(days=offset)
            activities = []
            rests = []
            for t in teams:
                e = lookup.get((t['id'], d))
                if not e:
                    continue
                if e['location_name']:
                    activities.append((t, e))
                else:
                    rests.append(t)
            if not activities and not rests:
                continue
            activities.sort(key=lambda item: (order.get(item[0]['name'], 99), item[0]['name']))
            lines.append(f'🚩 ({d:%d/%m/%y}) 星期{weekdays[d.weekday()]}')
            if activities:
                same_time = len({time_label(e) for _, e in activities}) == 1
                if same_time:
                    lines.append(f'时间：{time_label(activities[0][1])}')
                for i, (t, e) in enumerate(activities, 1):
                    place_label = '地点' if len(activities) == 1 else f'地点 {i}'
                    line = f"{place_label}：{e['location_name']}（{names.get(t['name'], t['name'])}）"
                    if not same_time:
                        line = f'时间：{time_label(e)}｜' + line
                    lines.append(line)
            for t in rests:
                lines.append(f"{t['name']}：休息 / Off")
            lines.append('')
        lines.extend([
            '🙏😇请每位弘法的师兄们在弘法发资料时心中祈求为对方种下善因：',
            '💓每发出一份资料的时候，心里默默地祈求："菩萨妈妈帮拿到资料的人能够种下与佛菩萨结缘的善因，等到因缘成熟的时候再得度" 💓',
            '',
            '🏳️‍🌈菜市弘法早上7am to 9am.',
            '🏳️‍🌈欢迎前来助缘，感恩师兄们无缘大慈，同体大悲，广度更多有缘众生'
        ])
        whatsapp = '\n'.join(lines)
        return page('''<div class="card hf-week"><style>
          .hf-week{font-size:19px;line-height:1.75}
          .hf-week h2{font-size:27px;margin:0 0 12px}
          .hf-week h3{font-size:23px;margin:25px 0 10px}
          .hf-week .hf-week-range{font-size:17px;color:#607369}
          .hf-week .hf-week-item{padding:13px 15px;margin:12px 0;background:#f6faf7;border:1px solid #dce9df;border-radius:12px}
          .hf-week .hf-week-item strong{font-size:19px}
          .hf-week .hf-week-place{font-size:19px;overflow-wrap:anywhere}
          .hf-week .hf-week-note{font-size:16px;color:#607369}
          .hf-week #weekText{width:100%;min-height:420px;box-sizing:border-box;white-space:pre-wrap;font-size:19px;line-height:1.7;padding:15px}
          .hf-week .hf-week-copy{font-size:19px;font-weight:700;margin:12px 0}
          .hf-week .hf-week-nav{font-size:18px}
          @media(max-width:520px){.hf-week{font-size:18px}.hf-week h2{font-size:24px}.hf-week #weekText{font-size:18px}.hf-week .hf-week-item strong,.hf-week .hf-week-place{font-size:18px}}
          </style><h2>📋 每周弘法时间表</h2>
          <p class="hf-week-range">{{start}} 至 {{end}}</p>
          {% for day,team,label in grid %}<div class="hf-week-item"><strong>📅 {{day}}｜{{team}}</strong><div class="hf-week-place">📍 {{label}}</div></div>
          {% else %}<p class="hf-week-note">本周暂无已发布的弘法安排。</p>{% endfor %}
          <h3>📱 WhatsApp 通知预览</h3>
          <p class="hf-week-note">仅包含已发布安排。</p>
          <textarea id="weekText" readonly>{{whatsapp}}</textarea>
          <button class="hf-week-copy" type="button" onclick="hongfaCopyWeek()">复制 WhatsApp 通知</button>
          <span id="hf-copy-status" role="status" aria-live="polite"></span>
          <script>
          function hongfaCopyWeek(){
            var box=document.getElementById('weekText');
            var status=document.getElementById('hf-copy-status');
            function fallback(){
              box.focus();box.select();box.setSelectionRange(0,box.value.length);
              var copied=false;
              try{copied=document.execCommand('copy');}catch(e){}
              status.textContent=copied?'✅ 已复制，可以粘贴到 WhatsApp':'无法自动复制；文字已选中，请按 Ctrl+C 复制。';
            }
            if(navigator.clipboard && window.isSecureContext){
              navigator.clipboard.writeText(box.value).then(function(){status.textContent='✅ 已复制，可以粘贴到 WhatsApp';}).catch(fallback);
            }else{fallback();}
          }
          </script>
          <p class="hf-week-nav"><a href="?start={{prev}}">上一周</a>　<a href="?start={{nxt}}">下一周</a></p></div>''',
          grid=grid, whatsapp=whatsapp, start=start, end=end-timedelta(days=1),
          prev=start-timedelta(days=7), nxt=end)

    @bp.route('/plans/<int:event_id>/code',methods=['GET','POST'])
    @guard()
    def event_code(who,event_id):
        e=event_for(who,event_id,leader_only=True)
        if e['status'] != 'published' or not e['location_id'] or e['closed_at']:
            abort(409)
        today=datetime.now(TZ).date()
        # 现场码仅当天生成；不在数据库保存明文。
        if e['event_date'] != today:
            return page('''<div class="card"><h2>🔐 今日签到码｜{{e.team_name}}</h2>
              <p>{{e.event_date}}｜{{e.location_name}}</p>
              <p>⏳ 活动当天才可以生成签到码。</p></div>''',e=e)
        secret=current_app.secret_key
        if not secret:
            abort(503,description='服务器尚未配置固定 SECRET_KEY，无法安全生成签到码。')
        key=secret.encode('utf-8') if isinstance(secret,str) else secret
        # 同一活动同一天使用同一4位码，刷新页面不会改变。
        seed=hmac.new(key,f'hongfa-code-v1:{event_id}:{e["event_date"]}'.encode(),hashlib.sha256).digest()
        code=f'{1000+int.from_bytes(seed[:8],"big")%9000:04d}'
        digest=hashlib.sha256(f'{event_id}:{code}'.encode()).hexdigest()
        error=''
        if request.method=='POST':
            if not e['starts_at'] or not e['ends_at']:
                error='活动尚未设置开始及结束时间。'
            else:
                start=datetime.combine(today,e['starts_at'].replace(tzinfo=None),tzinfo=TZ)-timedelta(minutes=30)
                end=datetime.combine(today,e['ends_at'].replace(tzinfo=None),tzinfo=TZ)+timedelta(minutes=30)
                # 活动开始/结束时间自动延伸30分钟，但窗口仅限当天。
                start=max(start,datetime.combine(today,datetime.min.time(),tzinfo=TZ))
                end=min(end,datetime.combine(today,datetime.max.time(),tzinfo=TZ))
                if start>=end:
                    error='活动时间设置不正确。'
                else:
                    # 已生成时不重复写入；旧6位码须明确点击替换。
                    replace_old=request.form.get('replace_old')=='1'
                    db_query('''UPDATE hongfa_events SET code_hash=%s,code_valid_from=%s,code_valid_until=%s
                      WHERE id=%s AND status='published' AND closed_at IS NULL
                      AND (code_hash IS NULL OR (%s AND code_hash<>%s))''',
                      (digest,start,end,event_id,replace_old,digest))
            e=event_for(who,event_id,leader_only=True)
        ready=(e['code_hash']==digest and e['code_valid_from'] and e['code_valid_until'])
        legacy=bool(e['code_hash'] and not ready)
        return page('''<div class="card" style="max-width:640px;margin:auto;text-align:center">
          <h2 style="font-size:28px">🔐 今日签到码｜{{e.team_name}}</h2>
          <p style="font-size:19px">{{e.event_date}}｜{{e.location_name}}</p>
          {% if error %}<p class="warn">{{error}}</p>{% endif %}
          {% if ready %}<p style="font-size:20px">今日签到码</p>
          <div style="font-size:64px;font-weight:800;letter-spacing:10px;color:#137e4c">{{code}}</div>
          <p>有效时间：{{e.code_valid_from.strftime('%H:%M')}} – {{e.code_valid_until.strftime('%H:%M')}}</p>
          {% else %}{% if legacy %}<p class="warn">此活动已有旧版现场码，点击下方按钮将替换旧码。</p>{% endif %}
          <form method="post"><input type="hidden" name="csrf_token" value="{{csrf}}">
          {% if legacy %}<input type="hidden" name="replace_old" value="1">{% endif %}
          <button style="font-size:23px;padding:16px 25px" type="submit">🟢 一键生成今日签到码</button></form>{% endif %}
          <p style="font-size:18px"><a href="{{url_for('hongfa.join',event_id=e.id)}}">📱 师兄现场签到</a>　
          <a href="{{url_for('hongfa.event_attendance',event_id=e.id)}}">👥 查看出席</a></p>
          </div>''',e=e,error=error,code=code,ready=ready,legacy=legacy)
