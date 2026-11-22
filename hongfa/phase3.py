"""弘法第三批：现场签到/签退、结束活动、独立库存、报表。"""
import hashlib
import hmac
import secrets
from datetime import datetime, date, timedelta
from functools import wraps
from zoneinfo import ZoneInfo
from flask import request, redirect, url_for, abort, session
from psycopg2.extras import RealDictCursor
from db import db_query, get_conn
from attendance_service import find_volunteer, verify_pin_for_volunteer
from .qr_signing import valid_qr

TZ = ZoneInfo('Asia/Kuala_Lumpur')

def register_phase3(bp, actor, render, csrf_token, valid_csrf):
    def page(body, **data):
        return render('<p><a href="/hongfa/dashboard">🏠 弘法管理首页</a>　<a href="/hongfa/attendance">👥 出席</a>　<a href="/hongfa/stock">📚 库存</a>　<a href="/hongfa/report">📊 报表</a></p>'+body, csrf=csrf_token(), **data)

    def guard(fn):
        @wraps(fn)
        def inner(*a, **kw):
            who=actor()
            if not who: return redirect(url_for('hongfa.login'))
            if request.method=='POST' and not valid_csrf(): abort(400)
            return fn(who,*a,**kw)
        return inner

    def event(event_id, who=None, leader=False):
        row=db_query('''SELECT e.*,t.name AS team_name,l.name AS location_name FROM hongfa_events e
          JOIN hongfa_teams t ON t.id=e.team_id LEFT JOIN hongfa_locations l ON l.id=e.location_id
          WHERE e.id=%s''',(event_id,),fetchone=True)
        if not row: abort(404)
        if who and (leader and row['team_id'] not in who['team_ids'] or
                    not leader and not who['secretary'] and row['team_id'] not in who['team_ids']): abort(403)
        return row

    def check_qr(e, token):
        return (e['status']=='published' and e['location_id'] and not e['closed_at']
                and e['event_date']==datetime.now(TZ).date() and valid_qr(e,token))

    def volunteer_id_from_form():
        """CHE 包含纯义工 8xx；STW 仍使用 STW-数字。"""
        branch = (request.form.get('branch') or 'CHE').upper()
        digits = (request.form.get('number') or '').strip()
        if branch not in ('CHE', 'STW') or not digits.isascii() or not digits.isdigit() or not 0 < int(digits) < 1000000:
            return ''
        number = int(digits)
        return str(number) if branch == 'CHE' and 800 <= number <= 899 else f'{branch}-{number}'

    def identity_ui(prefix):
        """共用签到和签退的分会按钮、身份切换与手机排版。"""
        return f'''<style>
          .hf-public{{max-width:580px;margin:auto}}.hf-public h2{{font-size:1.55rem;margin-top:0}}
          .hf-public .hf-sub{{background:#edf7f0;border-radius:12px;padding:12px;margin:12px 0 20px}}
          .hf-public .hf-field{{margin:14px 0}}.hf-public label{{display:block;font-weight:650;margin-bottom:7px}}
          .hf-public input{{width:100%;margin:0;min-height:48px}}
          .hf-public .hf-branch-row{{display:flex;gap:10px;align-items:center}}
          .hf-public .hf-branch-row input{{flex:1;min-width:0}}
          .hf-public .hf-branch{{width:112px;min-width:112px;min-height:48px;font-weight:750;background:#16854e}}
          .hf-public .hf-branch.stw{{background:#bb3030}}
          .hf-public .hf-kind-row{{display:flex;gap:10px;margin:16px 0}}
          .hf-public .hf-kind{{flex:1;background:#edf2ef;color:#244334;border:2px solid #d3e3d8}}
          .hf-public .hf-kind.selected{{background:#16854e;color:white;border-color:#16854e}}
          .hf-public .hf-kind.guest.selected{{background:#48647a;border-color:#48647a}}
          .hf-public .hf-submit{{width:100%;margin-top:14px;font-weight:700}}
          .hf-public .hf-hint{{font-size:13px;color:#65766c;margin-top:6px}}
          .hf-public [hidden]{{display:none!important}}
          @media(max-width:420px){{.hf-public .hf-kind-row{{gap:6px}}.hf-public .hf-kind{{padding:12px 6px;font-size:15px}}}}
        </style>
        <div class="hf-kind-row" role="group" aria-label="签到身份">
          <button class="hf-kind selected" type="button" data-kind="volunteer" onclick="hfKind_{prefix}('volunteer')">👤 已有义工编号</button>
          <button class="hf-kind guest" type="button" data-kind="guest" onclick="hfKind_{prefix}('guest')">🌱 新师兄（无编号）</button>
        </div>
        <input type="hidden" name="kind" id="hf-kind-{prefix}" value="volunteer">
        <script>
          function hfKind_{prefix}(kind){{
            document.getElementById('hf-kind-{prefix}').value=kind;
            document.getElementById('hf-vol-{prefix}').hidden=(kind!=='volunteer');
            document.getElementById('hf-guest-{prefix}').hidden=(kind!=='guest');
            document.querySelectorAll('#hf-form-{prefix} .hf-kind').forEach(function(b){{
              b.classList.toggle('selected',b.dataset.kind===kind);
            }});
          }}
          function hfBranch_{prefix}(){{
            var field=document.getElementById('hf-branch-{prefix}');
            var btn=document.getElementById('hf-branch-btn-{prefix}');
            field.value=field.value==='CHE'?'STW':'CHE';
            btn.textContent=field.value;
            btn.classList.toggle('stw',field.value==='STW');
            btn.setAttribute('aria-label','当前分会 '+field.value+'，点击切换');
          }}
        </script>'''

    @bp.route('/join/<int:event_id>',methods=['GET','POST'])
    def join(event_id):
        e=event(event_id)
        if request.method=='GET' and not check_qr(e,request.args.get('token','')):
            abort(403,description='请扫描负责人现场展示的今日二维码。')
        error=''; guest_token=None
        if request.method=='POST':
            if not valid_csrf(): abort(400)
            token=(request.form.get('token') or '').strip()
            if not check_qr(e,token):
                error='二维码无效、不是今天的二维码，或活动已结束。'
            else:
                kind=request.form.get('kind')
                if kind=='volunteer':
                    vid=volunteer_id_from_form()
                    vol=find_volunteer(vid) if vid else None
                    if not vol or str(vol.get('编号','')).upper()!=vid.upper() or vol.get('状态') not in ('在册','',None) or not (vol.get('PIN') or '') or not verify_pin_for_volunteer(vol,request.form.get('pin','')):
                        error='义工编号或 PIN 不正确。'
                    else:
                        person_key='v:'+vid.upper()
                        error=_insert_attendance(e,person_key,vid,None,None,None,token=token)
                        if not error:
                            session['hongfa_book_volunteer'] = {'event_id': e['id'], 'vid': vid, 'date': e['event_date'].isoformat(), 'issued': datetime.now(TZ).timestamp()}
                elif kind=='guest':
                    name=(request.form.get('name') or '').strip()[:100]
                    phone=''.join(c for c in (request.form.get('phone') or '') if c.isascii() and c.isdigit())
                    if len(name)<2 or not 9<=len(phone)<=15:
                        error='新人须填写姓名及有效电话号码。'
                    else:
                        guest_token=secrets.token_urlsafe(18)
                        person_key='g:'+phone
                        error=_insert_attendance(e,person_key,None,name,phone,hashlib.sha256(guest_token.encode()).hexdigest(),token=token)
                        if error: guest_token=None
                else: error='请选择签到身份。'
        return render('''<div class="card hf-public"><h2>🪷 现场弘法签到</h2>
          <div class="hf-sub"><b>{{e.team_name}}</b>｜{{e.event_date}}<br>📍 {{e.location_name}}</div>
          {% if error %}<p class="warn" role="alert">{{error}}</p>{% endif %}
          {% if success %}<h3>✅ 签到成功</h3>
          <p>签到已记录，请在弘法结束后自行签退。</p>
          {% if guest_token %}<p><b>新人签退凭证（请立即保存）：</b><br><span style="overflow-wrap:anywhere">{{guest_token}}</span></p>{% endif %}
          {% if book_access %}<p><a class="btn" style="display:block;text-align:center;font-size:21px;padding:17px" href="{{url_for('hongfa.books')}}">📚 结缘经书／快速扣书</a></p>{% endif %}
          <p><a class="btn" href="{{url_for('hongfa.self_checkout')}}">👋 前往自助签退</a></p>
          {% else %}<form id="hf-form-join" method="post"><input type="hidden" name="csrf_token" value="{{csrf}}">
          <input type="hidden" name="token" value="{{token}}">
          {{identity_ui|safe}}
          <div id="hf-vol-join">
            <div class="hf-field"><label for="hf-number-join">义工编号</label>
              <div class="hf-branch-row"><input type="hidden" id="hf-branch-join" name="branch" value="CHE">
              <button type="button" class="hf-branch" id="hf-branch-btn-join" onclick="hfBranch_join()" aria-label="当前分会 CHE，点击切换">CHE</button>
              <input id="hf-number-join" name="number" inputmode="numeric" pattern="[0-9]+" placeholder="只输入数字，例如 69 或 819"></div>
              <p class="hf-hint">🟢 CHE（包括 8xx 纯义工）｜🔴 STW；点击左侧按钮切换</p></div>
            <div class="hf-field"><label for="hf-pin-join">本人 PIN</label><input id="hf-pin-join" name="pin" type="password" autocomplete="current-password" placeholder="输入本人 PIN"></div>
          </div>
          <div id="hf-guest-join" hidden>
            <div class="hf-field"><label for="hf-name">中文姓名</label><input id="hf-name" name="name" autocomplete="name" placeholder="请输入姓名"></div>
            <div class="hf-field"><label for="hf-phone-join">电话号码</label><input id="hf-phone-join" name="phone" type="tel" inputmode="tel" autocomplete="tel" placeholder="例如 012-3456789"></div>
            <p class="hf-hint">新人签到成功后，请保存页面显示的签退凭证。</p>
          </div>
          <button class="hf-submit" type="submit">✅ 确认签到</button></form>{% endif %}</div>''',e=e,error=error,token=(request.args.get('token') if request.method=='GET' else request.form.get('token','')),success=request.method=='POST' and not error,guest_token=guest_token,book_access=(request.method=='POST' and not error and request.form.get('kind')=='volunteer'),csrf=csrf_token(),identity_ui=identity_ui('join'))

    def _insert_attendance(e,person_key,vid,name,phone,token_hash,token=None,operator=None):
        # 数据库唯一索引 + 人员键事务锁，避免并发重复；日期跨活动检查在同一事务中。
        try:
            with get_conn() as conn:
                with conn.cursor(cursor_factory=RealDictCursor) as cur:
                    cur.execute('SELECT pg_advisory_xact_lock(hashtextextended(%s,0))',(person_key,))
                    cur.execute('SELECT id,event_date,location_id,status,closed_at FROM hongfa_events WHERE id=%s FOR UPDATE',(e['id'],))
                    locked=cur.fetchone()
                    if not locked or locked['status']!='published' or locked['closed_at'] or not locked['location_id'] or locked['event_date']!=datetime.now(TZ).date() or (operator is None and not valid_qr(locked,token)):
                        return '签到已关闭或二维码无效，请重新进入。'
                    cur.execute('''SELECT 1 FROM hongfa_attendance a JOIN hongfa_events ev ON ev.id=a.event_id
                      WHERE a.person_key=%s AND ev.event_date=%s LIMIT 1''',(person_key,e['event_date']))
                    if cur.fetchone(): return '今天已经有弘法签到记录，不能重复签到。'
                    cur.execute('''INSERT INTO hongfa_attendance(event_id,person_key,volunteer_id,guest_name,guest_phone,guest_checkout_hash)
                      VALUES (%s,%s,%s,%s,%s,%s) RETURNING id''',(e['id'],person_key,vid,name,phone,token_hash))
                    inserted_id=cur.fetchone()['id']
                    if operator:
                        cur.execute('INSERT INTO hongfa_checkin_audit(attendance_id,operator_id,method) VALUES (%s,%s,%s)',(inserted_id,operator,'leader'))
            return ''
        except Exception as ex:
            if getattr(ex,'pgcode',None)=='23505': return '该人员已经签到，请勿重复。'
            raise

    @bp.route('/attendance/<int:event_id>/leader-checkin',methods=['GET','POST'])
    @guard
    def leader_checkin(who,event_id):
        e=event(event_id,who,leader=True)
        if e['status']!='published' or e['closed_at'] or not e['location_id'] or e['event_date']!=datetime.now(TZ).date():
            abort(409)
        error=''; candidate=None; success=False
        if request.method=='POST':
            vid=volunteer_id_from_form()
            vol=find_volunteer(vid) if vid else None
            if not vol or str(vol.get('编号','')).upper()!=vid.upper() or vol.get('状态') not in ('在册','',None):
                error='找不到有效在册义工。'
            elif request.form.get('action')=='lookup':
                candidate={'id':vid,'name':next((str(vol[k]).strip() for k in ('中文姓名','姓名','名字','name') if vol.get(k)),vid)}
            elif request.form.get('action')=='confirm':
                if request.form.get('confirm_present')!='yes':
                    error='请先确认义工本人已在现场。'
                else:
                    error=_insert_attendance(e,'v:'+vid.upper(),vid,None,None,None,operator=who['id'])
                    success=not error
            else: abort(400)
        return page('''<div class="card hf-public"><h2>👥 负责人现场义工签到</h2>
          <div class="hf-sub"><b>{{e.team_name}}</b>｜{{e.event_date}}｜{{e.location_name}}</div>
          {% if error %}<p class="warn">{{error}}</p>{% endif %}
          {% if success %}<h3>✅ 现场签到成功</h3>{% endif %}
          <form method="post"><input type="hidden" name="csrf_token" value="{{csrf}}">
          <div class="hf-branch-row"><input type="hidden" name="branch" id="leader-branch" value="{{request.form.get('branch','CHE')}}">
          <button type="button" onclick="var b=document.getElementById('leader-branch');b.value=b.value==='CHE'?'STW':'CHE';this.textContent=b.value;this.style.background=b.value==='CHE'?'#16854e':'#bb3030'">{{request.form.get('branch','CHE')}}</button>
          <input name="number" inputmode="numeric" pattern="[0-9]+" required placeholder="输入义工编号数字" value="{{request.form.get('number','')}}"></div>
          <button name="action" value="lookup">🔎 查询姓名</button>
          {% if candidate %}<div class="hf-sub"><b style="font-size:23px">{{candidate.name}}｜{{candidate.id}}</b></div>
          
          <label><input type="checkbox" name="confirm_present" value="yes" required style="width:auto"> 我确认义工本人已到现场</label><br>
          <button name="action" value="confirm">✅ 确认签到</button>{% endif %}
          </form><p><a href="{{url_for('hongfa.event_attendance',event_id=e.id)}}">查看出席记录</a></p></div>''',e=e,error=error,success=success,candidate=candidate,request=request)

    @bp.route('/checkout',methods=['GET','POST'])
    def self_checkout():
        error=''; done=False
        if request.method=='POST':
            if not valid_csrf(): abort(400)
            kind=request.form.get('kind')
            if kind=='volunteer':
                vid=volunteer_id_from_form()
                vol=find_volunteer(vid) if vid else None
                if not vol or str(vol.get('编号','')).upper()!=vid.upper() or not (vol.get('PIN') or '') or not verify_pin_for_volunteer(vol,request.form.get('pin','')):
                    error='编号或 PIN 不正确。'
                else:
                    count=db_query('''UPDATE hongfa_attendance a SET check_out_at=now(),checkout_method='self'
                      FROM hongfa_events e WHERE a.event_id=e.id AND a.person_key=%s AND a.check_out_at IS NULL
                      AND e.event_date=%s AND e.closed_at IS NULL RETURNING a.id''',('v:'+vid.upper(),datetime.now(TZ).date()),fetchall=True) or []
                    done=bool(count)
            elif kind=='guest':
                phone=''.join(c for c in (request.form.get('phone') or '') if c.isascii() and c.isdigit())
                token=(request.form.get('token') or '').strip()
                digest=hashlib.sha256(token.encode()).hexdigest() if token else ''
                count=db_query('''UPDATE hongfa_attendance a SET check_out_at=now(),checkout_method='self'
                  FROM hongfa_events e WHERE a.event_id=e.id AND a.person_key=%s AND a.guest_checkout_hash=%s
                  AND a.check_out_at IS NULL AND e.event_date=%s AND e.closed_at IS NULL RETURNING a.id''',
                  ('g:'+phone,digest,datetime.now(TZ).date()),fetchall=True) or []
                done=bool(count)
            else: error='请选择身份。'
            if not error and not done: error='找不到今天尚未签退的记录，或活动已结束。'
        return render('''<div class="card hf-public"><h2>👋 弘法自助签退</h2>
          <div class="hf-sub">弘法结束后，请登记实际签退时间。</div>
          {% if done %}<h3>✅ 已记录实际签退时间</h3>{% endif %}
          {% if error %}<p class="warn" role="alert">{{error}}</p>{% endif %}
          <form id="hf-form-checkout" method="post"><input type="hidden" name="csrf_token" value="{{csrf}}">
          {{identity_ui|safe}}
          <div id="hf-vol-checkout">
            <div class="hf-field"><label for="hf-number-checkout">义工编号</label>
              <div class="hf-branch-row"><input type="hidden" id="hf-branch-checkout" name="branch" value="CHE">
              <button type="button" class="hf-branch" id="hf-branch-btn-checkout" onclick="hfBranch_checkout()" aria-label="当前分会 CHE，点击切换">CHE</button>
              <input id="hf-number-checkout" name="number" inputmode="numeric" pattern="[0-9]+" placeholder="只输入数字，例如 69 或 819"></div>
              <p class="hf-hint">🟢 CHE（包括 8xx 纯义工）｜🔴 STW</p></div>
            <div class="hf-field"><label for="hf-pin-checkout">本人 PIN</label><input id="hf-pin-checkout" name="pin" type="password" autocomplete="current-password"></div>
          </div>
          <div id="hf-guest-checkout" hidden>
            <div class="hf-field"><label for="hf-phone-checkout">新人电话号码</label><input id="hf-phone-checkout" name="phone" type="tel" inputmode="tel" autocomplete="tel"></div>
            <div class="hf-field"><label for="hf-token">签到时保存的签退凭证</label><input id="hf-token" name="token" autocomplete="off" placeholder="请输入签到成功时保存的凭证"></div>
          </div>
          <button class="hf-submit" type="submit">✅ 确认签退</button></form></div>''',csrf=csrf_token(),error=error,done=done,identity_ui=identity_ui('checkout'))

    @bp.route('/attendance')
    @guard
    def attendance(who):
        if who['secretary']:
            rows=db_query('''SELECT e.id,e.event_date,t.name AS team_name,e.status,e.closed_at,
              COUNT(a.id) AS total,COUNT(a.id) FILTER(WHERE a.check_out_at IS NULL) AS open_count
              FROM hongfa_events e JOIN hongfa_teams t ON t.id=e.team_id
              LEFT JOIN hongfa_attendance a ON a.event_id=e.id
              WHERE e.event_date >= %s GROUP BY e.id,t.name ORDER BY e.event_date DESC LIMIT 100''',
              (datetime.now(TZ).date()-timedelta(days=45),),fetchall=True) or []
        else:
            rows=db_query('''SELECT e.id,e.event_date,t.name AS team_name,e.status,e.closed_at,
              COUNT(a.id) AS total,COUNT(a.id) FILTER(WHERE a.check_out_at IS NULL) AS open_count
              FROM hongfa_events e JOIN hongfa_teams t ON t.id=e.team_id
              LEFT JOIN hongfa_attendance a ON a.event_id=e.id
              WHERE e.team_id=ANY(%s) AND e.event_date >= %s
              GROUP BY e.id,t.name ORDER BY e.event_date DESC LIMIT 100''',
              (who['team_ids'],datetime.now(TZ).date()-timedelta(days=45)),fetchall=True) or []
        return page('''<div class="card"><h2>👥 弘法活动出席</h2>
          {% for e in rows %}<p>{{e.event_date}}｜{{e.team_name}}｜{{e.status}}｜出席 {{e.total}}｜未签退 {{e.open_count}}
          <a href="{{url_for('hongfa.event_attendance',event_id=e.id)}}">查看</a></p>
          {% else %}<p>暂无活动。</p>{% endfor %}</div>''',rows=rows)

    @bp.route('/attendance/<int:event_id>')
    @guard
    def event_attendance(who,event_id):
        e=event(event_id,who)
        rows=db_query('''SELECT volunteer_id,guest_name,guest_phone,check_in_at,check_out_at,checkout_method,
          ROUND(EXTRACT(EPOCH FROM (COALESCE(check_out_at,now())-check_in_at))/3600.0,2) AS hours,
          (SELECT operator_id FROM hongfa_checkin_audit h WHERE h.attendance_id=hongfa_attendance.id) AS checkin_operator
          FROM hongfa_attendance WHERE event_id=%s ORDER BY check_in_at''',(event_id,),fetchall=True) or []
        return page('''<div class="card"><h2>👥 {{e.team_name}}｜{{e.event_date}}</h2>
          <p>实际出席 {{rows|length}} 人</p>
          {% for a in rows %}<p>{{a.volunteer_id or a.guest_name}}｜{{a.check_in_at}} → {{a.check_out_at or '未签退'}}{% if a.checkin_operator %}｜负责人代签到：{{a.checkin_operator}}{% endif %}
          {% if a.check_out_at %}｜{{a.hours}} 小时{% endif %}</p>{% else %}<p>暂无签到。</p>{% endfor %}
          {% if e.team_id in who.team_ids and e.status=='published' and not e.closed_at and e.event_date == today %}
          <form method="post" action="{{url_for('hongfa.close_event',event_id=e.id)}}" onsubmit="return confirm('确定结束活动？未签退人员将以当前实际时间签退。')">
          <input type="hidden" name="csrf_token" value="{{csrf}}"><button>🔒 结束活动并统一签退</button></form>{% endif %}
          {% if e.team_id in who.team_ids and e.status=='published' and e.closed_at and e.event_date >= today %}
          <p class="hf-muted">活动已结束。如属误操作，可恢复活动；原有签到及签退记录保持不变。</p>
          <form method="post" action="{{url_for('hongfa.reopen_event',event_id=e.id)}}" onsubmit="return confirm('确定恢复这场弘法活动？原有签到和签退记录不会更改，需重新生成现场码。')">
          <input type="hidden" name="csrf_token" value="{{csrf}}"><button style="background:#16854e">🔓 恢复活动</button></form>{% endif %}</div>''',e=e,rows=rows,who=who,today=datetime.now(TZ).date())

    @bp.route('/attendance/<int:event_id>/close',methods=['POST'])
    @guard
    def close_event(who,event_id):
        e=event(event_id,who,leader=True)
        if e['event_date'] != datetime.now(TZ).date(): abort(409)
        with get_conn() as conn:
            with conn.cursor() as cur:
                cur.execute('''UPDATE hongfa_events SET closed_at=now(),closed_by=%s,code_hash=NULL,
                  code_valid_from=NULL,code_valid_until=NULL WHERE id=%s AND status='published' AND closed_at IS NULL RETURNING closed_at''',(who['id'],event_id))
                row=cur.fetchone()
                if row:
                    cur.execute('''UPDATE hongfa_attendance SET check_out_at=%s,checkout_method='leader_close'
                      WHERE event_id=%s AND check_out_at IS NULL''',(row[0],event_id))
        return redirect(url_for('hongfa.event_attendance',event_id=event_id))

    @bp.route('/attendance/<int:event_id>/reopen',methods=['POST'])
    @guard
    def reopen_event(who,event_id):
        e=event(event_id,who,leader=True)
        if e['event_date'] < datetime.now(TZ).date(): abort(409)
        # 仅恢复活动开放状态；绝不撤销已经记录的签退。
        db_query('''UPDATE hongfa_events SET closed_at=NULL,closed_by=NULL,
          code_hash=NULL,code_valid_from=NULL,code_valid_until=NULL
          WHERE id=%s AND team_id=ANY(%s) AND status='published' AND closed_at IS NOT NULL''',
          (event_id,who['team_ids']))
        return redirect(url_for('hongfa.event_attendance',event_id=event_id))

    @bp.route('/stock',methods=['GET','POST'])
    @guard
    def stock(who):
        error=''
        if request.method=='POST':
            try:
                team_id=int(request.form.get('team_id','0'))
                if team_id not in who['team_ids']: abort(403)
                item_id=int(request.form.get('item_id','0'))
                qty=int(request.form.get('quantity','0'))
                typ=request.form.get('movement_type')
                if not 1<=qty<=100000 or typ not in ('receipt','distribution','adjust_in','adjust_out'):
                    raise ValueError('数量或操作类型无效。')
                ev_id=int(request.form.get('event_id','0')) if typ=='distribution' else None
                if typ=='distribution' and not ev_id: raise ValueError('结缘扣减必须选择已发布的本组活动。')
                reqkey=(request.form.get('request_key') or '').strip()
                if not reqkey: raise ValueError('请求编号缺失，请刷新页面。')
                with get_conn() as conn:
                    with conn.cursor(cursor_factory=RealDictCursor) as cur:
                        cur.execute('SELECT id FROM hongfa_teams WHERE id=%s FOR UPDATE',(team_id,))
                        if not cur.fetchone(): raise ValueError('组别无效。')
                        cur.execute('SELECT id FROM hongfa_items WHERE id=%s AND active=true',(item_id,))
                        if not cur.fetchone(): raise ValueError('经书无效。')
                        if ev_id:
                            cur.execute("SELECT id FROM hongfa_events WHERE id=%s AND team_id=%s AND status='published'",(ev_id,team_id))
                            if not cur.fetchone(): raise ValueError('活动不属于本组或未发布。')
                        cur.execute('SELECT id FROM hongfa_stock_movements WHERE request_key=%s',(reqkey,))
                        if cur.fetchone(): raise ValueError('此操作已登记，请勿重复提交。')
                        cur.execute('SELECT COALESCE(SUM(quantity),0) AS balance FROM hongfa_stock_movements WHERE team_id=%s AND item_id=%s',(team_id,item_id))
                        balance=int(cur.fetchone()['balance'])
                        delta=qty if typ in ('receipt','adjust_in') else -qty
                        if balance+delta<0: raise ValueError(f'库存不足，目前只有 {balance} 本。')
                        cur.execute('''INSERT INTO hongfa_stock_movements(team_id,item_id,event_id,quantity,movement_type,created_by,request_key)
                          VALUES (%s,%s,%s,%s,%s,%s,%s)''',(team_id,item_id,ev_id,delta,typ,who['id'],reqkey))
                return redirect(url_for('hongfa.stock'))
            except (ValueError,TypeError) as ex: error=str(ex)
        teams=db_query('SELECT id,name FROM hongfa_teams WHERE active=true ORDER BY id',fetchall=True) or []
        if not who['secretary']: teams=[t for t in teams if t['id'] in who['team_ids']]
        items=db_query('SELECT id,title FROM hongfa_items WHERE active=true ORDER BY title',fetchall=True) or []
        events=db_query("SELECT id,team_id,event_date FROM hongfa_events WHERE status='published' AND event_date>=%s ORDER BY event_date DESC LIMIT 100",(datetime.now(TZ).date()-timedelta(days=90),),fetchall=True) or []
        events=[e for e in events if e['team_id'] in who['team_ids']]
        balances=db_query('''SELECT m.team_id,t.name AS team_name,i.title,SUM(m.quantity) AS balance
          FROM hongfa_stock_movements m JOIN hongfa_teams t ON t.id=m.team_id JOIN hongfa_items i ON i.id=m.item_id
          GROUP BY m.team_id,t.name,i.title ORDER BY t.name,i.title''',fetchall=True) or []
        balances=[b for b in balances if who['secretary'] or b['team_id'] in who['team_ids']]
        return page('''<div class="card"><h2>📚 各组经书库存</h2>
          {% for b in balances %}<p>{{b.team_name}}｜{{b.title}}｜<b>{{b.balance}}</b> 本</p>{% else %}<p>暂无库存记录。</p>{% endfor %}
          {% if error %}<p class="warn">{{error}}</p>{% endif %}
          {% if who.team_ids %}<h3>负责人登记领取／结缘</h3>
          {% if not items %}<p>经书目录为空，请先由秘书新增书名。</p>{% endif %}
          <form method="post"><input type="hidden" name="csrf_token" value="{{csrf}}">
          <input type="hidden" name="request_key" value="{{request_key}}">
          <label>组别</label><select name="team_id">{% for t in teams if t.id in who.team_ids %}<option value="{{t.id}}">{{t.name}}</option>{% endfor %}</select>
          <label>经书</label><select name="item_id">{% for i in items %}<option value="{{i.id}}">{{i.title}}</option>{% endfor %}</select>
          <label>操作</label><select name="movement_type"><option value="receipt">藏经阁领取（增加）</option><option value="distribution">弘法结缘（扣减）</option><option value="adjust_in">库存调整增加</option><option value="adjust_out">库存调整减少</option></select>
          <label>数量（正整数）</label><input type="number" name="quantity" min="1" required>
          <label>结缘对应活动（领取可不选）</label><select name="event_id"><option value="">请选择</option>
          {% for e in events %}<option value="{{e.id}}">#{{e.id}}｜{{e.event_date}}｜组 {{e.team_id}}</option>{% endfor %}</select>
          <button>保存库存记录</button></form>{% endif %}
          {% if who.secretary %}<p>秘书只有查看库存权限，不可修改。</p>{% endif %}</div>''',who=who,teams=teams,items=items,events=events,balances=balances,error=error,request_key=secrets.token_urlsafe(24))

    @bp.route('/items',methods=['GET','POST'])
    @guard
    def items(who):
        if not who['secretary']: abort(403)
        error=''
        if request.method=='POST':
            title=(request.form.get('title') or '').strip()[:150]
            if not title: error='请输入经书名称。'
            else:
                db_query('INSERT INTO hongfa_items(title) SELECT %s WHERE NOT EXISTS(SELECT 1 FROM hongfa_items WHERE title=%s)',(title,title))
                return redirect(url_for('hongfa.items'))
        rows=db_query('SELECT title FROM hongfa_items ORDER BY title',fetchall=True) or []
        return page('''<div class="card"><h2>📖 经书目录（秘书维护）</h2>
          {% if error %}<p class="warn">{{error}}</p>{% endif %}
          <form method="post"><input type="hidden" name="csrf_token" value="{{csrf}}"><input name="title" maxlength="150" placeholder="书名" required><button>新增书名</button></form>
          {% for r in rows %}<p>{{r.title}}</p>{% endfor %}</div>''',rows=rows,error=error)

    @bp.route('/report')
    @guard
    def report(who):
        raw=request.args.get('month') or datetime.now(TZ).strftime('%Y-%m')
        try: first=date.fromisoformat(raw+'-01')
        except ValueError: abort(400)
        last=(first.replace(day=28)+timedelta(days=4)).replace(day=1)
        rows=db_query('''SELECT e.team_id,t.name AS team_name,e.event_date,l.name AS location_name,
          COUNT(a.id) AS attendees,
          COALESCE(ROUND(SUM(EXTRACT(EPOCH FROM (a.check_out_at-a.check_in_at))/3600.0),2),0) AS total_hours
          FROM hongfa_events e JOIN hongfa_teams t ON t.id=e.team_id
          LEFT JOIN hongfa_locations l ON l.id=e.location_id
          LEFT JOIN hongfa_attendance a ON a.event_id=e.id
          WHERE e.event_date >= %s AND e.event_date < %s
          GROUP BY e.id,e.team_id,t.name,e.event_date,l.name ORDER BY e.event_date,t.name''',(first,last),fetchall=True) or []
        rows=[r for r in rows if who['secretary'] or r['team_id'] in who['team_ids']]
        dist=db_query('''SELECT m.team_id,t.name AS team_name,i.title,COALESCE(SUM(-m.quantity),0) AS distributed
          FROM hongfa_stock_movements m JOIN hongfa_events e ON e.id=m.event_id
          JOIN hongfa_teams t ON t.id=m.team_id JOIN hongfa_items i ON i.id=m.item_id
          WHERE m.movement_type='distribution' AND e.event_date >= %s AND e.event_date < %s
          GROUP BY m.team_id,t.name,i.title ORDER BY t.name,i.title''',(first,last),fetchall=True) or []
        dist=[r for r in dist if who['secretary'] or r['team_id'] in who['team_ids']]
        locations=[]
        # 地点统计避免因 attendance JOIN 放大 visits。
        if who['secretary']:
            locations=db_query('''SELECT l.name,COUNT(DISTINCT e.id) AS visits,MAX(e.event_date) AS last_visit,
              COUNT(a.id) AS attendees FROM hongfa_locations l JOIN hongfa_events e ON e.location_id=l.id
              LEFT JOIN hongfa_attendance a ON a.event_id=e.id
              WHERE e.status='published' AND e.closed_at IS NOT NULL AND e.event_date < %s
              GROUP BY l.id,l.name ORDER BY visits DESC,l.name''',(last,),fetchall=True) or []
        return page('''<div class="card"><h2>📊 弘法月报｜{{month}}</h2>
          <form method="get"><input type="month" name="month" value="{{month}}"><button>查询</button></form>
          <h3>出席及实际时数（仅已签退计算时数）</h3>
          {% for r in rows %}<p>{{r.event_date}}｜{{r.team_name}}｜{{r.location_name or 'Off'}}｜{{r.attendees}} 人｜{{r.total_hours}} 小时</p>{% else %}<p>暂无活动。</p>{% endfor %}
          <h3>经书结缘</h3>{% for r in dist %}<p>{{r.team_name}}｜{{r.title}}｜{{r.distributed}} 本</p>{% else %}<p>暂无结缘。</p>{% endfor %}
          {% if who.secretary %}<h3>地点历史（截至本月末）</h3>{% for r in locations %}<p>{{r.name}}｜{{r.visits}} 次｜{{r.attendees}} 人次｜最近 {{r.last_visit}}</p>{% endfor %}{% endif %}</div>''',rows=rows,dist=dist,locations=locations,month=raw,who=who)
