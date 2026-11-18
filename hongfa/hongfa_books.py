"""弘法书籍快速登记：独立模块，复用现有库存流水表。"""
import json
import secrets
from datetime import datetime
from zoneinfo import ZoneInfo
from functools import wraps
from flask import request, redirect, url_for, abort, session
from psycopg2.extras import RealDictCursor
from db import db_query, get_conn

TZ = ZoneInfo('Asia/Kuala_Lumpur')


def register_books(bp, actor, render, csrf_token, valid_csrf):
    def guard(fn):
        @wraps(fn)
        def wrapped(*args, **kwargs):
            who = actor()
            if not who:
                grant = session.get('hongfa_book_volunteer')
                if not isinstance(grant, dict):
                    return redirect(url_for('hongfa.login'))
                try:
                    event_id = int(grant['event_id'])
                    vid = str(grant['vid'])
                    issued = float(grant['issued'])
                    if not 0 <= datetime.now(TZ).timestamp() - issued <= 8 * 3600 or grant['date'] != datetime.now(TZ).date().isoformat():
                        raise ValueError()
                except (ValueError, TypeError, KeyError):
                    session.pop('hongfa_book_volunteer', None)
                    abort(403)
                verified = db_query('''SELECT e.team_id FROM hongfa_attendance a
                    JOIN hongfa_events e ON e.id=a.event_id
                    WHERE e.id=%s AND e.event_date=%s AND e.status='published'
                      AND e.closed_at IS NULL AND a.volunteer_id=%s LIMIT 1''',
                    (event_id, datetime.now(TZ).date(), vid), fetchone=True)
                if not verified:
                    abort(403)
                who = {'id': vid, 'secretary': False, 'team_ids': [verified['team_id']], 'book_volunteer': True, 'book_event': event_id}

            if request.method == 'POST' and not valid_csrf():
                abort(400)
            return fn(who, *args, **kwargs)
        return wrapped

    def books_page(who, message='', error=''):
        is_volunteer = who.get('book_volunteer', False)
        teams = db_query('SELECT id,name FROM hongfa_teams WHERE active=true ORDER BY id', fetchall=True) or []
        teams = [t for t in teams if t['id'] in who['team_ids']]
        items = db_query('SELECT id,title FROM hongfa_items WHERE active=true ORDER BY title', fetchall=True) or []
        categories = db_query('SELECT item_id,category FROM hongfa_book_categories', fetchall=True) or []
        favorites = db_query('SELECT team_id,item_id,sort_order FROM hongfa_book_favorites WHERE team_id=ANY(%s) ORDER BY sort_order,item_id', (who['team_ids'],), fetchall=True) or [] if who['team_ids'] else []
        balances = db_query('''SELECT team_id,item_id,SUM(quantity) AS qty
            FROM hongfa_stock_movements WHERE team_id=ANY(%s)
            GROUP BY team_id,item_id''', (who['team_ids'],), fetchall=True) or [] if who['team_ids'] else []
        stock = {f"{b['team_id']}:{b['item_id']}": int(b['qty']) for b in balances}
        today = datetime.now(TZ).date()
        events = db_query('''SELECT id,team_id FROM hongfa_events WHERE event_date=%s
            AND status='published' AND location_id IS NOT NULL AND team_id=ANY(%s)
            AND closed_at IS NULL''', (today, who['team_ids']), fetchall=True) or [] if who['team_ids'] else []
        event_map = {str(e['team_id']): e['id'] for e in events}
        return render('''<style>
        .hb-tabs,.hb-filters{display:flex;gap:7px;flex-wrap:wrap;margin:12px 0}
        .hb-tabs button,.hb-filters button{font-size:16px;padding:10px 13px;background:#e3eee6;color:#244d35}
        .hb-tabs button.on,.hb-filters button.on{background:#16854e;color:#fff}
        .hb-grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:9px}
        .hb-book{border:1px solid #dce9df;border-radius:12px;padding:12px;min-width:0;background:#fbfdfb}
        .hb-name{font-size:17px;font-weight:750;min-height:52px;overflow-wrap:anywhere}
        .hb-sub{font-size:13px;color:#66776f;margin:5px 0}
        .hb-step{display:flex;align-items:center;justify-content:space-between;gap:5px}
        .hb-step button{width:40px;height:42px;padding:0;font-size:24px}
        .hb-step b{font-size:22px}.hb-picked{border:2px solid #16854e;border-radius:12px;padding:12px;background:#eaf5ec;margin:12px 0}
        .hb-bottom{position:fixed;bottom:0;left:0;right:0;background:#fff;padding:10px 14px max(10px,env(safe-area-inset-bottom));box-shadow:0 -3px 12px #0002;z-index:30}
        .hb-bottom button{display:block;width:min(730px,100%);margin:auto;font-size:20px;font-weight:800}
        .hb-area{padding-bottom:85px}.hb-hidden{display:none!important}
        @media(max-width:520px){.hb-grid{grid-template-columns:repeat(2,minmax(0,1fr))}.hb-name{font-size:15px;min-height:54px}.hb-step button{width:37px}}
        </style><div class="card hb-area"><h2>📚 弘法书籍｜快速登记</h2>
        {% if not is_volunteer %}<p><a href="/hongfa/dashboard">🏠 返回弘法管理首页</a></p>{% endif %}
        <p class="hf-muted">{% if is_volunteer %}已绑定今日签到，可登记本组结缘书籍。{% else %}负责人可扣书、点货及调整库存。{% endif %}</p>
        {% if error %}<p class="warn">{{error}}</p>{% endif %}{% if message %}<p style="color:#16854e">{{message}}</p>{% endif %}
        {% if not is_volunteer %}<p><a class="btn" id="hb-manage-link" href="{{url_for('hongfa.books_manage')}}">⚙️ 管理书籍</a></p>{% endif %}
        {% if not teams %}<p>秘书可查看原有库存页面；本页面需要负责人组别权限。</p>{% else %}
        <label>弘法组</label><select id="hb-team" onchange="hbCart={};hbDraw()">{% for t in teams %}<option value="{{t.id}}">{{t.name}}</option>{% endfor %}</select>
        {% if not is_volunteer and (who.secretary or who.leader_team_ids) %}<p><a class="btn" href="{{url_for('hongfa.books_add')}}" style="font-size:19px;font-weight:750">➕ 新增书籍</a></p>{% endif %}
        <div class="hb-tabs"><button id="hb-out-btn" class="on" onclick="hbTab('out')">📖 结缘扣书</button>{% if not is_volunteer %}<button id="hb-stock-btn" onclick="hbTab('stock')">📦 点货／补货</button>{% endif %}</div>
        <div id="hb-out-note" class="hf-muted">仅允许登记本组今天已发布且尚未结束的弘法活动；一次可选多种书。</div>
        <input id="hb-search" placeholder="🔍 输入书名立即搜索" oninput="hbDraw()">
        <div class="hb-filters" id="hb-filters"></div><div id="hb-picked"></div><div id="hb-books" class="hb-grid"></div>
        <div id="hb-bottom" class="hb-bottom"><button id="hb-submit" onclick="hbSubmit()">确认结缘（0本）</button></div>
        <form id="hb-form" method="post" class="hb-hidden"><input type="hidden" name="csrf_token" value="{{csrf}}"><input type="hidden" name="request_key" value="{{request_key}}"><input type="hidden" name="team_id" id="hb-form-team"><input type="hidden" name="mode" id="hb-mode"><input type="hidden" name="lines" id="hb-lines"></form>
        {% endif %}</div>
        <script>
        const hbItems={{items|tojson}},hbStock={{stock|tojson}},hbEvents={{events|tojson}},hbCategories={{categories|tojson}},hbFavorites={{favorites|tojson}};
        let hbCart={},hbView='out',hbFilter='常用',hbBusy=false;
        const hbEl=id=>document.getElementById(id);
        function hbCategory(name,id){const saved=hbCategories.find(x=>x.item_id===id);if(saved)return saved.category;if(name.includes('佛言佛语'))return '佛言佛语';if(name.includes('白话佛法'))return '白话佛法';if(/手册|指南|小本|入门/.test(name))return '手册／小本';if(/[\u4e00-\u9fff]/.test(name))return '其他中文';return '英文';}
        function hbTab(v){hbView=v;hbCart={};hbFilter='常用';hbEl('hb-out-btn').classList.toggle('on',v==='out');hbEl('hb-stock-btn')?.classList.toggle('on',v==='stock');hbEl('hb-out-note').classList.toggle('hb-hidden',v!=='out');hbDraw();}
        function hbSetFilter(v){hbFilter=v;hbDraw();}
        function hbChange(id,d){hbCart[id]=Math.max(0,(hbCart[id]||0)+d);if(!hbCart[id])delete hbCart[id];hbDraw();}
        function hbDraw(){if(!hbEl('hb-team'))return;const t=hbEl('hb-team').value,q=hbEl('hb-search').value.trim().toLowerCase();
          const filters=[...new Set(['常用','全部','佛言佛语','白话佛法','手册／小本','其他中文','英文',...hbCategories.map(x=>x.category)])];
          const box=hbEl('hb-filters');box.replaceChildren();
          for(const f of filters){const b=document.createElement('button');b.type='button';b.className=hbFilter===f?'on':'';b.textContent=(f==='常用'?'⭐ ':'')+f;b.onclick=()=>hbSetFilter(f);box.appendChild(b);}
          const fav=hbFavorites.filter(x=>String(x.team_id)===t).map(x=>x.item_id);
          const manage=hbEl('hb-manage-link');if(manage)manage.href={{url_for('hongfa.books_manage')|tojson}}+'?team_id='+encodeURIComponent(t);
          let arr=hbItems.filter(x=>q?x.title.toLowerCase().includes(q):(hbFilter==='全部'||hbCategory(x.title,x.id)===hbFilter||hbFilter==='常用'));
          if(!q&&hbFilter==='常用'){arr=arr.filter(x=>fav.includes(x.id));if(!arr.length)arr=hbItems.slice(0,12);else arr.sort((a,b)=>fav.indexOf(a.id)-fav.indexOf(b.id));}
          arr.sort((a,b)=>(hbCart[b.id]||0)-(hbCart[a.id]||0));
          const card=x=>{const n=hbCart[x.id]||0,bal=hbStock[t+':'+x.id];return `<div class="hb-book"><div class="hb-name"></div><div class="hb-sub">库存：${bal===undefined?'未点货':bal+'本'}</div><div class="hb-step"><button type="button" onclick="hbChange(${x.id},-1)">－</button><b>${n}</b><button type="button" onclick="hbChange(${x.id},1)">＋</button></div></div>`;};
          hbEl('hb-books').innerHTML=arr.map(card).join('')||'<p>没有找到书籍。</p>';
          hbEl('hb-books').querySelectorAll('.hb-name').forEach((e,i)=>e.textContent=arr[i].title);
          const chosen=hbItems.filter(x=>hbCart[x.id]>0);hbEl('hb-picked').innerHTML=chosen.length?'<div class="hb-picked"><b>✅ 已选书籍</b><div id="hb-picked-list"></div></div>':'';
          if(chosen.length){hbEl('hb-picked-list').textContent=chosen.map(x=>x.title+' × '+hbCart[x.id]).join('；');}
          const sum=Object.values(hbCart).reduce((a,b)=>a+b,0);hbEl('hb-submit').textContent=hbView==='stock'?'保存点货数量':'确认结缘（'+sum+'本）';hbEl('hb-submit').disabled=!sum||hbBusy;
          if(hbView==='stock')hbEl('hb-out-note').classList.add('hb-hidden');
        }
        function hbSubmit(){if(hbBusy)return;const t=hbEl('hb-team').value,lines=Object.entries(hbCart).filter(x=>x[1]>0).map(x=>({item_id:Number(x[0]),quantity:x[1]}));if(!lines.length)return;
          if(hbView==='out'&&!hbEvents[t]){alert('本组今天没有可登记的已发布活动。');return;}
          if(hbView==='out')for(const x of lines){const b=hbStock[t+':'+x.item_id];if(b===undefined){alert('请先点货：'+hbItems.find(i=>i.id===x.item_id).title);return;}if(b<x.quantity){alert('库存不足：'+hbItems.find(i=>i.id===x.item_id).title);return;}}
          if(hbView==='stock'){const next=prompt('请输入所选书籍的实际库存总数（一次只能选择一种书）');if(next===null)return;if(lines.length!==1||!/^(0|[1-9][0-9]*)$/.test(next)||Number(next)>1000000){alert('请只选择一本书，并输入正确的实际数量。');return;}lines[0].quantity=Number(next);}
          if(!confirm(hbView==='out'?'确认结缘并扣除 '+lines.reduce((s,x)=>s+x.quantity,0)+' 本？':'确认保存实际点货数量？'))return;
          hbBusy=true;hbEl('hb-form-team').value=t;hbEl('hb-mode').value=hbView;hbEl('hb-lines').value=JSON.stringify(lines);hbEl('hb-form').submit();}
        hbDraw();
        </script>''', who=who, csrf=csrf_token(), request_key=secrets.token_urlsafe(24),
                      teams=teams, items=items, stock=stock, events=event_map, categories=categories, favorites=favorites, message=message, error=error, is_volunteer=is_volunteer)

    @bp.route('/books', methods=['GET', 'POST'])
    @guard
    def books(who):
        if request.method == 'GET':
            return books_page(who, message=('✅ 新书籍已加入目录，请到点货／补货设置库存。' if request.args.get('added') == '1' else ('✅ 已保存库存记录。' if request.args.get('saved') == '1' else '')))
        try:
            team_id = int(request.form.get('team_id', 0))
            if team_id not in who['team_ids']:
                abort(403)
            mode = request.form.get('mode')
            if mode not in ('out', 'stock'):
                raise ValueError('操作无效。')
            if who.get('book_volunteer') and mode != 'out':
                abort(403)
            lines = json.loads(request.form.get('lines') or '[]')
            if not isinstance(lines, list) or not 1 <= len(lines) <= 60:
                raise ValueError('请选择书籍。')
            parsed = {}
            for line in lines:
                item_id = int(line['item_id'])
                qty = int(line['quantity'])
                if item_id in parsed or item_id < 1 or not 0 <= qty <= 1000000:
                    raise ValueError('数量或书籍重复。')
                if mode == 'out' and qty < 1:
                    raise ValueError('结缘数量必须大于零。')
                parsed[item_id] = qty
            if mode == 'stock' and len(parsed) != 1:
                raise ValueError('点货每次只能选择一本书。')
            reqkey = (request.form.get('request_key') or '').strip()
            if len(reqkey) < 20:
                raise ValueError('请求编号无效。')
            with get_conn() as conn:
                with conn.cursor(cursor_factory=RealDictCursor) as cur:
                    cur.execute('SELECT id FROM hongfa_teams WHERE id=%s FOR UPDATE', (team_id,))
                    if not cur.fetchone():
                        raise ValueError('组别无效。')
                    cur.execute('SELECT 1 FROM hongfa_stock_movements WHERE request_key LIKE %s LIMIT 1', (reqkey + ':%',))
                    if cur.fetchone():
                        raise ValueError('这一笔已保存，请勿重复提交。')
                    event_id = None
                    if mode == 'out':
                        cur.execute('''SELECT id FROM hongfa_events WHERE team_id=%s AND event_date=%s
                            AND status='published' AND closed_at IS NULL AND location_id IS NOT NULL FOR UPDATE''',
                            (team_id, datetime.now(TZ).date()))
                        row = cur.fetchone()
                        if not row:
                            raise ValueError('本组今天没有可登记的弘法活动。')
                        event_id = row['id']
                        if who.get('book_volunteer') and event_id != who['book_event']:
                            abort(403)
                    cur.execute('SELECT id FROM hongfa_items WHERE id=ANY(%s) AND active=true', (list(parsed),))
                    if len(cur.fetchall()) != len(parsed):
                        raise ValueError('书籍目录已变更，请刷新。')
                    for item_id, qty in parsed.items():
                        cur.execute('''SELECT COALESCE(SUM(quantity),0) AS balance,COUNT(*) AS entries
                            FROM hongfa_stock_movements WHERE team_id=%s AND item_id=%s''', (team_id, item_id))
                        row = cur.fetchone()
                        balance = int(row['balance'])
                        if mode == 'out':
                            if not row['entries']:
                                raise ValueError('这本书尚未点货。')
                            delta = -qty
                            if balance < qty:
                                raise ValueError('库存不足，请刷新后重试。')
                            typ = 'distribution'
                        else:
                            delta = qty - balance
                            typ = 'adjust_in' if delta >= 0 else 'adjust_out'
                        if delta == 0:
                            continue
                        cur.execute('''INSERT INTO hongfa_stock_movements
                            (team_id,item_id,event_id,quantity,movement_type,created_by,request_key)
                            VALUES (%s,%s,%s,%s,%s,%s,%s)''',
                            (team_id,item_id,event_id,delta,typ,who['id'],reqkey + ':' + str(item_id)))
            return redirect(url_for('hongfa.books', saved='1'))
        except (ValueError, TypeError, KeyError, json.JSONDecodeError) as exc:
            return books_page(who, error=str(exc))

    @bp.route('/books/manage', methods=['GET', 'POST'])
    @guard
    def books_manage(who):
        if who.get('book_volunteer') or not (who['secretary'] or who['team_ids']):
            abort(403)
        teams = db_query('SELECT id,name FROM hongfa_teams WHERE active=true ORDER BY id', fetchall=True) or []
        if not who['secretary']:
            teams = [t for t in teams if t['id'] in who['team_ids']]
        selected = request.values.get('team_id', type=int)
        if selected not in [t['id'] for t in teams]:
            selected = teams[0]['id'] if teams else None
        error = ''
        if request.method == 'POST':
            action = request.form.get('action')
            if action in ('category','rename') and not (who['secretary'] or who.get('leader_team_ids')): abort(403)
            item_id = request.form.get('item_id', type=int)
            if not db_query('SELECT id FROM hongfa_items WHERE id=%s AND active=true', (item_id,), fetchone=True):
                abort(400)
            if action == 'favorite':
                if selected is None: abort(403)
                with get_conn() as conn:
                    with conn.cursor() as cur:
                        cur.execute('SELECT id FROM hongfa_teams WHERE id=%s FOR UPDATE', (selected,))
                        cur.execute('SELECT sort_order FROM hongfa_book_favorites WHERE team_id=%s AND item_id=%s', (selected,item_id))
                        old = cur.fetchone()
                        if old:
                            cur.execute('DELETE FROM hongfa_book_favorites WHERE team_id=%s AND item_id=%s', (selected,item_id))
                        else:
                            cur.execute('SELECT COALESCE(MAX(sort_order),0)+1 FROM hongfa_book_favorites WHERE team_id=%s', (selected,))
                            order = cur.fetchone()[0]
                            cur.execute('INSERT INTO hongfa_book_favorites(team_id,item_id,sort_order) VALUES (%s,%s,%s)', (selected,item_id,order))
            elif action in ('up','down'):
                if selected is None: abort(403)
                with get_conn() as conn:
                    with conn.cursor() as cur:
                        cur.execute('SELECT id FROM hongfa_teams WHERE id=%s FOR UPDATE', (selected,))
                        cur.execute('SELECT item_id FROM hongfa_book_favorites WHERE team_id=%s ORDER BY sort_order,item_id', (selected,))
                        ids=[r[0] for r in cur.fetchall()]
                        if item_id in ids:
                            pos=ids.index(item_id); dest=pos+(-1 if action=='up' else 1)
                            if 0<=dest<len(ids): ids[pos],ids[dest]=ids[dest],ids[pos]
                            for n,i in enumerate(ids,1):
                                cur.execute('UPDATE hongfa_book_favorites SET sort_order=%s WHERE team_id=%s AND item_id=%s', (n,selected,i))
            elif action == 'category':
                category=' '.join((request.form.get('category') or '').split())
                if not 1<=len(category)<=40 or category in ('常用','全部'):
                    error='分类名称须为1至40字，不能使用常用或全部。'
                else:
                    db_query('''INSERT INTO hongfa_book_categories(item_id,category) VALUES (%s,%s)
                        ON CONFLICT (item_id) DO UPDATE SET category=EXCLUDED.category''', (item_id,category))
            elif action == 'rename':
                title=' '.join((request.form.get('title') or '').split())
                if not 1<=len(title)<=150:
                    error='书籍名称须为1至150字。'
                else:
                    with get_conn() as conn:
                        with conn.cursor() as cur:
                            cur.execute('SELECT pg_advisory_xact_lock(hashtextextended(%s,0))', ('hongfa_book_title:'+title.casefold(),))
                            cur.execute('SELECT id FROM hongfa_items WHERE lower(title)=lower(%s) AND id<>%s LIMIT 1', (title,item_id))
                            if cur.fetchone(): error='已有相同书名。'
                            else: cur.execute('UPDATE hongfa_items SET title=%s WHERE id=%s', (title,item_id))
            else: abort(400)
            if not error:
                return redirect(url_for('hongfa.books_manage',team_id=selected,saved=1))
        items=db_query('SELECT id,title FROM hongfa_items WHERE active=true ORDER BY title', fetchall=True) or []
        cats=db_query('SELECT item_id,category FROM hongfa_book_categories', fetchall=True) or []
        favs=db_query('SELECT item_id,sort_order FROM hongfa_book_favorites WHERE team_id=%s ORDER BY sort_order,item_id', (selected,),fetchall=True) or [] if selected else []
        cmap={x['item_id']:x['category'] for x in cats}
        fids={x['item_id']:x['sort_order'] for x in favs}
        for x in items:
            x['category']=cmap.get(x['id']) or ('佛言佛语' if '佛言佛语' in x['title'] else '白话佛法' if '白话佛法' in x['title'] else '手册／小本' if any(k in x['title'] for k in ('手册','指南','小本','入门')) else '其他中文' if any('\u4e00'<=c<='\u9fff' for c in x['title']) else '英文')
            x['fav_order']=fids.get(x['id'])
        return render('''<div class="card"><h2>⚙️ 管理书籍</h2>
          <p><a href="{{url_for('hongfa.books')}}">← 返回快速扣书</a>　{% if can_catalog %}<a href="{{url_for('hongfa.books_add')}}">➕ 新增书籍</a>{% endif %}</p>
          <p class="hf-muted">书名和分类三组共用；⭐ 常用及顺序只影响所选组。修改不会改变库存。</p>
          {% if error %}<p class="warn">{{error}}</p>{% endif %}{% if request_saved %}<p style="color:green">✅ 已保存</p>{% endif %}
          <label>选择弘法组（设置常用）</label><select onchange="location.href='?team_id='+this.value">{% for t in teams %}<option value="{{t.id}}" {% if t.id==selected %}selected{% endif %}>{{t.name}}</option>{% endfor %}</select>
          <input id="hb-manage-search" placeholder="🔍 搜索书籍名称" oninput="hbManageFilter()">
          <div id="hb-manage-list">{% for x in items %}<div class="hb-manage-row" data-title="{{x.title|lower}}" style="border-bottom:1px solid #ddd;padding:13px 0">
            <b>{{x.title}}</b> {% if x.fav_order %}<span style="color:#16854e">⭐ 第{{x.fav_order}}位</span>{% endif %}
            <div style="display:flex;flex-wrap:wrap;gap:5px;margin:8px 0">
            <form method="post"><input type="hidden" name="csrf_token" value="{{csrf}}"><input type="hidden" name="team_id" value="{{selected}}"><input type="hidden" name="item_id" value="{{x.id}}"><input type="hidden" name="action" value="favorite"><button style="padding:8px">{{'取消常用' if x.fav_order else '⭐ 设为常用'}}</button></form>
            {% if x.fav_order %}{% for a,label in [('up','⬆️'),('down','⬇️')] %}<form method="post"><input type="hidden" name="csrf_token" value="{{csrf}}"><input type="hidden" name="team_id" value="{{selected}}"><input type="hidden" name="item_id" value="{{x.id}}"><button name="action" value="{{a}}" style="padding:8px">{{label}}</button></form>{% endfor %}{% endif %}</div>
            {% if can_catalog %}<form method="post" style="display:flex;gap:6px;align-items:center"><input type="hidden" name="csrf_token" value="{{csrf}}"><input type="hidden" name="team_id" value="{{selected}}"><input type="hidden" name="item_id" value="{{x.id}}"><input type="hidden" name="action" value="category"><input name="category" value="{{x.category}}" maxlength="40" aria-label="分类" style="min-width:0;flex:1"><button style="white-space:nowrap">保存分类</button></form>
            <form method="post" style="display:flex;gap:6px;align-items:center"><input type="hidden" name="csrf_token" value="{{csrf}}"><input type="hidden" name="team_id" value="{{selected}}"><input type="hidden" name="item_id" value="{{x.id}}"><input type="hidden" name="action" value="rename"><input name="title" value="{{x.title}}" maxlength="150" aria-label="书籍名称" style="min-width:0;flex:1"><button style="white-space:nowrap">改书名</button></form>{% endif %}
          </div>{% endfor %}</div></div>
          <script>function hbManageFilter(){var q=document.getElementById('hb-manage-search').value.trim().toLowerCase();document.querySelectorAll('.hb-manage-row').forEach(function(r){r.style.display=r.dataset.title.includes(q)?'':'none';});}</script>''',csrf=csrf_token(),teams=teams,selected=selected,items=items,error=error,can_catalog=bool(who['secretary'] or who.get('leader_team_ids')),request_saved=request.args.get('saved')=='1')

    @bp.route('/books/add', methods=['GET', 'POST'])
    @guard
    def books_add(who):
        # 只有负责人或秘书可维护共用书籍目录；签到义工不能新增。
        if who.get('book_volunteer') or not (who['secretary'] or who.get('leader_team_ids')):
            abort(403)
        error = ''
        if request.method == 'POST':
            title = ' '.join((request.form.get('title') or '').split())
            if not title or len(title) > 150:
                error = '请输入1至150字的书籍名称。'
            else:
                # 共用目录，按忽略大小写的书名检查重复；不修改各组库存。
                with get_conn() as conn:
                    with conn.cursor() as cur:
                        cur.execute('SELECT pg_advisory_xact_lock(hashtextextended(%s,0))', ('hongfa_book_title:' + title.casefold(),))
                        cur.execute('SELECT id FROM hongfa_items WHERE lower(title)=lower(%s) LIMIT 1', (title,))
                        if cur.fetchone():
                            error = '这本书已经存在，请直接搜索原有书籍。'
                        else:
                            cur.execute('INSERT INTO hongfa_items(title) VALUES (%s)', (title,))
                if not error:
                    return redirect(url_for('hongfa.books', added='1'))
        return render('''<div class="card"><h2>➕ 新增书籍</h2>
            <p><a href="{{url_for('hongfa.books')}}">← 返回书籍管理</a></p>
            <p class="hf-muted">三组共用书籍名称，但库存分别计算。新增后请到「点货／补货」输入实际数量。</p>
            {% if error %}<p class="warn" role="alert">{{error}}</p>{% endif %}
            <form method="post"><input type="hidden" name="csrf_token" value="{{csrf}}">
            <label for="book-title">📖 书籍名称</label>
            <input id="book-title" name="title" maxlength="150" required autofocus placeholder="例如：佛言佛语 9册">
            <button type="submit" style="font-size:20px;font-weight:750">💾 保存新书籍</button></form></div>''',
            csrf=csrf_token(), error=error)
