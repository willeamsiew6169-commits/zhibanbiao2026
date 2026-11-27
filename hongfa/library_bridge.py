"""Two-stage book transfer. All stock writes use the existing two ledgers."""
import os
import re
import secrets
from uuid import UUID, uuid4

from flask import abort, flash, redirect, render_template_string, request, session, url_for
from psycopg2.extras import RealDictCursor

from db import db_query, get_conn

WRITE_FLAG = 'HONGFA_LIBRARY_TRANSFERS_WRITE'
CONFIRM_FLAG = 'HONGFA_LIBRARY_CONFIRM_WRITE'
WRITE_VALUE = 'enabled-after-excel-reconciliation'


def writes_enabled():
    return os.environ.get(WRITE_FLAG) == WRITE_VALUE


def confirmation_enabled():
    return os.environ.get(CONFIRM_FLAG) == WRITE_VALUE


def lock_catalog_shared(cur):
    cur.execute("SELECT pg_advisory_xact_lock_shared(hashtext('hongfa_library_catalog'))")


def installed():
    row = db_query("""SELECT has_table_privilege(to_regclass('public.hongfa_library_transfers'), 'SELECT')
                      AND has_table_privilege(to_regclass('public.hongfa_library_book_map'), 'SELECT') AS installed""",
                   fetchone=True)
    return bool(row and row['installed'])


def pending_count():
    if not installed():
        return None
    return db_query("SELECT COUNT(*) AS n FROM hongfa_library_transfers WHERE status='pending'",
                    fetchone=True)['n']


def quantity_lines(form):
    items = {}
    for name, values in form.lists():
        if not name.startswith('qty_'):
            continue
        if len(values) != 1 or not re.fullmatch(r'qty_[1-9][0-9]{0,17}', name):
            raise ValueError('书籍编号重复或无效。')
        value = values[0].strip()
        if value in ('', '0'):
            continue
        if not re.fullmatch(r'[1-9][0-9]{0,5}', value) or int(value) > 100000:
            raise ValueError('领取数量须为 1 至 100000 的整数。')
        items[int(name[4:])] = int(value)
    if not 1 <= len(items) <= 100:
        raise ValueError('每次请选择 1 至 100 种书籍。')
    return items


def _lock_items(cur, ids):
    cur.execute("""SELECT id,item_code,name,category,balance,is_active FROM library_items
                   WHERE id=ANY(%s) ORDER BY id FOR UPDATE""", (sorted(ids),))
    return {row['id']: row for row in cur.fetchall()}


def _mapping_matches(book, mapping):
    return (book['item_code'], book['name'], book['category']) == (
        mapping['item_code'], mapping['book_name'], mapping['category'])


def _mapped_item(cur, book):
    cur.execute("""SELECT m.*, i.active FROM hongfa_library_book_map m
                   JOIN hongfa_items i ON i.id=m.hongfa_item_id
                   WHERE m.library_item_id=%s""", (book['id'],))
    mapping = cur.fetchone()
    if mapping:
        if not _mapping_matches(book, mapping) or not mapping['active']:
            raise ValueError(f"《{book['name']}》的书号、名称、分类或弘法目录状态已变更，请先对账。")
        return mapping['hongfa_item_id']
    cur.execute("""SELECT 1 FROM library_items
                   WHERE name=%s AND id<>%s AND is_active=true LIMIT 1""", (book['name'], book['id']))
    if cur.fetchone():
        raise ValueError(f"藏经阁有重名书籍《{book['name']}》，请先核对对应关系。")
    cur.execute("""INSERT INTO hongfa_items(title) VALUES (%s) ON CONFLICT (title) DO NOTHING""",
                (book['name'],))
    cur.execute("SELECT id,active FROM hongfa_items WHERE title=%s FOR UPDATE", (book['name'],))
    item = cur.fetchone()
    if not item or not item['active']:
        raise ValueError(f"《{book['name']}》的同名弘法书籍不可用，请先核对。")
    cur.execute("""INSERT INTO hongfa_library_book_map(library_item_id,hongfa_item_id,item_code,book_name,category)
                   VALUES (%s,%s,%s,%s,%s) ON CONFLICT DO NOTHING RETURNING library_item_id""",
                (book['id'], item['id'], book['item_code'], book['name'], book['category']))
    if not cur.fetchone():
        raise ValueError(f"《{book['name']}》已经对应其他书籍，请先核对。")
    return item['id']


def register_receipt(team_id, actor_id, key, quantities):
    """Make a receipt and reserve library stock without changing its balance."""
    if not 1 <= len(quantities) <= 100 or any(type(qty) is not int or not 1 <= qty <= 100000
                                               for qty in quantities.values()):
        raise ValueError('领取数量无效。')
    with get_conn() as conn, conn.cursor(cursor_factory=RealDictCursor) as cur:
        lock_catalog_shared(cur)
        cur.execute("SELECT pg_advisory_xact_lock(hashtext('hongfa_library_receipt'), hashtext(%s))", (key,))
        cur.execute("SELECT id FROM hongfa_teams WHERE id=%s AND active=true FOR UPDATE", (team_id,))
        if not cur.fetchone():
            raise ValueError('组别已停用。')
        cur.execute("""SELECT team_id, received_by, library_item_id, quantity
                       FROM hongfa_library_transfers WHERE request_key=%s""", (key,))
        saved = cur.fetchall()
        if saved:
            if (any(row['team_id'] != team_id or row['received_by'] != actor_id for row in saved)
                    or {row['library_item_id']: row['quantity'] for row in saved} != quantities):
                raise ValueError('该请求编号已使用，请重新打开页面。')
            return False
        books = _lock_items(cur, quantities)
        if len(books) != len(quantities):
            raise ValueError('书籍已变更，请重新打开页面。')
        cur.execute("""SELECT library_item_id, SUM(quantity) AS quantity
                       FROM hongfa_library_transfers WHERE status='pending' AND library_item_id=ANY(%s)
                       GROUP BY library_item_id""", (sorted(quantities),))
        pending = {row['library_item_id']: row['quantity'] for row in cur.fetchall()}
        for item_id, book in books.items():
            if not book['is_active'] or (book['balance'] or 0) - pending.get(item_id, 0) < quantities[item_id]:
                raise ValueError(f"《{book['name']}》可领取库存不足或已停用。")
        # Stable title ordering also serializes concurrent shared-catalog inserts.
        for item_id, book in sorted(books.items(), key=lambda pair: (pair[1]['name'], pair[0])):
            hongfa_item_id = _mapped_item(cur, book)
            qty = quantities[item_id]
            cur.execute("""INSERT INTO hongfa_library_transfers
                           (request_key,team_id,library_item_id,item_code,book_name,category,quantity,received_by)
                           VALUES (%s,%s,%s,%s,%s,%s,%s,%s)""",
                        (key, team_id, item_id, book['item_code'], book['name'], book['category'], qty, actor_id))
            cur.execute("""INSERT INTO hongfa_stock_movements
                           (team_id,item_id,event_id,quantity,movement_type,created_by,request_key)
                           VALUES (%s,%s,NULL,%s,'received',%s,%s)""",
                        (team_id, hongfa_item_id, qty, actor_id, f'library:{key}:{item_id}'))
    return True


def confirm_receipts(ids, admin, label, write_audit):
    with get_conn() as conn, conn.cursor(cursor_factory=RealDictCursor) as cur:
        lock_catalog_shared(cur)
        cur.execute("SELECT id,library_item_id FROM hongfa_library_transfers WHERE id=ANY(%s)", (ids,))
        refs = cur.fetchall()
        if len(refs) != len(ids):
            raise ValueError('部分领取记录已不存在。')
        books = _lock_items(cur, {row['library_item_id'] for row in refs})
        cur.execute("""SELECT t.* FROM hongfa_library_transfers t
                       WHERE t.id=ANY(%s) ORDER BY t.library_item_id,t.id FOR UPDATE OF t""", (ids,))
        rows = cur.fetchall()
        confirmed = 0
        for row in rows:
            if row['status'] == 'confirmed':
                continue
            book = books[row['library_item_id']]
            if not book['is_active'] or not _mapping_matches(book, row):
                raise ValueError(f"《{row['book_name']}》书籍资料已变更，请先对账。")
            qty = row['quantity']
            before = book['balance'] or 0
            if before < qty:
                raise ValueError(f"《{row['book_name']}》账面库存不足，请先对账。")
            cur.execute("""INSERT INTO library_transactions
                           (item_code,item_name,transaction_type,quantity,volunteer_id,volunteer_name,handled_by,remark)
                           VALUES (%s,%s,'out',%s,%s,
                             COALESCE((SELECT name FROM volunteers WHERE id::text=%s LIMIT 1),%s),%s,%s)
                           RETURNING id""",
                        (book['item_code'], book['name'], qty, row['received_by'], row['received_by'],
                         row['received_by'], label, f"弘法领取 #{row['id']}"))
            transaction_id = cur.fetchone()['id']
            cur.execute("""UPDATE hongfa_library_transfers SET status='confirmed',confirmed_by=%s,
                           confirmed_at=now(),library_transaction_id=%s WHERE id=%s""",
                        (str(admin['id']), transaction_id, row['id']))
            cur.execute("UPDATE library_items SET balance=balance-%s WHERE id=%s", (qty, book['id']))
            write_audit(cur, 'HONGFA_TRANSFER_OUT', item_code=book['item_code'], item_name=book['name'],
                        quantity_change=-qty, balance_before=before, balance_after=before-qty,
                        transaction_id=transaction_id, detail=f"确认弘法领取 #{row['id']}",
                        metadata={'transfer_id': row['id'], 'team_id': row['team_id']}, admin=admin)
            book['balance'] = before - qty
            confirmed += 1
    return confirmed


def _history(team_ids=None, pending=False, offset=0):
    where, params = [], []
    if team_ids is not None:
        where.append('r.team_id=ANY(%s)')
        params.append(team_ids)
    if pending:
        where.append("r.status='pending'")
    return db_query("""SELECT r.*,t.name AS team_name,
                       to_char(r.received_at AT TIME ZONE 'Asia/Kuala_Lumpur','YYYY-MM-DD HH24:MI') AS received_my,
                       to_char(r.confirmed_at AT TIME ZONE 'Asia/Kuala_Lumpur','YYYY-MM-DD HH24:MI') AS confirmed_my,
                       COALESCE(v.name,r.received_by) AS receiver_name,
                       COALESCE(cv.name,a.volunteer_id,r.confirmed_by) AS confirmer_name
                       FROM hongfa_library_transfers r JOIN hongfa_teams t ON t.id=r.team_id
                       LEFT JOIN volunteers v ON v.id::text=r.received_by
                       LEFT JOIN library_admin_users a ON a.id::text=r.confirmed_by
                       LEFT JOIN volunteers cv ON cv.id::text=a.volunteer_id::text """
                    + (' WHERE ' + ' AND '.join(where) if where else '')
                    + ' ORDER BY r.received_at DESC,r.id DESC LIMIT 101 OFFSET %s', (*params, offset), fetchall=True) or []


HISTORY = """
<div class="bridge-scroll"><table class="bridge-table">
<thead><tr>{% if pending and can_confirm %}<th>确认</th>{% endif %}<th>书籍／组别</th><th>数量</th><th>领取人／时间</th><th>确认状态</th></tr></thead>
<tbody>{% for row in records %}<tr>
{% if pending and can_confirm %}<td class="bridge-checkbox"><input type="checkbox" name="transfer_id" value="{{ row.id }}" aria-label="选择 {{ row.book_name }}"></td>{% endif %}
<td class="bridge-book">{{ row.book_name }}<br>{{ row.team_name }}</td><td data-label="数量">{{ row.quantity }}</td>
<td data-label="领取人／时间">{{ row.receiver_name }}<br>{{ row.received_my }}</td>
<td data-label="确认状态">{% if row.status == 'confirmed' %}已确认<br>{{ row.confirmer_name }} · {{ row.confirmed_my }}{% else %}待确认{% endif %}</td>
</tr>{% else %}<tr><td colspan="5">暂无记录</td></tr>{% endfor %}</tbody></table></div>
"""

STYLE = """
<style>
.bridge-scroll{overflow-x:auto}.bridge-table{width:100%;border-collapse:collapse;font-size:16px}
.bridge-table th,.bridge-table td{text-align:left;padding:10px;border-bottom:1px solid #ddd;vertical-align:top}
.bridge-form{display:flex;flex-wrap:wrap;align-items:end;gap:10px;margin:12px 0}
.bridge-form label{display:grid;gap:4px}.bridge-qty{width:6rem}
.bridge-form input,.bridge-form select,.bridge-table input,.bridge-action{font-size:16px;min-height:44px}
.bridge-table input[type=checkbox]{width:24px}.bridge-action{margin:12px 0}
@media(max-width:600px){
  .bridge-catalog thead{position:absolute;clip-path:inset(50%);width:1px;height:1px;overflow:hidden}
  .bridge-catalog tr{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));border-bottom:1px solid #ddd}
  .bridge-catalog td{display:block;border:0;padding:6px}.bridge-catalog .bridge-book{grid-column:1/-1}
  .bridge-catalog td[data-label]::before{content:attr(data-label) ' ';font-weight:700}
}
</style>
"""


def register_hongfa_library(bp, actor, render, csrf_token, valid_csrf):
    @bp.route('/library-books', methods=['GET', 'POST'])
    def library_books():
        who = actor()
        if not who:
            abort(403)
        can_record = bool(who['team_ids']) and not who['secretary']
        if request.method == 'POST':
            if not can_record or not writes_enabled():
                abort(403)
            if not request.form.get('csrf_token', '').isascii() or not valid_csrf():
                abort(400)
        ready = installed()
        if not ready:
            return render('<div class="card"><h2>藏经阁</h2><p>库存连接尚未安装。</p><a href="/hongfa/dashboard">返回</a></div>'), 503
        teams = db_query('SELECT id,name FROM hongfa_teams WHERE active=true ORDER BY name', fetchall=True) or []
        teams = [team for team in teams if team['id'] in who['team_ids']]
        ids = [team['id'] for team in teams]
        selected = request.values.get('team_id', type=int) or (ids[0] if ids else None)
        if not who['secretary'] and selected not in ids:
            abort(403)
        error, key = '', str(uuid4())
        if request.method == 'POST':
            try:
                key = str(UUID(request.form.get('request_key', '')))
                register_receipt(selected, who['id'], key, quantity_lines(request.form))
                return redirect(url_for('hongfa.library_books', team_id=selected, saved=1))
            except ValueError as exc:
                error = str(exc)
        books = []
        query = request.args.get('q', '').strip()[:100]
        if not who['secretary']:
            books = db_query("""SELECT b.id,b.name,b.category,COALESCE(b.balance,0) AS balance,
                               COALESCE(p.qty,0) AS pending,GREATEST(COALESCE(b.balance,0)-COALESCE(p.qty,0),0) AS available,
                               COALESCE(s.qty,0) AS team_balance
                               FROM library_items b
                               LEFT JOIN hongfa_library_book_map m ON m.library_item_id=b.id
                               LEFT JOIN hongfa_items h ON h.title=b.name AND h.active=true
                               LEFT JOIN hongfa_book_favorites f ON f.item_id=COALESCE(m.hongfa_item_id,h.id) AND f.team_id=%s
                               LEFT JOIN LATERAL (SELECT SUM(quantity) AS qty FROM hongfa_library_transfers
                                   WHERE library_item_id=b.id AND status='pending') p ON true
                               LEFT JOIN LATERAL (SELECT SUM(quantity) AS qty FROM hongfa_stock_movements
                                   WHERE item_id=COALESCE(m.hongfa_item_id,h.id) AND team_id=%s) s ON true
                               WHERE b.is_active=true AND (%s='' OR b.name ILIKE %s OR b.category ILIKE %s)
                               ORDER BY f.sort_order NULLS LAST,b.name,b.id LIMIT 150""",
                             (selected, selected, query, f'%{query}%', f'%{query}%'), fetchall=True) or []
        page = max(0, request.args.get('page', 0, type=int))
        records = _history(None if who['secretary'] else [selected], offset=page * 100)
        return render(STYLE + """
          <div class="card"><h2>📚 藏经阁</h2><a href="/hongfa/dashboard">返回管理首页</a>
          {% if error %}<p role="alert">{{ error }}</p>{% elif saved %}<p role="status">已登记领取。</p>{% endif %}
          {% if not secretary %}
          <form method="get" class="bridge-form"><label>弘法组别<select name="team_id">
          {% for t in teams %}<option value="{{t.id}}" {% if t.id==selected %}selected{% endif %}>{{t.name}}</option>{% endfor %}
          </select></label><label>书名或分类<input type="search" name="q" value="{{query}}"></label><button>搜索</button></form>
          {% if not enabled %}<p>库存对账完成前暂停领取写入。</p>{% endif %}
          <form method="post" action="{{url_for('hongfa.library_books',team_id=selected,q=query)}}">
          <input type="hidden" name="csrf_token" value="{{csrf}}"><input type="hidden" name="request_key" value="{{key}}">
          <div class="bridge-scroll"><table class="bridge-table bridge-catalog"><thead><tr><th>书籍／分类</th><th>账面</th><th>待确认</th><th>可领取</th><th>本组库存</th>
          {% if enabled %}<th>领取数</th>{% endif %}</tr></thead><tbody>
          {% for b in books %}<tr><td class="bridge-book"><span class="bridge-book-name" data-library-item-id="{{b.id}}">{{b.name}}</span><br>{{b.category or '—'}}</td>
          <td data-label="账面">{{b.balance}}</td><td data-label="待确认">{{b.pending}}</td>
          <td data-label="可领取">{{b.available}}</td><td data-label="本组库存">{{b.team_balance}}</td>
          {% if enabled %}<td data-label="领取数"><input class="bridge-qty" type="number" inputmode="numeric" min="0" max="{{[b.available,100000]|min}}" step="1"
          name="qty_{{b.id}}" aria-label="{{b.name}} 领取数量" {% if not b.available %}disabled{% endif %}></td>{% endif %}</tr>
          {% else %}<tr><td colspan="6">未找到书籍</td></tr>{% endfor %}</tbody></table></div>
          {% if enabled and books %}<button class="bridge-action">确认领取</button>{% endif %}</form>{% endif %}
          <h3>历史领取记录</h3>""" + HISTORY + """
          {% if more %}<a href="{{url_for('hongfa.library_books',team_id=selected,q=query,page=page+1)}}">下一页</a>{% endif %}</div>
          """, secretary=who['secretary'], teams=teams, selected=selected, books=books, query=query,
                      enabled=writes_enabled(), csrf=csrf_token(), key=key,
                      records=records[:100], more=len(records) > 100, page=page, error=error, saved=request.args.get('saved'))


def register_library_confirmation(bp, get_admin, can, label, audit):
    @bp.route('/admin/hongfa-transfers', methods=['GET', 'POST'])
    @bp.route('/admin/hongfa-transfer-history', endpoint='hongfa_transfer_history', methods=['GET'])
    def hongfa_transfers():
        admin = get_admin()
        pending = request.endpoint == 'library.hongfa_transfers'
        may_confirm = admin and can('can_batch_out')
        if not admin or not (may_confirm if pending else may_confirm or can('can_view_records')):
            abort(403)
        if not installed():
            abort(503)
        token = session.setdefault('library_transfer_csrf', secrets.token_urlsafe(32))
        error = ''
        if request.method == 'POST':
            if not confirmation_enabled() or not may_confirm:
                abort(403)
            if not secrets.compare_digest(token.encode(), request.form.get('csrf_token', '').encode()):
                abort(400)
            try:
                raw = request.form.getlist('transfer_id')
                if not 1 <= len(raw) <= 100 or any(not re.fullmatch(r'[1-9][0-9]{0,17}', value) for value in raw):
                    raise ValueError('请选择 1 至 100 条领取记录。')
                count = confirm_receipts(sorted({int(value) for value in raw}), admin, label(admin), audit)
                flash(f'已确认 {count} 笔。', 'hongfa_transfer')
                return redirect(url_for('library.hongfa_transfers'))
            except ValueError as exc:
                error = str(exc)
        page = max(0, request.args.get('page', 0, type=int))
        records = _history(pending=pending, offset=page * 100)
        return render_template_string("""<!doctype html><html lang="zh-CN"><head><meta charset="utf-8">
          <meta name="viewport" content="width=device-width, initial-scale=1"><title>弘法领取</title>
          """ + STYLE + """<link rel="stylesheet" href="/static/css/toolbox.css">
          <link rel="stylesheet" href="/static/css/library-inventory.css">
          </head><body class="library-inventory"><main class="inv">
          <header class="inv-head"><h1>{{ '弘法领取确认' if pending else '弘法领取历史' }}</h1></header>
          <nav class="inv-tabs" aria-label="弘法领取管理">
          {% if may_confirm %}<a {% if pending %}aria-current="page"{% endif %} href="{{ url_for('library.hongfa_transfers') }}">待确认</a>{% endif %}
          <a {% if not pending %}aria-current="page"{% endif %} href="{{ url_for('library.hongfa_transfer_history') }}">历史记录</a></nav>
          {% for msg in get_flashed_messages(category_filter=['hongfa_transfer']) %}<p class="inv-note" role="status">{{msg}}</p>{% endfor %}
          {% if error %}<p class="alert alert-error" role="alert">{{error}}</p>{% endif %}
          {% if pending and not can_confirm %}<p class="inv-note">库存对账完成前暂停确认。</p>{% endif %}
          <div class="inv-card">
          <form method="post"><input type="hidden" name="csrf_token" value="{{token}}">""" + HISTORY + """
          {% if pending and can_confirm and records %}<button class="btn-tool btn-green bridge-action">确认所选领取</button>{% endif %}</form></div>
          {% if more %}<nav class="inv-pages"><a href="?page={{page+1}}">下一页</a></nav>{% endif %}
          <a class="inv-back" href="/library/admin/home">↩ 返回负责人首页</a></main></body></html>
          """, pending=pending, may_confirm=may_confirm, can_confirm=may_confirm and confirmation_enabled(),
                                      page=page, more=len(records) > 100,
                                      records=records[:100], token=token, error=error)
