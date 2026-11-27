"""Offline stock/permission/report checks. Add --all for the existing full suite."""
from concurrent.futures import ThreadPoolExecutor
from contextlib import closing, contextmanager
from datetime import datetime, timezone
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from io import BytesIO
from pathlib import Path
import os
import re
import sqlite3
import subprocess
import sys
from tempfile import TemporaryDirectory
from threading import Thread
from types import ModuleType
import unittest
from unittest.mock import MagicMock, Mock, patch
from uuid import uuid4

from flask import Blueprint, Flask, render_template_string
from openpyxl import load_workbook
from werkzeug.datastructures import MultiDict


ROOT = Path(__file__).resolve().parent
blocked_db = ModuleType('db')
blocked_db.get_conn = blocked_db.db_query = Mock(side_effect=AssertionError('No production database'))
attendance_stub = ModuleType('attendance_service')
attendance_stub.find_volunteer = attendance_stub.verify_pin_for_volunteer = Mock(
    side_effect=AssertionError('No production attendance'))
with patch.dict(sys.modules, {'db': blocked_db, 'attendance_service': attendance_stub}):
    from hongfa import library_bridge as bridge
    from hongfa import hongfa_book_reports as reports
    from hongfa import hongfa_books as books
    from hongfa import phase3

PGLITE = ROOT / 'output/hongfa-bridge-pg/package/dist'


class Cursor:
    """Translate the limited PostgreSQL DML to a disposable SQLite fixture."""
    def __init__(self, cursor):
        self.cursor = cursor

    def __enter__(self):
        return self

    def __exit__(self, *args):
        self.cursor.close()

    def execute(self, sql, args=()):
        sql = ' '.join(sql.split())
        if sql.startswith('SELECT pg_advisory_xact_lock'):
            return
        sql = re.sub(r' FOR UPDATE(?: OF \w+)?', '', sql)
        sql = re.sub(r'(\w+)::text', r'CAST(\1 AS text)', sql)
        sql = sql.replace('now()', 'CURRENT_TIMESTAMP')
        sql, bound = self.bind(sql, args)
        self.cursor.execute(sql, bound)

    @staticmethod
    def bind(sql, args):
        # Only the service's id lists use ANY. Values remain parameters.
        parts = re.split(r'(=ANY\(%s\)|%s)', sql)
        values = iter(args)
        bound = []
        for i, part in enumerate(parts):
            if part == '=ANY(%s)':
                value = next(values)
                parts[i] = ' IN (' + ','.join('?' for _ in value) + ')'
                bound.extend(value)
            elif part == '%s':
                parts[i] = '?'
                bound.append(next(values))
        return ''.join(parts), bound

    def fetchone(self):
        row = self.cursor.fetchone()
        return dict(row) if row is not None else None

    def fetchall(self):
        return [dict(row) for row in self.cursor.fetchall()]


class Connection:
    def __init__(self, db):
        self.db = db

    def cursor(self, **kwargs):
        return Cursor(self.db.cursor())


class StockTests(unittest.TestCase):
    def setUp(self):
        tmp = TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.path = Path(tmp.name) / 'books.sqlite'
        with closing(sqlite3.connect(self.path)) as conn, conn:
            conn.executescript("""
            CREATE TABLE library_items(id INTEGER PRIMARY KEY,item_code TEXT UNIQUE,name TEXT,
              category TEXT,balance INTEGER,is_active BOOLEAN);
            CREATE TABLE hongfa_teams(id INTEGER PRIMARY KEY,active BOOLEAN);
            CREATE TABLE hongfa_items(id INTEGER PRIMARY KEY,title TEXT UNIQUE,active BOOLEAN DEFAULT TRUE);
            CREATE TABLE hongfa_library_book_map(library_item_id INTEGER PRIMARY KEY,hongfa_item_id INTEGER UNIQUE,
              item_code TEXT,book_name TEXT,category TEXT);
            CREATE TABLE hongfa_library_transfers(id INTEGER PRIMARY KEY,request_key TEXT,team_id INTEGER,
              library_item_id INTEGER,item_code TEXT,book_name TEXT,category TEXT,quantity INTEGER,received_by TEXT,status TEXT DEFAULT 'pending',
              confirmed_by TEXT,confirmed_at TEXT,library_transaction_id INTEGER,UNIQUE(request_key,library_item_id));
            CREATE TABLE hongfa_stock_movements(team_id INTEGER,item_id INTEGER,event_id INTEGER,quantity INTEGER,
              movement_type TEXT CHECK(movement_type IN ('received','distributed','adjustment')),
              created_by TEXT,request_key TEXT UNIQUE);
            CREATE TABLE library_transactions(id INTEGER PRIMARY KEY,item_code TEXT,item_name TEXT,
              transaction_type TEXT,quantity INTEGER,volunteer_id TEXT,volunteer_name TEXT,handled_by TEXT,remark TEXT);
            CREATE TABLE volunteers(id TEXT,name TEXT);
            CREATE TABLE audit(transaction_id INTEGER,quantity_change INTEGER);
            INSERT INTO hongfa_teams VALUES (1,true),(2,true),(3,false);
            INSERT INTO library_items VALUES (10,'A','心经','经书',100,true),(20,'B','新书','书籍',40,true);
            INSERT INTO hongfa_items VALUES (9,'心经',true);
            INSERT INTO volunteers VALUES ('CHE-1','测试负责人');
            INSERT INTO hongfa_stock_movements VALUES (1,9,NULL,10,'received','CHE-9','old');
            """)
        hook = patch.object(bridge, 'get_conn', self.connect)
        hook.start()
        self.addCleanup(hook.stop)

    @contextmanager
    def connect(self):
        with closing(sqlite3.connect(self.path, timeout=5)) as conn, conn:
            conn.row_factory = sqlite3.Row
            conn.execute('BEGIN IMMEDIATE')
            yield Connection(conn)

    def rows(self, sql):
        with closing(sqlite3.connect(self.path)) as conn:
            return conn.execute(sql).fetchall()

    def receipt(self, quantities, team=1, key=None):
        return bridge.register_receipt(team, 'CHE-1', key or str(uuid4()), quantities)

    def confirm(self, ids):
        def audit(cur, action, *, transaction_id, quantity_change, **kwargs):
            self.assertEqual(action, 'HONGFA_TRANSFER_OUT')
            cur.execute('INSERT INTO audit VALUES (%s,%s)', (transaction_id, quantity_change))
        return bridge.confirm_receipts(ids, {'id': 88}, '藏经阁负责人', audit)

    def test_existing_new_and_batch_receipt_keeps_library_balance(self):
        self.assertTrue(self.receipt({10: 20, 20: 5}))
        self.assertEqual(self.rows('SELECT title FROM hongfa_items ORDER BY id'), [('心经',), ('新书',)])
        self.assertEqual(self.rows('SELECT library_item_id,hongfa_item_id FROM hongfa_library_book_map ORDER BY 1'),
                         [(10, 9), (20, 10)])
        self.assertEqual(self.rows('SELECT balance FROM library_items ORDER BY id'), [(100,), (40,)])
        self.assertEqual(self.rows('SELECT SUM(quantity) FROM hongfa_stock_movements GROUP BY item_id ORDER BY item_id'),
                         [(30,), (5,)])
        self.assertEqual(self.rows("SELECT library_item_id,SUM(quantity) FROM hongfa_library_transfers WHERE status='pending' GROUP BY 1"),
                         [(10, 20), (20, 5)])
        self.assertEqual(self.rows("SELECT b.balance-SUM(r.quantity) FROM library_items b JOIN hongfa_library_transfers r ON r.library_item_id=b.id WHERE b.id=10 AND r.status='pending'"),
                         [(80,)])

    def test_batch_confirmation_and_replays_use_existing_out_and_audit(self):
        key = str(uuid4())
        self.assertTrue(self.receipt({10: 20, 20: 4}, key=key))
        self.assertFalse(self.receipt({10: 20, 20: 4}, key=key))
        with self.assertRaisesRegex(ValueError, '请求编号'):
            self.receipt({10: 20, 20: 4}, team=2, key=key)
        self.assertEqual(self.confirm([1]), 1)
        self.assertEqual(self.confirm([1, 2]), 1)
        self.assertEqual(self.confirm([1, 2]), 0)
        self.assertEqual(self.rows('SELECT balance FROM library_items ORDER BY id'), [(80,), (36,)])
        self.assertEqual(self.rows('SELECT transaction_type,quantity,volunteer_name FROM library_transactions ORDER BY id'),
                         [('out', 20, '测试负责人'), ('out', 4, '测试负责人')])
        self.assertEqual(self.rows('SELECT quantity_change FROM audit ORDER BY transaction_id'), [(-20,), (-4,)])
        self.assertEqual(self.rows('SELECT status,confirmed_by FROM hongfa_library_transfers ORDER BY id'),
                         [('confirmed', '88'), ('confirmed', '88')])
        self.assertEqual(self.rows('SELECT SUM(quantity) FROM hongfa_stock_movements'), [(34,)])

    def test_two_teams_compete_for_the_same_available_books(self):
        def take(team):
            try:
                return self.receipt({10: 60}, team=team)
            except ValueError:
                return False
        with ThreadPoolExecutor(max_workers=2) as pool:
            self.assertCountEqual(pool.map(take, (1, 2)), [True, False])
        self.receipt({10: 40}, team=2)
        totals = dict(self.rows('SELECT team_id,SUM(quantity) FROM hongfa_stock_movements GROUP BY team_id'))
        self.assertIn(totals, ({1: 70, 2: 40}, {1: 10, 2: 100}))
        with self.assertRaisesRegex(ValueError, '库存不足'):
            self.receipt({10: 1}, team=1)
        self.assertEqual(self.rows('SELECT balance FROM library_items WHERE id=10'), [(100,)])

    def test_changed_excel_identity_blocks_receipts_and_confirmation(self):
        self.receipt({10: 10})
        with closing(sqlite3.connect(self.path)) as conn, conn:
            conn.execute("UPDATE library_items SET name='其他法物' WHERE id=10")
        with self.assertRaisesRegex(ValueError, '对账'):
            self.confirm([1])
        with self.assertRaisesRegex(ValueError, '对账'):
            self.receipt({10: 10})
        self.assertEqual(self.rows('SELECT status FROM hongfa_library_transfers'), [('pending',)])
        self.assertEqual(self.rows('SELECT COUNT(*) FROM library_transactions'), [(0,)])

    def test_exact_title_bound_to_another_item_rolls_back_batch(self):
        self.receipt({10: 10})
        with closing(sqlite3.connect(self.path)) as conn, conn:
            conn.execute("UPDATE library_items SET is_active=false WHERE id=10")
            conn.execute("INSERT INTO library_items VALUES (30,'C','心经','经书',60,true)")
        with self.assertRaisesRegex(ValueError, '已经对应'):
            self.receipt({20: 5, 30: 4})
        self.assertEqual(self.rows('SELECT COUNT(*) FROM hongfa_library_transfers'), [(1,)])
        self.assertEqual(self.rows("SELECT COUNT(*) FROM hongfa_items WHERE title='新书'"), [(0,)])

    def test_audit_failure_reverts_the_entire_confirmation(self):
        self.receipt({10: 10, 20: 5})
        audit = Mock(side_effect=[None, RuntimeError('audit unavailable')])
        with self.assertRaisesRegex(RuntimeError, 'audit unavailable'):
            bridge.confirm_receipts([1, 2], {'id': 88}, 'admin', audit)
        self.assertEqual(self.rows('SELECT status FROM hongfa_library_transfers'), [('pending',), ('pending',)])
        self.assertEqual(self.rows('SELECT COUNT(*) FROM library_transactions'), [(0,)])
        self.assertEqual(self.rows('SELECT balance FROM library_items ORDER BY id'), [(100,), (40,)])

    def test_inactive_team_and_illegal_quantities(self):
        with self.assertRaisesRegex(ValueError, '组别'):
            self.receipt({10: 1}, team=3)
        for value in ('-1', '1.2', '1e2', '100001', '0'):
            with self.subTest(value=value), self.assertRaises(ValueError):
                bridge.quantity_lines(MultiDict([('qty_10', value)]))
        with self.assertRaises(ValueError):
            bridge.quantity_lines(MultiDict([('qty_10', '5'), ('qty_10', '8')]))


class RouteTests(unittest.TestCase):
    def setUp(self):
        self.who = None
        self.admin = None
        self.perms = set()
        self.csrf_ok = False
        self.app = Flask(__name__)
        self.app.secret_key = 'offline-fixture'
        self.app.testing = True
        hf = Blueprint('hongfa', __name__, url_prefix='/hongfa')
        lib = Blueprint('library', __name__, url_prefix='/library')
        bridge.register_hongfa_library(hf, lambda: self.who, render_template_string, lambda: 'csrf',
                                       lambda: self.csrf_ok)
        bridge.register_library_confirmation(lib, lambda: self.admin, self.perms.__contains__,
                                             lambda admin: 'admin', Mock())
        self.app.register_blueprint(hf)
        self.app.register_blueprint(lib)
        self.session = self.app.test_client()

    def test_default_gate_and_role_boundaries(self):
        self.assertEqual(self.session.get('/hongfa/library-books').status_code, 403)
        self.assertEqual(self.session.get('/library/admin/hongfa-transfers').status_code, 403)
        self.who = {'id': 'CHE-1', 'secretary': False, 'team_ids': [1]}
        with patch.dict(os.environ, {bridge.WRITE_FLAG: ''}):
            self.assertEqual(self.session.post('/hongfa/library-books?team_id=1').status_code, 403)
        self.who = {'id': 'CHE-1', 'secretary': True, 'team_ids': [1]}
        with patch.dict(os.environ, {bridge.WRITE_FLAG: bridge.WRITE_VALUE}):
            self.assertEqual(self.session.post('/hongfa/library-books?team_id=1').status_code, 403)
        self.admin = {'id': 88}
        self.perms.add('can_view_records')
        self.assertEqual(self.session.post('/library/admin/hongfa-transfers').status_code, 403)

    def test_cross_team_read_and_post_csrf(self):
        self.who = {'id': 'CHE-1', 'secretary': False, 'team_ids': [1]}
        with patch.object(bridge, 'installed', return_value=True), patch.object(
                bridge, 'db_query', return_value=[{'id': 1, 'name': '组一'}, {'id': 2, 'name': '组二'}]):
            self.assertEqual(self.session.get('/hongfa/library-books?team_id=2').status_code, 403)
        with patch.dict(os.environ, {bridge.WRITE_FLAG: bridge.WRITE_VALUE}):
            self.assertEqual(self.session.post('/hongfa/library-books?team_id=1').status_code, 400)
        self.admin = {'id': 88}
        self.perms.add('can_batch_out')
        with patch.dict(os.environ, {bridge.CONFIRM_FLAG: bridge.WRITE_VALUE}), patch.object(bridge, 'installed', return_value=True):
            self.assertEqual(self.session.post('/library/admin/hongfa-transfers').status_code, 400)

    def test_authorized_batch_routes_and_secretary_history(self):
        manager = {'id': 'CHE-1', 'secretary': False, 'team_ids': [1]}
        self.who = manager
        self.csrf_ok = True
        key = str(uuid4())
        form = {'csrf_token': 'csrf', 'request_key': key, 'qty_10': '20', 'qty_20': '5'}
        with patch.dict(os.environ, {bridge.WRITE_FLAG: bridge.WRITE_VALUE, bridge.CONFIRM_FLAG: bridge.WRITE_VALUE}), patch.object(
                bridge, 'installed', return_value=True), patch.object(
                bridge, 'db_query', return_value=[{'id': 1, 'name': '组一'}]), patch.object(
                bridge, 'register_receipt', return_value=True) as write, patch.object(
                bridge, '_history', return_value=[]) as history:
            self.assertEqual(self.session.post('/hongfa/library-books?team_id=2', data=form).status_code, 403)
            self.assertEqual(self.session.post('/hongfa/library-books?team_id=1', data=form).status_code, 302)
            write.assert_called_once_with(1, 'CHE-1', key, {10: 20, 20: 5})
            self.who = {**manager, 'secretary': True, 'team_ids': []}
            response = self.session.get('/hongfa/library-books')
            self.assertEqual(response.status_code, 200)
            self.assertNotIn('确认领取', response.get_data(as_text=True))
            history.assert_called_with(None, offset=0)

            self.admin = {'id': 88}
            self.perms.add('can_batch_out')
            with self.session.session_transaction() as login:
                login['library_transfer_csrf'] = 'approval'
            with patch.object(bridge, 'confirm_receipts', return_value=2) as confirm:
                response = self.session.post('/library/admin/hongfa-transfers',
                                             data={'csrf_token': 'approval', 'transfer_id': ['2', '1', '2']})
            self.assertEqual(response.status_code, 302)
            self.assertTrue(response.location.endswith('/admin/hongfa-transfers'))
            self.assertEqual(confirm.call_args.args[:3], ([1, 2], {'id': 88}, 'admin'))

    def test_confirmation_tabs_and_switches_are_independent(self):
        self.who = {'id': 'CHE-1', 'secretary': False, 'team_ids': [1]}
        self.csrf_ok = True
        self.admin = {'id': 88}
        self.perms.add('can_batch_out')
        key = str(uuid4())
        approval = {'csrf_token': 'approval', 'transfer_id': '1'}
        receipt = {'csrf_token': 'csrf', 'request_key': key, 'qty_10': '20'}
        with self.session.session_transaction() as login:
            login['library_transfer_csrf'] = 'approval'
        with patch.object(bridge, 'installed', return_value=True), patch.object(
                bridge, 'db_query', return_value=[{'id': 1, 'name': '组一'}]), patch.object(
                bridge, '_history', return_value=[]) as history, patch.object(
                bridge, 'register_receipt', return_value=True) as receive, patch.object(
                bridge, 'confirm_receipts', return_value=1) as confirm:
            with patch.dict(os.environ, {bridge.WRITE_FLAG: bridge.WRITE_VALUE, bridge.CONFIRM_FLAG: ''}):
                self.assertEqual(self.session.post('/hongfa/library-books?team_id=1', data=receipt).status_code, 302)
                self.assertEqual(self.session.post('/library/admin/hongfa-transfers', data=approval).status_code, 403)
                pending = self.session.get('/library/admin/hongfa-transfers').get_data(as_text=True)
                self.assertIn('/library/admin/hongfa-transfer-history', pending)
                self.assertIn('暂停确认', pending)
                history.assert_called_with(pending=True, offset=0)
                historical = self.session.get('/library/admin/hongfa-transfer-history').get_data(as_text=True)
                self.assertIn('/library/admin/hongfa-transfers', historical)
                history.assert_called_with(pending=False, offset=0)
            with patch.dict(os.environ, {bridge.WRITE_FLAG: '', bridge.CONFIRM_FLAG: bridge.WRITE_VALUE}):
                self.assertEqual(self.session.post('/hongfa/library-books?team_id=1', data=receipt).status_code, 403)
                self.assertEqual(self.session.post('/library/admin/hongfa-transfers', data=approval).status_code, 302)
            receive.assert_called_once()
            confirm.assert_called_once()

    def test_phone_catalog_shows_quantities_and_search(self):
        from playwright.sync_api import sync_playwright

        self.who = {'id': 'CHE-1', 'secretary': False, 'team_ids': [1]}
        rows = [{'id': 10, 'name': '心经', 'category': '经书', 'balance': 100,
                 'pending': 20, 'available': 80, 'team_balance': 30}]

        def query(sql, args=(), **kwargs):
            if 'FROM hongfa_teams' in sql:
                return [{'id': 1, 'name': '组一'}]
            self.assertIn('ORDER BY f.sort_order NULLS LAST', sql)
            self.assertIn('%心经%', args)
            return rows

        with patch.dict(os.environ, {bridge.WRITE_FLAG: bridge.WRITE_VALUE}), patch.object(
                bridge, 'installed', return_value=True), patch.object(
                bridge, '_history', return_value=[]), patch.object(bridge, 'db_query', side_effect=query):
            response = self.session.get('/hongfa/library-books?team_id=1&q=心经')
        self.assertEqual(response.status_code, 200)

        with sync_playwright() as playwright:
            browser = playwright.chromium.launch()
            page = browser.new_page(viewport={'width': 390, 'height': 700}, is_mobile=True, has_touch=True)
            page.route('**/*', lambda route: route.abort())
            page.set_content('<meta name="viewport" content="width=device-width,initial-scale=1">'
                             + response.get_data(as_text=True))
            qty = page.get_by_role('spinbutton', name='心经 领取数量')
            self.assertEqual(qty.get_attribute('max'), '80')
            self.assertGreaterEqual(qty.bounding_box()['height'], 44)
            qty.fill('20')
            self.assertEqual(page.get_by_role('button', name='确认领取').count(), 1)
            self.assertEqual(page.locator('td[data-label="待确认"]').inner_text(), '20')
            self.assertEqual(page.locator('td[data-label="本组库存"]').inner_text(), '30')
            self.assertTrue(page.evaluate('document.documentElement.scrollWidth <= window.innerWidth'))
            browser.close()


class LegacyMovementTests(unittest.TestCase):
    def setUp(self):
        app = Flask(__name__)
        bp = Blueprint('hongfa', __name__, url_prefix='/hongfa')
        actor = lambda: {'id': 'CHE-1', 'secretary': False, 'team_ids': [1]}
        books.register_books(bp, actor, render_template_string, lambda: 'csrf', lambda: True)
        with patch.dict(sys.modules, {'db': blocked_db, 'hongfa.hongfa_book_reports': reports,
                                     'attendance_service': attendance_stub}):
            phase3.register_phase3(bp, actor, render_template_string, lambda: 'csrf', lambda: True)
        app.register_blueprint(bp)
        self.client = app.test_client()

    def record(self, module, path, data, rows, items=None):
        connection = MagicMock()
        cursor = connection.__enter__.return_value.cursor.return_value.__enter__.return_value
        cursor.fetchone.side_effect = rows
        cursor.fetchall.return_value = items or []
        with patch.object(module, 'get_conn', return_value=connection):
            response = self.client.post(path, data=data)
        self.assertEqual(response.status_code, 302)
        operations = [call.args[1] for call in cursor.execute.call_args_list
                      if 'INSERT INTO hongfa_stock_movements' in call.args[0]]
        self.assertEqual(len(operations), 1)
        return operations[0][2:5]

    def test_existing_manual_stock_form_uses_canonical_values(self):
        for option, qty, event, value in [('receipt', 3, None, 'received'),
                                          ('distribution', -3, 7, 'distributed'),
                                          ('adjust_in', 3, None, 'adjustment'),
                                          ('adjust_out', -3, None, 'adjustment')]:
            with self.subTest(option=option):
                rows = [{'id': 1}, {'id': 9}]
                if event:
                    rows.append({'id': event})
                rows += [None, {'balance': 10}]
                data = dict(team_id=1, item_id=9, quantity=3, movement_type=option,
                            request_key='legacy-form', event_id=event or 0)
                self.assertEqual(self.record(phase3, '/hongfa/stock', data, rows), (event, qty, value))

    def test_existing_book_form_distribution_and_zero_count(self):
        items = [{'id': 9}]
        common = dict(team_id=1, request_key='offline-request-number-123')
        self.assertEqual(self.record(books, '/hongfa/books',
                                     dict(common, mode='out', lines='[{"item_id":9,"quantity":3}]'),
                                     [{'id': 1}, None, {'id': 7}, {'balance': 10, 'entries': 1}], items),
                         (7, -3, 'distributed'))
        self.assertEqual(self.record(books, '/hongfa/books',
                                     dict(common, mode='stock', lines='[{"item_id":9,"quantity":0}]'),
                                     [{'id': 1}, None, {'balance': 0, 'entries': 0}, {'?column?': 1}], items),
                         (None, 0, 'adjustment'))


class ReportAndMigrationTests(unittest.TestCase):
    def test_month_year_receipts_types_and_excel_settings(self):
        app = Flask(__name__)
        bp = Blueprint('hongfa', __name__, url_prefix='/hongfa')
        reports.register_book_reports(bp, lambda: {'secretary': True}, render_template_string)
        app.register_blueprint(bp)
        types = [('received', 20), ('receipt', 5), ('distributed', -2), ('distribution', -3), ('adjustment', 1)]
        movements = [dict(team_id=1, team_name='一组', item_id=9, title='心经', quantity=qty,
                          movement_type=typ, created_at=datetime(2026, 10, 31, 16, 0, tzinfo=timezone.utc))
                     for typ, qty in types]

        def query(sql, args=(), **kwargs):
            if 'FROM hongfa_teams WHERE' in sql:
                return [{'id': 1, 'name': '一组'}]
            if 'JOIN hongfa_events' in sql:
                self.assertIn("('distributed','distribution')", sql)
                return []
            self.assertEqual(args[1].utcoffset().total_seconds(), 28800)
            return movements

        with patch.object(reports, 'db_query', side_effect=query):
            for month in (11, 0):
                response = app.test_client().get(f'/hongfa/books/report/download?year=2026&month={month}')
                self.assertEqual(response.status_code, 200)
                workbook = load_workbook(BytesIO(response.data))
                summary = workbook['库存月年汇总']
                self.assertEqual(list(summary.values)[1], ('一组', '心经', 0, 25, 5, 1, 21))
                self.assertEqual(summary.freeze_panes, 'B2')
                self.assertTrue(summary.auto_filter.ref)

    def test_migration_repeat_guards_and_no_historical_rewrite(self):
        schema = (ROOT / 'hongfa/library_bridge.sql').read_text(encoding='utf-8')
        for declaration in re.findall(r'^CREATE TABLE .*$', schema, re.M):
            self.assertIn('IF NOT EXISTS', declaration)
        for name in ('hongfa_library_protect_reserved', 'hongfa_library_protect_transaction'):
            self.assertIn(f'DROP TRIGGER IF EXISTS {name}', schema)
            self.assertIn(f'CREATE TRIGGER {name}', schema)
        self.assertIn('UNIQUE (request_key, library_item_id)', schema)
        self.assertIn('NEW.balance < pending', schema)
        self.assertIn('BEFORE UPDATE OR DELETE ON public.library_transactions', schema)
        zero = (ROOT / 'hongfa/stocktake_zero.sql').read_text(encoding='utf-8')
        self.assertIn('IF EXISTS', zero)
        self.assertIn("CHECK (quantity <> 0 OR movement_type = 'adjustment')", zero)
        self.assertNotRegex(zero, r'(?im)^\s*(UPDATE|DELETE FROM)\b')

    @unittest.skipUnless((PGLITE / 'pglite.wasm').is_file(), 'optional local PGlite fixture is absent')
    def test_postgresql_migrations_execute_twice_and_enforce_guards(self):
        from playwright.sync_api import sync_playwright

        # Only serve test engine assets, never the worktree or production config.
        with ThreadingHTTPServer(('127.0.0.1', 0), partial(SimpleHTTPRequestHandler, directory=PGLITE)) as server:
            thread = Thread(target=server.serve_forever, daemon=True)
            thread.start()
            try:
                origin = f'http://127.0.0.1:{server.server_port}'
                with sync_playwright() as playwright:
                    browser = playwright.chromium.launch()
                    page = browser.new_page(service_workers='block')
                    page.route('**/*', lambda route: route.continue_()
                               if route.request.url.startswith(origin + '/') else route.abort())
                    page.goto(origin)
                    base = (ROOT / 'hongfa/schema.sql').read_text(encoding='utf-8')
                    migrations = [(ROOT / 'hongfa' / name).read_text(encoding='utf-8')
                                  for name in ('stocktake_zero.sql', 'library_bridge.sql')]
                    rows = page.evaluate("""async ({base,migrations}) => {
                      const {PGlite} = await import('/index.js');
                      const pg = await PGlite.create();
                      await pg.exec('CREATE ROLE anon; CREATE ROLE authenticated;');
                      await pg.exec(base);
                      await pg.exec(`CREATE TABLE library_items(id integer PRIMARY KEY,item_code text UNIQUE,name text,category text,balance integer);
                        CREATE TABLE library_transactions(id integer PRIMARY KEY,status text);
                        INSERT INTO hongfa_teams(name) VALUES ('team');
                        INSERT INTO hongfa_items(title) VALUES ('book');
                        INSERT INTO library_items VALUES (1,'A','book','test',100);
                        INSERT INTO hongfa_stock_movements(team_id,item_id,quantity,movement_type,created_by,request_key)
                          VALUES(1,1,10,'received','tester','prior');`);
                      for (const ddl of migrations) { await pg.exec(ddl); await pg.exec(ddl); }
                      await pg.exec(`INSERT INTO hongfa_stock_movements(team_id,item_id,quantity,movement_type,created_by,request_key)
                          VALUES(1,1,0,'adjustment','tester','zero');
                        INSERT INTO hongfa_library_book_map VALUES(1,1,'A','book','test',now());
                        INSERT INTO hongfa_library_transfers(request_key,team_id,library_item_id,item_code,book_name,category,quantity,received_by)
                          VALUES('10000000-0000-0000-0000-000000000000',1,1,'A','book','test',20,'tester');`);
                      async function rejects(sql) {
                        try { await pg.exec(sql); } catch(e) { if(e.code==='23514') return; throw e; }
                        throw new Error('expected a check violation');
                      }
                      await rejects("INSERT INTO hongfa_stock_movements(team_id,item_id,quantity,movement_type,created_by) VALUES(1,1,0,'received','tester')");
                      await rejects('UPDATE library_items SET balance=19 WHERE id=1');
                      await pg.exec(`BEGIN; INSERT INTO library_transactions VALUES (1,'active');
                        UPDATE hongfa_library_transfers SET status='confirmed',confirmed_by='tester',
                          confirmed_at=now(),library_transaction_id=1 WHERE id=1;
                        UPDATE library_items SET balance=balance-20 WHERE id=1; COMMIT;`);
                      await rejects("UPDATE library_transactions SET status='cancelled'");
                      await rejects('DELETE FROM library_transactions');
                      const result = await pg.query(`SELECT (SELECT balance FROM library_items) AS balance,
                          (SELECT SUM(quantity)::int FROM hongfa_stock_movements) AS historical_total,
                          (SELECT status FROM hongfa_library_transfers) AS status`);
                      await pg.close();
                      return result.rows;
                    }""", {'base': base, 'migrations': migrations})
                    self.assertEqual(rows, [{'balance': 80, 'historical_total': 10, 'status': 'confirmed'}])
                    browser.close()
            finally:
                server.shutdown()
                thread.join()


if __name__ == '__main__':
    suite = unittest.defaultTestLoader.loadTestsFromModule(sys.modules[__name__])
    with patch.dict(sys.modules, {'db': blocked_db, 'attendance_service': attendance_stub}):
        status = int(not unittest.TextTestRunner(verbosity=2).run(suite).wasSuccessful())
    if '--all' in sys.argv:
        status |= subprocess.run([sys.executable, '-B', 'test_duty_v2_business_rules.py', '--with-mobile'],
                                 cwd=ROOT).returncode
        status |= subprocess.run([sys.executable, '-B', 'whatsapp_agent/test_whatsapp_popup.py'],
                                 cwd=ROOT).returncode
    sys.exit(status)
