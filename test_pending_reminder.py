"""Offline test: run from project root or this extracted folder. No application imports or real DB."""
import ast
import sqlite3
from datetime import date, timedelta
from pathlib import Path
from jinja2 import Environment

root = Path(__file__).resolve().parent
if not (root / 'schedule').exists():
    root = Path.cwd()

def load_functions(path, names, env):
    tree = ast.parse(path.read_text(encoding='utf-8-sig'))
    selected = [n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name in names]
    assert len(selected) == len(names)
    exec(compile(ast.Module(body=selected, type_ignores=[]), str(path), 'exec'), env)

class Cursor:
    def __enter__(self): return self
    def __exit__(self, *args): pass
    def execute(self, sql, params):
        self.result = db.execute(sql.replace('%s', '?'), params)
    def fetchone(self): return self.result.fetchone()
class Connection:
    def cursor(self, **kwargs): return Cursor()

def query(sql, params, fetchone=False, fetchall=False):
    result = db.execute(sql.replace('%s', '?'), params)
    return result.fetchone() if fetchone else result.fetchall()

db = sqlite3.connect(':memory:')
db.row_factory = sqlite3.Row
db.executescript('''CREATE TABLE volunteer_schedule_signups(id INTEGER PRIMARY KEY, signup_date TEXT, name TEXT, volunteer_id TEXT, role TEXT, start_time TEXT, end_time TEXT, status TEXT);
CREATE TABLE schedule_day_flags(flag_date TEXT, need_setup_master_table BOOLEAN, need_remove_master_table BOOLEAN, setup_people TEXT, remove_people TEXT, remarks TEXT, is_published BOOLEAN, need_republish BOOLEAN, last_schedule_change TEXT);''')
for day, n in [('2026-11-21',2),('2026-11-22',2),('2026-11-28',1),('2026-12-01',2),('2026-12-15',7),('2027-01-01',3)]:
    db.executemany('INSERT INTO volunteer_schedule_signups VALUES(NULL,?,?,?,?,?,?,?)', [(day,'Same volunteer','CHE-1','值班','10:00','12:00','pending')]*n)
for role,status in [('膳食','pending'),('膳食组','pending'),('值班','assigned'),('值班','cancelled')]:
    db.execute('INSERT INTO volunteer_schedule_signups VALUES(NULL,?,?,?,?,?,?,?)',('2026-11-20','Excluded','CHE-2',role,None,None,status))
original = list(map(tuple,db.execute('SELECT * FROM volunteer_schedule_signups')))
env = dict(timedelta=timedelta, malaysia_today=lambda: today, db_query=query, RealDictCursor=None)
services = root/'schedule'/'services'
load_functions(services/'assignment_service.py', {'load_schedule_admin_dashboard_data'},env)
load_functions(services/'schedule_calendar_queries.py', {'get_pending_signup_counts','get_pending_signup_details'},env)
source = (root/'schedule'/'schedule_page_templates.py').read_text(encoding='utf-8-sig')
compile(source,'schedule_page_templates.py','exec')
html = next(ast.literal_eval(n.value) for n in ast.parse(source).body if isinstance(n,ast.Assign) and any(isinstance(t,ast.Name) and t.id=='SCHEDULE_HTML' for t in n.targets))
Environment().parse(html)
section = html.split('<!-- 未安排提醒 -->',1)[1].split('<!-- 供台报名提醒 -->',1)[0]
for iso,expected in [('2026-11-20',(0,2,5)),('2026-11-30',(0,2,5)),('2026-12-01',(2,0,9)),('2026-12-31',(0,3,9)),('2027-01-01',(3,0,3)),('2027-02-01',(0,0,0))]:
    today=date.fromisoformat(iso)
    counts,_,_=env['load_schedule_admin_dashboard_data'](iso,conn=Connection())
    assert tuple(counts[k] for k in ('today','tomorrow','month'))==expected,counts
    assert counts==env['get_pending_signup_counts']()
    rows=[dict(r) for r in env['get_pending_signup_details']()]
    assert len(rows)==expected[2]
    for r in rows:
        r['signup_date']=date.fromisoformat(r['signup_date']); assert (r['signup_date'].year,r['signup_date'].month)==(today.year,today.month)
    rendered=Environment().from_string(section).render(pending_counts=counts,pending_details=rows,pending_tomorrow_label='test')
    assert f'本月：<b>{expected[2]}</b> 笔' in rendered
    assert f'查看未安排名单（{expected[2]} 笔）' in rendered
    assert rendered.count('查看当天')==expected[2]
    assert f'今日：<b>{expected[0]}</b> 笔' in rendered and f'明日：<b>{expected[1]}</b> 笔' in rendered
    print(iso,expected,'PASS')
assert original==list(map(tuple,db.execute('SELECT * FROM volunteer_schedule_signups')))
for path in (root/'schedule').rglob('*.py'):
    if path.name in {'assignment_service.py','schedule_calendar_queries.py','schedule_page_templates.py'}: ast.parse(path.read_text(encoding='utf-8-sig'))
print('PASS: counts, details, template, unchanged signup data; SQLite simulation; no Supabase access')
