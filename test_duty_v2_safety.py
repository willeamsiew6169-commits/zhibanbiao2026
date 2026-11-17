"""第三轮离线事务安全测试。不会连接 Supabase / Render，不会修改真实数据。
放在项目根目录，与 schedule_builder_V2_safe.py 同级运行。
"""
import ast
import copy
import pathlib
import sys
import types
import unittest
from unittest.mock import patch

HERE = pathlib.Path(__file__).resolve().parent
SOURCE = HERE / 'schedule_builder_V2_safe.py'

class FakeDB:
    def __init__(self, locked=False):
        self.signup = dict(id=101, volunteer_id='CHE-101', name='测试义工',
                           start_time='10:00am', end_time='2:00pm', role='值班', status='pending')
        self.rows = [dict(id=1, signup_id=101, volunteer_id='CHE-101', assigned_place='观音堂',
                          start_time='10:00am', end_time='2:00pm', locked_by_admin=locked),
                     dict(id=2, signup_id=202, volunteer_id='CHE-202', assigned_place='活动中心',
                          start_time='10:00am', end_time='2:00pm', locked_by_admin=False)]
        self.committed = False
        self.insert_failure = False
        self.statements = []
        self.next_id = 10
        self.pending = None
        self.depth = 0
    def __enter__(self):
        if self.depth == 0:
            self.pending = copy.deepcopy(self.rows)
            self.committed = False
        self.depth += 1
        return self
    def __exit__(self, exc_type, exc, tb):
        self.depth -= 1
        if self.depth == 0 and (exc_type is not None or not self.committed):
            self.pending = None
        return False
    def cursor(self, **kwargs): return self
    def execute(self, sql, args=()):
        norm = ' '.join(sql.lower().split())
        self.statements.append(norm)
        if 'pg_advisory_xact_lock' in norm:
            self.result = []
        elif 'from volunteer_schedule_signups' in norm and 'for update' in norm:
            self.result = [copy.deepcopy(self.signup)]
        elif 'from volunteer_schedule_assignments' in norm and 'for update' in norm:
            self.result = copy.deepcopy(self.pending)
        elif norm.startswith('delete from volunteer_schedule_assignments'):
            target = [r for r in self.pending if r['signup_id'] == args[1] and not r['locked_by_admin']]
            self.pending = [r for r in self.pending if r not in target]
            self.result = [{'id': r['id']} for r in target]
        elif norm.startswith('insert into volunteer_schedule_assignments'):
            if self.insert_failure:
                raise RuntimeError('模拟 INSERT 数据库失败')
            self.pending.append(dict(id=self.next_id, signup_id=args[0], volunteer_id=args[1],
                                     assigned_place=args[5], start_time=args[6], end_time=args[7],
                                     locked_by_admin=False))
            self.next_id += 1
            self.result = []
        elif norm.startswith('update volunteer_schedule_signups'):
            self.result = []
        else:
            raise AssertionError('未识别 SQL：' + norm[:150])
    def fetchone(self): return self.result[0] if self.result else None
    def fetchall(self): return self.result
    def commit(self):
        self.rows = copy.deepcopy(self.pending)
        self.committed = True

class SafetyTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        source = SOURCE.read_text(encoding='utf-8-sig')
        tree = ast.parse(source, filename=str(SOURCE))
        functions = [n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == '_replace_duty_signup_atomic']
        if len(functions) != 1:
            raise AssertionError('修正版缺少唯一的 _replace_duty_signup_atomic')
        namespace = {'RealDictCursor': object, 'load_place_history_cache': lambda **kw: {}, 'get_conn': None}
        cls.namespace = namespace
        exec(compile(ast.Module(body=functions, type_ignores=[]), str(SOURCE), 'exec'), namespace)
    def setUp(self):
        self.db = FakeDB()
        self.namespace['get_conn'] = lambda: self.db
        self.planner = types.ModuleType('schedule.builders.duty_v2')
        self.planner.plan = lambda people, **kw: [('橙观音堂', dict(start_time='10:00am', end_time='2:00pm'))]
        self.modules = patch.dict(sys.modules, {'schedule': types.ModuleType('schedule'),
                         'schedule.builders': types.ModuleType('schedule.builders'),
                         'schedule.builders.duty_v2': self.planner})
        self.modules.start()
        self.addCleanup(self.modules.stop)
    def run_replace(self):
        return self.namespace['_replace_duty_signup_atomic'](101, '2026-11-20')
    def test_01_successful_atomic_replace(self):
        self.assertEqual(self.run_replace(), (1, 1))
        self.assertTrue(self.db.committed)
        self.assertEqual(len([r for r in self.db.rows if r['signup_id'] == 101]), 1)
    def test_02_insert_failure_rolls_back(self):
        original = copy.deepcopy(self.db.rows)
        self.db.insert_failure = True
        with self.assertRaisesRegex(RuntimeError, 'INSERT'):
            self.run_replace()
        self.assertEqual(self.db.rows, original)
    def test_03_planner_failure_preserves_original(self):
        original = copy.deepcopy(self.db.rows)
        self.planner.plan = lambda *a, **kw: (_ for _ in ()).throw(RuntimeError('模拟计算失败'))
        with self.assertRaisesRegex(RuntimeError, '计算失败'):
            self.run_replace()
        self.assertEqual(self.db.rows, original)
    def test_04_admin_lock_prevents_changes(self):
        self.db.rows[0]['locked_by_admin'] = True
        original = copy.deepcopy(self.db.rows)
        self.assertEqual(self.run_replace(), (0, 0))
        self.assertEqual(self.db.rows, original)
    def test_05_other_volunteers_unchanged(self):
        other = copy.deepcopy(self.db.rows[1])
        self.run_replace()
        self.assertEqual([r for r in self.db.rows if r['signup_id'] == 202], [other])
    def test_06_repeat_does_not_duplicate(self):
        self.run_replace()
        self.db.committed = False
        self.run_replace()
        self.assertEqual(len([r for r in self.db.rows if r['signup_id'] == 101]), 1)
    def test_07_empty_plan_preserves_original(self):
        original = copy.deepcopy(self.db.rows)
        self.planner.plan = lambda *a, **kw: []
        with self.assertRaisesRegex(ValueError, '未能生成'):
            self.run_replace()
        self.assertEqual(self.db.rows, original)

if __name__ == '__main__':
    print('第三轮 V2 事务安全离线测试（模拟数据库；不连接云端）', flush=True)
    unittest.main(verbosity=2)
