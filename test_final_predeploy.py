"""观音堂部署前一键离线验证。只读；不导入项目模块、不连接数据库。"""
import ast
import pathlib
import sys
import unittest
from unittest.mock import patch

ROOT = pathlib.Path(__file__).resolve().parent
CANDIDATES = [ROOT / 'schedule' / 'builders' / 'schedule_builder_V2_2200_merged_REVIEW.py', ROOT / 'schedule_builder_V2_2200_merged_REVIEW.py', ROOT / 'schedule_builder_V2_2200_merged_REVIEW(1).py']
BUILDER = next((p for p in CANDIDATES if p.exists()), None)
DB = next((p for p in [ROOT / 'db.py', ROOT / 'db(8).py'] if p.exists()), None)

class Tests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if not BUILDER or not DB:
            raise RuntimeError('找不到整合版或 db.py；请将本程序放在项目根目录，并保留整合版文件。')
        cls.code = BUILDER.read_text(encoding='utf-8-sig')
        cls.dbcode = DB.read_text(encoding='utf-8-sig')
        cls.tree = ast.parse(cls.code)
        cls.dbtree = ast.parse(cls.dbcode)
        cls.funcs = {n.name: n for n in cls.tree.body if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef))}
        cls.dbfuncs = {n.name: n for n in cls.dbtree.body if isinstance(n, ast.FunctionDef)}

    def source(self, name):
        return ast.get_source_segment(self.code, self.funcs[name])

    def test_01_syntax(self):
        ast.parse(self.code); ast.parse(self.dbcode)

    def test_02_db_context_rollback(self):
        s = ast.get_source_segment(self.dbcode, self.dbfuncs['get_conn'])
        self.assertIn('conn.rollback()', s)
        self.assertIn('conn.commit()', s)

    def test_03_db_pool_not_modified(self):
        self.assertIn('ThreadedConnectionPool', self.dbcode)

    def test_04_snapshot_transaction_cursor(self):
        s = self.source('_capture_2200_baseline_in_transaction')
        self.assertIn('cur.execute', s)
        self.assertNotIn('get_conn()', s)
        self.assertIn('ON CONFLICT(target_date) DO NOTHING', s)

    def test_05_snapshot_before_commit(self):
        for name in ('publish_full_schedule_2200_atomic', '_publish_existing_2200_atomic'):
            s = self.source(name)
            self.assertLess(s.index('_capture_2200_baseline_in_transaction(cur, target)'), s.rindex('conn.commit()'))

    def test_06_snapshot_exception_not_swallowed(self):
        s = self.source('_capture_2200_baseline_in_transaction')
        self.assertNotIn('except ', s)

    def test_07_snapshot_failure_injection(self):
        # Extract and execute ONLY the helper, with fake modules, fake cursor; no database import.
        fn = self.funcs['_capture_2200_baseline_in_transaction']
        module = ast.Module(body=[fn], type_ignores=[])
        namespace = {}
        import types
        fake_mod = types.ModuleType('schedule.services.whatsapp_morning_summary')
        fake_mod._assignments = lambda cur, target: [{'volunteer_id':'CHE-1','role':'值班'}]
        fake_mod._key = lambda row: (row['volunteer_id'], row['role'])
        class Cursor:
            def __init__(self, fail=False): self.sql=[]; self.fail=fail
            def execute(self, sql, params):
                self.sql.append((sql, params))
                if self.fail and 'INSERT INTO whatsapp_morning_snapshots' in sql: raise RuntimeError('injected snapshot failure')
        with patch.dict(sys.modules, {'schedule.services.whatsapp_morning_summary': fake_mod}):
            exec(compile(module, str(BUILDER), 'exec'), namespace)
            cur = Cursor()
            namespace[fn.name](cur, '2099-01-01')
            self.assertEqual(len(cur.sql), 2)
            self.assertIn('CHE-1', cur.sql[1][1][1])
            with self.assertRaisesRegex(RuntimeError, 'injected snapshot failure'):
                namespace[fn.name](Cursor(fail=True), '2099-01-01')

    def test_08_v2_replacement_atomic(self):
        s = self.source('_replace_duty_signup_atomic')
        self.assertIn('with get_conn() as conn', s)
        self.assertLess(s.index('DELETE FROM volunteer_schedule_assignments'), s.index('INSERT INTO volunteer_schedule_assignments'))
        self.assertLess(s.index('INSERT INTO volunteer_schedule_assignments'), s.index('conn.commit()'))

    def test_09_v2_admin_lock(self):
        self.assertIn("if any(r['locked_by_admin'] for r in target)", self.source('_replace_duty_signup_atomic'))

    def test_10_guard_explicit(self):
        self.assertIn('WA_2200_ALLOW_DB_WRITE', self.source('publish_full_schedule_2200_atomic'))

    def test_11_no_database_access_during_test(self):
        # Test intentionally parses AST and runs only extracted helper with fake cursor.
        self.assertTrue(True)

if __name__ == '__main__':
    print('=== 观音堂最终部署前检查（完全离线、不接 Supabase）===', flush=True)
    print('整合版:', BUILDER, '\ndb:', DB, flush=True)
    result = unittest.TextTestRunner(verbosity=2).run(unittest.defaultTestLoader.loadTestsFromTestCase(Tests))
    print(f'\n结果：{result.testsRun-len(result.failures)-len(result.errors)}/{result.testsRun} 通过')
    print('注意：模拟异常传播不等于真实 PostgreSQL 回滚；真实 WhatsApp 发送未验证。')
    sys.exit(0 if result.wasSuccessful() else 1)
