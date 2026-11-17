"""Offline checks only. No Supabase, Render, or WhatsApp access."""
import ast
import difflib
import json
from pathlib import Path

HERE = Path(__file__).resolve().parent
ORIGINAL = HERE / "schedule_builder.py"
PATCHED = HERE / 'schedule_builder_2200_snapshot_safe.py'


def check(label, condition):
    print(('PASS' if condition else 'FAIL') + ' | ' + label)
    if not condition:
        raise AssertionError(label)


original = ORIGINAL.read_text(encoding='utf-8')
patched = PATCHED.read_text(encoding='utf-8')
ast.parse(patched)
check('Python syntax', True)

diff = list(difflib.unified_diff(original.splitlines(), patched.splitlines()))
added = [line[1:] for line in diff if line.startswith('+') and not line.startswith('+++') and line[1:].strip()]
removed = [line[1:] for line in diff if line.startswith('-') and not line.startswith('---') and line[1:].strip()]
check('No existing nonblank lines removed', not removed)
check('Only helper and two calls added', len([line for line in added if '_capture_2200_baseline_in_transaction(cur, target)' in line]) == 3)

tree = ast.parse(patched)
funcs = {node.name: node for node in tree.body if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))}
for name in ('publish_full_schedule_2200_atomic', '_publish_existing_2200_atomic'):
    fn = funcs[name]
    code = ast.get_source_segment(patched, fn)
    check(name + ' invokes snapshot before commit', code.index('_capture_2200_baseline_in_transaction(cur, target)') < code.rindex('conn.commit()'))

helper = funcs['_capture_2200_baseline_in_transaction']
code = ast.get_source_segment(patched, helper)
check('Snapshot uses non-overwriting insert', 'ON CONFLICT(target_date) DO NOTHING' in code)
check('Snapshot uses existing transaction cursor', 'cur.execute' in code and 'get_conn(' not in code and 'commit(' not in code)

# Execute only helper in isolation, with in-memory DB cursor and stub imports.
import sys
import types
schedule = types.ModuleType('schedule'); schedule.__path__ = []
services = types.ModuleType('schedule.services'); services.__path__ = []
summary = types.ModuleType('schedule.services.whatsapp_morning_summary')
summary._assignments = lambda cur, target: [{'volunteer_id':'CHE-1','role':'值班'}, {'volunteer_id':'CHE-2','role':'卫生'}]
summary._key = lambda row: (row['volunteer_id'], row['role'])
sys.modules.update({'schedule':schedule, 'schedule.services':services, 'schedule.services.whatsapp_morning_summary':summary})
namespace = {}
exec(compile(ast.Module(body=[helper], type_ignores=[]), '<snapshot-helper>', 'exec'), namespace)

class FakeCursor:
    def __init__(self, fail=False): self.queries=[]; self.fail=fail
    def execute(self, sql, params):
        self.queries.append((sql,params))
        if self.fail and 'INSERT INTO whatsapp_morning_snapshots' in sql:
            raise RuntimeError('simulated snapshot database failure')

cur = FakeCursor()
namespace['_capture_2200_baseline_in_transaction'](cur, '2026-11-18')
check('Snapshot insert contains both role keys', json.loads(cur.queries[-1][1][1]) == [['CHE-1','值班'], ['CHE-2','卫生']])
try:
    namespace['_capture_2200_baseline_in_transaction'](FakeCursor(fail=True), '2026-11-18')
except RuntimeError as exc:
    check('Snapshot failure propagates to caller', str(exc) == 'simulated snapshot database failure')
else:
    check('Snapshot failure propagates to caller', False)
print('\nOffline checks passed. Database rollback and end-to-end delivery are NOT verified.')
