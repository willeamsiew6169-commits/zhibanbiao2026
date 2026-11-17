"""第二轮只读离线整合测试。放在项目根目录执行 python test_duty_v2_integration.py。
只读取 schedule_builder.py 源码并抽取纯函数，不导入主程序，不连接数据库。"""
import ast
import datetime
import pathlib
import re
import unittest
from schedule.builders.duty_v2 import plan

SOURCE = pathlib.Path(__file__).resolve().parent / 'schedule' / 'builders' / 'schedule_builder.py'
NEEDED = {'_wa6_minutes', '_wa6_gaps', '_wa6_time', 'plan_2200_safe_duty_merge'}
tree = ast.parse(SOURCE.read_text(encoding='utf-8-sig'), filename=str(SOURCE))
functions = [node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name in NEEDED]
if {f.name for f in functions} != NEEDED:
    raise RuntimeError('主程序缺少 22:00 安全合并函数；不能继续')
module = ast.fix_missing_locations(ast.Module(body=functions, type_ignores=[]))
ns = {'re': re, 'datetime': datetime.datetime}
exec(compile(module, str(SOURCE), 'exec'), ns)
merge = ns['plan_2200_safe_duty_merge']

def signup(sid, start='10:00', end='14:00'):
    return dict(id=sid, volunteer_id=f'CHE-{sid}', name=f'义工{sid}', role='值班', start_time=start, end_time=end, status='pending')

def row(sid, start='10:00', end='14:00', place='观音堂'):
    return dict(signup_id=sid, volunteer_id=f'CHE-{sid}', name=f'义工{sid}', role='值班', start_time=start, end_time=end, assigned_place=place, shift_label='橙', status='assigned')

class V2IntegrationOffline(unittest.TestCase):
    def test_01_2200_empty_existing_adds(self):
        self.assertEqual(len(merge([signup(1)], [], [row(1)])), 1)
    def test_02_2200_preserve_existing(self):
        self.assertEqual(merge([signup(1)], [row(1)], [row(1, place='活动中心')]), [])
    def test_03_2200_partial_gap(self):
        a = merge([signup(1)], [row(1, '10:00','12:00')], [row(1,'10:00','14:00')]); self.assertEqual([(r['start_time'],r['end_time']) for r in a],[('12:00:00','14:00:00')])
    def test_04_2200_no_theoretical_raises(self):
        with self.assertRaises(RuntimeError): merge([signup(1)], [], [])
    def test_05_2200_incomplete_theoretical_raises(self):
        with self.assertRaises(RuntimeError): merge([signup(1)], [], [row(1,'10:00','12:00')])
    def test_06_2200_multiple_signups(self):
        self.assertEqual(len(merge([signup(1),signup(2)], [row(1)], [row(1),row(2)])),1)
    def test_07_2200_cancelled_skipped(self):
        s=signup(1);s['status']='cancelled';self.assertEqual(merge([s],[],[]),[])
    def test_08_2200_unrelated_roles_preserved(self):
        other=dict(row(99),role='卫生');self.assertEqual(len(merge([signup(1)],[other],[row(1)])),1)
    def test_09_2200_existing_unchanged_in_memory(self):
        existing=[row(1)];before=[dict(x) for x in existing];merge([signup(1),signup(2)],existing,[row(1),row(2)]);self.assertEqual(existing,before)
    def test_10_2200_clips_to_gap(self):
        a=merge([signup(1,'10:00','12:00')],[],[row(1,'09:00','13:00')]);self.assertEqual((a[0]['start_time'],a[0]['end_time']),('10:00:00','12:00:00'))
    def test_11_v2_late_keeps_existing(self):
        existing=[row(1)];before=[dict(x) for x in existing];plan([signup(2)],existing=existing);self.assertEqual(existing,before)
    def test_12_v2_late_fills_opposite(self):
        r=plan([signup(2)],existing=[row(1)]);self.assertTrue(all(key.endswith('活动中心') for key,_ in r))
    def test_13_v2_full_signup(self):
        r=plan([signup(2)],existing=[row(1)]);self.assertEqual(min(ns['_wa6_minutes'](p['start_time']) for _,p in r),600);self.assertEqual(max(ns['_wa6_minutes'](p['end_time']) for _,p in r),840)
    def test_14_2200_missing_place_fails(self):
        bad=row(1);bad['assigned_place']=None
        with self.assertRaises(RuntimeError):merge([signup(1)],[],[bad])
    def test_15_2200_invalid_interval_fails(self):
        with self.assertRaises(ValueError):merge([signup(1,'14:00','10:00')],[],[row(1)])

if __name__=='__main__':
    print('第二轮 V2 / 22:00 安全合并离线测试（不连接数据库）')
    unittest.main(verbosity=2)
