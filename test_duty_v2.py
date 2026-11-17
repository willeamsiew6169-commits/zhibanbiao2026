"""观音堂 V2 离线模拟：在项目根目录运行 python test_duty_v2.py。
只导入 schedule.builders.duty_v2；不连接数据库、不写入文件、不发送通知。
失败表示待修复，不要部署。退出码非零表示有失败。
"""
import unittest
from collections import defaultdict
from schedule.builders.duty_v2 import plan
from schedule.builders.time_utils import parse_min

G, A = '观音堂', '活动中心'

def person(code, start='10:00', end='14:00'):
    return {'volunteer_id': code, 'name': code, 'start_time': start, 'end_time': end}

def segments(rows):
    out = []
    for key, p in rows:
        place = A if key.endswith(A) else G if key.endswith(G) else None
        out.append((p['volunteer_id'], parse_min(p['start_time']), parse_min(p['end_time']), place))
    return out

def coverage(rows, start, end, place):
    return any(s <= start and e >= end and loc == place for _,s,e,loc in segments(rows))

def assert_exact_signup(test, people, rows):
    for p in people:
        code = p['volunteer_id']
        own = sorted((s,e) for who,s,e,_ in segments(rows) if who == code)
        test.assertTrue(own, f'{code} 未被安排')
        test.assertEqual(own[0][0], parse_min(p['start_time']))
        test.assertEqual(own[-1][1], parse_min(p['end_time']))
        for (_,e),(s,_) in zip(own,own[1:]):
            test.assertEqual(e,s,f'{code} 有重叠或空白')

class DutyV2OfflineTests(unittest.TestCase):
    def test_01_single(self):
        ps=[person('CHE-1')]; assert_exact_signup(self,ps,plan(ps))
    def test_02_two_full_coverage(self):
        r=plan([person('CHE-1'),person('CHE-2')]); self.assertTrue(all(coverage(r,s,e,p) for s,e in [(600,720),(720,840)] for p in (G,A)))
    def test_03_two_swap(self):
        r=segments(plan([person('CHE-1'),person('CHE-2')])); self.assertEqual({loc for who,s,e,loc in r if who=='CHE-1'}, {G,A}); self.assertEqual({loc for who,s,e,loc in r if who=='CHE-2'}, {G,A})
    def test_04_three_same(self):
        ps=[person(f'CHE-{i}') for i in range(1,4)]; r=plan(ps); assert_exact_signup(self,ps,r); self.assertTrue(all(coverage(r,600,840,p) or (coverage(r,600,720,p) and coverage(r,720,840,p)) for p in (G,A)))
    def test_05_four_same(self):
        ps=[person(f'CHE-{i}') for i in range(1,5)]; assert_exact_signup(self,ps,plan(ps))
    def test_06_short(self):
        ps=[person('CHE-1','10:00','11:00')]; assert_exact_signup(self,ps,plan(ps))
    def test_07_half_hour(self):
        ps=[person('CHE-1','10:00','14:00'),person('CHE-2','10:30','14:00')]; assert_exact_signup(self,ps,plan(ps))
    def test_08_partial_overlap(self):
        ps=[person('CHE-1','10:00','14:00'),person('CHE-2','11:00','13:00')]; assert_exact_signup(self,ps,plan(ps))
    def test_09_handover(self):
        ps=[person('CHE-1'),person('CHE-2','10:00','12:00'),person('CHE-3','12:00','14:00')]; r=plan(ps); self.assertTrue(all(coverage(r,s,e,p) for s,e in [(600,720),(720,840)] for p in (G,A)))
    def test_10_cross_1400(self):
        ps=[person('CHE-1','13:00','16:00')]; assert_exact_signup(self,ps,plan(ps))
    def test_11_green(self):
        ps=[person('CHE-1','08:00','10:00')]; self.assertTrue(all(k.startswith('绿') for k,_ in plan(ps)))
    def test_12_orange(self):
        ps=[person('CHE-1','10:00','14:00')]; self.assertTrue(all(k.startswith('橙') for k,_ in plan(ps)))
    def test_13_yellow(self):
        ps=[person('CHE-1','14:00','18:00')]; self.assertTrue(all(k.startswith('黄') for k,_ in plan(ps)))
    def test_14_cross_green_orange(self):
        ps=[person('CHE-1','09:00','11:00')]; self.assertEqual({k[0] for k,_ in plan(ps)}, {'绿','橙'})
    def test_15_cross_orange_yellow(self):
        ps=[person('CHE-1','13:00','15:00')]; self.assertEqual({k[0] for k,_ in plan(ps)}, {'橙','黄'})
    def test_16_long_shift(self):
        ps=[person('CHE-1','10:00','18:00')]; assert_exact_signup(self,ps,plan(ps))
    def test_17_fixed_place(self):
        ps=[person('STW-152')]; self.assertTrue(all(k.endswith(A) for k,_ in plan(ps)))
    def test_18_fixed_and_other(self):
        ps=[person('STW-152'),person('CHE-2')]; r=plan(ps); self.assertTrue(all(k.endswith(A) for k,p in r if p['volunteer_id']=='STW-152'))
    def test_19_history_gyt(self):
        ps=[person('CHE-1')]; h={'CHE-1':{'counts':{G:9,A:1}}}; self.assertTrue(all(k.endswith(A) for k,_ in plan(ps,history=h)))
    def test_20_history_activity(self):
        ps=[person('CHE-1')]; h={'CHE-1':{'counts':{G:1,A:9}}}; self.assertTrue(all(k.endswith(G) for k,_ in plan(ps,history=h)))
    def test_21_history_tie(self):
        ps=[person('CHE-1')]; assert_exact_signup(self,ps,plan(ps,history={'CHE-1':{'counts':{G:5,A:5}}}))
    def test_22_new_no_history(self):
        ps=[person('CHE-999')]; assert_exact_signup(self,ps,plan(ps,history={}))
    def test_23_late_fill_other(self):
        ps=[person('CHE-2')]; existing=[dict(person('CHE-1'),assigned_place=G)]; r=plan(ps,existing=existing); self.assertTrue(all(k.endswith(A) for k,_ in r))
    def test_24_late_fill_gap(self):
        ps=[person('CHE-3')]; existing=[dict(person('CHE-1'),assigned_place=G),dict(person('CHE-2','10:00','12:00'),assigned_place=A)]; r=plan(ps,existing=existing); self.assertTrue(coverage(r,720,840,A))
    def test_25_late_full_signup(self):
        ps=[person('CHE-3')]; existing=[dict(person('CHE-1'),assigned_place=G)]; assert_exact_signup(self,ps,plan(ps,existing=existing))
    def test_26_existing_not_returned(self):
        ps=[person('CHE-3')]; existing=[dict(person('CHE-1'),assigned_place=G)]; self.assertEqual({p['volunteer_id'] for _,p in plan(ps,existing=existing)}, {'CHE-3'})
    def test_27_late_sequential(self):
        base=[dict(person('CHE-1'),assigned_place=G)]; first=plan([person('CHE-2')],existing=base); frozen=base+[dict(p,assigned_place=A if k.endswith(A) else G) for k,p in first]; second=plan([person('CHE-3')],existing=frozen); self.assertEqual({p['volunteer_id'] for _,p in second},{'CHE-3'})
    def test_28_invalid_time(self):
        self.assertEqual(plan([person('CHE-1','14:00','10:00')]),[])
    def test_29_empty(self):
        self.assertEqual(plan([]),[])
    def test_30_deterministic(self):
        ps=[person('CHE-1'),person('CHE-2'),person('CHE-3')]; self.assertEqual(plan(ps),plan(list(reversed(ps))))

if __name__ == '__main__':
    print('观音堂智能排班 V2 — 离线模拟（不连接数据库）', flush=True)
    unittest.main(verbosity=2)
