"""Offline tests; no database imports or connections."""
import sys, types, unittest
from unittest.mock import patch
# Isolate calendar dependency; explicit meal/ordinary cases below.
calendar = types.ModuleType('lunar_rules')
calendar.is_meal_festival_day = lambda day: False
with patch.dict(sys.modules, {'lunar_rules': calendar}):
    from schedule.builders import duty_v2 as v

def person(i, s='10:00', e='18:00', place=None):
    p = dict(volunteer_id=i, start_time=s, end_time=e)
    if place: p['assigned_place'] = place
    return p

def spans(rows, i='A'):
    return [(v._start(p), v._end(p), k[1:]) for k,p in rows if p['volunteer_id']==i]

class BoundaryTests(unittest.TestCase):
    def test_standard(self):
        for patch_mode in (False, True):
            r=v.plan([person('A'),person('B')],patch_mode=patch_mode)
            for i in ('A','B'):
                a,b=spans(r,i); self.assertEqual((a[0],a[1],b[0],b[1]),(600,840,840,1080)); self.assertNotEqual(a[2],b[2])
    def test_fixed_center(self):
        for patch_mode in (False,True):
            r=v.plan([person('A'),person('STW-152')],patch_mode=patch_mode)
            self.assertEqual(spans(r),[(600,840,'观音堂'),(840,1080,'观音堂')])
    def test_existing_late_handover(self):
        existing=[person('B','10:00','15:00','活动中心'),person('C','15:00','18:00','观音堂')]
        r=v.plan([person('A')],existing=existing,patch_mode=True)
        self.assertEqual(spans(r),[(600,840,'观音堂'),(840,900,'观音堂'),(900,1080,'活动中心')])
    def test_busy_lunch(self):
        existing=[person('B','10:00','14:00','观音堂'),person('C','10:00','14:00','活动中心'),person('D','11:00','14:00','活动中心'),person('E','14:00','18:00','活动中心')]
        r=v.plan([person('A')],existing=existing,patch_mode=True)
        self.assertEqual(spans(r),[(600,840,'活动中心'),(840,1080,'观音堂')])
    def test_preserved(self):
        r=v.plan([person('A')],existing=[person('B',place='活动中心')],patch_mode=True,preserved=[(600,720,'观音堂')])
        self.assertEqual(spans(r),[(600,840,'观音堂'),(840,1080,'观音堂')])
    def test_reported_roster(self):
        r=v.plan([person('A'),person('STW-152','11:00'),person('B','10:00','12:00'),person('C','12:00','14:00'),person('D','12:00','14:00')])
        self.assertEqual(spans(r),[(600,840,'活动中心'),(840,1080,'观音堂')])
        for minute in range(600,1080):
            self.assertEqual({k[1:] for k,p in r if v._start(p)<=minute<v._end(p)},set(v.PLACES))
    def test_meal(self):
        with patch.object(v,'is_meal_festival_day',return_value=True):
            r=v.plan_for_date([person('A'),person('STW-152')],target_date='2026-11-28',patch_mode=True)
        self.assertEqual(spans(r),[(600,840,'观音堂'),(840,1080,'观音堂')])
if __name__=='__main__': unittest.main(verbosity=2)
