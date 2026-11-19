# Offline mock checks; no database connection or writes.
import sys, types, unittest
from datetime import date, datetime
from decimal import Decimal
from unittest.mock import patch
sys.modules['db'] = types.ModuleType('db')
sys.modules['db'].db_query = lambda *a, **k: None
import reports_service as r
D=date(2026,11,15)
class Test(unittest.TestCase):
    def test_live(self):
        a=dict(id=1, volunteer_id='806', volunteer_name='A', day=D, role='值班', start_time='9:05am',end_time='2:12pm',hours=Decimal('4.92'),remarks='assignment_ids=537,536',data_source='cloud')
        def log(vid,role,aid,place,remark=''):
            return dict(id=1, volunteer_id=vid,name=vid,day=D,actual_role=role,assignment_id=aid,actual_place=place,remarks=remark,check_in_time=datetime(2026,11,15,16),check_out_time=datetime(2026,11,15,18))
        farm=log('CHE-232','佛学班',None,'农舍','历史补录：农舍佛学班，实际出席2小时')
        with patch.object(r,'db_query',side_effect=[[a],[log('806','值班',537,'活动中心'),log('806','值班',536,'观音堂'),farm,farm,log('X','值班',None,'观音堂')]]):
            rows=r.load_live_attendance(D,date(2026,11,16))
        self.assertEqual(len(rows),2)
        self.assertEqual(sum(x['hours'] for x in rows),Decimal('6.92'))
        self.assertEqual(rows[0]['place'],'活动中心、观音堂')
        self.assertEqual(rows[1]['start_time'],'16:00')
    def test_cutoff(self):
        with patch.object(r,'db_query',return_value=[]),patch.object(r,'load_live_attendance',return_value=[]) as live:
            r.load_attendance_rows(2026,10); live.assert_not_called()
            r.load_attendance_rows(2026,11); live.assert_called_once_with(D,date(2026,12,1))
if __name__=='__main__': unittest.main()
