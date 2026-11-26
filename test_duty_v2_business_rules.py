"""V2 business contracts. Runs production source offline with synthetic data."""

import copy
import logging
import os
from pathlib import Path
import sys
from datetime import date
from types import SimpleNamespace
import unittest
from unittest.mock import MagicMock, patch

from schedule.builders.duty_v2 import plan
from schedule.builders.time_utils import min_to_ampm, parse_min
from test_duty_v2 import A, G, person
from test_duty_v2_full_day import ProductionFakeDB, load_functions
from test_duty_v2_integration import ns as merge_helpers


ROOT = Path(__file__).resolve().parent
BUILDER = ROOT / 'schedule/builders/schedule_builder.py'
DAY = '2026-11-26'


def member(code, start, end, **fields):
    return dict(person(code, start, end), **fields)


def existing(code, start, end, place):
    return member(code, start, end, assigned_place=place)


def segments(result):
    return [(parse_min(row['start_time']), parse_min(row['end_time']), key[1:])
            for key, row in result]


def locations_at(minute, rows):
    return [place for start, end, place in rows if start <= minute < end]


class BusinessPlannerTests(unittest.TestCase):
    def assert_staffing(self, rows, start, end, hall, center):
        for minute in range(start, end):
            locations = locations_at(minute, rows)
            self.assertEqual((locations.count(G), locations.count(A)), (hall, center),
                             f'{min_to_ampm(minute)}: {locations}')

    def assert_four_four(self, result):
        rows = segments(result)
        self.assertEqual([(start, end) for start, end, _ in rows],
                         [(600, 840), (840, 1080)])
        self.assertCountEqual([place for _, _, place in rows], [G, A])

    def test_both_locations_remain_covered_during_staggered_handover(self):
        rows = plan([member('CHE-1', '10:00', '14:00'),
                     member('CHE-2', '10:00', '12:00'),
                     member('CHE-3', '12:00', '14:00')])
        self.assert_staffing(segments(rows), 600, 840, 1, 1)

    def assert_lunch_priority(self, patch_mode):
        # Both 4+4 directions leave the same late coverage gap. Putting the
        # newcomer at the hall during lunch leaves the center one person short.
        frozen = [
            existing('CHE-2', '13:00', '14:00', A),
            existing('CHE-3', '10:00', '16:00', G),
            existing('CHE-4', '10:00', '16:00', G),
            existing('CHE-5', '10:00', '15:00', A),
        ]
        rows = segments(plan([member('CHE-1', '10:00', '18:00')],
                             existing=frozen, patch_mode=patch_mode))
        rows += [(parse_min(row['start_time']), parse_min(row['end_time']),
                  row['assigned_place']) for row in frozen]
        self.assert_staffing(rows, 660, 780, 2, 2)

    def test_lunch_center_staffing_has_priority_before_publish(self):
        self.assert_lunch_priority(patch_mode=False)

    def test_lunch_center_staffing_has_priority_during_full_day_patch(self):
        self.assert_lunch_priority(patch_mode=True)

    def test_center_gets_odd_extra_person_before_1400_outside_lunch(self):
        rows = plan([member('CHE-1', '10:00', '14:00'),
                     member('CHE-2', '10:00', '14:00'),
                     member('CHE-3', '13:00', '14:00')])
        self.assert_staffing(segments(rows), 780, 840, 1, 2)

    def test_four_workers_split_two_and_two(self):
        rows = plan([member(f'CHE-{i}', '10:00', '14:00') for i in range(4)])
        self.assert_staffing(segments(rows), 600, 840, 2, 2)

    def test_five_workers_give_the_center_three(self):
        rows = plan([member(f'CHE-{i}', '10:00', '14:00') for i in range(5)])
        self.assert_staffing(segments(rows), 600, 840, 2, 3)

    def test_non_full_day_prefers_different_morning_and_afternoon_locations(self):
        rows = segments(plan([member('CHE-1', '10:00', '17:00')]))
        self.assertCountEqual([place for _, _, place in rows], [G, A])
        self.assertEqual([(s, e) for s, e, _ in rows], [(600, 840), (840, 1020)])

    def test_two_signups_for_one_id_rotate_even_with_different_display_names(self):
        morning = dict(member('CHE-1', '12:00', '14:00', signup_id=1), name='morning')
        afternoon = dict(member('CHE-1', '14:00', '15:00', signup_id=2), name='afternoon')
        rows = plan([morning, afternoon])
        places = {row['signup_id']: key[1:] for key, row in rows}
        self.assertEqual(set(places), {1, 2})
        self.assertCountEqual(places.values(), [G, A])

    def test_short_cross_boundary_signup_can_rotate_at_1400(self):
        rows = segments(plan([member('CHE-1', '12:00', '15:00')]))
        self.assertEqual([(s, e) for s, e, _ in rows], [(720, 840), (840, 900)])
        self.assertCountEqual([place for _, _, place in rows], [G, A])

    def test_duty_starts_at_the_other_location_from_cleaning(self):
        signup = member('CHE-1', '10:00', '18:00')
        for cleaning, expected in ((G, A), (A, G)):
            with self.subTest(cleaning=cleaning):
                self.assertEqual(segments(plan([signup], cleaning={'CHE-1': cleaning}))[0][2],
                                 expected)

    def test_history_breaks_ties_but_does_not_override_coverage(self):
        signup = member('CHE-1', '10:00', '12:00')
        for frequent, infrequent in ((G, A), (A, G)):
            with self.subTest(frequent=frequent):
                history = {'CHE-1': {'counts': {frequent: 9, infrequent: 1}}}
                self.assertEqual(segments(plan([signup], history=history)),
                                 [(600, 720, infrequent)])
                frozen = [existing('CHE-2', '10:00', '12:00', infrequent)]
                self.assertEqual(segments(plan([signup], history=history, existing=frozen)),
                                 [(600, 720, frequent)])

    def test_full_day_has_only_1400_switch_before_publish(self):
        signup = member('CHE-1', '10:00', '18:00')
        frozen = [existing('CHE-2', '10:00', '15:00', A),
                  existing('CHE-3', '15:00', '18:00', G)]
        self.assert_four_four(plan([signup], existing=frozen))

    def test_full_day_patch_prefers_four_four_when_coverage_allows(self):
        signup = member('CHE-1', '10:00', '18:00')
        frozen = [existing('CHE-2', '10:00', '14:00', G),
                  existing('CHE-3', '14:00', '18:00', A)]
        result = plan([signup], existing=frozen, patch_mode=True)
        self.assert_four_four(result)
        self.assertEqual(segments(result), [(600, 840, A), (840, 1080, G)])

    def test_full_day_patch_can_move_switch_to_avoid_an_empty_location(self):
        signup = member('CHE-1', '10:00', '18:00')
        frozen = [existing('CHE-2', '10:00', '15:00', A),
                  existing('CHE-3', '15:00', '18:00', G)]
        rows = segments(plan([signup], existing=frozen, patch_mode=True))
        self.assertEqual(rows, [(600, 840, G), (840, 900, G), (900, 1080, A)])
        self.assert_staffing(rows + [(600, 900, A), (900, 1080, G)], 600, 1080, 1, 1)

    def test_stw_152_is_fixed_at_center_in_both_modes(self):
        for patch_mode in (False, True):
            with self.subTest(patch_mode=patch_mode):
                rows = plan([member('STW-152', '10:00', '18:00')], patch_mode=patch_mode)
                self.assertEqual(segments(rows), [(600, 840, A), (840, 1080, A)])

    def test_special_day_designated_workers_and_1300_cutoff(self):
        history = {'CHE-228': {'counts': {A: 99}}}
        with patch('lunar_rules.get_special_day_info', return_value={'lunar_day': 15}):
            for patch_mode in (False, True):
                for code in ('STW-158', 'CHE-228', 'CHE-116', 'CHE-184', 'CHE-50'):
                    with self.subTest(code=code, patch_mode=patch_mode):
                        rows = plan([member(code, '10:00', '18:00')], target_date=DAY,
                                    patch_mode=patch_mode, history=history)
                        self.assertEqual(segments(rows), [(600, 840, A), (840, 1080, G)])
                for start, expected in (('12:59', A), ('13:00', G)):
                    with self.subTest(start=start, patch_mode=patch_mode):
                        rows = plan([member('CHE-228', start, '14:00')], target_date=DAY,
                                    patch_mode=patch_mode, history=history)
                        self.assertEqual(segments(rows), [(parse_min(start), 840, expected)])


class BusinessEntryPointTests(unittest.TestCase):
    def setUp(self):
        self.db = ProductionFakeDB()
        self.db.signup.update(start_time='10:00am', end_time='6:00pm')
        self.handlers = dict(merge_helpers, RealDictCursor=object, os=os, logging=logging,
                             __name__='offline_business', min_to_ampm=min_to_ampm,
                             get_conn=lambda: self.db, load_place_history_cache=lambda **kw: {})
        load_functions(BUILDER, {
            '_replace_duty_signup_atomic', '_v2_cleaning_rows', '_v2_cleaning_context',
            '_validate_full_duty_before_publish', '_publish_existing_2200_atomic',
            'publish_full_schedule_2200_atomic',
        }, self.handlers)

    def fill(self):
        return self.handlers['_replace_duty_signup_atomic'](101, DAY)

    def test_patch_preserves_published_rows_and_fills_only_the_gap(self):
        before = copy.deepcopy(self.db.rows)
        self.assertEqual(self.fill(), (1, 0))
        self.assertEqual(self.db.rows[:-1], before)
        added = self.db.rows[-1]
        self.assertEqual((added['signup_id'], added['assigned_place'],
                          parse_min(added['start_time']), parse_min(added['end_time'])),
                         (101, A, 840, 1080))

    def test_patch_protects_an_admin_locked_assignment(self):
        self.db.rows[0].update(locked_by_admin=True, assignment_source='admin')
        before = copy.deepcopy(self.db.rows)
        self.assertEqual(self.fill(), (0, 0))
        self.assertEqual(self.db.rows, before)

    def test_repeated_patch_keeps_exactly_the_same_rows(self):
        self.assertEqual(self.fill(), (1, 0))
        filled = copy.deepcopy(self.db.rows)
        self.assertEqual(self.fill(), (0, 0))
        self.assertEqual(self.db.rows, filled)

    def test_manual_and_2200_publish_both_reject_conflicting_automatic_rows(self):
        signup = self.db.signup
        proposed = [dict(row, assigned_place=key[1:]) for key, row in
                    plan([dict(signup, signup_id=101)])]
        old = [dict(row, assigned_place=G, assignment_source='system',
                    locked_by_admin=False, status='assigned') for row in proposed]
        conn = MagicMock()
        conn.__enter__.return_value = conn
        cur = conn.cursor.return_value.__enter__.return_value
        cur.fetchone.side_effect = [{'n': len(old)}, {'is_published': False}]
        cur.fetchall.side_effect = [[signup], old]
        with patch.dict(os.environ, {'WA_2200_ALLOW_DB_WRITE': '1'}), patch.dict(self.handlers, {
            'get_conn': lambda: conn,
            'calculate_schedule_for_date': lambda target: (None, None, proposed, None, None),
        }):
            with self.assertRaisesRegex(ValueError, '强制'):
                self.handlers['publish_full_schedule_2200_atomic'](DAY)
        conn.commit.assert_not_called()

        publish = MagicMock()
        rerun = MagicMock()
        repair = MagicMock(return_value={'failed': []})
        dependencies = SimpleNamespace(repair_unassigned_duty_for_date=repair,
                                       rebuild_published_cleaning_assignments=MagicMock(),
                                       _wa6_minutes=merge_helpers['_wa6_minutes'])
        namespace = dict(date=date, session={'schedule_login': True},
                         schedule_bp=SimpleNamespace(route=lambda *args, **kw: lambda func: func),
                         request=SimpleNamespace(form={'date': DAY}),
                         db_query=MagicMock(side_effect=[{'n': len(old)}, [signup], old]),
                         is_schedule_published=lambda target: False,
                         publish_schedule_for_date=publish, run_schedule_for_date=rerun,
                         redirect=lambda url: ('redirect', 302), url_for=lambda *args, **kw: '/')
        load_functions(ROOT / 'schedule/schedule_web.py', {'publish_schedule'}, namespace)
        with patch.dict(sys.modules, {'schedule.builders.schedule_builder': dependencies}):
            _, status = namespace['publish_schedule']()
        repair.assert_called_once_with(DAY)
        rerun.assert_not_called()
        self.assertEqual((status, publish.call_count), (409, 0),
                         '22:00 rejects the conflict. Manual publish must also reject it.')


class November26ReferenceTests(unittest.TestCase):
    # Continuous windows inferred from the administrator's adjusted roster.
    # The real signup IDs, volunteer IDs and signup boundaries are unknown.
    WINDOWS = {
        'Anna': (600, 1020),
        '程秋萍': (600, 720),
        '陈天赐': (660, 840),
        '陈丽珍': (720, 900),
        '廖秋華': (600, 840),
        '温汉聪': (600, 1080),
        '邹运金': (720, 960),
    }
    CROSS_SHIFT = ('Anna', '陈丽珍', '邹运金', '温汉聪')

    @classmethod
    def setUpClass(cls):
        people = [member(f'CHE-{index}', min_to_ampm(start), min_to_ampm(end), name=name)
                  for index, (name, (start, end)) in enumerate(cls.WINDOWS.items(), 1)]
        cls.rows = plan(people, target_date=DAY)

    def test_adjusted_roster_windows_and_feasible_two_location_coverage(self):
        self.assertEqual({row['name'] for _, row in self.rows}, set(self.WINDOWS))
        for name, (start, end) in self.WINDOWS.items():
            with self.subTest(name=name):
                own = segments([(key, row) for key, row in self.rows if row['name'] == name])
                scheduled = sorted(minute for a, b, _ in own for minute in range(a, b))
                self.assertEqual(scheduled, list(range(start, end)))
        rows = segments(self.rows)
        for minute in range(600, 1020):
            self.assertEqual(set(locations_at(minute, rows)), {G, A}, min_to_ampm(minute))
        for minute in range(1020, 1080):
            self.assertEqual(len(locations_at(minute, rows)), 1, min_to_ampm(minute))

    def test_four_cross_shift_workers_change_location_at_1400(self):
        places = {}
        for name in self.CROSS_SHIFT:
            start, end = self.WINDOWS[name]
            own = segments([(key, row) for key, row in self.rows if row['name'] == name])
            self.assertEqual([(a, b) for a, b, _ in own], [(start, 840), (840, end)], name)
            places[name] = tuple(place for _, _, place in own)
        stayed = {name: morning for name, (morning, afternoon) in places.items()
                  if morning == afternoon}
        self.assertEqual(stayed, {}, f'跨班地点分配为 {places}')


if __name__ == '__main__':
    # Block the application DB module and load only the five offline modules.
    with patch.dict(sys.modules, {'db': None}):
        suite = unittest.defaultTestLoader.loadTestsFromNames([
            'test_duty_v2', 'test_duty_v2_integration', 'test_duty_v2_safety',
            'test_duty_v2_full_day', '__main__',
        ])
        result = unittest.TextTestRunner(verbosity=2).run(suite)
    sys.exit(0 if result.wasSuccessful() else 1)
