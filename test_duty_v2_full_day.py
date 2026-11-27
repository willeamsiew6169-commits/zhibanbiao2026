"""Offline full-day coverage and production transaction regressions."""
import ast
import copy
from datetime import datetime, timedelta
import logging
import os
from pathlib import Path
import sys
from types import SimpleNamespace
import unittest
from unittest.mock import MagicMock, patch

from schedule.builders.duty_v2 import plan
from schedule.builders.time_utils import min_to_ampm
from test_duty_v2 import A, G, person, segments, assert_exact_signup
from test_duty_v2_integration import ns
from test_duty_v2_safety import FakeDB


def full(code='CHE-101'):
    return person(code, '10:00am', '6:00pm')


def existing(code, start, end, place):
    return dict(person(code, start, end), assigned_place=place)


def five_three_coverage():
    return [existing('CHE-201', '10:00am', '3:00pm', A),
            existing('CHE-202', '3:00pm', '6:00pm', G)]


def manual_five_three(signup):
    return [
        dict(signup, signup_id=signup['id'], start_time=start, end_time=end,
             assigned_place=place, status='assigned', assignment_source='admin', locked_by_admin=True)
        for start, end, place in [('10:00am', '3:00pm', G), ('3:00pm', '6:00pm', A)]
    ]


def load_functions(path, names, namespace):
    tree = ast.parse(path.read_text(encoding='utf-8-sig'))
    nodes = [node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name in names]
    assert {node.name for node in nodes} == names
    exec(compile(ast.Module(body=nodes, type_ignores=[]), str(path), 'exec'), namespace)


class FullDayPlannerTests(unittest.TestCase):
    def assert_four_four(self, rows, code='CHE-101'):
        own = [(a, b, place) for who, a, b, place in segments(rows) if who == code]
        self.assertEqual([(a, b) for a, b, _ in own], [(600, 840), (840, 1080)])
        self.assertEqual({place for _, _, place in own}, {G, A})
        self.assertEqual([key[0] for key, member in rows if member['volunteer_id'] == code], ['橙', '黄'])

    def test_before_publish_one_to_many(self):
        for count in (1, 2, 5):
            with self.subTest(count=count):
                people = [full(f'CHE-{index}') for index in range(count)]
                rows = plan(people)
                for member in people:
                    self.assert_four_four(rows, member['volunteer_id'])

    def test_mixed_full_and_half_day_keeps_coverage(self):
        people = [full(), person('CHE-2'), person('CHE-3', '14:00', '18:00')]
        rows = plan(people)
        self.assert_four_four(rows)
        assert_exact_signup(self, people, rows)

    def test_five_three_example_before_publish_still_four_four(self):
        self.assert_four_four(plan([full()], existing=five_three_coverage()))

    def test_five_three_example_late_fills_empty_hall(self):
        base = five_three_coverage()
        before = copy.deepcopy(base)
        self.assertEqual(segments(plan([full()], existing=base, patch_mode=True)),
                         [('CHE-101', 600, 840, G), ('CHE-101', 840, 900, G),
                          ('CHE-101', 900, 1080, A)])
        self.assertEqual(base, before)

    def test_late_four_four_wins_when_both_halves_can_be_covered(self):
        base = [existing('CHE-201', '10:00am', '2:00pm', G),
                existing('CHE-202', '2:00pm', '6:00pm', A)]
        rows = plan([full()], existing=base, patch_mode=True)
        self.assert_four_four(rows)
        self.assertEqual([place for _, _, _, place in segments(rows)], [A, G])

    def test_late_four_four_wins_over_history_with_no_other_workers(self):
        self.assert_four_four(plan([full()], patch_mode=True,
                                   history={'CHE-101': {'counts': {G: 100}}}))

    def test_lunch_center_priority_selects_four_four_direction(self):
        base = [existing('CHE-201', '10:00am', '6:00pm', G),
                existing('CHE-202', '10:00am', '6:00pm', A)]
        for late in (False, True):
            rows = plan([full()], existing=base, patch_mode=late)
            self.assert_four_four(rows)
            self.assertEqual(segments(rows)[0][-1], A)

    def test_ordinary_day_locations_and_stw_152_fixed_center(self):
        with patch('lunar_rules.solar_to_lunar', return_value=SimpleNamespace(day=2, month=10)):
            self.assert_four_four(plan([full('CHE-228')], target_date='2026-11-20'), 'CHE-228')
            for late in (False, True):
                rows = plan([full('STW-152')], target_date='2026-11-20', patch_mode=late)
                self.assertEqual(segments(rows), [('STW-152', 600, 840, A), ('STW-152', 840, 1080, A)])

    def test_all_special_day_full_duty_ids_prefer_center_then_hall(self):
        with patch('lunar_rules.solar_to_lunar', return_value=SimpleNamespace(day=15, month=10)):
            for code in ('STW-158', 'CHE-228', 'CHE-116', 'CHE-184', 'CHE-50'):
                for late in (False, True):
                    with self.subTest(code=code, patch_mode=late):
                        rows = plan([full(code)], target_date='2026-11-20', patch_mode=late,
                                    history={code: {'counts': {A: 9}}}, cleaning={code: A})
                        self.assertEqual(segments(rows), [(code, 600, 840, A), (code, 840, 1080, G)])

    def test_special_day_non_full_duty_keeps_1300_exception(self):
        with patch('lunar_rules.solar_to_lunar', return_value=SimpleNamespace(day=15, month=10)):
            for late in (False, True):
                history = {'CHE-228': {'counts': {A: 9}}}
                rows = plan([person('CHE-228')], target_date='2026-11-20', patch_mode=late, history=history)
                self.assertTrue(all(place == A for _, _, _, place in segments(rows)))
                rows = plan([person('CHE-228', '13:00', '15:00')], target_date='2026-11-20',
                            patch_mode=late, history=history)
                self.assertTrue(all(place == G for _, _, _, place in segments(rows)))

    def test_special_day_full_duty_preserves_coverage_with_two_designated_workers(self):
        with patch('lunar_rules.solar_to_lunar', return_value=SimpleNamespace(day=15, month=10)):
            rows = plan([full('CHE-228'), full('CHE-116')], target_date='2026-11-20')
        for code in ('CHE-228', 'CHE-116'):
            self.assert_four_four(rows, code)
        self.assertEqual({place for _, start, _, place in segments(rows) if start == 600}, {A, G})

    def test_cleaning_coordination_breaks_tie_between_four_four_directions(self):
        self.assertEqual(segments(plan([full()], cleaning={'CHE-101': G}))[0][-1], A)

    def test_non_full_day_plan_is_unchanged_between_modes(self):
        members = [person('CHE-1', '09:00', '15:00'), person('CHE-2', '11:00', '17:00')]
        self.assertEqual(plan(members), plan(members, patch_mode=True))
        assert_exact_signup(self, members, plan(members))


class ProductionFakeDB(FakeDB):
    def execute(self, sql, args=()):
        if 'select volunteer_id, assigned_place' in ' '.join(sql.lower().split()) and "role='卫生'" in sql:
            self.result = []
        else:
            super().execute(sql, args)


class FullDayTransactionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        path = Path(__file__).resolve().parent / 'schedule/builders/schedule_builder.py'
        names = {'_replace_duty_signup_atomic', '_validate_full_duty_before_publish', '_plan_meal_duty_gaps',
                 '_publish_existing_2200_atomic', 'publish_full_schedule_2200_atomic'}
        cls.handlers = dict(ns, RealDictCursor=object, min_to_ampm=min_to_ampm,
                            __name__='offline_2200', logging=logging, os=os,
                            load_place_history_cache=lambda **kwargs: {}, _v2_cleaning_rows=lambda rows: {})
        load_functions(path, names, cls.handlers)

    def setUp(self):
        self.db = ProductionFakeDB()
        self.db.signup.update(full())
        self.handlers['get_conn'] = lambda: self.db
        self.ordinary_day = patch('schedule.builders.duty_v2.is_meal_center_day', return_value=False)
        self.ordinary_day.start()
        self.addCleanup(self.ordinary_day.stop)

    def fill(self):
        return self.handlers['_replace_duty_signup_atomic'](101, '2026-11-20')

    def test_partial_fill_preserves_rows_and_repeated_fill_is_noop(self):
        old = copy.deepcopy(self.db.rows)
        self.assertEqual(self.fill(), (1, 0))
        self.assertEqual(self.db.rows[:len(old)], old)
        self.assertEqual(self.db.rows[-1]['assigned_place'], A)
        self.assertEqual(self.db.rows[-1]['start_time'], '2:00pm')
        filled = copy.deepcopy(self.db.rows)
        self.assertEqual(self.fill(), (0, 0))
        self.assertEqual(self.db.rows, filled)
        self.assertFalse(any(sql.startswith(('delete', 'update volunteer_schedule_assignments'))
                             for sql in self.db.statements))

    def test_failed_insert_rolls_back_and_locked_signup_stays_unchanged(self):
        self.db.insert_failure = True
        old = copy.deepcopy(self.db.rows)
        with self.assertRaisesRegex(RuntimeError, 'INSERT'):
            self.fill()
        self.assertEqual(self.db.rows, old)
        self.db.rows[0]['locked_by_admin'] = True
        locked = copy.deepcopy(self.db.rows)
        self.assertEqual(self.fill(), (0, 0))
        self.assertEqual(self.db.rows, locked)

    def test_prepublication_validator_detects_incompatible_existing_places(self):
        validate = self.handlers['_validate_full_duty_before_publish']
        signup = self.db.signup
        good = [dict(row, assigned_place=key[1:]) for key, row in plan([dict(signup, signup_id=101)])]
        validate([signup], good)
        bad = [dict(row, assigned_place=G) for row in good]
        with self.assertRaisesRegex(ValueError, '报名 101.*强制'):
            validate([signup], bad)

    def test_only_existing_admin_and_locked_rows_can_be_exceptions(self):
        validate = self.handlers['_validate_full_duty_before_publish']
        signup = self.db.signup
        rows = manual_five_three(signup)
        self.assertEqual(validate([signup], rows, existing=rows), [101])
        for source, locked in ((None, True), ('system', True), ('admin', False), ('admin', None)):
            with self.subTest(source=source, locked=locked):
                unverified = [dict(row, assignment_source=source, locked_by_admin=locked) for row in rows]
                with self.assertRaisesRegex(ValueError, '报名 101.*强制'):
                    validate([signup], unverified, existing=unverified)
        with self.assertRaisesRegex(ValueError, '报名 101.*强制'):
            validate([signup], rows)

    def test_manual_exception_cannot_hide_invalid_time_place_gap_or_overlap(self):
        validate = self.handlers['_validate_full_duty_before_publish']
        signup = self.db.signup
        rows = manual_five_three(signup)
        for changes in ({'start_time': '2:30pm'}, {'start_time': '3:30pm'},
                        {'end_time': '7:00pm'}, {'assigned_place': ''}):
            with self.subTest(changes=changes):
                invalid = [rows[0], dict(rows[1], **changes)]
                with self.assertRaisesRegex(ValueError, '报名 101.*无效'):
                    validate([signup], invalid, existing=invalid)

    def test_mixed_exception_only_accepts_irregularity_caused_by_manual_rows(self):
        validate = self.handlers['_validate_full_duty_before_publish']
        signup = self.db.signup
        manual, last = manual_five_three(signup)
        automatic = dict(last, assignment_source='system', locked_by_admin=False)
        self.assertEqual(validate([signup], [manual, automatic], existing=[manual]), [101])
        valid_manual = dict(manual, end_time='2:00pm')
        wrong_late = dict(automatic, start_time='2:00pm', assigned_place=G)
        with self.assertRaisesRegex(ValueError, '报名 101.*强制'):
            validate([signup], [valid_manual, wrong_late], existing=[valid_manual, wrong_late])

    def test_2200_manual_exception_publishes_then_queues_whatsapp(self):
        date = '2026-11-20'
        signup = self.db.signup
        old = manual_five_three(signup)
        before = copy.deepcopy(old)
        proposed = [dict(row, assigned_place=key[1:]) for key, row in plan([dict(signup, signup_id=101)])]

        publish_conn = MagicMock()
        publish_conn.__enter__.return_value = publish_conn
        publish_cur = publish_conn.cursor.return_value.__enter__.return_value
        publish_cur.fetchone.side_effect = [{'n': 2}, {'is_published': False}]
        publish_cur.fetchall.side_effect = [[signup], old, old]
        publish_cur.rowcount = 1

        coverage_conn = MagicMock()
        coverage_conn.__enter__.return_value = coverage_conn
        coverage_cur = coverage_conn.cursor.return_value.__enter__.return_value
        coverage_cur.fetchall.side_effect = [[signup], old]
        coverage_cur.fetchone.return_value = {'n': 2}
        job_conn = MagicMock()
        job_conn.__enter__.return_value = job_conn
        job_cur = job_conn.cursor.return_value.__enter__.return_value
        job_cur.fetchone.side_effect = [None, {'id': 8}]

        root = Path(__file__).resolve().parent
        notification_connections = iter([coverage_conn, job_conn])
        service = dict(os=os, LOG=logging.getLogger('offline_whatsapp'), datetime=datetime, timedelta=timedelta,
                       malaysia_now=lambda: datetime(2026, 11, 19, 22, 5),
                       get_conn=lambda: next(notification_connections), _scalar=lambda *args: 2,
                       _settings=lambda keys: dict.fromkeys(keys, True))
        load_functions(root / 'schedule/services/whatsapp_publish_2200.py',
                       {'_execute_locked', 'finalize_published_roster', '_validate_coverage',
                        '_dry_interval', '_dry_minutes', '_dry_merge', '_dry_covers'}, service)

        snapshot = MagicMock()
        modules = {
            'schedule.builders.schedule_builder': SimpleNamespace(
                publish_full_schedule_2200_atomic=self.handlers['publish_full_schedule_2200_atomic']),
            'schedule.services.publish_service': SimpleNamespace(
                is_schedule_published=lambda target: publish_conn.commit.called),
            'schedule.services.whatsapp_service': SimpleNamespace(
                build_whatsapp_from_assigned=lambda target: 'offline roster'),
            'schedule.services.whatsapp_morning_summary': SimpleNamespace(
                capture_2200_snapshot=snapshot, _assignments=lambda cur, target: old,
                _key=lambda row: (row['signup_id'], row['start_time'], row['end_time'])),
        }
        dependencies = {'get_conn': lambda: publish_conn, '_capture_2200_baseline_in_transaction': MagicMock(),
                        'calculate_schedule_for_date': lambda target: (None, None, proposed, None, None)}
        with patch.dict(sys.modules, modules), patch.dict(self.handlers, dependencies), \
                patch.dict(os.environ, {'WA_2200_ALLOW_DB_WRITE': '1'}), self.assertLogs('offline_2200', 'WARNING') as logs:
            result = service['_execute_locked'](date, 'offline', True, True)
        self.assertEqual(result, {'status': 'queued', 'target_date': date, 'job_id': 8})
        self.assertIn('101', logs.output[0])
        self.assertEqual(old, before)
        self.assertFalse(any('INSERT INTO volunteer_schedule_assignments' in call.args[0]
                             for call in publish_cur.execute.call_args_list))
        publish_conn.commit.assert_called_once()
        snapshot.assert_called()
        job_conn.commit.assert_called_once()
        self.assertTrue(any('INSERT INTO whatsapp_notification_jobs' in call.args[0]
                            for call in job_cur.execute.call_args_list))

    def test_publish_validation_retains_stw_152_fixed_center_exception(self):
        validate = self.handlers['_validate_full_duty_before_publish']
        signup = dict(self.db.signup, volunteer_id='STW-152', signup_id=101)
        rows = [dict(row, assigned_place=key[1:]) for key, row in plan([signup])]
        validate([signup], rows)
        validate([signup], [dict(rows[0], end_time='6:00pm')])
        with self.assertRaisesRegex(ValueError, '报名 101.*固定活动中心'):
            validate([signup], [rows[0], dict(rows[1], assigned_place=G)])

    def test_2200_conflict_stops_before_assignment_writes_and_publish(self):
        conn = MagicMock()
        cur = conn.cursor.return_value.__enter__.return_value
        signup = self.db.signup
        proposed = [dict(row, assigned_place=key[1:])
                    for key, row in plan([dict(signup, signup_id=101)])]
        old = [dict(proposed[0], end_time='3:00pm', assignment_source='system', locked_by_admin=False)]
        cur.fetchone.return_value = {'is_published': False}
        cur.fetchall.side_effect = [[signup], old]
        with patch.dict(self.handlers, {'get_conn': lambda: conn}):
            conn.__enter__.return_value = conn
            with self.assertRaisesRegex(ValueError, '报名 101.*强制'):
                self.handlers['_publish_existing_2200_atomic']('2026-11-20', proposed)
        statements = [' '.join(call.args[0].upper().split()) for call in cur.execute.call_args_list]
        self.assertFalse(any(sql.startswith(('INSERT INTO VOLUNTEER_SCHEDULE_ASSIGNMENTS',
                                             'UPDATE SCHEDULE_DAY_FLAGS')) for sql in statements))
        conn.commit.assert_not_called()

    def test_2200_published_state_skips_all_assignment_work(self):
        conn = MagicMock()
        conn.__enter__.return_value = conn
        cur = conn.cursor.return_value.__enter__.return_value
        cur.fetchone.return_value = {'is_published': True}
        with patch.dict(self.handlers, {'get_conn': lambda: conn}):
            result = self.handlers['_publish_existing_2200_atomic']('2026-11-20', [])
        self.assertEqual(result['status'], 'already_published')
        cur.fetchall.assert_not_called()
        conn.commit.assert_not_called()
