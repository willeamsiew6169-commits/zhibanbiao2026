"""V2 business contracts. Runs production source offline with synthetic data."""

import ast
import copy
import logging
import os
import subprocess
from pathlib import Path
import sys
from datetime import date, datetime, time, timedelta
from types import SimpleNamespace
import unittest
from unittest.mock import MagicMock, patch

from schedule.builders.duty_v2 import plan, plan_for_date
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


class MealStaffingTests(unittest.TestCase):
    @staticmethod
    def roster(hall, center, start='11:00', end='13:00'):
        return [existing(f'CHE-{i}', start, end, place)
                for i, place in enumerate([G] * hall + [A] * center, 1)]

    def test_confirmed_extra_hall_targets_preserve_existing_assignments(self):
        for hall, center in ((2, 5), (3, 9), (3, 10)):
            with self.subTest(hall=hall, center=center):
                frozen = self.roster(hall, center)
                original = copy.deepcopy(frozen)
                rows = plan([member('CHE-99', '11:00', '13:00')],
                            existing=frozen, patch_mode=True, meal_activity_window=(660, 780))
                self.assertEqual(segments(rows), [(660, 780, G)])
                self.assertEqual(frozen, original)

    def test_meal_target_uses_concurrent_staffing_only_within_activity_window(self):
        frozen = self.roster(3, 4) + self.roster(0, 6, '13:00', '14:00')
        signup = member('CHE-99', '11:00', '13:00')
        self.assertEqual(segments(plan([signup], existing=frozen,
                                      meal_activity_window=(660, 780))), [(660, 780, A)])
        for window in (None, (780, 840)):
            with self.subTest(window=window):
                self.assertEqual(segments(plan([signup], existing=frozen,
                                              meal_activity_window=window)), [(660, 780, G)])

    def test_coverage_takes_precedence_over_meal_hall_target(self):
        rows = plan([member('CHE-99', '11:00', '13:00')], existing=self.roster(1, 0),
                    meal_activity_window=(660, 780))
        self.assertEqual(segments(rows), [(660, 780, A)])


class MealEntryTests(unittest.TestCase):
    MEAL_DAY = '2026-11-24'
    NORMAL_DAY = '2026-11-26'

    def setUp(self):
        settings_file = ROOT / 'schedule/services/settings_service.py'
        self.settings = {}
        self.settings_handlers = dict(datetime=datetime, time=time, timedelta=timedelta,
                                      get_schedule_settings=lambda: self.settings)
        from lunar_rules import is_meal_festival_day
        self.settings_handlers['is_meal_festival_day'] = is_meal_festival_day
        load_functions(settings_file, {'current_meal_event_date', '_meal_manual_override', 'is_meal_signup_open'},
                       self.settings_handlers)

    def test_fixed_date_zero_meal_signups_before_and_after_signup_closes(self):
        signup = member('CHE-99', '13:30', '14:00')
        frozen = MealStaffingTests.roster(3, 4, '13:30', '14:00')
        for now, open_now in ((datetime(2026, 11, 23, 22), True), (datetime(2026, 11, 24, 14), False)):
            with self.subTest(now=now), patch('lunar_rules.solar_to_lunar',
                                             side_effect=lambda day: SimpleNamespace(
                                                 month=10, day=15 if str(day) == self.MEAL_DAY else 2)):
                self.assertIs(self.settings_handlers['is_meal_signup_open'](now), open_now)
                self.assertEqual(segments(plan_for_date([signup], target_date=self.MEAL_DAY, existing=frozen)),
                                 [(810, 840, A)])

    def test_ordinary_day_with_an_undated_temporary_switch(self):
        self.settings.update(meal_manual_event_date='', meal_manual_state='true',
                             meal_manual_expires_at='2026-11-27T14:00:00')
        with patch('lunar_rules.solar_to_lunar', return_value=SimpleNamespace(month=10, day=2)):
            self.assertTrue(self.settings_handlers['is_meal_signup_open'](datetime(2026, 11, 26, 10)))
            frozen = MealStaffingTests.roster(3, 4, '13:30', '14:00')
            rows = plan_for_date([member('CHE-99', '13:30', '14:00')],
                                 target_date=self.NORMAL_DAY, existing=frozen)
        self.assertEqual(segments(rows), [(810, 840, G)])

    def test_meal_scoring_stops_at_1400(self):
        with patch('lunar_rules.solar_to_lunar', return_value=SimpleNamespace(month=10, day=15)):
            for start, end, expected in [('13:30', '14:00', A), ('14:00', '14:30', G)]:
                with self.subTest(start=start):
                    frozen = MealStaffingTests.roster(3, 4, start, end)
                    rows = plan_for_date([member('CHE-99', start, end)], target_date=self.MEAL_DAY,
                                         existing=frozen, patch_mode=True)
                    self.assertEqual(segments(rows), [(parse_min(start), parse_min(end), expected)])

    def test_ordinary_lunch_four_and_six_concurrent_workers(self):
        with patch('lunar_rules.solar_to_lunar', return_value=SimpleNamespace(month=10, day=2)):
            for total, hall in ((4, 1), (6, 2)):
                with self.subTest(total=total):
                    rows = plan_for_date([member(f'CHE-{i}', '11:00', '13:00') for i in range(total)],
                                         target_date=self.NORMAL_DAY)
                    self.assertCountEqual(locations_at(720, segments(rows)), [G] * hall + [A] * (total - hall))
        short = [member(code, '13:30', '14:30') for code in ('CHE-1', 'CHE-2')]
        rotated = plan(short)
        for code in ('CHE-1', 'CHE-2'):
            self.assertCountEqual([key[1:] for key, row in rotated if row['volunteer_id'] == code], [G, A])

    def test_dated_plan_prioritizes_the_hall_for_one_worker_with_fixed_exception(self):
        history = {'CHE-99': {'counts': {G: 10}}}
        for start, end in [('10:00', '12:00'), ('14:00', '16:00')]:
            for code, expected in [('CHE-99', G), ('STW-152', A)]:
                with self.subTest(start=start, code=code):
                    rows = plan_for_date([member(code, start, end)], target_date=self.NORMAL_DAY, history=history)
                    self.assertEqual(segments(rows), [(parse_min(start), parse_min(end), expected)])

    def test_cleaning_at_both_locations_does_not_impose_a_duty_location(self):
        handlers = {}
        load_functions(BUILDER, {'_v2_cleaning_context'}, handlers)
        assignments = {'佛堂卫生': [member('CHE-1', '08:00', '09:00'), member('CHE-2', '08:00', '09:00')],
                       '楼梯卫生': [member('CHE-2', '08:00', '09:00'), member('CHE-3', '08:00', '09:00')]}
        self.assertEqual(handlers['_v2_cleaning_context'](assignments), {'CHE-1': G, 'CHE-3': A})

    def test_shared_validator_accepts_an_automatic_handover_that_improves_coverage(self):
        signup = member('CHE-99', '10:00', '18:00', id=99, role='值班')
        rows = [dict(existing('CHE-99', '10:00', '15:00', G), signup_id=99, role='值班'),
                dict(existing('CHE-99', '15:00', '18:00', A), signup_id=99, role='值班')]
        other = [dict(existing('CHE-1', '10:00', '15:00', A), signup_id=1, role='值班'),
                 dict(existing('CHE-2', '15:00', '18:00', G), signup_id=2, role='值班')]
        handlers = dict(merge_helpers)
        load_functions(BUILDER, {'_validate_full_duty_before_publish'}, handlers)
        validate = handlers['_validate_full_duty_before_publish']
        with self.assertRaisesRegex(ValueError, '强制'):
            validate([signup], rows)
        self.assertEqual(validate([signup], rows + other), [])

    def test_formal_entrypoints_use_the_target_date(self):
        functions = {}
        for filename, names in (
                (BUILDER, {'auto_assign_duty', 'auto_assign_new_duty', '_replace_duty_signup_atomic',
                           '_plan_meal_duty_gaps', 'calculate_schedule_for_date', 'assign_jobs',
                           'run_schedule_for_date', 'publish_full_schedule_2200_atomic',
                           '_publish_existing_2200_atomic', 'repair_unassigned_duty_for_date', 'patch_schedule_for_date'}),
                (ROOT / 'schedule/schedule_web.py', {'edit_assigned_place', 'publish_schedule'})):
            functions.update({node.name: node for node in ast.parse(filename.read_text(encoding='utf-8')).body
                              if isinstance(node, ast.FunctionDef) and node.name in names})
        for name in ('auto_assign_duty', 'auto_assign_new_duty', '_replace_duty_signup_atomic',
                     '_plan_meal_duty_gaps', 'edit_assigned_place'):
            with self.subTest(name=name):
                calls = [node for node in ast.walk(functions[name]) if isinstance(node, ast.Call)
                         and isinstance(node.func, ast.Name) and node.func.id == 'plan_for_date']
                self.assertEqual(len(calls), 1)
                self.assertTrue(any(kw.arg == 'target_date' and isinstance(kw.value, ast.Name)
                                    for kw in calls[0].keywords))
        for name, callee in [('assign_jobs', 'auto_assign_duty'), ('assign_jobs', 'auto_assign_new_duty'),
                             ('calculate_schedule_for_date', 'assign_jobs'), ('run_schedule_for_date', 'calculate_schedule_for_date'),
                             ('publish_full_schedule_2200_atomic', 'calculate_schedule_for_date'),
                             ('_publish_existing_2200_atomic', '_plan_meal_duty_gaps'),
                             ('publish_schedule', 'run_schedule_for_date'), ('publish_schedule', 'repair_unassigned_duty_for_date'),
                             ('repair_unassigned_duty_for_date', 'patch_schedule_for_date'),
                             ('patch_schedule_for_date', '_replace_duty_signup_atomic')]:
            self.assertTrue(any(isinstance(node, ast.Name) and node.id == callee
                                for node in ast.walk(functions[name])), (name, callee))

    def test_2200_meal_gaps_score_the_existing_locked_roster(self):
        signup = member('CHE-99', '13:30', '14:00', id=99, role='值班')
        old = [dict(row, signup_id=i, role='值班', assignment_source='admin', locked_by_admin=True)
               for i, row in enumerate(MealStaffingTests.roster(3, 4, '13:30', '14:00'), 1)]
        before = copy.deepcopy(old)
        proposal = dict(signup, signup_id=99, assigned_place=G)
        conn = MagicMock()
        conn.__enter__.return_value = conn
        cursor = conn.cursor.return_value.__enter__.return_value
        cursor.fetchone.return_value = {'is_published': False}
        cursor.fetchall.side_effect = [[signup], old, old + [dict(proposal, assigned_place=A)]]
        cursor.rowcount = 1
        snapshot = MagicMock()
        deps = dict(merge_helpers, load_place_history_cache=lambda **kwargs: {}, _v2_cleaning_rows=lambda rows: {},
                    get_conn=lambda: conn, RealDictCursor=object, _capture_2200_baseline_in_transaction=snapshot)
        load_functions(BUILDER, {'_plan_meal_duty_gaps', '_publish_existing_2200_atomic',
                                  '_validate_full_duty_before_publish'}, deps)
        with patch('lunar_rules.solar_to_lunar', return_value=SimpleNamespace(month=10, day=15)):
            deps['_publish_existing_2200_atomic'](self.MEAL_DAY, [proposal])
        inserts = [call.args[1] for call in cursor.execute.call_args_list
                   if call.args[0].lstrip().startswith('INSERT INTO volunteer_schedule_assignments')]
        self.assertEqual([(row[0], row[6]) for row in inserts], [(99, A)])
        conn.commit.assert_called_once()
        snapshot.assert_called_once_with(cursor, self.MEAL_DAY)
        self.assertEqual(old, before)


class EmptyRosterCoverageTests(unittest.TestCase):
    def assert_full_coverage(self, rows):
        intervals = segments(rows)
        for minute in range(600, 1080):
            self.assertCountEqual(locations_at(minute, intervals), [G, A], min_to_ampm(minute))

    def assert_shared_validator_accepts(self, people, rows):
        handlers = dict(merge_helpers)
        load_functions(BUILDER, {'_validate_full_duty_before_publish'}, handlers)
        signups = [dict(person, id=person['signup_id'], role='值班') for person in people]
        assignments = [dict(row, assigned_place=key[1:]) for key, row in rows]
        self.assertEqual(handlers['_validate_full_duty_before_publish'](signups, assignments), [])

    def test_full_duty_skips_rotation_when_the_other_worker_is_fixed(self):
        people = [member('CHE-99', '10:00', '18:00', signup_id=1),
                  member('STW-152', '10:00', '18:00', signup_id=2)]
        for day in (2, 15):
            with self.subTest(lunar_day=day), patch('lunar_rules.solar_to_lunar',
                                                  return_value=SimpleNamespace(month=10, day=day)):
                rows = plan_for_date(people, target_date=DAY)
                self.assert_full_coverage(rows)
                self.assertEqual(segments([(key, row) for key, row in rows if row['signup_id'] == 1]),
                                 [(600, 840, G), (840, 1080, G)])
                self.assert_shared_validator_accepts(people, rows)

    def test_full_duty_delays_rotation_until_the_other_worker_arrives(self):
        people = [member('CHE-99', '10:00', '18:00', signup_id=1),
                  member('STW-152', '10:00', '15:30', signup_id=2),
                  member('CHE-11', '15:30', '18:00', signup_id=3)]
        rows = plan_for_date(people, target_date=DAY)
        self.assert_full_coverage(rows)
        self.assertEqual(segments([(key, row) for key, row in rows if row['signup_id'] == 1]),
                         [(600, 840, G), (840, 930, G), (930, 1080, A)])
        self.assert_shared_validator_accepts(people, rows)

    def test_full_duty_keeps_1400_rotation_when_coverage_is_already_complete(self):
        people = [member('CHE-99', '10:00', '18:00', signup_id=1),
                  member('CHE-11', '10:00', '18:00', signup_id=2)]
        history = {person['volunteer_id']: {'counts': {G: 100}} for person in people}
        rows = plan_for_date(people, target_date=DAY, history=history)
        self.assert_full_coverage(rows)
        for person in people:
            own = segments([(key, row) for key, row in rows if row['signup_id'] == person['signup_id']])
            self.assertEqual([(a, b) for a, b, _ in own], [(600, 840), (840, 1080)])
            self.assertCountEqual([place for _, _, place in own], [G, A])


class BusinessPlannerTests(unittest.TestCase):
    def assert_staffing(self, rows, start, end, hall, center):
        for minute in range(start, end):
            locations = locations_at(minute, rows)
            self.assertEqual((locations.count(G), locations.count(A)), (hall, center),
                             f'{min_to_ampm(minute)}: {locations}')

    def uncovered_by_place(self, rows):
        return {
            place: sum(
                not any(start <= minute < end and assigned == place
                        for start, end, assigned in rows)
                for minute in range(600, 1080)
            )
            for place in (G, A)
        }

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

    def test_full_day_uses_1500_handover_to_remove_60_minute_gap(self):
        """A nonstandard switch is justified only by strictly better coverage."""
        signup = member('CHE-1', '10:00', '18:00')
        frozen = [existing('CHE-2', '10:00', '15:00', A),
                  existing('CHE-3', '15:00', '18:00', G)]
        frozen_segments = [
            (parse_min(row['start_time']), parse_min(row['end_time']), row['assigned_place'])
            for row in frozen
        ]
        standard_options = [
            [(600, 840, G), (840, 1080, A)],
            [(600, 840, A), (840, 1080, G)],
            [(600, 840, G), (840, 1080, G)],
            [(600, 840, A), (840, 1080, A)],
        ]
        best_standard_gaps = min(
            (self.uncovered_by_place(option + frozen_segments) for option in standard_options),
            key=lambda gaps: sum(gaps.values()),
        )
        self.assertEqual(best_standard_gaps, {G: 60, A: 0})

        result = segments(plan([signup], existing=frozen))
        self.assertEqual(result, [(600, 840, G), (840, 900, G), (900, 1080, A)])
        self.assertEqual(self.uncovered_by_place(result + frozen_segments), {G: 0, A: 0})

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
        with patch('lunar_rules.solar_to_lunar', return_value=SimpleNamespace(day=15, month=10)):
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
            '_validate_full_duty_before_publish', '_publish_existing_2200_atomic', '_plan_meal_duty_gaps',
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
                                       _wa6_minutes=merge_helpers['_wa6_minutes'],
                                       _validate_full_duty_before_publish=self.handlers['_validate_full_duty_before_publish'])
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
        locked = [dict(old[0], end_time='3:00pm', assignment_source='admin', locked_by_admin=True),
                  dict(old[1], start_time='3:00pm', assigned_place=A)]
        namespace['db_query'].side_effect = [{'n': len(locked)}, [signup], locked]
        with patch.dict(sys.modules, {'schedule.builders.schedule_builder': dependencies}):
            _, status = namespace['publish_schedule']()
        self.assertEqual((status, publish.call_count), (302, 1),
                         'Manual publication must preserve the valid locked exception accepted by 22:00.')


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

    def test_formal_dated_plan_covers_both_and_places_final_worker_at_the_hall(self):
        people = [member(f'CHE-{index}', min_to_ampm(start), min_to_ampm(end), name=name)
                  for index, (name, (start, end)) in enumerate(self.WINDOWS.items(), 1)]
        rows = plan_for_date(people, target_date=DAY)
        for minute in range(600, 1020):
            self.assertEqual(set(locations_at(minute, segments(rows))), {G, A}, min_to_ampm(minute))
        for minute in range(1020, 1080):
            self.assertEqual(locations_at(minute, segments(rows)), [G], min_to_ampm(minute))
        for name in self.CROSS_SHIFT:
            own = segments([(key, row) for key, row in rows if row['name'] == name])
            self.assertEqual([(a, b) for a, b, _ in own], [(self.WINDOWS[name][0], 840), (840, self.WINDOWS[name][1])])
            self.assertCountEqual([place for _, _, place in own], [G, A])


if __name__ == '__main__':
    # Block the application DB module and load only the five offline modules.
    with patch.dict(sys.modules, {'db': None}):
        suite = unittest.defaultTestLoader.loadTestsFromNames([
            'test_duty_v2', 'test_duty_v2_integration', 'test_duty_v2_safety',
            'test_duty_v2_full_day', '__main__',
        ])
        result = unittest.TextTestRunner(verbosity=2).run(suite)
    mobile_status = 0
    if '--with-mobile' in sys.argv:
        mobile_status = subprocess.run([sys.executable, '-B', str(ROOT / 'test_mobile_volunteer.py')],
                                       cwd=ROOT).returncode
    sys.exit(0 if result.wasSuccessful() and mobile_status == 0 else 1)
