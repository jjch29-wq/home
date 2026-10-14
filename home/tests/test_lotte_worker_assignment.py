"""검사별 담당자 및 일일 근무 실적 중복 방지 검증."""
from pathlib import Path
import ast
import sys
from types import SimpleNamespace
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.lotte.src.worker_assignment import build_worker_assignments, worker_names, update_work_details, apply_work_details, merge_worker_attendance


class WorkerAssignmentTests(unittest.TestCase):
    def test_attendance_combines_methods_without_repeating_time_or_ot(self):
        rows = merge_worker_attendance([
            dict(names='주진철', company='함감속기', method='UT', ot_hours='09:00~18:00', ot_amount='1'),
            dict(names='주진철', company='함감속기', method='MT', ot_hours='', ot_amount='0'),
            dict(names='주진철', company='함감속기', method='UT', ot_hours='09:00~18:00', ot_amount='1')])
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]['method'], 'UT, MT')
        self.assertEqual(rows[0]['ot_hours'], '09:00~18:00')
        self.assertEqual(rows[0]['ot_amount'], '1')

    def test_attendance_keeps_distinct_workers_methods_separate(self):
        rows = merge_worker_attendance([
            dict(names='주진철', company='함감속기', method='UT'),
            dict(names='다른담당', company='함감속기', method='MT')])
        self.assertEqual([row['method'] for row in rows], ['UT', 'MT'])
    def setUp(self):
        self.workers = {
            1: dict(name='김검사', time='(주간) 8', ot='2', meal='10000', method=''),
            2: dict(name='이검사', time='(야간) 4', ot='1', meal='5000', method='RT'),
        }

    def test_distinct_teams_keep_their_own_time(self):
        pt, rt = build_worker_assignments([
            dict(method='PT', workers=[1]), dict(method='RT', workers=[2])], self.workers)
        self.assertEqual((pt['User'], pt['User2']), ('김검사', ''))
        self.assertEqual((rt['User'], rt['User2']), ('', '이검사'))
        self.assertEqual(rt['WorkTime2'], '(야간) 4')
        self.assertEqual(rt['OT2'], '1')
        self.assertEqual(rt['WorkerMethod2'], 'RT')

    def test_shared_worker_is_identified_in_each_item_but_counted_once(self):
        rows = build_worker_assignments([
            dict(method='PT', workers=[1]), dict(method='UT', workers=[1])], self.workers)
        self.assertEqual(rows[1]['User'], '김검사')
        self.assertEqual(rows[1]['WorkerMethod'], 'UT')
        self.assertEqual(rows[0]['WorkTime'], '(주간) 8')
        for field in ('WorkTime', 'OT', 'Meal'):
            self.assertEqual(rows[1][field], '')

    def test_legacy_items_use_worker_method_and_empty_selection_is_preserved(self):
        pt, rt, empty = build_worker_assignments([
            dict(method='PT'), dict(method='RT'), dict(method='UT', workers=[])], self.workers)
        self.assertEqual(pt['User2'], '')
        self.assertEqual(rt['User2'], '이검사')
        self.assertEqual(empty['User'], '')

    def test_single_item_assignment(self):
        row, = build_worker_assignments([dict(method='UT', workers=[2])], self.workers)
        self.assertEqual(row['User'], '')
        self.assertEqual(row['User2'], '이검사')
        self.assertEqual(row['WorkerMethod2'], 'UT')

    def test_items_created_before_workers_inherit_current_workers(self):
        row, = build_worker_assignments([dict(method='UT', workers=[], workers_explicit=False)], self.workers)
        self.assertEqual(row['User'], '김검사')
        self.assertEqual(row['User2'], '')
        empty, = build_worker_assignments([dict(method='UT', workers=[], workers_explicit=True)], self.workers)
        self.assertEqual(empty['User'], '')

    def test_display_includes_workers_when_first_slot_is_empty(self):
        self.assertEqual(worker_names(dict(User='', User2='이검사', User3='nan',
                                          User4='김검사', User5='이검사')), '이검사, 김검사')

    def test_loaded_record_saves_item_and_workers_without_changing_quantity(self):
        record = dict(검사품명='', User='', User2='이전담당', Usage=6, 검사비=172000,
                      차량번호='76마3422', 주행거리='100')
        workers = {1: dict(name='부장 주진철', time='(주간) 09:00~18:00',
                           ot='1', meal='10000', method='UT')}
        self.assertTrue(update_work_details(record, 'Gear Shaft', workers))
        self.assertEqual(record['검사품명'], 'Gear Shaft')
        self.assertEqual(record['User'], '부장 주진철')
        self.assertEqual(record['WorkTime'], '(주간) 09:00~18:00')
        self.assertEqual(record['User2'], '')
        self.assertEqual((record['Usage'], record['검사비'], record['주행거리']), (6, 172000, '100'))
        self.assertFalse(update_work_details(record, 'Gear Shaft', workers))

    def test_edit_with_numeric_empty_excel_columns(self):
        import pandas as pd
        frame = pd.DataFrame({'검사품명': [float('nan')], 'User': [float('nan')],
                              'WorkTime': [float('nan')], 'OT': [0.0], 'Meal': [0.0]})
        record = frame.loc[0].to_dict()
        update_work_details(record, 'Gear, Shaft', {
            1: dict(name='주진철', time='(주간) 09:00~18:00', ot='', meal='10,000', method='UT')})
        apply_work_details(frame, 0, record)
        self.assertEqual(frame.at[0, '검사품명'], 'Gear, Shaft')
        self.assertEqual(frame.at[0, 'User'], '주진철')
        self.assertEqual(frame.at[0, 'OT'], 0.0)
        self.assertEqual(frame.at[0, 'Meal'], 10000.0)
        self.assertEqual(frame.at[0, 'OT10'], 0.0)

    def test_session_preserves_item_assignments_and_worker_details(self):
        source = Path(__file__).resolve().parents[1] / 'src/site_apps/lotte/src/app.py'
        tree = ast.parse(source.read_text(encoding='utf-8-sig'))
        method = next(node for node in ast.walk(tree)
                      if isinstance(node, ast.FunctionDef) and node.name == '_get_usage_session_data')
        namespace = {}
        exec(compile(ast.Module(body=[method], type_ignores=[]), str(source), 'exec'), namespace)
        field = SimpleNamespace(get=lambda: '', get_date=lambda: '2026-10-13')
        app = SimpleNamespace(rtk_entries={}, _multi_materials=[dict(method='PT', workers=[1])])
        for attribute in ('ent_daily_date', 'cb_daily_site', 'cb_daily_company',
                          'ent_daily_applied_code', 'ent_daily_report_no', 'cb_daily_equip',
                          'cb_daily_material', 'ent_daily_inspection_item', 'cb_daily_test_method',
                          'cb_daily_unit', 'ent_daily_test_amount', 'ent_daily_unit_price',
                          'ent_daily_travel_cost', 'ent_daily_meal_cost', 'ent_daily_test_fee',
                          'ent_daily_note'):
            setattr(app, attribute, field)
        app.worker_group1 = SimpleNamespace(cb_name=field, cb_shift=field,
            ent_worktime=field, ent_ot=field, get_meal=lambda: '10000', get_method=lambda: 'PT')
        session = namespace['_get_usage_session_data'](app)
        self.assertEqual(session['multi_materials'], app._multi_materials)
        self.assertEqual(session['workers'][0]['meal'], '10000')
        self.assertEqual(session['workers'][0]['method'], 'PT')


if __name__ == '__main__':
    unittest.main()
