"""검사별 담당자 및 일일 근무 실적 중복 방지 검증."""
from pathlib import Path
import ast
import sys
from types import SimpleNamespace
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.lotte.src.worker_assignment import build_worker_assignments, worker_names


class WorkerAssignmentTests(unittest.TestCase):
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

    def test_display_includes_workers_when_first_slot_is_empty(self):
        self.assertEqual(worker_names(dict(User='', User2='이검사', User3='nan',
                                          User4='김검사', User5='이검사')), '이검사, 김검사')

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
