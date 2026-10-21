import ast
from copy import deepcopy
from pathlib import Path
import sys
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch, mock_open

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / 'home/src'))
from site_apps.central.src.rt_quantities import rt_counts, film_grades, sync_rt_widgets, corrected_rt_history, rt_result
from site_apps.central.src.rt_quantities import resolve_rt_row, widget_rt_values, retest_parent, default_rt_copy, effective_section_info


class Field:
    def __init__(self, value=''):
        self.value = str(value)

    def get(self, *args):
        return self.value

    def set(self, value):
        self.value = str(value)

    def delete(self, *args):
        self.value = ''

    def insert(self, index, value):
        self.value = str(value)


def widgets(info='N/1,이물질/4,N/1'):
    return {key: Field(value) for key, value in {'검사방법': 'RT', '구간정보': info,
            'RT_OR': '3', 'RT_RE': '', '결과': '합격', '용접사': 'W-2026-A-08'}.items()}


class RtQuantitiesTests(unittest.TestCase):
    def test_new_joint_and_retake_default_readings_are_editable(self):
        original = {'검사방법': 'RT', 'Joint No.': 'R06', '관경': '100A',
                    '구간정보': 'N/1,이물질/4,N/1', 'RT_OR': '3'}
        retake = default_rt_copy(dict(original, **{'Joint No.': 'R06R', '구간정보': '', 'RT_OR': '0'}), [original])
        self.assertEqual(retake['구간정보'], ',N/1')
        self.assertEqual(rt_counts(retake), (1, 0))
        self.assertEqual(rt_result(effective_section_info(retake), 1), '합격')
        next_joint = default_rt_copy(dict(retake, **{'Joint No.': 'R07', '구간정보': ''}), [original, retake])
        self.assertEqual(next_joint['구간정보'], 'N/1,N/1,N/1')
        self.assertEqual(rt_counts(next_joint), (3, 0))
        edited = dict(retake, 구간정보=',이물질/4')
        self.assertEqual(default_rt_copy(edited, [original]), edited)
        self.assertEqual(rt_counts(edited), (1, 1))

    def test_zero_original_is_recovered_from_entered_film_grades(self):
        for info, rejected, result in (('N/1,N/1,N/1', 0, '합격'),
                                        ('N/1,이물질/4,N/1', 1, '불합격')):
            row = widgets(info)
            row['RT_OR'].set('0')
            row['RT_RE'].set('0')
            sync_rt_widgets(row, from_details=True)
            self.assertEqual(row['RT_OR'].get(), '3')
            self.assertEqual(row['RT_RE'].get(), str(rejected))
            self.assertEqual(row['결과'].get(), result)
            self.assertEqual(rt_counts({'RT_OR': '0', '구간정보': info}), (3, rejected))

    def test_retake_only_counts_failed_position_and_uses_new_grade(self):
        original = {'검사방법': 'RT', 'Joint No.': 'R06', '구간정보': 'N/1,이물질/4,N/1', 'RT_OR': '3'}
        copied = dict(original, **{'Joint No.': 'R06R'})
        retake = resolve_rt_row(copied, [original])
        self.assertEqual(retake['재촬영구간'], [2])
        self.assertEqual(retake['구간정보'], ',이물질/4')
        self.assertEqual(rt_counts(retake), (1, 1))
        retake['구간정보'] = ',N/1'
        self.assertEqual(rt_counts(retake), (1, 0))
        self.assertEqual(original['구간정보'], 'N/1,이물질/4,N/1')
        self.assertEqual(retest_parent('R06R'), 'R06')
        self.assertEqual(retest_parent('R06R1'), 'R06')
        self.assertEqual(retest_parent('R06'), '')
        history = {'2026-10-20': {'ndt_results': [original, copied],
                    'qty_data': {'RT_소계': {'전일누계': '0'}}}}
        corrected = corrected_rt_history(history)
        self.assertEqual(corrected['2026-10-20']['qty_data']['RT_소계']['금일작업'], '4')

    def test_result_from_grades_one_two_pass_three_four_fail(self):
        for value in ('N/1,N/2,N/1', '*/1,*/2,*/2'):
            with self.subTest(value=value):
                self.assertEqual(rt_result(value, 3), '합격')
        for value in ('N/1,이물질/3,N/2', 'N/1,이물질/4,N/1', '이물질/4,N/', '1/4'):
            with self.subTest(value=value):
                self.assertEqual(rt_result(value, 3), '불합격')
        for value in ('', 'N/1', 'N/1,N/', '/1,/2,/1'):
            with self.subTest(value=value):
                self.assertEqual(rt_result(value, 3), '')

    def test_result_does_not_remain_pass_after_clearing_or_incomplete_edit(self):
        row = widgets('N/1,N/2,N/1')
        sync_rt_widgets(row)
        self.assertEqual(row['결과'].get(), '합격')
        row['구간정보'].set('N/1,N/2,N/')
        sync_rt_widgets(row, from_details=True)
        self.assertEqual(row['결과'].get(), '')
        row['구간정보'].set('N/1,N/2,이물질/4')
        sync_rt_widgets(row, from_details=True)
        self.assertEqual(row['결과'].get(), '불합격')
        self.assertEqual(row['RT_RE'].get(), '1')
        row['구간정보'].set('')
        sync_rt_widgets(row, from_details=True)
        self.assertEqual(row['결과'].get(), '')
    def test_three_films_one_reject_is_three_total(self):
        row = {'RT_OR': '3', 'RT_RE': '1'}
        self.assertEqual(rt_counts(row), (3, 1))
        self.assertEqual(rt_counts(dict(row, 구간정보='N/1,이물질/4,N/1')), (3, 1))
        self.assertEqual(rt_counts(dict(row, 구간정보='1/N,4/SI,1/N')), (3, 1))
        self.assertEqual(rt_counts(dict(row, RT_RE='0', 구간정보='N/1,이물질/4,N/1')), (3, 1))

    def test_input_updates_reject_quantity_and_result_without_increasing_or(self):
        row = widgets()
        sync_rt_widgets(row)
        self.assertEqual(row['RT_OR'].get(), '3')
        self.assertEqual(row['RT_RE'].get(), '1')
        self.assertEqual(row['결과'].get(), '불합격')
        row['구간정보'].set('N/1,N/1,N/1')
        sync_rt_widgets(row, from_details=True)
        self.assertEqual(row['RT_RE'].get(), '0')
        self.assertEqual(row['결과'].get(), '합격')
        row['구간정보'].set('')
        row['RT_RE'].set('1')
        sync_rt_widgets(row, from_details=True)
        self.assertEqual(row['RT_RE'].get(), '0')

    def test_bad_counts_are_rejected_and_retest_is_not_inferred_from_re(self):
        for row in ({'RT_OR': 3, 'RT_RE': 4}, {'RT_OR': 'nan'}, {'RT_OR': 3, 'RT_RE': -.5}):
            with self.subTest(row=row), self.assertRaises(ValueError):
                rt_counts(row)
        self.assertIsNone(film_grades('이물질/'))
        row = widgets()
        row['결과'].set('재촬영')
        sync_rt_widgets(row)
        self.assertEqual(row['결과'].get(), '재촬영')

    def test_report_cumulative_counts_only_or_and_keeps_re_separate(self):
        history = {}
        for day, total, rejects, old_sum in (('2026-10-19', 3, 1, 4), ('2026-10-20', 6, 2, 12)):
            history[day] = {'ndt_results': [{'검사방법': 'RT', '관경': '100A', 'RT_OR': total, 'RT_RE': rejects}],
                            'qty_data': {key: {'전일누계': '0', '금일작업': str(old_sum), '총누계': str(old_sum), '예상량': '100'}
                                         for key in ('RT_150A~100A', 'RT_소계')}}
        original = deepcopy(history)
        corrected = corrected_rt_history(history)
        self.assertEqual(corrected['2026-10-19']['qty_data']['RT_소계']['금일작업'], '3')
        self.assertEqual(corrected['2026-10-20']['qty_data']['RT_소계']['금일작업'], '6')
        self.assertEqual(corrected['2026-10-20']['qty_data']['RT_소계']['총누계'], '9')
        self.assertEqual(history, original)
        self.assertEqual(corrected['2026-10-20']['ndt_results'][0]['RT_RE'], 2)
        self.assertEqual(corrected['2026-10-19']['qty_data']['RT_소계']['불량'], '1')
        self.assertEqual(corrected['2026-10-19']['qty_data']['RT_소계']['불량률'], '33.3')
        self.assertEqual(corrected['2026-10-20']['qty_data']['RT_소계']['불량'], '3')
        self.assertEqual(corrected['2026-10-20']['qty_data']['RT_소계']['불량률'], '33.3')

    def test_native_auto_calculation_counts_three_total_and_one_defect(self):
        source = ROOT / 'home/src/site_apps/central/src/daily_work_log_tab.py'
        tree = ast.parse(source.read_text(encoding='utf-8'))
        cls = next(n for n in tree.body if isinstance(n, ast.ClassDef) and n.name == 'DailyWorkLogTab')
        method = next(n for n in cls.body if isinstance(n, ast.FunctionDef) and n.name == 'auto_calculate_and_save')
        namespace = {'rt_counts': rt_counts, 'corrected_rt_history': corrected_rt_history,
                     'tk': SimpleNamespace(END='end'), 'messagebox': Mock()}
        exec(compile(ast.fix_missing_locations(ast.Module(body=[method], type_ignores=[])), '<actual auto totals>', 'exec'), namespace)
        app = SimpleNamespace(_validate_rt_rows=lambda: True, _validate_hot_rows=lambda: True,
                _refresh_hot_summary=Mock(), date_entry=Field('2026-10-19'), load_history=lambda: {},
                equip_entries={}, save_current_history=Mock())
        app.ndt_grid_entries = [dict(widgets(), 관경=Field('100A'), 근무구분=Field(''))]
        keys = []
        for name, prefix in (('paut_keys', 'PAUT_'), ('rt_keys', 'RT_')):
            values = next(ast.literal_eval(n.value) for n in ast.walk(method) if isinstance(n, ast.Assign)
                          and any(isinstance(t, ast.Name) and t.id == name for t in n.targets))
            keys += [prefix + value for value in values] + [prefix + '소계']
        app.qty_entries = {key: {field: Field() for field in ('예상량', '전일누계', '금일작업', '총누계', '공정률', '불량', '불량률')}
                           for key in keys}
        app.personnel_entries = {key: Field() for key in ('검사원_누계', '검사원_인원', '검사원_현장대리인', '안전_누계', '안전_인원', '안전_현장대리인')}
        namespace['auto_calculate_and_save'](app)
        self.assertEqual(app.qty_entries['RT_소계']['금일작업'].get(), '3')
        self.assertEqual(app.qty_entries['RT_소계']['총누계'].get(), '3')
        self.assertEqual(app.qty_entries['RT_소계']['불량'].get(), '1')
        self.assertEqual(app.qty_entries['RT_소계']['불량률'].get(), '33.3')
        app.save_current_history.assert_called_once()
        namespace['messagebox'].showinfo.assert_called_once()
        namespace['messagebox'].reset_mock()
        app.save_current_history.reset_mock()
        namespace['auto_calculate_and_save'](app, show_completion=False)
        app.save_current_history.assert_called_once()
        namespace['messagebox'].showinfo.assert_not_called()

    def test_native_monthly_total_expressions_do_not_add_rt_re(self):
        source = ROOT / 'home/src/site_apps/central/src/services/monthly_report_manager.py'
        tree = ast.parse(source.read_text(encoding='utf-8'))
        for target, values in (('tot_v', {'ori_v': 3, 're_v': 1}), ('total_qty', {'total_ori': 3, 'total_re': 1})):
            assignment = next(n for n in ast.walk(tree) if isinstance(n, ast.Assign)
                              and any(isinstance(t, ast.Name) and t.id == target for t in n.targets))
            expression = compile(ast.Expression(assignment.value), '<actual monthly total>', 'eval')
            self.assertEqual(eval(expression, dict(values, method='RT')), 3)
            self.assertEqual(eval(expression, dict(values, method='PAUT')), 4)

    def test_native_daily_save_records_rt_or_3_and_rt_re_1_and_keeps_photos(self):
        source = ROOT / 'home/src/site_apps/central/src/daily_work_log_tab.py'
        tree = ast.parse(source.read_text(encoding='utf-8'))
        cls = next(n for n in tree.body if isinstance(n, ast.ClassDef) and n.name == 'DailyWorkLogTab')
        methods = [n for n in cls.body if isinstance(n, ast.FunctionDef) and n.name in
                   ('_rt_context', '_validate_rt_rows', 'get_current_data', 'save_current_history')]
        namespace = {'rt_counts': rt_counts, 'sync_rt_widgets': sync_rt_widgets, 'messagebox': Mock(),
                     'widget_rt_values': widget_rt_values, 'resolve_rt_row': resolve_rt_row}
        module = ast.Module(body=[ast.ClassDef(name='App', bases=[], keywords=[], body=methods, decorator_list=[])], type_ignores=[])
        exec(compile(ast.fix_missing_locations(module), '<actual daily save>', 'exec'), namespace)
        app = namespace['App']()
        app.ndt_grid_entries = [widgets() for _ in range(45)]
        app.date_entry = Field('2026-10-19')
        app.weather_entry = Field('맑음')
        app.remarks_text = Field('메모')
        app.qty_entries = app.equip_entries = app.personnel_entries = {}
        app._normalize_welder_id = lambda name: name
        app._selected_date_key = lambda: '2026-10-19'
        app._validate_hot_rows = Mock(return_value=True)
        app.load_history = lambda: {'2026-10-19': {'process_photos': [{'file': 'photo.jpg'}]}}
        app.save_history = Mock()
        with patch('builtins.open', mock_open()):
            app.save_current_history()
        saved = app.save_history.call_args.args[0]['2026-10-19']
        self.assertEqual(len(saved['ndt_results']), 45)
        self.assertEqual(saved['ndt_results'][0]['RT_OR'], '3')
        self.assertEqual(saved['ndt_results'][0]['RT_RE'], '1')
        self.assertEqual(saved['process_photos'], [{'file': 'photo.jpg'}])


if __name__ == '__main__':
    unittest.main()
