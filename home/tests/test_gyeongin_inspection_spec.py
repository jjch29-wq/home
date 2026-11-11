"""Free-form inspection specs stay separate from pipe calculations and report groups."""
from pathlib import Path
import sys
import unittest
import ast
import math
from types import SimpleNamespace
from unittest.mock import Mock, patch
from test_gyeongin_aux_edit_preservation import Field
import pandas as pd
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.gyeongin.src.services.report_inspection import collect_inspection_details


class InspectionSpecTests(unittest.TestCase):
    def test_non_pipe_rt_is_included_without_diameter(self):
        rows = pd.DataFrame([{'검사방법': 'RT', '검사규격': '두께 12 mm', 'Usage': 10, '불량수': 2}])
        result = collect_inspection_details(rows, {})
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]['pipe_size'], '두께 12 mm')
        self.assertEqual(result[0]['quantity'], 10)
        self.assertEqual(result[0]['reject_films'], 2)

    def test_specs_are_not_merged_and_rep_remains_separate(self):
        rows = pd.DataFrame([
            {'검사방법': 'RT', '검사규격': spec, '관경(Inch)': '', 'Usage': 10, '검사구분': kind}
            for spec, kind in [('두께 12 mm', 'ORI'), ('두께 20 mm', 'ORI'), ('두께 12 mm', 'REP')]])
        self.assertEqual(len(collect_inspection_details(rows, {})), 2)

    def test_spec_never_becomes_a_pipe_length_lookup(self):
        rows = pd.DataFrame([{'검사방법': 'MT', '검사규격': '300', '관경(Inch)': '', 'Usage': 20}])
        result = collect_inspection_details(rows, {'300A': 2})
        self.assertEqual(result[0]['points'], 0)

    def test_existing_pipe_calculation_is_preserved(self):
        rows = pd.DataFrame([{'검사방법': 'MT', '관경(Inch)': '300', 'Usage': 20}])
        result = collect_inspection_details(rows, {'300A': 2})
        self.assertEqual(result[0]['points'], 10)
        self.assertEqual(result[0]['pipe_size'], '300')


class InspectionTargetTests(unittest.TestCase):
    def callback(self, name, app, **extra):
        path = Path(__file__).resolve().parents[1] / 'src/site_apps/gyeongin/src/views/daily_usage_view.py'
        tree = ast.parse(path.read_text(encoding='utf-8'))
        fn = next(n for n in ast.walk(tree) if isinstance(n, ast.FunctionDef) and n.name == name)
        scope = {'self': app, 'math': math, 'tk': SimpleNamespace(END='end'), **extra}
        exec(compile(ast.Module(body=[fn], type_ignores=[]), str(path), 'exec'), scope)
        return scope[name]

    def test_target_switch_changes_labels_and_clears_pipe_only(self):
        app = SimpleNamespace(inspection_target_var=Field('배관 외'), cb_daily_test_method=Field('RT'),
                              cb_ndt_report_pipe=Mock(), ndt_report_pipe_var=Field('300A'))
        labels = {name: Mock() for name in ('ori_count_label', 'rep_count_label', 'ori_qty_label', 'rep_qty_label', 'calc_button')}
        update = self.callback('update_inspection_target', app, **labels)
        update()
        labels['ori_count_label'].configure.assert_called_with(text='[ORI] 검사 개소:')
        labels['rep_count_label'].configure.assert_called_with(text='[REP] 재검사 개소:')
        labels['ori_qty_label'].configure.assert_called_with(text='촬영 매수:')
        self.assertEqual(app.ndt_report_pipe_var.get(), '')
        app.cb_ndt_report_pipe.configure.assert_called_with(state='disabled')
        app.inspection_target_var.set('배관'); update()
        labels['ori_count_label'].configure.assert_called_with(text='[ORI] 조인트 수:')
        app.cb_ndt_report_pipe.configure.assert_called_with(state='normal')

    def test_non_pipe_fee_does_not_recalculate_quantities(self):
        app = SimpleNamespace(cb_daily_test_method=Field('RT'), inspection_target_var=Field('배관 외'),
                              ent_daily_test_amount=Field('12'), ent_daily_unit_price=Field('100'),
                              ent_daily_test_fee=Field())
        with patch('site_apps.gyeongin.src.services.inspection_quantity.quantity_per_joint', side_effect=AssertionError('pipe calculation')):
            self.callback('_calculate_ndt_fee', app)()
        self.assertEqual(app.ent_daily_test_amount.get(), '12')
        self.assertEqual(app.ent_daily_test_fee.get(), '1,200')

    def test_manual_rt_originals_and_repeats_stay_separate(self):
        app = SimpleNamespace(inspection_target_var=Field('배관 외'), cb_daily_test_method=Field('RT'),
                              ndt_ori_qty_var=Field('12'), ndt_rep_qty_var=Field('3'),
                              ent_daily_test_amount=Field(), update_daily_test_fee_calc=Mock())
        update = self.callback('sync_manual_quantities', app)
        update()
        self.assertEqual(app.ent_daily_test_amount.get(), '12')
        self.assertEqual(app.ndt_rep_qty_var.get(), '3')
        app.cb_daily_test_method.set('MT'); update()
        self.assertEqual(app.ent_daily_test_amount.get(), '15')
        app.inspection_target_var.set('배관'); app.ent_daily_test_amount.set('99'); update()
        self.assertEqual(app.ent_daily_test_amount.get(), '99')


if __name__ == '__main__':
    unittest.main()
