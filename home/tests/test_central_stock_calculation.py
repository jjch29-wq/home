"""Inventory display and reorder checks share the same consumption calculation."""
from pathlib import Path
import re
import ast
from unittest.mock import Mock
import sys
from types import SimpleNamespace
import unittest
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.central.src.models.stock_history import material_history
from site_apps.central.src.models.stock_calculation import build_stock_lookups, stock_breakdown
from site_apps.central.src.models.material_model import calculate_current_stock_impl


def fixture(materials, transactions=(), daily=()):
    def clean(value): return '' if pd.isna(value) else str(value).strip()
    def normalize(value): return re.sub(r'\.0$', '', clean(value))
    return SimpleNamespace(
        materials_df=pd.DataFrame(materials),
        transactions_df=pd.DataFrame(transactions), daily_usage_df=pd.DataFrame(daily),
        clean_nan=clean, normalize_id=normalize,
        _is_consumable_material=lambda name, method: any(k in name.lower() for k in ['film', '세척제']))


class StockCalculationTests(unittest.TestCase):
    def check_stock(self, app, expected):
        before = [df.copy(deep=True) for df in (app.materials_df, app.transactions_df, app.daily_usage_df)]
        lookups = build_stock_lookups(app)
        row = app.materials_df.iloc[0]
        self.assertEqual(stock_breakdown(app, row, lookups)['current_stock'], expected)
        self.assertEqual(calculate_current_stock_impl(app, row['MaterialID'], lookups), expected)
        self.assertEqual(calculate_current_stock_impl(app, row['MaterialID']), expected)
        summary, rows = material_history(app, row)
        self.assertEqual(summary['current_stock'], expected)
        self.assertAlmostEqual(summary['stored_qty'] + sum(r['increase'] - r['decrease'] for r in rows), expected)
        for actual, original in zip((app.materials_df, app.transactions_df, app.daily_usage_df), before):
            pd.testing.assert_frame_equal(actual, original)

    def test_reorder_alert_uses_consumption_and_reorder_threshold(self):
        app = fixture([{'MaterialID':'1', '품목명':'film', '수량':100, '재고하한':95}], daily=[
            {'MaterialID':'1', '품목명':'film', '검사방법':'RT', 'FilmCount':10}])
        app.calculate_current_stock = lambda mat_id, **kw: calculate_current_stock_impl(app, mat_id, **kw)
        source = Path(__file__).resolve().parents[1] / 'src/site_apps/central/src/app.py'
        fn = next(n for n in ast.walk(ast.parse(source.read_text(encoding='utf-8')))
                  if isinstance(n, ast.FunctionDef) and n.name == 'show_low_stock')
        messages = Mock()
        ns = {'pd':pd, 'messagebox':messages}
        exec(compile(ast.Module(body=[fn], type_ignores=[]), str(source), 'exec'), ns)
        ns['show_low_stock'](app)
        messages.showwarning.assert_called_once()
        self.assertIn('90', messages.showwarning.call_args.args[1])
        messages.showinfo.assert_not_called()

    def test_equipment_primary_auxiliary_and_id_history_do_not_reduce_stock(self):
        app = fixture([{'MaterialID':'2', '품목명':'PAUT', '모델명':'MX2', '수량':1}], daily=[
            {'Date':'2026-10-23', 'Site':'A', 'MaterialID':'', '장비명':'MX2', 'Usage':30, '검사방법':'PAUT'},
            {'Date':'2026-10-24', 'Site':'B', 'MaterialID':'film', '보조장비':'Yoke; MX2', '검사방법':'RT'},
            {'Date':'2026-10-25', 'Site':'C', 'MaterialID':'2', '장비명':'old label', '검사방법':'PAUT'},
            {'Date':'2026-10-26', 'Site':'D', 'MaterialID':'3', '장비명':'MX20', '검사방법':'PAUT'},
        ])
        summary, rows = material_history(app, app.materials_df.iloc[0])
        self.assertEqual(summary['current_stock'], 1)
        self.assertEqual([r['site'] for r in rows], ['C', 'B', 'A'])
        self.assertTrue(all(r['kind'] == '장비 사용' and r['increase'] == r['decrease'] == 0 for r in rows))

    def test_direct_chemical_and_json_mirror_are_not_added_together(self):
        self.check_stock(fixture([{'MaterialID':'3','품목명':'세척제','수량':20}], daily=[
            {'MaterialID':'','검사방법':'MT','NDT_세척제':3,'ndt_data':'{"세척제":3}'}]),17)
        self.check_stock(fixture([{'MaterialID':'3','품목명':'세척제','수량':20}], daily=[
            {'MaterialID':'','검사방법':'MT','ndt_data':'{"세척제":3}'}]),17)

    def test_schema_default_zero_still_recovers_legacy_json_usage(self):
        self.check_stock(fixture([{'MaterialID':'3','품목명':'세척제','수량':20}], daily=[
            {'MaterialID':'','검사방법':'MT','NDT_세척제':0,'ndt_data':'{"세척제":3}'}]),17)

    def test_daily_film_use_reduces_reorder_stock(self):
        self.check_stock(fixture([{'MaterialID':'1', '품목명':'film', '수량':100}], daily=[
            {'MaterialID':'1.0', '품목명':'film', '검사방법':'RT', 'FilmCount':10}]), 90)

    def test_manual_in_out_and_auto_deduction_are_not_double_counted(self):
        self.check_stock(fixture([{'MaterialID':'1', '품목명':'film', '수량':100}], transactions=[
            {'MaterialID':'1', 'Type':'IN', 'Quantity':20, 'Note':''},
            {'MaterialID':'1', 'Type':'OUT', 'Quantity':-5, 'Note':''},
            {'MaterialID':'1', 'Type':'OUT', 'Quantity':-10, 'Note':'자동 차감'}], daily=[
            {'MaterialID':'1', '품목명':'film', '검사방법':'RT', 'FilmCount':10}]), 105)

    def test_inspection_quantity_does_not_consume_equipment(self):
        self.check_stock(fixture([{'MaterialID':'2', '품목명':'Yoke', '수량':1}], daily=[
            {'MaterialID':'2', '장비명':'Yoke', '검사방법':'MT', 'Usage':30}]), 1)

    def test_chemical_consumption_and_empty_daily_data(self):
        material = [{'MaterialID':'3', '품목명':'세척제', '수량':20}]
        self.check_stock(fixture(material), 20)
        self.check_stock(fixture(material, daily=[
            {'MaterialID':'', '검사방법':'MT', 'NDT_세척제':3}]), 17)


if __name__ == '__main__':
    unittest.main()
