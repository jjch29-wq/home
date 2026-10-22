import ast
from decimal import Decimal
from pathlib import Path
from types import SimpleNamespace
import sys
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.central.src.services.billing_totals import grouped_billing_totals, grouped_supply_amount


def record(qty, price, method='PAUT', material='300A 이상', work_time='일반'):
    return dict(qty=qty, unit_price=price, ndt_type=method, material_type=material,
                work_time=work_time, loc='[열배관] 중앙지사',
                subtotal=int(Decimal(str(qty)) * price))


class BillingTotalsTest(unittest.TestCase):
    def test_quantity_is_summed_before_truncation(self):
        rows = [record(qty, 201863) for qty in
                (10.2136, 10.2136, 4.0024, 4.4686, 17.8744, 17.8744, 3.0018, 11.0066)]
        self.assertEqual(sum(r['subtotal'] for r in rows), 15877612)
        self.assertEqual(grouped_supply_amount(rows, {}), 15877615)
        mt = [record(qty, 75433, 'MT', 'MT') for qty in (.8784, 18.6816, 6.2)]
        self.assertEqual(sum(r['subtotal'] for r in mt), 1943153)
        self.assertEqual(grouped_supply_amount(mt, {}), 1943154)

    def test_different_shift_and_material_remain_separate(self):
        rows = [record(1, 201863), record(1, 284015, work_time='야간'),
                record(6, 138793, 'RT', '3 1/3 x 6"')]
        self.assertEqual(len(grouped_billing_totals(rows, {})), 3)
        self.assertEqual(grouped_supply_amount(rows[:1], {}), 201863)
        self.assertEqual(grouped_supply_amount([], {}), 0)

    def test_ui_uses_grouped_amount_and_preserves_four_decimal_quantity(self):
        source = Path(__file__).resolve().parents[1] / 'src/site_apps/central/src/ndt_billing_tab.py'
        tree = ast.parse(source.read_text(encoding='utf-8'))
        method = next(n for n in ast.walk(tree) if isinstance(n, ast.FunctionDef)
                      and n.name == 'update_qty_summary')
        namespace = dict(grouped_billing_totals=grouped_billing_totals,
                         grouped_supply_amount=grouped_supply_amount)
        exec(compile(ast.Module(body=[method], type_ignores=[]), str(source), 'exec'), namespace)
        class Var:
            def set(self, value):
                self.value = value
        quantity, subtotal = Var(), Var()
        rows = [record(10.2136, 201863), record(10.2136, 201863)]
        app = SimpleNamespace(records=rows,
                              contract_vars={'열배관_일반_PAUT_300A 이상': {'curr_qty': quantity, 'c_price': 201863}},
                              subtotal_var=subtotal, export_billing_data=lambda: None)
        namespace['update_qty_summary'](app)
        self.assertEqual(quantity.value, '20.4272')
        self.assertIn(f'{grouped_supply_amount(rows, app.contract_vars):,}', subtotal.value)
        self.assertEqual(rows[0]['subtotal'], 2061747)


if __name__ == '__main__':
    unittest.main()
