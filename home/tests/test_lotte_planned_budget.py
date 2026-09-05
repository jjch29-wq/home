"""Exercise the actual widget calculations without starting the desktop app."""
import ast
import json
from pathlib import Path
import sys
import types
import unittest

SRC = Path(__file__).resolve().parents[1] / 'src'
sys.path.insert(0, str(SRC))
from services.lotte_planned_budget import load_planned_budget


class Field:
    def __init__(self, value=''):
        self.value = str(value)

    def get(self):
        return self.value

    def cget(self, key):
        return self.value

    def config(self, **kwargs):
        self.value = kwargs['text']


def widget_class(name):
    tree = ast.parse((SRC / 'views/components.py').read_text(encoding='utf-8-sig'))
    node = next(n for n in tree.body if isinstance(n, ast.ClassDef) and n.name == name)
    node.bases = []
    scope = {}
    exec(compile(ast.Module(body=[node], type_ignores=[]), str(SRC / 'views/components.py'), 'exec'), scope)
    return scope[name]


class PlannedBudgetTests(unittest.TestCase):
    def test_workbook_totals_using_widget_calculations(self):
        baseline = load_planned_budget()
        source = json.loads((SRC / 'services/lotte_planned_budget.json').read_text(encoding='utf-8'))['cells']
        labor = widget_class('LaborCostDetailWidget').__new__(widget_class('LaborCostDetailWidget'))
        labor.ranks = list(baseline['labor'])[:8]
        labor.special_types = list(baseline['labor'])[8:]
        labor.exact_rates = True
        labor.entries = {k: {col: Field(v) for col, v in row.items()} for k, row in baseline['labor'].items()}
        labor.totals = {k: Field() for k in labor.entries}
        for label in ('t1_personnel_sum', 't1_days_sum', 't1_cost_sum',
                      't2_personnel_sum', 't2_hours_sum', 't2_cost_sum', 'grand_total'):
            setattr(labor, 'lbl_' + label, Field())
        labor.on_change_callback = None
        for key in labor.entries:
            labor._on_input_change(key)
        self.assertEqual(labor.raw_total, source['J30'])
        self.assertEqual(labor.lbl_t1_cost_sum.get(), '182,645,000')
        self.assertEqual(labor.lbl_t2_cost_sum.get(), '9,425,000')

        cls = widget_class('ExpenseProfitDetailWidget')
        expense = cls.__new__(cls)
        expense.entries = {k: [{**{col: Field(v) for col, v in row.items()}, 'amount': Field()}
                                for row in rows] for k, rows in baseline['expense'].items()}
        expense.budget_mode = 'planned'
        expense.master_app = None
        expense.get_labor_total = lambda: labor.raw_total
        expense.get_material_total = lambda: sum(float(r['qty']) * float(r['price']) for r in baseline['material'])
        expense.get_revenue = lambda: baseline['revenue']
        for label in ('insurance_base', 'insurance_amount', 'exp_total', 'exp_vat',
                      'sales_cost_total', 'indirect_base', 'indirect_total', 'grand_total_cost',
                      'prof_revenue', 'prof_total_cost', 'prof_op_profit', 'prof_margin'):
            setattr(expense, 'lbl_' + label, Field())
        captured = []
        expense.on_change_callback = lambda *args: captured.append(args)
        expense.calculate_all()
        checks = {'insurance_amount': 'J78', 'exp_total': 'J90', 'exp_vat': 'J91',
                  'sales_cost_total': 'J93', 'indirect_total': 'J98',
                  'grand_total_cost': 'J100', 'prof_op_profit': 'F104'}
        for label, cell in checks.items():
            with self.subTest(cell=cell):
                actual = getattr(expense, 'lbl_' + label).get().replace('₩ ', '')
                self.assertEqual(actual, f"{source[cell]:,.0f}")
        self.assertEqual(expense.lbl_prof_margin.get(), '21.09%')
        self.assertAlmostEqual(captured[-1][0] + captured[-1][1], source['J90'])
        self.assertAlmostEqual(expense.get_total_cost(), source['J90'] - source['J74'])
        self.assertEqual(expense.entries['depreciation'][4]['qty'].get(), '0')

    def test_all_changed_python_files_parse(self):
        for name in ('views/components.py', 'services/lotte_planned_budget.py', '롯데건설 바이오로직스.py'):
            ast.parse((SRC / name).read_text(encoding='utf-8-sig'))


if __name__ == '__main__':
    unittest.main()
