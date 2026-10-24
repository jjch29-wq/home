"""Exercise the actual daily aggregation loop, including 250A night work."""
import ast
from pathlib import Path
from types import SimpleNamespace
import unittest


class Field:
    def __init__(self, value): self.value = value
    def get(self): return self.value


class PautNightTotalsTests(unittest.TestCase):
    def test_all_paut_sizes_split_day_and_night_without_changing_subtotal(self):
        path = Path(__file__).resolve().parents[1] / 'src/site_apps/central/src/daily_work_log_tab.py'
        tree = ast.parse(path.read_text(encoding='utf-8-sig'))
        cls = next(n for n in tree.body if isinstance(n, ast.ClassDef) and n.name == 'DailyWorkLogTab')
        method = next(n for n in cls.body if isinstance(n, ast.FunctionDef) and n.name == 'auto_calculate_and_save')
        loop = next(n for n in method.body if isinstance(n, ast.For) and isinstance(n.target, ast.Name) and n.target.id == 'row_entries')
        keys = ['300A이상', '250A', '200A', '150A~125A', '100A이하']
        quantities = {f'PAUT_{key}{suffix}':0.0 for key in keys for suffix in ('', '-야간')}
        rows = []
        for size in ('300A', '250A', '200A', '150A', '100A'):
            for shift in ('주간', '야간'):
                rows.append({key:Field(value) for key,value in {
                    '검사방법':'PAUT', '관경':size, '근무구분':shift, 'PAUT':'1.6802'}.items()})
        namespace = {'self':SimpleNamespace(ndt_grid_entries=rows),
                     'today_qty':quantities, 'today_rejects':{}}
        exec(compile(ast.Module(body=[loop],type_ignores=[]),str(path),'exec'),namespace)
        for key in quantities:
            self.assertAlmostEqual(quantities[key],1.6802,key)
        self.assertAlmostEqual(sum(quantities.values()),16.802)


if __name__ == '__main__':
    unittest.main()
