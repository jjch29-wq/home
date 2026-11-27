"""Monthly OT identity, normalized material filtering, and empty popup clearing."""
import ast
import datetime
from pathlib import Path
import re
import sys
from types import SimpleNamespace
import unittest
from unittest.mock import MagicMock, Mock
import pandas as pd
sys.path.insert(0,str(Path(__file__).resolve().parents[1] / "src"))

SOURCE = Path(__file__).resolve().parents[1] / 'src/site_apps/central/src/views/monthly_usage_view.py'
TEXT = SOURCE.read_text(encoding='utf-8')
TREE = ast.parse(TEXT)
sys.path.insert(0,str(SOURCE.parents[3]))


def compile_function(fn, namespace):
    ast.fix_missing_locations(fn)
    exec(compile(ast.Module(body=[fn],type_ignores=[]),str(SOURCE),'exec'),namespace)
    return namespace[fn.name]


class MonthlyAggregationTests(unittest.TestCase):
    def test_distinct_workers_are_counted_even_with_same_timestamp(self):
        fn = next(n for n in ast.walk(TREE) if isinstance(n,ast.FunctionDef) and n.name=='sync_dedup_and_calc_ot')
        app = SimpleNamespace(clean_nan=lambda x: '' if pd.isna(x) else str(x).strip(),
            _safe_format_datetime=lambda x,f:pd.to_datetime(x).strftime(f),
            _calculate_ot_from_worktime=lambda *args:(2,20000))
        ns = {'self':app,'pd':pd,'datetime':datetime,'_re':re,
              'matched_pairs':[('User','WorkTime','OT')],'seen_m_times':set(),'seen_m_contents':set()}
        calc = compile_function(fn,ns)
        base = {'Date':'2026-10-23','Site':'A','WorkTime':'09:00~20:00','MaterialID':'',
                'OT':'20000','EntryTime':'2026-10-23 10:00:00'}
        first = calc(pd.Series(dict(base,User='A')))
        second = calc(pd.Series(dict(base,User='B')))
        repeated = calc(pd.Series(dict(base,User='A')))
        self.assertEqual(first.iloc[1] + second.iloc[1],16000)
        self.assertEqual(repeated.iloc[1],0)
        ns['seen_m_times'].clear()
        ns['seen_m_contents'].clear()
        calc(pd.Series(dict(base,User='A')))
        changed_shift = calc(pd.Series(dict(base,User='A',WorkTime='20:00~23:00')))
        self.assertEqual(changed_shift.iloc[1],13000)

    def prefix(self, app):
        tree = ast.parse(TEXT)
        fn = next(n for n in tree.body if isinstance(n,ast.FunctionDef) and n.name=='update_monthly_usage_view_impl')
        stop = TEXT[:TEXT.index('    # Populate site filter options')].count('\n') + 1
        fn.body = [n for n in fn.body if n.lineno < stop]
        fn.body.append(ast.Return(value=ast.Name(id='df',ctx=ast.Load())))
        return compile_function(fn,{'pd':pd})(app)

    def fixture(self):
        def field(value): return SimpleNamespace(get=lambda:value)
        def tv():
            obj=Mock(); obj.get_children.return_value=['old']; return obj
        return SimpleNamespace(monthly_usage_tree=tv(),site_summary_tree=tv(),worker_summary_tree=tv(),
            detached_windows={'monthly':{'tree':tv(),'site_tree':tv(),'worker_tree':tv()}},
            cb_filter_year=field('전체'),cb_filter_month=field('전체'),cb_filter_site_monthly=field('전체'),
            cb_filter_method_monthly=field('전체'),cb_filter_material_monthly=field('film'),
            normalize_site_name=lambda v:v, normalize_id=lambda v: re.sub(r'\.0$', '',str(v).strip()),
            materials_df=pd.DataFrame([{'MaterialID':10001,'품목명':'film'}]),
            daily_usage_df=pd.DataFrame([{'Date':'2026-10-23','Site':'A','MaterialID':'10001.0'}]))

    def test_material_filter_matches_equivalent_ids(self):
        result=self.prefix(self.fixture())
        self.assertEqual(len(result),1)

    def test_empty_result_clears_all_detached_tables(self):
        app=self.fixture()
        app.cb_filter_material_monthly=SimpleNamespace(get=lambda:'missing')
        self.assertIsNone(self.prefix(app))
        for tree in app.detached_windows['monthly'].values():
            tree.delete.assert_called_once_with('old')

    def test_material_filter_matches_display_name_partial_keyword_and_ndt_labels(self):
        # 1) Display label with [자동등록] spec (e.g., Carestream T200-3⅓*6" [자동등록]) matches 품목명
        app = self.fixture()
        app.materials_df = pd.DataFrame([
            {'MaterialID': 481.0, '품목명': 'Carestream T200-3⅓*6"', '모델명': 'Carestream T200-3⅓*6"', '규격': '자동등록'},
            {'MaterialID': 406.0, '품목명': 'Carestream T200-3⅓*12"', '모델명': '', '규격': ''},
        ])
        app.daily_usage_df = pd.DataFrame([
            {'Date': '2026-09-29', 'Site': '중앙지사', '검사방법': 'RT', 'MaterialID': 481.0},
            {'Date': '2026-10-06', 'Site': '중앙지사', '검사방법': 'RT', 'MaterialID': 406.0},
            {'Date': '2026-10-07', 'Site': '중앙지사', '검사방법': 'MT', 'MaterialID': ''},
        ])
        app.cb_filter_material_monthly = SimpleNamespace(get=lambda: 'Carestream T200-3⅓*6" [자동등록]')
        res_disp = self.prefix(app)
        self.assertEqual(len(res_disp), 1)
        self.assertEqual(res_disp.iloc[0]['MaterialID'], 481.0)

        # 2) Partial keyword search (e.g., 'T200') matches both T200 films
        app._monthly_view_dirty = True
        app.cb_filter_material_monthly = SimpleNamespace(get=lambda: 'T200')
        res_partial = self.prefix(app)
        self.assertEqual(len(res_partial), 2)

        # 3) Virtual NDT label ('MT약품') matches MT rows without MaterialID
        app._monthly_view_dirty = True
        app.cb_filter_material_monthly = SimpleNamespace(get=lambda: 'MT약품')
        res_mt = self.prefix(app)
        self.assertEqual(len(res_mt), 1)
        self.assertEqual(res_mt.iloc[0]['검사방법'], 'MT')

    def test_worker_summary_counts_unique_worked_dates_not_clock_hours(self):
        import tkinter as tk
        tree = ast.parse(TEXT)
        fn = next(n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == '_populate_monthly_summary_trees_impl')
        start = TEXT[:TEXT.index('    # --- Populate Worker Summary ---')].count('\n') + 1
        fn.body = [n for n in fn.body if n.lineno >= start]

        inserted_rows = []
        worker_tree = MagicMock()
        worker_tree.insert.side_effect = lambda parent, idx, values=(), **kw: inserted_rows.append((values, kw.get('tags', ())))

        app = SimpleNamespace(
            clean_nan=lambda x: '' if pd.isna(x) else str(x).strip(),
            _safe_format_datetime=lambda x, f: '' if (pd.isna(x) or not str(x).strip()) else pd.to_datetime(x).strftime(f),
            worker_summary_tree=worker_tree,
        )
        df = pd.DataFrame([
            # Worker A: 2026-10-20 has day shift + empty secondary material row + night shift -> 1.0 day, OT preserved
            {'Date': '2026-10-20', 'Site': '현장1', 'MaterialID': '481', 'User': '부장 주진철', 'WorkTime': '(주간) 09:00~18:00', 'OT': ''},
            {'Date': '2026-10-20', 'Site': '현장1', 'MaterialID': '406', 'User': '부장 주진철', 'WorkTime': '', 'OT': ''},
            {'Date': '2026-10-20', 'Site': '현장2', 'MaterialID': '', 'User': '부장 주진철', 'WorkTime': '(야간) 20:00~23:00', 'OT': '3'},
            # Worker A: 2026-10-21 night shift -> +1.0 day (total 2.0 days, NOT 9 + 1 + 20 + 20 = 50.0)
            {'Date': '2026-10-21', 'Site': '현장1', 'MaterialID': '', 'User': '부장 주진철', 'WorkTime': '(야간) 20:00~05:00', 'OT': '9'},
            # Worker B: 2026-10-22 explicit 0.5 day + secondary empty row -> 0.5 day
            {'Date': '2026-10-22', 'Site': '현장1', 'MaterialID': '481', 'User': '계장 장승대', 'WorkTime': '0.5주간', 'OT': ''},
            {'Date': '2026-10-22', 'Site': '현장1', 'MaterialID': '406', 'User': '계장 장승대', 'WorkTime': '', 'OT': ''},
        ])

        def is_active(val):
            if val is None:
                return False
            s = str(val).strip().lower()
            if s in ('', '0', '0.0', '0.00', 'nan', 'none', '-'):
                return False
            try:
                return abs(float(s.replace(',', ''))) > 0.001
            except ValueError:
                return bool(s)

        ns = {
            'pd': pd,
            're': re,
            'tk': tk,
            'df': df,
            'detached': None,
            'is_active': is_active,
        }
        # Replace `self` parameter call with `(self)`
        fn.args.args = [ast.arg(arg='self')]
        run_worker_summary = compile_function(fn, ns)
        run_worker_summary(app)

        rows_by_name = {vals[0]: vals for vals, _ in inserted_rows}
        self.assertEqual(rows_by_name['계장 장승대'][1], '0.5')
        self.assertEqual(rows_by_name['부장 주진철'][1], '2.0')
        # OT hours across both shifts (day 4.0 + night 8.0 = total 12.0) must still sum accurately
        self.assertEqual(rows_by_name['부장 주진철'][2], '4.0')
        self.assertEqual(rows_by_name['부장 주진철'][3], '8.0')
        self.assertEqual(rows_by_name['부장 주진철'][5], '12.0')
        self.assertEqual(rows_by_name['--- 전체 누계 ---'][1], '2.5')


if __name__ == '__main__': unittest.main()
