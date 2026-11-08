"""Monthly OT identity, normalized material filtering, and empty popup clearing."""
import ast
import datetime
from pathlib import Path
import re
import sys
from types import SimpleNamespace
import unittest
from unittest.mock import Mock
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

if __name__ == '__main__': unittest.main()
