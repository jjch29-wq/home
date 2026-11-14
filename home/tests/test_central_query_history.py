"""Inquiry edits and deletions preserve other stock records and roll back failures."""
import ast
import json
import sys
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock
import unittest
import pandas as pd
sys.path.insert(0,str(Path(__file__).resolve().parents[1] / "src"))

SOURCE = Path(__file__).resolve().parents[1] / 'src/site_apps/central/src/app.py'
TREE = ast.parse(SOURCE.read_text(encoding='utf-8'))


class QueryHistoryTests(unittest.TestCase):
    def test_entry_companies_use_catalog_but_inquiry_keeps_history(self):
        class Combo(dict):
            def __init__(self, value=''):
                super().__init__()
                self.value=value
            def get(self): return self.value
            def set(self,value): self.value=value
        fn=next(n for n in ast.walk(TREE) if isinstance(n,ast.FunctionDef)
                and n.name=='refresh_inquiry_filters')
        ns={'pd':pd,'MARKER_PATTERN':__import__('re').compile(r'\[.*?\]')}
        exec(compile(ast.Module(body=[fn],type_ignores=[]),str(SOURCE),'exec'),ns)
        app=SimpleNamespace(companies=['한국지역난방공사'],
            daily_usage_df=pd.DataFrame({'업체명':['한국지역난방공사 중앙지사']}),
            cb_daily_company=Combo('한국지역난방공사 중앙지사'),
            cb_daily_filter_company=Combo())
        ns['refresh_inquiry_filters'](app)
        self.assertEqual(app.cb_daily_company['values'],['한국지역난방공사'])
        self.assertEqual(app.cb_daily_filter_company['values'],
                         ['전체','한국지역난방공사','한국지역난방공사 중앙지사'])
        self.assertEqual(app.cb_daily_company.get(),'한국지역난방공사 중앙지사')

    def setUp(self):
        self.messages = Mock()
        ns = {'pd':pd,'messagebox':self.messages}
        for name in ('save_daily_usage_edits','delete_daily_usage_entry'):
            fn = next(n for n in ast.walk(TREE) if isinstance(n,ast.FunctionDef) and n.name == name)
            exec(compile(ast.Module(body=[fn],type_ignores=[]),str(SOURCE),'exec'),ns)
        self.methods = ns
        tv = Mock()
        tv.selection.return_value = ['row']
        tv.item.return_value = ('7',)
        self.app = SimpleNamespace(root=Mock(), daily_usage_tree=tv,
            daily_usage_df=pd.DataFrame([{'Date':'2026-10-23','Site':'A','MaterialID':'1','Usage':10},
                                        {'Date':'2026-10-23','Site':'A','MaterialID':'2','Usage':20}],index=[7,8]),
            transactions_df=pd.DataFrame([{'Date':'2026-10-23','Site':'A','MaterialID':'1','Type':'OUT','Note':'A 현장 사용 수동 출고'},
                                         {'Date':'2026-10-23','Site':'A','MaterialID':'2','Type':'OUT','Note':'A 현장 사용 수동 출고'}]),
            materials_df=pd.DataFrame(), monthly_usage_df=pd.DataFrame(), save_data=Mock(return_value=True),
            get_material_display_name=lambda _: '', update_daily_usage_view=Mock(),
            update_recent_entries_view=Mock(), update_stock_view=Mock(), update_transaction_view=Mock(),
            refresh_inquiry_filters=Mock(), update_monthly_usage_view=Mock(), update_material_combo=Mock())
        self.app._safe_set_daily_df = lambda idx,key,value,**kw: self.app.daily_usage_df.loc.__setitem__((idx,key),value)
        self.data = {'Date':'2026-10-23','Site':'A','MaterialID':'1','검사방법':'PAUT','검사량':15}

    def test_edit_keeps_other_transactions_and_refreshes(self):
        original = self.app.transactions_df.copy(deep=True)
        self.assertTrue(self.methods['save_daily_usage_edits'](self.app,7,self.data))
        pd.testing.assert_frame_equal(original,self.app.transactions_df)
        self.assertEqual(self.app.daily_usage_df.loc[7,'Usage'],15)
        self.assertEqual(self.app.daily_usage_df.loc[8,'Usage'],20)
        self.app.update_recent_entries_view.assert_called_once()
        self.app.save_data.assert_called_once_with(require_primary=True)

    def test_edit_failure_rolls_back_all_frames(self):
        original = self.app.daily_usage_df.copy(deep=True)
        self.app.save_data.return_value = False
        self.assertFalse(self.methods['save_daily_usage_edits'](self.app,7,self.data))
        pd.testing.assert_frame_equal(original,self.app.daily_usage_df)
        self.app.update_stock_view.assert_not_called()

    def test_zero_chemical_edit_updates_legacy_and_restores_stock(self):
        from test_central_stock_calculation import fixture
        from site_apps.central.src.models.stock_calculation import build_stock_lookups, stock_breakdown
        reference=fixture([{'MaterialID':'3','품목명':'세척제','수량':20}])
        for name in ('materials_df','transactions_df','normalize_id','clean_nan','_is_consumable_material'):
            setattr(self.app,name,getattr(reference,name))
        self.app.daily_usage_df['NDT_세척제']=[3,0]
        self.app.daily_usage_df['ndt_data']=[json.dumps({'세척제':3,'현상제':2}),'']
        self.assertTrue(self.methods['save_daily_usage_edits'](self.app,7,dict(self.data,**{'NDT_세척제':0})))
        legacy=json.loads(self.app.daily_usage_df.loc[7,'ndt_data'])
        self.assertEqual(legacy,{'세척제':0,'현상제':2})
        stock=stock_breakdown(self.app,self.app.materials_df.iloc[0],build_stock_lookups(self.app))
        self.assertEqual(stock['current_stock'],20)

    def test_failed_zero_edit_restores_direct_and_json_values(self):
        self.app.daily_usage_df['NDT_세척제']=[3,0]
        self.app.daily_usage_df['ndtdata']=[json.dumps({'세척제':3}),'']
        before=self.app.daily_usage_df.copy(deep=True)
        self.app.save_data.return_value=False
        self.assertFalse(self.methods['save_daily_usage_edits'](self.app,7,dict(self.data,**{'NDT_세척제':0})))
        pd.testing.assert_frame_equal(self.app.daily_usage_df,before)

    def test_delete_preserves_other_record_ids_and_transactions(self):
        original = self.app.transactions_df.copy(deep=True)
        self.methods['delete_daily_usage_entry'](self.app)
        self.assertEqual(list(self.app.daily_usage_df.index),[8])
        pd.testing.assert_frame_equal(original,self.app.transactions_df)

    def test_delete_save_failure_restores_records(self):
        original = self.app.daily_usage_df.copy(deep=True)
        self.app.save_data.return_value = False
        self.methods['delete_daily_usage_entry'](self.app)
        pd.testing.assert_frame_equal(original,self.app.daily_usage_df)
        self.app.update_stock_view.assert_not_called()

    def test_refresh_failure_after_success_does_not_undo_saved_edit(self):
        self.app.update_daily_usage_view.side_effect = RuntimeError('refresh failed')
        self.assertTrue(self.methods['save_daily_usage_edits'](self.app,7,self.data))
        self.assertEqual(self.app.daily_usage_df.loc[7,'Usage'],15)
        self.messages.showwarning.assert_called_once()

if __name__ == '__main__': unittest.main()
