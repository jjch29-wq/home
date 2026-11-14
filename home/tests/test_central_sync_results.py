import ast
from pathlib import Path
from types import SimpleNamespace
import sys
import tempfile
import unittest
from unittest.mock import Mock, patch
import pandas as pd

sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'src'))
from site_apps.central.src import sync_google_sheet as sync


class SyncResultTests(unittest.TestCase):
    def test_unchanged_legacy_zero_alias_does_not_zero_actual_usage(self):
        with tempfile.TemporaryDirectory() as directory:
            path=Path(directory)/'stock.xlsx'
            row=dict(Date='2026-11-14',Site='A',User='worker',Usage=10,검사량=10,
                     수량=0,EntryTime='2026-11-14 10:00:00',검사방법='MT')
            pd.DataFrame([row]).to_excel(path,sheet_name='DailyUsage',index=False)
            with patch.object(sync,'fetch_records_from_web_app',return_value=([],[row])), \
                 patch.object(sync,'save_sync_config'), \
                 patch.object(sync,'push_batch_to_web_app',return_value={'status':'success'}):
                result=sync.two_way_reconcile_with_web_app(path,'https://mock.invalid')
            self.assertEqual(result['updated_count'],0)
            self.assertEqual(pd.read_excel(path).iloc[0]['Usage'],10)

    def test_cloud_quantity_updates_all_ledger_aliases(self):
        for method in ('MT','RT'):
            with self.subTest(method=method), tempfile.TemporaryDirectory() as directory:
                path=Path(directory)/'stock.xlsx'
                first=dict(Date='2026-11-14',Site='A',User='worker',Usage=10,
                           검사량=10,수량=10,FilmCount=10 if method=='RT' else 0,
                           EntryTime='2026-11-14 10:00:00',검사방법=method)
                pd.DataFrame([first]).to_excel(path,sheet_name='DailyUsage',index=False)
                with patch.object(sync,'fetch_records_from_web_app',return_value=([],[dict(first,수량=20)])), \
                     patch.object(sync,'save_sync_config'), \
                     patch.object(sync,'push_batch_to_web_app',return_value={'status':'success'}):
                    result=sync.two_way_reconcile_with_web_app(path,'https://mock.invalid')
                    for _ in range(2):
                        repeated=sync.two_way_reconcile_with_web_app(path,'https://mock.invalid')
                        self.assertEqual(repeated['updated_count'],0)
                self.assertEqual(result['updated_count'],1)
                row=pd.read_excel(path).iloc[0]
                for key in ('Usage','검사량','수량'):
                    self.assertEqual(row[key],20)
                self.assertEqual(row['FilmCount'],20 if method=='RT' else 0)

    def test_zero_edit_is_stable_even_with_stale_secondary_aliases(self):
        with tempfile.TemporaryDirectory() as directory:
            path=Path(directory)/'stock.xlsx'
            old=dict(Date='2026-11-14',Site='A',User='worker',Usage=10,검사량=10,
                     수량=10,FilmCount=10,EntryTime='2026-11-14 10:00:00',검사방법='RT')
            pd.DataFrame([old]).to_excel(path,sheet_name='DailyUsage',index=False)
            with patch.object(sync,'fetch_records_from_web_app',return_value=([],[dict(old,수량=0)])), \
                 patch.object(sync,'save_sync_config'), \
                 patch.object(sync,'push_batch_to_web_app',return_value={'status':'success'}):
                for _ in range(3):
                    result=sync.two_way_reconcile_with_web_app(path,'https://mock.invalid')
                    self.assertEqual(result['status'],'success')
                    row=pd.read_excel(path).iloc[0]
                    self.assertEqual([row[k] for k in ('수량','검사량','Usage','FilmCount')],[0,0,0,0])

    def test_summary_preserves_updates_deletes_and_conflicts(self):
        with patch.object(sync,'sync_local_and_drive_files',return_value={'status':'success','added_count':0}), \
             patch.object(sync,'load_sync_config',return_value={'web_app_url':'https://mock.invalid'}), \
             patch.object(sync,'get_all_inventory_paths',return_value=[Path('mock.xlsx')]), \
             patch.object(sync,'run_sync',return_value=dict(status='success',added_count=0,updated_count=2,deleted_count=1,conflict_count=3)):
            result=sync.sync_all_central_sources()
        self.assertEqual([result[k] for k in ('added_count','updated_count','deleted_count','conflict_count')],[0,2,1,3])

    def test_update_delete_refresh_and_conflict_notice_keep_current_input(self):
        source=Path(__file__).resolve().parents[1]/'src/site_apps/central/main.py'
        fn=next(n for n in ast.walk(ast.parse(source.read_text(encoding='utf-8')))
                if isinstance(n,ast.FunctionDef) and n.name=='_apply_sync_result')
        ns={}
        exec(compile(ast.Module(body=[fn],type_ignores=[]),str(source),'exec'),ns)
        for changes in ({'updated_count':1},{'deleted_count':1},{'conflict_count':2}):
            app=SimpleNamespace(load_data=Mock(),update_daily_usage_view=Mock(),
                update_recent_entries_view=Mock(),update_stock_view=Mock(),
                daily_save_status=Mock(),current_input='unsaved')
            ns['_apply_sync_result'](app,changes)
            if changes.get('conflict_count'):
                app.load_data.assert_not_called()
                self.assertIn('2건 보류',app.daily_save_status.config.call_args.kwargs['text'])
            else:
                app.load_data.assert_called_once()
                app.update_stock_view.assert_called_once()
            self.assertEqual(app.current_input,'unsaved')


if __name__=='__main__': unittest.main()
