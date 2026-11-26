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
    def test_delete_post_sends_the_request_id_at_top_level(self):
        with (
            patch.object(sync, "load_sync_config", return_value={}),
            patch("requests.post") as post,
        ):
            post.return_value.json.return_value = {"status": "success"}
            result = sync.push_delete_to_web_app(
                "", web_app_url="https://mock.invalid", client_request_id="job-second"
            )
        self.assertEqual(result["status"], "success")
        payload = post.call_args.kwargs["json"]
        self.assertEqual(payload["action"], "delete")
        self.assertEqual(payload["client_request_id"], "job-second")
        self.assertEqual(payload["entry_time"], "")

    def test_mobile_delete_removes_only_matching_id_at_shared_timestamp(self):
        first = dict(Date="2026-11-14", Site="A", User="worker", Usage=10,
                     EntryTime="2026-11-14 10:00:00", 검사방법="MT", ClientRequestID="first")
        second = dict(first, ClientRequestID="second")
        with tempfile.TemporaryDirectory() as directory:
            inventory = Path(directory) / "stock.xlsx"
            pd.DataFrame([first, second]).to_excel(inventory, sheet_name="DailyUsage", index=False)
            with (
                patch.object(sync, "fetch_records_from_web_app",
                             side_effect=[(["ClientRequestID"], [first, second]),
                                          ([], [second]), (["ClientRequestID"], [second])]),
                patch.object(sync, "push_batch_to_web_app", return_value={"status": "success"}),
            ):
                self.assertEqual(sync.two_way_reconcile_with_web_app(inventory, "https://mock.invalid")["status"], "success")
                self.assertEqual(sync.two_way_reconcile_with_web_app(inventory, "https://mock.invalid")["status"], "error")
                self.assertEqual(len(pd.read_excel(inventory)), 2)
                result = sync.two_way_reconcile_with_web_app(inventory, "https://mock.invalid")
                self.assertEqual(result["status"], "success")
                self.assertEqual(result["deleted_count"], 1)
            self.assertEqual(pd.read_excel(inventory)["ClientRequestID"].tolist(), ["second"])

    def test_identical_same_second_jobs_with_distinct_ids_survive_pull(self):
        first = dict(Date="2026-11-14", Site="A", User="worker", Usage=10,
                     EntryTime="2026-11-14 10:00:00", 검사방법="MT", ClientRequestID="first")
        second = dict(first, ClientRequestID="second")
        with tempfile.TemporaryDirectory() as directory:
            inventory = Path(directory) / "stock.xlsx"
            pd.DataFrame([first]).to_excel(inventory, sheet_name="DailyUsage", index=False)
            with (
                patch.object(sync, "fetch_records_from_web_app", return_value=(["ClientRequestID"], [first, second])),
                patch.object(sync, "push_batch_to_web_app", return_value={"status": "success"}),
            ):
                for expected in (1, 0):
                    result = sync.two_way_reconcile_with_web_app(inventory, "https://mock.invalid")
                    self.assertEqual(result["status"], "success")
                    self.assertEqual(result["added_count"], expected)
            self.assertCountEqual(pd.read_excel(inventory)["ClientRequestID"], ["first", "second"])

    def test_pending_delete_targets_id_even_when_timestamp_is_shared(self):
        entry = "2026-11-14 10:00:00"
        first = {"EntryTime": entry, "ClientRequestID": "first"}
        second = {"EntryTime": entry, "ClientRequestID": "second"}
        with tempfile.TemporaryDirectory() as directory:
            inventory = Path(directory) / "stock.xlsx"
            with sync.queue_inventory_deletions(inventory, [sync.inventory_deletion_key(first)]):
                pd.DataFrame([second]).to_excel(inventory, sheet_name="DailyUsage", index=False)
            with (
                patch.object(sync, "fetch_records_from_web_app", return_value=(["ClientRequestID"], [first, second])),
                patch.object(sync, "push_delete_to_web_app", return_value={"status": "success"}) as push,
            ):
                sync._flush_pending_deletes(inventory, "https://mock.invalid")
            push.assert_called_once_with("", web_app_url="https://mock.invalid", client_request_id="first")
            self.assertEqual(pd.read_excel(inventory)["ClientRequestID"].tolist(), ["second"])

    def test_incomplete_cloud_response_does_not_clear_pending_delete(self):
        entry = "2026-11-14 10:00:00"
        with tempfile.TemporaryDirectory() as directory:
            inventory = Path(directory) / "stock.xlsx"
            with sync.queue_inventory_deletions(inventory, [entry]):
                pd.DataFrame({"EntryTime": []}).to_excel(inventory, sheet_name="DailyUsage", index=False)
            pending = inventory.with_name(".mobile_pending_deletes.json")
            with patch.object(sync.urllib.request, "urlopen") as open_url:
                response = open_url.return_value.__enter__.return_value
                response.read.return_value = b'{"status":"success"}'
                with self.assertRaises(ValueError):
                    sync._flush_pending_deletes(inventory, "https://mock.invalid")
                self.assertEqual(sync._load_pending_deletes(pending), {entry})
                response.read.return_value = b'{"status":"success","headers":["EntryTime"],"rows":[]}'
                sync._flush_pending_deletes(inventory, "https://mock.invalid")
                self.assertEqual(sync._load_pending_deletes(pending), set())

    def test_pending_delete_with_stale_inventory_copy_blocks_sync(self):
        entry = "2026-11-14 10:00:00"
        with tempfile.TemporaryDirectory() as directory:
            inventory = Path(directory) / "stock.xlsx"
            alternate = Path(directory) / "other" / "stock.xlsx"
            alternate.parent.mkdir()
            with sync.queue_inventory_deletions(inventory, [entry]):
                pd.DataFrame({"EntryTime": []}).to_excel(inventory, sheet_name="DailyUsage", index=False)
            pd.DataFrame({"EntryTime": [entry]}).to_excel(alternate, sheet_name="DailyUsage", index=False)
            for url in ("", "https://mock.invalid"):
                with (
                    self.subTest(url_configured=bool(url)),
                    patch.object(sync, "load_sync_config", return_value={"web_app_url": url}),
                    patch.object(sync, "get_all_inventory_paths", return_value=[inventory, alternate]),
                    patch.object(sync, "push_delete_to_web_app") as push,
                    patch.object(sync, "sync_local_and_drive_files") as merge,
                ):
                    result = sync.sync_all_central_sources()
                    self.assertEqual(result["status"], "error")
                    self.assertIn("다른 재고 사본", result["message"])
                    push.assert_not_called()
                    merge.assert_not_called()

    def test_pending_pc_delete_survives_failure_and_retries(self):
        entry = "2026-11-14 10:00:00"
        with tempfile.TemporaryDirectory() as directory:
            inventory = Path(directory) / "stock.xlsx"
            with sync.queue_inventory_deletions(inventory, [entry]):
                pd.DataFrame({"EntryTime": []}).to_excel(inventory, sheet_name="DailyUsage", index=False)
                self.assertEqual(sync.sync_all_central_sources()["status"], "busy")
            pending = inventory.with_name(".mobile_pending_deletes.json")
            with (
                patch.object(sync, "fetch_records_from_web_app", return_value=([], [{"EntryTime": entry}])),
                patch.object(sync, "push_delete_to_web_app",
                             side_effect=[{"status": "error"}, {"status": "success"}]) as push,
            ):
                with self.assertRaises(RuntimeError):
                    sync._flush_pending_deletes(inventory, "https://mock.invalid")
                self.assertEqual(sync._load_pending_deletes(pending), {entry})
                sync._flush_pending_deletes(inventory, "https://mock.invalid")
                self.assertEqual(sync._load_pending_deletes(pending), set())
                self.assertEqual(push.call_count, 2)

    def test_startup_setting_controls_only_immediate_sync(self):
        from site_apps.central import main

        for enabled in (False, True):
            with (
                self.subTest(enabled=enabled),
                patch.object(main, "is_development", return_value=False),
                patch.object(sync, "load_sync_config", return_value={"auto_sync_on_startup": enabled}),
                patch.object(sync, "sync_all_central_sources", return_value={"status": "success"}) as run_sync,
            ):
                root = Mock()

                def setup_ui(app, window):
                    app.root = window
                    app.daily_save_status = Mock()

                with patch.object(main.MaterialManager, "__init__", setup_ui):
                    app = main.CentralMaterialManager(root)

                self.assertEqual(run_sync.call_count, int(enabled))
                root.after.assert_called_once_with(15000, app._start_periodic_sync)

    def setUp(self):
        self._dev_patcher = patch.object(sync, 'is_development', return_value=False)
        self._dev_patcher.start()

    def tearDown(self):
        self._dev_patcher.stop()

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

    def test_overlapping_sync_is_busy_and_lock_is_released(self):
        def first_sync():
            self.assertEqual(sync.sync_all_central_sources()['status'], 'busy')
            raise OSError('simulated failure')

        with patch.object(sync, '_sync_all_central_sources', side_effect=first_sync):
            with self.assertRaises(OSError):
                sync.sync_all_central_sources()
        with patch.object(sync, '_sync_all_central_sources', return_value={'status': 'success'}):
            self.assertEqual(sync.sync_all_central_sources()['status'], 'success')

    def test_update_delete_refresh_and_conflict_notice_keep_current_input(self):
        source=Path(__file__).resolve().parents[1]/'src/site_apps/central/main.py'
        functions = [n for n in ast.walk(ast.parse(source.read_text(encoding='utf-8')))
                     if isinstance(n, ast.FunctionDef)
                     and n.name in {'_apply_sync_result', '_on_periodic_sync_error'}]
        ns={}
        exec(compile(ast.Module(body=functions,type_ignores=[]),str(source),'exec'),ns)
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
        timestamp = '2026-11-26 09:30:00'
        ns['_apply_sync_result'](app, {'last_success_at': timestamp})
        self.assertIn(timestamp, app.daily_save_status.config.call_args.kwargs['text'])
        app._on_periodic_sync_error = Mock()
        app.load_data.reset_mock()
        ns['_apply_sync_result'](app, {'status': 'error', 'added_count': 1, 'message': 'offline'})
        app.load_data.assert_called_once()
        app._on_periodic_sync_error.assert_called_once_with('offline')
        self.assertEqual(app.current_input, 'unsaved')
        ns['_on_periodic_sync_error'](app, 'offline')
        status = app.daily_save_status.config.call_args.kwargs
        self.assertEqual(status['foreground'], 'red')
        self.assertIn(timestamp, status['text'])


if __name__=='__main__': unittest.main()
