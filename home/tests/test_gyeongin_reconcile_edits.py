"""Exercise cloud reconciliation against temporary workbooks, without network."""
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.gyeongin.src import sync_google_sheet as sync


class ReconcileEditTests(unittest.TestCase):
    def test_fractional_time_copy_is_not_imported_again(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'inventory.xlsx'
            current = {'Date': '2026-11-09', 'Site': 'Edited', 'Usage': 16.9,
                       '수량': 0, 'EntryTime': '2026-11-11 00:00:56', 'User': 'Worker'}
            stale = dict(current, Site='Original', EntryTime='2026-11-11 00:00:56.304000')
            stale['수량'] = 0.0
            pd.DataFrame([current]).to_excel(path, sheet_name='DailyUsage', index=False)
            result = sync.sync_records_to_excel([stale], path)
            self.assertEqual(result['added_count'], 0)
            self.assertEqual(list(pd.read_excel(path)['Site']), ['Edited'])
            self.assertEqual(sync.record_keys(current), sync.record_keys(stale))

    def test_edit_retry_cloud_edit_and_new_rows(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'inventory.xlsx'
            original = {'Date': '2026-11-09', 'Site': 'Original', 'Usage': 16.9,
                        'EntryTime': '2026-11-09 18:00:00', 'User': 'Worker'}
            pd.DataFrame([original]).to_excel(path, sheet_name='DailyUsage', index=False)
            state = {}
            cloud = [original.copy()]
            def store(value):
                state.clear()
                state.update(value)
                return True
            with patch.object(sync, 'load_sync_config', return_value={'web_app_url': 'test'}), \
                 patch.object(sync, 'save_sync_config'), \
                 patch.object(sync, 'load_sync_state', side_effect=lambda: dict(state)), \
                 patch.object(sync, 'save_sync_state', side_effect=store), \
                 patch.object(sync, 'fetch_records_from_web_app', side_effect=lambda url: ([], cloud)), \
                 patch.object(sync, 'push_batch_to_web_app', return_value={'status': 'success'}) as push:
                self.assertEqual(sync.two_way_reconcile_with_web_app(path)['status'], 'success')
                edited = dict(original, Site='Edited', Usage=20)
                pd.DataFrame([edited]).to_excel(path, sheet_name='DailyUsage', index=False)
                push.return_value = {'status': 'error'}
                self.assertEqual(sync.two_way_reconcile_with_web_app(path)['status'], 'error')
                self.assertEqual(pd.read_excel(path).iloc[0]['Site'], 'Edited')
                push.return_value = {'status': 'success'}
                self.assertEqual(sync.two_way_reconcile_with_web_app(path)['status'], 'success')
                self.assertEqual(push.call_args.kwargs['updates'][0]['Site'], 'Edited')
                self.assertEqual(len(pd.read_excel(path)), 1)
                cloud[:] = [dict(edited, Site='Mobile edit'),
                            dict(original, EntryTime='2026-11-10 18:00:00', Site='New job')]
                self.assertEqual(sync.two_way_reconcile_with_web_app(path)['status'], 'success')
                saved = pd.read_excel(path)
                self.assertEqual(list(saved['Site']), ['Mobile edit', 'New job'])
                self.assertEqual(sync.two_way_reconcile_with_web_app(path)['pulled_count'], 0)
                self.assertEqual(len(pd.read_excel(path)), 2)

    def test_batch_reports_failed_update(self):
        with patch.object(sync, 'load_sync_config', return_value={'web_app_url': 'test'}), \
             patch.object(sync, 'push_record_to_web_app', return_value={'status': 'error'}):
            result = sync.push_batch_to_web_app(updates=[{'EntryTime': '2026-11-09 18:00:00'}])
        self.assertEqual(result['status'], 'error')


if __name__ == '__main__':
    unittest.main()
