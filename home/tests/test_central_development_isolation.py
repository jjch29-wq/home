"""Development must not discover or contact production storage."""
from pathlib import Path
import sys
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.central.src import data_sync, sync_google_sheet as sync
from site_apps.central.mobile import server
from site_apps.central import run_mobile


class DevelopmentIsolationTests(unittest.TestCase):
    def test_shared_drive_config_is_ignored(self):
        with patch.object(data_sync, 'is_development', return_value=True):
            self.assertIsNone(data_sync._configured_data_dir())
            site = Path(data_sync.__file__).resolve().parents[1]
            self.assertEqual(Path(data_sync.get_business_data_dir()), site / 'data')
            self.assertEqual(Path(data_sync.get_history_path()), site / 'src/daily_work_history.json')
            self.assertIsNone(data_sync.get_data_connection_error())

    def test_inventory_discovery_stays_local(self):
        with patch.object(sync, 'is_development', return_value=True), patch.object(server, 'is_development', return_value=True):
            self.assertEqual(sync.get_all_inventory_paths(), [sync.DEFAULT_INVENTORY_PATH])
            self.assertEqual(server.get_inventory_paths(), [str(sync.DEFAULT_INVENTORY_PATH)])

    def test_explicit_cloud_urls_cannot_bypass_isolation(self):
        with patch.object(sync, 'is_development', return_value=True), \
             patch('urllib.request.urlopen', side_effect=AssertionError('Network forbidden')), \
             patch('requests.post', side_effect=AssertionError('Network forbidden')):
            url = 'https://production.invalid/exec'
            self.assertEqual(sync.load_sync_config()['web_app_url'], '')
            with self.assertRaises(RuntimeError):
                sync.fetch_records_from_web_app(url)
            for result in (sync.push_record_to_web_app({}, web_app_url=url),
                           sync.push_delete_to_web_app('entry', web_app_url=url),
                           sync.push_batch_to_web_app(saves=[{}], web_app_url=url),
                           sync.two_way_reconcile_with_web_app(web_app_url=url),
                           sync.run_sync(url), sync.sync_local_and_drive_files(),
                           sync.sync_all_central_sources()):
                self.assertTrue(result['disabled'])

    def test_mobile_uses_separate_port_without_tunnel(self):
        with patch('site_apps.central.src.runtime_profile.is_development', return_value=True), \
             patch.object(run_mobile, 'start_tunnel') as tunnel, patch('uvicorn.run') as run:
            run_mobile.main()
            tunnel.assert_not_called()
            run.assert_called_once_with('site_apps.central.mobile.server:app', host='127.0.0.1', port=8001, reload=False)


if __name__ == '__main__':
    unittest.main()
