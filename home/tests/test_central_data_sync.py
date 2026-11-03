import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch


sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.central.src import data_sync


class CentralDataSyncGuardTest(unittest.TestCase):
    def test_missing_share_reports_error_and_uses_read_fallback(self):
        with patch.object(data_sync, '_configured_data_dir', return_value=r'F:\missing\central'), \
             patch.object(data_sync.os.path, 'isdir', return_value=False):
            self.assertIn(r'F:\missing\central', data_sync.get_data_connection_error())
            self.assertEqual(
                data_sync.get_data_dir(),
                str(Path(data_sync.__file__).resolve().parent),
            )
            with self.assertRaises(data_sync.SharedDataUnavailableError):
                data_sync.require_shared_data_available()

    def test_available_share_is_selected_without_error(self):
        with tempfile.TemporaryDirectory() as shared_dir, \
             patch.object(data_sync, '_configured_data_dir', return_value=shared_dir):
            self.assertIsNone(data_sync.get_data_connection_error())
            self.assertEqual(data_sync.get_data_dir(), shared_dir)


if __name__ == '__main__':
    unittest.main()
