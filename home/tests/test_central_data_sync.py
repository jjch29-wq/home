import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch


sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.central.src import data_sync


class CentralDataSyncGuardTest(unittest.TestCase):
    def test_corrupt_external_inventory_fails_attendance_export(self):
        from site_apps.central.src.utils import export_helper

        with (
            tempfile.TemporaryDirectory() as directory,
            patch.dict(os.environ, {"PMI_CENTRAL_DATA_ROOT": directory}),
            patch.object(data_sync, "is_development", return_value=False),
        ):
            inventory = Path(directory) / "data" / "Material_Inventory.xlsx"
            inventory.parent.mkdir()
            inventory.write_bytes(b"not an Excel file")
            with self.assertRaisesRegex(RuntimeError, "운영 재고의 작업일보를 읽을 수 없습니다"):
                export_helper.load_central_work_log_attendance()

    def test_attendance_error_does_not_fall_back_to_legacy_export(self):
        from site_apps.central.src.utils import export_helper

        with (
            patch.object(export_helper.os.path, "exists", return_value=True),
            patch.object(export_helper, "export_budget_to_standard_template",
                         side_effect=export_helper.AttendanceExportError("work log unavailable")),
            patch.object(export_helper, "_export_budget_legacy_excel") as legacy,
        ):
            with self.assertRaisesRegex(export_helper.AttendanceExportError, "work log unavailable"):
                export_helper.export_budget_estimation_to_excel("unused.xlsx", {}, [], [], [])
            legacy.assert_not_called()

    def test_report_history_resolves_to_external_root(self):
        from site_apps.central.src.reports import report_issue_history

        with (
            tempfile.TemporaryDirectory() as directory,
            patch.dict(os.environ, {"PMI_CENTRAL_DATA_ROOT": directory}),
            patch.object(data_sync, "is_development", return_value=False),
        ):
            self.assertEqual(
                report_issue_history.resolve_history_path(),
                Path(directory) / "report_issue_history.json",
            )

    def test_external_root_is_the_only_mobile_inventory(self):
        from site_apps.central.src import sync_google_sheet as sync

        with (
            tempfile.TemporaryDirectory() as directory,
            patch.dict(os.environ, {"PMI_CENTRAL_DATA_ROOT": directory}),
            patch.object(data_sync, "is_development", return_value=False),
            patch.object(sync, "is_development", return_value=False),
        ):
            inventory = Path(directory) / "data" / "Material_Inventory.xlsx"
            with self.assertRaises(FileNotFoundError):
                sync.get_all_inventory_paths()
            config = {"web_app_url": "https://mock.invalid"}
            self.assertTrue(sync.save_sync_config(config))
            inventory.touch()
            self.assertEqual(sync.get_all_inventory_paths(), [inventory])
            self.assertEqual(sync.load_sync_config(), config)
            self.assertTrue(inventory.with_name("google_sync_config.json").is_file())
            sync._save_sync_state(inventory, {"2026-11-14 10:00:00"})
            self.assertTrue(inventory.with_name(".mobile_sync_state.json").is_file())

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
