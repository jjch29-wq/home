import unittest
from unittest.mock import patch
import tempfile
import pandas as pd
from pathlib import Path
from home.src.site_apps.central.src.sync_google_sheet import (
    load_sync_config,
    save_sync_config,
    sync_records_to_excel,
    sync_all_central_sources,
    get_all_inventory_paths
)

class TestCentralGoogleSync(unittest.TestCase):
    def setUp(self):
        # These tests exercise production reconciliation using mocked cloud APIs.
        # Development network blocking is covered separately by isolation tests.
        guard = patch('home.src.site_apps.central.src.sync_google_sheet.is_development', return_value=False)
        guard.start()
        self.addCleanup(guard.stop)

    def test_sync_records_to_excel(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            test_excel_path = Path(tmp_dir) / "Material_Inventory.xlsx"

            initial_daily_df = pd.DataFrame([{
                "Date": "2026-11-08",
                "Site": "중앙지사",
                "User": "부장 주진철",
                "검사방법": "PAUT",
                "수량": 0.8401,
                "관경(Inch)": "250A",
                "구분": "열수송관",
                "단가": 325451.0,
                "EntryTime": "2026-11-08 10:00:00"
            }])

            with pd.ExcelWriter(test_excel_path, engine="openpyxl") as writer:
                initial_daily_df.to_excel(writer, sheet_name="DailyUsage", index=False)

            # Test duplicate prevention (same EntryTime or composite key)
            duplicate_rows = [{
                "Date": "2026-11-08",
                "Site": "중앙지사",
                "User": "부장 주진철",
                "검사방법": "PAUT",
                "수량": 0.8401,
                "관경(Inch)": "250A",
                "구분": "열수송관",
                "단가": 325451.0,
                "EntryTime": "2026-11-08 10:00:00"
            }]

            res1 = sync_records_to_excel(duplicate_rows, inventory_path=test_excel_path)
            self.assertEqual(res1["status"], "success")
            self.assertEqual(res1["added_count"], 0)
            self.assertEqual(res1["skipped_count"], 1)

            # Test adding new row
            new_rows = [{
                "Date": "2026-11-08",
                "Site": "중앙지사",
                "User": "대리 우명광",
                "검사방법": "RT",
                "수량": 3.0,
                "관경(Inch)": "200A",
                "구분": "열배관",
                "단가": 205412.0,
                "EntryTime": "2026-11-08 14:00:00"
            }]

            res2 = sync_records_to_excel(new_rows, inventory_path=test_excel_path)
            self.assertEqual(res2["status"], "success")
            self.assertEqual(res2["added_count"], 1)

            with pd.ExcelFile(test_excel_path) as ef:
                read_df = ef.parse("DailyUsage")
            self.assertEqual(len(read_df), 2)
            self.assertEqual(read_df.iloc[1]["User"], "대리 우명광")
            self.assertEqual(read_df.iloc[1]["검사방법"], "RT")

    def test_sync_all_central_sources_callable(self):
        from unittest.mock import patch

        with patch(
            "home.src.site_apps.central.src.sync_google_sheet.sync_local_and_drive_files",
            return_value={"status": "success", "added_count": 0, "message": "mock drive"},
        ), patch(
            "home.src.site_apps.central.src.sync_google_sheet.load_sync_config",
            return_value={"web_app_url": ""},
        ):
            res = sync_all_central_sources()
        self.assertEqual(res["status"], "success")
        self.assertIn("added_count", res)

    def test_two_way_reconcile_in_place_update(self):
        from unittest.mock import patch
        from home.src.site_apps.central.src.sync_google_sheet import two_way_reconcile_with_web_app

        with tempfile.TemporaryDirectory() as tmp_dir:
            test_excel = Path(tmp_dir) / "Material_Inventory.xlsx"
            initial_df = pd.DataFrame([{
                "Date": "2026-11-09",
                "Site": "중앙지사",
                "User": "부장 주진철",
                "검사방법": "PAUT",
                "수량": 1.0,
                "단가": 325451.0,
                "관경(Inch)": "250A",
                "구분": "열수송관",
                "Note": "초기 메모",
                "EntryTime": "2026-11-09 10:00:00"
            }])
            with pd.ExcelWriter(test_excel, engine="openpyxl") as writer:
                initial_df.to_excel(writer, sheet_name="DailyUsage", index=False)

            # Cloud has updated quantity and note for the same EntryTime
            cloud_records = [{
                "Date": "2026-11-09",
                "Site": "중앙지사",
                "User": "부장 주진철",
                "검사방법": "PAUT",
                "수량": 2.5,
                "단가": 325451.0,
                "관경(Inch)": "250A",
                "구분": "열수송관",
                "Note": "모바일에서 수정된 메모",
                "EntryTime": "2026-11-09 10:00:00"
            }]

            with patch("home.src.site_apps.central.src.sync_google_sheet.fetch_records_from_web_app", return_value=(["Date"], cloud_records)), \
                 patch("home.src.site_apps.central.src.sync_google_sheet.push_batch_to_web_app", return_value={"status": "success"}):
                res = two_way_reconcile_with_web_app(inventory_path=test_excel, web_app_url="https://mock.url")
                self.assertEqual(res["status"], "success")

            with pd.ExcelFile(test_excel) as ef:
                updated_df = ef.parse("DailyUsage")

            self.assertEqual(len(updated_df), 1)
            self.assertEqual(float(updated_df.iloc[0]["수량"]), 2.5)
            self.assertEqual(str(updated_df.iloc[0]["Note"]), "모바일에서 수정된 메모")

    def test_mobile_delete_is_applied_after_successful_snapshot(self):
        from unittest.mock import patch
        from home.src.site_apps.central.src.sync_google_sheet import two_way_reconcile_with_web_app

        with tempfile.TemporaryDirectory() as tmp_dir:
            test_excel = Path(tmp_dir) / "Material_Inventory.xlsx"
            record = {
                "Date": "2026-11-10",
                "Site": "중앙지사",
                "User": "부장 주진철",
                "검사방법": "PAUT",
                "수량": 1.0,
                "관경(Inch)": "250A",
                "EntryTime": "2026-11-10 10:00:00",
            }
            with pd.ExcelWriter(test_excel, engine="openpyxl") as writer:
                pd.DataFrame([record]).to_excel(writer, sheet_name="DailyUsage", index=False)

            with patch(
                "home.src.site_apps.central.src.sync_google_sheet.fetch_records_from_web_app",
                return_value=(["EntryTime"], [record]),
            ), patch("home.src.site_apps.central.src.sync_google_sheet.push_batch_to_web_app") as push_mock:
                first = two_way_reconcile_with_web_app(test_excel, "https://mock.url")
                self.assertEqual(first["status"], "success")
                push_mock.assert_not_called()

            with patch(
                "home.src.site_apps.central.src.sync_google_sheet.fetch_records_from_web_app",
                return_value=(["EntryTime"], []),
            ), patch("home.src.site_apps.central.src.sync_google_sheet.push_batch_to_web_app") as push_mock:
                second = two_way_reconcile_with_web_app(test_excel, "https://mock.url")
                self.assertEqual(second["status"], "success")
                self.assertEqual(second["deleted_count"], 1)
                push_mock.assert_not_called()

            with pd.ExcelFile(test_excel) as ef:
                deleted_df = ef.parse("DailyUsage")
            self.assertEqual(len(deleted_df), 0)

    def test_push_batch_reports_partial_failure(self):
        from unittest.mock import patch
        from home.src.site_apps.central.src.sync_google_sheet import push_batch_to_web_app

        with patch(
            "home.src.site_apps.central.src.sync_google_sheet.push_record_to_web_app",
            return_value={"status": "error", "message": "mock failure"},
        ):
            result = push_batch_to_web_app(
                saves=[{"EntryTime": "2026-11-10 10:00:00"}],
                web_app_url="https://mock.url",
            )

        self.assertEqual(result["status"], "error")
        self.assertEqual(result["added_count"], 0)
        self.assertIn("mock failure", result["message"])

if __name__ == "__main__":
    unittest.main()

