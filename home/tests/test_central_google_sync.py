import unittest
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
        res = sync_all_central_sources()
        self.assertEqual(res["status"], "success")
        self.assertIn("added_count", res)

if __name__ == "__main__":
    unittest.main()

