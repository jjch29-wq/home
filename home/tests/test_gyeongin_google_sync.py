import unittest
import tempfile
import pandas as pd
from pathlib import Path
from home.src.site_apps.gyeongin.src.sync_google_sheet import (
    load_sync_config,
    save_sync_config,
    sync_records_to_excel
)

class TestGyeonginGoogleSync(unittest.TestCase):
    def test_sync_records_to_excel(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            test_excel_path = Path(tmp_dir) / "Material_Inventory.xlsx"
            
            initial_daily_df = pd.DataFrame([{
                "Date": "2026-11-08",
                "Site": "경인사무소",
                "User": "부장 주진철",
                "검사방법": "MT",
                "수량": 5.0,
                "검사품명": "Tank",
                "단위": "P,M,I/D",
                "단가": 4200.0,
                "EntryTime": "2026-11-08 10:00:00"
            }])
            
            with pd.ExcelWriter(test_excel_path, engine="openpyxl") as writer:
                initial_daily_df.to_excel(writer, sheet_name="DailyUsage", index=False)

            # Test duplicate prevention (same EntryTime or composite key)
            duplicate_rows = [{
                "Date": "2026-11-08",
                "Site": "경인사무소",
                "User": "부장 주진철",
                "검사방법": "MT",
                "수량": 5.0,
                "검사품명": "Tank",
                "단위": "P,M,I/D",
                "단가": 4200.0,
                "EntryTime": "2026-11-08 10:00:00"
            }]
            
            res1 = sync_records_to_excel(duplicate_rows, inventory_path=test_excel_path)
            self.assertEqual(res1["status"], "success")
            self.assertEqual(res1["added_count"], 0)
            self.assertEqual(res1["skipped_count"], 1)

            # Test adding new row
            new_rows = [{
                "Date": "2026-11-08",
                "Site": "경인사무소",
                "User": "대리 우명광",
                "검사방법": "RT",
                "수량": 3.0,
                "검사품명": "Piping",
                "단위": "매",
                "단가": 25500.0,
                "EntryTime": "2026-11-08 14:00:00"
            }]
            
            res2 = sync_records_to_excel(new_rows, inventory_path=test_excel_path)
            self.assertEqual(res2["status"], "success")
            self.assertEqual(res2["added_count"], 1)
            
            # Verify result in Excel
            read_df = pd.read_excel(test_excel_path, sheet_name="DailyUsage")
            self.assertEqual(len(read_df), 2)
            self.assertEqual(read_df.iloc[1]["User"], "대리 우명광")
            self.assertEqual(read_df.iloc[1]["검사방법"], "RT")

if __name__ == "__main__":
    unittest.main()
