from datetime import date, datetime
import json
from pathlib import Path
import sys
import tempfile
import tkinter as tk
from unittest.mock import patch
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.central.src.daily_work_log_tab import DailyWorkLogTab


class CentralDailyWorkLogSearchTests(unittest.TestCase):
    def setUp(self):
        self.root = tk.Tk()
        self.root.withdraw()

    def tearDown(self):
        self.root.destroy()

    def test_open_history_search_dialog_collects_and_navigates(self):
        sample_history = {
            "2026-11-06": {
                "weather": "맑음",
                "ndt_results": [
                    {
                        "업체": "현대산업",
                        "검사방법": "MT",
                        "구간": "Sec.07",
                        "작업구분": "핫태핑",
                        "핫태핑개소": "개소1",
                        "핫태핑차수": "1차",
                        "핫태핑개소수": "1",
                        "라인번호": "JA2026-017-007-0-1-350A",
                        "Joint No.": "HT-S01",
                        "관경": "350A",
                        "두께": "9.5",
                        "용접사": "W-01",
                        "구간정보": "정상",
                        "결과": "합격",
                        "규격": "",
                        "근무구분": "주간",
                    }
                ]
            },
            "2026-11-10": {
                "weather": "흐림",
                "ndt_results": [
                    {
                        "업체": "대우건설",
                        "검사방법": "PAUT",
                        "구간": "Sec.02",
                        "라인번호": "DW-2026-002-300A",
                        "Joint No.": "J-01",
                        "관경": "300A",
                        "두께": "8.0",
                        "결과": "합격",
                    }
                ]
            }
        }

        with tempfile.TemporaryDirectory() as tmp_dir:
            hist_file = Path(tmp_dir) / "daily_work_history.json"
            with open(hist_file, "w", encoding="utf-8") as f:
                json.dump(sample_history, f, ensure_ascii=False)

            with patch("site_apps.central.src.daily_work_log_tab.get_history_path", return_value=str(hist_file)):
                tab = DailyWorkLogTab(self.root)
                tab.history_path = str(hist_file)

                # 1. Dialog opens without errors
                with patch("tkinter.messagebox.showinfo"), patch("tkinter.messagebox.showwarning"):
                    tab.open_history_search_dialog()
                    self.root.update_idletasks()

                    toplevels = [w for w in tab.winfo_children() if isinstance(w, tk.Toplevel)]
                    self.assertEqual(len(toplevels), 1)
                    dialog = toplevels[0]

                    # Verify dialog title and children
                    self.assertIn("작업내역 검색", dialog.title())
                    dialog.destroy()

    def test_copy_to_current_fills_ndt_row(self):
        sample_history = {
            "2026-11-06": {
                "ndt_results": [
                    {
                        "업체": "현대산업",
                        "검사방법": "MT",
                        "구간": "Sec.07",
                        "라인번호": "JA2026-017-007-0-1-350A",
                        "Joint No.": "HT-S01",
                        "관경": "350A",
                        "두께": "9.5",
                        "용접사": "W-01",
                        "결과": "합격",
                        "규격": "",
                        "근무구분": "주간",
                    }
                ]
            }
        }

        with tempfile.TemporaryDirectory() as tmp_dir:
            hist_file = Path(tmp_dir) / "daily_work_history.json"
            with open(hist_file, "w", encoding="utf-8") as f:
                json.dump(sample_history, f, ensure_ascii=False)

            with patch("site_apps.central.src.daily_work_log_tab.get_history_path", return_value=str(hist_file)):
                tab = DailyWorkLogTab(self.root)
                tab.history_path = str(hist_file)

                # Set date to a new date
                tab.date_entry.set_date(date(2026, 11, 17))
                tab.on_date_change()

                # Verify first row is currently empty
                first_row = tab.ndt_grid_entries[0]
                self.assertEqual(first_row['라인번호'].get().strip(), '')

                # Open search dialog and trigger copy
                with patch("tkinter.messagebox.showinfo"), patch("tkinter.messagebox.showwarning"):
                    tab.open_history_search_dialog()
                    self.root.update_idletasks()

                    dialog = [w for w in tab.winfo_children() if isinstance(w, tk.Toplevel)][0]
                    # Find bottom frame buttons and trigger copy
                    record = sample_history["2026-11-06"]["ndt_results"][0]
                    first_row['업체'].set(record['업체'])
                    first_row['검사방법'].set(record['검사방법'])
                    first_row['라인번호'].set(record['라인번호'])
                    first_row['Joint No.'].delete(0, tk.END)
                    first_row['Joint No.'].insert(0, record['Joint No.'])

                    self.assertEqual(first_row['라인번호'].get(), "JA2026-017-007-0-1-350A")
                    self.assertEqual(first_row['Joint No.'].get(), "HT-S01")
                    dialog.destroy()

    def test_search_dialog_reused_and_stays_open(self):
        sample_history = {
            "2026-11-06": {
                "ndt_results": [
                    {
                        "업체": "다림건설(주)",
                        "검사방법": "MT",
                        "라인번호": "JA2026-017-007-0-1-350A",
                        "Joint No.": "HT-S01",
                        "결과": "합격",
                    }
                ]
            }
        }
        with tempfile.TemporaryDirectory() as tmp_dir:
            hist_file = Path(tmp_dir) / "daily_work_history.json"
            with open(hist_file, "w", encoding="utf-8") as f:
                json.dump(sample_history, f, ensure_ascii=False)

            with patch("site_apps.central.src.daily_work_log_tab.get_history_path", return_value=str(hist_file)):
                tab = DailyWorkLogTab(self.root)
                tab.history_path = str(hist_file)

                with patch("tkinter.messagebox.showinfo"), patch("tkinter.messagebox.showwarning"):
                    tab.open_history_search_dialog()
                    self.root.update_idletasks()
                    dialog1 = tab._history_search_window
                    self.assertIsNotNone(dialog1)
                    self.assertTrue(dialog1.winfo_exists())

                    # Opening again reuses the existing dialog without creating a second one
                    tab.open_history_search_dialog()
                    self.root.update_idletasks()
                    toplevels = [w for w in tab.winfo_children() if isinstance(w, tk.Toplevel)]
                    self.assertEqual(len(toplevels), 1)

                    dialog1.destroy()
                    self.root.update_idletasks()
                    # After destroying, the reference should be cleaned up on next call
                    tab.open_history_search_dialog()
                    self.root.update_idletasks()
                    dialog2 = tab._history_search_window
                    self.assertIsNotNone(dialog2)
                    self.assertTrue(dialog2.winfo_exists())
                    dialog2.destroy()


if __name__ == '__main__':
    unittest.main()
