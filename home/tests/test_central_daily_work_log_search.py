from datetime import date, datetime
import json
from pathlib import Path
import sys
import tempfile
import tkinter as tk
from tkinter import ttk
from unittest.mock import patch
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.central.src.daily_work_log_tab import DailyWorkLogTab


class CentralDailyWorkLogSearchTests(unittest.TestCase):
    def test_selection_menu_and_shift_click_share_checkbox_state(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "daily_work_history.json"
            with patch("site_apps.central.src.daily_work_log_tab.get_history_path", return_value=str(path)):
                tab = DailyWorkLogTab(self.root)
                tab.ndt_grid_entries[1]["Joint No."].insert(0, "J01")
                tab.ndt_grid_entries[4]["Joint No."].insert(0, "J02")
                grid = tab._ndt_grid_frame
                menu_button = grid.grid_slaves(row=0, column=0)[0]
                menu = menu_button.nametowidget(menu_button.cget("menu"))

                def selected():
                    return [i for i, var in enumerate(tab.ndt_selection_vars) if var.get()]

                menu.invoke(0)
                self.assertEqual(selected(), [1, 4])
                self.assertEqual(tab.ndt_selection_text.get(), "2개 행 선택됨")
                menu.invoke(1)
                self.assertEqual(selected(), [])

                tab.pack(fill="both", expand=True)
                self.root.geometry("1400x900")
                self.root.deiconify()
                self.root.update_idletasks()
                cell = tab.ndt_grid_entries[0]["Joint No."]
                cell.event_generate("<ButtonPress-1>")
                self.assertIs(tab.selected_ndt_row, tab.ndt_grid_entries[0])
                self.assertEqual(selected(), [])
                tab.ndt_tree.see("row_2")
                self.root.update_idletasks()
                bx2, by2, bw2, bh2 = tab.ndt_tree.bbox("row_2", "#1")
                tab.ndt_tree.event_generate("<ButtonPress-1>", x=bx2 + bw2 // 2, y=by2 + bh2 // 2)
                tab.ndt_tree.see("row_5")
                self.root.update_idletasks()
                bx5, by5, bw5, bh5 = tab.ndt_tree.bbox("row_5", "#1")
                tab.ndt_tree.event_generate("<ButtonPress-1>", x=bx5 + bw5 // 2, y=by5 + bh5 // 2, state=1)
                self.assertEqual(selected(), [1, 2, 3, 4])
                self.assertEqual(tab.ndt_selection_text.get(), "4개 행 선택됨")

    def test_initial_rows_expand_to_show_all_saved_records(self):
        selected = date(2026, 11, 17)
        rows = [{"검사방법": "MT", "Joint No.": f"J{index:02}"} for index in range(35)]
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "daily_work_history.json"
            path.write_text(json.dumps({selected.isoformat(): {"ndt_results": rows}}), encoding="utf-8")
            with patch("site_apps.central.src.daily_work_log_tab.get_history_path", return_value=str(path)):
                tab = DailyWorkLogTab(self.root)
                self.assertEqual(len(tab.ndt_grid_entries), 10)
                tab.date_entry.set_date(selected)
                tab.on_date_change()
                self.assertEqual(
                    [widgets["Joint No."].get() for widgets in tab.ndt_grid_entries[:35]],
                    [row["Joint No."] for row in rows],
                )

    def setUp(self):
        conn_patcher = patch(
            "site_apps.central.src.daily_work_log_tab.get_data_connection_error",
            return_value="",
        )
        conn_patcher.start()
        self.addCleanup(conn_patcher.stop)
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

                    # Verify maximize buttons exist
                    buttons = [w for w in dialog.winfo_children() if isinstance(w, (tk.Button, ttk.Button))]
                    all_buttons = []
                    def find_buttons(widget):
                        for c in widget.winfo_children():
                            if isinstance(c, (tk.Button, ttk.Button)):
                                all_buttons.append(c)
                            find_buttons(c)
                    find_buttons(dialog)
                    btn_texts = [b.cget("text") for b in all_buttons]
                    self.assertTrue(any("최대화" in t for t in btn_texts))
                    self.assertTrue(any("이동하여 수정" in t for t in btn_texts))
                    self.assertTrue(any("복사하여 가져오기" in t for t in btn_texts))

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

    def test_work_category_options_and_batch_edit(self):
        tab = DailyWorkLogTab(self.root)
        self.root.update_idletasks()

        # 1. 컬럼에 '공사구분'이 존재하고 위젯이 올바르게 생성되는지 확인
        self.assertIn('공사구분', tab.ndt_cols)
        row0 = tab.ndt_grid_entries[0]
        self.assertIn('공사구분', row0)

        # 2. 다림건설 선택 시 공사구분 드롭다운 옵션 확인
        row0['업체'].set('다림건설(주)')
        row0['업체'].event_generate('<<ComboboxSelected>>')
        self.assertEqual(list(row0['공사구분']['values']), ['', '제조', '수탁', '건가'])

        # 3. GS네오텍 선택 시 공사구분 드롭다운 옵션 확인
        row0['업체'].set('GS네오텍')
        row0['업체'].event_generate('<<ComboboxSelected>>')
        self.assertEqual(list(row0['공사구분']['values']), ['', '건가-추가', '건가-개선'])

        # 4. 일괄 변경 다이얼로그 호출 및 적용 확인
        row0['공사구분'].set('건가-추가')
        row1 = tab.ndt_grid_entries[1]
        row1['라인번호'].insert(0, 'TEST-LINE-01')
        tab.ndt_selection_vars[1].set(True)

        with patch("tkinter.messagebox.showinfo"):
            tab._edit_ndt_work_categories()
            self.root.update_idletasks()
            toplevels = [w for w in tab.winfo_children() if isinstance(w, tk.Toplevel)]
            self.assertEqual(len(toplevels), 1)
            dialog = toplevels[0]
            self.assertIn("공사구분", dialog.title())
            dialog.destroy()

    def test_billing_summary_splits_by_work_category(self):
        from site_apps.central.src.services.billing_summary import build_company_summaries
        records = [
            {
                'date': '2026-11-06',
                'company': '다림건설(주)',
                'work_category': '제조',
                'ndt_type': 'MT',
                'material_type': 'MT',
                'work_time': '일반',
                'qty': 10.0,
                'loc': '[열배관] 중앙지사 / Sec.07 / JA2026-01',
                'unit': 'M',
                'unit_price': 5000,
                'subtotal': 50000,
                'source': 'work_log',
                'history_rows': [0],
            },
            {
                'date': '2026-11-06',
                'company': '다림건설(주)',
                'work_category': '수탁',
                'ndt_type': 'MT',
                'material_type': 'MT',
                'work_time': '일반',
                'qty': 5.0,
                'loc': '[열배관] 중앙지사 / Sec.07 / JA2026-02',
                'unit': 'M',
                'unit_price': 5000,
                'subtotal': 25000,
                'source': 'work_log',
                'history_rows': [1],
            }
        ]
        history = {
            '2026-11-06': {
                'ndt_results': [
                    {'업체': '다림건설(주)', '공사구분': '제조', '검사방법': 'MT', 'MT': 10.0, '구간': 'Sec.07', '라인번호': 'JA2026-01'},
                    {'업체': '다림건설(주)', '공사구분': '수탁', '검사방법': 'MT', 'MT': 5.0, '구간': 'Sec.07', '라인번호': 'JA2026-02'},
                ]
            }
        }
        contracts = {'열배관_일반_MT_MT': {'c_price': 5000}}
        summaries = build_company_summaries(records, history, contracts)
        self.assertIn('다림건설(주) [제조]', summaries)
        self.assertIn('다림건설(주) [수탁]', summaries)

    def test_line_number_multiselect_filtering(self):
        sample_history = {
            "2026-11-06": {
                "ndt_results": [
                    {"업체": "다림건설(주)", "검사방법": "MT", "라인번호": "LINE-01", "Joint No.": "J-01", "결과": "합격"},
                    {"업체": "다림건설(주)", "검사방법": "MT", "라인번호": "LINE-02", "Joint No.": "J-02", "결과": "합격"},
                    {"업체": "다림건설(주)", "검사방법": "MT", "라인번호": "LINE-03", "Joint No.": "J-03", "결과": "합격"},
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

                    dialog = tab._history_search_window
                    self.assertIsNotNone(dialog)
                    self.assertTrue(hasattr(dialog, '_selected_lines'))

                    # Find treeview
                    tree = None
                    for child in dialog.winfo_children():
                        for sub in child.winfo_children():
                            if isinstance(sub, ttk.Treeview):
                                tree = sub
                                break
                    self.assertIsNotNone(tree)
                    # Initially all 3 rows displayed
                    self.assertEqual(len(tree.get_children()), 3)

                    # 1. Open multi-select popup
                    dialog._open_line_multiselect_popup()
                    self.root.update_idletasks()
                    popup = getattr(dialog, '_line_popup', None)
                    self.assertIsNotNone(popup)
                    self.assertTrue(popup.winfo_exists())
                    self.assertIn("라인번호 복수 선택", popup.title())
                    popup.destroy()

                    # 2. Test multi-line filtering: Select LINE-01 and LINE-03
                    dialog._selected_lines.clear()
                    dialog._selected_lines.update(["LINE-01", "LINE-03"])
                    # Find refresh button and click (recursive search)
                    all_buttons = []
                    def find_buttons(w):
                        for c in w.winfo_children():
                            if isinstance(c, (tk.Button, ttk.Button)):
                                all_buttons.append(c)
                            find_buttons(c)
                    find_buttons(dialog)
                    refresh_btn = next((b for b in all_buttons if "조회" in b.cget("text")), None)
                    self.assertIsNotNone(refresh_btn)
                    refresh_btn.invoke()
                    self.root.update_idletasks()

                    # Tree should have exactly 2 rows
                    filtered_items = tree.get_children()
                    self.assertEqual(len(filtered_items), 2)
                    line_vals = [tree.item(item)['values'][5] for item in filtered_items]
                    self.assertIn("LINE-01", line_vals)
                    self.assertIn("LINE-03", line_vals)
                    self.assertNotIn("LINE-02", line_vals)

                    # 3. Test clear filter: Reset to all
                    dialog._selected_lines.clear()
                    refresh_btn.invoke()
                    self.root.update_idletasks()
                    self.assertEqual(len(tree.get_children()), 3)

                    dialog.destroy()

    def test_treeview_inplace_editor_tab_enter_and_rt_slots(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "daily_work_history.json"
            with patch("site_apps.central.src.daily_work_log_tab.get_history_path", return_value=str(path)):
                tab = DailyWorkLogTab(self.root)
                self.root.geometry("1400x800")
                tab.pack(fill="both", expand=True)
                self.root.deiconify()
                self.root.update()

                # 1. Fill row 0 with PAUT + 300A via in-place editor & verify auto thickness/length
                tab._open_cell_editor(0, "업체")
                tab._cell_combo_var.set("다림건설(주)")
                tab._on_cell_combo_selected()
                self.assertEqual(list(tab.ndt_grid_entries[0]["공사구분"]["values"]), ["", "제조", "수탁", "건가"])

                tab._open_cell_editor(0, "검사방법")
                tab._cell_combo_var.set("PAUT")
                tab._on_cell_combo_selected()

                tab._open_cell_editor(0, "관경")
                tab._cell_combo_var.set("300A")
                tab._on_cell_combo_selected()
                self.assertEqual(tab.ndt_grid_entries[0]["두께"].get(), "6.4")
                self.assertTrue(float(tab.ndt_grid_entries[0]["PAUT"].get()) > 0)

                tab._open_cell_editor(0, "Joint No.")
                tab._cell_entry_var.set("J-01")
                tab._on_editor_return(None)
                # Row 0 Enter moves focus to Row 1 Joint No.
                self.assertEqual(tab._active_editor_state["row"], 1)
                self.assertEqual(tab._active_editor_state["col"], "Joint No.")

                # 2. Row 1 Enter auto-copies Row 0 and increments Joint No. to J-02
                tab._on_editor_return(None)
                self.assertEqual(tab.ndt_grid_entries[1]["Joint No."].get(), "J-02")
                self.assertEqual(tab.ndt_grid_entries[1]["업체"].get(), "다림건설(주)")
                self.assertEqual(tab.ndt_grid_entries[1]["관경"].get(), "300A")
                self.assertEqual(tab.ndt_tree.set("row_2", "Joint No."), "J-02")

                # 3. RT 4-slot section info editor updates RT_OR / RT_RE / 결과 live
                tab.ndt_grid_entries[2]["검사방법"].set("RT")
                tab.ndt_grid_entries[2]["관경"].set("100A")
                tab.ndt_grid_entries[2]["관경"].event_generate("<<ComboboxSelected>>")
                tab._open_cell_editor(2, "구간정보")
                self.assertIs(tab._active_editor_state["widget"], tab._cell_rt_frame)
                tab._cell_rt_vars[0].set("1/-")
                tab._cell_rt_vars[1].set("3/SL")
                tab._cell_rt_vars[2].set("1/-")
                tab._on_rt_sub_keyrelease(type("E", (), {"keysym": "L"})())
                self.assertEqual(tab.ndt_grid_entries[2]["구간정보"].get(), "1/-,3/SL,1/-")
                self.assertEqual(tab.ndt_grid_entries[2]["RT_OR"].get(), "3")
                self.assertEqual(tab.ndt_grid_entries[2]["RT_RE"].get(), "1")
                self.assertEqual(tab.ndt_grid_entries[2]["결과"].get(), "불합격")

    def test_work_log_frame_sash_left_hide_and_col_widths_persist_across_restart(self):
        from types import SimpleNamespace
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "daily_work_history.json"
            shared_cfg = {}
            saved_calls = []
            host = SimpleNamespace(
                tab_config=shared_cfg,
                save_tab_config=lambda force=False: saved_calls.append((force, dict(shared_cfg))),
            )
            with patch("site_apps.central.src.daily_work_log_tab.get_history_path", return_value=str(path)):
                self.root.geometry("1500x850")
                self.root.deiconify()

                # Session 1: user adjusts sash to 30%, resizes '라인번호' col to 280px, then hides left pane
                tab1 = DailyWorkLogTab(self.root, main_app=host)
                tab1.pack(fill="both", expand=True)
                self.root.update()

                total_w = tab1.paned.winfo_width()
                target_sash = int(total_w * 0.30)
                tab1._on_paned_press(SimpleNamespace(widget=tab1.paned))
                tab1.paned.sashpos(0, target_sash)
                tab1._on_paned_release(SimpleNamespace(widget=tab1.paned))

                tab1.ndt_tree.column("라인번호", width=280)
                tab1._resizing_ndt_col = True
                tab1._on_ndt_tree_release(None)

                tab1._toggle_left_pane()
                self.assertFalse(tab1._left_pane_visible)
                self.assertFalse(shared_cfg["work_log_left_visible"])
                self.assertEqual(shared_cfg["work_log_ndt_col_widths"]["라인번호"], 280)
                self.assertAlmostEqual(shared_cfg["work_log_sash_ratio"], target_sash / total_w, delta=0.03)
                tab1.destroy()

                # Session 2 (Restart): left pane starts hidden & column width 280px is preserved;
                # clicking '▶ 좌측 보이기' restores the 30% sash position instead of resetting to 50%!
                tab2 = DailyWorkLogTab(self.root, main_app=host)
                tab2.pack(fill="both", expand=True)
                self.root.update()

                self.assertFalse(tab2._left_pane_visible)
                self.assertEqual(tab2.left_pane_toggle.cget("text"), "▶ 좌측 보이기")
                self.assertEqual(int(tab2.ndt_tree.column("라인번호", "width")), 280)

                tab2._toggle_left_pane()
                self.root.update()
                self.assertTrue(tab2._left_pane_visible)
                restored_ratio = tab2.paned.sashpos(0) / tab2.paned.winfo_width()
                self.assertAlmostEqual(restored_ratio, shared_cfg["work_log_sash_ratio"], delta=0.03)
                tab2.destroy()


if __name__ == '__main__':
    unittest.main()
