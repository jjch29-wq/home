# -*- coding: utf-8 -*-
"""한국지역난방 중앙지사 NDT 성적서 탭 탑재 및 격리 검증 테스트."""
import sys
import unittest
from pathlib import Path
from types import SimpleNamespace
import tkinter as tk
from tkinter import ttk
from unittest.mock import Mock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.central.src.views.ndt_report_tab import NDTReportTab
from site_apps.central.src.reports.ndt_report_app import PMIReportApp


class TestCentralNDTReportTab(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        try:
            cls.root = tk.Tk()
            cls.root.withdraw()
            cls.has_tk = True
        except Exception:
            cls.has_tk = False

    @classmethod
    def tearDownClass(cls):
        if getattr(cls, 'has_tk', False) and cls.root:
            try:
                cls.root.destroy()
            except Exception:
                pass

    def test_ndt_report_tab_initialization(self):
        """NDTReportTab 프레임 생성 및 PMIReportApp 임베딩 정상 초기화 검증."""
        if not self.has_tk:
            self.skipTest("Tkinter display not available in current environment")

        notebook = ttk.Notebook(self.root)
        tab = NDTReportTab(notebook, main_app=None)
        notebook.add(tab, text="비파괴검사 성적서/보고서")

        self.assertIsNotNone(tab.report_app)
        self.assertIsInstance(tab.report_app, PMIReportApp)
        self.assertTrue(tab.report_app.is_embedded)
        self.assertIsNotNone(tab.report_app.mode_notebook)

        # 모드 탭들 정상 생성 확인
        tabs = tab.report_app.mode_notebook.tabs()
        self.assertGreaterEqual(len(tabs), 4)

    def test_report_generation_can_remain_unissued(self):
        app = SimpleNamespace(
            config={}, root=object(), log=Mock(), populate_preview=Mock(),
            paut_extracted_data=[], mt_extracted_data=[], rt_extracted_data=[],
            kogas_extracted_data=[],
        )
        workbook = SimpleNamespace(
            worksheets=[{"AA5": SimpleNamespace(value="SIS/PAUT-012")}]
        )
        with patch(
            "site_apps.central.src.reports.ndt_report_app.messagebox.askyesno",
            return_value=False,
        ), patch(
            "site_apps.central.src.reports.report_issue_history.record_issue"
        ) as record_issue:
            issued = PMIReportApp._record_report_issue(
                app, "PAUT", workbook, "report.xlsx", [{"_source_id": "row-1"}]
            )

        self.assertFalse(issued)
        record_issue.assert_not_called()
        app.log.assert_called_once()

    def test_issue_history_change_refreshes_source_linked_numbers(self):
        sourced = {"_source_id": "row-1", "Report No": "OLD"}
        manual = {"Report No": "MANUAL"}
        data = [sourced, manual]
        app = SimpleNamespace(
            root=object(),
            _get_mode_info=Mock(return_value=(None, None, data, None)),
            populate_preview=Mock(),
        )

        def invoke_change(_parent, on_change=None, **_kwargs):
            on_change()

        with patch(
            "site_apps.central.src.reports.report_issue_history.show_history",
            side_effect=invoke_change,
        ):
            PMIReportApp.show_report_issue_history(app, "PAUT")

        self.assertNotIn("Report No", sourced)
        self.assertEqual(manual["Report No"], "MANUAL")
        app.populate_preview.assert_called_once_with(data, switch_tab=False, mode="PAUT")

    def test_sidebar_toggle_restores_width_and_preserves_widgets(self):
        if not self.has_tk:
            self.skipTest("Tkinter display not available in current environment")
        window = tk.Toplevel(self.root)
        self.addCleanup(window.destroy)
        window.geometry("850x400")
        paned = tk.PanedWindow(window, orient="horizontal")
        paned.pack(fill="both", expand=True)
        sidebar = tk.Frame(paned)
        preview = tk.Frame(paned)
        paned.add(sidebar, width=300)
        paned.add(preview)
        entry = ttk.Entry(sidebar)
        entry.pack()
        entry.insert(0, "작성 중")
        PMIReportApp._add_sidebar_toggle(SimpleNamespace(is_embedded=True), paned, sidebar)
        window.update_idletasks()
        paned.sash_place(0, 280, 0)
        window.update_idletasks()
        width = paned.sash_coord(0)[0]
        button = window.pack_slaves()[0].winfo_children()[0]
        button.invoke()
        window.update_idletasks()
        self.assertFalse(sidebar.winfo_ismapped())
        self.assertTrue(preview.winfo_ismapped())
        self.assertTrue(button.winfo_ismapped())
        button.invoke()
        window.update_idletasks()
        self.assertTrue(sidebar.winfo_ismapped())
        self.assertEqual(paned.sash_coord(0)[0], width)
        self.assertEqual(entry.get(), "작성 중")

    def test_pipe_size_material_mapping_and_gapji_dropdown(self):
        """관경별 KS 재질 매핑(20A~600A -> SPPS380, 650A~900A -> SPW400), 갑지 드롭다운 및 수동 변경 보존 검증."""
        from site_apps.central.src.reports.pipe_material_spec import (
            KDHC_PIPE_MATERIALS,
            infer_kdhc_pipe_material,
            material_for_size_a,
        )

        # 1. 규격표 경계값(20A~600A vs 650A~900A) 검증
        for sz in (20, 25, 32, 40, 50, 65, 80, 100, 125, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600):
            self.assertEqual(material_for_size_a(sz), 'SPPS380')
        for sz in (650, 700, 750, 800, 850, 900):
            self.assertEqual(material_for_size_a(sz), 'SPW400')

        self.assertEqual(
            infer_kdhc_pipe_material([{'Line No.': 'JA2026-013-011-0-2-700A(S)', 'selected': True}], mode='PAUT'),
            'SPW400',
        )
        self.assertEqual(
            infer_kdhc_pipe_material([{'Size': '400A', 'selected': True}, {'Size': '750A', 'selected': True}], mode='PAUT'),
            'SPPS380 / SPW400',
        )
        self.assertEqual(
            infer_kdhc_pipe_material([{'Size': '700A', 'selected': True}], mode='MT', item_name='HOT TAPPING'),
            'SPPS250',
        )

        if not self.has_tk:
            self.skipTest("Tkinter display not available in current environment")

        notebook = ttk.Notebook(self.root)
        tab = NDTReportTab(notebook, main_app=None)
        app = tab.report_app
        app.save_settings = lambda: None  # 테스트 중 설정 파일 쓰기 차단

        # PAUT 서브탭 생성 및 갑지 재질 Combobox 확인
        app._ensure_subtab_built(4)
        app.current_mode = 'PAUT'
        app._load_report_info_for_mode('PAUT')

        combos = []
        def collect_combos(w):
            for ch in w.winfo_children():
                if isinstance(ch, ttk.Combobox):
                    combos.append(ch)
                collect_combos(ch)
        collect_combos(app.paut_tab_notebook)
        self.assertTrue(
            any(tuple(cb.cget('values')) == KDHC_PIPE_MATERIALS and str(cb.cget('state')) != 'readonly' for cb in combos),
            "PAUT 갑지 재질 칸이 편집 가능한 KDHC 표준 재질 드롭다운이어야 합니다.",
        )

        # 2. 700A 행 반영 시 갑지 재질이 SPW400으로 자동 세팅됨
        app.gapji_material.set('SPPS380')
        app.paut_extracted_data = [
            {'Date': '2026-11-20', 'Line No.': 'JA2026-013-011-0-2-700A(S)', 'Joint No.': 'S20', 'Size': '700A', 'selected': True, 'date_filtered': True, 'line_filtered': True, 'size_filtered': True},
            {'Date': '2026-11-20', 'Line No.': 'JA2026-013-011-0-2-400A(S)', 'Joint No.': 'S01', 'Size': '400A', 'selected': False, 'date_filtered': True, 'line_filtered': True, 'size_filtered': True},
        ]
        app.populate_preview(app.paut_extracted_data, switch_tab=False, mode='PAUT')
        self.assertEqual(app.gapji_material.get(), 'SPW400')

        # 3. 동일 관경 상태에서 사용자가 수동으로 재질을 바꾸면 재렌더링 시 유지됨
        app.gapji_material.set('A106 Gr.B')
        app._on_gapji_material_selected('PAUT')
        app.populate_preview(app.paut_extracted_data, switch_tab=False, mode='PAUT')
        self.assertEqual(app.gapji_material.get(), 'A106 Gr.B')

        # 4. 선택 행 관경이 400A로 바뀌면 SPPS380으로 자동 갱신됨
        app.paut_extracted_data[0]['selected'] = False
        app.paut_extracted_data[1]['selected'] = True
        app.populate_preview(app.paut_extracted_data, switch_tab=False, mode='PAUT')
        self.assertEqual(app.gapji_material.get(), 'SPPS380')

        # 5. MT 모드에서 PIPE 품명이어도 SPW400 선택이 강제로 SPPS380으로 덮어써지지 않음
        app._ensure_subtab_built(3)
        app.current_mode = 'MT'
        app.gapji_item.set('PIPE')
        app.gapji_material.set('SPW400')
        app._on_gapji_material_selected('MT')
        self.assertEqual(app.gapji_material.get(), 'SPW400')
        self.assertEqual(app.report_info_by_mode['MT']['material'], 'SPW400')


    def test_gapji_material_selection_restores_combobox_cursor(self):
        callbacks = []
        root = SimpleNamespace(after_idle=lambda callback: callbacks.append(callback))
        widget = Mock()
        widget.winfo_exists.return_value = True
        app = SimpleNamespace(root=root)

        PMIReportApp._restore_gapji_input_focus(app, widget)

        self.assertEqual(len(callbacks), 1)
        callbacks[0]()
        widget.focus_set.assert_called_once_with()
        widget.icursor.assert_called_once_with(tk.END)
        widget.selection_clear.assert_called_once_with()

    def test_paut_print_settings_clear_template_scale_for_one_page_width(self):
        import openpyxl
        import tempfile

        app = object.__new__(PMIReportApp)
        app.current_mode = 'PAUT'
        app.config = {}
        app.log = lambda _message: None
        workbook = openpyxl.Workbook()
        cover = workbook.active
        cover.title = '갑지'
        data = workbook.create_sheet('을지')

        for sheet, context, expected_end in (
            (cover, 'COVER', '$AJ$51'),
            (data, 'DATA', '$AH$47'),
        ):
            sheet.page_setup.scale = 125
            app.force_print_settings(sheet, context=context)
            self.assertTrue(sheet.sheet_properties.pageSetUpPr.fitToPage)
            self.assertEqual(sheet.page_setup.fitToWidth, 1)
            self.assertEqual(sheet.page_setup.fitToHeight, 1)
            self.assertIsNone(sheet.page_setup.scale)
            self.assertEqual(sheet.page_setup.orientation, 'portrait')
            self.assertTrue(str(sheet.print_area).endswith(expected_end))

        from site_apps.central.src.reports.paut_report_layout import align_report_heights
        align_report_heights(
            [cover, data],
            measured_widths={'갑지': 600.0, '을지': 580.0},
            measured_heights={'갑지': 800.0, '을지': 760.0},
        )
        self.assertIsNotNone(cover.page_setup.scale)
        self.assertIsNotNone(data.page_setup.scale)
        app.force_print_settings(cover, context='COVER')
        app.force_print_settings(data, context='DATA')

        with tempfile.TemporaryDirectory() as temp_dir:
            saved_path = Path(temp_dir) / 'paut_print_settings.xlsx'
            workbook.save(saved_path)
            reopened = openpyxl.load_workbook(saved_path)
            try:
                for sheet, expected_end in (
                    (reopened['갑지'], '$AJ$51'),
                    (reopened['을지'], '$AH$47'),
                ):
                    self.assertTrue(sheet.sheet_properties.pageSetUpPr.fitToPage)
                    self.assertEqual(sheet.page_setup.fitToWidth, 1)
                    self.assertEqual(sheet.page_setup.fitToHeight, 1)
                    self.assertIsNone(sheet.page_setup.scale)
                    self.assertTrue(str(sheet.print_area).endswith(expected_end))
            finally:
                reopened.close()


if __name__ == '__main__':
    unittest.main()
