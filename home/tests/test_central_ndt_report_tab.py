# -*- coding: utf-8 -*-
"""한국지역난방 중앙지사 NDT 성적서 탭 탑재 및 격리 검증 테스트."""
import sys
import unittest
from pathlib import Path
from types import SimpleNamespace
import tkinter as tk
from tkinter import ttk

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


if __name__ == '__main__':
    unittest.main()
