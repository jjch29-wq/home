# -*- coding: utf-8 -*-
"""한국지역난방 중앙지사 NDT 성적서 탭 탑재 및 격리 검증 테스트."""
import sys
import unittest
from pathlib import Path
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


if __name__ == '__main__':
    unittest.main()
