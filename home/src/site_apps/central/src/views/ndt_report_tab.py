# -*- coding: utf-8 -*-
"""비파괴검사 성적서/보고서 전용 탭 모듈 (한국지역난방 중앙지사 전용).

중앙지사 자재 관리 앱 내에서 비파괴검사 성적서(RT, MT, PT, PAUT)를 직접 생성하고,
중앙지사 일보 데이터와 연동하여 원클릭으로 성적서를 출력합니다.
"""
import tkinter as tk
from tkinter import ttk, messagebox
import logging

from site_apps.central.src.reports.ndt_report_app import PMIReportApp

logger = logging.getLogger(__name__)


class NDTReportTab(ttk.Frame):
    """비파괴검사 성적서/보고서 통합 탭 컨테이너."""

    def __init__(self, parent, main_app=None, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.main_app = main_app
        self.report_app = None

        self._setup_view()

    def _setup_view(self):
        """보고서 앱 인스턴스 생성 및 프레임 마운트."""
        try:
            self.report_app = PMIReportApp(self, central_app=self.main_app, is_embedded=True)
            logger.info("NDTReportTab: PMIReportApp initialized successfully.")
        except Exception as e:
            logger.error(f"Failed to initialize PMIReportApp inside NDTReportTab: {e}", exc_info=True)
            err_label = ttk.Label(
                self,
                text=f"비파괴검사 보고서 모듈을 불러오는 중 오류가 발생했습니다:\n{e}",
                foreground="red",
                font=("Malgun Gothic", 11)
            )
            err_label.pack(expand=True, padx=20, pady=20)

    def on_closing(self):
        """앱 종료 시 보고서 설정 저장 위임."""
        if self.report_app and hasattr(self.report_app, 'on_closing'):
            try:
                self.report_app.on_closing()
            except Exception as e:
                logger.warning(f"Error during NDTReportTab closing: {e}")

