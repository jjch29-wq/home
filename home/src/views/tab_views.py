"""View objects that own the application's notebook tabs."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Callable

from tkinter import ttk


@dataclass
class BaseTabView:
    app: object
    notebook: ttk.Notebook
    title: str
    frame_attribute: str
    setup_method: str

    def build(self) -> ttk.Frame:
        frame = ttk.Frame(self.notebook)
        setattr(self.app, self.frame_attribute, frame)
        self.notebook.add(frame, text=self.title)
        setup: Callable[[], None] = getattr(self.app, self.setup_method)
        setup()
        return frame

    @classmethod
    def attach(cls, app, notebook, frame):
        """Attach a view object to a frame created by legacy bootstrap code."""
        view = cls(app, notebook)
        setattr(app, view.frame_attribute, frame)
        getattr(app, view.setup_method)()
        return view


class StockTabView(BaseTabView):
    def __init__(self, app, notebook):
        super().__init__(app, notebook, "현재 재고 현황", "tab_stock", "setup_stock_tab")


class InOutTabView(BaseTabView):
    def __init__(self, app, notebook):
        super().__init__(app, notebook, "입출고 관리", "tab_inout", "setup_inout_tab")


class ImportExportTabView(BaseTabView):
    def __init__(self, app, notebook):
        super().__init__(app, notebook, "데이터 가져오기/내보내기", "tab_import", "setup_import_tab")


class MonthlyUsageTabView(BaseTabView):
    def __init__(self, app, notebook):
        super().__init__(app, notebook, "월별 집계", "tab_monthly_usage", "setup_monthly_usage_tab")


class DailyUsageTabView(BaseTabView):
    def __init__(self, app, notebook):
        super().__init__(app, notebook, "현장별 일일 사용량 기입", "tab_daily_usage", "setup_daily_usage_tab")


class DailyUsageQueryTabView(BaseTabView):
    def __init__(self, app, notebook):
        super().__init__(app, notebook, "현장 일일기록 조회 및 관리", "tab_daily_usage_query", "setup_daily_usage_query_tab")


class BudgetTabView(BaseTabView):
    def __init__(self, app, notebook):
        super().__init__(app, notebook, "공사실행예산서", "tab_budget", "setup_budget_tab")


class NdtBillingTabView(BaseTabView):
    def __init__(self, app, notebook):
        super().__init__(app, notebook, "기성 정산 (NDT)", "tab_ndt_billing", "setup_ndt_billing_tab")


TAB_VIEW_TYPES = (
    StockTabView,
    InOutTabView,
    ImportExportTabView,
    MonthlyUsageTabView,
    DailyUsageTabView,
    DailyUsageQueryTabView,
    BudgetTabView,
    NdtBillingTabView,
)
