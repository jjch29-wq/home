"""Excel report orchestration separated from the main window class."""

from __future__ import annotations

from services.excel_exporter import (
    export_central_daily_work_report_impl,
    export_daily_work_report_impl,
)


def export_daily_work_report(app, *args, **kwargs):
    return export_daily_work_report_impl(app, *args, **kwargs)


def export_central_daily_work_report(app, *args, **kwargs):
    return export_central_daily_work_report_impl(app, *args, **kwargs)

