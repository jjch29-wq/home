"""Calendar previews must remain read-only and reflect current saved records."""
import datetime
import json
from pathlib import Path
import sys
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.central.src.views.calendar_work_summary import CalendarWorkSummary, daily_summary, work_log_summary


class CalendarSummaryTests(unittest.TestCase):
    def test_escape_closes_calendar_and_preview_and_cancels_pending_hover(self):
        preview = CalendarWorkSummary.__new__(CalendarWorkSummary)
        preview.calendar, preview.entry, preview.popup_parent = Mock(), Mock(), Mock()
        preview.hover_popup, preview.hover_timer, preview.hover_widget = Mock(), 'pending', Mock()
        self.assertEqual(preview.close_calendar(), 'break')
        preview.calendar.after_cancel.assert_called_once_with('pending')
        preview.hover_popup.withdraw.assert_called_once()
        preview.popup_parent.withdraw.assert_called_once()
        preview.entry.focus_set.assert_called_once()
        self.assertIsNone(preview.hover_widget)
        preview.popup_parent.winfo_ismapped.return_value = False
        preview.entry.focus_set.reset_mock()
        self.assertIsNone(preview.close_calendar())
        preview.entry.focus_set.assert_not_called()

    def test_hover_bindings_survive_refresh_and_schedule_preview(self):
        preview = CalendarWorkSummary.__new__(CalendarWorkSummary)
        preview.calendar = Mock()
        preview.calendar.after.return_value = 'timer-1'
        preview.calendar._get_day_coords.return_value = (0, 0)
        cell = Mock()
        cell.cget.return_value = '23'
        preview.calendar._calendar = [[cell]]
        preview.bound_hover_cells = set()
        preview.hover_widget = preview.hover_timer = preview.hover_popup = None
        day = datetime.date(2026, 10, 23)
        preview.bind_hover_cells({day: 'RT 18매'})
        preview.bind_hover_cells({day: 'RT 20매'})
        self.assertEqual(cell.bind.call_count, 4)
        self.assertEqual(preview.hover_texts[cell], 'RT 20매')
        preview.queue_hover(SimpleNamespace(widget=cell))
        preview.calendar.after.assert_called_once_with(350, preview.show_hover)
        preview.queue_hover(SimpleNamespace(widget=cell))
        self.assertEqual(preview.calendar.after.call_count, 1)
        preview.hide()
        preview.calendar.after_cancel.assert_called_once_with('timer-1')
        self.assertIsNone(preview.hover_widget)

    def test_hover_popup_is_raised_and_child_unmap_does_not_cancel_it(self):
        preview = CalendarWorkSummary.__new__(CalendarWorkSummary)
        preview.calendar = Mock()
        preview.popup_parent = Mock()
        cell = Mock()
        cell.winfo_rootx.return_value = 700
        cell.winfo_rooty.return_value = 250
        cell.winfo_width.return_value = 40
        cell.winfo_screenwidth.return_value = 1920
        cell.winfo_screenheight.return_value = 1080
        preview.hover_widget, preview.hover_popup = cell, None
        preview.hover_texts = {cell: 'RT 18매'}
        module = 'site_apps.central.src.views.calendar_work_summary'
        with patch(module + '.tk.Toplevel') as top, patch(module + '.ttk.Label') as label:
            top.return_value.winfo_reqwidth.return_value = 400
            top.return_value.winfo_reqheight.return_value = 200
            preview.show_hover()
            top.return_value.attributes.assert_called_with('-topmost', True)
            top.return_value.deiconify.assert_called_once()
            top.return_value.lift.assert_called_once()
            label.return_value.config.assert_called_with(text='RT 18매')
            top.return_value.withdraw.reset_mock()
            preview.hide(SimpleNamespace(widget=Mock()))
            top.return_value.withdraw.assert_not_called()
            preview.hide(SimpleNamespace(widget=preview.popup_parent))
            top.return_value.withdraw.assert_called_once()

    def test_work_log_preview_distinguishes_saved_report_from_export(self):
        data = {'weather': '맑음', 'ndt_results': [
            {'검사방법': 'RT', '관경': '150A', '라인번호': 'L-1', 'RT_OR': '18', 'RT_RE': '0'},
            {'검사방법': ''}], 'personnel_data': {'검사원_인원': '3', '안전_인원': '1'},
            'remarks': '작업 완료'}
        text = work_log_summary(data)
        for value in ('저장됨 (출력 여부는 별도)', '일보 검사 내역 1건', 'RT_OR 18',
                      '150A', 'L-1', '검사원 인원 3명', '날씨: 맑음', '작업 완료'):
            self.assertIn(value, text)
        self.assertIn('저장 이력 없음', work_log_summary(None))
        self.assertNotIn('RT_RE 0', text)

    def test_work_log_calendar_uses_its_own_source_and_keeps_green_markers(self):
        calendar = Mock()
        calendar.get_displayed_month.return_value = (10, 2026)
        calendar.calevent_create.side_effect = range(200)
        preview = CalendarWorkSummary.__new__(CalendarWorkSummary)
        preview.calendar, preview.event_ids = calendar, []
        # No cb_daily_site on this tab: reading the input tab's selection is wrong.
        preview.app = SimpleNamespace()
        preview.site_getter = lambda: '중앙지사'
        preview.frame_getter = lambda: pd.DataFrame([
            {'Date': '2026-10-23', 'Site': '중앙지사', '검사방법': 'RT',
             'User': '김', 'WorkTime': '09:00~18:00'},
            {'Date': '2026-10-23', 'Site': '다른 현장', '검사방법': 'UT'}])
        preview.history_getter = Mock(return_value={'2026-10-23': {'ndt_results': []}})
        preview.refresh()
        call = next(c for c in calendar.calevent_create.call_args_list
                    if c.args[0] == datetime.date(2026, 10, 23))
        self.assertEqual(call.args[2], ['saved_work_log'])
        self.assertIn('작업일보: 저장됨', call.args[1])
        self.assertIn('근무시간: 김: 09:00~18:00', call.args[1])
        self.assertNotIn('UT', call.args[1])
        calendar.calevent_create.reset_mock()
        # Immediately after deletion, the green marker and saved message go away.
        preview.refresh(history={})
        call = next(c for c in calendar.calevent_create.call_args_list
                    if c.args[0] == datetime.date(2026, 10, 23))
        self.assertEqual(call.args[2], [])
        self.assertIn('작업일보: 저장 이력 없음', call.args[1])
        self.assertIn('검사 작업 1건', call.args[1])

    def test_work_and_vehicle_summary_preserves_original_participation(self):
        row = {'검사방법': 'RT', '검사품명': 'Pipe', '관경(Inch)': '150A',
               '검사량': 18, 'Unit': '매', 'User': '김', 'User2': '이',
               '작업별인원': json.dumps([{'User3': '박'}]),
               '차량번호': '81두1580', '주행거리': 60, '차량점검': '양호'}
        text = daily_summary(datetime.date(2026, 10, 23), '중앙지사', [row])
        for value in ('검사 작업 1건', 'RT · 150A · Pipe · 18매', '박, 김, 이', '81두1580 · 60km'):
            self.assertIn(value, text)

    def test_vehicle_only_record_is_not_an_inspection(self):
        text = daily_summary(datetime.date(2026, 10, 23), '중앙지사', [
            {'MaterialID': '차량점검', 'User': '차량점검', '검사방법': 'RT',
             '차량번호': '81두1580', '차량점검': '양호', 'DailyMemo': float('nan')}])
        self.assertIn('검사 작업 0건', text)
        self.assertIn('81두1580', text)
        self.assertNotIn('참여자:', text)
        self.assertNotIn('nan', text)

    def test_empty_and_long_day(self):
        day = datetime.date(2026, 10, 23)
        self.assertIn('저장된 작업 없음', daily_summary(day, '중앙지사', []))
        text = daily_summary(day, '', [{'Site': '현장A', '검사방법': 'RT'}] * 20)
        self.assertIn('검사 작업 20건', text)
        self.assertIn('외 14건', text)
        self.assertIn('현장A:', text)

    def test_refresh_filters_current_site_and_replaces_old_previews(self):
        calendar = Mock()
        calendar.get_displayed_month.return_value = (10, 2026)
        calendar.calevent_create.side_effect = range(200)
        site = Mock()
        site.get.return_value = '중앙지사'
        frame = pd.DataFrame([
            {'Date': '2026-10-23', 'Site': '중앙지사', '검사방법': 'RT'},
            {'Date': '2026-10-23', 'Site': '다른 현장', '검사방법': 'UT'},
            {'Date': 'invalid', 'Site': '중앙지사', '검사방법': 'MT'},
        ])
        original = frame.copy(deep=True)
        preview = CalendarWorkSummary.__new__(CalendarWorkSummary)
        preview.calendar, preview.event_ids = calendar, [99]
        preview.app = SimpleNamespace(cb_daily_site=site, daily_usage_df=frame)
        preview.refresh()
        calendar.calevent_remove.assert_called_once_with(99)
        self.assertEqual(calendar.calevent_create.call_count, 49)
        day = datetime.date(2026, 10, 23)
        text = next(c.args[1] for c in calendar.calevent_create.call_args_list if c.args[0] == day)
        self.assertIn('RT', text)
        self.assertNotIn('UT', text)
        marked = next(c for c in calendar.calevent_create.call_args_list if c.args[0] == day)
        self.assertEqual(marked.args[2], ['saved_daily_work'])
        empty_day = next(c for c in calendar.calevent_create.call_args_list
                         if c.args[0] == datetime.date(2026, 10, 22))
        self.assertEqual(empty_day.args[2], [])
        calendar._display_selection.assert_called_once()
        pd.testing.assert_frame_equal(frame, original)
        site.get.return_value = '다른 현장'
        calendar.calevent_create.reset_mock()
        preview.refresh()
        text = next(c.args[1] for c in calendar.calevent_create.call_args_list if c.args[0] == day)
        self.assertIn('UT', text)
        self.assertNotIn('RT', text)
        # Deleting the last saved job removes its marker on the next refresh.
        preview.app.daily_usage_df = frame.iloc[0:0]
        calendar.calevent_create.reset_mock()
        preview.refresh()
        self.assertTrue(all(c.args[2] == [] for c in calendar.calevent_create.call_args_list))
        preview.hide()
        self.assertIsNone(calendar.tooltip_wrapper.current_widget)


if __name__ == '__main__':
    unittest.main()
