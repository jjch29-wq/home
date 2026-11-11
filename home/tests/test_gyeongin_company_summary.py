from copy import deepcopy
from pathlib import Path
from tempfile import TemporaryDirectory
import sys
import unittest
from types import SimpleNamespace
from unittest.mock import Mock, mock_open, patch
import openpyxl
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.gyeongin.src.company_usage_summary import first_usage_date, summarize_company_usage
from site_apps.gyeongin.src.daily_work_log_exporter import DailyWorkLogExporter
from site_apps.gyeongin.src.daily_work_log_tab import DailyWorkLogTab


class CompanySummaryTests(unittest.TestCase):
    def row(self, day, company='업체A', qty=10, method='RT', unit='매', **extra):
        return {'Date': day, '업체명': company, 'Usage': qty, '검사방법': method, 'Unit': unit, **extra}

    def test_inclusive_period_crosses_year_and_ignores_chemical_manufacturer(self):
        frame = pd.DataFrame([self.row('2025-12-31', qty=10, 회사코드='약품회사'),
                              self.row('2026-01-01', qty=3), self.row('2026-01-02', qty=100),
                              self.row('2025-12-30', qty=100), self.row('2026-01-01', '업체B', 8)])
        before = frame.copy(deep=True)
        rows, skipped = summarize_company_usage(frame, '2025-12-31', '2026-01-01')
        self.assertEqual(skipped, 0)
        self.assertEqual(rows[0], dict(company='업체A', method='RT', unit='매', previous=10, today=3, total=13))
        self.assertEqual(rows[1]['company'], '업체B')
        pd.testing.assert_frame_equal(frame, before)

    def test_units_and_methods_never_mix_and_missing_company_is_explicit(self):
        frame = pd.DataFrame([self.row('2026-11-11', method='MT', unit='M'),
                              self.row('2026-11-11', method='MT', unit='개소'),
                              self.row('2026-11-11', company='', method='PT', unit='M', 회사코드='약품회사')])
        rows, _ = summarize_company_usage(frame, '2026-11-01', '2026-11-11')
        self.assertEqual(len(rows), 3)
        self.assertEqual(rows[0]['company'], '(업체명 미입력)')

    def test_rt_rep_and_zero_quantity_chemical_rows_do_not_inflate_totals(self):
        frame = pd.DataFrame([self.row('2026-11-11', qty=10),
                              self.row('2026-11-11', qty=3, 검사구분='REP'),
                              self.row('2026-11-11', qty=0, 검사량=10, NDT_침투제=2)])
        rows, _ = summarize_company_usage(frame, '2026-11-11', '2026-11-11')
        self.assertEqual(rows[0]['total'], 10)
        self.assertEqual(rows[0]['previous'], 0)

    def test_repeated_refresh_uses_records_instead_of_previous_totals(self):
        frame = pd.DataFrame([self.row('2026-11-10'), self.row('2026-11-11', qty=2)])
        first, _ = summarize_company_usage(frame, '2026-11-10', '2026-11-11')
        self.assertEqual(first, summarize_company_usage(frame, '2026-11-10', '2026-11-11')[0])
        self.assertEqual(summarize_company_usage(frame, '2026-11-11', '2026-11-11')[0][0]['total'], 2)

    def test_invalid_dates_and_quantities_are_reported_and_backwards_period_rejected(self):
        frame = pd.DataFrame([self.row('bad date'), self.row('2026-11-11', qty='bad'),
                              self.row('2026-11-11', qty=-1), self.row('2026-11-11', qty=float('inf'))])
        self.assertEqual(summarize_company_usage(frame, '2026-11-01', '2026-11-11'), ([], 4))
        with self.assertRaises(ValueError):
            summarize_company_usage(frame, '2026-11-12', '2026-11-11')

    def test_default_start_is_first_record_not_start_of_year(self):
        frame = pd.DataFrame([self.row('2025-06-04'), self.row('2026-11-11')])
        self.assertEqual(str(first_usage_date(frame, '2026-11-11')), '2025-06-04')
        self.assertEqual(summarize_company_usage(pd.DataFrame(), '2026-11-11', '2026-11-11'), ([], 0))

    def test_export_preserves_all_groups_and_matches_screen_period(self):
        rows = [dict(company=f'업체{i:02d}', method='RT', unit='매', previous=10, today=2, total=12)
                for i in range(25)]
        data = dict(date='2026-11-11', summary_start_date='2025-06-04', company_summary=rows)
        before = deepcopy(data)
        with TemporaryDirectory() as directory:
            target = Path(directory) / 'daily.xlsx'
            DailyWorkLogExporter().generate_excel(data, target)
            book = openpyxl.load_workbook(target)
            try:
                detail = book['업체별 실적']
                self.assertEqual(detail['B2'].value, '2025-06-04')
                self.assertEqual(detail.max_row, 29)
                self.assertEqual(detail['A29'].value, '업체24')
                self.assertEqual(detail['F29'].value, 12)
                self.assertEqual(book.active['A8'].value, '업체명')
                self.assertIn('25행', book.active['A28'].value)
                self.assertNotIn('예상량', [cell.value for row in book.active for cell in row])
            finally:
                book.close()
        self.assertEqual(data, before)

    def make_tab(self, start='2026-11-01'):
        def date_field(value):
            return SimpleNamespace(get=lambda: value, get_date=lambda: pd.Timestamp(value).date())
        history = {'2026-11-11': {'qty_data': {'old': {'예상량': '129'}},
                                'equip_data': {'RT장비': {'금일': '1', '누계': '3'}},
                                'personnel_data': {'검사원_인원': '2', '검사원_누계': '8'}}}
        tab = SimpleNamespace(
            main_app=SimpleNamespace(daily_usage_df=pd.DataFrame([self.row('2026-11-11', qty=4)])),
            date_entry=date_field('2026-11-11'), company_start_date=date_field(start),
            _summary_start_initialized=True, company_summary_tree=Mock(), company_summary_status=Mock(),
            _validate_rt_rows=lambda: True, _validate_hot_rows=lambda: True,
            load_history=lambda: history, _selected_date_key=lambda: '2026-11-11',
            weather_entry=Mock(), remarks_text=Mock(), qty_entries={}, equip_entries={},
            personnel_entries={}, ndt_grid_entries=[], save_history=Mock(), _capture_log_input=Mock())
        tab.company_summary_tree.get_children.return_value = ()
        tab.weather_entry.get.return_value = ''
        tab.remarks_text.get.return_value = ''
        tab.refresh_company_summary = lambda: DailyWorkLogTab.refresh_company_summary(tab)
        return tab, history

    def test_save_snapshots_current_summary_without_erasing_legacy_history(self):
        tab, history = self.make_tab()
        tab.company_memos = {'업체B': '추가 검사 예정'}
        tab._company_memo_company = '업체A'
        tab.company_memo_text = Mock()
        tab.company_memo_text.get.return_value = '성적서 오전 제출'
        with patch('builtins.open', mock_open()):
            DailyWorkLogTab.save_current_history(tab)
        saved = tab.save_history.call_args.args[0]['2026-11-11']
        self.assertEqual(saved['summary_start_date'], '2026-11-01')
        self.assertEqual(saved['company_summary'][0]['today'], 4)
        self.assertEqual(saved['qty_data'], {'old': {'예상량': '129'}})
        self.assertEqual(saved['company_memos'], {'업체A': '성적서 오전 제출', '업체B': '추가 검사 예정'})
        self.assertEqual(saved['equip_data'], {'RT장비': {'금일': '1', '누계': '3'}})
        self.assertEqual(saved['personnel_data'], {'검사원_인원': '2', '검사원_누계': '8'})

    def test_auto_calculation_saves_without_retired_equipment_and_people_widgets(self):
        tab = SimpleNamespace(_validate_rt_rows=lambda: True, _validate_hot_rows=lambda: True,
                              _refresh_hot_summary=Mock(), refresh_company_summary=lambda: True,
                              save_current_history=Mock())
        DailyWorkLogTab.auto_calculate_and_save(tab, show_completion=False)
        tab.save_current_history.assert_called_once_with()

    def test_invalid_range_does_not_save_or_replace_existing_history(self):
        tab, history = self.make_tab('2026-11-12')
        before = deepcopy(history)
        self.assertFalse(DailyWorkLogTab.save_current_history(tab))
        tab.save_history.assert_not_called()
        self.assertEqual(history, before)


if __name__ == '__main__':
    unittest.main()
