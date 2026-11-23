"""Daily reports distinguish inspection shifts and keep detail text readable."""
import contextlib
import datetime
import io
from pathlib import Path
import sys
import tempfile
import unittest

import openpyxl
import pandas as pd

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / 'home/src'))
from site_apps.central.src.services.report_inspection import (
    collect_inspection_details, collect_method_totals, inspection_method_shift,
)
from site_apps.central.src.daily_work_report_manager import DailyWorkReportManager
from site_apps.central.src.vehicle_records import format_vehicle_numbers


def records():
    return pd.DataFrame([
        {'검사방법': 'PAUT', '작업형태': '일반', '관경(Inch)': '250A', '조인트수': 14,
         'Usage': 11.7614, 'Unit': 'M', '단가': 325451, '검사비': 3827759, '출장비': 0},
        {'검사방법': 'PAUT', '작업형태': '야간', '관경(Inch)': '250A', '조인트수': 8,
         'Usage': 6.7208, 'Unit': 'M', '단가': 229486, '검사비': 1542349, '출장비': 0},
    ])


class DailyReportShiftTests(unittest.TestCase):
    def test_vehicle_numbers_split_legacy_separators_and_remove_duplicates(self):
        self.assertEqual(format_vehicle_numbers([
            '90서8616 (스타렉스) || 95가0175 (스타렉스)',
            '90서8616 (스타렉스), 12가3456 (승합, 디젤)', None]),
            '90서8616 (스타렉스)\n95가0175 (스타렉스)\n12가3456 (승합, 디젤)')
        self.assertEqual(format_vehicle_numbers(float('nan')), '')

    def test_multiple_vehicles_are_centered_with_enough_row_height(self):
        template = ROOT / 'home/src/site_apps/central/resources/Template_DailyWorkReport.xlsx'
        for count in (1, 2, 3):
            with self.subTest(count=count), tempfile.TemporaryDirectory(dir=ROOT / 'tmp') as directory:
                vehicles = ['90서8616 (스타렉스)', '95가0175 (스타렉스)', '12가3456 (승합차)'][:count]
                data = {'date': datetime.date(2026, 10, 24), 'car_no': ' || '.join(vehicles)}
                output = Path(directory) / 'vehicles.xlsx'
                with contextlib.redirect_stdout(io.StringIO()):
                    DailyWorkReportManager(str(template)).generate_report(data, output)
                book = openpyxl.load_workbook(output)
                try:
                    sheet = book.active
                    self.assertEqual(sheet['N9'].value, '\n'.join(vehicles))
                    self.assertEqual(sheet['N9'].alignment.horizontal, 'center')
                    self.assertEqual(sheet['N9'].alignment.vertical, 'center')
                    self.assertTrue(sheet['N9'].alignment.wrap_text)
                    self.assertFalse(sheet['N9'].alignment.shrink_to_fit)
                    self.assertGreaterEqual(sheet.row_dimensions[9].height, count * 15 + 4)
                finally:
                    book.close()

    def test_long_equipment_names_expand_row_height_and_preserve_readable_font(self):
        template = ROOT / 'home/src/site_apps/central/resources/Template_DailyWorkReport.xlsx'
        long_equipment = (
            'PAUT MX2(OMNI2-103085), JIREH Scanner, 5L16-A10 PAUT PROBE(A10) (S1873), '
            'SA10-N55S-IHC PAUT PROBE-WEDGE(SA10), SA2-N55S-IHC PAUT PROBE-WEDGE(SA2), '
            '5L64-A2 PAUT PROBE(A2) (S1599), 5L64-A2 PAUT PROBE(A2) (S1694)'
        )
        with tempfile.TemporaryDirectory(dir=ROOT / 'tmp') as directory:
            data = {'date': datetime.date(2026, 11, 19), 'equipment': long_equipment}
            output = Path(directory) / 'equipment.xlsx'
            with contextlib.redirect_stdout(io.StringIO()):
                DailyWorkReportManager(str(template)).generate_report(data, output)
            book = openpyxl.load_workbook(output)
            try:
                sheet = book.active
                self.assertEqual(sheet['E9'].value, long_equipment)
                self.assertTrue(sheet['E9'].alignment.wrap_text)
                self.assertLessEqual(sheet['E9'].font.size, 8.5)
                # Ensure row 9 height is substantially expanded to fit all equipment lines
                self.assertGreaterEqual(sheet.row_dimensions[9].height, 50)
                # Ensure whole sheet still fits within 1-page printable boundary
                visible_rows = [sheet.row_dimensions[r] for r in range(1, 53) if not sheet.row_dimensions[r].hidden]
                total_height = sum(r.height or 15 for r in visible_rows)
                self.assertLessEqual(total_height, 277 / 25.4 * 72 + 1)
            finally:
                book.close()

    def test_day_night_totals_preserve_saved_quantities_prices_and_amounts(self):
        result = collect_method_totals(records())
        self.assertEqual(list(result), ['PAUT(주간)', 'PAUT(야간)'])
        self.assertAlmostEqual(sum(v['qty'] for v in result.values()), 18.4822)
        self.assertEqual(result['PAUT(주간)']['price'], 325451)
        self.assertEqual(result['PAUT(야간)']['price'], 229486)
        self.assertEqual(sum(v['total'] for v in result.values()), 5370108)

    def test_distinct_rates_and_units_do_not_merge_under_maximum_price(self):
        frame = records()
        frame = pd.concat([frame, pd.DataFrame([
            dict(frame.iloc[0], 단가=100, Usage=1, 검사비=100),
            dict(frame.iloc[0], Unit='POINT', Usage=2, 검사비=200),
        ])], ignore_index=True)
        result = collect_method_totals(frame)
        self.assertEqual(len(result), 4)
        self.assertEqual(sum(v['total'] for v in result.values()), frame['검사비'].sum())

    def test_saved_work_type_takes_precedence_over_worker_hours(self):
        self.assertEqual(inspection_method_shift({'검사방법': 'PAUT', '작업형태': '일반',
            'WorkTime': '(야간) 22:00~익일08:00'}), ('PAUT', '주간'))
        self.assertEqual(inspection_method_shift({'검사방법': 'RT', 'WorkTime': '(야간) 22:00~익일08:00'}),
                         ('RT', '야간'))
        self.assertEqual(inspection_method_shift({'검사방법': 'RT'}), ('RT', '미구분'))

    def test_mixed_worker_shifts_do_not_duplicate_or_split_the_quantity(self):
        frame = pd.DataFrame([{'검사방법': 'UT', 'Usage': 10, '단가': 100, '검사비': 1000,
                               'WorkTime': '(주간) 09:00~18:00', 'WorkTime2': '(야간) 22:00~익일08:00'}])
        result = collect_method_totals(frame)
        self.assertEqual(list(result), ['UT(주야간)'])
        self.assertEqual(result['UT(주야간)']['qty'], 10)

    def test_details_use_same_shift_groups_and_keep_point_counts(self):
        details = collect_inspection_details(records(), {'250A': .8401})
        self.assertEqual([(d['shift'], d['points']) for d in details], [('주간', 14), ('야간', 8)])
        self.assertAlmostEqual(sum(d['quantity'] for d in details), 18.4822)

    def test_native_workbook_has_both_rows_and_ten_point_details(self):
        frame = records()
        data = {'date': datetime.date(2026, 10, 24), 'methods': collect_method_totals(frame),
                'inspection_details': collect_inspection_details(frame, {'250A': .8401})}
        template = ROOT / 'home/src/site_apps/central/resources/Template_DailyWorkReport.xlsx'
        with tempfile.TemporaryDirectory(dir=ROOT / 'tmp') as directory:
            output = Path(directory) / 'daily.xlsx'
            with contextlib.redirect_stdout(io.StringIO()):
                DailyWorkReportManager(str(template)).generate_report(data, output)
            book = openpyxl.load_workbook(output)
            try:
                sheet = book.active
                self.assertEqual(sheet['B13'].value, 'PAUT(주간)')
                self.assertEqual(sheet['B14'].value, 'PAUT(야간)')
                self.assertAlmostEqual(sheet['H13'].value + sheet['H14'].value, 18.4822)
                self.assertEqual(sheet['Q13'].value + sheet['Q14'].value, 5370108)
                self.assertEqual(sheet['B18'].font.sz, 10)
                self.assertTrue(sheet['B18'].alignment.wrap_text)
                self.assertIn('PAUT(주간) 250A: 14 POINT', sheet['B18'].value)
                self.assertIn('PAUT(야간) 250A: 8 POINT', sheet['B18'].value)
                self.assertGreaterEqual(sum(sheet.row_dimensions[r].height for r in range(18, 26)), 53)
            finally:
                book.close()

    def test_pad_breakdown_renders_in_inspection_details(self):
        pad_details = [
            {'work_type': 'PAD(C심)', 'pipe_size': '800A', 'method': 'MT', 'points': 3, 'quantity': 7.6605, 'joints': ['S01', 'S02', 'S03']},
            {'work_type': 'PAD(L심)', 'pipe_size': '1000mm', 'method': 'MT', 'points': 4, 'quantity': 4.0, 'joints': ['S04', 'S05', 'S06', 'S07']},
            {'work_type': 'PAD(CAP)', 'pipe_size': '100A', 'method': 'MT', 'points': 1, 'quantity': 0.3591, 'joints': ['S08']},
        ]
        data = {
            'date': datetime.date(2026, 11, 23),
            'inspection_item': 'PAD',
            'methods': {'MT(주간)': {'unit': 'M', 'qty': 12.0196, 'price': 75433, 'travel': 0, 'total': 906675}},
            'pad_details': pad_details,
            'note': '[PAD 세부내역] PAD(C심) 800A: 3개소(7.6605m) / PAD(L심): 4개소(4.0000m) / PAD(CAP) 100A: 1개소(0.3591m)'
        }
        template = ROOT / 'home/src/site_apps/central/resources/Template_DailyWorkReport.xlsx'
        with tempfile.TemporaryDirectory(dir=ROOT / 'tmp') as directory:
            output = Path(directory) / 'pad_report.xlsx'
            with contextlib.redirect_stdout(io.StringIO()):
                DailyWorkReportManager(str(template)).generate_report(data, output)
            book = openpyxl.load_workbook(output)
            try:
                sheet = book.active
                self.assertEqual(sheet['B13'].value, 'MT(주간)')
                self.assertAlmostEqual(sheet['H13'].value, 12.0196)
                b18_val = sheet['B18'].value
                self.assertIn('PAD(C심) 800A : 3 POINT / 검사길이 7.6605 M', b18_val)
                self.assertIn('PAD(L심) 1000mm : 4 POINT / 검사길이 4.0000 M', b18_val)
                self.assertIn('PAD(CAP) 100A : 1 POINT / 검사길이 0.3591 M', b18_val)
                self.assertIn('PAD 소계', b18_val)
                self.assertIn('총 검사길이 12.0196 M', b18_val)
                self.assertIn('Joint: S01, S02, S03', b18_val)
            finally:
                book.close()


if __name__ == '__main__':
    unittest.main()
