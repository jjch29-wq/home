import sys
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

import openpyxl

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / 'home/src'))
from site_apps.central.src.daily_work_log_exporter import DailyWorkLogExporter


class DailyLogPrintTests(unittest.TestCase):
    def test_many_inspections_continue_on_pages_without_squeezing_height(self):
        rows = [{'검사방법': 'RT', 'Joint No.': f'R{index:02}', '관경': '100A',
                 '구간정보': 'N/1,N/1,N/1', 'RT_OR': '3', 'RT_RE': '0'} for index in range(1, 76)]
        with TemporaryDirectory(dir=ROOT / 'tmp') as directory:
            target = Path(directory) / 'daily.xlsx'
            DailyWorkLogExporter().generate_excel({'date': '2026-10-20', 'ndt_results': rows}, target)
            book = openpyxl.load_workbook(target)
            try:
                sheet = book.active
                self.assertEqual(sheet.page_setup.fitToWidth, 1)
                self.assertEqual(sheet.page_setup.fitToHeight, 0)
                self.assertTrue(sheet.sheet_properties.pageSetUpPr.fitToPage)
                self.assertTrue(sheet.print_options.horizontalCentered)
                self.assertEqual(sheet.print_title_rows, '$30:$31')
                self.assertEqual(sheet['F106'].value, 'R75')
                self.assertIn('$T$107', str(sheet.print_area))
                self.assertEqual(sheet['P106'].value, 3)
            finally:
                book.close()

    def test_line_number_formatting_and_row_height(self):
        rows = [
            {'검사방법': 'PAUT', '라인번호': 'JA2026-017-007-0-1-350A', 'Joint No.': 'S01'},
            {'검사방법': 'PAUT', '라인번호': 'JA2026-017-007-0-1-300A(가배관)', 'Joint No.': 'S02'},
            {'검사방법': 'MT', '라인번호': '송경고등학교 내 주차장 300A(가배관)', 'Joint No.': 'S03'},
            {'검사방법': 'MT', '작업구분': '핫태핑', '라인번호': 'JA2026-017-007-0-1-300A', '핫태핑개소': 'A', '핫태핑차수': '1차', 'Joint No.': 'S04'},
        ]
        with TemporaryDirectory(dir=ROOT / 'tmp') as directory:
            target = Path(directory) / 'daily_line_test.xlsx'
            DailyWorkLogExporter().generate_excel({'date': '2026-10-20', 'ndt_results': rows}, target)
            book = openpyxl.load_workbook(target)
            try:
                sheet = book.active
                # Column E widened to 26
                self.assertEqual(sheet.column_dimensions['E'].width, 26)
                
                # Normal short line: normal font, nowrap
                c32 = sheet['E32']
                self.assertFalse(c32.alignment.wrap_text)
                self.assertTrue(c32.alignment.shrink_to_fit)
                self.assertEqual(sheet.row_dimensions[32].height, 20)
                
                # Moderate line: scaled to 8.5pt, nowrap
                c33 = sheet['E33']
                self.assertFalse(c33.alignment.wrap_text)
                self.assertEqual(c33.font.size, 8.5)
                self.assertEqual(sheet.row_dimensions[33].height, 20)
                
                # Long line: scaled to 7.5pt, nowrap, height 20
                c34 = sheet['E34']
                self.assertFalse(c34.alignment.wrap_text)
                self.assertEqual(c34.font.size, 7.5)
                self.assertEqual(sheet.row_dimensions[34].height, 20)
                
                # Multiline (핫태핑): wrap_text=True, height >= 34
                c35 = sheet['E35']
                self.assertTrue(c35.alignment.wrap_text)
                self.assertGreaterEqual(sheet.row_dimensions[35].height, 34)
            finally:
                book.close()


if __name__ == '__main__':
    unittest.main()
