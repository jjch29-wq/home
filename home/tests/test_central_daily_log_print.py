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


if __name__ == '__main__':
    unittest.main()
