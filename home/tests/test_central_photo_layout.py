"""Check photo/caption boundaries after workbook serialization."""
import sys
import tempfile
import unittest
from pathlib import Path

import openpyxl
from PIL import Image
from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor
from openpyxl.utils.units import pixels_to_EMU

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.central.src.services.monthly_report_manager import MonthlyReportManager


class PhotoLayoutTest(unittest.TestCase):
    def test_fixed_ndt_top_spacer_uses_ten_points(self):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.cell(100, 2, '월간 용역 진도 보고서')
        sheet.cell(104, 6, '3.0 비파괴검사현황')
        manager = MonthlyReportManager('unused.xlsx')

        inserted = manager._insert_fixed_ndt_top_spacers(sheet)

        self.assertEqual(inserted, 1)
        self.assertEqual(sheet.row_dimensions[106].height, 25.0)
        self.assertEqual(sheet.cell(100, 2).value, '월간 용역 진도 보고서')

    def test_sec16_marker_is_retained_after_photo_layout(self):
        with tempfile.TemporaryDirectory() as folder:
            folder = Path(folder)
            source = folder / 'photo.png'
            Image.new('RGB', (320, 240), 'steelblue').save(source)
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.cell(610, 3, '1.1 위상배열초음파탐상검사(SEC.16)')
            for row in range(600, 686):
                sheet.row_dimensions[row].height = 2.7
            sheet.print_area = 'B600:W685'
            manager = MonthlyReportManager('unused.xlsx')
            manager._history_path = str(folder / 'history.json')

            manager._populate_process_photo_pages(sheet, [dict(
                process='PAUT', file_path=str(source), date='2026-10-16',
                joint_no='S01', description='검사 공정사진', location='Sec.03',
            )])

            self.assertEqual(len(sheet._images), 1)
            self.assertIn('SEC.16', str(sheet.cell(610, 3).value).upper())

    def test_both_corners_stay_inside_frame_above_caption(self):
        with tempfile.TemporaryDirectory() as folder:
            folder = Path(folder)
            source = folder / 'photo.png'
            Image.new('RGB', (480, 640), 'steelblue').save(source)
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.cell(610, 3, '1.1 위상배열초음파탐상검사(SEC.16)')
            for row in range(600, 686):
                sheet.row_dimensions[row].height = 2.7
            sheet.print_area = 'B600:W685'
            manager = MonthlyReportManager('unused.xlsx')
            manager._history_path = str(folder / 'history.json')
            photos = [dict(process='PAUT', file_path=str(source),
                           date='2026-10-16', joint_no=f'S{i:02}',
                           description='검사 공정사진', location='Sec.16')
                      for i in range(18)]
            manager._populate_process_photo_pages(sheet, photos)
            output = folder / 'report.xlsx'
            workbook.save(output)
            checked = openpyxl.load_workbook(output).active
            self.assertEqual(len(checked._images), 18)
            self.assertIn('SEC.16', str(checked.cell(610, 3).value).upper())
            inset = pixels_to_EMU(6)
            for i, picture in enumerate(checked._images):
                anchor = picture.anchor
                self.assertIsInstance(anchor, TwoCellAnchor)
                self.assertEqual(anchor._from.colOff, inset)
                self.assertEqual(anchor._from.rowOff, inset)
                self.assertEqual(anchor.to.colOff, -inset)
                self.assertEqual(anchor.to.rowOff, -inset)
                self.assertEqual(anchor.to.row - anchor._from.row, 16)
                caption_row = anchor.to.row + 1
                caption = checked.cell(caption_row, anchor._from.col + 1)
                self.assertIn('검사 공정사진', caption.value)
                self.assertFalse(caption.alignment.wrap_text)
                self.assertTrue(caption.alignment.shrink_to_fit)
                self.assertEqual(checked.row_dimensions[caption_row].height, 30)
                if i % 2:
                    left = checked._images[i - 1].anchor
                    self.assertEqual(left._from.row, anchor._from.row)
                    self.assertEqual(left.to.row, anchor.to.row)
                    self.assertEqual(left._from.rowOff, anchor._from.rowOff)
                    self.assertEqual(left.to.rowOff, anchor.to.rowOff)


if __name__ == '__main__':
    unittest.main()
