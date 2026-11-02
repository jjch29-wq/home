"""Import wrapped owner-export cells without splitting report data lines."""
import sys
import tempfile
import unittest
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.central.src.services.monthly_report_manager import MonthlyReportManager


class OwnerReportTest(unittest.TestCase):
    def test_long_names_paginate_by_height_and_keep_following_photo_aligned(self):
        from PIL import Image as PILImage
        from openpyxl.drawing.image import Image
        from openpyxl.worksheet.pagebreak import Break
        with tempfile.TemporaryDirectory() as folder:
            folder = Path(folder)
            source = folder / 'owner.xlsx'
            workbook = openpyxl.Workbook()
            inputs = workbook.active
            inputs.title = '작업정보'
            names = [f'서울특별시서초구양재동열수송관현장{i}' for i in range(3)]
            for row, name in enumerate(names, 5):
                inputs.cell(row, 1, '202609')
                inputs.cell(row, 5, name)
            workbook.save(source)
            target = openpyxl.Workbook()
            sheet = target.active
            sheet['F1'] = '월간 용역 진도 보고서'
            sheet['F5'] = '6.0 방사선안전관리통합정보망 보고자료'
            sheet['B10'] = '1. 방사선안전관리통합정보망 보고자료'
            sheet.column_dimensions['F'].width = 6
            for row in range(14, 57):
                sheet.row_dimensions[row].height = 2
            sheet.row_breaks.append(Break(id=56))
            sheet['B60'] = '다음 사진'
            from io import BytesIO
            photo_buffer = BytesIO()
            PILImage.new('RGB', (100, 100), 'red').save(photo_buffer, format='PNG')
            photo_buffer.seek(0)
            sheet.add_image(Image(photo_buffer), 'C60')
            template = folder / 'template.xlsx'
            target.save(template)
            target = openpyxl.load_workbook(template)
            sheet = target.active
            manager = MonthlyReportManager('unused.xlsx')
            manager._populate_owner_report_rows(sheet, '2026-09', str(source))
            result_path = folder / 'report.xlsx'
            manager._save_workbook_with_fallback(target, str(result_path))
            result = openpyxl.load_workbook(result_path).active
            found = [result.cell(r, 6).value for r in range(1, result.max_row + 1)
                     if result.cell(r, 6).value in names]
            self.assertEqual(found, names)
            body_starts = [r + 4 for r in range(1, result.max_row + 1)
                           if str(result.cell(r, 2).value or '').startswith(sheet['B10'].value)]
            self.assertEqual(len(body_starts), 3)
            for start in body_starts:
                self.assertAlmostEqual(sum(
                    0 if result.row_dimensions[r].hidden else result.row_dimensions[r].height
                    for r in range(start, start + 43)), 86)
            photo_row = next(r for r in range(1, result.max_row + 1)
                             if result.cell(r, 2).value == '다음 사진')
            self.assertEqual(result._images[0].anchor._from.row + 1, photo_row)

    def test_body_joins_hard_breaks_and_keeps_codes_and_headers(self):
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / 'owner.xlsx'
            workbook = openpyxl.Workbook()
            inputs = workbook.active
            inputs.title = '작업정보'
            values = {
                1: '2026\n09', 2: '0993', 3: '서울\r\n검사',
                4: '2946\n81', 5: '한국지역난방공사 중앙지사', 6: 150,
                7: '2026-09-\n29', 8: '사용시설\n이외',
                9: '추진\n철', 10: '박광\n복', 15: '08:\n30',
                16: '16:10', 17: 6, 18: '파이프',
                19: 'SPPS38\n0', 20: 4.5, 25: 0.79, 26: 1,
            }
            for col, value in values.items():
                inputs.cell(5, col, value)
            # Owner exports include registered locations with no work at all.
            for col, value in {1:'202609', 2:'0000', 9:'0', 17:0,
                               18:'0', 19:'0', 20:0, 25:0, 26:0}.items():
                inputs.cell(6, col, value)
            # A genuine zero-quantity job with crew information must survive.
            for col, value in {1:'202609', 2:'0111', 9:'추진철', 17:0}.items():
                inputs.cell(7, col, value)
            workbook.save(source)

            target = openpyxl.Workbook()
            sheet = target.active
            sheet['B1'] = '1. 방사선안전관리통합정보망 보고자료'
            sheet['D2'] = '검사\n업체명'
            sheet['D2'].alignment = Alignment(wrap_text=True)
            manager = MonthlyReportManager('unused.xlsx')
            manager._populate_owner_report_rows(sheet, '2026-09', str(source))
            output = Path(folder) / 'report.xlsx'
            manager._save_workbook_with_fallback(target, str(output))
            result = openpyxl.load_workbook(output).active

            self.assertEqual(result['B5'].value, '202609')
            self.assertEqual(result['C5'].value, '0993')
            self.assertEqual(result['C6'].value, '0111')
            self.assertIsNone(result['C7'].value)
            self.assertEqual(result['D5'].value, '서울검사')
            self.assertEqual(result['E5'].value, '294681')
            self.assertEqual(result['F5'].value, values[5])
            self.assertEqual(result['H5'].value, '2026-09-29')
            self.assertEqual(result['J5'].value, '추진철')
            self.assertEqual(result['K5'].value, '박광복')
            self.assertEqual(result['P5'].value, '08:30')
            self.assertEqual(result['T5'].value, 'SPPS380')
            self.assertEqual(result['U5'].value, 4.5)
            for row in result.iter_rows(min_row=5, max_row=5, min_col=2, max_col=23):
                for cell in row:
                    if cell.column == 6:
                        self.assertTrue(cell.alignment.wrap_text)
                        self.assertFalse(cell.alignment.shrink_to_fit)
                        self.assertEqual(cell.font.sz, 6)
                    else:
                        self.assertFalse(cell.alignment.wrap_text)
                        self.assertTrue(cell.alignment.shrink_to_fit)
            self.assertGreater(result.row_dimensions[5].height, 15)
            self.assertAlmostEqual(sum(
                0 if result.row_dimensions[r].hidden else result.row_dimensions[r].height
                for r in range(5, 48)), 43 * 15)
            self.assertEqual(result['D2'].value, '검사\n업체명')
            self.assertTrue(result['D2'].alignment.wrap_text)

    def test_strips_redundant_branch_prefix_from_workplace_name(self):
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / 'owner.xlsx'
            workbook = openpyxl.Workbook()
            inputs = workbook.active
            inputs.title = '작업정보'
            inputs.cell(5, 1, '202609')
            inputs.cell(5, 5, '한국지역난방공사 중앙지사 서울특별시 서초구')
            inputs.cell(5, 9, '주진철')
            inputs.cell(5, 17, 6)
            workbook.save(source)

            target = openpyxl.Workbook()
            sheet = target.active
            sheet['B1'] = '1. 방사선안전관리통합정보망 보고자료'
            manager = MonthlyReportManager('unused.xlsx')
            manager._populate_owner_report_rows(sheet, '2026-09', str(source))
            self.assertEqual(sheet['F5'].value, '서울특별시 서초구')

    def test_last_visible_empty_row_keeps_bottom_border_when_trailing_row_hidden(self):
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / 'owner.xlsx'
            workbook = openpyxl.Workbook()
            inputs = workbook.active
            inputs.title = '작업정보'
            inputs.cell(5, 1, '202609')
            inputs.cell(5, 5, '서울특별시 서초구')
            inputs.cell(5, 9, '주진철')
            inputs.cell(5, 17, 6)
            workbook.save(source)

            target = openpyxl.Workbook()
            sheet = target.active
            sheet['B1'] = '1. 방사선안전관리통합정보망 보고자료'
            for row in range(5, 48):
                sheet.row_dimensions[row].height = 13.35
            manager = MonthlyReportManager('unused.xlsx')
            manager._populate_owner_report_rows(sheet, '2026-09', str(source))
            self.assertTrue(sheet.row_dimensions[47].hidden)
            self.assertFalse(sheet.row_dimensions[46].hidden)
            for col in range(2, 24):
                self.assertEqual(sheet.cell(46, col).border.bottom.style, 'thin')

    def test_inspector_qualification_row_expands_without_changing_page_height(self):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.row_dimensions[176].height = 17.25
        sheet.row_dimensions[177].height = 17.25
        sheet.cell(177, 2, '8. 검사원의 자격')
        sheet.row_dimensions[178].height = 14.25
        sheet.row_dimensions[180].height = 17.25
        sheet.cell(180, 2, '가. 본 검사에 투입되는 검사자는')
        sheet.row_dimensions[181].height = 17.25
        sheet.row_dimensions[182].height = 20.25
        sheet.cell(182, 3, '구 분')
        sheet.row_dimensions[183].height = 85.15
        sheet.cell(183, 3, '현장대리인\n(책임자)')
        sheet.cell(183, 7, '○ 다음 각 호중 하나에 해당하는 자격을 갖춘 자')
        sheet.row_dimensions[188].height = 115.15
        sheet.cell(188, 5, 'RT')
        sheet.cell(188, 20, '비파괴검사법\n제13조,\n제58조')

        tracked_rows = range(168, 190)
        before_total = sum(
            float(sheet.row_dimensions[r].height or 15.0)
            for r in tracked_rows
        )
        manager = MonthlyReportManager('unused.xlsx')
        manager._fit_inspector_qualification_rows(sheet)
        after_total = sum(
            float(sheet.row_dimensions[r].height or 15.0)
            for r in tracked_rows
        )
        self.assertEqual(sheet.row_dimensions[183].height, 100.0)
        self.assertEqual(sheet.row_dimensions[188].height, 118.0)
        self.assertAlmostEqual(before_total, after_total)

    def test_ensure_valid_print_area_restores_missing_or_corrupted_print_area(self):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.cell(800, 23, 'end')
        manager = MonthlyReportManager('unused.xlsx')

        sheet.print_area = None
        manager._ensure_valid_print_area(sheet)
        self.assertIn('$B$1:$W$838', str(sheet.print_area))

        sheet.print_area = 'Y519:Y519'
        manager._set_print_area_end(sheet, 750)
        self.assertIn('$B$1:$W$750', str(sheet.print_area))


if __name__ == '__main__':
    unittest.main()
