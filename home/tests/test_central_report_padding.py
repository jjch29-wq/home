"""Padding belongs to the preceding report page, never a blank page."""
import sys
import unittest
import os
from pathlib import Path
from unittest.mock import patch

import openpyxl
from openpyxl.worksheet.pagebreak import Break

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.central.src.services.monthly_report_manager import MonthlyReportManager


class ReportPaddingTest(unittest.TestCase):
    def test_safety_training_folder_follows_history_data_drive(self):
        manager = MonthlyReportManager('unused.xlsx')
        manager._history_path = (
            r'G:\내 드라이브\07_Antigravity\PMI_Data\central'
            r'\daily_work_history.json'
        )

        with patch.object(os.path, 'isdir', side_effect=lambda path: path.startswith('G:')):
            folder = manager._resolve_safety_training_dir()

        self.assertTrue(folder.startswith('G:'))
        self.assertTrue(folder.endswith('교육일지'))

    def test_ndt_detail_continuations_use_sixty_body_rows(self):
        manager = MonthlyReportManager('unused.xlsx')
        # 반복 문서 제목 아래 본문 60행 중 4행은 표 머리글,
        # 1행은 TOTAL이므로 검사 데이터 55행을 수용한다.
        for method in ('PAUT', 'RT', 'MT', 'PT'):
            with self.subTest(method=method):
                self.assertIsNone(manager._detail_continuation_overflow_row(100, 45))
                self.assertIsNone(manager._detail_continuation_overflow_row(100, 55))
                self.assertEqual(manager._detail_continuation_overflow_row(100, 56), 155)

    def test_last_record_stays_with_total_even_with_one_record_left(self):
        manager = MonthlyReportManager('unused.xlsx')
        sheet = openpyxl.Workbook().active
        with patch.object(manager, '_paut_print_scale', return_value=100):
            for count in (1, 4):
                with patch.object(manager, '_first_print_overflow_row', return_value=20 + count):
                    self.assertEqual(manager._detail_page_overflow_row(sheet, 20, count),
                                     20 + count - 1)

    def test_four_added_rows_consume_padding_without_moving_next_header(self):
        sheet = openpyxl.Workbook().active
        for row in range(1, 62):
            sheet.row_dimensions[row].height = 16.5
        sheet.cell(10, 2, 'TOTAL')
        sheet.cell(61, 6, 'NEXT HEADER')
        sheet.row_breaks.append(Break(id=60))
        manager = MonthlyReportManager('unused.xlsx')
        manager.table_markers = {'next': 61}
        manager._insert_rows_safely(sheet, 6, 4)
        manager.table_markers['next'] += 4
        consumed = manager._consume_detail_page_whitespace(sheet, 14, 66)
        self.assertEqual(consumed, 66)
        self.assertEqual(sheet.cell(14, 2).value, 'TOTAL')
        self.assertEqual(sheet.cell(61, 6).value, 'NEXT HEADER')
        self.assertEqual(manager.table_markers['next'], 61)
        self.assertEqual([b.id for b in sheet.row_breaks.brk], [60])

    def test_only_unused_height_is_consumed_and_data_is_preserved(self):
        sheet = openpyxl.Workbook().active
        sheet.cell(10, 2, 'TOTAL')
        sheet.cell(12, 2, 'FOLLOWING TABLE')
        sheet.row_dimensions[13].height = 10
        sheet.row_dimensions[14].height = 20
        sheet.row_breaks.append(Break(id=14))
        manager = MonthlyReportManager('unused.xlsx')
        consumed = manager._consume_detail_page_whitespace(sheet, 10, 16.5)
        self.assertEqual(consumed, 16.5)
        self.assertEqual(sheet.row_dimensions[14].height, 3.5)
        self.assertEqual(sheet.cell(12, 2).value, 'FOLLOWING TABLE')
        self.assertEqual([b.id for b in sheet.row_breaks.brk], [14])
        consumed = manager._consume_detail_page_whitespace(sheet, 10, 100)
        self.assertEqual(consumed, 13.5)
        self.assertEqual(sheet.cell(12, 2).value, 'FOLLOWING TABLE')

    def test_when_page_is_full_new_detail_rows_overflow(self):
        sheet = openpyxl.Workbook().active
        for row in range(1, 70):
            sheet.row_dimensions[row].height = 16.5
        manager = MonthlyReportManager('unused.xlsx')
        self.assertIsNone(manager._first_print_overflow_row(sheet, 10, 30))
        self.assertIsNotNone(manager._first_print_overflow_row(sheet, 10, 60))

    def test_paut_tail_uses_actual_width_fit_instead_of_full_size(self):
        sheet = openpyxl.Workbook().active
        sheet.print_area = 'B1:W60'
        sheet.page_setup.fitToWidth = 1
        sheet.page_setup.scale = None
        for dimension in sheet.column_dimensions.values():
            dimension.width = 7
        from openpyxl.utils import get_column_letter
        for col in range(2, 24):
            sheet.column_dimensions[get_column_letter(col)].width = 7
        for row in range(1, 11):
            sheet.row_dimensions[row].height = 16
        for row in range(11, 48):
            sheet.row_dimensions[row].height = 16.5
        manager = MonthlyReportManager('unused.xlsx')
        self.assertIsNotNone(manager._first_print_overflow_row(sheet, 11, 37))
        scale = manager._paut_print_scale(sheet)
        self.assertLess(scale, 100)
        self.assertIsNone(manager._first_print_overflow_row(sheet, 11, 37, scale))

    def test_moves_header_boundary_after_inserted_padding(self):
        sheet = openpyxl.Workbook().active
        for start in (1, 66, 112):
            sheet.cell(start, 6, '월 간 용 역 진 도 보 고 서')
            sheet.cell(start + 4, 6, '3.0 비파괴검사현황')
        sheet.cell(111, 6, 'PAUT 마지막 검사행')
        for end in (65, 111, 176):
            sheet.row_breaks.append(Break(id=end))
        manager = MonthlyReportManager('unused.xlsx')
        manager._pad_ndt_continuation_pages(sheet)
        self.assertEqual(sorted(b.id for b in sheet.row_breaks.brk), [65, 130, 195])
        self.assertEqual(sheet.cell(111, 6).value, 'PAUT 마지막 검사행')
        self.assertEqual(sheet.cell(131, 6).value, '월 간 용 역 진 도 보 고 서')
        self.assertFalse(any(sheet.cell(r, 6).value for r in range(112, 131)))
        manager._align_report_breaks_to_repeated_headers(sheet)
        self.assertNotIn(111, [b.id for b in sheet.row_breaks.brk])
        self.assertIn(130, [b.id for b in sheet.row_breaks.brk])

    def test_pt_spacer_is_compacted_to_keep_document_header(self):
        sheet = openpyxl.Workbook().active
        sheet.cell(20, 2, 'RT TOTAL')
        sheet.cell(23, 2, '2.4 비파괴검사 세부현황(PT)')
        manager = MonthlyReportManager('unused.xlsx')

        manager._compact_ndt_method_spacers(sheet)

        self.assertEqual(sheet.row_dimensions[21].height, 1.0)
        self.assertEqual(sheet.row_dimensions[22].height, 1.0)
        self.assertFalse(sheet.row_dimensions[21].hidden)

    def test_only_weld_defect_trailing_rows_are_compacted(self):
        sheet = openpyxl.Workbook().active
        sheet.cell(300, 6, '4.0 용접결함, 불량률 현황')
        sheet.cell(310, 2, 'TABLE TOTAL')
        sheet.row_breaks.append(Break(id=315))
        manager = MonthlyReportManager('unused.xlsx')

        compacted = manager._compact_weld_defect_trailing_spacers(sheet)

        self.assertEqual(compacted, 5)
        for row in range(311, 316):
            self.assertEqual(sheet.row_dimensions[row].height, 1.0)
            self.assertFalse(sheet.row_dimensions[row].hidden)

    def test_formatted_blank_safety_rows_are_not_bulk_compacted(self):
        sheet = openpyxl.Workbook().active
        sheet.cell(300, 6, '월간 용역 진도 보고서')
        sheet.cell(310, 2, '방사선안전관리 표')
        sheet.cell(316, 6, '월간 용역 진도 보고서')
        manager = MonthlyReportManager('unused.xlsx')

        manager._collapse_blank_rows_before_report_headers(sheet)

        self.assertIsNone(sheet.row_dimensions[311].height)
        self.assertIsNone(sheet.row_dimensions[315].height)

    def test_pt_overflow_gets_repeated_document_header(self):
        sheet = openpyxl.Workbook().active
        sheet.cell(100, 6, '월간 용역 진도 보고서')
        sheet.cell(104, 6, '3.0 비파괴검사현황')
        sheet.cell(120, 6, '2.4 비파괴검사 세부현황(PT)')
        sheet.cell(124, 2, 'TOTAL')
        sheet.row_breaks.append(Break(id=99))
        sheet.row_breaks.append(Break(id=124))
        manager = MonthlyReportManager('unused.xlsx')
        manager.table_markers = {'next': 125}

        with patch.object(manager, '_first_print_overflow_row', return_value=118):
            with patch.object(manager, '_paut_print_scale', return_value=88):
                inserted = manager._ensure_pt_continuation_header(sheet)

        self.assertEqual(inserted, 7)
        self.assertEqual(sheet.cell(120, 6).value, '월간 용역 진도 보고서')
        self.assertEqual(sheet.cell(124, 6).value, '3.0 비파괴검사현황')
        self.assertEqual(sheet.cell(127, 6).value, '2.4 비파괴검사 세부현황(PT)')
        self.assertIn(119, [int(brk.id) for brk in sheet.row_breaks.brk])
        self.assertEqual(manager.table_markers['next'], 132)

if __name__ == '__main__':
    unittest.main()
