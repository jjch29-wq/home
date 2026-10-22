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
    def test_photo_frame_tail_fits_without_resizing_photos_or_caption(self):
        from types import SimpleNamespace
        manager = MonthlyReportManager('unused.xlsx')
        sheet = openpyxl.Workbook().active
        sheet.page_setup.scale = 100
        sheet.page_margins.top = sheet.page_margins.bottom = .75
        sheet.page_margins.footer = .3
        sheet.print_area = 'B1:W45'
        sheet.cell(5, 6, '6.0 사진대지')
        sheet.cell(41, 2, '사진 설명')
        for r in range(1, 46):
            sheet.row_dimensions[r].height = 16
        anchor = SimpleNamespace(_from=SimpleNamespace(row=10), to=SimpleNamespace(row=39))
        sheet._images.append(SimpleNamespace(anchor=anchor))
        sheet.row_breaks.append(Break(id=45))
        self.assertIsNotNone(manager._first_print_overflow_row(sheet, 1, 45, 103))
        manager._fit_photo_page_trailing_rows(sheet)
        self.assertIsNone(manager._first_print_overflow_row(sheet, 1, 45, 103))
        self.assertTrue(all(sheet.row_dimensions[r].height == 16 for r in range(1, 42)))
        self.assertEqual(sheet.cell(41, 2).value, '사진 설명')
        self.assertEqual((anchor._from.row, anchor.to.row), (10, 39))
        self.assertEqual([b.id for b in sheet.row_breaks.brk], [45])

    def test_rt_gap_compaction_keeps_last_record_and_total_on_page(self):
        manager = MonthlyReportManager('unused.xlsx')
        sheet = openpyxl.Workbook().active
        sheet.page_setup.scale = 100
        for r in range(1, 39):
            sheet.row_dimensions[r].height = 18
        sheet.cell(15, 2, 'MT TOTAL')
        sheet.cell(20, 2, '2.3 비파괴검사 세부현황(RT)')
        sheet.cell(21, 6, 'Line No.')
        for r in range(23, 38):
            sheet.cell(r, 3, r - 22)
        sheet.cell(38, 2, 'TOTAL')
        self.assertIsNotNone(manager._detail_page_overflow_row(sheet, 23, 15))
        manager._compact_rt_leading_gap(sheet, 21)
        self.assertIsNone(manager._detail_page_overflow_row(sheet, 23, 15))
        self.assertEqual([sheet.row_dimensions[r].height for r in range(16, 20)], [2] * 4)
        for r in (15, 20, 21, 22, 23, 37, 38):
            self.assertEqual(sheet.row_dimensions[r].height, 18)
        self.assertEqual(sheet.cell(37, 3).value, 15)
        self.assertEqual(sheet.cell(38, 2).value, 'TOTAL')
        self.assertFalse(sheet.row_breaks.brk)

    def test_defect_tables_stay_together_by_reclaiming_empty_gap(self):
        manager = MonthlyReportManager('unused.xlsx')
        sheet = openpyxl.Workbook().active
        sheet.page_setup.scale = 100
        sheet.cell(1, 6, '월간 용역 진도 보고서')
        sheet.cell(5, 6, '4.0 용접결함, 불량률 현황')
        sheet.cell(14, 2, '총계')
        sheet.cell(35, 2, '1.2 PAUT 용접 불량률')
        sheet.cell(38, 2, 'ID')
        sheet.cell(38, 3, '성명')
        sheet.cell(47, 2, '총계')
        for r in range(1, 48):
            sheet.row_dimensions[r].height = 20
        sheet.row_breaks.append(Break(id=47))
        originals = {f'W-{i}': 100 for i in range(8)}
        manager._write_welder_defect_summary(sheet, {}, {}, originals, '1.2', ('PAUT',))
        self.assertEqual([b.id for b in sheet.row_breaks.brk], [47])
        self.assertIsNone(manager._first_print_overflow_row(sheet, 1, 47, 103))
        self.assertEqual(sheet.row_dimensions[20].height, 2)
        for r in (1, 5, 14, 35, 38, 39, 46, 47):
            self.assertEqual(sheet.row_dimensions[r].height, 20)
        self.assertEqual([sheet.cell(r, 2).value for r in range(39, 47)], sorted(originals))

    def test_final_limit_does_not_copy_empty_mt_header_into_blank_page(self):
        manager = MonthlyReportManager('unused.xlsx')
        sheet = openpyxl.Workbook().active
        sheet.cell(1, 6, '월간 용역 진도 보고서')
        sheet.cell(5, 6, '3.0 비파괴검사현황')
        sheet.cell(49, 2, '1.2.2 MT')
        sheet.cell(50, 6, 'Line No.')
        sheet.cell(52, 2, '해당 월 검사실적 없음')
        sheet.cell(53, 2, 'TOTAL')
        sheet.row_breaks.append(Break(id=70))
        manager._enforce_final_ndt_body_limit(sheet)
        self.assertEqual([b.id for b in sheet.row_breaks.brk], [70])
        self.assertEqual(sum('Line No.' == sheet.cell(r, 6).value
                             for r in range(1, sheet.max_row + 1)), 1)

    def test_final_limit_moves_whole_table_without_duplicate_heading(self):
        manager = MonthlyReportManager('unused.xlsx')
        sheet = openpyxl.Workbook().active
        sheet.cell(1, 6, '월간 용역 진도 보고서')
        sheet.cell(5, 6, '3.0 비파괴검사현황')
        sheet.cell(55, 2, '1.2.2 MT')
        sheet.cell(56, 6, 'Line No.')
        sheet.cell(58, 2, '업체')
        sheet.cell(58, 3, 1)
        sheet.cell(59, 2, 'TOTAL')
        sheet.row_breaks.append(Break(id=59))
        manager._enforce_final_ndt_body_limit(sheet)
        self.assertEqual(sum('Line No.' == sheet.cell(r, 6).value
                             for r in range(1, sheet.max_row + 1)), 1)
        self.assertEqual(sheet.cell(65, 3).value, 1)
        self.assertEqual([b.id for b in sheet.row_breaks.brk], [54, 66])

    def test_final_limit_repairs_late_insertions_and_survives_save(self):
        import io
        manager = MonthlyReportManager('unused.xlsx')
        sheet = openpyxl.Workbook().active
        sheet.cell(1, 6, '월간 용역 진도 보고서')
        sheet.cell(5, 6, '3.0 비파괴검사현황')
        sheet.cell(9, 2, '2.1 PAUT')
        sheet.cell(10, 6, 'Line No.')
        sheet.merge_cells('F10:G11')
        for index in range(58):
            sheet.cell(12 + index, 2, index + 1)
            sheet.row_dimensions[12 + index].height = 9
        sheet.cell(70, 2, 'TOTAL')
        sheet.row_breaks.append(Break(id=70))
        sheet.cell(71, 6, '4.0 용접결함, 불량률 현황')
        manager._enforce_final_ndt_body_limit(sheet)
        buffer = io.BytesIO()
        sheet.parent.save(buffer)
        buffer.seek(0)
        result = openpyxl.load_workbook(buffer).active
        self.assertEqual(min(b.id for b in result.row_breaks.brk), 58)
        start = 1
        for end in sorted(b.id for b in result.row_breaks.brk):
            self.assertLessEqual(end - start + 1, 58)
            start = end + 1
        records = [result.cell(r, 2).value for r in range(1, result.max_row + 1)
                   if isinstance(result.cell(r, 2).value, int)]
        self.assertEqual(records, list(range(1, 59)))
        self.assertEqual(sum(result.cell(r, 2).value == 'TOTAL'
                             for r in range(1, result.max_row + 1)), 1)
        count = result.max_row
        manager._enforce_final_ndt_body_limit(result)
        self.assertEqual(result.max_row, count)

    def test_defect_table_continuations_do_not_unmerge_next_table_header(self):
        sheet = openpyxl.Workbook().active
        # The middle block is a continuation without its own TOTAL. The old
        # search used the next table's TOTAL and consumed that table's header.
        for header in (700, 715, 729):
            sheet.cell(header + 1, 2, 'ID')
            for col in (16, 17, 18):
                sheet.cell(header, col, f'HEADER {col}')
                sheet.merge_cells(start_row=header, end_row=header + 1,
                                  start_column=col, end_column=col)
            sheet.cell(header + 2, 16, header)
        for total in (704, 733):
            sheet.cell(total, 2, '총계')
            sheet.cell(total, 16, 123)
        manager = MonthlyReportManager('unused.xlsx')
        manager._widen_weld_defect_table(sheet)
        self.assertEqual(sheet.cell(717, 16).value, 715)
        self.assertEqual(sheet.cell(731, 16).value, 729)
        self.assertEqual(sheet.cell(733, 16).value, 123)
        merges = {str(m) for m in sheet.merged_cells.ranges}
        self.assertIn('P729:Q730', merges)
        self.assertIn('P715:Q716', merges)

    def test_summary_compaction_only_reclaims_blank_table_gaps(self):
        manager = MonthlyReportManager('unused.xlsx')
        manager.table_markers.update(qty=10, paut_1=20, mt_1=40,
                                     rt_1=60, pt_1=80, paut_2=90)
        sheet = openpyxl.Workbook().active
        for r in range(1, 91):
            sheet.row_dimensions[r].height = 15
        for header in (20, 40, 60, 80):
            sheet.cell(header, 2, '표 머리글')
            sheet.cell(header + 2, 2, '검사 데이터')
            sheet.cell(header + 3, 2, 'TOTAL')
        sheet.cell(25, 6, '보존할 소제목')
        sheet.merge_cells('B26:C26')
        manager._compact_ndt_summary_page(sheet)
        self.assertLess(sheet.row_dimensions[24].height, 15)
        for r in (20, 22, 23, 25, 26, 80, 82, 83, 90):
            self.assertEqual(sheet.row_dimensions[r].height, 15)
        self.assertEqual(sheet.cell(22, 2).value, '검사 데이터')
        self.assertIn('B26:C26', {str(m) for m in sheet.merged_cells.ranges})

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

    def test_detail_continuations_fill_available_height_and_preserve_records(self):
        manager = MonthlyReportManager('unused.xlsx')
        sheet = openpyxl.Workbook().active
        sheet.page_setup.scale = 100
        sheet.cell(1, 6, '월간 용역 진도 보고서')
        sheet.cell(5, 6, '3.0 비파괴검사현황')
        sheet.cell(30, 2, 'PAUT')
        sheet.cell(31, 2, '순번')
        for r in range(1, 33):
            sheet.row_dimensions[r].height = 12
        for index in range(160):
            sheet.cell(33 + index, 2, index + 1)
            sheet.row_dimensions[33 + index].height = 9
        sheet.cell(193, 2, 'TOTAL')
        sheet.row_breaks.append(Break(id=193))
        start, remaining = 33, 160
        counts = []
        while remaining:
            split = manager._detail_continuation_overflow_row(sheet, start, remaining)
            if split is None:
                counts.append(remaining)
                break
            count = split - start
            self.assertGreater(count, 0)
            counts.append(count)
            inserted = manager._insert_table_continuation(sheet, split, 1, 31)
            sheet.row_breaks.append(Break(id=split - 1))
            remaining -= count
            start = split + inserted
            self.assertLess(len(counts), 10)
        records = [sheet.cell(r, 2).value for r in range(1, sheet.max_row + 1)
                   if isinstance(sheet.cell(r, 2).value, int)]
        self.assertEqual(records, list(range(1, 161)))
        self.assertEqual(counts[1], 47)
        self.assertLess(counts[0], counts[1])
        page_start = 1
        for end in sorted(b.id for b in sheet.row_breaks.brk):
            self.assertLessEqual(end - page_start + 1, 58)
            self.assertIsNone(manager._first_print_overflow_row(
                sheet, page_start, end - page_start + 1, 103))
            page_start = end + 1
        for r in range(1, sheet.max_row + 1):
            if isinstance(sheet.cell(r, 2).value, int):
                self.assertEqual(sheet.row_dimensions[r].height, 9)

    def test_body_limit_excludes_seven_title_rows_and_reserves_total(self):
        manager = MonthlyReportManager('unused.xlsx')
        sheet = openpyxl.Workbook().active
        sheet.page_setup.scale = 100
        for r in range(1, 150):
            sheet.row_dimensions[r].height = 5
        # Page 101:107 is the title, 108:111 is the table heading.
        sheet.row_breaks.append(Break(id=100))
        self.assertIsNone(manager._detail_page_overflow_row(sheet, 112, 46))
        self.assertEqual(manager._detail_page_overflow_row(sheet, 112, 47), 158)
        # With more records following, use the 58th row for data instead
        # of leaving it empty for a TOTAL that belongs on the final page.
        self.assertEqual(manager._detail_page_overflow_row(sheet, 112, 48), 159)

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
