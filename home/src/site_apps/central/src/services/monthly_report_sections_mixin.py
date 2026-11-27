import openpyxl
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
from openpyxl.cell.cell import Cell, MergedCell
import os
import tempfile

# [FIX] Patch openpyxl MergedCell to prevent 'read-only' error when assigning to merged cells
if not hasattr(MergedCell, '_patched_value'):
    MergedCell._patched_value = True
    MergedCell.value = property(lambda self: None, lambda self, val: None)

import json
import datetime
from collections import defaultdict
import re



class MonthlyReportSectionsMixin:
    def _populate_radiation_network_report(self, ws, target_ym):
        """Replace the section 6.0 screenshot with a real editable cell table."""
        title_row = None
        for row in range(1, ws.max_row + 1):
            row_text = ''.join(
                str(ws.cell(row=row, column=col).value or '')
                for col in range(1, min(ws.max_column, 30) + 1)
            ).replace(' ', '')
            if '1.방사선안전관리통합정보망보고자료' in row_text:
                title_row = row
                break
        if title_row is None:
            return

        table_start = title_row + 1
        next_page_start = None
        for row in range(table_start + 1, min(ws.max_row, table_start + 80) + 1):
            row_text = ''.join(
                str(ws.cell(row=row, column=col).value or '')
                for col in range(1, min(ws.max_column, 30) + 1)
            ).replace(' ', '')
            if '6.0사진대지' in row_text:
                next_page_start = max(table_start + 1, row - 4)
                break
        page_end = (next_page_start - 1) if next_page_start else table_start + 46

        # Remove only the captured report image. Logos in the common header are
        # anchored above table_start and must remain on the page.
        kept_images = []
        for image in list(ws._images):
            anchor = getattr(image, 'anchor', None)
            if hasattr(anchor, '_from'):
                image_row = anchor._from.row + 1
                if table_start <= image_row <= page_end:
                    continue
            kept_images.append(image)
        ws._images = kept_images

        source_path = os.path.join(
            os.path.expanduser('~'), 'Downloads', '발주자 보고.xlsx'
        )
        records = []
        if os.path.exists(source_path):
            try:
                source_wb = openpyxl.load_workbook(source_path, data_only=True)
                source_ws = source_wb['작업정보'] \
                    if '작업정보' in source_wb.sheetnames else source_wb.active
                target_key = str(target_ym or '').replace('-', '')[:6]

                def _text(value):
                    if value is None:
                        return ''
                    if isinstance(value, (datetime.datetime, datetime.date)):
                        return value.strftime('%Y-%m-%d')
                    return str(value).strip()

                def _clock(value):
                    text = _text(value)
                    if len(text) == 4 and text.isdigit():
                        return f'{text[:2]}:{text[2:]}'
                    return text

                for row in range(5, source_ws.max_row + 1):
                    year_month = _text(source_ws.cell(row, 1).value).replace('-', '')[:6]
                    if target_key and year_month != target_key:
                        continue
                    workers = [
                        _text(source_ws.cell(row, col).value)
                        for col in range(10, 15)
                    ]
                    workers = ', '.join(value for value in workers if value)

                    source_name = ''
                    source_activity = ''
                    source_count = ''
                    for name, activity_col, count_col in (
                        ('Ir-192', 21, 22), ('Co-60', 23, 24),
                        ('Se-75', 25, 26), ('Am-241', 27, 28),
                    ):
                        activity = source_ws.cell(row, activity_col).value
                        count = source_ws.cell(row, count_col).value
                        if activity not in (None, '') or count not in (None, ''):
                            source_name = name
                            source_activity = activity if activity is not None else ''
                            source_count = count if count is not None else ''
                            break

                    generator = ' / '.join(
                        value for value in (
                            _text(source_ws.cell(row, 29).value),
                            _text(source_ws.cell(row, 30).value),
                        ) if value
                    )
                    records.append([
                        year_month,
                        _text(source_ws.cell(row, 2).value),
                        _text(source_ws.cell(row, 3).value),
                        _text(source_ws.cell(row, 4).value),
                        _text(source_ws.cell(row, 5).value),
                        source_ws.cell(row, 6).value or '',
                        _text(source_ws.cell(row, 7).value),
                        _text(source_ws.cell(row, 8).value),
                        _text(source_ws.cell(row, 9).value),
                        workers,
                        _clock(source_ws.cell(row, 15).value),
                        _clock(source_ws.cell(row, 16).value),
                        source_ws.cell(row, 17).value or '',
                        _text(source_ws.cell(row, 18).value),
                        _text(source_ws.cell(row, 19).value),
                        source_ws.cell(row, 20).value or '',
                        source_name,
                        source_activity,
                        source_count,
                        generator,
                        source_ws.cell(row, 31).value or '',
                        _text(source_ws.cell(row, 32).value),
                    ])
                source_wb.close()
            except Exception as exc:
                print(f'[WARN] 발주자 보고자료를 읽지 못했습니다: {exc}')

        # Protect the following photo page when a month contains more rows than
        # the original screenshot area. The page break and print area move with
        # the inserted rows; drawings need the same explicit anchor adjustment.
        planned_data_start = table_start + 2
        available_rows = max(1, page_end - planned_data_start + 1)
        extra_rows = max(0, len(records) - available_rows)
        if extra_rows and next_page_start:
            self._insert_rows_safely(ws, next_page_start, extra_rows)
            for image in list(ws._images):
                anchor = getattr(image, 'anchor', None)
                if not hasattr(anchor, '_from') or anchor._from.row + 1 < next_page_start:
                    continue
                anchor._from.row += extra_rows
                if hasattr(anchor, 'to'):
                    anchor.to.row += extra_rows
            page_end += extra_rows

        # Clear the old screenshot area while retaining the surrounding page.
        for merged in list(ws.merged_cells.ranges):
            if merged.min_row >= table_start and merged.max_row <= page_end:
                ws.unmerge_cells(str(merged))
        for row in range(table_start, page_end + 1):
            for col in range(2, 24):
                ws.cell(row=row, column=col).value = None

        group_row = table_start
        header_row = table_start + 1
        data_start = table_start + 2
        groups = (
            (2, 7, '작업장 정보'),
            (8, 14, '작업 기본정보'),
            (15, 17, '검사대상물'),
            (18, 20, '최대사용선원'),
            (21, 22, '방사선발생장치'),
            (23, 23, '비고'),
        )
        headers = (
            '년월', '업체코드', '검사업체명', '작업장코드', '작업장명',
            '일일최대량', '작업일', '작업장소', '작업조장', '작업자',
            '시작', '종료', '작업수량', '이름', '재질', '최대두께',
            '선원', '방사능량', '수량', '최대전압/전류', '발생장치수량', '비고',
        )
        dark_fill = PatternFill('solid', fgColor='BFBFBF')
        orange_fill = PatternFill('solid', fgColor='F4B183')
        yellow_fill = PatternFill('solid', fgColor='FFF200')
        header_font = Font(name='맑은 고딕', size=7, bold=True)
        body_font = Font(name='맑은 고딕', size=7)
        center = Alignment(
            horizontal='center', vertical='center', wrap_text=True,
            shrink_to_fit=True,
        )

        for start_col, end_col, label in groups:
            if start_col < end_col:
                ws.merge_cells(
                    start_row=group_row, start_column=start_col,
                    end_row=group_row, end_column=end_col,
                )
            cell = ws.cell(group_row, start_col, label)
            cell.fill = dark_fill
            cell.font = header_font
            cell.alignment = center
        for offset, label in enumerate(headers, start=2):
            cell = ws.cell(header_row, offset, label)
            cell.fill = orange_fill if offset <= 7 else yellow_fill
            cell.font = header_font
            cell.alignment = center

        ws.row_dimensions[group_row].height = 20
        ws.row_dimensions[header_row].height = 34
        if not records:
            ws.merge_cells(
                start_row=data_start, start_column=2,
                end_row=data_start, end_column=23,
            )
            ws.cell(data_start, 2).value = (
                f'{target_ym} 발주자 보고자료 없음'
            )
            ws.cell(data_start, 2).alignment = center
            ws.cell(data_start, 2).font = body_font
            records_end = data_start
        else:
            for row_offset, record in enumerate(records):
                target_row = data_start + row_offset
                for col_offset, value in enumerate(record, start=2):
                    cell = ws.cell(target_row, col_offset, value)
                    cell.font = body_font
                    cell.alignment = center
                ws.row_dimensions[target_row].height = 24
            records_end = data_start + len(records) - 1

        for row in range(group_row, records_end + 1):
            for col in range(2, 24):
                ws.cell(row, col).border = self.border_thin

    @staticmethod
    def _load_owner_report_workbook(source_path):
        """Load an owner report workbook supporting both .xlsx and legacy .xls."""
        ext = os.path.splitext(str(source_path))[1].lower()
        if ext != '.xls':
            return openpyxl.load_workbook(source_path, data_only=True)

        # Some portals save OOXML content with a .xls extension.
        try:
            with open(source_path, 'rb') as fh:
                raw = fh.read()
            if raw[:4] == b'PK\x03\x04':
                from io import BytesIO
                return openpyxl.load_workbook(BytesIO(raw), data_only=True)
        except OSError:
            pass

        # Fast path: use xlrd if installed in the runtime environment.
        try:
            import xlrd  # type: ignore

            xls_wb = xlrd.open_workbook(source_path)
            mem_wb = openpyxl.Workbook()
            first = True
            for sheet_name in xls_wb.sheet_names():
                xls_ws = xls_wb.sheet_by_name(sheet_name)
                if first:
                    dst_ws = mem_wb.active
                    dst_ws.title = sheet_name
                    first = False
                else:
                    dst_ws = mem_wb.create_sheet(title=sheet_name)
                for r in range(xls_ws.nrows):
                    for c in range(xls_ws.ncols):
                        cell = xls_ws.cell(r, c)
                        val = cell.value
                        if cell.ctype == xlrd.XL_CELL_DATE:
                            try:
                                val = xlrd.xldate_as_datetime(val, xls_wb.datemode)
                            except Exception:
                                pass
                        elif cell.ctype == xlrd.XL_CELL_NUMBER and isinstance(val, float) and val.is_integer():
                            val = int(val)
                        dst_ws.cell(r + 1, c + 1, val)
            return mem_wb
        except ImportError:
            pass

        # Fallback on Windows: convert .xls -> temporary .xlsx via Excel COM.
        import tempfile
        com_initialized = False
        excel = None
        wb_com = None
        try:
            import pythoncom  # type: ignore
            import win32com.client  # type: ignore

            try:
                pythoncom.CoInitialize()
                com_initialized = True
            except Exception:
                pass

            with tempfile.TemporaryDirectory(prefix='central_owner_xls_') as tmp_dir:
                tmp_xlsx = os.path.join(tmp_dir, 'owner_converted.xlsx')
                excel = win32com.client.DispatchEx('Excel.Application')
                excel.Visible = False
                excel.DisplayAlerts = False
                wb_com = excel.Workbooks.Open(os.path.abspath(source_path), ReadOnly=True)
                wb_com.SaveAs(tmp_xlsx, FileFormat=51)
                wb_com.Close(False)
                wb_com = None
                excel.Quit()
                excel = None
                # Load fully into memory before temp directory is removed.
                with open(tmp_xlsx, 'rb') as fh:
                    from io import BytesIO
                    return openpyxl.load_workbook(BytesIO(fh.read()), data_only=True)
        except Exception as exc:
            raise ValueError(
                f"'.xls' 발주자 보고 파일을 읽을 수 없습니다: {exc}\n"
                "Excel에서 '.xlsx' 형식으로 다시 저장한 뒤 선택해 주세요."
            ) from exc
        finally:
            if wb_com is not None:
                try:
                    wb_com.Close(False)
                except Exception:
                    pass
            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    def _populate_owner_report_rows(self, ws, target_ym, source_path):
        """Fill the existing section 6.0 template rows from a selected file."""
        title_row = None
        for row in range(1, ws.max_row + 1):
            row_text = ''.join(
                str(ws.cell(row=row, column=col).value or '')
                for col in range(1, min(ws.max_column, 30) + 1)
            ).replace(' ', '')
            if '1.방사선안전관리통합정보망보고자료' in row_text:
                title_row = row
                break

        if title_row is None:
            if not source_path:
                return
            raise ValueError('템플릿에서 6.0 발주자 보고자료 표를 찾을 수 없습니다.')

        # The edited V70 template uses rows 630~632 for its three-level header.
        data_start = title_row + 4
        data_end = data_start + 42
        for row in range(data_start, data_end + 1):
            for col in range(2, 24):
                cell = ws.cell(row=row, column=col)
                if type(cell).__name__ != 'MergedCell':
                    cell.value = None

        records = []
        if source_path and os.path.exists(source_path):
            source_wb = self._load_owner_report_workbook(source_path)
            try:
                if '작업정보' not in source_wb.sheetnames:
                    raise ValueError("발주자 보고 파일에 '작업정보' 시트가 없습니다.")
                source_ws = source_wb['작업정보']
                target_key = str(target_ym or '').replace('-', '')[:6]

                def _text(value):
                    if value is None:
                        return ''
                    if isinstance(value, (datetime.datetime, datetime.date)):
                        return value.strftime('%Y-%m-%d')
                    # Owner-export cells may contain hard line breaks as well
                    # as Excel's visual wrapping (e.g. "2946\n81" or a name).
                    # Join those fragments without corrupting codes/leading
                    # zeros or removing meaningful ordinary spaces.
                    return re.sub(r'[\r\n\u2028\u2029]+', '', str(value)).strip()

                def _clock(value):
                    text = _text(value)
                    digits = re.sub(r'\D', '', text)
                    if len(digits) == 3:
                        digits = f'0{digits}'
                    if len(digits) == 4:
                        return f'{digits[:2]}:{digits[2:]}'
                    return text

                for row in range(5, source_ws.max_row + 1):
                    year_month = _text(source_ws.cell(row, 1).value)
                    year_month = re.sub(r'\D', '', year_month)[:6]
                    if target_key and year_month != target_key:
                        continue
                    if self._owner_is_empty_work_row(source_ws, row):
                        continue
                    record = [
                    year_month,
                    _text(source_ws.cell(row, 2).value),
                    _text(source_ws.cell(row, 3).value),
                    _text(source_ws.cell(row, 4).value),
                    self._clean_owner_workplace_name(_text(source_ws.cell(row, 5).value)),
                    source_ws.cell(row, 6).value or '',
                    _text(source_ws.cell(row, 7).value),
                    _text(source_ws.cell(row, 8).value),
                    _text(source_ws.cell(row, 9).value),
                    *[
                        _text(source_ws.cell(row, col).value)
                        for col in range(10, 15)
                    ],
                    _clock(source_ws.cell(row, 15).value),
                    _clock(source_ws.cell(row, 16).value),
                    source_ws.cell(row, 17).value or '',
                    _text(source_ws.cell(row, 18).value),
                    _text(source_ws.cell(row, 19).value),
                    source_ws.cell(row, 20).value or '',
                    source_ws.cell(row, 25).value or '',
                    source_ws.cell(row, 26).value or '',
                    ]
                    records.append(record)
            finally:
                source_wb.close()
        elif source_path and not os.path.exists(source_path):
            raise FileNotFoundError(f'발주자 보고 파일을 찾을 수 없습니다: {source_path}')

        if not records:
            ws.merge_cells(start_row=data_start, start_column=2, end_row=data_start, end_column=23)
            no_data_cell = ws.cell(row=data_start, column=2)
            no_data_cell.value = "해당 월 방사선 투과검사(RT) 실적 없음"
            no_data_cell.font = Font(name='맑은 고딕', size=10, bold=False)
            no_data_cell.alignment = Alignment(horizontal='center', vertical='center')
            return

        body_font = Font(name='맑은 고딕', size=6)
        body_alignment = Alignment(
            horizontal='center', vertical='center',
            wrap_text=False, shrink_to_fit=True,
        )
        workplace_font = Font(name='맑은 고딕', size=6)
        workplace_alignment = Alignment(
            horizontal='center', vertical='center',
            wrap_text=True, shrink_to_fit=False,
        )

        # 첫 페이지의 43개 행을 모두 사용한 뒤에는 16페이지 형식의
        # 공통 문서 헤더와 6.0 표 헤더를 반복한 계속 페이지를 만든다.
        import copy
        from openpyxl.worksheet.pagebreak import Break

        rows_per_page = data_end - data_start + 1
        default_height = float(ws.sheet_format.defaultRowHeight or 15)
        reference_heights = [
            float(ws.row_dimensions[row].height
                  if ws.row_dimensions[row].height is not None else default_height)
            for row in range(data_start, data_end + 1)
        ]
        page_capacity = sum(reference_heights)
        base_height = reference_heights[0]
        pages = [[]]
        used_height = 0.0
        for record in records:
            height = self._owner_workplace_row_height(ws, record[4], base_height)
            if height > page_capacity:
                raise ValueError('작업장명이 발주자 보고 한 페이지 높이를 초과합니다.')
            if pages[-1] and (len(pages[-1]) >= rows_per_page
                              or used_height + height > page_capacity + 0.01):
                pages.append([])
                used_height = 0.0
            pages[-1].append((record, height))
            used_height += height
        page_starts = [data_start]
        insert_row = data_end + 1
        previous_breaks = sorted(
            int(brk.id) for brk in ws.row_breaks.brk if int(brk.id) < title_row
        )
        document_header_row = previous_breaks[-1] + 1 if previous_breaks else 1
        table_header_end = data_start - 1

        reference_styles = [
            copy.copy(ws.cell(data_start, col)._style)
            for col in range(1, ws.max_column + 1)
        ]
        for _ in pages[1:]:
            inserted_header_rows = self._insert_status_continuation(
                ws, insert_row, document_header_row, title_row, table_header_end
            )
            body_start = insert_row + inserted_header_rows
            self._insert_rows_safely(ws, body_start, rows_per_page)
            for image in list(ws._images):
                anchor = getattr(image, 'anchor', None)
                if not hasattr(anchor, '_from') or anchor._from.row + 1 < body_start:
                    continue
                anchor._from.row += rows_per_page
                if hasattr(anchor, 'to'):
                    anchor.to.row += rows_per_page
            for row in range(body_start, body_start + rows_per_page):
                ws.row_dimensions[row].height = base_height
                for col, style in enumerate(reference_styles, 1):
                    ws.cell(row, col)._style = copy.copy(style)
            if not any(int(brk.id) == insert_row - 1 for brk in ws.row_breaks.brk):
                ws.row_breaks.append(Break(id=insert_row - 1))

            page_starts.append(body_start)
            insert_row = body_start + rows_per_page

        thin_bottom = Side(style='thin')
        for body_start, page in zip(page_starts, pages):
            for offset, (record, height) in enumerate(page):
                target_row = body_start + offset
                ws.row_dimensions[target_row].height = height
                ws.row_dimensions[target_row].hidden = False
                for column, value in enumerate(record, start=2):
                    cell = ws.cell(row=target_row, column=column)
                    cell.value = value
                    cell.font = workplace_font if column == 6 else body_font
                    cell.alignment = workplace_alignment if column == 6 else body_alignment
            # Growing name cells must consume the reserved blank table rows,
            # rather than push the following photos off their printed page.
            remaining_height = max(0.0, page_capacity - sum(h for _, h in page))
            empty_count = rows_per_page - len(page)
            visible_empty = min(empty_count, max(1, int(remaining_height / base_height)))
            for offset in range(empty_count):
                dimension = ws.row_dimensions[body_start + len(page) + offset]
                visible = offset < visible_empty and remaining_height > 0
                dimension.height = remaining_height / visible_empty if visible else 0
                dimension.hidden = not visible
            visible_rows = len(page) + (visible_empty if remaining_height > 0 else 0)
            for row_offset in range(visible_rows):
                row_idx = body_start + row_offset
                for column in range(2, 24):
                    cell = ws.cell(row=row_idx, column=column)
                    border = copy.copy(cell.border) if cell.border else Border()
                    if border.bottom is None or border.bottom.style is None:
                        border.bottom = copy.copy(thin_bottom)
                        cell.border = border

    @staticmethod
    def _owner_is_empty_work_row(ws, row):
        """Skip explicit zero-work placeholders, preserving genuine zero-qty work."""
        quantity = ws.cell(row, 17).value
        try:
            if quantity in (None, '') or float(str(quantity).replace(',', '')) != 0:
                return False
        except (TypeError, ValueError):
            return False
        # Registration metadata and 09:00~09:01 timestamps alone do not prove
        # inspection work. A crew, subject, source/device or remark does.
        for col in list(range(9, 15)) + list(range(18, 33)):
            value = ws.cell(row, col).value
            if value is None or str(value).strip() == '':
                continue
            try:
                if float(str(value).strip()) == 0:
                    continue
            except (TypeError, ValueError):
                pass
            return False
        return True

    @staticmethod
    def _clean_owner_workplace_name(name):
        """Strip the redundant branch prefix when location details follow."""
        text = str(name or '').strip()
        cleaned = re.sub(
            r'^한국지역난방공사\s*(?:중앙지사)?\s*',
            '',
            text,
        ).strip()
        return cleaned or text

    @staticmethod
    def _owner_workplace_row_height(ws, name, minimum):
        """Fit the complete workplace name at 6pt without shrinking its text."""
        import unicodedata

        digit_width = 8 if ws.parent._fonts[0].name in (
            '굴림', '굴림체', '돋움', '돋움체') else 7
        width_pixels = float(ws.column_dimensions['F'].width or 8.43) * digit_width + 5
        available_points = max(6.0, (width_pixels - 8) * 0.75)
        lines, used = 1, 0.0
        for character in str(name or ''):
            width = 6.0 if unicodedata.east_asian_width(character) in ('W', 'F') else 3.5
            if used and used + width > available_points:
                lines += 1
                used = 0.0
            used += width
        return max(float(minimum), lines * 9.5 + 3.0)

    def _trim_trailing_blank_print_pages(self, ws):
        """Stop printing at the end of the last page containing real content.

        Dynamic row insertion expands the template print area and shifts every
        manual row break.  That can leave a fully blank page after the last
        report page.  Preserve the complete last content page by trimming at
        its following page boundary, rather than at the last populated row.
        """
        from openpyxl.utils import get_column_letter

        print_area_text = str(ws.print_area or '')
        area_match = re.search(
            r"\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)",
            print_area_text.upper(),
        )
        if not area_match:
            return

        min_col_name, min_row_text, max_col_name, max_row_text = area_match.groups()
        from openpyxl.utils.cell import column_index_from_string
        min_col = column_index_from_string(min_col_name)
        max_col = column_index_from_string(max_col_name)
        min_row = int(min_row_text)
        print_max_row = int(max_row_text)

        last_content_row = min_row
        for row in ws.iter_rows(
            min_row=min_row,
            max_row=min(print_max_row, ws.max_row),
            min_col=min_col,
            max_col=max_col,
        ):
            if any(cell.value not in (None, '') for cell in row):
                last_content_row = row[0].row

        # Images are printable content even though their cells have no value.
        for image in list(ws._images):
            anchor = getattr(image, 'anchor', None)
            if not hasattr(anchor, '_from'):
                continue
            image_bottom = anchor._from.row + 1
            if hasattr(anchor, 'to'):
                image_bottom = max(image_bottom, anchor.to.row + 1)
            last_content_row = max(last_content_row, image_bottom)

        later_breaks = sorted(
            int(brk.id)
            for brk in ws.row_breaks.brk
            if last_content_row <= int(brk.id) < print_max_row
        )
        if not later_breaks:
            return

        last_page_end = later_breaks[0]
        ws.print_area = (
            f"${get_column_letter(min_col)}${min_row}:"
            f"${get_column_letter(max_col)}${last_page_end}"
        )
        ws.row_breaks.brk = [
            brk for brk in ws.row_breaks.brk
            if int(brk.id) < last_page_end
        ]

    def _trim_print_area_after_last_process_photo(self, ws):
        """End printing at the page containing the final real process photo."""
        photo_rows = []
        for image in list(ws._images):
            anchor = getattr(image, 'anchor', None)
            if not hasattr(anchor, '_from'):
                continue
            image_row = anchor._from.row + 1
            # Report-header logos are about 186x50. Process photos are much
            # larger; ignore logos on otherwise empty trailing photo pages.
            if (
                image_row >= 600
                and float(image.width) >= 250
                and float(image.height) >= 150
            ):
                photo_rows.append(image_row)
        if not photo_rows:
            return

        last_photo_row = max(photo_rows)
        following_breaks = sorted(
            int(brk.id) for brk in ws.row_breaks.brk
            if int(brk.id) >= last_photo_row
        )
        if following_breaks:
            self._set_print_area_end(ws, following_breaks[0])

    def _rt_reject_defects(self, row):
        """Return reject-grade RT defects encoded as ``grade/defect``."""
        method = str(row.get('검사방법', '') or '').strip().upper()
        if method != 'RT':
            return []

        from site_apps.central.src.rt_quantities import film_grades, effective_section_info
        return [defect for grade, defect in (film_grades(effective_section_info(row)) or []) if grade >= 3]

    def _rt_film_count(self, row):
        """Count populated RT section-info slots (one slot equals one film)."""
        if str(row.get('검사방법', '') or '').strip().upper() != 'RT':
            return 0
        from site_apps.central.src.rt_quantities import rt_counts
        return rt_counts(row)[0]

    def _compact_weld_defect_table_gap(self, ws, page_start, title_row, total_row):
        """Reclaim the empty gap between defect tables before splitting data."""
        scale = self._report_height_scale(ws)
        if self._first_print_overflow_row(ws, page_start,
                                          total_row - page_start + 1, scale) is None:
            return
        previous_total = next((r for r in range(title_row - 1, page_start - 1, -1)
                               if any(str(ws.cell(r, c).value or '').strip().upper()
                                      in ('TOTAL', '총계')
                                      for c in range(1, min(ws.max_column, 23) + 1))), None)
        if previous_total is None:
            return
        merged_rows = {r for m in ws.merged_cells.ranges
                       for r in range(m.min_row, m.max_row + 1)}
        image_rows = {im.anchor._from.row + 1 for im in ws._images
                      if hasattr(im.anchor, '_from')}
        for r in range(previous_total + 1, title_row):
            if r in merged_rows or r in image_rows:
                continue
            if not any(ws.cell(r, c).value is not None
                       for c in range(1, min(ws.max_column, 23) + 1)):
                dim = ws.row_dimensions[r]
                if not dim.hidden:
                    dim.height = min(float(dim.height or ws.sheet_format.defaultRowHeight or 15), 2.0)

    def _write_welder_defect_summary(
        self, ws, defects_by_welder, retests_by_welder,
        original_counts_by_welder, section_number, section_keywords,
    ):
        """Write welder defects and A/B quantities into a 4.0 summary table."""
        import copy

        has_data = bool(defects_by_welder or retests_by_welder or original_counts_by_welder)

        title_row = None
        for row in range(1, ws.max_row + 1):
            row_text = ''.join(
                str(ws.cell(row=row, column=col).value or '')
                for col in range(1, ws.max_column + 1)
            ).replace(' ', '').replace('\n', '').upper()
            if (
                section_number in row_text
                and ('불량률' in row_text or '불량율' in row_text)
                and any(keyword.upper() in row_text for keyword in section_keywords)
            ):
                title_row = row
                break
        if title_row is None:
            return

        header_row = None
        id_col = None
        name_col = None
        header_cols = {}
        for row in range(title_row + 1, min(title_row + 7, ws.max_row + 1)):
            current_headers = {}
            current_id_col = None
            for col in range(1, ws.max_column + 1):
                label = ''.join(str(ws.cell(row=row, column=col).value or '').split()).upper()
                if label == 'ID':
                    current_id_col = col
                elif label:
                    current_headers.setdefault(label, col)
            if current_id_col is not None:
                header_row = row
                id_col = current_id_col
                header_cols = current_headers
                name_col = current_headers.get('성명')
                break
        if header_row is None:
            return

        retest_col = None
        original_count_col = None
        defect_rate_col = None
        for row in range(title_row + 1, header_row + 1):
            for col in range(1, ws.max_column + 1):
                label = ''.join(str(ws.cell(row=row, column=col).value or '').split()).upper()
                if '불합격부' in label and '재검사' in label and ('(A)' in label or label.endswith('A')):
                    retest_col = col
                if '당초' in label and '시공물량' in label and ('(B)' in label or label.endswith('B')):
                    original_count_col = col
                if '용접' in label and ('불량률' in label or '불량율' in label) and 'A/B' in label:
                    defect_rate_col = col
            if (
                retest_col is not None
                and original_count_col is not None
                and defect_rate_col is not None
            ):
                break

        total_row = None
        for row in range(header_row + 1, min(header_row + 30, ws.max_row + 1)):
            values = [
                ''.join(str(ws.cell(row=row, column=col).value or '').split()).upper()
                for col in range(1, ws.max_column + 1)
            ]
            if any(value in {'TOTAL', '총계'} for value in values):
                total_row = row
                break
        if total_row is None:
            return

        # Older templates reserve D:O for defects and omit CL. Expand only
        # this table, keeping the worksheet's other sections in place.
        if 'CL' not in header_cols and retest_col == 16 and original_count_col == 17 and defect_rate_col == 18:
            start_row = title_row + 1
            merges = [m for m in list(ws.merged_cells.ranges)
                      if m.min_row >= start_row and m.max_row <= total_row
                      and m.max_col >= 16 and m.min_col <= 18]
            # The defect-group heading ends immediately before column P.
            group_merges = [m for m in list(ws.merged_cells.ranges)
                            if m.min_row >= start_row and m.max_row < header_row
                            and m.min_col == 4 and m.max_col == 15]
            for merged in merges + group_merges:
                ws.unmerge_cells(str(merged))
            for row in range(start_row, total_row + 1):
                for col in range(18, 15, -1):
                    source = ws.cell(row, col)
                    target = ws.cell(row, col + 1)
                    target.value = source.value
                    target._style = copy.copy(source._style)
                cell = ws.cell(row, 16)
                cell.value = None
                cell._style = copy.copy(ws.cell(row, 15)._style)
            for merged in merges:
                ws.merge_cells(start_row=merged.min_row, end_row=merged.max_row,
                               start_column=merged.min_col + 1, end_column=merged.max_col + 1)
            for merged in group_merges:
                ws.merge_cells(start_row=merged.min_row, end_row=merged.max_row,
                               start_column=merged.min_col, end_column=16)
            ws.cell(header_row, 16).value = 'CL'
            header_cols['CL'] = 16
            retest_col, original_count_col, defect_rate_col = 17, 18, 19

        data_start = header_row + 1
        available_rows = total_row - data_start

        if not has_data:
            msg_col = id_col if id_col else 2
            ws.cell(row=data_start, column=msg_col).value = "당월 해당 공정 검사 실적 없음"

            # 텍스트 가운데 정렬 및 글씨 진하게
            align = copy.copy(ws.cell(row=data_start, column=msg_col).alignment)
            if align:
                align.horizontal = 'center'
                align.vertical = 'center'
                ws.cell(row=data_start, column=msg_col).alignment = align

            font = copy.copy(ws.cell(row=data_start, column=msg_col).font)
            if font:
                font.bold = True
                ws.cell(row=data_start, column=msg_col).font = font

            table_end_col = defect_rate_col
            if not table_end_col:
                cols_found = [c for c in [retest_col, original_count_col] + list(header_cols.values()) if c is not None]
                table_end_col = max(cols_found) if cols_found else msg_col + 15

            try:
                ws.merge_cells(start_row=data_start, start_column=msg_col, end_row=data_start, end_column=table_end_col)
            except Exception:
                pass

            # 나머지 데이터 행들은 값을 지우거나 그대로 둠
            for r in range(data_start + 1, total_row):
                ws.row_dimensions[r].hidden = True

            # 합계(TOTAL) 행의 값들도 초기화
            for c in range(msg_col + 1, ws.max_column + 1):
                cell = ws.cell(row=total_row, column=c)
                if type(cell).__name__ != 'MergedCell':
                    cell.value = ""

            return
        welders = sorted(
            set(defects_by_welder)
            | set(retests_by_welder)
            | set(original_counts_by_welder)
        )
        if len(welders) > available_rows:
            extra_rows = len(welders) - available_rows
            style_source_row = max(data_start, total_row - 1)
            row_style = [
                copy.copy(ws.cell(row=style_source_row, column=col)._style)
                for col in range(1, ws.max_column + 1)
            ]
            row_height = ws.row_dimensions[style_source_row].height
            self._insert_rows_safely(ws, total_row, extra_rows)
            for row in range(total_row, total_row + extra_rows):
                ws.row_dimensions[row].height = row_height
                for col, style in enumerate(row_style, start=1):
                    ws.cell(row=row, column=col)._style = copy.copy(style)
            total_row += extra_rows

        # Keep every body row on the same horizontal grid as the first data
        # row.  The template contains a few reserved blank rows whose styles
        # are correct but whose wide summary cells are not merged.  Likewise,
        # dynamically inserted rows above TOTAL only receive cell styles.
        # Those rows expose internal column lines (notably on the RT summary
        # page) unless the reference row's single-row merges are replicated.
        table_end_col = defect_rate_col or original_count_col or retest_col
        if table_end_col is None:
            table_end_col = max(
                [id_col] + [col for col in header_cols.values() if col is not None]
            )
        for merged in list(ws.merged_cells.ranges):
            if (
                merged.min_row == data_start
                and merged.max_row == data_start
                and merged.min_col <= table_end_col <= merged.max_col
            ):
                table_end_col = merged.max_col
                break
        reference_merges = [
            (merged.min_col, merged.max_col)
            for merged in list(ws.merged_cells.ranges)
            if merged.min_row == data_start
            and merged.max_row == data_start
            and merged.min_col >= id_col
            and merged.max_col <= table_end_col
        ]
        if reference_merges:
            for row in range(data_start + 1, total_row):
                for merged in list(ws.merged_cells.ranges):
                    if (
                        merged.min_row == row
                        and merged.max_row == row
                        and merged.min_col <= table_end_col
                        and merged.max_col >= id_col
                    ):
                        ws.unmerge_cells(str(merged))
                for min_col, max_col in reference_merges:
                    ws.merge_cells(
                        start_row=row,
                        start_column=min_col,
                        end_row=row,
                        end_column=max_col,
                    )

        defect_columns = {
            code: header_cols.get(code)
            for code in {code for counts in defects_by_welder.values() for code in counts}
            if header_cols.get(code) is not None
        }
        for offset, welder in enumerate(welders):
            row = data_start + offset
            id_cell = ws.cell(row=row, column=id_col)
            id_cell.value = welder
            id_alignment = copy.copy(id_cell.alignment)
            id_alignment.wrap_text = False
            id_alignment.shrink_to_fit = True
            id_cell.alignment = id_alignment
            if name_col is not None:
                name_cell = ws.cell(row=row, column=name_col)
                name_cell.value = self.WELDER_NAMES.get(welder.upper(), '')
                name_alignment = copy.copy(name_cell.alignment)
                name_alignment.wrap_text = False
                name_alignment.shrink_to_fit = True
                name_cell.alignment = name_alignment
            for code, count in defects_by_welder.get(welder, {}).items():
                col = defect_columns.get(code)
                if col is not None:
                    ws.cell(row=row, column=col).value = count
            if retest_col is not None:
                ws.cell(row=row, column=retest_col).value = retests_by_welder.get(welder, 0)
            if original_count_col is not None:
                ws.cell(row=row, column=original_count_col).value = (
                    original_counts_by_welder.get(welder, 0)
                )
            if defect_rate_col is not None:
                retest_qty = retests_by_welder.get(welder, 0)
                original_qty = original_counts_by_welder.get(welder, 0)
                rate_cell = ws.cell(row=row, column=defect_rate_col)
                rate_cell.value = retest_qty / original_qty if original_qty else None
                rate_cell.number_format = '0.00%'
                rate_alignment = copy.copy(rate_cell.alignment)
                rate_alignment.wrap_text = False
                rate_alignment.shrink_to_fit = True
                rate_cell.alignment = rate_alignment

        # Recalculate each defect total without disturbing the other totals.
        for col in set(defect_columns.values()):
            total = sum(
                self._safe_float(ws.cell(row=row, column=col).value)
                for row in range(data_start, data_start + len(welders))
            )
            cell = ws.cell(row=total_row, column=col)
            if type(cell).__name__ != 'MergedCell':
                cell.value = int(total)
        if retest_col is not None:
            cell = ws.cell(row=total_row, column=retest_col)
            if type(cell).__name__ != 'MergedCell':
                cell.value = sum(retests_by_welder.values())
        if original_count_col is not None:
            cell = ws.cell(row=total_row, column=original_count_col)
            if type(cell).__name__ != 'MergedCell':
                cell.value = sum(original_counts_by_welder.values())
        if defect_rate_col is not None:

            total_retest = sum(retests_by_welder.values())
            total_original = sum(original_counts_by_welder.values())
            total_rate_cell = ws.cell(row=total_row, column=defect_rate_col)
            total_rate_cell.value = (
                total_retest / total_original if total_original else None
            )
            total_rate_cell.number_format = '0.00%'
            total_rate_alignment = copy.copy(total_rate_cell.alignment)
            total_rate_alignment.wrap_text = False
            total_rate_alignment.shrink_to_fit = True
            total_rate_cell.alignment = total_rate_alignment

        # 14페이지 표가 인쇄 가능 영역을 넘으면 공통 문서 헤더와 현재
        # 용접사 현황 표 헤더를 반복하고 남은 행을 새 페이지에서 계속한다.
        remaining_start = data_start
        remaining_count = len(welders) + 1  # TOTAL 행 포함
        previous_breaks = sorted(
            int(brk.id) for brk in ws.row_breaks.brk
            if int(brk.id) < title_row
        )
        document_header_row = previous_breaks[-1] + 1 if previous_breaks else 1
        self._compact_weld_defect_table_gap(ws, document_header_row, title_row, total_row)
        while remaining_count > 0:
            continuation_row = self._first_print_overflow_row(
                ws, remaining_start, remaining_count, self._report_height_scale(ws)
            )
            if continuation_row is None:
                break
            rows_on_page = continuation_row - remaining_start
            if rows_on_page <= 0:
                break
            inserted_rows = self._insert_status_continuation(
                ws, continuation_row, document_header_row, title_row, header_row
            )
            from openpyxl.worksheet.pagebreak import Break
            if not any(
                int(brk.id) == continuation_row - 1 for brk in ws.row_breaks.brk
            ):
                ws.row_breaks.append(Break(id=continuation_row - 1))
            remaining_count -= rows_on_page
            remaining_start = continuation_row + inserted_rows

    def _populate_process_photo_pages(self, ws, process_photos, doc_num='01'):
        """Place registered process photos into the template's 7.0 pages."""
        from openpyxl.drawing.image import Image as XLImage
        from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
        from openpyxl.utils import get_column_letter
        from openpyxl.utils.units import pixels_to_EMU, points_to_pixels
        from PIL import Image as PILImage, ImageOps
        from io import BytesIO
        import copy

        from site_apps.central.src.data_sync import resolve_process_photo_path
        valid_photos = []
        for photo in process_photos:
            process = str(photo.get('process', '')).strip().upper()
            if process not in {'PAUT', 'MT', 'RT', 'PT'}:
                continue
            stored_path = str(photo.get('file_path', '') or '')
            image_path = resolve_process_photo_path(self._history_path, stored_path)
            if image_path:
                saved_photo = dict(photo)
                saved_photo['_resolved_path'] = image_path
                valid_photos.append(saved_photo)
        process_photos = valid_photos

        original_photo_starts = []
        for row in reversed(range(600, ws.max_row + 1)):
            row_text = ''.join(
                str(ws.cell(row=row, column=col).value or '')
                for col in range(1, ws.max_column + 1)
            ).replace(' ', '').upper()
            if 'SEC.16' in row_text:
                # Insert 21 blank rows to accommodate the 3rd row of photos (6 photos total)
                self._insert_rows_safely(ws, row + 42, 21)

        for row in range(600, ws.max_row + 1):
            row_text = ''.join(
                str(ws.cell(row=row, column=col).value or '')
                for col in range(1, ws.max_column + 1)
            ).replace(' ', '').upper()
            if 'SEC.16' in row_text and any(
                keyword in row_text
                for keyword in ('PAUT', '위상배열초음파탐상검사', '자분탐상검사', '방사선투과검사')
            ):
                original_photo_starts.append(row - 9)
        photo_section_start = min(original_photo_starts) if original_photo_starts else None

        # Page 16 is the layout reference for the following photo pages.  Read
        # its seven actual header-row heights instead of estimating them; the
        # photo header otherwise renders noticeably shorter on page 17.
        reference_header_heights = None
        reference_section_font = None
        reference_header_start = None
        reference_logo_anchor = None
        reference_logo_size = None
        if photo_section_start is not None:
            for row in range(1, photo_section_start):
                row_text = ''.join(
                    str(ws.cell(row=row, column=col).value or '')
                    for col in range(1, min(ws.max_column, 23) + 1)
                ).replace(' ', '').replace('\n', '')
                if '6.0\ubc29\uc0ac\uc120\uc548\uc804\uad00\ub9ac\ud1b5\ud569\uc815\ubcf4\ub9dd\ubcf4\uace0\uc790\ub8cc' in row_text:
                    reference_start = max(1, row - 4)
                    reference_header_start = reference_start
                    reference_header_heights = [
                        ws.row_dimensions[reference_start + offset].height
                        for offset in range(7)
                    ]
                    reference_section_font = copy.copy(
                        ws.cell(row=row, column=6).font
                    )
                    for image in list(ws._images):
                        anchor = getattr(image, 'anchor', None)
                        if not hasattr(anchor, '_from'):
                            continue
                        image_row = anchor._from.row + 1
                        image_col = anchor._from.col + 1
                        if (
                            reference_start <= image_row <= reference_start + 6
                            and image_col <= 5
                        ):
                            reference_logo_anchor = copy.deepcopy(anchor)
                            reference_logo_size = (image.width, image.height)
                            break

        # Clear every legacy/template photo before building the photo pages.
        # Some templates contain a sample ledger image anchored above the
        # detected SEC.16 page start (the May 11 image is anchored at row 718,
        # while the generated PAUT slots begin at row 739).  Page-local cleanup
        # therefore cannot see it.  Rows from 600 onward are the report's photo
        # section; retain only the small document-header logos there.
        cleaned_images = []
        for existing_image in list(ws._images):
            anchor = getattr(existing_image, 'anchor', None)
            if not hasattr(anchor, '_from'):
                cleaned_images.append(existing_image)
                continue
            image_row = anchor._from.row + 1
            image_col = anchor._from.col + 1
            is_header_logo = (
                image_col <= 5
                and float(existing_image.width) < 300
                and float(existing_image.height) < 100
            )
            if image_row >= 600 and not is_header_logo:
                continue
            cleaned_images.append(existing_image)
        ws._images = cleaned_images

        pt_photos = [
            photo for photo in process_photos
            if str(photo.get('process', '')).upper() == 'PT'
        ]
        pt_title_row = self._ensure_pt_photo_page(ws) if pt_photos else None

        processes_with_photos = {
            str(photo.get('process', '')).upper()
            for photo in process_photos
            if str(photo.get('process', '')).strip()
        }
        self._remove_empty_process_photo_pages(ws, processes_with_photos)

        page_layouts = {}
        title_keywords = {
            'PAUT': ('위상배열초음파탐상검사', 'PAUT'),
            'MT': ('자분탐상검사',),
            'RT': ('방사선투과검사',),
        }
        for row in range(1, ws.max_row + 1):
            row_text = ''.join(
                str(ws.cell(row=row, column=col).value or '')
                for col in range(1, ws.max_column + 1)
            ).replace(' ', '').upper()
            for process, keywords in title_keywords.items():
                if process in page_layouts:
                    continue
                if (
                    any(keyword.upper() in row_text for keyword in keywords)
                    and 'SEC.16' in row_text
                    and row > 600
                ):
                    page_layouts[process] = {
                        'title_row': row,
                        'image_row': row + 2,
                        'caption_row': row + 42,
                    }
        # PT may have moved when empty PAUT/MT/RT pages were deleted, so locate
        # every surviving page again instead of relying on its old row number.
        if pt_title_row is not None:
            for row in range(600, ws.max_row + 1):
                row_text = ''.join(
                    str(ws.cell(row=row, column=col).value or '')
                    for col in range(1, ws.max_column + 1)
                ).replace(' ', '').upper()
                if '침투탐상검사' in row_text and 'SEC.16' in row_text:
                    page_layouts['PT'] = {
                        'title_row': row,
                        'image_row': row + 2,
                        'caption_row': row + 42,
                    }
                    break
        descriptions = {
            'PAUT': '위상배열초음파탐상검사',
            'MT': '자분탐상검사',
            'RT': '방사선투과검사',
            'PT': '침투탐상검사',
        }
        # Add as many continuation pages as needed (six photos per page).
        for process, layout in sorted(
            page_layouts.items(), key=lambda item: item[1]['title_row'], reverse=True
        ):
            photo_count = sum(
                1 for photo in process_photos
                if str(photo.get('process', '')).upper() == process
            )
            extra_pages = max(0, (photo_count - 1) // 6)
            source_title_row = layout['title_row']
            for continuation_no in range(2, extra_pages + 2):
                source_title_row = self._clone_process_photo_page(
                    ws, source_title_row, continuation_no
                )

        # All lower pages may have moved. Rebuild a list of every surviving
        # base/continuation page for each process.
        layouts_by_process = {process: [] for process in descriptions}
        scan_keywords = dict(title_keywords)
        scan_keywords['PT'] = ('침투탐상검사',)
        for row in range(600, ws.max_row + 1):
            row_text = ''.join(
                str(ws.cell(row=row, column=col).value or '')
                for col in range(1, ws.max_column + 1)
            ).replace(' ', '').upper()
            if 'SEC.16' not in row_text:
                continue
            for process, keywords in scan_keywords.items():
                if any(keyword.upper() in row_text for keyword in keywords):
                    layouts_by_process[process].append({
                        'title_row': row,
                        'image_row': row + 2,
                    })
                    break

        for process, process_layouts in layouts_by_process.items():
            photos = sorted(
                [
                    photo for photo in process_photos
                    if str(photo.get('process', '')).upper() == process
                ],
                key=lambda photo: (
                    str(photo.get('date', '')), str(photo.get('joint_no', ''))
                ),
            )
            for page_index, layout in enumerate(process_layouts):
                page_photos = photos[page_index * 6:(page_index + 1) * 6]

                # Always clear photos embedded in the source template, even
                # when this particular page has no newly registered photos.
                # Previously the early continue below left old photo ledgers
                # (for example, the May 11 sample) on unused/extra pages.
                page_start_row = layout['title_row'] - 9
                page_end_row = page_start_row + 74
                kept_images = []
                for existing_image in list(ws._images):
                    anchor = getattr(existing_image, 'anchor', None)
                    if hasattr(anchor, '_from'):
                        r = anchor._from.row + 1
                        c = anchor._from.col + 1
                        is_header_logo = (
                            c <= 5
                            and float(existing_image.width) < 300
                            and float(existing_image.height) < 100
                        )
                        if page_start_row <= r <= page_end_row and not is_header_logo:
                            continue
                    kept_images.append(existing_image)
                ws._images = kept_images

                if not page_photos:
                    continue

                photo_slots = (
                    ('C', layout['image_row'], 3, 10, layout['image_row'] + 16),
                    ('L', layout['image_row'], 12, 20, layout['image_row'] + 16),
                    ('C', layout['image_row'] + 21, 3, 10, layout['image_row'] + 37),
                    ('L', layout['image_row'] + 21, 12, 20, layout['image_row'] + 37),
                    ('C', layout['image_row'] + 42, 3, 10, layout['image_row'] + 58),
                    ('L', layout['image_row'] + 42, 12, 20, layout['image_row'] + 58),
                )

                # C:J and L:T center the ledger inside the B:W document frame.
                # Their pixel widths differ by only three pixels; column K is
                # retained as a consistent gutter between the two photos.

                # dynamically copy styles from the second row to the third row
                import copy
                source_start = layout['image_row'] + 21
                source_end = layout['image_row'] + 37
                target_start = layout['image_row'] + 42

                for offset in range(source_end - source_start + 1):
                    source_r = source_start + offset
                    target_r = target_start + offset

                    # copy row height
                    if source_r in ws.row_dimensions:
                        ws.row_dimensions[target_r].height = ws.row_dimensions[source_r].height

                # Hide gaps so that the 3rd row fits exactly onto one A4 page
                for gap_offset in [18, 19, 20, 39, 40, 41, 61, 62]:
                    ws.row_dimensions[layout['image_row'] + gap_offset].hidden = True

                    for col in range(2, 24):
                        source_cell = ws.cell(row=source_r, column=col)
                        target_cell = ws.cell(row=target_r, column=col)

                        target_cell._style = copy.copy(source_cell._style)
                        target_cell.alignment = copy.copy(source_cell.alignment)
                        target_cell.number_format = source_cell.number_format
                        target_cell.font = copy.copy(source_cell.font)
                        target_cell.border = copy.copy(source_cell.border)
                        target_cell.fill = copy.copy(source_cell.fill)

                # Clone merge ranges for the 3rd row
                new_merges = []
                for merged in list(ws.merged_cells.ranges):
                    if merged.min_row >= source_start and merged.max_row <= source_end:
                        new_merges.append((
                            merged.min_row + 21, merged.min_col,
                            merged.max_row + 21, merged.max_col
                        ))

                for mr in new_merges:
                    ws.merge_cells(start_row=mr[0], start_column=mr[1], end_row=mr[2], end_column=mr[3])

                # Center the two-column ledger within the B:W document frame
                # and use part of the former bottom whitespace for the photos.
                # A modest four-percent height increase keeps the page inside
                # its existing manual boundary while improving vertical balance.
                default_row_height = ws.sheet_format.defaultRowHeight or 15
                for frame_offset in (0, 21, 42):
                    frame_start = layout['image_row'] + frame_offset
                    frame_end = frame_start + 15
                    for frame_row in range(frame_start, frame_end + 1):
                        current_height = (
                            ws.row_dimensions[frame_row].height
                            or default_row_height
                        )
                        ws.row_dimensions[frame_row].height = (
                            float(current_height) * 1.04
                        )

                for index, photo in enumerate(page_photos):
                    (
                        anchor_col, anchor_row, caption_col,
                        caption_end_col, caption_row,
                    ) = photo_slots[index]

                    # Size and center the photo inside the real cell frame.
                    # The caption row is excluded from the available height.
                    frame_width = 0
                    for col in range(caption_col, caption_end_col + 1):
                        width = ws.column_dimensions[
                            get_column_letter(col)
                        ].width
                        if width is None:
                            width = ws.sheet_format.defaultColWidth or 8.43
                        frame_width += int(float(width) * 7 + 5)

                    default_row_height = ws.sheet_format.defaultRowHeight or 15
                    frame_height = sum(
                        points_to_pixels(
                            ws.row_dimensions[row].height or default_row_height
                        )
                        for row in range(anchor_row, caption_row)
                    )
                    # Keep a clear gap above the caption separator; drawings
                    # render over cell borders when they touch the boundary.
                    # Use the same six-pixel inset on every side.  Each source
                    # photo is center-cropped to this inner rectangle, so the
                    # frame has equal top, bottom, left, and right whitespace.
                    frame_inset = 6
                    max_width = max(1, frame_width - frame_inset * 2)
                    max_height = max(1, frame_height - frame_inset * 2)

                    # Respect phone-camera EXIF orientation, then fill the
                    # landscape photo slot with a centered crop.  Portrait
                    # photos no longer appear as a narrow vertical strip.
                    target_size = (int(max_width), int(max_height))
                    with PILImage.open(photo['_resolved_path']) as source_image:
                        source_image = ImageOps.exif_transpose(source_image)
                        if source_image.mode not in ('RGB', 'RGBA'):
                            source_image = source_image.convert('RGB')
                        fitted_image = ImageOps.fit(
                            source_image,
                            target_size,
                            method=PILImage.Resampling.LANCZOS,
                            centering=(0.5, 0.5),
                        )
                        image_buffer = BytesIO()
                        fitted_image.save(image_buffer, format='PNG')
                    image_buffer.seek(0)
                    image = XLImage(image_buffer)
                    image.width, image.height = target_size

                    # Bind both corners to the actual cell frame. A fixed
                    # pixel extent can cross the caption border after Excel
                    # resolves row heights. Both columns use identical insets.
                    marker = AnchorMarker(
                        col=caption_col - 1,
                        colOff=pixels_to_EMU(frame_inset),
                        row=anchor_row - 1,
                        rowOff=pixels_to_EMU(frame_inset),
                    )
                    image.anchor = TwoCellAnchor(
                        editAs='twoCell',
                        _from=marker,
                        to=AnchorMarker(
                            col=caption_end_col,
                            colOff=-pixels_to_EMU(frame_inset),
                            row=caption_row - 1,
                            rowOff=-pixels_to_EMU(frame_inset),
                        ),
                    )
                    ws.add_image(image)

                    joint_no = str(photo.get('joint_no', '') or '').strip()
                    location = str(photo.get('location', '') or '').strip()
                    section = str(photo.get('section', '') or '').strip()
                    line_no = str(photo.get('line_no', '') or '').strip()
                    # Always prefer the 2.x detail table's Section (구간) for
                    # the caption location slot; fall back to location or line
                    # number only when section is absent.
                    if section and section.casefold() != joint_no.casefold():
                        location = section
                    elif not location or location.casefold() == joint_no.casefold():
                        location = next(
                            (
                                value for value in (section, line_no)
                                if value and value.casefold() != joint_no.casefold()
                            ),
                            location,
                        )
                    caption = ' / '.join(filter(None, [
                        str(photo.get('date', '')),
                        location,
                        joint_no,
                        str(photo.get('description', '')),
                    ]))
                    caption = ' '.join(caption.split())
                    for merged in list(ws.merged_cells.ranges):
                        if (
                            merged.min_row <= caption_row <= merged.max_row
                            and merged.min_col <= caption_end_col
                            and merged.max_col >= caption_col
                        ):
                            ws.unmerge_cells(str(merged))
                    ws.merge_cells(
                        start_row=caption_row, start_column=caption_col,
                        end_row=caption_row, end_column=caption_end_col,
                    )
                    caption_cell = ws.cell(row=caption_row, column=caption_col)
                    caption_cell.value = caption
                    caption_cell.alignment = Alignment(
                        horizontal='center', vertical='center',
                        wrap_text=False, shrink_to_fit=True,
                    )
                    caption_cell.font = Font(name='맑은 고딕', size=8)

                    caption_cell.font = Font(
                        name='\ub9d1\uc740 \uace0\ub515', size=9,
                    )
                    ws.row_dimensions[caption_row].height = 30

                    # Make the image and its description read as one framed
                    # item.  Keep the caption inside the bottom of that frame.
                    frame_top = anchor_row
                    frame_range = ws.iter_rows(
                        min_row=frame_top, max_row=caption_row,
                        min_col=caption_col, max_col=caption_end_col,
                    )
                    frame_side = Side(style='thin', color='808080')
                    for frame_row in frame_range:
                        for frame_cell in frame_row:
                            border = copy.copy(frame_cell.border)
                            if frame_cell.row == frame_top:
                                border.top = copy.copy(frame_side)
                            if frame_cell.row == caption_row:
                                border.bottom = copy.copy(frame_side)
                                border.top = copy.copy(frame_side)
                            if frame_cell.column == caption_col:
                                border.left = copy.copy(frame_side)
                            if frame_cell.column == caption_end_col:
                                border.right = copy.copy(frame_side)
                            frame_cell.border = border

        # The template contains additional legacy photo-log pages after the
        # three SEC.16 pages.  Print only the generated process-photo pages;
        # when no valid photo exists, stop at the preceding report page.
        surviving_titles = [
            layout['title_row']
            for layouts in layouts_by_process.values()
            for layout in layouts
        ]
        if photo_section_start is not None:
            print_end_row = (
                max(surviving_titles) + 65
                if surviving_titles else photo_section_start - 1
            )
            self._set_print_area_end(ws, print_end_row)

        # Rebuild the photo-page header deterministically.  Searching another
        # page for a header can select the radiation-safety page because it has
        # the same outer geometry but different text and internal merges.
        thin_black = Side(style='thin', color='000000')
        no_side = Side(style=None)
        header_fill = PatternFill('solid', fgColor='FFFFFF')
        ordered_photo_titles = sorted(surviving_titles)
        total_photo_pages = len(ordered_photo_titles)
        for photo_page_number, title_row in enumerate(
            ordered_photo_titles, start=1
        ):
            header_start = title_row - 9
            header_end = header_start + 6

            if reference_header_heights:
                for offset, height in enumerate(reference_header_heights):
                    ws.row_dimensions[header_start + offset].height = height
            else:
                for row in range(header_start, header_start + 4):
                    ws.row_dimensions[row].height = 10.5
                for row in range(header_start + 4, header_end + 1):
                    ws.row_dimensions[row].height = 6

            # The merged photo-title band spans three worksheet rows, while
            # the revision-number band spans two.  Normalize the three lower
            # rows so both visible merged cells have exactly the same height.
            default_height = ws.sheet_format.defaultRowHeight or 15
            revision_height = sum(
                float(ws.row_dimensions[row].height or default_height)
                for row in range(header_start + 2, header_start + 4)
            )
            photo_row_height = revision_height / 3.0
            for row in range(header_start + 4, header_end + 1):
                ws.row_dimensions[row].height = photo_row_height

            branch_name = ''
            for row in range(header_start, header_end + 1):
                for col in range(1, min(ws.max_column, 23) + 1):
                    value = str(ws.cell(row=row, column=col).value or '').strip()
                    if col <= 5 and '\uc9c0\uc0ac' in value and not branch_name:
                        branch_name = value

            for merged in list(ws.merged_cells.ranges):
                if not (
                    merged.max_row < header_start
                    or merged.min_row > header_end
                ):
                    ws.unmerge_cells(str(merged))

            # Apply the complete grid before merging.  Styling only the anchor
            # cell of a merged range leaves gaps at its far and outer edges.
            for row in range(header_start, header_end + 1):
                for col in range(2, 24):
                    cell = ws.cell(row=row, column=col)
                    cell.fill = header_fill
                    cell.border = Border(
                        left=thin_black,
                        right=thin_black,
                        top=thin_black,
                        bottom=thin_black,
                    )

            header_blocks = (
                (header_start, 2, header_start + 3, 5),
                (header_start + 4, 2, header_end, 5),
                (header_start, 6, header_start + 3, 15),
                (header_start + 4, 6, header_end, 15),
                (header_start, 16, header_start + 1, 23),
                (header_start + 2, 16, header_start + 3, 23),
                (header_start + 4, 16, header_end, 23),
            )
            for min_row, min_col, max_row, max_col in header_blocks:
                ws.merge_cells(
                    start_row=min_row, start_column=min_col,
                    end_row=max_row, end_column=max_col,
                )
                block = ws.cell(row=min_row, column=min_col)
                block.fill = header_fill
                block.alignment = Alignment(
                    horizontal='center', vertical='center', wrap_text=True,
                )

            # The logo and branch name form one visual left panel, so there is
            # no horizontal divider above "\uc911\uc559\uc9c0\uc0ac".  Reassert the
            # black top/left perimeter after merging because openpyxl can drop
            # the far edge of merged ranges.
            for col in range(2, 6):
                logo_edge = ws.cell(header_start + 3, col)
                logo_border = copy.copy(logo_edge.border)
                logo_border.bottom = no_side
                logo_edge.border = logo_border

                branch_edge = ws.cell(header_start + 4, col)
                branch_border = copy.copy(branch_edge.border)
                branch_border.top = no_side
                branch_edge.border = branch_border

            # For merged ranges, Excel/openpyxl rebuilds the far-edge border
            # from the top-left cell when the workbook is reopened.  Clear
            # those source borders too, otherwise the divider reappears even
            # though every cell along the shared edge was cleared above.
            logo_block = ws.cell(header_start, 2)
            logo_block_border = copy.copy(logo_block.border)
            logo_block_border.bottom = no_side
            logo_block.border = logo_block_border

            branch_block = ws.cell(header_start + 4, 2)
            branch_block_border = copy.copy(branch_block.border)
            branch_block_border.top = no_side
            branch_block.border = branch_block_border

            for col in range(2, 24):
                top_cell = ws.cell(header_start, col)
                top_border = copy.copy(top_cell.border)
                top_border.top = thin_black
                top_cell.border = top_border
            for row in range(header_start, header_end + 1):
                left_cell = ws.cell(row, 2)
                left_border = copy.copy(left_cell.border)
                left_border.left = thin_black
                left_cell.border = left_border

            # Match the page-16 logo position as well as the cell heights.
            if reference_logo_anchor is not None and reference_header_start is not None:
                for image in list(ws._images):
                    anchor = getattr(image, 'anchor', None)
                    if not hasattr(anchor, '_from'):
                        continue
                    image_row = anchor._from.row + 1
                    image_col = anchor._from.col + 1
                    if header_start <= image_row <= header_end and image_col <= 5:
                        new_anchor = copy.deepcopy(reference_logo_anchor)
                        row_shift = header_start - reference_header_start
                        new_anchor._from.row += row_shift
                        if hasattr(new_anchor, 'to'):
                            new_anchor.to.row += row_shift
                        image.anchor = new_anchor
                        if reference_logo_size:
                            image.width, image.height = reference_logo_size
                        break

            ws.cell(header_start, 2).value = None
            branch_cell = ws.cell(header_start + 4, 2)
            branch_cell.value = branch_name or '\uc911 \uc559 \uc9c0 \uc0ac'
            branch_cell.fill = PatternFill('solid', fgColor='FFFFFF')
            branch_cell.font = Font(
                name='\ub9d1\uc740 \uace0\ub515', size=8, bold=True,
            )
            branch_cell.alignment = Alignment(
                horizontal='center', vertical='top',
                wrap_text=False, shrink_to_fit=True,
            )
            ws.cell(header_start, 6).value = (
                '[2026\ub144 \uc911\uc559\uc9c0\uc0ac \uc5f4\uc218\uc1a1\uad00 '
                '\ube44\ud30c\uad34\uac80\uc0ac\uc6a9\uc5ed \ub2e8\uac00\uacc4\uc57d]\n'
                '\uc6d4 \uac04 \uc6a9 \uc5ed \uc9c4 \ub3c4 \ubcf4 \uace0 \uc11c'
            )
            ws.cell(header_start, 6).font = Font(
                name='\ub9d1\uc740 \uace0\ub515', size=9, bold=True,
            )
            ws.cell(header_start + 4, 6).value = '6.0 \uc0ac\uc9c4\ub300\uc9c0'
            ws.cell(header_start + 4, 6).font = (
                copy.copy(reference_section_font)
                if reference_section_font
                else Font(name='\ub9d1\uc740 \uace0\ub515', size=10, bold=True)
            )
            ws.cell(header_start, 16).value = (
                f'\ubb38\uc11c\ubc88\ud638 : \uc6d4\uac04\uc6a9\uc5ed\uc9c4\ub3c4\ubcf4\uace0\uc11c {doc_num}\ud638'
            )
            ws.cell(header_start + 2, 16).value = '\uac1c\uc815\ubc88\ud638 :        0'
            ws.cell(header_start + 4, 16).value = (
                f'\ucabd \ubc88 \ud638 :    {photo_page_number}    of    '
                f'{total_photo_pages}'
            )
            for row in (header_start, header_start + 2, header_start + 4):
                ws.cell(row, 16).alignment = Alignment(
                    horizontal='left', vertical='center',
                    wrap_text=False, indent=1,
                )
                ws.cell(row, 16).font = Font(
                    name='\ub9d1\uc740 \uace0\ub515', size=8,
                )

    def _set_print_area_end(self, ws, end_row):
        """Set the vertical print limit and discard later manual breaks."""
        from openpyxl.worksheet.cell_range import CellRange

        target_end = max(1, int(end_row))
        try:
            if ws.print_area:
                raw_range = str(ws.print_area).split('!')[-1].replace('$', '')
                print_range = CellRange(raw_range)
                if (
                    print_range.min_row != 1
                    or print_range.min_col != 2
                    or print_range.max_col != 23
                ):
                    print_range = CellRange(f"B1:W{target_end}")
                else:
                    print_range.max_row = max(print_range.min_row, target_end)
            else:
                print_range = CellRange(f"B1:W{target_end}")
            ws.print_area = f"$B$1:$W${print_range.max_row}"
            ws.row_breaks.brk = [
                brk for brk in ws.row_breaks.brk
                if int(brk.id) < print_range.max_row
            ]
        except Exception:
            ws.print_area = f"$B$1:$W${target_end}"

    def _ensure_valid_print_area(self, ws):
        """Restore a valid $B$1:$W$<end_row> print area if missing or corrupted."""
        try:
            for view in ws.views.sheetView:
                for sel in getattr(view, 'selection', ()) or ():
                    if getattr(sel, 'activeCell', None) and str(sel.activeCell).startswith(('Y', 'Z')):
                        sel.activeCell = 'B1'
                        sel.sqref = 'B1'
        except Exception:
            pass
        pa_text = str(ws.print_area or '').upper()
        pa_match = re.search(
            r"\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)",
            pa_text,
        )
        if (
            pa_match
            and pa_match.group(1) == 'B'
            and int(pa_match.group(2)) == 1
            and pa_match.group(3) == 'W'
            and int(pa_match.group(4)) >= 100
        ):
            ws.print_area = f"$B$1:$W${int(pa_match.group(4))}"
            return
        last_row = max(ws.max_row, 838)
        ws.print_area = f"$B$1:$W${last_row}"

    def _remove_empty_process_photo_pages(self, ws, processes_with_photos):
        """Delete complete template pages for processes without photos."""
        from openpyxl.worksheet.cell_range import CellRange

        keywords = {
            'PAUT': ('위상배열초음파탐상검사', 'PAUT'),
            'MT': ('자분탐상검사',),
            'RT': ('방사선투과검사',),
            'PT': ('침투탐상검사',),
        }
        pages_to_delete = []
        for row in range(600, ws.max_row + 1):
            row_text = ''.join(
                str(ws.cell(row=row, column=col).value or '')
                for col in range(1, ws.max_column + 1)
            ).replace(' ', '').upper()
            if 'SEC.16' not in row_text:
                continue
            for process, process_keywords in keywords.items():
                if process in processes_with_photos:
                    continue
                if any(keyword.upper() in row_text for keyword in process_keywords):
                    # Photo-page titles sit nine rows below the manual page
                    # boundary (PAUT 676:729, MT 730:783, RT 784:837).
                    pages_to_delete.append((row - 9, process))
                    break

        page_rows = 75
        for page_start, _process in sorted(pages_to_delete, reverse=True):
            page_end = page_start + page_rows - 1

            # Remove drawings owned by the deleted page and shift every later
            # logo/photo upward with the worksheet content.
            kept_images = []
            for image in list(ws._images):
                anchor = getattr(image, 'anchor', None)
                if not hasattr(anchor, '_from'):
                    kept_images.append(image)
                    continue
                image_row = anchor._from.row + 1
                if page_start <= image_row <= page_end:
                    continue
                if image_row > page_end:
                    anchor._from.row -= page_rows
                    if hasattr(anchor, 'to'):
                        anchor.to.row -= page_rows
                kept_images.append(image)
            ws._images = kept_images

            self._delete_rows_safely(ws, page_start, page_rows)

            if ws.print_area:
                try:
                    print_range = CellRange(str(ws.print_area).replace('$', ''))
                    if print_range.max_row > page_end:
                        print_range.max_row -= page_rows
                    elif print_range.max_row >= page_start:
                        print_range.max_row = page_start - 1
                    ws.print_area = str(print_range)
                except Exception:
                    pass

    def _clone_process_photo_page(self, ws, source_title_row, continuation_no):
        """Clone one 54-row photo page directly after its source page."""
        import copy
        from io import BytesIO
        from openpyxl.drawing.image import Image as XLImage
        from openpyxl.worksheet.pagebreak import Break

        page_rows = 75
        source_start = source_title_row - 9
        source_end = source_start + page_rows - 1
        insert_row = source_end + 1
        max_col = ws.max_column
        cell_snapshot = [
            [
                (
                    ws.cell(row=row, column=col).value,
                    copy.copy(ws.cell(row=row, column=col)._style),
                    copy.copy(ws.cell(row=row, column=col).alignment),
                    ws.cell(row=row, column=col).number_format,
                )
                for col in range(1, max_col + 1)
            ]
            for row in range(source_start, source_end + 1)
        ]
        row_snapshot = [
            (ws.row_dimensions[row].height, ws.row_dimensions[row].hidden)
            for row in range(source_start, source_end + 1)
        ]
        merge_snapshot = [
            (
                merged.min_row - source_start, merged.min_col,
                merged.max_row - source_start, merged.max_col,
            )
            for merged in list(ws.merged_cells.ranges)
            if merged.min_row >= source_start and merged.max_row <= source_end
        ]
        image_snapshot = []
        for image in list(ws._images):
            anchor = getattr(image, 'anchor', None)
            if hasattr(anchor, '_from') and source_start <= anchor._from.row + 1 <= source_end:
                image_snapshot.append((image, copy.deepcopy(anchor)))

        self._insert_rows_safely(ws, insert_row, page_rows)
        for existing_image in list(ws._images):
            anchor = getattr(existing_image, 'anchor', None)
            if hasattr(anchor, '_from') and anchor._from.row + 1 >= insert_row:
                anchor._from.row += page_rows
                if hasattr(anchor, 'to'):
                    anchor.to.row += page_rows

        for offset, saved_row in enumerate(cell_snapshot):
            target_row = insert_row + offset
            height, hidden = row_snapshot[offset]
            ws.row_dimensions[target_row].height = height
            ws.row_dimensions[target_row].hidden = hidden
            for col, (value, style, alignment, number_format) in enumerate(saved_row, 1):
                cell = ws.cell(row=target_row, column=col)
                cell.value = value
                cell._style = copy.copy(style)
                cell.alignment = copy.copy(alignment)
                cell.number_format = number_format
        for min_row, min_col, max_row, max_col in merge_snapshot:
            ws.merge_cells(
                start_row=insert_row + min_row, start_column=min_col,
                end_row=insert_row + max_row, end_column=max_col,
            )

        row_shift = insert_row - source_start
        for source_image, saved_anchor in image_snapshot:
            buffer = BytesIO()
            image_ref = source_image.ref
            if hasattr(image_ref, 'copy') and hasattr(image_ref, 'save'):
                image_ref.copy().save(
                    buffer, format=getattr(image_ref, 'format', None) or 'PNG'
                )
            elif isinstance(image_ref, (str, os.PathLike)):
                with open(image_ref, 'rb') as stream:
                    buffer.write(stream.read())
            else:
                old_position = image_ref.tell() if hasattr(image_ref, 'tell') else None
                if hasattr(image_ref, 'seek'):
                    image_ref.seek(0)
                buffer.write(image_ref.read())
                if old_position is not None and hasattr(image_ref, 'seek'):
                    image_ref.seek(old_position)
            buffer.seek(0)
            new_image = XLImage(buffer)
            new_image.width = source_image.width
            new_image.height = source_image.height
            new_image._photo_continuation_stream = buffer
            new_image.anchor = copy.deepcopy(saved_anchor)
            new_image.anchor._from.row += row_shift
            if hasattr(new_image.anchor, 'to'):
                new_image.anchor.to.row += row_shift
            ws.add_image(new_image)

        new_title_row = insert_row + 9
        title_cell = ws.cell(row=new_title_row, column=3)
        base_title = re.sub(
            r'\s*\(계속\s*\d+\)\s*$', '', str(title_cell.value or '').strip()
        )
        title_cell.value = f"{base_title} (계속 {continuation_no})"
        page_end = insert_row + page_rows - 1
        if not any(int(brk.id) == page_end for brk in ws.row_breaks.brk):
            ws.row_breaks.append(Break(id=page_end))
        return new_title_row

    def _ensure_pt_photo_page(self, ws):
        """Clone a template photo page after RT when PT photos exist."""
        import copy
        from io import BytesIO
        from openpyxl.drawing.image import Image as XLImage
        from openpyxl.worksheet.pagebreak import Break

        # Reuse an existing generated/template PT page when present.
        for row in range(600, ws.max_row + 1):
            text = ''.join(
                str(ws.cell(row=row, column=col).value or '')
                for col in range(1, ws.max_column + 1)
            ).replace(' ', '').upper()
            if '침투탐상검사' in text and ('SEC.16' in text or 'PT' in text):
                return row

        source_title_row = None
        rt_title_row = None
        for row in range(600, ws.max_row + 1):
            text = ''.join(
                str(ws.cell(row=row, column=col).value or '')
                for col in range(1, ws.max_column + 1)
            ).replace(' ', '').upper()
            if source_title_row is None and '자분탐상검사' in text and 'SEC.16' in text:
                source_title_row = row
            if rt_title_row is None and '방사선투과검사' in text and 'SEC.16' in text:
                rt_title_row = row
        if source_title_row is None or rt_title_row is None:
            return None

        page_rows = 75
        source_start = source_title_row - 9
        source_end = source_start + page_rows - 1
        insert_row = rt_title_row - 9 + page_rows
        max_col = ws.max_column

        cells = []
        for row in range(source_start, source_end + 1):
            cells.append([
                (
                    ws.cell(row=row, column=col).value,
                    copy.copy(ws.cell(row=row, column=col)._style),
                    copy.copy(ws.cell(row=row, column=col).alignment),
                    ws.cell(row=row, column=col).number_format,
                )
                for col in range(1, max_col + 1)
            ])
        row_dimensions = [
            (
                ws.row_dimensions[row].height,
                ws.row_dimensions[row].hidden,
            )
            for row in range(source_start, source_end + 1)
        ]
        merges = [
            (
                merged.min_row - source_start, merged.min_col,
                merged.max_row - source_start, merged.max_col,
            )
            for merged in list(ws.merged_cells.ranges)
            if merged.min_row >= source_start and merged.max_row <= source_end
        ]
        source_images = []
        for image in list(ws._images):
            anchor = getattr(image, 'anchor', None)
            if not hasattr(anchor, '_from'):
                continue
            image_row = anchor._from.row + 1
            if source_start <= image_row <= source_end:
                source_images.append((image, copy.deepcopy(anchor)))

        self._insert_rows_safely(ws, insert_row, page_rows)
        # openpyxl does not move drawings when rows are inserted.  Keep every
        # existing logo/photo below the new PT page aligned with its content.
        for existing_image in list(ws._images):
            anchor = getattr(existing_image, 'anchor', None)
            if not hasattr(anchor, '_from') or anchor._from.row + 1 < insert_row:
                continue
            anchor._from.row += page_rows
            if hasattr(anchor, 'to'):
                anchor.to.row += page_rows
        for offset, row_cells in enumerate(cells):
            target_row = insert_row + offset
            height, hidden = row_dimensions[offset]
            ws.row_dimensions[target_row].height = height
            ws.row_dimensions[target_row].hidden = hidden
            for col, (value, style, alignment, number_format) in enumerate(
                row_cells, start=1
            ):
                cell = ws.cell(row=target_row, column=col)
                cell.value = value
                cell._style = copy.copy(style)
                cell.alignment = copy.copy(alignment)
                cell.number_format = number_format
        for min_row, min_col, max_row, max_col in merges:
            ws.merge_cells(
                start_row=insert_row + min_row, start_column=min_col,
                end_row=insert_row + max_row, end_column=max_col,
            )

        row_shift = insert_row - source_start
        for source_image, saved_anchor in source_images:
            image_buffer = BytesIO()
            image_ref = source_image.ref
            if hasattr(image_ref, 'copy') and hasattr(image_ref, 'save'):
                image_ref.copy().save(
                    image_buffer,
                    format=getattr(image_ref, 'format', None) or 'PNG',
                )
            elif isinstance(image_ref, (str, os.PathLike)):
                with open(image_ref, 'rb') as stream:
                    image_buffer.write(stream.read())
            else:
                old_position = image_ref.tell() if hasattr(image_ref, 'tell') else None
                if hasattr(image_ref, 'seek'):
                    image_ref.seek(0)
                image_buffer.write(image_ref.read())
                if old_position is not None and hasattr(image_ref, 'seek'):
                    image_ref.seek(old_position)
            image_buffer.seek(0)
            new_image = XLImage(image_buffer)
            new_image.width = source_image.width
            new_image.height = source_image.height
            new_image._process_photo_page_stream = image_buffer
            new_image.anchor = copy.deepcopy(saved_anchor)
            new_image.anchor._from.row += row_shift
            if hasattr(new_image.anchor, 'to'):
                new_image.anchor.to.row += row_shift
            ws.add_image(new_image)

        # Keep RT and the generated PT page as separate printed pages.
        rt_page_end = insert_row - 1
        pt_page_end = insert_row + page_rows - 1
        existing_breaks = {int(brk.id) for brk in ws.row_breaks.brk}
        if rt_page_end not in existing_breaks:
            ws.row_breaks.append(Break(id=rt_page_end))
        if pt_page_end not in existing_breaks:
            ws.row_breaks.append(Break(id=pt_page_end))

        pt_title_row = insert_row + (source_title_row - source_start)
        ws.cell(row=pt_title_row, column=3).value = '1.4 침투탐상검사(PT)(SEC.16)'
        return pt_title_row


    def _complete_ndt_quantity_totals(self, ws):
        """Fill missing TOTAL cells in summary/detail and subcontract subtotals."""
        active, method = False, None
        max_col = min(ws.max_column, 23)
        for row in range(1, ws.max_row + 1):
            section = ''.join(str(ws.cell(row, 6).value or '').split())
            if section.startswith('3.0'):
                active = True
            elif active and section.startswith('4.0'):
                break
            if not active:
                continue
            label = ''.join(str(ws.cell(row, col).value or '')
                            for col in range(1, max_col + 1))
            if 'LINENO' in ''.join(label.split()).replace('.', '').upper():
                qty_label = str(ws.cell(row, 17).value or '').upper()
                method = next((m for m in ('PAUT', 'MT', 'PT') if m in qty_label),
                              'RT' if '촬영' in qty_label else None)
                continue
            target = ws.cell(row, 20)
            if method is None or not isinstance(target, Cell) or target.value not in (None, ''):
                continue
            try:
                original = self._safe_float(ws.cell(row, 17).value)
                retest = self._safe_float(ws.cell(row, 18).value)
            except (TypeError, ValueError):
                continue
            if original == 0 and retest == 0:
                continue
            target.value = int(original + retest) if method == 'RT' else original + retest
            target.number_format = '0' if method == 'RT' else '0.0000'
            target.alignment = self.align_center

    def _annotate_ndt_overruns(self, ws):
        """Expose exceeded estimates without inventing change-order reasons."""
        active = False
        for row in range(1, ws.max_row + 1):
            section = ''.join(str(ws.cell(row, 6).value or '').split())
            if section.startswith('3.0'):
                active = True
            if not active:
                continue
            if str(ws.cell(row, 2).value or '').replace(' ', '').startswith('1.2SECTION'):
                break
            try:
                expected = float(str(ws.cell(row, 7).value).replace(',', ''))
                cumulative = float(str(ws.cell(row, 13).value).replace(',', ''))
            except (TypeError, ValueError):
                continue
            if expected > 0 and cumulative > expected + 0.0000001:
                note = ws.cell(row, 21)
                existing = str(note.value or '').strip()
                if isinstance(note, Cell) and '예상량 초과' not in existing:
                    note.value = f'{existing} / 예상량 초과' if existing else '예상량 초과'
                    note.alignment = Alignment(horizontal='center', vertical='center',
                                               wrap_text=False, shrink_to_fit=True)

    def _layout_ndt_pages(self, ws, capacity_points=None):
        """Rebuild the complete 3.0 body to one shared height for both scopes.

        Remove template page boundaries before filling pages in source order.
        A subtotal, TOTAL, empty-result message or change of inspection method
        must not prevent the following content from using available space.
        Only the document/table headings needed at an actual split repeat.
        """
        self._reflow_ndt_section(ws)
        # Reserve 15pt below the header and 20pt above the logo before
        # pagination, then transfer the latter into an actual blank row.
        self._lower_ndt_page_titles(ws, extra_height=35.0)
        self._normalize_ndt_parent_caption_style(ws)
        self._space_ndt_section_caption(ws)
        self._paginate_reflowed_ndt(ws, capacity_points=capacity_points)
        self._insert_fixed_ndt_top_spacers(ws, reserved_height=20.0)
        self._space_ndt_summary_caption(ws)

    def _center_enlarge_report_logos(self, ws, factor=1.5):
        """Center visible logo artwork above the branch label, enlarged 1.5x."""
        from io import BytesIO
        from PIL import Image as PILImage, ImageChops
        from openpyxl.drawing.image import Image as XLImage
        from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
        from openpyxl.drawing.xdr import XDRPositiveSize2D
        from openpyxl.utils import get_column_letter
        from openpyxl.utils.units import pixels_to_EMU

        starts = sorted({m.min_row for m in ws.merged_cells.ranges
                         if m.min_col == 6 and m.max_col == 15
                         and m.max_row >= m.min_row + 2
                         and '월간용역진도보고서' in ''.join(
                             str(ws.cell(m.min_row, 6).value or '').split())})
        if not starts:
            return
        # The Korean template's Dotum/Gulim 11 digit-width metric is 8px.
        digit_width = 8 if ws.parent._fonts[0].name in (
            '굴림', '굴림체', '돋움', '돋움체') else 7
        widths = {}
        for col in range(2, 6):
            dimension = ws.column_dimensions[get_column_letter(col)]
            widths[col] = pixels_to_EMU(int(float(dimension.width or 8.43)
                                           * digit_width + 5))
        frame_width = sum(widths.values())
        default = float(ws.sheet_format.defaultRowHeight or 15)
        for image in list(ws._images):
            anchor = image.anchor
            if not hasattr(anchor, 'ext') or not hasattr(anchor, '_from'):
                continue
            start = next((r for r in starts
                          if r <= anchor._from.row + 1 <= r + 6), None)
            if start is None or anchor._from.col > 4 or image.height >= 100:
                continue
            ref = image.ref
            position = ref.tell() if hasattr(ref, 'tell') else None
            if hasattr(ref, 'seek'):
                ref.seek(0)
            artwork = PILImage.open(ref).convert('RGB')
            bbox = ImageChops.difference(
                artwork, PILImage.new('RGB', artwork.size, 'white')).getbbox()
            if hasattr(ref, 'seek') and position is not None:
                ref.seek(position)
            if not bbox:
                continue
            width = int(anchor.ext.cx * factor * (bbox[2] - bbox[0]) / artwork.width)
            height = int(anchor.ext.cy * factor * (bbox[3] - bbox[1]) / artwork.height)
            buffer = BytesIO()
            artwork.crop(bbox).save(buffer, format='PNG')
            buffer.seek(0)
            replacement = XLImage(buffer)
            replacement._report_header_stream = buffer
            # Center on the actual artwork, excluding the asset's white tail.
            horizontal = max(0, (frame_width - width) // 2)
            col = 2
            while col < 5 and horizontal >= widths[col]:
                horizontal -= widths[col]
                col += 1
            top_height = sum(float(ws.row_dimensions[r].height or default)
                             for r in range(start, start + 4)) * 12700
            vertical = max(0, int((top_height - height) / 2))
            row = start
            while row < start + 6:
                row_height = int(float(ws.row_dimensions[row].height or default) * 12700)
                if vertical < row_height:
                    break
                vertical -= row_height
                row += 1
            replacement.anchor = OneCellAnchor(
                _from=AnchorMarker(col=col - 1, colOff=horizontal,
                                   row=row - 1, rowOff=vertical),
                ext=XDRPositiveSize2D(width, height))
            ws._images[ws._images.index(image)] = replacement

    def _fix_report_section_header_labels(self, ws):
        """Keep TOC, general-work and photo headers correctly numbered."""
        starts = sorted({m.min_row for m in ws.merged_cells.ranges
                         if m.min_col == 6 and m.max_col == 15
                         and m.max_row >= m.min_row + 2})
        for index, start in enumerate(starts):
            end = starts[index + 1] if index + 1 < len(starts) else ws.max_row + 1
            section = ws.cell(start + 4, 6)
            compact = ''.join(str(section.value or '').split())
            if compact == '목차':
                body = ''.join(str(ws.cell(row, col).value or '')
                               for row in range(start + 7, min(end, start + 40))
                               for col in range(1, min(ws.max_column, 23) + 1))
                if any(token in ''.join(body.split())
                       for token in ('용역명', '현장조직도', '검사용역의범위')):
                    section.value = '1.0 용역업무일반'
            elif '사진대지' in compact:
                section.value = '7.0 사진대지'
        # Section-local page counts must follow the actual final sections,
        # not the template's old 1.0/2.0 boundary or unused photo pages.
        print_end = ws.max_row
        if ws.print_area:
            ends = re.findall(r'\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?(\d+)',
                              str(ws.print_area))
            if ends:
                print_end = max(map(int, ends))
        groups = {}
        for start in starts:
            if start > print_end:
                continue
            text = ''.join(str(ws.cell(start + 4, 6).value or '').split())
            match = re.match(r'^(\d+\.0)', text)
            key = match.group(1) if match else 'toc' if text == '목차' else None
            if key is not None:
                groups.setdefault(key, []).append(start)
        for group in groups.values():
            for number, start in enumerate(group, 1):
                ws.cell(start + 4, 16).value = (
                    f'  쪽  번 호 :      {number}     of     {len(group)}'
                )

    def _space_photo_page_caption(self, ws, gap=5.0):
        """Separate the photo caption from the report frame using its spacer."""
        from copy import copy

        for row in range(1, ws.max_row):
            cells = [cell for cell in ws[row] if cell.value is not None]
            if not any(''.join(str(cell.value).split()) == '1.사진대지'
                       for cell in cells):
                continue
            # Borrow space from the blank row below, preserving photo anchors
            # and total page height rather than moving the whole photo grid.
            if any(cell.value is not None for cell in ws[row + 1]):
                continue
            default = float(ws.sheet_format.defaultRowHeight or 15)
            below = ws.row_dimensions[row + 1]
            caption = ws.row_dimensions[row]
            if below.hidden:
                continue
            height = float(below.height if below.height is not None else default)
            amount = min(float(gap), max(0.0, height - 1.0))
            below.height = height - amount
            caption.height = float(caption.height if caption.height is not None
                                   else default) + amount
            caption.hidden = False
            for cell in cells:
                alignment = copy(cell.alignment)
                alignment.vertical = 'bottom'
                cell.alignment = alignment

    def _save_workbook_with_fallback(self, wb, output_path):
        """Save the report, avoiding a hard failure when Excel locks the file.

        Windows denies overwriting an xlsx that is already open in Excel.  In
        that case keep the user's selected folder and save a timestamped copy.
        Build the complete workbook beside the destination before replacing it.
        """
        # Continuation photo pages are inserted before the template's existing
        # MT/RT pages, so their new breaks can be appended out of row order.
        # Excel ignores those later out-of-order boundaries and splits the
        # following header between two printed pages. Normalize only after all
        # row insertions and spacing adjustments, including the fallback save.
        for sheet in wb.worksheets:
            breaks_by_row = {}
            for page_break in sheet.row_breaks.brk:
                page_break.man = True
                breaks_by_row[int(page_break.id)] = page_break
            sheet.row_breaks.brk = [
                breaks_by_row[row] for row in sorted(breaks_by_row)
            ]
        output_path = os.path.abspath(output_path)
        output_dir = os.path.dirname(output_path)
        temporary_path = None
        try:
            with tempfile.NamedTemporaryFile(
                dir=output_dir, prefix=".monthly-report-", suffix=".xlsx", delete=False
            ) as temporary:
                temporary_path = temporary.name
            wb.save(temporary_path)
            try:
                os.replace(temporary_path, output_path)
                return output_path
            except PermissionError:
                stem, extension = os.path.splitext(output_path)
                timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S_%f')
                fallback_path = f'{stem}_{timestamp}{extension or ".xlsx"}'
                os.rename(temporary_path, fallback_path)
                return fallback_path
        except PermissionError as error:
            raise PermissionError(
                    '\uc120\ud0dd\ud55c \ud3f4\ub354\uc5d0 \ubcf4\uace0\uc11c\ub97c \uc800\uc7a5\ud560 \uc218 \uc5c6\uc2b5\ub2c8\ub2e4. '
                    '\uae30\uc874 Excel \ud30c\uc77c\uc744 \ub2eb\uac70\ub098 OneDrive \ub3d9\uae30\ud654\ub97c \ud655\uc778\ud55c \ud6c4, '
                    '\ub2e4\ub978 \ud3f4\ub354\ub97c \uc120\ud0dd\ud574 \ub2e4\uc2dc \uc2dc\ub3c4\ud574 \uc8fc\uc138\uc694.\n'
                    f'\uc800\uc7a5 \uacbd\ub85c: {output_path}'
            ) from error
        finally:
            if temporary_path:
                try:
                    os.unlink(temporary_path)
                except OSError:
                    pass

    def _safe_float(self, val):
        try:
            return float(str(val).replace(',', ''))
        except (ValueError, TypeError):
            return 0.0

    def _format_num(self, val):
        try:
            f = float(str(val).replace(',', ''))
            if f == 0: return "-"
            if f % 1 == 0: return int(f)
            return f
        except:
            return val if val else "-"

    def _write_row(self, ws, row_idx, values):
        """배열의 값들을 row_idx 행의 1번 컬럼부터 순서대로 적음. None은 건너뜀"""
        for col_idx, val in enumerate(values):
            if val is not None:
                cell = ws.cell(row=row_idx, column=col_idx+1)
                cell.value = val
                cell.font = self.font_normal
                cell.alignment = self.align_center
                cell.border = self.border_thin

    def _clone_safety_training_header(self, ws, start_row, header_row, insert_row):
        import copy
        from io import BytesIO
        from openpyxl.drawing.image import Image as XLImage
        from openpyxl.styles import Border, Side
        from openpyxl.worksheet.pagebreak import Break, RowBreak

        doc_header_start = 1
        for row in range(start_row - 1, 0, -1):
            val = ''.join(str(ws.cell(row=row, column=c).value or '') for c in range(1, 10)).replace(' ', '')
            if '단가계약' in val or '월간용역진도보고서' in val:
                doc_header_start = max(1, row - 1)
                break

        doc_header_end = doc_header_start + 4
        source_rows = list(range(doc_header_start, doc_header_end + 1)) + [start_row - 1] + list(range(start_row, header_row + 1))
        clone_height = len(source_rows)
        max_col = ws.max_column

        cell_snapshot, row_snapshot = [], []
        for r in source_rows:
            row_data = []
            for c in range(1, max_col + 1):
                cell = ws.cell(row=r, column=c)
                row_data.append((cell.value, copy.copy(cell._style), copy.copy(cell.alignment), cell.number_format))
            cell_snapshot.append(row_data)
            row_snapshot.append((ws.row_dimensions[r].height, ws.row_dimensions[r].hidden))

        merge_snapshot = []
        for merged in list(ws.merged_cells.ranges):
            for i, r in enumerate(source_rows):
                if merged.min_row == r and merged.max_row == r:
                    merge_snapshot.append((i, merged.min_col, merged.max_col))
                    break

        image_snapshot = []
        for image in list(ws._images):
            anchor = getattr(image, 'anchor', None)
            if hasattr(anchor, '_from'):
                img_r = anchor._from.row + 1
                if doc_header_start <= img_r <= doc_header_end:
                    image_snapshot.append((image, copy.deepcopy(anchor), img_r - doc_header_start))

        # Give bottom border to previous page
        thick = Side(style='medium')
        for c in range(1, max_col + 1):
            cell = ws.cell(row=insert_row - 1, column=c)
            if cell.border and (cell.border.left or cell.border.right):
                new_border = copy.copy(cell.border)
                new_border.bottom = thick
                cell.border = new_border

        self._insert_rows_safely(ws, insert_row, clone_height)

        for existing_image in list(ws._images):
            anchor = getattr(existing_image, 'anchor', None)
            if hasattr(anchor, '_from') and anchor._from.row + 1 >= insert_row:
                anchor._from.row += clone_height
                if hasattr(anchor, 'to'):
                    anchor.to.row += clone_height

        for offset, row_data in enumerate(cell_snapshot):
            target_row = insert_row + offset
            height, hidden = row_snapshot[offset]
            ws.row_dimensions[target_row].height = height
            ws.row_dimensions[target_row].hidden = hidden
            for col, (value, style, alignment, number_format) in enumerate(row_data, 1):
                cell = ws.cell(row=target_row, column=col)
                cell.value = value
                cell._style = copy.copy(style)
                cell.alignment = copy.copy(alignment)
                cell.number_format = number_format

        for offset, min_col, max_col in merge_snapshot:
            try:
                ws.merge_cells(start_row=insert_row+offset, start_column=min_col, end_row=insert_row+offset, end_column=max_col)
            except ValueError:
                pass

        for source_image, saved_anchor, offset in image_snapshot:
            buffer = BytesIO()
            image_ref = source_image.ref
            if hasattr(image_ref, 'copy') and hasattr(image_ref, 'save'):
                image_ref.copy().save(buffer, format=getattr(image_ref, 'format', None) or 'PNG')
            elif isinstance(image_ref, (str, os.PathLike)):
                with open(image_ref, 'rb') as stream:
                    buffer.write(stream.read())
            else:
                old_position = image_ref.tell() if hasattr(image_ref, 'tell') else None
                if hasattr(image_ref, 'seek'):
                    image_ref.seek(0)
                buffer.write(image_ref.read())
                if old_position is not None and hasattr(image_ref, 'seek'):
                    image_ref.seek(old_position)
            buffer.seek(0)
            new_image = XLImage(buffer)
            new_image.width = source_image.width
            new_image.height = source_image.height

            new_anchor = copy.deepcopy(saved_anchor)
            new_anchor._from.row = (insert_row + offset) - 1
            if hasattr(new_anchor, 'to'):
                new_anchor.to.row = new_anchor._from.row + (saved_anchor.to.row - saved_anchor._from.row)
            new_image.anchor = new_anchor
            ws.add_image(new_image)

        if not ws.row_breaks:
            ws.row_breaks = RowBreak()
        ws.row_breaks.append(Break(id=insert_row - 1))

        return insert_row + clone_height

    def _resolve_safety_training_dir(self):
        """Find the designated education folder on the active data drive."""
        import ntpath

        relative_parts = (
            '내 드라이브', '01_업무_및_회사', 'Office', '착공',
            '2026년 중앙지사 열수송관 비파괴검사 단가계약', '교육일지',
        )
        history_drive = ntpath.splitdrive(
            str(getattr(self, '_history_path', '') or '')
        )[0]
        drives = []
        for drive in (history_drive, 'F:', 'G:'):
            if drive and drive.upper() not in {item.upper() for item in drives}:
                drives.append(drive)
        candidates = [os.path.join(drive + '\\', *relative_parts) for drive in drives]
        return next(
            (candidate for candidate in candidates if os.path.isdir(candidate)),
            candidates[0] if candidates else os.path.join('F:\\', *relative_parts),
        )

    def _populate_safety_training_log(self, ws, target_ym, ndt_groups, target_dates=None, report_scope="month"):
        """Populate the 2.0 Safety Management Training Status table using values from 지역난방_안전관리교육."""
        try:
            import sys
            import copy
            import glob
            import os
            from openpyxl.styles import Alignment

        except Exception:
            return

        try:
            year, month = map(int, target_ym.split('-'))
        except:
            return

        # Create training contents based on report_scope
        if report_scope == "cumulative":
            start_year, start_month = 2026, 8
            total_months = (year - start_year) * 12 + (month - start_month) + 1
        else:
            start_year, start_month = year, month
            total_months = 1

        training_logs = []
        # Prefer the actual workbook produced by 지역난방_안전관리교육.py.
        # Its 교육현황 sheet contains the authoritative date, category,
        # content, duration, instructor, and location entered by the user.
        # 교육 원본은 사용자가 지정한 교육일지 폴더만 신뢰한다.
        # 프로그램 폴더나 바탕화면의 복사본을 함께 검색하면 더 최근에
        # 수정된 임시 PAUT 파일이 지정 폴더의 RT 원본을 덮어쓸 수 있다.
        safety_training_dir = self._resolve_safety_training_dir()
        target_months = []
        for offset in range(max(0, total_months)):
            source_year = start_year + (start_month - 1 + offset) // 12
            source_month = (start_month - 1 + offset) % 12 + 1
            target_months.append((source_year, source_month))

        for source_year, source_month in target_months:
            safety_filename = (
                f'\uc9c0\uc5ed\ub09c\ubc29_\uc548\uc804\uad00\ub9ac\uad50\uc721_'
                f'{source_year}{source_month:02d}.xlsx'
            )
            safety_files = glob.glob(
                os.path.join(safety_training_dir, safety_filename)
            )
            if not safety_files:
                print(
                    '[WARN] Safety training source not found in designated '
                    f'folder: {os.path.join(safety_training_dir, safety_filename)}'
                )
                continue

            newest_safety_file = max(set(safety_files), key=os.path.getmtime)
            try:
                safety_wb = openpyxl.load_workbook(
                    newest_safety_file, data_only=True
                )
                summary_name = '\uad50\uc721\ud604\ud669'
                summary_ws = safety_wb[summary_name] \
                    if summary_name in safety_wb.sheetnames else None
                if summary_ws is not None:
                    target_key = f'{source_year}{source_month:02d}'
                    for source_row in range(2, summary_ws.max_row + 1):
                        raw_date = summary_ws.cell(source_row, 1).value
                        if isinstance(raw_date, (datetime.datetime, datetime.date)):
                            date_text = raw_date.strftime('%Y-%m-%d')
                        else:
                            date_text = str(raw_date or '').strip()
                        date_key = re.sub(r'\D', '', date_text)[:6]
                        if date_key != target_key:
                            continue
                        full_content = str(
                            summary_ws.cell(source_row, 3).value or ''
                        ).strip()
                        # The education workbook keeps the full training text.
                        # The monthly progress report only needs the two main
                        # subjects: common safety and process-specific safety.
                        content_parts = [
                            part.strip() for part in full_content.split('/')
                            if part.strip()
                        ]
                        representative_topics = []
                        for section_number in ('1', '2'):
                            for part_index, part in enumerate(content_parts):
                                if not re.match(
                                    rf'^{section_number}\.\s+', part
                                ):
                                    continue
                                heading = re.sub(
                                    rf'^{section_number}\.\s*', '', part
                                ).strip()
                                detail = content_parts[part_index + 1] \
                                    if part_index + 1 < len(content_parts) else ''
                                representative_topics.append(
                                    f'{heading}: {detail}' if detail else heading
                                )
                                break
                        report_content = '\n'.join(representative_topics) \
                            if representative_topics else full_content
                        training_logs.append({
                            'date': date_text.split()[0],
                            'category': str(summary_ws.cell(source_row, 2).value or '').strip(),
                            'content': report_content,
                            'time': str(summary_ws.cell(source_row, 4).value or '').strip(),
                            'instructor': str(summary_ws.cell(source_row, 5).value or '').strip(),
                            'location': '서울검사 사무실',
                        })
                safety_wb.close()
            except Exception as exc:
                print(f'[WARN] Failed to load safety training summary: {exc}')

        # A safety workbook can contain the same education record more than
        # once (for example after copying or regenerating its summary sheet).
        # Keep the first occurrence, but preserve genuinely different classes
        # held on the same date.
        if training_logs:
            unique_training_logs = []
            seen_training_logs = set()
            for log in training_logs:
                duplicate_key = tuple(
                    ' '.join(str(log.get(field, '') or '').split()).casefold()
                    for field in (
                        'date', 'category', 'content', 'time',
                        'instructor', 'location',
                    )
                )
                if duplicate_key in seen_training_logs:
                    continue
                seen_training_logs.add(duplicate_key)
                unique_training_logs.append(log)
            training_logs = unique_training_logs

        category_val = "월간안전교육"
        time_val = "1시간"
        instructor_val = "현장소장"
        location_val = "사무실"

        # Find the table in the template
        start_row = None
        for row in range(1, ws.max_row + 1):
            val = str(ws.cell(row=row, column=2).value or '').replace(' ', '')
            if '2.0안전관리교육' in val or '2.0안전관리' in val:
                start_row = row
                break

        if not start_row:
            return

        header_row = None
        for row in range(start_row + 1, start_row + 10):
            val = str(ws.cell(row=row, column=2).value or '').replace(' ', '')
            if '일자' in val:
                header_row = row
                break

        if not header_row:
            return

        data_row = header_row + 1
        for merged in ws.merged_cells.ranges:
            if merged.min_row <= header_row <= merged.max_row:
                data_row = max(data_row, merged.max_row + 1)

        cols = {'date': 2, 'category': 4, 'content': 8, 'time': 18, 'instructor': 20, 'location': 22}
        for col in range(2, 24):
            val = str(ws.cell(row=header_row, column=col).value or '').replace(' ', '')
            if '일자' in val: cols['date'] = col
            elif '구분' in val: cols['category'] = col
            elif '내용' in val: cols['content'] = col
            elif '시간' in val: cols['time'] = col
            elif '강사' in val: cols['instructor'] = col
            elif '장소' in val: cols['location'] = col

        reference_row = data_row
        row_merges = []
        record_row_span = 1
        for merged_range in list(ws.merged_cells.ranges):
            if merged_range.min_row == reference_row:
                record_row_span = max(
                    record_row_span,
                    merged_range.max_row - merged_range.min_row + 1,
                )
                row_merges.append((
                    merged_range.min_col,
                    merged_range.max_col,
                    merged_range.max_row - merged_range.min_row,
                ))

        max_rows_per_page = 7
        current_data_row = data_row
        for i, log in enumerate(training_logs):
            # 교육내용이 여러 페이지로 늘어나는 경우에도 매 페이지마다
            # 공통 문서 헤더와 2.0 교육현황 표 헤더를 반복한다.
            if i > 0 and i % max_rows_per_page == 0:
                current_data_row = self._clone_safety_training_header(ws, start_row, header_row, current_data_row)

            current_row = current_data_row

            # Copy formatting and merges for rows beyond the first one
            if current_row > reference_row:
                for row_offset in range(record_row_span):
                    for col in range(2, 24):
                        ref_cell = ws.cell(
                            row=reference_row + row_offset, column=col
                        )
                        new_cell = ws.cell(
                            row=current_row + row_offset, column=col
                        )
                        if ref_cell.has_style:
                            new_cell.font = copy.copy(ref_cell.font)
                            new_cell.border = copy.copy(ref_cell.border)
                            new_cell.fill = copy.copy(ref_cell.fill)
                            new_cell.number_format = copy.copy(ref_cell.number_format)
                            new_cell.alignment = copy.copy(ref_cell.alignment)

                for min_col, max_col, max_row_offset in row_merges:
                    try:
                        ws.merge_cells(
                            start_row=current_row,
                            start_column=min_col,
                            end_row=current_row + max_row_offset,
                            end_column=max_col,
                        )
                    except ValueError:
                        pass

            ws.cell(row=current_row, column=cols.get('date', 2)).value = log['date']
            ws.cell(row=current_row, column=cols.get('category', 4)).value = log.get('category') or category_val
            ws.cell(row=current_row, column=cols.get('content', 8)).value = log['content']
            ws.cell(row=current_row, column=cols.get('time', 18)).value = log.get('time') or time_val
            ws.cell(row=current_row, column=cols.get('instructor', 20)).value = log.get('instructor') or instructor_val
            ws.cell(row=current_row, column=cols.get('location', 22)).value = '서울검사 사무실'

            c = ws.cell(row=current_row, column=cols.get('content', 8))
            align = copy.copy(c.alignment) if c.alignment else Alignment()
            align.wrap_text = True
            align.vertical = 'center'
            align.horizontal = 'left'
            c.alignment = align

            for c_idx in [cols.get('date',2), cols.get('category',4), cols.get('time',18), cols.get('instructor',20), cols.get('location',22)]:
                cell = ws.cell(row=current_row, column=c_idx)
                align_center = copy.copy(cell.alignment) if cell.alignment else Alignment()
                align_center.wrap_text = True
                align_center.vertical = 'center'
                align_center.horizontal = 'center'
                cell.alignment = align_center

            lines = log['content'].split('\n')
            estimated_height = sum(max(1, len(line)//45 + 1) for line in lines) * 20
            ws.row_dimensions[current_row].height = max(45, estimated_height)

            current_data_row += record_row_span
