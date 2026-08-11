import openpyxl
from openpyxl.styles import Alignment, Border, Side, Font
from openpyxl.cell.cell import Cell
import os
import json
import datetime
from collections import defaultdict
import re

class MonthlyReportManager:
    def __init__(self, template_path):
        self.template_path = template_path
        # 템플릿의 각 표 시작 위치 (헤더 기준)
        # 이 위치는 템플릿_최종완성본_V70 기준이며, 동적 삽입 시 밀릴 수 있으므로 
        # 위에서부터 아래로 순차적으로 삽입(insert_rows)하면서 작업해야 합니다.
        self.table_markers = {
            'qty': 382, # 1.1 비파괴검사 물량표
            'paut_1': 403, # 1.2.1 PAUT
            'mt_1': 412, # 1.2.2 MT
            'rt_1': 429, # 1.2.3 RT
            'pt_1': 436, # 1.2.4 PT
            'paut_2': 448, # 2.1 PAUT 세부
            'mt_2': 481, # 2.2 MT 세부
            'rt_2': 490, # 2.3 RT 세부
            'pt_2': 497 # 2.4 PT 세부
        }
        
        self.font_normal = Font(name='맑은 고딕', size=10)
        self.align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
        thin = Side(style='thin')
        self.border_thin = Border(left=thin, right=thin, top=thin, bottom=thin)

    def _insert_rows_safely(self, ws, insert_idx, amount):
        """
        openpyxl의 병합 셀 깨집 버그를 우회하여 안전하게 행을 삽입합니다.
        삽입 위치 아래의 병합 셀 구조를 보존합니다.
        """
        from openpyxl.worksheet.cell_range import CellRange
        from openpyxl.worksheet.pagebreak import Break
        # 1. 모든 병합 정보 백업
        old_merges = [str(m) for m in ws.merged_cells.ranges]
        # 2. 병합 정보 초기화 (에러 원인 제거)
        ws.merged_cells.ranges = []
        # 3. 안전하게 빈 행 삽입
        ws.insert_rows(insert_idx, amount)
        # 4. 백업된 병합 정보를 위치에 맞게 수정하여 다시 적용
        for m_str in old_merges:
            cr = CellRange(m_str)
            if cr.min_row >= insert_idx:
                cr.shift(row_shift=amount)   # 삽입 위치 아래 병합셀 밀어냄
            elif cr.min_row < insert_idx <= cr.max_row:
                cr.expand(down=amount)       # 삽입 위치에 걸치는 병합셀 늘림
            ws.merge_cells(str(cr))

        # 5. 페이지 브레이크(로고/쪽번호 위치) 보정
        if ws.row_breaks and ws.row_breaks.brk:
            new_breaks = []
            for brk in ws.row_breaks.brk:
                if brk.id >= insert_idx:
                    new_breaks.append(Break(id=brk.id + amount))
                else:
                    new_breaks.append(brk)
            ws.row_breaks.brk = new_breaks

        # 6. 인쇄 영역(print_area) 보정
        if ws.print_area:
            try:
                pa = CellRange(str(ws.print_area).replace('$', ''))
                if pa.max_row >= insert_idx:
                    pa.max_row += amount
                    ws.print_area = str(pa)
            except Exception:
                pass
    def _delete_rows_safely(self, ws, delete_idx, amount):
        """
        openpyxl의 병합 셀 깨집 버그를 우회하여 안전하게 행을 삭제합니다.
        """
        from openpyxl.worksheet.cell_range import CellRange
        from openpyxl.worksheet.pagebreak import Break
        # 1. 삭제 범위에 완전히 포함되지 않는 병합 보존
        old_merges = []
        for m in ws.merged_cells.ranges:
            if m.min_row >= delete_idx and m.max_row < delete_idx + amount:
                continue  # 삭제될 행 안에만 있는 병합은 제거
            old_merges.append(str(m))

        ws.merged_cells.ranges = []
        # 2. 행 삭제
        ws.delete_rows(delete_idx, amount)
        # 3. 병합 정보 복원 (삭제 위치 기준으로 조정)
        for m_str in old_merges:
            cr = CellRange(m_str)
            if cr.min_row >= delete_idx + amount:
                cr.shift(row_shift=-amount)   # 삭제 아래 병합셀 당김
            elif cr.max_row >= delete_idx + amount:
                cr.max_row -= amount          # 범위에 걸치는 병합셀 축소
            try:
                ws.merge_cells(str(cr))
            except Exception:
                pass

        # Adjust page breaks after deleted rows.
        if ws.row_breaks and ws.row_breaks.brk:
            new_breaks = []
            for brk in ws.row_breaks.brk:
                if brk.id >= delete_idx + amount:
                    new_breaks.append(Break(id=brk.id - amount))
                elif brk.id < delete_idx:
                    new_breaks.append(brk)
            ws.row_breaks.brk = new_breaks

    def _repair_merged_right_borders(self, ws):
        """Restore right edges that openpyxl can drop from merged template cells.

        Excel stores a merged range's visible outline across both the anchor cell
        and the cells on the perimeter.  After a load/save cycle, the right edge
        can disappear when the template only carries the outline on the anchor.
        Reapply that edge to the full right perimeter before saving.
        """
        import copy

        for merged in list(ws.merged_cells.ranges):
            anchor = ws.cell(row=merged.min_row, column=merged.min_col)

            # Prefer an existing right edge.  For framed merged cells whose right
            # edge was already lost, the matching left edge is the intended style.
            right_side = anchor.border.right
            if not right_side.style:
                for row in range(merged.min_row, merged.max_row + 1):
                    candidate = ws.cell(row=row, column=merged.max_col).border.right
                    if candidate.style:
                        right_side = candidate
                        break

            is_framed = bool(
                anchor.border.left.style
                and (anchor.border.top.style or anchor.border.bottom.style)
            )
            if not right_side.style and is_framed:
                right_side = anchor.border.left

            if not right_side.style:
                continue

            # Keep the anchor edge as well as every cell on the visible perimeter.
            targets = [anchor]
            targets.extend(
                ws.cell(row=row, column=merged.max_col)
                for row in range(merged.min_row, merged.max_row + 1)
            )
            for cell in targets:
                border = copy.copy(cell.border)
                border.right = copy.copy(right_side)
                cell.border = border

    def _insert_continuation_header(self, ws, insert_row, source_row, row_count=6):
        """Insert a copy of an existing page header before continued table rows."""
        import copy

        max_col = ws.max_column
        cell_snapshot = []
        for r_offset in range(row_count):
            row_data = []
            for col in range(1, max_col + 1):
                cell = ws.cell(row=source_row + r_offset, column=col)
                row_data.append({
                    'value': cell.value,
                    'style': copy.copy(cell._style),
                    'number_format': cell.number_format,
                })
            cell_snapshot.append(row_data)

        row_heights = [
            ws.row_dimensions[source_row + offset].height
            for offset in range(row_count)
        ]
        source_merges = []
        for merged in list(ws.merged_cells.ranges):
            if merged.min_row >= source_row and merged.max_row < source_row + row_count:
                source_merges.append((
                    merged.min_row - source_row, merged.min_col,
                    merged.max_row - source_row, merged.max_col,
                ))

        # Images are drawing objects, so insert_rows() neither moves nor copies
        # them.  Keep snapshots of images anchored in the source header.
        source_images = []
        for image in list(ws._images):
            anchor = getattr(image, 'anchor', None)
            if not hasattr(anchor, '_from'):
                continue
            image_row = anchor._from.row + 1
            if source_row <= image_row < source_row + row_count:
                source_images.append((image, copy.deepcopy(anchor)))

        self._insert_rows_safely(ws, insert_row, row_count)

        # Move existing drawings below the insertion together with their cells.
        for image in list(ws._images):
            anchor = getattr(image, 'anchor', None)
            if not hasattr(anchor, '_from') or anchor._from.row + 1 < insert_row:
                continue
            anchor._from.row += row_count
            if hasattr(anchor, 'to'):
                anchor.to.row += row_count

        for r_offset, row_data in enumerate(cell_snapshot):
            ws.row_dimensions[insert_row + r_offset].height = row_heights[r_offset]
            for col, saved in enumerate(row_data, start=1):
                cell = ws.cell(row=insert_row + r_offset, column=col)
                cell.value = saved['value']
                cell._style = copy.copy(saved['style'])
                cell.number_format = saved['number_format']

        for min_ro, min_col, max_ro, max_col_idx in source_merges:
            ws.merge_cells(
                start_row=insert_row + min_ro, start_column=min_col,
                end_row=insert_row + max_ro, end_column=max_col_idx,
            )

        # Duplicate source-header logos/images at the continuation header.
        image_row_shift = insert_row - source_row
        for source_image, saved_anchor in source_images:
            from io import BytesIO
            from openpyxl.drawing.image import Image as XLImage

            image_buffer = BytesIO()
            image_ref = source_image.ref
            if hasattr(image_ref, 'copy') and hasattr(image_ref, 'save'):
                pil_copy = image_ref.copy()
                image_format = getattr(image_ref, 'format', None) or 'PNG'
                pil_copy.save(image_buffer, format=image_format)
            elif isinstance(image_ref, (str, os.PathLike)):
                with open(image_ref, 'rb') as image_file:
                    image_buffer.write(image_file.read())
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
            # Retain the stream for the whole workbook-save lifecycle.
            new_image._continuation_stream = image_buffer
            new_image.anchor = copy.deepcopy(saved_anchor)
            new_image.anchor._from.row += image_row_shift
            if hasattr(new_image.anchor, 'to'):
                new_image.anchor.to.row += image_row_shift
            ws.add_image(new_image)

    def _insert_table_continuation(self, ws, insert_row, document_header_row,
                                   table_header_row):
        """Start a continuation page with the document and table headers.

        ``table_header_row`` is the first of the table's two column-header
        rows; the subsection title is the row immediately above it.  The
        resulting continuation block is ten rows high:

        * seven rows of the repeated document header;
        * one subsection-title row;
        * two table-column-header rows.

        Existing content at and below ``insert_row`` is moved down intact.
        """
        import copy

        title_row = table_header_row - 1
        source_rows = range(title_row, table_header_row + 2)
        max_col = ws.max_column

        cell_snapshot = []
        for source_row in source_rows:
            row_snapshot = []
            for col in range(1, max_col + 1):
                cell = ws.cell(row=source_row, column=col)
                row_snapshot.append({
                    'value': cell.value,
                    'style': copy.copy(cell._style),
                    'number_format': cell.number_format,
                })
            cell_snapshot.append(row_snapshot)

        row_heights = [ws.row_dimensions[row].height for row in source_rows]
        source_merges = []
        for merged in list(ws.merged_cells.ranges):
            if merged.min_row >= title_row and merged.max_row <= table_header_row + 1:
                source_merges.append((
                    merged.min_row - title_row, merged.min_col,
                    merged.max_row - title_row, merged.max_col,
                ))

        self._insert_continuation_header(
            ws, insert_row, document_header_row, row_count=7
        )
        table_insert_row = insert_row + 7
        self._insert_rows_safely(ws, table_insert_row, 3)

        # Keep images below the repeated document header aligned with the rows
        # moved by the three-row subsection/table-header insertion.
        for image in list(ws._images):
            anchor = getattr(image, 'anchor', None)
            if not hasattr(anchor, '_from') or anchor._from.row + 1 < table_insert_row:
                continue
            anchor._from.row += 3
            if hasattr(anchor, 'to'):
                anchor.to.row += 3

        for offset, row_snapshot in enumerate(cell_snapshot):
            target_row = table_insert_row + offset
            ws.row_dimensions[target_row].height = row_heights[offset]
            for col, saved in enumerate(row_snapshot, start=1):
                cell = ws.cell(row=target_row, column=col)
                cell.value = saved['value']
                cell._style = copy.copy(saved['style'])
                cell.number_format = saved['number_format']

        for min_ro, min_col, max_ro, max_col in source_merges:
            ws.merge_cells(
                start_row=table_insert_row + min_ro, start_column=min_col,
                end_row=table_insert_row + max_ro, end_column=max_col,
            )

        # Mark the copied subsection as a continuation without depending on
        # a particular language or fixed title column.
        for col in range(1, max_col + 1):
            cell = ws.cell(row=table_insert_row, column=col)
            if str(cell.value or '').strip():
                cell.value = f"{cell.value} (계속)"
                break

        return 10

    def _first_print_overflow_row(self, ws, start_row, row_count):
        """Return the first data row that falls onto the next printed page."""
        break_ids = sorted(int(brk.id) for brk in ws.row_breaks.brk)
        previous_breaks = [row for row in break_ids if row < start_row]
        page_start = (previous_breaks[-1] + 1) if previous_breaks else 1

        # A4 portrait printable height, adjusted by the template's print scale.
        margins = ws.page_margins
        scale = float(ws.page_setup.scale or 100) / 100.0
        usable_inches = 11.69 - float(margins.top or 0) \
            - float(margins.bottom or 0) - float(margins.footer or 0)
        capacity_points = usable_inches * 72.0 / scale
        default_height = float(ws.sheet_format.defaultRowHeight or 15)

        used_points = 0.0
        for row in range(page_start, start_row):
            used_points += float(ws.row_dimensions[row].height or default_height)

        for row in range(start_row, start_row + row_count):
            row_height = float(ws.row_dimensions[row].height or default_height)
            if used_points + row_height > capacity_points:
                return row
            used_points += row_height
        return None

    def _find_next_document_header(self, ws, after_row, search_rows=80):
        """Locate the next real monthly-report header instead of assuming an offset."""
        end_row = min(ws.max_row, after_row + search_rows)
        for row in range(after_row + 1, end_row + 1):
            for col in range(1, min(ws.max_column, 30) + 1):
                value = str(ws.cell(row=row, column=col).value or '')
                compact = value.replace(' ', '').replace('\n', '')
                if '월간용역진도보고서' in compact:
                    return row

        # Before placeholder replacement, the contract-title row may not yet
        # contain the final report wording.  The fixed 3.0 title is four rows
        # below the beginning of the document header in this template.
        for row in range(after_row + 1, end_row + 1):
            for col in range(1, min(ws.max_column, 30) + 1):
                value = str(ws.cell(row=row, column=col).value or '')
                compact = value.replace(' ', '').replace('\n', '')
                if '3.0비파괴검사현황' in compact:
                    return max(after_row + 1, row - 4)

        # Final structural fallback: template page headers begin immediately
        # after the next explicit horizontal page break.
        later_breaks = sorted(
            int(brk.id) for brk in ws.row_breaks.brk if int(brk.id) >= after_row
        )
        if later_breaks:
            return later_breaks[0] + 1
        return None

    def _renumber_ndt_status_pages(self, ws):
        """Renumber every generated page belonging to section 3.0."""
        header_rows = []
        for row in range(1, ws.max_row + 1):
            value = str(ws.cell(row=row, column=6).value or '')
            compact = value.replace(' ', '').replace('\n', '')
            if '3.0비파괴검사현황' in compact:
                header_rows.append(row)

        total_pages = len(header_rows)
        for page_number, header_row in enumerate(header_rows, start=1):
            ws.cell(row=header_row, column=16).value = (
                f"  쪽  번 호 :      {page_number}     of     {total_pages}"
            )

    def generate_report(self, history_path, target_ym, output_path, doc_num="01", create_date=None):

        if create_date is None:
            from datetime import datetime
            create_date = str(datetime.today().date())
        """
        history_path: daily_work_history.json 경로
        target_ym: "YYYY-MM" 형식 (예: "2026-08")
        """
        if not os.path.exists(self.template_path):
            raise FileNotFoundError(f"Template not found at {self.template_path}")
            
        if not os.path.exists(history_path):
            raise FileNotFoundError(f"History file not found at {history_path}")

        with open(history_path, 'r', encoding='utf-8') as f:
            history_data = json.load(f)

        # 1. 데이터 필터링 및 집계
        target_dates = sorted([d for d in history_data.keys() if d.startswith(target_ym)])
        if not target_dates:
            print(f"No data found for {target_ym}")
            return None

        # 집계 구조 (1.2 요약용: 업체+구간+라인+관경+규격별 그룹)
        qty_summary = {}
        ndt_groups = {
            'PAUT': defaultdict(lambda: {'count': 0, 'qty': 0.0}),
            'MT': defaultdict(lambda: {'count': 0, 'qty': 0.0}),
            'RT': defaultdict(lambda: {'count': 0, 'qty': 0.0}),
            'PT': defaultdict(lambda: {'count': 0, 'qty': 0.0})
        }
        # 세부 현황용 (2.x 전체 레코드 - 날짜순)
        ndt_details = {'PAUT': [], 'MT': [], 'RT': [], 'PT': []}

        # 물량표(qty) 처리를 위해 시작일의 전일누계와 전체 금일작업 합산을 구함
        first_day_data = history_data[target_dates[0]]
        last_day_data = history_data[target_dates[-1]]

        for key, val in first_day_data.get('qty_data', {}).items():
            qty_summary[key] = {
                '예상량': val.get('예상량', ''),
                '전월누계': val.get('전일누계', '0'), # 월간 보고서이므로 첫날의 전일누계가 전월누계
                '금월작업': 0.0,
                '총누계': val.get('총누계', '0'),
                '공정률': val.get('공정률', ''),
                '불량': 0,
                '불량률': ''
            }

        # 날짜별 금일작업 누적
        for d in target_dates:
            day_data = history_data[d]
            for key, val in day_data.get('qty_data', {}).items():
                if key not in qty_summary:
                    qty_summary[key] = {'예상량': val.get('예상량', ''), '전월누계': '0', '금월작업': 0.0, '총누계': '0', '공정률': '', '불량': 0, '불량률': ''}
                
                try:
                    today_work = float(str(val.get('금일작업', '0')).replace(',', '') or 0)
                    qty_summary[key]['금월작업'] += today_work
                except ValueError:
                    pass
                
                # 불량(결함) 수량도 누적 (숫자일 경우)
                try:
                    defect = float(str(val.get('불량', '0')).replace(',', '') or 0)
                    qty_summary[key]['불량'] += defect
                except ValueError:
                    pass

            # NDT 결과 집계
            for row in day_data.get('ndt_results', []):
                method = str(row.get('검사방법', '')).strip().upper()
                if not method or method not in ndt_groups:
                    continue
                
                company = str(row.get('업체', '')).strip()
                if not company: continue # 빈 행 무시
                
                section = str(row.get('구간', '')).strip()
                line_no = str(row.get('라인번호', '')).strip()
                size = str(row.get('관경', '')).strip()
                spec = str(row.get('규격', '')).strip()
                
                # 수량 (길이 또는 매수)
                qty_val = 0.0
                if method == 'PAUT': qty_val = self._safe_float(row.get('PAUT'))
                elif method == 'MT': qty_val = self._safe_float(row.get('MT'))
                elif method == 'PT': qty_val = self._safe_float(row.get('PT'))
                elif method == 'RT':
                    # RT는 OR + RE 매수 합산
                    or_val = self._safe_float(row.get('RT_OR'))
                    re_val = self._safe_float(row.get('RT_RE'))
                    qty_val = or_val + re_val
                
                group_key = (company, section, line_no, size, spec)
                ndt_groups[method][group_key]['count'] += 1
                ndt_groups[method][group_key]['qty'] += qty_val

                # 2.x 세부현황용: 개별 레코드 저장 (날짜 포함)
                if method == 'RT':
                    d_ori = self._safe_float(row.get('RT_OR'))
                    d_re  = self._safe_float(row.get('RT_RE'))
                else:
                    d_ori = qty_val
                    d_re  = 0.0
                if d_ori + d_re > 0:
                    ndt_details[method].append({
                        'key': group_key,
                        'ori': d_ori, 're': d_re, 'qty': d_ori + d_re
                    })

        # 마지막 날의 총누계와 공정률을 최종 반영
        for key, val in last_day_data.get('qty_data', {}).items():
            if key in qty_summary:
                qty_summary[key]['총누계'] = val.get('총누계', '0')
                qty_summary[key]['공정률'] = val.get('공정률', '')
                if val.get('불량률'):
                    qty_summary[key]['불량률'] = val.get('불량률', '')

        # 2. 엑셀 쓰기 준비
        wb = openpyxl.load_workbook(self.template_path)
        
        # 엑셀 열 때 '명명된 범위' 복구 창이 뜨는 것을 방지하기 위해 불필요한 이름 정의 삭제
        if hasattr(wb, 'defined_names'):
            try:
                wb.defined_names.clear()
            except:
                pass
                
        # 외부 수식 참조(externalLink) 찌꺼기로 인한 복구 창 방지
        if hasattr(wb, '_external_links'):
            wb._external_links = []
            
        # [표지] 시트의 병합 셀이 openpyxl 저장 시 풀리는 버그 방지를 위한 백업
        ws = wb.active

        # The template already has the page 6/7 boundary at row 229. Insert
        # the missing page 5/6 boundary at row 192 without replacing row 229
        # or any later boundary (including the page 7/8 break at row 278).
        from openpyxl.worksheet.pagebreak import Break
        template_breaks = sorted(ws.row_breaks.brk, key=lambda brk: int(brk.id))
        if not any(int(brk.id) == 192 for brk in template_breaks):
            template_breaks.append(Break(id=192))
        ws.row_breaks.brk = sorted(
            template_breaks, key=lambda brk: int(brk.id)
        )

        # Capture the template's compact body-row height before dynamic inserts
        # move row 466. Header rows keep their original heights.
        detail_body_row_height = (
            ws.row_dimensions[466].height
            or ws.sheet_format.defaultRowHeight
            or 15
        )




        # 역순으로 채워야 행 삽입 시 위에 있는 인덱스(row number)가 변하지 않음
        # 따라서 아래에서부터 표를 처리합니다.
        
        # 표 위치들을 리스트로 묶어서 역순 처리
        # is_detail=True: 2.x 세부(전체 레코드), False: 1.x 요약(그룹)
        sections = [
            ('pt_2', 'PT', True), ('rt_2', 'RT', True), ('mt_2', 'MT', True), ('paut_2', 'PAUT', True),
            ('pt_1', 'PT', False), ('rt_1', 'RT', False), ('mt_1', 'MT', False), ('paut_1', 'PAUT', False)
        ]
        section_item_counts = {}
        
        for section_key, method, is_detail in sections:
            header_row = self.table_markers[section_key]
            start_row = header_row + 2
            
            # TOTAL 행을 찾음 (header_row 이후 빈 행이 아닌 "TOTAL"이 있는 행)
            total_row = start_row
            for r in range(start_row, start_row + 20):
                val = str(ws.cell(row=r, column=2).value or "").strip()
                if "TOTAL" in val.upper() or "총계" in val:
                    total_row = r
                    break
            
            # 데이터 삽입: 2.x는 전체 레코드, 1.x는 요약 그룹
            if is_detail:
                data_items = [
                    (rec['key'], {'count': 1, 'qty': rec['qty'], 'ori': rec['ori'], 're': rec['re']})
                    for rec in ndt_details[method]
                ]
            else:
                data_items = [
                    (k, {'count': v['count'], 'qty': v['qty'], 'ori': v['qty'], 're': 0.0})
                    for k, v in ndt_groups[method].items()
                ]
            num_items = len(data_items)
            section_item_counts[section_key] = num_items
            
            def _write_safe(r, c, val):
                cell = ws.cell(row=r, column=c)
                if type(cell).__name__ == 'MergedCell':
                    new_cell = Cell(ws, row=r, column=c)
                    ws._cells[(r, c)] = new_cell
                    cell = new_cell
                cell.value = val

            def safe_merge(sr, sc, er, ec):
                from openpyxl.utils import get_column_letter
                coord = f"{get_column_letter(sc)}{sr}:{get_column_letter(ec)}{er}"
                overlaps = []
                for m in list(ws.merged_cells.ranges):
                    # Check for overlap
                    if m.bounds[0] <= ec and m.bounds[2] >= sc and m.bounds[1] <= er and m.bounds[3] >= sr:
                        if str(m) == coord:
                            return # Already perfectly merged
                        overlaps.append(m)
                
                # Unmerge any overlapping regions first to prevent Excel file corruption
                for m in overlaps:
                    try:
                        ws.unmerge_cells(str(m))
                    except:
                        pass
                
                try: ws.merge_cells(coord)
                except: pass
            
            # 기존에 2개의 빈 행이 있다고 가정하고(TOTAL 위에), 모자라면 삽입
            existing_empty_rows = total_row - start_row
            
            if num_items > existing_empty_rows:
                rows_to_insert = num_items - existing_empty_rows
                # 안전한 행 삽입: 아래 시트(13페이지 등) 병합 셀 구조 보존
                self._insert_rows_safely(ws, start_row, rows_to_insert)
                total_row += rows_to_insert
                # 모든 기존 마커 업데이트 (삽입된 행 아래에 있는 마커들만 += rows_to_insert)
                for k, v in self.table_markers.items():
                    if v >= start_row:
                        self.table_markers[k] += rows_to_insert

                # 버퍼 행 정리: TOTAL 이후 빈 행을 rows_to_insert만큼 삭제
                # 로고/쪽번호 헤더 블록이 TOTAL 바로 다음에 위치하도록 당김
                scan_row = total_row + 1
                cleaned = 0
                while cleaned < rows_to_insert:
                    is_empty = all(
                        not str(ws.cell(row=scan_row, column=c).value or '').strip()
                        for c in range(1, 25)
                    )
                    if is_empty:
                        self._delete_rows_safely(ws, scan_row, 1)
                        total_row -= 1  # 삭제로 없어진 행 보정
                        for k, v in self.table_markers.items():
                            if v >= scan_row:
                                self.table_markers[k] -= 1
                        cleaned += 1
                        # scan_row 유지 (삭제 후 다음 행이 같은 위치로 올라옴)
                    else:
                        break  # 내용 있는 행(헤더 블록) 만나면 중지
            else:
                rows_to_insert = 0

            current_row = start_row
            for idx, ((comp, sec, line, size, spec), vals) in enumerate(data_items):
                count = vals['count']
                qty = vals['qty']
                unit = "매" if method == "RT" else "m"

                # Clear reusable template values before writing this detail row.
                for clear_col in range(17, 24):
                    clear_cell = ws.cell(row=current_row, column=clear_col)
                    if type(clear_cell).__name__ != 'MergedCell':
                        clear_cell.value = None
                
                # 셀 위치: 2:업체, 3:순번, 4:Section, 6:Line No., 10:관경, 12:용접개소, 14:규격, 16:단위, 17:길이
                ws.cell(row=current_row, column=2).value = comp
                ws.cell(row=current_row, column=3).value = str(idx+1)
                ws.cell(row=current_row, column=4).value = sec
                ws.cell(row=current_row, column=6).value = line
                ws.cell(row=current_row, column=10).value = size
                ws.cell(row=current_row, column=12).value = count
                ws.cell(row=current_row, column=14).value = spec
                ws.cell(row=current_row, column=16).value = unit
                
                ws.cell(row=current_row, column=16).value = unit

                ori_v = vals.get('ori', qty)
                re_v  = vals.get('re', 0.0)
                tot_v = ori_v + re_v

                def _fmt(v):
                    return f"{v:.4f}" if v % 1 != 0 else str(int(v))

                _write_safe(current_row, 17, _fmt(ori_v) if ori_v > 0 else '')
                if re_v > 0:
                    _write_safe(current_row, 18, _fmt(re_v))  # RE' 앵커(col18)
                # PAUT는 항상 TOTAL, 세부현황도 TOTAL 표시
                if method == 'PAUT' or is_detail:
                    _write_safe(current_row, 20, _fmt(tot_v) if tot_v > 0 else '')
                
                # 병합 처리 (템플릿 양식에 맞춤)
                safe_merge(current_row, 4, current_row, 5) # Section
                safe_merge(current_row, 6, current_row, 9) # Line No.
                safe_merge(current_row, 10, current_row, 11) # 관경
                safe_merge(current_row, 12, current_row, 13) # 용접개소
                safe_merge(current_row, 14, current_row, 15) # 규격
                
                # PAUT/MT/RT/PT 모두 동일: RE'(col18~19)만 병합, ORI'(17)·TOTAL(20)은 개별 셀
                safe_merge(current_row, 18, current_row, 19) # RE'
                safe_merge(current_row, 20, current_row, 21) # TOTAL
                safe_merge(current_row, 22, current_row, 23) # 비고
                    
                # Line No. 글씨 다 보이도록 높이 조절
                if current_row >= 401:
                    ws.row_dimensions[current_row].height = detail_body_row_height

                # 폰트, 정렬 적용 (병합된 칸 전체에 정렬 속성을 먹여야 엑셀이 줄바꿈을 정상 인식함)
                for c_idx in range(2, 24):
                    c_cell = ws.cell(row=current_row, column=c_idx)
                    c_cell.font = self.font_normal
                    c_cell.alignment = self.align_center
                    c_cell.border = self.border_thin

                # Keep long Line No. values on one line inside the widened F:I area.
                ws.cell(row=current_row, column=6).alignment = Alignment(
                    horizontal='center', vertical='center',
                    wrap_text=False, shrink_to_fit=True
                )

                current_row += 1

            # Dynamic row operations can move the TOTAL row; locate it again.
            for candidate_row in range(current_row, current_row + 40):
                candidate_value = str(
                    ws.cell(row=candidate_row, column=2).value or ''
                ).strip().upper()
                if 'TOTAL' in candidate_value:
                    total_row = candidate_row
                    break

            if num_items == 0:
                # Keep the method visible while clearly distinguishing "no
                # activity" from a missing or failed data load.
                safe_merge(start_row, 2, start_row, 23)
                no_data_cell = ws.cell(row=start_row, column=2)
                no_data_cell.value = '해당 월 검사실적 없음'
                no_data_cell.font = self.font_normal
                no_data_cell.alignment = self.align_center
                no_data_cell.border = self.border_thin
                if start_row >= 401:
                    ws.row_dimensions[start_row].height = detail_body_row_height
                current_row = start_row + 1
                
            # 빈 행(데이터 없는 행)들 스타일 정리 및 PAUT 병합 처리 준비
            for r in range(current_row, total_row):
                if r >= 401:
                    ws.row_dimensions[r].height = detail_body_row_height
                # 빈 행도 기본적으로 다른 칸들도 병합을 유지해야 템플릿 양식이 안 깨짐
                safe_merge(r, 4, r, 5)
                safe_merge(r, 6, r, 9)
                safe_merge(r, 10, r, 11)
                safe_merge(r, 12, r, 13)
                safe_merge(r, 14, r, 15)
                
                # 빈 행도 RE'(col18~19)만 병합 (ORI'·TOTAL 개별 유지)
                safe_merge(r, 18, r, 19)
                safe_merge(r, 20, r, 21) # TOTAL
                safe_merge(r, 22, r, 23) # 비고
                for c_idx in range(2, 24):
                    cell = ws.cell(row=r, column=c_idx)
                    cell.border = self.border_thin

            # TOTAL 행 값 쓰기 및 스타일 보정
            total_ori = sum(vals.get('ori', vals['qty']) for _, vals in data_items)
            total_re = sum(vals.get('re', 0.0) for _, vals in data_items)
            total_qty = total_ori + total_re
            
            # TOTAL 행의 17, 18, 19, 20 칸 배경색(파란색)을 동일하게 맞춤 (첫 번째 셀 서식 복사)
            fill_style = ws.cell(row=total_row, column=2).fill
            if method == 'PAUT':
                import copy
                for c_idx in range(17, 21):
                    c = ws.cell(row=total_row, column=c_idx)
                    if fill_style:
                        c.fill = copy.copy(fill_style)

            safe_merge(total_row, 22, total_row, 23) # 비고
            if total_row >= 401:
                ws.row_dimensions[total_row].height = detail_body_row_height

            if num_items == 0:
                _write_safe(total_row, 17, 0)
                _write_safe(total_row, 20, 0)
                ws.cell(row=total_row, column=17).alignment = self.align_center
                ws.cell(row=total_row, column=20).alignment = self.align_center

            if total_qty > 0:
                from openpyxl.styles import Alignment as _Align
                shrink_align = _Align(horizontal='center', vertical='center', shrink_to_fit=True)
                if method == 'PAUT':
                    safe_merge(total_row, 20, total_row, 21)
                    _write_safe(total_row, 17, f"{total_ori:.4f}" if total_ori % 1 != 0 else str(int(total_ori)))
                    ws.cell(row=total_row, column=17).alignment = shrink_align
                    if total_re > 0:
                        _write_safe(total_row, 18, f"{total_re:.4f}" if total_re % 1 != 0 else str(int(total_re)))
                        ws.cell(row=total_row, column=18).alignment = shrink_align
                    _write_safe(total_row, 20, f"{total_qty:.4f}" if total_qty % 1 != 0 else str(int(total_qty)))
                    ws.cell(row=total_row, column=20).alignment = shrink_align
                else:
                    _write_safe(total_row, 17, f"{total_qty:.4f}" if total_qty % 1 != 0 else str(int(total_qty)))
                    ws.cell(row=total_row, column=17).alignment = shrink_align

            # When either a 1.2 summary table or a 2.x detail table overflows,
            # continue it on a real new page. Repeat both the seven-row document
            # header and the subsection/table header.  Repeat as many times as
            # necessary; following MT/RT/PT/4.0 content moves down as a block.
            remaining_start = start_row
            remaining_count = num_items
            previous_section_breaks = sorted(
                int(brk.id)
                for brk in ws.row_breaks.brk
                if int(brk.id) < header_row
            )
            document_header_row = (
                previous_section_breaks[-1] + 1
                if previous_section_breaks else 1
            )
            while remaining_count > 0:
                continuation_row = self._first_print_overflow_row(
                    ws, remaining_start, remaining_count
                )
                if continuation_row is None:
                    break

                rows_on_current_page = continuation_row - remaining_start
                inserted_rows = self._insert_table_continuation(
                    ws, continuation_row, document_header_row, header_row
                )
                from openpyxl.worksheet.pagebreak import Break
                if not any(
                    int(brk.id) == continuation_row - 1
                    for brk in ws.row_breaks.brk
                ):
                    ws.row_breaks.append(Break(id=continuation_row - 1))

                for key, marker_row in self.table_markers.items():
                    if marker_row >= continuation_row:
                        self.table_markers[key] += inserted_rows

                remaining_count -= rows_on_current_page
                remaining_start = continuation_row + inserted_rows

        # When the three method-detail tables on this page have no records,
        # distribute the unused vertical space between them instead of leaving
        # all tables crowded at the top.
        if all(section_item_counts.get(key, 0) == 0 for key in ('mt_2', 'rt_2', 'pt_2')):
            for current_key, next_key in (('mt_2', 'rt_2'), ('rt_2', 'pt_2')):
                current_marker = self.table_markers[current_key]
                next_marker = self.table_markers[next_key]
                current_total = None
                for row in range(current_marker + 2, next_marker):
                    if 'TOTAL' in str(ws.cell(row=row, column=2).value or '').upper():
                        current_total = row
                        break
                if current_total is None:
                    continue
                blank_spacer_rows = []
                for spacer_row in range(current_total + 1, next_marker):
                    is_blank = all(
                        not str(ws.cell(row=spacer_row, column=col).value or '').strip()
                        for col in range(1, 24)
                    )
                    if is_blank:
                        blank_spacer_rows.append(spacer_row)
                if blank_spacer_rows:
                    # Equal physical gap between method tables regardless of
                    # how many blank template rows each gap contains.
                    # Keep both gaps equal while ensuring rows 472:522 remain
                    # within one printed page ending at the row-522 break.
                    row_height = 48.0 / len(blank_spacer_rows)
                    for spacer_row in blank_spacer_rows:
                        ws.row_dimensions[spacer_row].height = row_height

        # 마지막으로 맨 위의 1.1 물량표 처리
        qty_start_row = self.table_markers['qty'] + 1
        qty_mapping = {
            'PAUT_300A이상': qty_start_row,
            'PAUT_300A이상-야간': qty_start_row + 1,
            'PAUT_250A': qty_start_row + 2,
            'PAUT_200A': qty_start_row + 3,
            'PAUT_200A-야간': qty_start_row + 4,
            'PAUT_소계': qty_start_row + 5,
            
            'RT_150A~100A': qty_start_row + 6,
            'RT_150A~100A-야간': qty_start_row + 7,
            'RT_80A이하': qty_start_row + 8,
            'RT_80A이하-야간': qty_start_row + 9,
            'RT_소계': qty_start_row + 10,
            
            'MT_전체(주간)': qty_start_row + 11,
            'MT_전체(야간)': qty_start_row + 12,
            
            'PT_전체(주간)': qty_start_row + 13,
            'PT_전체(야간)': qty_start_row + 14,
        }

        for key, r_idx in qty_mapping.items():
            if key in qty_summary:
                data = qty_summary[key]
                # 컬럼: 2:방법, 4:규격, 7:예상량, 9:전월누계, 11:금월작업, 13:총누계, 15:공정률, 17:불량, 19:불량률
                ws.cell(row=r_idx, column=7).value = self._format_num(data['예상량'])
                ws.cell(row=r_idx, column=9).value = self._format_num(data['전월누계'])
                ws.cell(row=r_idx, column=11).value = self._format_num(data['금월작업'])
                ws.cell(row=r_idx, column=13).value = self._format_num(data['총누계'])
                ws.cell(row=r_idx, column=15).value = data['공정률']
                ws.cell(row=r_idx, column=17).value = self._format_num(data['불량'])
                ws.cell(row=r_idx, column=19).value = data['불량률']

        # 사용자 요청: 관경/규격 축소, 단위/비고 확장
        ws.column_dimensions['C'].width = 4.0   # 순번
        ws.column_dimensions['J'].width = 5.0   # 관경 (8.0 → 5.0)
        ws.column_dimensions['N'].width = 2.5   # 규격 (3.5 → 2.5)
        ws.column_dimensions['O'].width = 8.0   # 단위 (5.6 → 8.0)
        ws.column_dimensions['P'].width = 8.0   # 비고 (5.0 → 8.0)
        # ORI(Q)=9, RE(R+S)=4.5+4.5=9, TOTAL(T+U)=4.5+4.5=9 → 동일 너비
        # Rebalance widths so Line No. fits on one line without widening the page.
        for col in ('F', 'G', 'H', 'I'):
            ws.column_dimensions[col].width = 6.5
        for col in ('J', 'K'):
            ws.column_dimensions[col].width = 3.5
        for col in ('L', 'M'):
            ws.column_dimensions[col].width = 3.0
        for col in ('N', 'O'):
            ws.column_dimensions[col].width = 3.5
        ws.column_dimensions['P'].width = 5.0

        ws.column_dimensions['Q'].width = 9.0
        ws.column_dimensions['R'].width = 4.5
        ws.column_dimensions['S'].width = 4.5
        ws.column_dimensions['T'].width = 4.5
        ws.column_dimensions['U'].width = 4.5
        ws.column_dimensions['V'].width = 0.0   # 비고 우측 끝 숨김
        ws.column_dimensions['V'].width = 2.25
        ws.column_dimensions['V'].hidden = False
        ws.column_dimensions['W'].width = 0.0   # 비고 우측 끝 숨김
        ws.column_dimensions['W'].width = 2.25
        ws.column_dimensions['W'].hidden = False

        # 사용자 요청: 모든 표의 좌측(2열), 우측(21열) 외곽선을 굵은 실선으로 복원/강제 설정
        medium = Side(style='medium')
        def apply_outer_borders(start_r, end_r):
            for r in range(start_r, end_r + 1):
                lc = ws.cell(row=r, column=2)
                lc.border = Border(left=medium, right=lc.border.right, top=lc.border.top, bottom=lc.border.bottom)
                rc = ws.cell(row=r, column=23)
                rc.border = Border(left=rc.border.left, right=medium, top=rc.border.top, bottom=rc.border.bottom)
                
        # 1. 상단 물량표 외곽선 교정
        if self.table_markers['qty'] > 0:
            apply_outer_borders(self.table_markers['qty'] + 1, self.table_markers['qty'] + 14)
            
        # 2. 하단 데이터 표 외곽선 교정 (헤더 제외, 데이터 행부터 TOTAL 전까지)
        for section_key in ['paut_1', 'paut_2', 'rt_1', 'rt_2', 'mt_1', 'mt_2', 'pt_1', 'pt_2']:
            if section_key not in self.table_markers: continue
            
            s_row = self.table_markers[section_key]
            if s_row > 0:
                e_row = s_row
                while True:
                    val = str(ws.cell(row=e_row, column=2).value).strip().upper()
                    if 'TOTAL' in val or '계' == val or e_row > s_row + 200:
                        break
                    e_row += 1
                apply_outer_borders(s_row, e_row - 1)

        # 문자열 치환 (문서번호, 연월, 날짜, 지사명 등)
        y, m = target_ym.split('-')
        import datetime
        replacements = {
            '[[보고서_연월]]': f"{y}년 {int(m)}월",
            '[[보고서_월]]': f"{int(m)}월",
            '[[문서번호]]': doc_num,
            '[[작성일자]]': create_date,
            '[[계약명]]': "2026년 중앙지사 열수송관 비파괴검사용역 단가계약",
            '[[지사명]]': "중앙지사",
            '2025년 동탄지사 열수송관 비파괴검사용역 단가계약': "2026년 중앙지사 열수송관 비파괴검사용역 단가계약",
            '2025년 동탄지사 열배관  비파괴검사용역 단가계약': "2026년 중앙지사 열수송관 비파괴검사용역 단가계약",
            '2025년 동탄지사 열배관 비파괴검사용역 단가계약': "2026년 중앙지사 열수송관 비파괴검사용역 단가계약",
            '동 탄 지 사': "중 앙 지 사",
            '분 당 사 업 소': "중 앙 지 사"
        }
        for sheet in wb.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value:
                        val_str = str(cell.value)
                        
                        if isinstance(cell.value, str):
                            for k, v in replacements.items():
                                if k in val_str:
                                    cell.value = val_str.replace(k, str(v))
                                    val_str = str(cell.value)
                                    
                        # 타이틀 텍스트 무조건 포맷 강제 적용 (모든 페이지의 제목 칸 타겟)
                        val_no_spaces = val_str.replace(' ', '').replace('\xa0', '').replace('\n', '')
                        
                        is_title = False
                        if ('단가계약' in val_no_spaces or '계약명' in val_no_spaces) and ('월간용역' in val_no_spaces or '보고서' in val_no_spaces):
                            is_title = True
                        elif cell.column == 6 and isinstance(cell.value, str) and val_no_spaces.startswith('='):
                            # 수식으로 제목을 참조하는 경우 (예: =F41, =F$41, =$F$41 등)
                            match = re.search(r'\$?F\$?(\d+)', val_no_spaces, re.IGNORECASE)
                            if match:
                                ref_row = int(match.group(1))
                                # 참조된 셀의 실제 값이 제목 키워드를 포함하는지 확인
                                ref_val = str(sheet.cell(row=ref_row, column=6).value or '').replace(' ', '').replace('\xa0', '').replace('\n', '')
                                if ('단가계약' in ref_val or '계약명' in ref_val) and ('월간용역' in ref_val or '보고서' in ref_val):
                                    is_title = True
                                # 참조 셀에 값이 없는 경우 행 번호 패턴으로 폴백 (41, 81, 121... 등 40행 간격)
                                elif ref_row % 40 == 1:
                                    is_title = True
                                    
                        if is_title:
                            contract_name = replacements.get('[[계약명]]', '2026년 중앙지사 열수송관 비파괴검사용역 단가계약')
                            # 무조건 괄호와 줄바꿈이 포함된 포맷으로 강제 덮어쓰기 (수식도 문자열로 덮어씌움)
                            cell.value = f"【{contract_name}】\n월 간 용 역 진 도 보 고 서"
                            
                            # Alignment and Font are imported at module level.
                            import copy
                            
                            title_font = Font(name='맑은 고딕', size=9, bold=True)
                            title_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
                            cell.font = title_font
                            cell.alignment = title_align
                            
                            for r_idx in range(cell.row, cell.row + 4):
                                for c_idx in range(1, 30):
                                    c = sheet.cell(row=r_idx, column=c_idx)
                                    if c.alignment:
                                        new_align = copy.copy(c.alignment)
                                        new_align.wrap_text = True
                                        c.alignment = new_align
                                    else:
                                        c.alignment = Alignment(wrap_text=True)


        # [표지] 시트 병합 셀 강제 복구 (openpyxl 버그 방지)
        # Dynamic continuation pages change the section total, so recalculate
        # the 3.0 page labels after all row and page operations are complete.
        self._renumber_ndt_status_pages(ws)

        # Preserve the visible right edge of merged template boxes and headers.
        for sheet in wb.worksheets:
            self._repair_merged_right_borders(sheet)

        # Prevent blank right-side pages caused by the widened V:W remarks area.
        # Keep vertical pagination automatic while fitting the report to one
        # printed page horizontally.
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.page_setup.scale = None

        # Hide the original trailing spacer rows only while they are still
        # blank. Dynamic continuation pages may move real content into these
        # coordinates, which must never be hidden.
        for blank_row in range(570, 573):
            is_blank = all(
                not str(ws.cell(row=blank_row, column=col).value or '').strip()
                for col in range(1, min(ws.max_column, 23) + 1)
            )
            if is_blank:
                ws.row_dimensions[blank_row].height = 0
                ws.row_dimensions[blank_row].hidden = True

        wb.save(output_path)
        return output_path

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
