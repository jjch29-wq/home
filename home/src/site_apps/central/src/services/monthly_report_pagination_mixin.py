import openpyxl
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
from openpyxl.cell.cell import Cell, MergedCell
import os

# [FIX] Patch openpyxl MergedCell to prevent 'read-only' error when assigning to merged cells
if not hasattr(MergedCell, '_patched_value'):
    MergedCell._patched_value = True
    MergedCell.value = property(lambda self: None, lambda self, val: None)

import json
import datetime
from collections import defaultdict
import re



class MonthlyReportPaginationMixin:
    @staticmethod
    def _keep_single_lower_horizontal_border(cell):
        """Draw an internal row boundary from the upper row only.

        Excel/PDF renderers can separate coincident bottom and top borders by
        a pixel after scaling.  Detail rows therefore keep their bottom edge
        and omit the top edge; the preceding row supplies the one shared line.
        """
        import copy

        border = copy.copy(cell.border)
        border.top = Side(style=None)
        cell.border = border

    @staticmethod
    def _bridge_blank_detail_row_borders(ws):
        """Continue vertical grid lines through tiny pagination remnants."""
        import copy

        bridged = []
        for row in range(2, ws.max_row):
            dimension = ws.row_dimensions[row]
            height = float(
                dimension.height
                if dimension.height is not None
                else (ws.sheet_format.defaultRowHeight or 15)
            )
            if height > 2.1 or dimension.hidden:
                continue
            if any(ws.cell(row=row, column=col).value not in (None, '')
                   for col in range(2, min(ws.max_column, 23) + 1)):
                continue

            previous = row - 1
            while previous > 0 and previous >= row - 3 and not str(
                ws.cell(previous, 3).value or ''
            ).strip():
                previous -= 1
            following = row + 1
            while following <= ws.max_row and following <= row + 3 and not str(
                ws.cell(following, 3).value or ''
            ).strip():
                following += 1
            try:
                previous_seq = int(str(ws.cell(previous, 3).value).strip())
                following_seq = int(str(ws.cell(following, 3).value).strip())
            except (TypeError, ValueError):
                continue
            if following_seq != previous_seq + 1:
                continue

            for col in range(2, min(ws.max_column, 23) + 1):
                source = ws.cell(previous, col).border
                target = ws.cell(row, col)
                border = copy.copy(target.border)
                border.left = copy.copy(source.left)
                border.right = copy.copy(source.right)
                border.top = Side(style=None)
                border.bottom = (
                    copy.copy(source.bottom)
                    if row + 1 == following else Side(style=None)
                )
                target.border = border
            bridged.append(row)
        return bridged

    @staticmethod
    def _collapse_blank_detail_row_gaps(ws):
        """Hide tiny blank remnants between consecutive detail records."""
        collapsed = []
        row = 2
        while row < ws.max_row:
            dimension = ws.row_dimensions[row]
            height = float(
                dimension.height
                if dimension.height is not None
                else (ws.sheet_format.defaultRowHeight or 15)
            )
            is_blank = not any(
                str(ws.cell(row=row, column=col).value or '').strip()
                for col in range(2, min(ws.max_column, 23) + 1)
            )
            if dimension.hidden or height > 2.1 or not is_blank:
                row += 1
                continue

            start = row
            end = row
            while end + 1 < ws.max_row and end - start < 2:
                next_dim = ws.row_dimensions[end + 1]
                next_height = float(
                    next_dim.height
                    if next_dim.height is not None
                    else (ws.sheet_format.defaultRowHeight or 15)
                )
                next_blank = not any(
                    str(ws.cell(row=end + 1, column=col).value or '').strip()
                    for col in range(2, min(ws.max_column, 23) + 1)
                )
                if next_dim.hidden or next_height > 2.1 or not next_blank:
                    break
                end += 1

            try:
                previous_seq = int(str(ws.cell(start - 1, 3).value).strip())
                following_seq = int(str(ws.cell(end + 1, 3).value).strip())
            except (TypeError, ValueError):
                row = end + 1
                continue
            if following_seq != previous_seq + 1:
                row = end + 1
                continue

            for blank_row in range(start, end + 1):
                ws.row_dimensions[blank_row].height = 0
                ws.row_dimensions[blank_row].hidden = True
                collapsed.append(blank_row)
            row = end + 1
        return collapsed

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

        # openpyxl bug: manually shift row_dimensions down
        import copy
        existing_rows = sorted(ws.row_dimensions.keys(), reverse=True)
        for r in existing_rows:
            if r >= insert_idx:
                new_dim = copy.copy(ws.row_dimensions[r])
                new_dim.index = r + amount
                new_dim.worksheet = ws
                ws.row_dimensions[r + amount] = new_dim
                del ws.row_dimensions[r]
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
                raw_pa = str(ws.print_area).split('!')[-1].replace('$', '')
                pa = CellRange(raw_pa)
                if pa.max_row >= insert_idx:
                    pa.max_row += amount
                    ws.print_area = f"$B$1:$W${pa.max_row}"
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

        # openpyxl bug: manually shift row_dimensions up
        import copy
        existing_rows = sorted(ws.row_dimensions.keys())
        for r in existing_rows:
            if r >= delete_idx and r < delete_idx + amount:
                del ws.row_dimensions[r]
            elif r >= delete_idx + amount:
                new_dim = copy.copy(ws.row_dimensions[r])
                new_dim.index = r - amount
                new_dim.worksheet = ws
                ws.row_dimensions[r - amount] = new_dim
                del ws.row_dimensions[r]
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

        # Keep drawings aligned with the rows that remain. Repeated report
        # headers contain copied logos, so drawings wholly inside the deleted
        # block must disappear with that header.
        kept_images = []
        delete_end = delete_idx + amount - 1
        for image in list(ws._images):
            anchor = getattr(image, 'anchor', None)
            if not hasattr(anchor, '_from'):
                kept_images.append(image)
                continue
            image_start = anchor._from.row + 1
            image_end = (
                anchor.to.row + 1 if hasattr(anchor, 'to') else image_start
            )
            if image_start >= delete_idx and image_end <= delete_end:
                continue
            if image_start >= delete_idx + amount:
                anchor._from.row -= amount
                if hasattr(anchor, 'to'):
                    anchor.to.row -= amount
            kept_images.append(image)
        ws._images = kept_images

        # Keep the print area aligned with rows removed below it. Without this
        # adjustment, deleted continuation headings leave the old lower
        # boundary behind and Excel can print a trailing blank page.
        if ws.print_area:
            try:
                raw_area = str(ws.print_area).split('!')[-1].replace('$', '')
                area = CellRange(raw_area)
                if area.max_row >= delete_idx + amount:
                    area.max_row -= amount
                    ws.print_area = f"$B$1:$W${area.max_row}"
                elif area.max_row >= delete_idx:
                    area.max_row = max(area.min_row, delete_idx - 1)
                    ws.print_area = f"$B$1:$W${area.max_row}"
            except Exception:
                pass

    def _pull_rt_summary_total_to_previous_page(self, ws):
        """Use the free first-page row for RT TOTAL and drop its continuation."""
        import copy
        from openpyxl.worksheet.pagebreak import Break

        max_col = min(ws.max_column, 23)

        def row_text(row):
            return ''.join(
                str(ws.cell(row=row, column=col).value or '')
                for col in range(1, max_col + 1)
            )

        for page_end in sorted(int(brk.id) for brk in ws.row_breaks.brk):
            search_end = min(ws.max_row, page_end + 20)
            header_row = next((
                row for row in range(page_end + 1, search_end + 1)
                if 'LINENO' in ''.join(row_text(row).split()).replace('.', '').upper()
                and row > page_end + 1
                and '1.2.3' in row_text(row - 1)
            ), None)
            continuation = header_row - 1 if header_row is not None else None
            if continuation is None:
                continue
            total_row = next((
                row for row in range(
                    continuation + 1, min(search_end, continuation + 6) + 1
                )
                if 'TOTAL' in str(ws.cell(row=row, column=2).value or '').upper()
            ), None)
            if total_row is None or total_row - continuation != 3:
                continue

            source_cells = []
            for col in range(1, ws.max_column + 1):
                cell = ws.cell(total_row, col)
                source_cells.append((
                    cell.value,
                    copy.copy(cell._style),
                    copy.copy(cell.alignment),
                    cell.number_format,
                ))
            source_height = ws.row_dimensions[total_row].height
            source_merges = [
                (merged.min_col, merged.max_col)
                for merged in list(ws.merged_cells.ranges)
                if merged.min_row == total_row == merged.max_row
            ]

            insert_row = page_end + 1
            self._insert_rows_safely(ws, insert_row, 1)
            for image in ws._images:
                anchor = getattr(image, 'anchor', None)
                if hasattr(anchor, '_from') and anchor._from.row + 1 >= insert_row:
                    anchor._from.row += 1
                    if hasattr(anchor, 'to'):
                        anchor.to.row += 1
            for col, (value, style, alignment, number_format) in enumerate(
                source_cells, start=1
            ):
                cell = ws.cell(insert_row, col)
                cell.value = value
                cell._style = copy.copy(style)
                cell.alignment = copy.copy(alignment)
                cell.number_format = number_format
            ws.row_dimensions[insert_row].height = source_height
            ws.row_dimensions[insert_row].hidden = False
            for min_col, max_col_idx in source_merges:
                ws.merge_cells(
                    start_row=insert_row, start_column=min_col,
                    end_row=insert_row, end_column=max_col_idx,
                )

            ws.row_breaks.brk = [
                Break(id=page_end + 1) if int(brk.id) == page_end else brk
                for brk in ws.row_breaks.brk
            ]

            delete_start = continuation + 1
            delete_count = total_row - continuation + 1
            self._delete_rows_safely(ws, delete_start, delete_count)
            for key, marker in self.table_markers.items():
                if marker >= insert_row:
                    self.table_markers[key] += 1
                if marker >= delete_start + delete_count:
                    self.table_markers[key] -= delete_count
            return True
        return False

    def _remove_header_only_ndt_pages(self, ws):
        """Delete empty 3.0 pages that contain only repeated captions."""
        from openpyxl.worksheet.pagebreak import Break

        max_col = min(ws.max_column, 23)

        def row_text(row):
            return ''.join(
                str(ws.cell(row=row, column=col).value or '')
                for col in range(1, max_col + 1)
            )

        removed = 0
        while True:
            breaks = sorted({int(brk.id) for brk in ws.row_breaks.brk})
            page_starts = [1] + [value + 1 for value in breaks]
            page_ends = breaks + [ws.max_row]
            target = None
            for start, end in zip(page_starts, page_ends):
                texts = [row_text(row) for row in range(start, end + 1)]
                has_ndt_title = any(
                    str(ws.cell(row=row, column=6).value or '').strip().startswith('3.0')
                    for row in range(start, end + 1)
                )
                if not has_ndt_title:
                    continue
                has_table = any(
                    'LINENO' in ''.join(text.split()).replace('.', '').upper()
                    for text in texts
                )
                has_detail = any(
                    isinstance(ws.cell(row=row, column=3).value, (int, float))
                    for row in range(start, end + 1)
                )
                if not has_table and not has_detail:
                    target = (start, end)
                    break
            if target is None:
                break

            start, end = target
            self._delete_rows_safely(ws, start, end - start + 1)
            unique = sorted({int(brk.id) for brk in ws.row_breaks.brk})
            ws.row_breaks.brk = [Break(id=value) for value in unique]
            removed += 1
        return removed

    def _pull_complete_ndt_tail_to_previous_page(
        self, ws, max_rows=None, bottom_reserve=0.0,
    ):
        """Pull leading continuation rows into prior-page free space.

        Records move one by one in source order.  TOTAL moves only when it also
        fits; otherwise the continuation header and TOTAL remain on the next
        page.  This keeps the 3.0 section flowing continuously without reducing
        row heights or forcing an automatic Excel overflow page.
        """
        import copy
        from openpyxl.worksheet.pagebreak import Break

        max_col = min(ws.max_column, 23)
        default_height = float(ws.sheet_format.defaultRowHeight or 15)

        def text(row):
            return ''.join(
                str(ws.cell(row=row, column=col).value or '')
                for col in range(1, max_col + 1)
            )

        def height(row):
            dim = ws.row_dimensions[row]
            return 0.0 if dim.hidden else float(dim.height or default_height)

        def snapshot(first, last):
            cells = []
            for row in range(first, last + 1):
                row_cells = []
                for col in range(1, ws.max_column + 1):
                    cell = ws.cell(row, col)
                    row_cells.append((
                        cell.value, copy.copy(cell._style),
                        copy.copy(cell.alignment), cell.number_format,
                    ))
                cells.append(row_cells)
            merges = []
            for merged in list(ws.merged_cells.ranges):
                if merged.min_row >= first and merged.max_row <= last:
                    merges.append((
                        merged.min_row - first, merged.min_col,
                        merged.max_row - first, merged.max_col,
                    ))
            return cells, [ws.row_dimensions[r].height
                           for r in range(first, last + 1)], merges

        def restore(first, cells, heights, merges):
            for offset, row_cells in enumerate(cells):
                target = first + offset
                ws.row_dimensions[target].height = heights[offset]
                ws.row_dimensions[target].hidden = False
                for col, saved in enumerate(row_cells, start=1):
                    cell = ws.cell(target, col)
                    cell.value = saved[0]
                    cell._style = copy.copy(saved[1])
                    cell.alignment = copy.copy(saved[2])
                    cell.number_format = saved[3]
            for min_ro, min_col, max_ro, max_col_idx in merges:
                ws.merge_cells(
                    start_row=first + min_ro, start_column=min_col,
                    end_row=first + max_ro, end_column=max_col_idx,
                )

        def shift_images(first_row, amount):
            for image in ws._images:
                anchor = getattr(image, 'anchor', None)
                if not hasattr(anchor, '_from') or anchor._from.row + 1 < first_row:
                    continue
                anchor._from.row += amount
                if hasattr(anchor, 'to'):
                    anchor.to.row += amount

        moved = 0
        breaks = sorted({int(brk.id) for brk in ws.row_breaks.brk})
        reference_used_height = None
        if float(bottom_reserve) == 45.0:
            reference_starts = [1] + [value + 1 for value in breaks]
            reference_ends = breaks + [ws.max_row]
            for reference_start, reference_end in zip(
                reference_starts, reference_ends
            ):
                has_first_ndt_header = any(
                    str(ws.cell(row=row, column=6).value or '').strip().startswith('3.0')
                    for row in range(reference_start, min(reference_end, reference_start + 12) + 1)
                )
                if has_first_ndt_header:
                    reference_used_height = sum(
                        height(row)
                        for row in range(reference_start, reference_end + 1)
                    )
                    break
        for index in range(1, len(breaks)):
            previous_end = breaks[index - 1]
            next_end = breaks[index]
            previous_start = breaks[index - 2] + 1 if index > 1 else 1
            next_start = previous_end + 1

            next_header = next((
                row for row in range(next_start, min(next_end, next_start + 20) + 1)
                if 'LINENO' in ''.join(text(row).split()).replace('.', '').upper()
            ), None)
            if next_header is None:
                continue
            data_start = next_header + 2
            total_row = next((
                row for row in range(data_start, next_end + 1)
                if str(ws.cell(row, 2).value or '').strip().upper() == 'TOTAL'
            ), None)
            sequence_value = ws.cell(data_start, 3).value
            total_only = total_row == data_start
            if not total_only and not (
                isinstance(sequence_value, (int, float))
                or str(sequence_value or '').strip().isdigit()
            ):
                continue
            segment_end = total_row
            if segment_end is None:
                # The same long table can continue for several pages before
                # its TOTAL.  Use the consecutive detail records available on
                # this page instead of excluding the page entirely.
                segment_end = data_start - 1
                for row in range(data_start, next_end + 1):
                    value = ws.cell(row, 3).value
                    if not (
                        isinstance(value, (int, float))
                        or str(value or '').strip().isdigit()
                    ):
                        break
                    segment_end = row
                if segment_end < data_start:
                    continue

            previous_headers = [
                row for row in range(previous_start, previous_end + 1)
                if 'LINENO' in ''.join(text(row).split()).replace('.', '').upper()
            ]
            if not previous_headers:
                continue
            previous_header = previous_headers[-1]
            previous_signature = tuple(
                str(ws.cell(previous_header, col).value or '').strip().upper()
                for col in range(1, max_col + 1)
            )
            next_signature = tuple(
                str(ws.cell(next_header, col).value or '').strip().upper()
                for col in range(1, max_col + 1)
            )
            if previous_signature != next_signature:
                continue

            margins = ws.page_margins
            if reference_used_height is not None:
                # Page 10 is the reviewed 45px baseline. Matching its actual
                # worksheet-row height is more reliable than converting PDF
                # preview pixels into points at an unknown viewer zoom.
                capacity = reference_used_height
            else:
                target_bottom_inches = max(
                    float(margins.bottom or 0),
                    float(bottom_reserve) / 72.0,
                )
                capacity = (
                    11.69 - float(margins.top or 0) - target_bottom_inches
                ) * 72.0 / (self._report_height_scale(ws) / 100.0)
            free_height = capacity - sum(
                height(row) for row in range(previous_start, previous_end + 1)
            )
            move_end = data_start - 1
            used_height = 0.0
            for row in range(data_start, segment_end + 1):
                if max_rows is not None and row - data_start >= int(max_rows):
                    break
                row_height = height(row)
                if used_height + row_height > free_height + 0.5:
                    break
                used_height += row_height
                move_end = row
            if move_end < data_start:
                continue

            cells, heights, merges = snapshot(data_start, move_end)
            count = move_end - data_start + 1
            moved_complete_tail = (
                total_row is not None and move_end == total_row
            )
            caption_row = next_header - 1
            self._delete_rows_safely(ws, data_start, count)
            self._insert_rows_safely(ws, next_start, count)
            shift_images(next_start, count)
            restore(next_start, cells, heights, merges)

            # The inserted tail belongs to the previous page; the repeated
            # document header now starts immediately after the shifted break.
            old_previous_end = previous_end
            for brk in ws.row_breaks.brk:
                if int(brk.id) == old_previous_end:
                    brk.id = old_previous_end + count
                    break

            if moved_complete_tail:
                # Remove only the now-obsolete continuation caption and its
                # two column-header rows. Keep the document header and parent.
                obsolete_caption = caption_row + count
                self._delete_rows_safely(ws, obsolete_caption, 3)
            ws.row_breaks.brk = [
                Break(id=value)
                for value in sorted({int(brk.id) for brk in ws.row_breaks.brk})
            ]
            moved += count
            break
        return moved

    @staticmethod
    def _place_general_equipment_table_image(ws):
        """Anchor the PAUT equipment table in its reserved blank rows."""
        max_col = min(ws.max_column, 23)

        heading_row = next((
            row for row in range(290, min(ws.max_row, 306) + 1)
            if 'PAUT' in ''.join(
                str(ws.cell(row=row, column=col).value or '')
                for col in range(1, max_col + 1)
            ).upper()
        ), 299)
        target_row = heading_row + 1
        for image in ws._images:
            anchor = getattr(image, 'anchor', None)
            if not hasattr(anchor, '_from'):
                continue
            image_row = anchor._from.row + 1
            if not (290 <= image_row <= 305 and float(image.width) > 500):
                continue
            row_shift = target_row - image_row
            anchor._from.row += row_shift
            if hasattr(anchor, 'to'):
                anchor.to.row += row_shift
            return target_row
        return None

    def _repair_merged_right_borders(self, ws):
        """Restore outer edges that openpyxl can drop from merged cells.

        Excel stores a merged range's visible outline across both the anchor cell
        and the cells on the perimeter.  After a load/save cycle, the right edge
        can disappear when the template only carries the outline on the anchor.
        Reapply that edge to the full right perimeter before saving.
        """
        import copy

        for merged in list(ws.merged_cells.ranges):
            anchor = ws.cell(row=merged.min_row, column=merged.min_col)

            def has_style(side):
                return bool(getattr(side, 'style', None))

            left_side = anchor.border.left
            if has_style(left_side):
                left_targets = [anchor]
                left_targets.extend(
                    ws.cell(row=row, column=merged.min_col)
                    for row in range(merged.min_row, merged.max_row + 1)
                )
                for cell in left_targets:
                    border = copy.copy(cell.border)
                    border.left = copy.copy(left_side)
                    cell.border = border

            # Prefer an existing right edge.  For framed merged cells whose right
            # edge was already lost, the matching left edge is the intended style.
            right_side = anchor.border.right
            if not has_style(right_side):
                for row in range(merged.min_row, merged.max_row + 1):
                    candidate = ws.cell(row=row, column=merged.max_col).border.right
                    if has_style(candidate):
                        right_side = candidate
                        break

            is_framed = bool(
                has_style(anchor.border.left)
                and (has_style(anchor.border.top) or has_style(anchor.border.bottom))
            )
            if not has_style(right_side) and is_framed:
                right_side = anchor.border.left

            if not has_style(right_side):
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
                                   table_header_row, split_heading=False):
        """Start a continuation page with the document and table headers.

        ``table_header_row`` is the first of the table's two column-header
        rows; the subsection title is the row immediately above it.  The
        resulting continuation block is eleven rows high, or twelve when
        repeating the SECTION parent caption:

        * seven rows of the repeated document header;
        * one blank spacer row below the document header;
        * the SECTION parent caption for a continued 1.2.x table;
        * one subsection-title row;
        * two table-column-header rows.

        Existing content at and below ``insert_row`` is moved down intact.
        """
        import copy

        title_row = table_header_row - 1
        subsection = ''.join(str(ws.cell(title_row, col).value or '')
                             for col in range(1, ws.max_column + 1))
        parent_row = None
        parent_token = None
        if re.match(r'^\s*1\.2\.\d', subsection):
            parent_token = '1.2SECTION'
        elif re.match(r'^\s*2\.[1-9]', subsection):
            parent_token = '2.0비파괴검사세부현황'
        if parent_token:
            for row in range(title_row - 1, 0, -1):
                text = ''.join(str(ws.cell(row, col).value or '')
                               for col in range(1, min(ws.max_column, 23) + 1))
                if ''.join(text.split()).upper().startswith(parent_token):
                    parent_row = row
                    break
            # Some templates keep the parent title below the first continued
            # table (for example above MT). Reuse it even in that layout.
            if parent_row is None:
                for row in range(1, ws.max_row + 1):
                    text = ''.join(str(ws.cell(row, col).value or '')
                                   for col in range(1, min(ws.max_column, 23) + 1))
                    if ''.join(text.split()).upper().startswith(parent_token):
                        parent_row = row
                        break
        source_rows = ([parent_row] if parent_row is not None else []) + list(
            range(title_row, table_header_row + 2)
        )
        subsection_offset = 1 if parent_row is not None else 0
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
        row_offsets = {row: offset for offset, row in enumerate(source_rows)}
        for merged in list(ws.merged_cells.ranges):
            if (merged.min_row in row_offsets and merged.max_row in row_offsets
                    and all(row in row_offsets for row in range(
                        merged.min_row, merged.max_row + 1))):
                source_merges.append((
                    row_offsets[merged.min_row], merged.min_col,
                    row_offsets[merged.max_row], merged.max_col,
                ))

        if split_heading:
            # A height boundary may fall between the two heading rows. Keep
            # the upper fragment on this page; the complete heading snapshot
            # above is repeated on the next page. Never extend a merge across
            # the newly inserted document header.
            for merged in list(ws.merged_cells.ranges):
                if merged.min_row < insert_row <= merged.max_row:
                    top, left, right = merged.min_row, merged.min_col, merged.max_col
                    ws.unmerge_cells(str(merged))
                    ws.merge_cells(start_row=top, start_column=left,
                                   end_row=insert_row - 1, end_column=right)

        self._insert_continuation_header(
            ws, insert_row, document_header_row, row_count=7
        )
        # Leave the first row below the common report header empty.  The
        # subsection title/table starts on the second row below the header.
        spacer_row = insert_row + 7
        table_insert_row = spacer_row + 1
        inserted_rows = len(source_rows) + 1
        self._insert_rows_safely(ws, spacer_row, inserted_rows)

        # Keep images below the repeated document header aligned with the rows
        # moved by the spacer plus three-row subsection/table-header insertion.
        for image in list(ws._images):
            anchor = getattr(image, 'anchor', None)
            if not hasattr(anchor, '_from') or anchor._from.row + 1 < table_insert_row:
                continue
            anchor._from.row += inserted_rows
            if hasattr(anchor, 'to'):
                anchor.to.row += inserted_rows

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
            cell = ws.cell(row=table_insert_row + subsection_offset, column=col)
            if str(cell.value or '').strip():
                base_title = re.sub(
                    r'(?:\s*\(계속\))+$', '', str(cell.value)
                ).rstrip()
                cell.value = f"{base_title} (계속)"
                break

        return 7 + inserted_rows

    def _insert_page13_rt_continuation(self, ws, insert_row,
                                       document_header_row, rt_header_row):
        """Create a page-13-style continuation page for overflowing RT rows.

        Rows at ``insert_row`` and below already contain the overflowing RT
        records, the RT TOTAL row, and the unused lower-page frame. Inserting
        the copied document/RT headers at that exact point completes the page
        13 frame without duplicating PAUT, MT, or PT data. The safe insertion
        also moves the original page 14 and every later page break down.
        """
        return self._insert_table_continuation(
            ws,
            insert_row,
            document_header_row,
            rt_header_row,
        )

    def _move_table_to_continuation(self, ws, insert_row, document_header_row,
                                    table_header_row):
        """Repeat the parent above a whole table moved to the next page."""
        inserted = self._insert_table_continuation(
            ws, insert_row, document_header_row, table_header_row)
        if inserted == 11:
            self._delete_rows_safely(ws, insert_row + 7, 4)
            return 7
        # The original subsection title and two header rows move with the
        # table. Remove their copied equivalents, retaining the parent title.
        self._delete_rows_safely(ws, insert_row + inserted - 3, 3)
        if insert_row == table_header_row - 1 and insert_row > 1:
            previous = insert_row - 1
            text = ''.join(str(cell.value or '') for cell in ws[previous])
            compact = ''.join(text.split()).upper()
            if compact.startswith(('2.0비파괴검사세부현황', '1.2SECTION')):
                for cell in ws[previous]:
                    if cell.value is not None:
                        cell.value = None
                ws.row_dimensions[previous].height = 2.0
        return inserted - 3

    def _insert_status_continuation(self, ws, insert_row, document_header_row,
                                    title_row, header_row):
        """14페이지 용접사 현황의 공통 헤더와 전체 표 헤더를 반복한다."""
        import copy

        max_col = ws.max_column
        source_rows = list(range(title_row, header_row + 1))
        snapshots = []
        for source_row in source_rows:
            snapshots.append([
                (cell.value, copy.copy(cell._style), cell.number_format)
                for cell in ws[source_row]
            ])
        heights = [ws.row_dimensions[row].height for row in source_rows]
        merges = [
            (merged.min_row - title_row, merged.min_col,
             merged.max_row - title_row, merged.max_col)
            for merged in list(ws.merged_cells.ranges)
            if merged.min_row >= title_row and merged.max_row <= header_row
        ]

        self._insert_continuation_header(
            ws, insert_row, document_header_row, row_count=7
        )
        block_start = insert_row + 8  # 공통 헤더 아래 한 행은 공백
        block_rows = len(source_rows)
        self._insert_rows_safely(ws, insert_row + 7, block_rows + 1)

        # 두 번째 삽입 아래의 그림도 행과 함께 이동시킨다.
        for image in list(ws._images):
            anchor = getattr(image, 'anchor', None)
            if not hasattr(anchor, '_from') or anchor._from.row + 1 < insert_row + 7:
                continue
            anchor._from.row += block_rows + 1
            if hasattr(anchor, 'to'):
                anchor.to.row += block_rows + 1

        for offset, snapshot in enumerate(snapshots):
            target_row = block_start + offset
            ws.row_dimensions[target_row].height = heights[offset]
            for col, (value, style, number_format) in enumerate(snapshot, 1):
                cell = ws.cell(target_row, col)
                cell.value = value
                cell._style = copy.copy(style)
                cell.number_format = number_format
        for min_ro, min_col, max_ro, max_col_idx in merges:
            ws.merge_cells(
                start_row=block_start + min_ro, start_column=min_col,
                end_row=block_start + max_ro, end_column=max_col_idx,
            )

        for col in range(1, max_col + 1):
            cell = ws.cell(block_start, col)
            if str(cell.value or '').strip():
                cell.value = f"{cell.value} (계속)"
                break
        return 8 + block_rows

    def _consume_detail_page_whitespace(self, ws, total_row, added_height):
        """Pay for added detail rows from trailing whitespace on this page."""
        from openpyxl.worksheet.pagebreak import Break

        later_breaks = sorted(int(b.id) for b in ws.row_breaks.brk
                              if int(b.id) > total_row)
        if not later_breaks:
            return 0.0
        page_end = later_breaks[0]
        remaining = float(added_height)
        default_height = float(ws.sheet_format.defaultRowHeight or 15)
        while page_end > total_row and remaining > 0.001:
            if any(ws.cell(page_end, col).value is not None
                   and str(ws.cell(page_end, col).value).strip()
                   for col in range(1, ws.max_column + 1)):
                break
            if any(m.min_row < page_end <= m.max_row
                   for m in ws.merged_cells.ranges):
                break
            occupied = False
            for image in ws._images:
                anchor = getattr(image, 'anchor', None)
                if not hasattr(anchor, '_from'):
                    continue
                first = anchor._from.row + 1
                last = (anchor.to.row + 1 if hasattr(anchor, 'to') else
                        first + max(1, int(float(image.height) / 10)))
                if first <= page_end <= last:
                    occupied = True
                    break
            if occupied:
                break
            dimension = ws.row_dimensions[page_end]
            height = (0.0 if dimension.hidden else
                      float(dimension.height if dimension.height is not None
                            else default_height))
            if height > remaining + 0.001:
                dimension.height = height - remaining
                remaining = 0.0
                break
            deleted_row = page_end
            self._delete_rows_safely(ws, deleted_row, 1)
            for key, marker in self.table_markers.items():
                if marker > deleted_row:
                    self.table_markers[key] -= 1
            for image in ws._images:
                anchor = getattr(image, 'anchor', None)
                if hasattr(anchor, '_from') and anchor._from.row + 1 > deleted_row:
                    anchor._from.row -= 1
                    if hasattr(anchor, 'to'):
                        anchor.to.row -= 1
            page_end -= 1
            remaining = max(0.0, remaining - height)
        # Deleting the old last spacer removes its manual break as well.
        if not any(int(b.id) == page_end for b in ws.row_breaks.brk):
            ws.row_breaks.append(Break(id=page_end))
        ws.row_breaks.brk.sort(key=lambda b: int(b.id))
        return float(added_height) - remaining

    def _paut_print_scale(self, ws):
        """Estimate the width-fit scale used by the A4 portrait report."""
        from openpyxl.utils import get_column_letter
        from openpyxl.worksheet.cell_range import CellRange

        if ws.page_setup.scale:
            return float(ws.page_setup.scale)
        if not ws.page_setup.fitToWidth or not ws.print_area:
            return 100.0
        area = CellRange(str(ws.print_area).split('!')[-1].replace('$', ''))
        width_points = 0.0
        for col in range(area.min_col, area.max_col + 1):
            dimension = ws.column_dimensions[get_column_letter(col)]
            if dimension.hidden:
                continue
            width = dimension.width or ws.sheet_format.defaultColWidth or 8.43
            width_points += (float(width) * 7 + 5) * 0.75
        usable_width = (8.27 - float(ws.page_margins.left or 0)
                        - float(ws.page_margins.right or 0)) * 72
        return min(100.0, usable_width / max(width_points, 1) * 100)

    def _report_height_scale(self, ws):
        """Return the calibrated Excel width-fit scale for vertical paging.

        Excel's rendered width-fit pagination is about nine percent less than
        the scale estimated from OOXML column widths.  A five-percent
        correction split the first 3.0 page after only two RT rows even though
        the exported PDF still had room above the bottom margin.  This general
        calibration keeps the taller detail pages within Excel's print area;
        the shorter first 3.0 summary page has its own tighter calibration.
        """
        return self._paut_print_scale(ws) * 0.91

    def _ndt_reflow_height_scale(self, ws, page_number=1):
        """Return the 3.0 reflow scale for a summary or detail page.

        Use the same calibrated boundary on the first two 3.0 pages.  Giving
        the first page a looser boundary let its final record and TOTAL extend
        visibly lower than the continuation page. A shared scale keeps both
        pages within the same printable baseline, allowing TOTAL to split.
        """
        if page_number in (1, 2):
            # Excel normalizes saved column widths slightly, raising the
            # effective scale from about 88.14% to 88.47%. This small second-
            # page correction also provides a consistent first-page boundary.
            return self._paut_print_scale(ws) * 0.905
        return self._report_height_scale(ws)

    def _lower_ndt_page_titles(self, ws, extra_height=5.0):
        """Lower only 3.0 page content by enlarging its header spacer.

        The seven-row document header is copied to every continuation page.
        Increasing its final blank row before pagination moves each 3.0 title
        and table down equally, without changing pages in sections 1, 2, or 4.
        """
        max_column = min(ws.max_column, 23)

        def row_text(row):
            return ''.join(
                str(ws.cell(row=row, column=column).value or '')
                for column in range(1, max_column + 1)
            )

        for title_row in range(1, ws.max_row + 1):
            title = ''.join(
                str(ws.cell(row=title_row, column=6).value or '').split()
            )
            if '3.0비파괴검사현황' not in title:
                continue

            for header_start in range(max(1, title_row - 6), title_row + 1):
                if '월간용역진도보고서' not in ''.join(
                    row_text(header_start).split()
                ):
                    continue
                spacer_row = header_start + 6
                current_height = float(
                    ws.row_dimensions[spacer_row].height
                    or ws.sheet_format.defaultRowHeight
                    or 15
                )
                ws.row_dimensions[spacer_row].height = (
                    current_height + float(extra_height)
                )
                return spacer_row
        return None

    def _compact_ndt_summary_page(self, ws):
        """Reclaim blank summary spacing before creating continuation pages."""
        first = self.table_markers['qty']
        last_header = self.table_markers['pt_1']
        detail_header = self.table_markers['paut_2']
        end = next((r for r in range(last_header + 2, detail_header)
                    if 'TOTAL' in str(ws.cell(r, 2).value or '').upper()), None)
        if end is None:
            return
        boundaries = [int(b.id) for b in ws.row_breaks.brk if int(b.id) < first]
        start = max(boundaries, default=0) + 1
        default = float(ws.sheet_format.defaultRowHeight or 15)
        setup = ws.page_setup
        old_scale, old_width = setup.scale, setup.fitToWidth
        try:
            setup.scale, setup.fitToWidth = None, 1
            scale = self._paut_print_scale(ws) / 100
        finally:
            setup.scale, setup.fitToWidth = old_scale, old_width
        margins = ws.page_margins
        capacity = (11.69 - float(margins.top or 0)
                    - float(margins.bottom or 0)
                    - float(margins.footer or 0)) * 72 / scale * 0.97
        def height(row):
            dim = ws.row_dimensions[row]
            return 0 if dim.hidden else float(dim.height if dim.height is not None else default)
        excess = sum(height(r) for r in range(start, end + 1)) - capacity
        if excess <= 0:
            return
        # Only the blank gaps after summary TOTAL rows are adjustable. Preserve
        # document titles, table headers, records, merged cells, and drawings.
        merged_rows = {r for m in ws.merged_cells.ranges
                       for r in range(m.min_row, m.max_row + 1)}
        image_rows = {im.anchor._from.row + 1 for im in ws._images
                      if hasattr(im.anchor, '_from')}
        candidates = []
        markers = [self.table_markers[k] for k in
                   ('paut_1', 'mt_1', 'rt_1', 'pt_1')]
        for left, right in zip(markers, markers[1:]):
            total = next((r for r in range(left + 2, right)
                          if 'TOTAL' in str(ws.cell(r, 2).value or '').upper()), None)
            if total is None:
                continue
            for r in range(total + 1, right):
                if r in merged_rows or r in image_rows:
                    continue
                if not any(ws.cell(r, c).value is not None
                           for c in range(1, min(ws.max_column, 23) + 1)):
                    candidates.append(r)
        reclaimable = sum(max(0, height(r) - 2) for r in candidates)
        if reclaimable:
            ratio = min(1, excess / reclaimable)
            for r in candidates:
                current = height(r)
                if current > 2:
                    ws.row_dimensions[r].height = current - (current - 2) * ratio

    def _enforce_final_ndt_body_limit(self, ws):
        """Recheck completed NDT blocks against the actual printable height.

        Row counts are not a reliable page limit because summary/detail rows can
        have very different heights. Using a fixed 58-row boundary left a large
        blank area on mixed summary/PAUT pages even though more records fit.
        """
        from openpyxl.worksheet.pagebreak import Break

        def text(row):
            return ''.join(str(ws.cell(row, c).value or '')
                           for c in range(1, min(ws.max_column, 23) + 1)).replace(' ', '')

        start = 1
        while start <= ws.max_row:
            ends = sorted(int(b.id) for b in ws.row_breaks.brk if int(b.id) >= start)
            end = ends[0] if ends else ws.max_row
            is_ndt = any('3.0비파괴검사현황' in text(r)
                         for r in range(start, min(end, start + 6) + 1))
            if not is_ndt:
                start = end + 1
                continue
            split = self._first_print_overflow_row(
                ws, start, end - start + 1,
                self._report_height_scale(ws),
            )
            if split is None:
                start = end + 1
                continue
            # Locate the actual two-row table heading above the overflow.
            all_headers = [
                r for r in range(start + 7, end)
                if 'LINENO' in ''.join(text(r).split()).replace('.', '').upper()
            ]
            if not all_headers:
                start = end + 1
                continue
            preceding_headers = [r for r in all_headers if r < split]
            # When the height boundary falls in the gap before a table, move
            # that whole table instead of ignoring the overflow.
            header = preceding_headers[-1] if preceding_headers else all_headers[0]
            # A completed/empty table followed by template whitespace must not
            # generate a continuation containing only its copied heading.
            next_headers = [
                r for r in range(header + 2, end)
                if 'LINENO' in ''.join(text(r).split()).replace('.', '').upper()
            ]
            table_end = next_headers[0] - 2 if next_headers else end
            total = next((r for r in range(header + 2, table_end + 1)
                          if 'TOTAL' in text(r).upper()), table_end + 1)
            # When the printable boundary falls after a completed table but
            # before the following one, move the following table as a whole.
            # Otherwise the completed table is selected and the overflow is
            # incorrectly ignored, leaving a large blank area on this page.
            if split > total and next_headers:
                header = next_headers[0]
                later_headers = [
                    r for r in range(header + 2, end)
                    if 'LINENO' in ''.join(text(r).split()).replace('.', '').upper()
                ]
                table_end = later_headers[0] - 2 if later_headers else end
                total = next((r for r in range(header + 2, table_end + 1)
                              if 'TOTAL' in text(r).upper()), table_end + 1)
            def is_record(row):
                return any(str(ws.cell(row, c).value or '').strip().isdigit()
                           for c in (2, 3))
            if split != total and not any(
                is_record(r) for r in range(max(split, header + 2), total)
            ):
                start = end + 1
                continue
            # Do not separate the two-row heading from its first record.
            move_whole_table = split <= header + 2
            if split <= header + 2:
                split = header - 1
            if split <= start + 7:
                raise ValueError('검사현황 표 머리글이 인쇄 가능 높이를 초과합니다.')
            if move_whole_table:
                # The original caption and column heading move with the data.
                # Copying them too would leave an orphan duplicate above them.
                inserted = self._move_table_to_continuation(ws, split, start, header)
            else:
                inserted = self._insert_table_continuation(ws, split, start, header)
            ws.row_breaks.append(Break(id=split - 1))
            ws.row_breaks.brk.sort(key=lambda b: int(b.id))
            for key, row in self.table_markers.items():
                if row >= split:
                    self.table_markers[key] += inserted
            start = split

    def _stale_ndt_continuation_blocks(self, ws, start_row, end_row):
        """Return copied table-heading blocks that must be rebuilt.

        Template versions are inconsistent: some fixed continuation pages add
        ``(계속)`` to the subsection title and others repeat the original title
        unchanged.  Delete a heading only when the same normalized subsection
        was already seen.  If the first surviving heading temporarily carries
        ``(계속)``, keep it and remove only that suffix.
        """
        stale_blocks = []
        first_continuation_titles = []
        seen_titles = set()
        max_column = min(ws.max_column, 23)

        for row in range(start_row, end_row):
            title = ''.join(
                str(ws.cell(row=row, column=col).value or '')
                for col in range(1, max_column + 1)
            ).strip()
            if not title:
                continue

            # A subsection caption is immediately followed by the first
            # column-heading row. Looking two rows ahead incorrectly treats a
            # repeated parent caption such as ``2.0 비파괴검사세부현황`` as the
            # table title and deletes the real MT caption plus both headings.
            table_header_row = row + 1
            if table_header_row >= end_row:
                continue
            header_text = ''.join(
                str(ws.cell(row=table_header_row, column=col).value or '')
                for col in range(1, max_column + 1)
            )
            if 'LINENO' not in ''.join(header_text.split()).replace('.', '').upper():
                continue

            normalized = re.sub(
                r'(?:\s*\(계속(?:\s+\d+)?\))+$', '', title
            ).strip()
            normalized_key = ''.join(normalized.split()).upper()
            is_explicit_continuation = '(계속' in title
            if normalized_key in seen_titles:
                stale_blocks.append((row, table_header_row - row + 2))
            else:
                seen_titles.add(normalized_key)
                if is_explicit_continuation:
                    for col in range(1, max_column + 1):
                        cell = ws.cell(row=row, column=col)
                        value = str(cell.value or '').strip()
                        if not value:
                            continue
                        cleaned = re.sub(
                            r'(?:\s*\(계속(?:\s+\d+)?\))+$', '', value
                        ).rstrip()
                        first_continuation_titles.append((row, col, cleaned))
                        break

        return stale_blocks, first_continuation_titles

    def _compact_ndt_summary_title_rows(self, ws, start_row, end_row):
        """Reclaim first-page height from summary captions, not data rows."""
        compact_height = 13.75
        max_column = min(ws.max_column, 23)
        for row in range(start_row, end_row):
            text = ''.join(
                str(ws.cell(row=row, column=col).value or '')
                for col in range(1, max_column + 1)
            )
            compact = ''.join(text.split())
            if not re.match(r'^1\.(?:0|1|2)', compact):
                continue
            current = float(
                ws.row_dimensions[row].height
                or ws.sheet_format.defaultRowHeight
                or 15
            )
            # Only compact the 16.5-point caption rows. Table headings and
            # records use different heights and must remain untouched.
            if compact_height < current <= 17.0:
                ws.row_dimensions[row].height = compact_height

    @staticmethod
    def _remove_orphan_ndt_subsection_captions(ws, start_row, end_row):
        """Remove old stranded captions only when a complete copy also exists."""
        max_col = min(ws.max_column, 23)
        captions = []
        complete = set()
        for row in range(start_row, end_row):
            title = ''.join(str(ws.cell(row, col).value or '')
                            for col in range(1, max_col + 1)).strip()
            if not re.match(r'^(?:1\.2\.\d+|2\.[1-9]\d*)\s', title):
                continue
            normalized = re.sub(r'(?:\s*\(계속(?:\s+\d+)?\))+$', '', title).strip()
            key = ''.join(normalized.split()).upper()
            following = ''.join(str(ws.cell(row + 1, col).value or '')
                                for col in range(1, max_col + 1))
            has_heading = 'LINENO' in ''.join(following.split()).replace('.', '').upper()
            captions.append((row, key, has_heading))
            if has_heading:
                complete.add(key)
        for row, key, has_heading in captions:
            if has_heading or key not in complete:
                continue
            for cell in ws[row]:
                if isinstance(cell, Cell):
                    cell.value = None
            ws.row_dimensions[row].height = 0
            ws.row_dimensions[row].hidden = True

    def _reflow_ndt_section(self, ws):
        """Flow all 3.0 tables through available space before section 4.0.

        The template owns several fixed 3.0 pages. After dynamic rows are
        written those copied headers and manual breaks can strand PAUT/MT on
        one page and RT/PT on the next even though the first page has room.
        Remove only the repeated 3.0 document headers, keep the first header,
        clear the internal template breaks, and retain a hard boundary before
        4.0. ``_enforce_final_ndt_body_limit`` then repaginates the continuous
        3.0 body by actual printable height and inserts continuation headers
        only where data really overflows.
        """
        from openpyxl.worksheet.pagebreak import Break

        def row_text(row):
            return ''.join(
                str(ws.cell(row=row, column=col).value or '')
                for col in range(1, min(ws.max_column, 23) + 1)
            )

        def section_header_starts(section_token):
            starts = []
            for row in range(1, ws.max_row + 1):
                # Real report section captions are anchored in column F. The
                # table of contents repeats the same words in another column
                # and must not be mistaken for a printable section header.
                compact = ''.join(
                    str(ws.cell(row=row, column=6).value or '').split()
                )
                if section_token not in compact:
                    continue
                for candidate in range(max(1, row - 6), row + 1):
                    candidate_text = ''.join(row_text(candidate).split())
                    if '월간용역진도보고서' in candidate_text:
                        starts.append(candidate)
                        break
            return sorted(set(starts))

        ndt_headers = section_header_starts('3.0비파괴검사현황')
        if not ndt_headers:
            return

        # Delete bottom-up so every stored row remains valid. Each common
        # report header is exactly seven rows; subsection/table headings below
        # it remain and join the continuous 3.0 body.
        for header_start in reversed(ndt_headers[1:]):
            self._delete_rows_safely(ws, header_start, 7)

        # Remove continuation captions/table headers inherited from the fixed
        # template pages. The continuous body must contain records only;
        # pagination below will insert a fresh continuation block exactly at
        # each real page boundary. Keeping the old blocks creates empty or
        # duplicated headings in the middle of a printed page.
        current_defects = section_header_starts('4.0용접결함,불량')
        stale_scan_end = current_defects[0] if current_defects else ws.max_row + 1
        stale_continuations, first_continuation_titles = (
            self._stale_ndt_continuation_blocks(
            ws, ndt_headers[0], stale_scan_end
            )
        )
        # A table can be moved as a whole before reflow, making its only
        # surviving heading look like a continuation. Preserve that heading
        # and normalize its caption before row deletions shift its position.
        for row, col, cleaned_title in first_continuation_titles:
            ws.cell(row=row, column=col).value = cleaned_title
        for start_row, row_count in reversed(stale_continuations):
            self._delete_rows_safely(ws, start_row, row_count)

        first_ndt = section_header_starts('3.0비파괴검사현황')[0]
        # Templates in circulation use both "불량율" and "불량률".
        defect_headers = section_header_starts('4.0용접결함,불량')
        if not defect_headers:
            return
        first_defect = defect_headers[0]
        self._remove_orphan_ndt_subsection_captions(ws, first_ndt, first_defect)
        self._remove_duplicate_ndt_detail_captions(ws, first_ndt, first_defect)

        # The summary captions have generous vertical padding. Reducing only
        # those caption rows lets the RT subcontract subtotal and TOTAL stay on
        # page 10 without shrinking records or crossing the bottom margin.
        self._compact_ndt_summary_title_rows(
            ws, first_ndt + 7, first_defect
        )

        # Template spacer rows used to pad fixed pages still consume printable
        # height after their repeated headers are removed. Collapse only truly
        # empty, non-merged, image-free rows inside 3.0 so the next table can
        # use every remaining part of the current page.
        merged_rows = {
            row
            for merged in ws.merged_cells.ranges
            for row in range(merged.min_row, merged.max_row + 1)
        }
        image_rows = {
            image.anchor._from.row + 1
            for image in ws._images
            if hasattr(getattr(image, 'anchor', None), '_from')
        }
        for row in range(first_ndt + 7, first_defect):
            if row in merged_rows or row in image_rows:
                continue
            if any(
                str(ws.cell(row=row, column=col).value or '').strip()
                for col in range(1, min(ws.max_column, 23) + 1)
            ):
                continue
            ws.row_dimensions[row].hidden = False
            ws.row_dimensions[row].height = 2.0

        retained = [
            brk for brk in ws.row_breaks.brk
            if int(brk.id) < first_ndt or int(brk.id) >= first_defect - 1
        ]
        if not any(int(brk.id) == first_defect - 1 for brk in retained):
            retained.append(Break(id=first_defect - 1))
        ws.row_breaks.brk = sorted(retained, key=lambda brk: int(brk.id))

        # A copied continuation caption may already contain the suffix. The
        # repagination step is free to copy it again, so normalize it once.
        for row in range(first_ndt, first_defect):
            for col in range(1, min(ws.max_column, 23) + 1):
                cell = ws.cell(row=row, column=col)
                if isinstance(cell.value, str):
                    while '(계속) (계속)' in cell.value:
                        cell.value = cell.value.replace('(계속) (계속)', '(계속)')

    def _paginate_reflowed_ndt(self, ws):
        """Paginate the continuous 3.0 body and repeat only needed headers."""
        from openpyxl.worksheet.pagebreak import Break
        max_column = min(ws.max_column, 23)

        def text(row):
            return ''.join(
                str(ws.cell(row=row, column=col).value or '')
                for col in range(1, max_column + 1)
            )

        def title_row(token):
            for row in range(1, ws.max_row + 1):
                compact = ''.join(
                    str(ws.cell(row=row, column=6).value or '').split()
                )
                if token in compact:
                    for candidate in range(max(1, row - 6), row + 1):
                        if '월간용역진도보고서' in ''.join(
                            text(candidate).split()
                        ):
                            return candidate
            return None

        first_ndt = title_row('3.0비파괴검사현황')
        if first_ndt is None:
            return

        for _ in range(50):
            first_defect = title_row('4.0용접결함,불량')
            if first_defect is None:
                return
            internal_breaks = sorted(
                int(brk.id) for brk in ws.row_breaks.brk
                if first_ndt <= int(brk.id) < first_defect - 1
            )
            page_start = first_ndt
            page_number = 1
            changed = False
            for page_end in internal_breaks + [first_defect - 1]:
                overflow = self._first_print_overflow_row(
                    ws, page_start, page_end - page_start + 1,
                    reserve_footer=False,
                    capacity_points=self.NDT_PAGE_HEIGHT,
                )
                if overflow is None:
                    page_start = page_end + 1
                    page_number += 1
                    continue

                headers = [
                    row for row in range(page_start + 7, page_end)
                    if 'LINENO' in ''.join(text(row).split()).replace('.', '').upper()
                ]
                if not headers:
                    raise ValueError(
                        '검사현황 표 머리글을 찾을 수 없어 페이지를 나눌 수 없습니다. '
                        f'(페이지 {page_start}~{page_end}, 넘침 행 {overflow})'
                    )
                following = [row for row in headers if row > overflow]
                preceding = [row for row in headers if row <= overflow]
                header = (
                    following[0] if following and (
                        not preceding or overflow >= following[0] - 1
                    ) else preceding[-1] if preceding else headers[0]
                )

                split = overflow
                heading_group = None
                for candidate in headers:
                    caption = candidate - 1
                    group_start = caption
                    # A parent SECTION/2.0 title belongs with the table too;
                    # otherwise it can be stranded at the preceding foot.
                    for previous in range(caption - 1, max(page_start + 6, caption - 4), -1):
                        compact = ''.join(text(previous).split()).upper()
                        if not compact:
                            continue
                        if compact.startswith(('1.2SECTION', '2.0비파괴검사세부현황')):
                            group_start = previous
                        break
                    first_record = candidate + 2
                    while first_record < page_end and not text(first_record).strip():
                        first_record += 1
                    if group_start <= overflow <= first_record:
                        heading_group = (candidate, group_start, caption)
                        header, split = candidate, group_start
                        break
                if split <= page_start + 7:
                    raise ValueError(
                        '검사 표의 제목 높이가 인쇄 가능 영역을 초과합니다. '
                        '양식의 제목 행 높이와 인쇄 여백을 확인해 주세요.'
                    )

                table_complete = any(
                    str(ws.cell(row, 2).value or '').strip().upper() == 'TOTAL'
                    for row in range(header + 2, split)
                )
                if heading_group:
                    if split < heading_group[2]:
                        # The original parent and complete table heading move
                        # together. Copy only the report frame, no duplicates.
                        inserted = 7
                        self._insert_continuation_header(
                            ws, split, first_ndt, row_count=inserted)
                    else:
                        inserted = self._move_table_to_continuation(
                            ws, split, first_ndt, header)
                elif split <= header - 1 or table_complete:
                    # The complete next table is still below the boundary.
                    # Its original caption/headings move with it unchanged.
                    inserted = 7
                    self._insert_continuation_header(
                        ws, split, first_ndt, row_count=inserted)
                else:
                    inserted = self._insert_table_continuation(
                        ws, split, first_ndt, header,
                    )
                ws.row_breaks.append(Break(id=split - 1))
                ws.row_breaks.brk.sort(key=lambda brk: int(brk.id))
                for key, marker_row in self.table_markers.items():
                    if marker_row >= split:
                        self.table_markers[key] += inserted
                changed = True
                break
            if not changed:
                return
        raise ValueError('검사현황 페이지 분할이 50회를 초과했습니다.')

    def _insert_fixed_ndt_top_spacers(self, ws, spacer_height=20.0,
                                      reserved_height=0.0):
        """Insert a blank row above each 3.0 logo inside its page boundary.

        Transfer reserved height from the final header row to the new row
        so pagination retains the same total height on every 3.0 page.
        """
        max_column = min(ws.max_column, 23)

        def row_text(row):
            return ''.join(
                str(ws.cell(row=row, column=column).value or '')
                for column in range(1, max_column + 1)
            )

        header_starts = []
        for title_row in range(1, ws.max_row + 1):
            title = ''.join(
                str(ws.cell(row=title_row, column=6).value or '').split()
            )
            if '3.0비파괴검사현황' not in title:
                continue
            for candidate in range(max(1, title_row - 6), title_row + 1):
                if '월간용역진도보고서' in ''.join(
                    row_text(candidate).split()
                ):
                    header_starts.append(candidate)
                    break

        for header_start in reversed(sorted(set(header_starts))):
            spacer_row = header_start + 6
            current_height = float(
                ws.row_dimensions[spacer_row].height
                or ws.sheet_format.defaultRowHeight
                or 15
            )
            ws.row_dimensions[spacer_row].height = (
                current_height - float(reserved_height)
            )
            ws.row_dimensions[spacer_row].hidden = False
            # Leave the boundary at header_start - 1 before the blank row.
            # Later breaks, merges and row heights move with the existing rows.
            self._insert_rows_safely(ws, header_start, 1)
            ws.row_dimensions[header_start].height = float(spacer_height)
            ws.row_dimensions[header_start].hidden = False
            for image in list(ws._images):
                anchor = getattr(image, 'anchor', None)
                if hasattr(anchor, '_from') and anchor._from.row + 1 >= header_start:
                    anchor._from.row += 1
                    if hasattr(anchor, 'to'):
                        anchor.to.row += 1
            for key, marker_row in self.table_markers.items():
                if marker_row >= header_start:
                    self.table_markers[key] += 1

        return len(set(header_starts))

    def _insert_other_report_top_spacers(self, ws, spacer_height=20.0):
        """Give non-3.0 report pages the same blank row above the header."""
        starts = []
        framed_starts = {m.min_row for m in ws.merged_cells.ranges
                         if m.min_col == 6 and m.max_col == 15
                         and m.max_row >= m.min_row + 2}
        for row in range(1, ws.max_row + 1):
            if framed_starts and row not in framed_starts:
                continue
            text = ''.join(str(ws.cell(row, col).value or '')
                           for col in range(1, min(ws.max_column, 23) + 1))
            if '월간용역진도보고서' not in ''.join(text.split()):
                continue
            section = ''.join(str(ws.cell(r, 6).value or '')
                              for r in range(row, min(row + 7, ws.max_row + 1)))
            if '3.0비파괴검사현황' in ''.join(section.split()):
                continue
            starts.append(row)
        for start in reversed(starts):
            self._insert_rows_safely(ws, start, 1)
            ws.row_dimensions[start].height = float(spacer_height)
            ws.row_dimensions[start].hidden = False
            for image in ws._images:
                anchor = getattr(image, 'anchor', None)
                if hasattr(anchor, '_from') and anchor._from.row + 1 >= start:
                    anchor._from.row += 1
                    if hasattr(anchor, 'to'):
                        anchor.to.row += 1
            for key, marker in self.table_markers.items():
                if marker >= start:
                    self.table_markers[key] += 1
        return len(starts)

    def _remove_duplicate_ndt_detail_captions(self, ws, start, end):
        """Keep one 2.0 title in the continuous body before repagination."""
        seen = False
        for row in range(start, end):
            text = ''.join(str(ws.cell(row, col).value or '')
                           for col in range(1, min(ws.max_column, 23) + 1))
            if not ''.join(text.split()).startswith('2.0비파괴검사세부현황'):
                continue
            if not seen:
                seen = True
                continue
            for cell in ws[row]:
                if cell.value is not None:
                    cell.value = None
            ws.row_dimensions[row].height = 2.0

    def _space_ndt_summary_caption(self, ws, gap=6.0):
        """Move summary parent captions up without shifting their tables."""
        from copy import copy

        changed = False
        for row in range(2, ws.max_row + 1):
            text = ''.join(str(ws.cell(row, col).value or '')
                           for col in range(1, min(ws.max_column, 23) + 1))
            compact = ''.join(text.split()).upper()
            if not compact.startswith('1.0비파괴검사'):
                continue
            previous = row - 1
            if any(ws.cell(previous, col).value is not None
                   for col in range(1, min(ws.max_column, 23) + 1)):
                continue
            if any(m.min_row <= previous <= m.max_row
                   for m in ws.merged_cells.ranges):
                continue
            default = float(ws.sheet_format.defaultRowHeight or 15)
            before = ws.row_dimensions[previous]
            caption = ws.row_dimensions[row]
            if before.hidden or caption.hidden:
                continue
            height = float(before.height if before.height is not None else default)
            amount = min(float(gap), max(0.0, height - 1.0))
            before.height = height - amount
            caption.height = float(caption.height if caption.height is not None
                                   else default) + amount
            for cell in ws[row]:
                if cell.value is not None:
                    alignment = copy(cell.alignment)
                    alignment.vertical = 'top'
                    cell.alignment = alignment
            changed = changed or bool(amount)
        return changed

    def _normalize_ndt_parent_caption_style(self, ws):
        """Use the 1.0 caption's typography and left edge for SECTION titles."""
        from copy import copy

        reference = None
        targets = []
        for cells in ws:
            for cell in cells:
                text = ''.join(str(cell.value or '').split()).upper()
                if text.startswith('1.0비파괴검사'):
                    reference = cell
                elif text.startswith('1.2SECTION'):
                    targets.append(cell)
        if reference is None:
            return
        for source in targets:
            value = source.value
            merges = [m for m in ws.merged_cells.ranges
                      if m.min_row == source.row == m.max_row
                      and m.min_col <= source.column <= m.max_col]
            end_col = max([m.max_col for m in merges] + [source.column])
            for merged in merges:
                ws.unmerge_cells(str(merged))
            source.value = None
            target = ws.cell(source.row, reference.column, value)
            target.font = copy(reference.font)
            target.alignment = copy(reference.alignment)
            alignment = copy(target.alignment)
            alignment.vertical = 'top'
            target.alignment = alignment
            if merges and end_col > reference.column:
                ws.merge_cells(start_row=source.row, end_row=source.row,
                               start_column=reference.column, end_column=end_col)

    def _align_ndt_parent_top_gaps(self, ws):
        """Match continued parent captions to the first 1.0 header gap."""
        default = float(ws.sheet_format.defaultRowHeight or 15)

        def height(row):
            dim = ws.row_dimensions[row]
            return 0.0 if dim.hidden else float(
                dim.height if dim.height is not None else default)

        headers = []
        for row in range(1, ws.max_row + 1):
            text = ''.join(str(ws.cell(row, col).value or '')
                           for col in range(1, min(ws.max_column, 23) + 1))
            if '월간용역진도보고서' in ''.join(text.split()):
                headers.append(row)
        captions = []
        for start in headers:
            for row in range(start + 7, min(start + 14, ws.max_row + 1)):
                text = ''.join(str(ws.cell(row, col).value or '')
                               for col in range(1, min(ws.max_column, 23) + 1))
                compact = ''.join(text.split()).upper()
                if not compact:
                    continue
                if compact.startswith(('1.0비파괴검사', '1.2SECTION',
                                       '2.0비파괴검사세부현황')):
                    captions.append((start, row, compact))
                break
        reference = next(((start, row) for start, row, text in captions
                          if text.startswith('1.0비파괴검사')), None)
        if reference is None:
            return
        # The report-title text starts one row below the visual frame start.
        # Include its final blank frame row when comparing the visible gap;
        # continuation copies can otherwise retain an extra 6pt even after
        # the separate body spacer has been collapsed.
        desired = sum(height(row) for row in range(reference[0] + 6, reference[1]))
        for start, caption, text in captions:
            if text.startswith('1.0비파괴검사') or caption <= start + 7:
                continue
            rows = list(range(start + 6, caption))
            if any(any(cell.value is not None for cell in ws[row])
                   for row in rows):
                continue
            # These are blank body spacers, outside the repeated frame.
            current = sum(height(row) for row in rows)
            if current > desired:
                # Keep every spacer row visible. Hiding a row immediately
                # after a manual page break can make Excel repaginate the
                # repeated document header above the printable page.
                scale = desired / current if current else 0.0
                for row in rows:
                    ws.row_dimensions[row].height = height(row) * scale
                    ws.row_dimensions[row].hidden = False

    def _space_ndt_section_caption(self, ws, gap=4.0):
        """Add space below SECTION captions before calculating page breaks."""
        from copy import copy

        for row in range(1, ws.max_row + 1):
            text = ''.join(str(ws.cell(row, col).value or '')
                           for col in range(1, min(ws.max_column, 23) + 1))
            if not ''.join(text.split()).upper().startswith(
                '1.2SECTION별검사'.upper()
            ):
                continue
            dimension = ws.row_dimensions[row]
            height = float(dimension.height if dimension.height is not None
                           else (ws.sheet_format.defaultRowHeight or 15))
            dimension.height = height + float(gap)
            dimension.hidden = False
            for cell in ws[row]:
                if cell.value is not None:
                    alignment = copy(cell.alignment)
                    alignment.vertical = 'top'
                    cell.alignment = alignment

    def _detail_page_overflow_row(self, ws, start_row, row_count):
        overflow = self._first_print_overflow_row(
            ws, start_row, row_count + 1, self._report_height_scale(ws)
        )
        # TOTAL may start the next page without moving the final record.
        return overflow

    def _remove_stale_table_breaks(self, ws, start_row, total_row):
        """Remove template breaks owned by a table that has been rewritten.

        The template also places breaks on TOTAL rows. Keeping those boundaries
        forced every short summary/detail table onto a separate page and left
        most of the following page blank.
        """
        ws.row_breaks.brk = [
            brk for brk in ws.row_breaks.brk
            if not start_row <= int(brk.id) <= total_row
        ]

    def _detail_continuation_overflow_row(self, ws, start_row, row_count):
        """Apply the actual printable-height limit on continuations."""
        return self._detail_page_overflow_row(ws, start_row, row_count)

    def _first_print_overflow_row(
        self, ws, start_row, row_count, scale_override=None,
        reserve_footer=True, capacity_points=None,
    ):
        """Return the first data row that falls onto the next printed page."""
        break_ids = sorted(int(brk.id) for brk in ws.row_breaks.brk)
        previous_breaks = [row for row in break_ids if row < start_row]
        page_start = (previous_breaks[-1] + 1) if previous_breaks else 1

        # A4 portrait printable height, adjusted by the template's print scale.
        margins = ws.page_margins
        scale = float(
            scale_override if scale_override is not None
            else (ws.page_setup.scale or 100)
        ) / 100.0
        # Footer is measured from the paper edge, inside the bottom margin.
        # The final 3.0 reflow must not subtract that space a second time.
        usable_inches = 11.69 - float(margins.top or 0) \
            - float(margins.bottom or 0)
        if reserve_footer:
            usable_inches -= float(margins.footer or 0)
        if capacity_points is None:
            capacity_points = usable_inches * 72.0 / scale
        default_height = float(ws.sheet_format.defaultRowHeight or 15)

        used_points = 0.0
        for row in range(page_start, start_row):
            dim = ws.row_dimensions[row]
            if not dim.hidden:
                used_points += float(dim.height if dim.height is not None else default_height)

        for row in range(start_row, start_row + row_count):
            dim = ws.row_dimensions[row]
            row_height = (0.0 if dim.hidden else
                          float(dim.height if dim.height is not None else default_height))
            if used_points + row_height > capacity_points:
                return row
            used_points += row_height
        return None

    def _compact_blank_spacers_before_table(self, ws, page_start, table_header_row):
        """Reclaim only unused spacer height without changing data rows."""
        merged_rows = set()
        for merged in ws.merged_cells.ranges:
            merged_rows.update(range(merged.min_row, merged.max_row + 1))

        image_rows = set()
        for image in list(ws._images):
            anchor = getattr(image, 'anchor', None)
            if hasattr(anchor, '_from'):
                image_rows.add(anchor._from.row + 1)

        compacted = 0
        for row in range(page_start, table_header_row):
            if row in merged_rows or row in image_rows:
                continue
            if any(
                str(ws.cell(row=row, column=col).value or '').strip()
                for col in range(1, min(ws.max_column, 24) + 1)
            ):
                continue
            current_height = float(
                ws.row_dimensions[row].height
                or ws.sheet_format.defaultRowHeight
                or 15
            )
            if current_height > 1.0:
                ws.row_dimensions[row].height = 1.0
                compacted += 1

        # Keep the visual top/bottom whitespace balanced while reclaiming
        # enough printable height for the PAUT detail table.
        if compacted:
            ws.page_margins.top = 0.35
            ws.page_margins.bottom = 0.35
        return compacted

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

    def _align_report_breaks_to_repeated_headers(self, ws):
        """Start every copied report header on a real printed page.

        Dynamic detail rows can push a template page downward.  Copying the
        continuation header alone is not sufficient: unless a manual break is
        also restored immediately before *every* repeated document header,
        Excel may leave the following template header in the middle of the
        same printed page and create a header-only overflow page.

        The exact title-row offset varies between template pages.  Locate the
        contract/report-title row immediately above each major section and put
        the break directly above that row.  This also avoids mistaking the
        table-of-contents entries for real page headers.
        """
        import copy
        from openpyxl.worksheet.pagebreak import Break

        section_titles = (
            '3.0비파괴검사현황',
            '4.0용접결함,불량률현황',
            '5.0안전관리활동및교육현황',
            '6.0방사선안전관리통합정보망보고자료',
            '6.0사진대지',
            '7.0사진대지',
        )
        required_breaks = set()
        obsolete_breaks = set()
        for row in range(1, ws.max_row + 1):
            row_text = ''.join(
                str(ws.cell(row=row, column=col).value or '')
                for col in range(1, min(ws.max_column, 23) + 1)
            )
            compact = ''.join(row_text.split())
            if not any(title in compact for title in section_titles):
                continue

            header_start = None
            for candidate in range(max(1, row - 6), row + 1):
                candidate_text = ''.join(
                    str(ws.cell(row=candidate, column=col).value or '')
                    for col in range(1, min(ws.max_column, 23) + 1)
                )
                candidate_compact = ''.join(candidate_text.split())
                if '월간용역진도보고서' in candidate_compact:
                    header_start = candidate
                    break
            if header_start and header_start > 1:
                # Hidden spacer rows inside the seven-row report header make
                # Excel suppress their cell borders. Keep them at minimum
                # height instead and explicitly continue the outer frame.
                for spacer_row in range(header_start + 5, header_start + 7):
                    if not ws.row_dimensions[spacer_row].hidden:
                        continue
                    ws.row_dimensions[spacer_row].hidden = False
                    ws.row_dimensions[spacer_row].height = 1.0

                    left_cell = ws.cell(row=spacer_row, column=2)
                    left_border = copy.copy(left_cell.border)
                    left_border.left = copy.copy(
                        ws.cell(row=header_start, column=2).border.left
                    )
                    left_cell.border = left_border

                    right_cell = ws.cell(row=spacer_row, column=23)
                    right_border = copy.copy(right_cell.border)
                    right_border.right = copy.copy(
                        ws.cell(row=header_start, column=23).border.right
                    )
                    right_cell.border = right_border

                break_id = header_start - 1
                # The page boundary belongs directly above the visible report
                # header. Remove both the former one-row-early break and every
                # stale break left *inside* the seven-row header.  Otherwise a
                # header beginning at row 483 can retain the old break at 487,
                # creating a page that contains only rows 483~487.
                obsolete_breaks.add(header_start - 2)
                obsolete_breaks.update(
                    range(header_start, header_start + 7)
                )
                required_breaks.add(break_id)

        existing = {
            int(brk.id): brk for brk in ws.row_breaks.brk
            if int(brk.id) not in obsolete_breaks
        }
        for break_id in required_breaks:
            existing.setdefault(break_id, Break(id=break_id))
        ws.row_breaks.brk = [existing[key] for key in sorted(existing)]

    def _pad_ndt_continuation_pages(self, ws):
        """Give each generated 3.0 continuation a complete page-height block.

        A continuation used to consist only of its repeated header, remaining
        records, and TOTAL row.  The next template page therefore started
        immediately after TOTAL (for example at row 492), leaving a visibly
        short page.  Use the preceding 3.0 page span as the page-height
        blueprint and insert cleared spacer rows after TOTAL so the following
        document starts only after the full continuation page.
        """
        header_starts = []
        for row in range(1, ws.max_row + 1):
            title = ''.join(
                str(ws.cell(row=row, column=col).value or '')
                for col in range(1, min(ws.max_column, 23) + 1)
            )
            if '3.0비파괴검사현황' not in ''.join(title.split()):
                continue
            for candidate in range(max(1, row - 6), row + 1):
                candidate_text = ''.join(
                    str(ws.cell(row=candidate, column=col).value or '')
                    for col in range(1, min(ws.max_column, 23) + 1)
                )
                if '월간용역진도보고서' in ''.join(candidate_text.split()):
                    header_starts.append(candidate)
                    break

        if len(header_starts) < 3:
            return

        index = 1
        while index < len(header_starts) - 1:
            page_start = header_starts[index]
            next_start = header_starts[index + 1]
            # The continuation caption is finalized later in some generation
            # paths, so use the structural page span rather than label text.
            # A shorter span between two repeated 3.0 headers is precisely the
            # partial-page condition this method must eliminate.
            blueprint_rows = page_start - header_starts[index - 1]
            current_rows = next_start - page_start
            padding_rows = blueprint_rows - current_rows
            if padding_rows <= 0:
                index += 1
                continue

            # Size the padding from the *printable height*, not from a fixed
            # row height.  A fixed 13.5pt spacer can push the last few rows
            # onto an automatic Excel page (for example rows 483~487) while
            # the next report header still has its manual break at row 487.
            # That combination creates a nearly blank extra page.
            default_height = float(ws.sheet_format.defaultRowHeight or 15)
            existing_points = sum(
                float(ws.row_dimensions[row].height or default_height)
                for row in range(page_start, next_start)
            )
            margins = ws.page_margins
            scale = float(ws.page_setup.scale or 100) / 100.0
            usable_inches = (
                11.69
                - float(margins.top or 0)
                - float(margins.bottom or 0)
                - float(margins.footer or 0)
            )
            capacity_points = usable_inches * 72.0 / scale
            available_padding_points = max(
                float(padding_rows), capacity_points - existing_points
            )
            spacer_height = max(
                1.0,
                min(13.5, available_padding_points / padding_rows),
            )

            # A break before the next header belongs AFTER the new padding.
            # The generic row inserter shifts breaks at/after next_start, but
            # not this boundary at next_start - 1. Leaving it behind prints
            # the inserted whitespace as a separate, completely blank page.
            boundary_breaks = [
                brk for brk in ws.row_breaks.brk
                if int(brk.id) == next_start - 1
            ]
            self._insert_rows_safely(ws, next_start, padding_rows)
            for brk in boundary_breaks:
                brk.id += padding_rows

            # insert_rows() does not move drawings; keep every later logo and
            # photo aligned with the report section it belongs to.
            for image in list(ws._images):
                anchor = getattr(image, 'anchor', None)
                if not hasattr(anchor, '_from') or anchor._from.row + 1 < next_start:
                    continue
                anchor._from.row += padding_rows
                if hasattr(anchor, 'to'):
                    anchor.to.row += padding_rows

            # Keep every padding row visible while fitting the complete block
            # before the next manual report-header boundary on one page.
            for spacer_row in range(next_start, next_start + padding_rows):
                ws.row_dimensions[spacer_row].height = spacer_height

            for key, marker_row in self.table_markers.items():
                if marker_row >= next_start:
                    self.table_markers[key] += padding_rows
            for later_index in range(index + 1, len(header_starts)):
                header_starts[later_index] += padding_rows
            index += 1

    def _normalize_ndt_header_logos(self, ws):
        """Keep one logo in each 3.0 header and none inside its data tables."""
        import copy

        header_starts = []
        for row in range(1, ws.max_row + 1):
            row_text = ''.join(
                str(ws.cell(row=row, column=col).value or '')
                for col in range(1, min(ws.max_column, 23) + 1)
            )
            if '3.0비파괴검사현황' not in ''.join(row_text.split()):
                continue
            for candidate in range(max(1, row - 6), row + 1):
                candidate_text = ''.join(
                    str(ws.cell(row=candidate, column=col).value or '')
                    for col in range(1, min(ws.max_column, 23) + 1)
                )
                if '월간용역진도보고서' in ''.join(candidate_text.split()):
                    header_starts.append(candidate)
                    break
        if not header_starts:
            return

        logo_images = []
        for image in list(ws._images):
            anchor = getattr(image, 'anchor', None)
            if not hasattr(anchor, '_from'):
                continue
            image_row = anchor._from.row + 1
            image_col = anchor._from.col + 1
            if image_col <= 5 and float(image.width) < 300 and float(image.height) < 100:
                logo_images.append(image)

        reference = next(
            (
                image for image in logo_images
                if header_starts[0] <= image.anchor._from.row + 1 <= header_starts[0] + 6
            ),
            None,
        )
        if reference is None:
            return
        reference_anchor = copy.deepcopy(reference.anchor)
        reference_start = header_starts[0]

        keep_ids = set()
        for header_start in header_starts:
            candidates = [
                image for image in logo_images
                if header_start <= image.anchor._from.row + 1 <= header_start + 6
            ]
            if not candidates:
                continue
            keeper = candidates[0]
            keeper.anchor = copy.deepcopy(reference_anchor)
            row_shift = header_start - reference_start
            keeper.anchor._from.row += row_shift
            if hasattr(keeper.anchor, 'to'):
                keeper.anchor.to.row += row_shift
            keeper.width = reference.width
            keeper.height = reference.height
            keep_ids.add(id(keeper))

        continuation_bodies = []
        for index in range(len(header_starts) - 1):
            page_start = header_starts[index]
            page_end = header_starts[index + 1]
            page_text = ''.join(
                str(ws.cell(row=row, column=col).value or '')
                for row in range(page_start, page_end)
                for col in range(1, min(ws.max_column, 23) + 1)
            )
            if '(계속)' in page_text:
                continuation_bodies.append((page_start + 7, page_end))

        cleaned_images = []
        for image in list(ws._images):
            if image not in logo_images:
                cleaned_images.append(image)
                continue
            image_row = image.anchor._from.row + 1
            inside_ndt_header = any(
                header_start <= image_row <= header_start + 6
                for header_start in header_starts
            )
            inside_continuation_body = any(
                body_start <= image_row < body_end
                for body_start, body_end in continuation_bodies
            )
            if id(image) in keep_ids or (
                not inside_ndt_header and not inside_continuation_body
            ):
                cleaned_images.append(image)
        ws._images = cleaned_images

    def _normalize_all_report_header_logos(self, ws, reference_toc=False):
        """Place one identically anchored logo in every repeated report header."""
        import copy
        from io import BytesIO
        from openpyxl.drawing.image import Image as XLImage

        header_starts = []
        first_row = 1 if reference_toc else 300
        framed_starts = {m.min_row for m in ws.merged_cells.ranges
                         if m.min_col == 6 and m.max_col == 15
                         and m.max_row >= m.min_row + 2}
        for row in range(first_row, ws.max_row + 1):
            if reference_toc and framed_starts and row not in framed_starts:
                continue
            row_text = ''.join(
                str(ws.cell(row=row, column=col).value or '')
                for col in range(1, min(ws.max_column, 23) + 1)
            )
            if '월간용역진도보고서' in ''.join(row_text.split()):
                header_starts.append(row)
        if not header_starts:
            return

        logo_images = []
        for image in list(ws._images):
            anchor = getattr(image, 'anchor', None)
            if not hasattr(anchor, '_from'):
                continue
            if (
                anchor._from.col + 1 <= 5
                and float(image.width) < 300
                and float(image.height) < 100
            ):
                logo_images.append(image)

        reference_start = next(
            (
                start for start in header_starts
                if any(
                    start <= image.anchor._from.row + 1 <= start + 6
                    for image in logo_images
                )
            ),
            None,
        )
        if reference_start is None:
            return
        if reference_toc:
            toc_start = next((start for start in header_starts if any(
                ''.join(str(ws.cell(row, col).value or '').split()) == '목차'
                for row in range(start, min(start + 7, ws.max_row + 1))
                for col in range(1, min(ws.max_column, 23) + 1)
            ) and any(start <= image.anchor._from.row + 1 <= start + 6
                      for image in logo_images)), None)
            if toc_start is None:
                return
            reference_start = toc_start
        reference = next(
            image for image in logo_images
            if reference_start <= image.anchor._from.row + 1 <= reference_start + 6
        )
        reference_anchor = copy.deepcopy(reference.anchor)
        if reference_toc:
            from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor
            from openpyxl.drawing.xdr import XDRPositiveSize2D
            from openpyxl.utils.units import pixels_to_EMU
            from openpyxl.utils import get_column_letter

            extent = copy.deepcopy(getattr(reference_anchor, 'ext', None))
            if extent is None and hasattr(reference_anchor, 'to'):
                end = reference_anchor.to
                begin = reference_anchor._from
                width = end.colOff - begin.colOff
                for col in range(begin.col, end.col):
                    dim = ws.column_dimensions[get_column_letter(col + 1)]
                    if not dim.hidden:
                        width += pixels_to_EMU(int(float(dim.width or 8.43) * 7 + 5))
                height = end.rowOff - begin.rowOff
                for row in range(begin.row + 1, end.row + 1):
                    dim = ws.row_dimensions[row]
                    if not dim.hidden:
                        height += int(float(dim.height if dim.height is not None
                                            else (ws.sheet_format.defaultRowHeight or 15)) * 12700)
                if width > 0 and height > 0:
                    extent = XDRPositiveSize2D(width, height)

            # Anchor at the frame's first row with a physical offset. Copying
            # a row index alone changes the visible inset in taller headers.
            marker = copy.deepcopy(reference_anchor._from)
            default = float(ws.sheet_format.defaultRowHeight or 15)
            for row in range(reference_start, marker.row + 1):
                dim = ws.row_dimensions[row]
                if not dim.hidden:
                    marker.rowOff += int(float(
                        dim.height if dim.height is not None else default
                    ) * 12700)
            marker.row = reference_start - 1
            reference_anchor = OneCellAnchor(
                _from=marker,
                ext=extent or XDRPositiveSize2D(
                    pixels_to_EMU(reference.width),
                    pixels_to_EMU(reference.height),
                ),
            )

        def clone_logo(source_image):
            buffer = BytesIO()
            image_ref = source_image.ref
            if hasattr(image_ref, 'copy') and hasattr(image_ref, 'save'):
                image_format = getattr(image_ref, 'format', None) or 'PNG'
                image_ref.copy().save(buffer, format=image_format)
            elif isinstance(image_ref, (str, os.PathLike)):
                with open(image_ref, 'rb') as image_file:
                    buffer.write(image_file.read())
            else:
                old_position = image_ref.tell() if hasattr(image_ref, 'tell') else None
                if hasattr(image_ref, 'seek'):
                    image_ref.seek(0)
                buffer.write(image_ref.read())
                if old_position is not None and hasattr(image_ref, 'seek'):
                    image_ref.seek(old_position)
            buffer.seek(0)
            cloned = XLImage(buffer)
            cloned._report_header_stream = buffer
            cloned.width = source_image.width
            cloned.height = source_image.height
            return cloned

        keep_ids = set()
        for header_start in header_starts:
            candidates = [
                image for image in logo_images
                if header_start <= image.anchor._from.row + 1 <= header_start + 6
            ]
            if candidates:
                keeper = candidates[0]
            else:
                keeper = clone_logo(reference)
                ws.add_image(keeper)
                logo_images.append(keeper)
            keeper.anchor = copy.deepcopy(reference_anchor)
            row_shift = header_start - reference_start
            keeper.anchor._from.row += row_shift
            if hasattr(keeper.anchor, 'to'):
                keeper.anchor.to.row += row_shift
            if reference_toc:
                # Excel requires rowOff to lie within its anchor row. A
                # larger offset may have correct coordinates but a clipped
                # drawing in PDF. Convert the physical inset to valid cells.
                marker = keeper.anchor._from
                for _ in range(7):
                    dim = ws.row_dimensions[marker.row + 1]
                    row_emu = 0 if dim.hidden else int(float(
                        dim.height if dim.height is not None else default
                    ) * 12700)
                    if row_emu > 0 and marker.rowOff < row_emu:
                        break
                    marker.rowOff -= row_emu
                    marker.row += 1
            keeper.width = reference.width
            keeper.height = reference.height
            keep_ids.add(id(keeper))

        # After dynamic row insertion, old header logos can remain behind in
        # tables or page bottoms.  Every report page now has a normalized
        # keeper, so any other logo-sized drawing in the report body is stale.
        ws._images = [
            image for image in ws._images
            if not (
                image in logo_images
                and image.anchor._from.row + 1 >= first_row
                and id(image) not in keep_ids
            )
        ]

    def _normalize_report_header_top_borders(self, ws):
        """Use one thin top edge across every repeated document header.

        One legacy 3.0 title block in the template uses a medium top border
        while its left and right header blocks use thin borders.  Excel keeps
        that mismatch when the header is copied, making only the contract
        title above ``2026`` look darker.  Detect the document-header geometry
        instead of matching localized text so this remains stable for future
        report years and title wording.
        """
        import copy

        thin = Side(style='thin', color='000000')
        merged_ranges = list(ws.merged_cells.ranges)
        for title_block in merged_ranges:
            if not (
                title_block.min_row >= 300
                and title_block.min_col == 6
                and title_block.max_col == 15
                and title_block.max_row - title_block.min_row + 1 >= 3
            ):
                continue

            # A real document header always has the document-number block
            # immediately to its right.  This excludes similarly sized table
            # merges elsewhere in the report body.
            has_document_number_block = any(
                block.min_row == title_block.min_row
                and block.min_col == 16
                and block.max_col == 23
                and block.max_row <= title_block.max_row
                for block in merged_ranges
            )
            if not has_document_number_block:
                continue

            for column in range(title_block.min_col, title_block.max_col + 1):
                cell = ws.cell(title_block.min_row, column)
                border = copy.copy(cell.border)
                border.top = copy.copy(thin)
                cell.border = border

    def _collapse_blank_rows_before_report_headers(self, ws):
        """Prevent a blank row becoming a one-row printed page.

        Excel can insert an automatic page break immediately before the last
        blank spacer row, followed by the template's manual break immediately
        before the next document header.  The spacer then prints as a page by
        itself.  A zero-height blank spacer preserves the header boundary while
        eliminating that phantom page.
        """
        header_starts = []
        for row in range(300, ws.max_row + 1):
            row_text = ''.join(
                str(ws.cell(row=row, column=col).value or '')
                for col in range(1, min(ws.max_column, 23) + 1)
            )
            if '월간용역진도보고서' in ''.join(row_text.split()):
                header_starts.append(row)

        image_rows = {
            image.anchor._from.row + 1
            for image in ws._images
            if hasattr(getattr(image, 'anchor', None), '_from')
        }
        for header_start in header_starts[1:]:
            spacer_row = header_start - 1
            is_blank = all(
                not str(ws.cell(row=spacer_row, column=col).value or '').strip()
                for col in range(1, min(ws.max_column, 23) + 1)
            )
            previous_is_blank = spacer_row > 1 and all(
                not str(ws.cell(row=spacer_row - 1, column=col).value or '').strip()
                for col in range(1, min(ws.max_column, 23) + 1)
            )
            if is_blank and not previous_is_blank and spacer_row not in image_rows:
                ws.row_dimensions[spacer_row].height = 1.0
                ws.row_dimensions[spacer_row].hidden = False

    def _compact_weld_defect_trailing_spacers(self, ws):
        """Compact only unused rows at the end of the 4.0 defect page."""
        section_rows = []
        for row in range(1, ws.max_row + 1):
            row_text = ''.join(
                str(ws.cell(row=row, column=col).value or '')
                for col in range(1, min(ws.max_column, 23) + 1)
            )
            compact = ''.join(row_text.split())
            if '4.0' in compact and '용접' in compact and '불량' in compact:
                section_rows.append(row)
        if not section_rows:
            return 0
        section_row = section_rows[-1]

        breaks = sorted(int(brk.id) for brk in ws.row_breaks.brk)
        later_breaks = [break_id for break_id in breaks if break_id >= section_row]
        if not later_breaks:
            return 0
        page_end = later_breaks[0]
        last_content_row = page_end
        while last_content_row > section_row and all(
            not str(ws.cell(row=last_content_row, column=col).value or '').strip()
            for col in range(1, min(ws.max_column, 23) + 1)
        ):
            last_content_row -= 1

        compacted = 0
        for row in range(last_content_row + 1, page_end + 1):
            ws.row_dimensions[row].height = 1.0
            ws.row_dimensions[row].hidden = False
            compacted += 1
        return compacted

    def _fit_photo_page_trailing_rows(self, ws):
        """Prevent a trailing photo-frame spacer from printing alone."""
        print_rows = [int(r) for r in re.findall(r'\$?[A-Z]+\$?(\d+)', str(ws.print_area or ''))]
        print_end = max(print_rows) if print_rows else ws.max_row
        ends = sorted({int(b.id) for b in ws.row_breaks.brk if int(b.id) < print_end}
                      | {print_end})
        default = float(ws.sheet_format.defaultRowHeight or 15)
        scale = self._report_height_scale(ws) / 100
        margins = ws.page_margins
        capacity = (11.69 - float(margins.top or 0) - float(margins.bottom or 0)
                    - float(margins.footer or 0)) * 72 / scale
        def height(r):
            dim = ws.row_dimensions[r]
            return 0 if dim.hidden else float(dim.height if dim.height is not None else default)
        start = 1
        for end in ends:
            title = ''.join(str(ws.cell(r, c).value or '')
                            for r in range(start, min(end, start + 10) + 1)
                            for c in range(1, min(ws.max_column, 23) + 1))
            if '사진대지' not in ''.join(title.split()):
                start = end + 1
                continue
            image_bottom = start - 1
            for im in ws._images:
                anchor = getattr(im, 'anchor', None)
                if hasattr(anchor, '_from') and start <= anchor._from.row + 1 <= end:
                    bottom = getattr(getattr(anchor, 'to', None), 'row', end - 1) + 1
                    image_bottom = max(image_bottom, bottom)
            excess = sum(height(r) for r in range(start, end + 1)) - capacity
            for r in range(end, max(start - 1, image_bottom), -1):
                if excess <= 0:
                    break
                if any(ws.cell(r, c).value is not None
                       for c in range(1, min(ws.max_column, 23) + 1)):
                    break
                reduction = min(excess + .01, max(0, height(r) - 1))
                if reduction:
                    ws.row_dimensions[r].height = height(r) - reduction
                    excess -= reduction
            start = end + 1

    def _compact_rt_leading_gap(self, ws, header_row):
        """Reclaim the blank MT-to-RT gap before RT pagination is decided."""
        protected_rows = {r for m in ws.merged_cells.ranges
                          for r in range(m.min_row, m.max_row + 1)}
        protected_rows.update(im.anchor._from.row + 1 for im in ws._images
                              if hasattr(im.anchor, '_from'))
        # header_row - 1 is the subsection caption; leave it intact.
        row = header_row - 2
        while row > 1 and not any(
            ws.cell(row, col).value is not None
            for col in range(1, min(ws.max_column, 23) + 1)
        ):
            if row not in protected_rows and not ws.row_dimensions[row].hidden:
                current = ws.row_dimensions[row].height
                ws.row_dimensions[row].height = min(
                    float(current if current is not None else ws.sheet_format.defaultRowHeight or 15),
                    2.0,
                )
            row -= 1

    def _compact_ndt_method_spacers(self, ws):
        """Keep a trailing PT table with its repeated 3.0 page header."""
        for row in range(1, ws.max_row + 1):
            row_text = ''.join(
                str(ws.cell(row=row, column=col).value or '')
                for col in range(1, min(ws.max_column, 23) + 1)
            )
            compact = ''.join(row_text.split()).upper()
            if '2.4' not in compact or 'PT' not in compact:
                continue
            spacer_row = row - 1
            while spacer_row > 1 and all(
                not str(ws.cell(row=spacer_row, column=col).value or '').strip()
                for col in range(1, min(ws.max_column, 23) + 1)
            ):
                ws.row_dimensions[spacer_row].height = 1.0
                ws.row_dimensions[spacer_row].hidden = False
                spacer_row -= 1

    def _ensure_pt_continuation_header(self, ws):
        """Repeat the 3.0 document header when PT prints on a new page."""
        from openpyxl.worksheet.pagebreak import Break

        pt_title_row = None
        for row in range(1, ws.max_row + 1):
            row_text = ''.join(
                str(ws.cell(row=row, column=col).value or '')
                for col in range(1, min(ws.max_column, 23) + 1)
            )
            compact = ''.join(row_text.split()).upper()
            if '2.4' in compact and 'PT' in compact:
                pt_title_row = row
                break
        if pt_title_row is None:
            return 0

        pt_total_row = None
        for row in range(pt_title_row + 1, min(pt_title_row + 10, ws.max_row) + 1):
            if any(
                'TOTAL' in str(ws.cell(row=row, column=col).value or '').upper()
                for col in range(1, min(ws.max_column, 23) + 1)
            ):
                pt_total_row = row
                break
        if pt_total_row is None:
            return 0

        breaks = sorted(int(brk.id) for brk in ws.row_breaks.brk)
        previous = [break_id for break_id in breaks if break_id < pt_title_row]
        document_header_row = previous[-1] + 1 if previous else 1
        overflow = self._first_print_overflow_row(
            ws,
            document_header_row,
            pt_total_row - document_header_row + 1,
            self._paut_print_scale(ws),
        )
        if overflow is None or overflow > pt_total_row:
            return 0

        inserted_rows = 7
        self._insert_continuation_header(
            ws, pt_title_row, document_header_row, row_count=inserted_rows
        )
        if not any(int(brk.id) == pt_title_row - 1 for brk in ws.row_breaks.brk):
            ws.row_breaks.append(Break(id=pt_title_row - 1))
        ws.row_breaks.brk.sort(key=lambda brk: int(brk.id))
        for key, marker_row in self.table_markers.items():
            if marker_row >= pt_title_row:
                self.table_markers[key] += inserted_rows
        return inserted_rows

    def _move_mt_criteria_to_page6(self, ws):
        """Place the MT criteria block on page 6 before the page-7 header."""
        import copy
        from openpyxl.worksheet.pagebreak import Break

        header_start, header_end = 230, 236
        mt_start, mt_end = 237, 242
        break_ids = {int(brk.id) for brk in ws.row_breaks.brk}
        if 229 not in break_ids:
            return

        # Swap [page-7 header][MT criteria] into [MT criteria][page-7 header]
        # without changing any downstream row numbers.
        source_order = list(range(mt_start, mt_end + 1)) + list(
            range(header_start, header_end + 1)
        )
        target_rows = list(range(header_start, mt_end + 1))
        row_map = dict(zip(source_order, target_rows))
        max_col = ws.max_column

        cells = {}
        dimensions = {}
        for source_row in range(header_start, mt_end + 1):
            dimensions[source_row] = (
                ws.row_dimensions[source_row].height,
                ws.row_dimensions[source_row].hidden,
            )
            cells[source_row] = []
            for col in range(1, max_col + 1):
                cell = ws.cell(row=source_row, column=col)
                cells[source_row].append((
                    cell.value,
                    copy.copy(cell._style),
                    cell.number_format,
                ))

        affected_merges = [
            merged for merged in list(ws.merged_cells.ranges)
            if merged.min_row >= header_start and merged.max_row <= mt_end
        ]
        merge_snapshots = [
            (merged.min_row, merged.min_col, merged.max_row, merged.max_col)
            for merged in affected_merges
        ]
        for merged in affected_merges:
            ws.unmerge_cells(str(merged))

        for source_row in source_order:
            target_row = row_map[source_row]
            height, hidden = dimensions[source_row]
            ws.row_dimensions[target_row].height = height
            ws.row_dimensions[target_row].hidden = hidden
            for col, (value, style, number_format) in enumerate(
                cells[source_row], start=1
            ):
                cell = ws.cell(row=target_row, column=col)
                cell.value = value
                cell._style = copy.copy(style)
                cell.number_format = number_format

        for min_row, min_col, max_row, max_col_idx in merge_snapshots:
            ws.merge_cells(
                start_row=row_map[min_row],
                start_column=min_col,
                end_row=row_map[max_row],
                end_column=max_col_idx,
            )

        for image in list(ws._images):
            anchor = getattr(image, 'anchor', None)
            if not hasattr(anchor, '_from'):
                continue
            image_row = anchor._from.row + 1
            if image_row not in row_map:
                continue
            shift = row_map[image_row] - image_row
            anchor._from.row += shift
            if hasattr(anchor, 'to'):
                anchor.to.row += shift

        # Page 6 now includes the six MT rows and ends immediately before the
        # relocated page-7 document header at row 236.
        ws.row_breaks.brk = [
            brk for brk in ws.row_breaks.brk if int(brk.id) != 229
        ]
        ws.row_breaks.append(Break(id=235))
        ws.row_breaks.brk = sorted(
            ws.row_breaks.brk, key=lambda brk: int(brk.id)
        )

    def _merge_pages7_and8(self, ws):
        """Combine the two sparse final pages of section 1.0."""
        import copy

        break_ids = {int(brk.id) for brk in ws.row_breaks.brk}
        if not {235, 278, 325}.issubset(break_ids):
            return

        # The second page repeats the seven-row common report header. Remove
        # that duplicate, move its body upward, and keep drawings aligned.
        delete_start, delete_count = 279, 7
        delete_end = delete_start + delete_count - 1
        retained_images = []
        for image in list(ws._images):
            anchor = getattr(image, 'anchor', None)
            if not hasattr(anchor, '_from'):
                retained_images.append(image)
                continue
            image_row = anchor._from.row + 1
            if delete_start <= image_row <= delete_end:
                continue
            if image_row > delete_end:
                anchor._from.row -= delete_count
                if hasattr(anchor, 'to'):
                    anchor.to.row -= delete_count
            retained_images.append(image)
        ws._images = retained_images

        self._delete_rows_safely(ws, delete_start, delete_count)
        for key, marker_row in self.table_markers.items():
            if marker_row > delete_end:
                self.table_markers[key] -= delete_count

        # Keep the equipment-specification table image.  Its anchor is shifted
        # with the merged page above; place it in the reserved blank area below
        # the PAUT minimum-specification heading, not over the preceding MT/RT
        # paragraphs.
        for image in ws._images:
            anchor = getattr(image, 'anchor', None)
            if not hasattr(anchor, '_from'):
                continue
            image_row = anchor._from.row + 1
            if 290 <= image_row <= 300 and float(image.width) > 500:
                heading_row = next((
                    row for row in range(290, 306)
                    if 'PAUT' in ''.join(
                        str(ws.cell(row=row, column=col).value or '')
                        for col in range(1, min(ws.max_column, 23) + 1)
                    ).upper()
                    and '최소사양' in ''.join(
                        str(ws.cell(row=row, column=col).value or '')
                        for col in range(1, min(ws.max_column, 23) + 1)
                    )
                ), 299)
                target_row = heading_row + 1
                row_shift = target_row - image_row
                anchor._from.row += row_shift
                if hasattr(anchor, 'to'):
                    anchor.to.row += row_shift
                image.height = 180
                if hasattr(anchor, 'to'):
                    anchor.to.row = anchor._from.row + 7
                    anchor.to.rowOff = 0
                break

        # Keep the three-line paragraph beginning with item "너." below the
        # real equipment-specification table, with one blank row between them.
        paragraph_start = None
        for row in range(300, 315):
            row_text = ''.join(
                str(ws.cell(row=row, column=col).value or '')
                for col in range(1, min(ws.max_column, 23) + 1)
            )
            if 'Raw Data' in row_text:
                paragraph_start = row
                break
        if paragraph_start is not None:
            paragraph_rows = list(range(paragraph_start, paragraph_start + 3))
            target_rows = list(range(paragraph_start + 4, paragraph_start + 7))
            if all(
                not str(ws.cell(row=row, column=col).value or '').strip()
                for row in target_rows
                for col in range(1, min(ws.max_column, 23) + 1)
            ):
                source_order = target_rows + paragraph_rows
                destination_order = paragraph_rows + target_rows
                row_map = dict(zip(source_order, destination_order))
                affected_rows = set(source_order)
                max_col = ws.max_column

                cells = {}
                dimensions = {}
                for source_row in source_order:
                    dimensions[source_row] = (
                        ws.row_dimensions[source_row].height,
                        ws.row_dimensions[source_row].hidden,
                    )
                    cells[source_row] = [
                        (
                            ws.cell(source_row, col).value,
                            copy.copy(ws.cell(source_row, col)._style),
                            ws.cell(source_row, col).number_format,
                        )
                        for col in range(1, max_col + 1)
                    ]

                affected_merges = [
                    merged for merged in list(ws.merged_cells.ranges)
                    if merged.min_row in affected_rows
                    and merged.max_row in affected_rows
                ]
                merge_snapshots = [
                    (
                        merged.min_row, merged.min_col,
                        merged.max_row, merged.max_col,
                    )
                    for merged in affected_merges
                ]
                for merged in affected_merges:
                    ws.unmerge_cells(str(merged))

                for source_row in source_order:
                    target_row = row_map[source_row]
                    height, hidden = dimensions[source_row]
                    ws.row_dimensions[target_row].height = height
                    ws.row_dimensions[target_row].hidden = hidden
                    for col, (value, style, number_format) in enumerate(
                        cells[source_row], start=1
                    ):
                        cell = ws.cell(target_row, col)
                        cell.value = value
                        cell._style = copy.copy(style)
                        cell.number_format = number_format

                for min_row, min_col, max_row, max_col_idx in merge_snapshots:
                    ws.merge_cells(
                        start_row=row_map[min_row], start_column=min_col,
                        end_row=row_map[max_row], end_column=max_col_idx,
                    )

        # Remove the former boundary between pages 7 and 8. The old row-325
        # boundary has already shifted to row 318 during deletion.
        ws.row_breaks.brk = [
            brk for brk in ws.row_breaks.brk if int(brk.id) != 278
        ]

        # Preserve all text-row heights. Collapse only truly empty, unmerged
        # spacer rows.  Seven-point spacers still left rows 312-313 on an
        # automatically-created eighth page, so keep just a one-point gap.
        merged_rows = set()
        for merged in ws.merged_cells.ranges:
            merged_rows.update(range(merged.min_row, merged.max_row + 1))
        image_rows = set()
        for image in ws._images:
            anchor = getattr(image, 'anchor', None)
            if not hasattr(anchor, '_from'):
                continue
            image_start = anchor._from.row + 1
            image_end = (
                anchor.to.row + 1
                if hasattr(anchor, 'to')
                else image_start
            )
            image_rows.update(range(image_start, image_end + 1))
        for row in range(236, 319):
            if row in merged_rows or row in image_rows:
                continue
            is_blank = all(
                not str(ws.cell(row=row, column=col).value or '').strip()
                for col in range(1, min(ws.max_column, 23) + 1)
            )
            if is_blank:
                ws.row_dimensions[row].height = 1.0

        # Rows 314-318 are the unused tail of the former eighth page.  Remove
        # them only when they contain no values, merges, or drawings, then end
        # page 7 at row 313.  All following report content moves up intact.
        trailing_start, trailing_end = 314, 318
        trailing_rows = set(range(trailing_start, trailing_end + 1))
        trailing_is_blank = all(
            not str(ws.cell(row=row, column=col).value or '').strip()
            for row in trailing_rows
            for col in range(1, ws.max_column + 1)
        )
        trailing_has_merge = any(
            merged.min_row <= trailing_end and merged.max_row >= trailing_start
            for merged in ws.merged_cells.ranges
        )
        trailing_has_image = any(
            hasattr(getattr(image, 'anchor', None), '_from')
            and trailing_start <= image.anchor._from.row + 1 <= trailing_end
            for image in ws._images
        )
        if trailing_is_blank and not trailing_has_merge and not trailing_has_image:
            for image in ws._images:
                anchor = getattr(image, 'anchor', None)
                if not hasattr(anchor, '_from') or anchor._from.row + 1 <= trailing_end:
                    continue
                anchor._from.row -= len(trailing_rows)
                if hasattr(anchor, 'to'):
                    anchor.to.row -= len(trailing_rows)

            self._delete_rows_safely(ws, trailing_start, len(trailing_rows))
            for key, marker_row in self.table_markers.items():
                if marker_row > trailing_end:
                    self.table_markers[key] -= len(trailing_rows)

            from openpyxl.worksheet.pagebreak import Break
            if not any(int(brk.id) == 313 for brk in ws.row_breaks.brk):
                ws.row_breaks.append(Break(id=313))
            ws.row_breaks.brk = sorted(
                ws.row_breaks.brk, key=lambda brk: int(brk.id)
            )

        # Section 1.0 now contains five pages instead of six.
        page_cells = []
        for row in range(80, 319):
            value = str(ws.cell(row=row, column=16).value or '')
            if 'of' in value.lower():
                page_cells.append(ws.cell(row=row, column=16))
        total_pages = len(page_cells)
        for page_number, cell in enumerate(page_cells, start=1):
            cell.value = (
                f"  \ucabd  \ubc88  \ud638 :      {page_number}     "
                f"of     {total_pages}"
            )

    def _widen_weld_defect_table(self, ws):
        """Extend each section-4 table without crossing continuation headers."""
        import copy

        has_cl = any(str(ws.cell(r, 16).value or '').strip().upper() == 'CL'
                     for r in range(1, ws.max_row + 1))
        source_columns = (17, 18, 19) if has_cl else (16, 17, 18)
        target_ranges = ((17, 18), (19, 21), (22, 23)) if has_cl else ((16, 17), (18, 20), (21, 23))
        merge_keys = {(m.min_col, m.max_col, m.min_row, m.max_row)
                      for m in ws.merged_cells.ranges}
        headers = [r for r in range(1, ws.max_row)
                   if str(ws.cell(r + 1, 2).value or '').strip().upper() == 'ID'
                   and all((c, c, r, r + 1) in merge_keys for c in source_columns)]
        if not headers:
            return

        blocks = []
        for index, header in enumerate(headers):
            limit = headers[index + 1] - 1 if index + 1 < len(headers) else ws.max_row
            total = next((r for r in range(header + 2, limit + 1)
                          if '총계' in str(ws.cell(r, 2).value or '')), None)
            if total is None and index + 1 == len(headers):
                continue
            blocks.append((header, total if total is not None else limit))

        # Process distinct blocks using original coordinates. Snapshot all
        # source values before any widening can cover their original columns.
        for header, end in reversed(blocks):
            for row, bottom in [(header, header + 1)] + [
                    (r, r) for r in range(header + 2, end + 1)]:
                full_row = next((m for m in ws.merged_cells.ranges
                                 if m.min_row == row == m.max_row
                                 and m.min_col == 2 and m.max_col in (source_columns[-1], 23)), None)
                if full_row is not None:
                    if full_row.max_col == source_columns[-1]:
                        style = copy.copy(ws.cell(row, 2)._style)
                        ws.unmerge_cells(str(full_row))
                        for col in range(2, 24):
                            ws.cell(row, col)._style = copy.copy(style)
                        ws.merge_cells(start_row=row, end_row=row,
                                       start_column=2, end_column=23)
                    continue
                snapshots = [(ws.cell(row, c).value, copy.copy(ws.cell(row, c)._style))
                             for c in source_columns]
                for merged in list(ws.merged_cells.ranges):
                    if (merged.min_row <= bottom and merged.max_row >= row
                            and merged.min_col <= 23 and merged.max_col >= source_columns[0]):
                        ws.unmerge_cells(str(merged))
                for r in range(row, bottom + 1):
                    for c in range(source_columns[0], 24):
                        ws.cell(r, c).value = None
                for (value, style), (left, right) in zip(snapshots, target_ranges):
                    for r in range(row, bottom + 1):
                        for c in range(left, right + 1):
                            ws.cell(r, c)._style = copy.copy(style)
                    ws.cell(row, left).value = value
                    ws.merge_cells(start_row=row, end_row=bottom,
                                   start_column=left, end_column=right)

        # Column widths are sheet-wide in Excel, so rebalance B:O while
        # preserving their combined width. This widens ID/name on both the
        # upper and lower defect tables and proportionally narrows the twelve
        # defect-detail columns without changing the printed table width.
        detail_columns = [
            openpyxl.utils.get_column_letter(col) for col in range(4, source_columns[0])
        ]
        original_id = float(ws.column_dimensions['B'].width or 8.43)
        original_name = float(ws.column_dimensions['C'].width or 8.43)
        original_detail_widths = {
            col: float(ws.column_dimensions[col].width or 8.43)
            for col in detail_columns
        }
        original_total = (
            original_id + original_name + sum(original_detail_widths.values())
        )
        new_id_width = 11.5
        new_name_width = 7.0
        remaining_detail_width = original_total - new_id_width - new_name_width
        detail_scale = remaining_detail_width / sum(
            original_detail_widths.values()
        )
        ws.column_dimensions['B'].width = new_id_width
        ws.column_dimensions['C'].width = new_name_width
        for col, width in original_detail_widths.items():
            ws.column_dimensions[col].width = width * detail_scale

    def _keep_iuc_euc_headers_single_line(self, ws):
        """Keep the narrow IUC/EUC column headers on a single line."""
        import copy

        for row in ws.iter_rows():
            for cell in row:
                compact = ''.join(str(cell.value or '').split()).upper()
                if compact not in {'IUC', 'EUC'}:
                    continue

                cell.value = compact
                alignment = copy.copy(cell.alignment)
                alignment.wrap_text = False
                alignment.shrink_to_fit = True
                cell.alignment = alignment

    def _clear_broken_reference_placeholders(self, ws):
        """Remove stale template formulas whose referenced sheet was removed."""
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.strip().upper() == '=#REF!':
                    cell.value = None

    def _fill_repeated_header_document_numbers(self, ws, doc_num):
        """Restore the document number in every repeated report header."""
        document_text = f'\ubb38\uc11c\ubc88\ud638 : \uc6d4\uac04\uc6a9\uc5ed\uc9c4\ub3c4\ubcf4\uace0\uc11c {doc_num}\ud638'
        updated = set()
        for row in range(1, ws.max_row + 1):
            for col in range(16, min(ws.max_column, 23) + 1):
                value = str(ws.cell(row=row, column=col).value or '')
                compact = value.replace(' ', '').replace('\n', '')
                if '\uac1c\uc815\ubc88\ud638' not in compact:
                    continue
                document_row = row - 2
                if document_row < 1 or (document_row, col) in updated:
                    continue
                target = ws.cell(row=document_row, column=col)
                if not isinstance(target, Cell):
                    continue
                target.value = document_text
                target.alignment = Alignment(
                    horizontal='left', vertical='center',
                    wrap_text=False, shrink_to_fit=True, indent=1,
                )
                target.font = Font(name='\ub9d1\uc740 \uace0\ub515', size=8)
                updated.add((document_row, col))

    def _rebalance_toc_page_layout(self, ws):
        """Spread the table-of-contents entries without moving its page break.

        The template reserves several alternating-height rows below the last
        visible TOC entry. Reuse that height for the seven visible entries and
        their intervening gaps, then divide the remainder across the blank rows
        at the foot of the page. Because the total height from the first entry
        through the manual page break is unchanged, every following report page
        keeps its existing start row and print position.
        """
        default_height = float(ws.sheet_format.defaultRowHeight or 15)
        break_rows = sorted(int(brk.id) for brk in ws.row_breaks.brk)

        # Locate the real TOC page by its common-header caption and page label;
        # this avoids matching the larger body title farther down the page.
        toc_header_row = None
        for row in range(1, ws.max_row + 1):
            values = [
                str(ws.cell(row=row, column=col).value or '').strip()
                for col in range(1, min(ws.max_column, 23) + 1)
            ]
            if '목 차' in values and any('쪽' in value and '번호' in value for value in values):
                toc_header_row = row
                break
        if toc_header_row is None:
            return

        page_end = next((row for row in break_rows if row >= toc_header_row), None)
        if page_end is None:
            return

        expected_labels = {f'{number}.0' for number in range(1, 8)}
        entry_rows = {}
        for row in range(toc_header_row + 1, page_end + 1):
            for col in range(1, min(ws.max_column, 12) + 1):
                value = str(ws.cell(row=row, column=col).value or '').strip()
                if value in expected_labels:
                    entry_rows[value] = row
                    break
        ordered_entries = [entry_rows.get(f'{number}.0') for number in range(1, 8)]
        if any(row is None for row in ordered_entries):
            return

        first_entry = ordered_entries[0]
        last_entry = ordered_entries[-1]
        original_height = sum(
            float(ws.row_dimensions[row].height or default_height)
            for row in range(first_entry, page_end + 1)
        )

        # Move the complete 1.0-7.0 block down while keeping the TOC title and
        # column headings fixed. The same height is removed from the blank rows
        # below 7.0, so the page break and every following page remain stable.
        content_offset = 50.0
        pre_entry_spacer = first_entry - 1
        ws.row_dimensions[pre_entry_spacer].height = (
            float(ws.row_dimensions[pre_entry_spacer].height or default_height)
            + content_offset
        )

        entry_height = 34.0
        gap_height = 18.0
        for row in ordered_entries:
            ws.row_dimensions[row].height = entry_height

        for current, following in zip(ordered_entries, ordered_entries[1:]):
            rows_between = list(range(current + 1, following))
            if not rows_between:
                continue
            per_row_height = gap_height / len(rows_between)
            for row in rows_between:
                ws.row_dimensions[row].height = per_row_height

        bottom_rows = list(range(last_entry + 1, page_end + 1))
        used_height = entry_height * len(ordered_entries) + gap_height * 6
        remaining_height = original_height - used_height - content_offset
        if bottom_rows and remaining_height > 0:
            bottom_height = remaining_height / len(bottom_rows)
            for row in bottom_rows:
                ws.row_dimensions[row].height = bottom_height

    def _normalize_repeated_report_header_text(self, ws, doc_num):
        """Write repeated-header text to the real merged-cell anchors.

        Row insertion can leave copied text in cells immediately outside the
        header merges while the actual F:O title anchor retains a stale formula.
        Excel evaluates that blank reference as 0. Locate each report header by
        its section caption and write all visible labels to the fixed merged
        anchors instead of relying on the copied cell value or formula.
        """
        contract_name = '2026년 중앙지사 열수송관 비파괴검사용역 단가계약'
        report_title = f'【{contract_name}】\n월 간 용 역 진 도 보 고 서'
        document_text = f'문서번호 : 월간용역진도보고서 {doc_num}호'
        section_markers = (
            '목 차',
            '1.0 용역업무일반',
            '2.0 인원,장비/자재 투입현황',
            '3.0 비파괴검사현황',
            '4.0 용접결함,불량율 현황',
            '5.0 안전관리활동 및 교육현황',
            '6.0 방사선안전관리통합정보망 보고',
            '6.0 사진대지',
            '7.0 사진대지',
        )

        normalized_headers = set()
        for section_row in range(1, ws.max_row + 1):
            row_values = [
                str(ws.cell(row=section_row, column=col).value or '').strip()
                for col in range(1, min(ws.max_column, 23) + 1)
            ]
            compact_values = [''.join(value.split()) for value in row_values]
            section_text = next(
                (
                    marker for marker in section_markers
                    if ''.join(marker.split()) in compact_values
                ),
                None,
            )
            if section_text is None:
                continue

            header_row = section_row - 4
            if header_row < 1 or header_row in normalized_headers:
                continue

            # A real repeated header has the title and metadata merge anchors at
            # F and P. This guard excludes body headings and TOC list entries.
            title_merge = next(
                (
                    merged for merged in ws.merged_cells.ranges
                    if merged.min_row == header_row
                    and merged.min_col == 6
                    and merged.max_col >= 15
                ),
                None,
            )
            metadata_merge = next(
                (
                    merged for merged in ws.merged_cells.ranges
                    if merged.min_row == header_row
                    and merged.min_col == 16
                    and merged.max_col >= 23
                ),
                None,
            )
            if title_merge is None or metadata_merge is None:
                continue

            title_cell = ws.cell(row=header_row, column=6)
            title_cell.value = report_title
            title_cell.font = Font(name='맑은 고딕', size=9, bold=True)
            title_cell.alignment = Alignment(
                horizontal='center', vertical='center', wrap_text=True,
            )

            doc_cell = ws.cell(row=header_row, column=16)
            doc_cell.value = document_text
            revision_cell = ws.cell(row=header_row + 2, column=16)
            revision_cell.value = '  개정번호 :            0'
            for cell in (doc_cell, revision_cell):
                cell.font = Font(name='맑은 고딕', size=8)
                cell.alignment = Alignment(
                    horizontal='left', vertical='center',
                    wrap_text=False, shrink_to_fit=True, indent=1,
                )

            # Normalize the section caption anchor as well. The page-number
            # anchor is deliberately left intact and is recalculated later.
            section_cell = ws.cell(row=section_row, column=6)
            section_cell.value = section_text
            section_cell.alignment = Alignment(
                horizontal='center', vertical='center',
                wrap_text=False, shrink_to_fit=True,
            )
            normalized_headers.add(header_row)

    def _fix_ndt_section_labels(self, ws):
        """Correct known text errors retained in the report template."""
        for row in ws.iter_rows():
            for cell in row:
                if not isinstance(cell.value, str):
                    continue
                if 'SPETION' in cell.value:
                    cell.value = cell.value.replace('SPETION', 'SECTION')

    def _fix_pt_summary_bottom_border(self, ws):
        """Match merged PT/MT label bottom edges to their adjacent tables."""
        import copy

        # Dynamic row insertion moves the summary table, so a fixed row such
        # as 390 eventually targets the wrong cells.  Locate every vertically
        # merged PT label and copy the actual bottom-table edge from column D.
        for merged in list(ws.merged_cells.ranges):
            if not (
                merged.min_col == 2
                and merged.max_col == 3
                and merged.max_row > merged.min_row
            ):
                continue
            anchor = ws.cell(row=merged.min_row, column=2)
            if str(anchor.value or '').strip().upper() not in {'PT', 'MT'}:
                continue

            bottom_row = merged.max_row
            reference_side = copy.copy(
                ws.cell(row=bottom_row, column=4).border.bottom
            )
            anchor_border = copy.copy(anchor.border)
            anchor_border.bottom = copy.copy(reference_side)
            anchor.border = anchor_border
            for col in (2, 3):
                cell = ws.cell(row=bottom_row, column=col)
                border = copy.copy(cell.border)
                border.bottom = copy.copy(reference_side)
                cell.border = border

    def _fit_long_section_headers(self, ws):
        """Keep long section captions fully visible in their merged cells."""
        import copy

        radiation_title = (
            '\ubc29\uc0ac\uc120\uc548\uc804\uad00\ub9ac'
            '\ud1b5\ud569\uc815\ubcf4\ub9dd\ubcf4\uace0\uc790\ub8cc'
        )
        for row in ws.iter_rows():
            for cell in row:
                if not isinstance(cell.value, str):
                    continue
                compact = ''.join(cell.value.split())
                if radiation_title not in compact:
                    continue
                alignment = copy.copy(cell.alignment)
                # The chapter caption in the repeated report header needs to
                # fit its box, but the numbered subsection (for example B713)
                # is a normal left-aligned heading and must not be shrunk into
                # the width of column B alone.
                is_numbered_subsection = bool(
                    re.match(r'^\s*\d+\.', cell.value)
                ) and cell.column == 2
                alignment.horizontal = (
                    'left' if is_numbered_subsection else 'center'
                )
                alignment.vertical = 'center'
                alignment.wrap_text = False
                alignment.shrink_to_fit = not is_numbered_subsection
                cell.alignment = alignment

    def _fit_inspector_qualification_rows(self, ws):
        """Expand the qualification rows on Page 5 using only blank spacer rows."""
        default_height = float(ws.sheet_format.defaultRowHeight or 15.0)
        target_height = 100.0
        rt_target_height = 118.0
        for row in range(1, ws.max_row + 1):
            role_text = ''.join(str(ws.cell(row=row, column=3).value or '').split())
            criteria_text = str(ws.cell(row=row, column=7).value or '')
            if '현장대리인' not in role_text or '다음 각 호중' not in criteria_text:
                continue
            current_height = float(ws.row_dimensions[row].height or default_height)
            needed = max(0.0, target_height - current_height)

            rt_row = None
            for candidate in range(row + 1, min(ws.max_row + 1, row + 8)):
                cand_spec = ''.join(str(ws.cell(row=candidate, column=5).value or '').split())
                cand_basis = str(ws.cell(row=candidate, column=20).value or '')
                if cand_spec == 'RT' and '제58조' in cand_basis:
                    rt_row = candidate
                    break
            if rt_row is not None:
                rt_current = float(ws.row_dimensions[rt_row].height or default_height)
                if rt_current < rt_target_height:
                    needed += rt_target_height - rt_current

            if needed <= 0:
                continue

            for spacer_row in range(row - 1, max(0, row - 16), -1):
                if needed <= 0:
                    break
                if any(
                    str(ws.cell(row=spacer_row, column=col).value or '').strip()
                    for col in range(1, min(ws.max_column, 23) + 1)
                ):
                    continue
                spacer_height = float(ws.row_dimensions[spacer_row].height or default_height)
                reclaim = min(needed, max(0.0, spacer_height - 10.0))
                if reclaim > 0:
                    ws.row_dimensions[spacer_row].height = spacer_height - reclaim
                    needed -= reclaim

            if current_height < target_height:
                ws.row_dimensions[row].height = target_height
            if rt_row is not None and float(ws.row_dimensions[rt_row].height or default_height) < rt_target_height:
                ws.row_dimensions[rt_row].height = rt_target_height

    def _ensure_cover_cell_elements(self, ws, create_date):
        """Use cell-native cover elements that survive openpyxl save cycles."""
        import copy

        thin = Side(style='thin', color='000000')

        def _merge_exactly(cell_range):
            from openpyxl.worksheet.cell_range import CellRange

            target = CellRange(cell_range)
            for merged in list(ws.merged_cells.ranges):
                overlaps = not (
                    merged.max_row < target.min_row
                    or merged.min_row > target.max_row
                    or merged.max_col < target.min_col
                    or merged.min_col > target.max_col
                )
                if overlaps and str(merged) != cell_range:
                    ws.unmerge_cells(str(merged))
            if not any(str(merged) == cell_range for merged in ws.merged_cells.ranges):
                ws.merge_cells(cell_range)

        # The original diagonal was an unsupported Line shape. A diagonal-up
        # border on one merged cell produces the same slash and is preserved.
        _merge_exactly('B4:C7')
        diagonal_cell = ws['B4']
        diagonal_cell.value = None
        diagonal_cell.border = Border(
            left=thin, right=thin, top=thin, bottom=thin,
            diagonal=thin, diagonalUp=True,
        )

        # The original date was an unsupported text box. Write the creation
        # date from the export dialog into a merged cell instead.
        _merge_exactly('N4:T6')
        date_cell = ws['N4']
        date_cell.value = str(create_date or '')
        # Match the cover's revision-number typography exactly.
        date_cell.font = copy.copy(ws['N3'].font)
        date_cell.alignment = Alignment(
            horizontal='center', vertical='center', wrap_text=False,
        )
        # 문서번호/개정번호와 동일하게 날짜는 외곽선 없이 표시한다.
        # 병합 전 템플릿에서 남은 가장자리 선도 병합 영역 전체에서 제거한다.
        for row in ws.iter_rows(min_row=4, max_row=6, min_col=14, max_col=20):
            for cell in row:
                cell.border = Border()
