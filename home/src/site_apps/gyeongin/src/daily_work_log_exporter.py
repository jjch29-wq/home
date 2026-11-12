import openpyxl
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
from openpyxl.utils import get_column_letter
import os

class DailyWorkLogExporter:
    def __init__(self):
        # Define basic styles
        self.font_title = Font(name='맑은 고딕', size=16, bold=True)
        self.font_normal = Font(name='맑은 고딕', size=9)
        self.font_bold = Font(name='맑은 고딕', size=9, bold=True)
        self.font_small = Font(name='맑은 고딕', size=8)

        self.align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
        self.align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
        self.align_right = Alignment(horizontal='right', vertical='center', wrap_text=True)
        self.align_nowrap = Alignment(horizontal='center', vertical='center', wrap_text=False, shrink_to_fit=True)

        thin = Side(style='thin')
        self.border_thin = Border(left=thin, right=thin, top=thin, bottom=thin)

        self.fill_header = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')

    def generate_excel(self, data, output_path):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "작업감독일보"

        # 1. Set Column Widths (A to AA roughly)
        col_widths = {
            'A': 5, 'B': 12, 'C': 8, 'D': 12, 'E': 26, 'F': 8, 'G': 8, 'H': 8, 'I': 13, 'J': 7, 'K': 7, 'L': 7,
            'M': 7, 'N': 6, 'O': 10, 'P': 10, 'Q': 10, 'R': 7, 'S': 7, 'T': 7, 'U': 4, 'V': 4, 'W': 4, 'X': 4, 'Y': 4, 'Z': 4
        }
        for col, width in col_widths.items():
            ws.column_dimensions[col].width = width

        # Helper to set cell value and style
        def set_cell(coord, value, font=self.font_normal, align=self.align_center, border=self.border_thin, fill=None):
            cell = ws[coord]
            cell.value = value
            cell.font = font
            cell.alignment = align
            if border: cell.border = border
            if fill: cell.fill = fill
            return cell

        def merge_and_set(range_str, value, font=self.font_normal, align=self.align_center, border=self.border_thin, fill=None):
            ws.merge_cells(range_str)
            start_cell = range_str.split(':')[0]
            cell = set_cell(start_cell, value, font, align, None, fill)
            # Apply border to all cells in merged range to avoid missing lines
            for row in ws[range_str]:
                for c in row:
                    if border: c.border = border
            return cell

        # Reduce spacer row heights to keep 1-page fit clean
        ws.row_dimensions[1].height = 6
        ws.row_dimensions[3].height = 6
        ws.row_dimensions[6].height = 6

        # =========================================================================
        # HEADER SECTION
        # =========================================================================
        merge_and_set('A2:R2', "비파괴검사 결과서 및 작업/감독일보", font=self.font_title, border=None)

        # Left Info
        merge_and_set('A4:F4', "용 역 명 : 경인사무소 비파괴검사용역", font=self.font_normal, align=self.align_left, border=None)
        date_str = data.get('date')
        if not date_str:
            import re
            m = re.search(r'(\d{4})[-_]?(\d{2})[-_]?(\d{2})', str(output_path))
            if m:
                date_str = f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
            else:
                date_str = '2026년   월   일'
        weather_str = data.get('weather', '')
        merge_and_set('A5:F5', f"검사일자 : {date_str}          날씨 : {weather_str}", font=self.font_normal, align=self.align_left, border=None)

        # Right Info (Signatures)

        set_cell('O4', '현장대리인', font=self.font_small, fill=self.fill_header)
        set_cell('P4', '감독', font=self.font_small, fill=self.fill_header)
        merge_and_set('Q4:R4', '확인', font=self.font_small, fill=self.fill_header)

        set_cell('O5', '') # Signature space
        set_cell('P5', '')
        merge_and_set('Q5:R5', '')
        ws.row_dimensions[5].height = 30 # Make signature box taller

        # Add signature image
        try:
            from openpyxl.drawing.image import Image
            sign_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'signs', '주진철.png')
            if os.path.exists(sign_path):
                img = Image(sign_path)
                # Resize image to fit the cell (approx width 70, height 38)
                img.width = 50
                img.height = 35

                # Use OneCellAnchor to center the image in O5
                from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
                from openpyxl.drawing.xdr import XDRPositiveSize2D
                from openpyxl.utils.units import pixels_to_EMU

                # O is column index 14, row 5 is index 4
                marker = AnchorMarker(col=14, colOff=pixels_to_EMU(12), row=4, rowOff=pixels_to_EMU(3))
                size = XDRPositiveSize2D(pixels_to_EMU(img.width), pixels_to_EMU(img.height))
                img.anchor = OneCellAnchor(_from=marker, ext=size)

                ws.add_image(img)
        except Exception as e:
            print(f"Error adding signature: {e}")

        # =========================================================================
        # SECTION 1: 작업 물량 및 누계 현황 (A7:J22)
        # =========================================================================
        if 'company_summary' in data:
            merge_and_set('A7:K7', f"1. 업체별 검사 실적 ({data.get('summary_start_date', '')} ~ {data.get('date', '')})", font=self.font_bold, align=self.align_left, border=None)
            headers = [('A8:C8', '업체명'), ('D8:D8', '검사방법'), ('E8:E8', '단위'),
                       ('F8:G8', '전일까지'), ('H8:I8', '선택일 실적'), ('J8:K8', '누계')]
            for cells, title in headers:
                merge_and_set(cells, title, font=self.font_bold, fill=self.fill_header)
            summary = data.get('company_summary', [])
            visible = summary[:19] if len(summary) > 20 else summary
            for row_idx, detail in enumerate(visible, 9):
                merge_and_set(f'A{row_idx}:C{row_idx}', detail['company'], align=self.align_left)
                set_cell(f'D{row_idx}', detail['method'])
                set_cell(f'E{row_idx}', detail['unit'])
                for cells, key in ((f'F{row_idx}:G{row_idx}', 'previous'),
                                   (f'H{row_idx}:I{row_idx}', 'today'), (f'J{row_idx}:K{row_idx}', 'total')):
                    cell = merge_and_set(cells, detail[key])
                    cell.number_format = '#,##0.####'
            if not summary:
                merge_and_set('A9:K9', '조회 기간에 저장된 검사 실적이 없습니다.', align=self.align_left)
            if len(summary) > 20:
                merge_and_set('A28:K28', f"전체 {len(summary)}행은 '업체별 실적' 시트 참조", align=self.align_left)
            # Every group is exported, including those beyond the main page's capacity.
            detail_sheet = wb.create_sheet('업체별 실적')
            detail_sheet.append(['업체별 검사 실적'])
            detail_sheet.append(['조회 시작일', data.get('summary_start_date', ''), '검사일자', data.get('date', '')])
            detail_sheet.append(['집계 기준', '저장된 일일 사용량 · RT REP 총 검사량 제외 · 단위별 구분'])
            detail_sheet.append(['업체명', '검사방법', '단위', '전일까지', '선택일 실적', '누계'])
            for detail in summary:
                detail_sheet.append([detail[key] for key in ('company', 'method', 'unit', 'previous', 'today', 'total')])
            for row in detail_sheet.iter_rows(min_row=4):
                for cell in row:
                    cell.border = self.border_thin
                    cell.font = self.font_bold if cell.row == 4 else self.font_normal
                    cell.alignment = self.align_left if cell.column == 1 else self.align_center
                    if cell.row == 4:
                        cell.fill = self.fill_header
                    elif cell.column >= 4:
                        cell.number_format = '#,##0.####'
            for column, width in zip('ABCDEF', (32, 16, 18, 18, 18, 18)):
                detail_sheet.column_dimensions[column].width = width
            detail_sheet.freeze_panes = 'A5'
            detail_sheet.auto_filter.ref = f'A4:F{max(4, detail_sheet.max_row)}'
            detail_sheet.print_title_rows = '1:4'
            detail_sheet.page_setup.orientation = 'landscape'
            detail_sheet.page_setup.fitToWidth = 1
            detail_sheet.page_setup.fitToHeight = 0
            detail_sheet.sheet_properties.pageSetUpPr.fitToPage = True
        else:
            merge_and_set('A7:F7', "1. 작업 물량 및 누계 현황", font=self.font_bold, align=self.align_left, border=None)

            headers_qty = ['방법', '규격', '예상량', '전일 누계', '금일 작업', '총 누계', '공정률(%)', '불량', '불량률(%)', '비고']
            for i, header in enumerate(headers_qty):
                col_letter = get_column_letter(i+1)
                set_cell(f'{col_letter}8', header, font=self.font_bold, fill=self.fill_header, align=self.align_nowrap)

            qty_rows = [
                ('PAUT', '300A이상'), ('PAUT', '300A이상-야간'), ('PAUT', '250A'), ('PAUT', '250A-야간'), ('PAUT', '200A'), ('PAUT', '200A-야간'),
                ('PAUT', '150A~125A'), ('PAUT', '150A~125A-야간'), ('PAUT', '100A이하'), ('PAUT', '100A이하-야간'), ('PAUT', '소계'),
                ('RT', '150A~100A'), ('RT', '150A~100A-야간'), ('RT', '80A이하'), ('RT', '80A이하-야간'), ('RT', '소계'),
                ('MT', '전체(주간)'), ('MT', '전체(야간)'),
                ('PT', '전체(주간)'), ('PT', '전체(야간)')
            ]

            row_idx = 9
            for method, spec in qty_rows:
                set_cell(f'A{row_idx}', method)
                set_cell(f'B{row_idx}', spec)

                # Fetch data from input dictionary
                row_data = data.get('qty_data', {}).get(f"{method}_{spec}", {})

                is_m_unit = method in ['PAUT', 'MT', 'PT']

                def set_qty_cell(coord, val_str, force_m_format=False):
                    if not val_str:
                        set_cell(coord, '')
                        return
                    try:
                        fval = float(str(val_str).replace(',', ''))
                        cell = ws[coord]
                        cell.value = fval
                        if force_m_format:
                            cell.number_format = '0.0000'
                        elif fval % 1 != 0:
                            cell.number_format = '0.0'
                        else:
                            cell.number_format = '0'
                        cell.font = self.font_normal
                        cell.alignment = self.align_center
                        cell.border = self.border_thin
                    except ValueError:
                        set_cell(coord, val_str)

                set_qty_cell(f'C{row_idx}', row_data.get('예상량', ''), force_m_format=False)
                set_qty_cell(f'D{row_idx}', row_data.get('전일누계', ''), force_m_format=is_m_unit)
                set_qty_cell(f'E{row_idx}', row_data.get('금일작업', ''), force_m_format=is_m_unit)
                set_qty_cell(f'F{row_idx}', row_data.get('총누계', ''), force_m_format=is_m_unit)

                set_qty_cell(f'G{row_idx}', row_data.get('공정률', ''))
                set_qty_cell(f'H{row_idx}', row_data.get('불량', ''))
                set_qty_cell(f'I{row_idx}', row_data.get('불량률', ''))
                set_cell(f'J{row_idx}', row_data.get('비고', ''))

                row_idx += 1

            ws.merge_cells('A9:A19') # PAUT
            ws.merge_cells('A20:A24') # RT
            ws.merge_cells('A25:A26') # MT
            ws.merge_cells('A27:A28') # PT

        # =========================================================================
        # RIGHT SIDE SECTION: Equipment and Personnel (L7:P23)
        # =========================================================================
        # Equipment
        merge_and_set('L8:N8', '장비투입 현황(대)', font=self.font_bold, fill=self.fill_header)
        set_cell('L9', '장비명', font=self.font_bold, fill=self.fill_header)
        set_cell('M9', '금일', font=self.font_bold, fill=self.fill_header)
        set_cell('N9', '누계', font=self.font_bold, fill=self.fill_header)

        equip_rows = ['PAUT장비', 'PAUT프로브', 'PAUT스캐너', 'RT장비', 'MT장비']
        for i, eq in enumerate(equip_rows):
            r = 10 + i
            set_cell(f'L{r}', eq, align=self.align_nowrap)
            eq_data = data.get('equip_data', {}).get(eq, {})
            set_cell(f'M{r}', eq_data.get('금일', ''))
            set_cell(f'N{r}', eq_data.get('누계', ''))

        # Personnel
        merge_and_set('O8:R8', '금일 투입인원(명)', font=self.font_bold, fill=self.fill_header)
        set_cell('O9', '구분(관리/안전)', font=self.font_small, fill=self.fill_header, align=self.align_nowrap)
        set_cell('P9', '검사원', font=self.font_bold, fill=self.fill_header, align=self.align_nowrap)
        merge_and_set('Q9:R9', '안전담당', font=self.font_bold, fill=self.fill_header, align=self.align_nowrap)

        personnel = data.get('personnel_data', {})
        set_cell('O10', '인원', align=self.align_nowrap)
        set_cell('P10', personnel.get('검사원_인원', ''), align=self.align_nowrap)
        merge_and_set('Q10:R10', personnel.get('안전_인원', ''), align=self.align_nowrap)

        set_cell('O11', '현장대리인', align=self.align_nowrap)
        set_cell('P11', personnel.get('검사원_현장대리인', ''), align=self.align_nowrap)
        merge_and_set('Q11:R11', personnel.get('안전_현장대리인', ''), align=self.align_nowrap)

        set_cell('O12', '누계', align=self.align_nowrap)
        set_cell('P12', personnel.get('검사원_누계', ''), align=self.align_nowrap)
        merge_and_set('Q12:R12', personnel.get('안전_누계', ''), align=self.align_nowrap)

        # Merge empty blocks under personnel to match equipment height
        merge_and_set('O13:R14', '', border=self.border_thin)

        # Remarks (특이사항 및 작업계획)
        merge_and_set('L15:R15', '특이사항 및 작업계획', font=self.font_bold, fill=self.fill_header)
        remarks = data.get('remarks', '')
        if 'company_memos' in data:
            from site_apps.gyeongin.src.company_memos import memo_snapshot, report_memo_text
            company_memos = memo_snapshot(data['company_memos'])
            remarks = report_memo_text(remarks, company_memos)
            if company_memos:
                notes = wb.create_sheet('업체별 메모')
                notes.append(['날짜', '업체 / 공통', '특이사항 및 계획'])
                if data.get('remarks', '').strip():
                    notes.append([data.get('date', ''), '공통', data['remarks']])
                for company, text in sorted(company_memos.items()):
                    notes.append([data.get('date', ''), company, text])
                for row in notes:
                    for cell in row:
                        cell.data_type = 's'
                        cell.font = self.font_bold if cell.row == 1 else self.font_normal
                        cell.border = self.border_thin
                        cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
                        if cell.row == 1:
                            cell.fill = self.fill_header
                    lines = sum(max(1, (len(line) + 59) // 60) for line in str(row[2].value or '').split('\n'))
                    notes.row_dimensions[row[0].row].height = min(409, max(24, lines * 18))
                for column, width in zip('ABC', (16, 28, 90)):
                    notes.column_dimensions[column].width = width
                notes.freeze_panes = 'A2'
                notes.print_title_rows = '1:1'
                notes.page_setup.orientation = 'landscape'
                notes.page_setup.fitToWidth = 1
                notes.page_setup.fitToHeight = 0
                notes.sheet_properties.pageSetUpPr.fitToPage = True
        merge_and_set('L16:R27', remarks, align=Alignment(horizontal='left', vertical='top', wrap_text=True))

        # =========================================================================
        # SECTION 2: 비파괴검사결과서 (A29:T...)
        # =========================================================================
        merge_and_set('A29:F29', "2. 비파괴검사결과서", font=self.font_bold, align=self.align_left, border=None)

        raw_ndt = data.get('ndt_results', [])
        ndt_results = [r for r in raw_ndt if any(str(v).strip() for v in r.values())] if raw_ndt else []
        non_pipe_rows = [r for r in ndt_results if self._is_non_pipe_row(r)]
        only_non_pipe = bool(ndt_results) and len(non_pipe_rows) == len(ndt_results)
        start_row = 32
        total_rows = max(8, len(ndt_results) + 1)

        if only_non_pipe:
            self._write_non_pipe_section2(ws, start_row, total_rows, ndt_results, set_cell, merge_and_set)
        else:
            self._write_pipe_section2(ws, start_row, total_rows, ndt_results, set_cell, merge_and_set)
            if non_pipe_rows:
                self._write_non_pipe_detail_sheet(wb, data.get('date', ''), non_pipe_rows)

        # Fit the table width; additional inspection rows continue on new pages.
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1 if total_rows <= 15 else 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.print_area = f'A1:T{start_row + total_rows - 1}'
        ws.print_title_rows = '30:31'
        ws.print_options.horizontalCentered = True

        # Header/Footer (Auto Page Numbering)
        ws.oddHeader.right.text = "Page &P of &N"
        ws.evenHeader.right.text = "Page &P of &N"

        # Margins
        ws.page_margins.left = 0.5
        ws.page_margins.right = 0.5
        ws.page_margins.top = 0.4
        ws.page_margins.bottom = 0.4
        ws.page_margins.header = 0.3

        wb.save(output_path)
        return output_path

    @staticmethod
    def _is_non_pipe_row(res):
        target = str(res.get('대상유형', '') or '').strip()
        if target in ('판재/시험편', '기계부품', '구조물/설비', 'PMI'):
            return True
        if target == '배관':
            return False
        has_pipe = any(str(res.get(k, '') or '').strip() for k in ('라인번호', 'Joint No.', '관경'))
        has_general = any(str(res.get(k, '') or '').strip() for k in ('검사품명', '식별번호', '치수원문', '수량'))
        return has_general and not has_pipe

    @staticmethod
    def _row_ident(res):
        for key in ('식별번호', '부품번호', '시험편번호', '설비번호', '태그번호', 'Joint No.'):
            val = str(res.get(key, '') or '').strip()
            if val:
                return val
        return ''

    @staticmethod
    def _row_material(res):
        for key in ('재질', '판정재질', '요구재질'):
            val = str(res.get(key, '') or '').strip()
            if val:
                return val
        return ''

    def _write_non_pipe_section2(self, ws, start_row, total_rows, ndt_results, set_cell, merge_and_set):
        headers = [
            ('A30:A31', '순번'), ('B30:B31', '업체'), ('C30:C31', '검사방법'), ('D30:D31', '대상유형'),
            ('E30:E31', '검사품명'), ('F30:F31', '식별번호'), ('G30:G31', '재질'), ('H30:H31', '두께'),
            ('I30:I31', '도면번호'), ('J30:M31', '치수·사양원문'), ('N30:N31', '결과'), ('O30:O31', '적용규격'),
            ('P30:Q31', '수량'), ('R30:R31', '단위'), ('S30:T31', '위치/비고'),
        ]
        for rng, text in headers:
            merge_and_set(rng, text, font=self.font_small, fill=self.fill_header, align=self.align_nowrap)

        for i in range(total_rows):
            row_idx = start_row + i
            res = ndt_results[i] if i < len(ndt_results) else {}
            set_cell(f'A{row_idx}', i + 1)
            set_cell(f'B{row_idx}', res.get('업체', ''))
            set_cell(f'C{row_idx}', res.get('검사방법', ''))
            set_cell(f'D{row_idx}', res.get('대상유형', ''), align=self.align_nowrap)
            set_cell(f'E{row_idx}', res.get('검사품명', ''), align=self.align_left)
            set_cell(f'F{row_idx}', self._row_ident(res), align=self.align_nowrap)
            set_cell(f'G{row_idx}', self._row_material(res), align=self.align_nowrap)
            set_cell(f'H{row_idx}', res.get('두께', ''))
            set_cell(f'I{row_idx}', res.get('도면번호', ''), align=self.align_nowrap)
            merge_and_set(f'J{row_idx}:M{row_idx}', res.get('치수원문', ''))
            set_cell(f'N{row_idx}', res.get('결과', ''))
            set_cell(f'O{row_idx}', res.get('적용규격', '') or res.get('규격', ''), align=self.align_nowrap)
            merge_and_set(f'P{row_idx}:Q{row_idx}', res.get('수량', ''))
            set_cell(f'R{row_idx}', res.get('단위', ''))
            loc_or_note = ' / '.join(filter(None, (str(res.get('검사위치', '') or '').strip(), str(res.get('비고', '') or '').strip())))
            merge_and_set(f'S{row_idx}:T{row_idx}', loc_or_note, font=self.font_small)

            if i == len(ndt_results):
                ws.unmerge_cells(f'J{row_idx}:M{row_idx}')
                ws.unmerge_cells(f'P{row_idx}:Q{row_idx}')
                ws.unmerge_cells(f'S{row_idx}:T{row_idx}')
                ws.merge_cells(f'B{row_idx}:T{row_idx}')
                ws[f'B{row_idx}'].value = "- 이 하 여 백 -"
                ws[f'B{row_idx}'].alignment = Alignment(horizontal='center', vertical='center')
            ws.row_dimensions[row_idx].height = 22

    def _write_pipe_section2(self, ws, start_row, total_rows, ndt_results, set_cell, merge_and_set):
        headers_ndt = [
            ('A30:A31', '순번'), ('B30:B31', '업체'), ('C30:C31', '검사방법'), ('D30:D31', '구간(Sec.No)'), ('E30:E31', '라인번호'),
            ('F30:F31', 'Joint No.'), ('G30:G31', '관경'), ('H30:H31', '두께'), ('I30:I31', '용접사'),
            ('N30:N31', '결과'), ('O30:O31', '규격')
        ]
        for rng, text in headers_ndt:
            merge_and_set(rng, text, font=self.font_small, fill=self.fill_header, align=self.align_nowrap)

        merge_and_set('J30:M30', '구간정보(Start/Length)', font=self.font_small, fill=self.fill_header)
        set_cell('J31', '1', font=self.font_small, fill=self.fill_header)
        set_cell('K31', '2', font=self.font_small, fill=self.fill_header)
        set_cell('L31', '3', font=self.font_small, fill=self.fill_header)
        set_cell('M31', '4', font=self.font_small, fill=self.fill_header)

        merge_and_set('P30:Q30', 'RT매수', font=self.font_small, fill=self.fill_header)
        set_cell('P31', 'OR', font=self.font_small, fill=self.fill_header)
        set_cell('Q31', 'RE', font=self.font_small, fill=self.fill_header)

        merge_and_set('R30:R31', 'PAUT(m)', font=self.font_small, fill=self.fill_header)
        merge_and_set('S30:S31', 'MT(m)', font=self.font_small, fill=self.fill_header)
        merge_and_set('T30:T31', 'PT(m)', font=self.font_small, fill=self.fill_header)

        for i in range(total_rows):
            row_idx = start_row + i
            res = ndt_results[i] if i < len(ndt_results) else {}
            non_pipe = bool(res) and self._is_non_pipe_row(res)

            set_cell(f'A{row_idx}', i + 1)
            set_cell(f'B{row_idx}', res.get('업체', ''))
            set_cell(f'C{row_idx}', res.get('검사방법', ''))
            set_cell(f'D{row_idx}', res.get('대상유형', '') if non_pipe else res.get('구간', ''), align=self.align_nowrap)
            line_text = str(res.get('검사품명', '') or '') if non_pipe else str(res.get('라인번호', '') or '')
            if res.get('작업구분') == '핫태핑':
                hot_text = f"핫태핑 {res.get('핫태핑개소', '')} / {res.get('핫태핑차수', '')} / {res.get('핫태핑개소수', '1')}개소"
                line_text = f'{line_text}\n{hot_text}'.strip()

            line_str = line_text
            num_lines = line_str.count('\n') + 1
            if num_lines > 1:
                set_cell(f'E{row_idx}', line_text, font=self.font_small, align=self.align_center)
            else:
                line_display_w = sum(2 if ord(c) > 127 else 1 for c in line_str)
                if line_display_w >= 34:
                    import copy
                    font_line = copy.copy(self.font_normal)
                    font_line.size = 7.5
                elif line_display_w >= 24:
                    import copy
                    font_line = copy.copy(self.font_normal)
                    font_line.size = 8.5
                else:
                    font_line = self.font_normal
                set_cell(f'E{row_idx}', line_text, font=font_line, align=self.align_nowrap)
            set_cell(f'F{row_idx}', self._row_ident(res) if non_pipe else res.get('Joint No.', ''), align=self.align_nowrap)
            set_cell(f'G{row_idx}', self._row_material(res) if non_pipe else res.get('관경', ''))
            set_cell(f'H{row_idx}', res.get('두께', ''))
            set_cell(f'I{row_idx}', res.get('도면번호', '') if non_pipe else res.get('용접사', ''))

            method = res.get('검사방법', '').upper().strip()
            sec_info_raw = (res.get('치수원문', '') if non_pipe else res.get('구간정보', '')).strip()

            is_empty_row = not any(res.values())
            if not is_empty_row and not sec_info_raw:
                sec_info_raw = "관련지시 없음"

            sec_info = sec_info_raw.split(',')
            sec_info += [''] * (4 - len(sec_info))

            if non_pipe or method in ['PAUT', 'PT', 'MT'] or sec_info_raw == "관련지시 없음":
                set_cell(f'J{row_idx}', sec_info_raw)
                set_cell(f'K{row_idx}', '')
                set_cell(f'L{row_idx}', '')
                set_cell(f'M{row_idx}', '')
                ws.merge_cells(f'J{row_idx}:M{row_idx}')
            else:
                set_cell(f'J{row_idx}', sec_info[0])
                set_cell(f'K{row_idx}', sec_info[1])
                set_cell(f'L{row_idx}', sec_info[2])
                set_cell(f'M{row_idx}', sec_info[3])

            spec_val = res.get('적용규격', '') if non_pipe else res.get('규격', '')
            shift_val = res.get('근무구분', '')
            if shift_val and shift_val != '주간':
                spec_val = f"{spec_val}({shift_val})" if spec_val else shift_val

            set_cell(f'N{row_idx}', res.get('결과', ''))
            set_cell(f'O{row_idx}', spec_val, align=self.align_nowrap)
            if str(res.get('검사방법', '')).strip().upper() == 'RT':
                from site_apps.gyeongin.src.rt_quantities import rt_counts
                rt_total, rt_rejects = rt_counts(res)
                set_cell(f'P{row_idx}', rt_total)
                set_cell(f'Q{row_idx}', rt_rejects)
            else:
                set_cell(f'P{row_idx}', res.get('RT_OR', ''))
                set_cell(f'Q{row_idx}', res.get('RT_RE', ''))
            set_cell(f'R{row_idx}', res.get('PAUT', ''))
            set_cell(f'S{row_idx}', res.get('MT', '') or (f"{res.get('수량', '')}{res.get('단위', '')}" if non_pipe and method == 'MT' else ''))
            set_cell(f'T{row_idx}', res.get('PT', '') or (f"{res.get('수량', '')}{res.get('단위', '')}" if non_pipe and method == 'PT' else ''))

            if i == len(ndt_results):
                ws.merge_cells(f'B{row_idx}:T{row_idx}')
                ws[f'B{row_idx}'].value = "- 이 하 여 백 -"
                ws[f'B{row_idx}'].alignment = Alignment(horizontal='center', vertical='center')

            ws.row_dimensions[row_idx].height = max(34, num_lines * 17) if (num_lines > 1 or res.get('작업구분') == '핫태핑') else 20

    def _write_non_pipe_detail_sheet(self, wb, date_text, rows):
        ws = wb.create_sheet('일반·부품 검사결과')
        ws.append(['일반·기계부품·구조물·PMI 비파괴검사 내역', f'검사일자: {date_text}'])
        headers = ['순번', '업체', '검사방법', '대상유형', '검사품명', '식별번호', '재질', '치수원문', '수량', '단위', '결과', '적용규격', '검사절차서', '도면번호', '검사위치', '비고']
        ws.append(headers)
        for idx, res in enumerate(rows, start=1):
            ws.append([
                idx, res.get('업체', ''), res.get('검사방법', ''), res.get('대상유형', ''),
                res.get('검사품명', ''), self._row_ident(res), self._row_material(res),
                res.get('치수원문', ''), res.get('수량', ''), res.get('단위', ''),
                res.get('결과', ''), res.get('적용규격', '') or res.get('규격', ''),
                res.get('검사절차서', ''), res.get('도면번호', ''), res.get('검사위치', ''), res.get('비고', ''),
            ])
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.border = self.border_thin
                cell.font = self.font_bold if cell.row == 2 else self.font_normal
                cell.alignment = self.align_left if cell.column in (5, 8, 16) else self.align_center
                if cell.row == 2:
                    cell.fill = self.fill_header
        for col, width in zip('ABCDEFGHIJKLMNOP', (6, 18, 10, 13, 20, 16, 12, 24, 9, 8, 9, 14, 14, 14, 14, 22)):
            ws.column_dimensions[col].width = width
        ws.page_setup.orientation = 'landscape'
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True

if __name__ == "__main__":
    # Test generation
    exporter = DailyWorkLogExporter()
    test_data = {
        'date': '2026년 08월 05일',
        'weather': '맑음',
        'qty_data': {
            '300A이상': {'예상량': 129, '전일누계': 100, '금일작업': 10, '총누계': 110, '공정률': '85.2%'}
        },
        'remarks': '안전 작업 지시 이행\nPAUT 장비 점검 완료',
        'ndt_results': [
            {'검사방법': 'PAUT', '구간': 'S-1', '라인번호': 'L-1', 'Joint No.': 'J-1', '결과': '합격'}
        ]
    }
    # exporter.generate_excel(test_data, "test_daily_report.xlsx")
