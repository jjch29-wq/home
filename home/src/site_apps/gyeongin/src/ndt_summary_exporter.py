import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

class NDTSummaryExporter:
    def __init__(self, history):
        from site_apps.gyeongin.src.rt_quantities import corrected_rt_history
        self.history = corrected_rt_history(history)
        self.font_normal = Font(name='맑은 고딕', size=10)
        self.font_bold = Font(name='맑은 고딕', size=10, bold=True)
        
        self.align_center = Alignment(horizontal='center', vertical='center', wrap_text=False)
        
        thin = Side(border_style="thin", color="000000")
        self.border_thin = Border(left=thin, right=thin, top=thin, bottom=thin)
        self.fill_header = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')

    def generate(self, output_path):
        wb = openpyxl.Workbook()
        
        # Group data by method
        grouped_data = {}
        
        # Sort dates
        sorted_dates = sorted(list(self.history.keys()))
        
        for date in sorted_dates:
            daily_data = self.history[date]
            ndt_results = daily_data.get('ndt_results', [])
            
            for row in ndt_results:
                method = str(row.get('검사방법', '')).strip().upper()
                if not method:
                    continue
                    
                if method not in grouped_data:
                    grouped_data[method] = []
                    
                # Inject date into the row dict for convenience
                row_copy = dict(row)
                row_copy['일자'] = date
                grouped_data[method].append(row_copy)
                
        # If no data at all
        if not grouped_data:
            ws = wb.active
            ws.title = "데이터 없음"
            ws['A1'] = "출력할 비파괴검사 결과가 없습니다."
            wb.save(output_path)
            return

        # 1. Create Summary Sheet First
        ws_summary = wb.active
        ws_summary.title = "총괄 집계표"
        self._write_summary_sheet(ws_summary, grouped_data)

        # 2. Create sheets for each method
        for method, rows in grouped_data.items():
            ws = wb.create_sheet(title=method)
            self._write_sheet(ws, method, rows)
            
        wb.save(output_path)
        
    @staticmethod
    def _row_meta(r, method):
        target_type = str(r.get('대상유형', '') or '').strip()
        is_non_pipe = target_type in ('판재/시험편', '기계부품', '구조물/설비', 'PMI') or (
            target_type != '배관'
            and any(str(r.get(k, '') or '').strip() for k in ('검사품명', '식별번호', '치수원문', '수량'))
            and not any(str(r.get(k, '') or '').strip() for k in ('라인번호', 'Joint No.', '관경'))
        )
        if not target_type:
            target_type = '일반' if is_non_pipe else '배관'
        ident = next((str(r.get(k, '') or '').strip() for k in ('식별번호', '부품번호', '시험편번호', '설비번호', '태그번호', 'Joint No.') if str(r.get(k, '') or '').strip()), '')
        material = next((str(r.get(k, '') or '').strip() for k in ('재질', '판정재질', '요구재질') if str(r.get(k, '') or '').strip()), '')
        item_or_line = str(r.get('라인번호', '') or '').strip() or str(r.get('검사품명', '') or '').strip()
        size_or_spec = str(r.get('관경', '') or '').strip() or material or str(r.get('치수원문', '') or '').strip()
        sec = str(r.get('구간', '') or '').strip() or target_type
        unit = str(r.get('단위', '') or '').strip()
        raw_qty = str(r.get('수량', '') or '').strip().replace(',', '')
        method_val = str(r.get(method, '') or '').strip().replace(',', '')
        qty_val = 0.0
        if method == 'RT':
            from site_apps.gyeongin.src.rt_quantities import rt_counts
            or_val, re_val = rt_counts(r)
            try:
                qty_val = float(or_val) if or_val else 0.0
            except ValueError:
                qty_val = 0.0
        elif method_val:
            try:
                qty_val = float(method_val)
            except ValueError:
                qty_val = 0.0
            unit = unit or 'm'
        elif raw_qty:
            try:
                qty_val = float(raw_qty)
            except ValueError:
                qty_val = 0.0
            unit = unit or 'EA'
        elif not unit:
            unit = 'm' if not is_non_pipe else 'EA'
        return {
            'is_non_pipe': is_non_pipe,
            'target_type': target_type,
            'sec': sec,
            'item_or_line': item_or_line,
            'size_or_spec': size_or_spec,
            'ident': ident,
            'material': material,
            'unit': unit,
            'qty': qty_val,
        }

    def _write_summary_sheet(self, ws, grouped_data):
        headers = ['업체', '대상/구간(Sec.No)', '품명/라인번호', '관경/재질·치수', '검사방법', '검사 횟수(건)', '물량 누계']
        col_widths = {'A': 16, 'B': 16, 'C': 24, 'D': 18, 'E': 11, 'F': 16, 'G': 22}
        
        for col_idx, width in enumerate(col_widths.values(), start=1):
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = width
            
        # Write Title
        ws.merge_cells('A1:G1')
        title_cell = ws.cell(row=1, column=1, value="용 역 명 : 경인사무소 비파괴검사용역")
        title_cell.font = Font(name='맑은 고딕', size=14, bold=True)
        title_cell.alignment = Alignment(horizontal='left', vertical='center')
        ws.row_dimensions[1].height = 25

        for col_idx, header_text in enumerate(headers, start=1):
            cell = ws.cell(row=2, column=col_idx, value=header_text)
            cell.font = self.font_bold
            cell.alignment = self.align_center
            cell.fill = self.fill_header
            cell.border = self.border_thin
            
        # Aggregate data
        summary_dict = {}
        for method, rows in grouped_data.items():
            for r in rows:
                company = str(r.get('업체', '')).strip()
                meta = self._row_meta(r, method)
                key = (company, meta['sec'], meta['item_or_line'], meta['size_or_spec'], method, meta['unit'])
                if key not in summary_dict:
                    summary_dict[key] = {'count': 0, 'qty': 0.0, 'rt_or': 0, 'rt_re': 0}
                    
                summary_dict[key]['count'] += 1
                
                if method == 'RT':
                    from site_apps.gyeongin.src.rt_quantities import rt_counts
                    or_val, re_val = rt_counts(r)
                    try: summary_dict[key]['rt_or'] += int(float(or_val)) if or_val else 0
                    except ValueError: pass
                    try: summary_dict[key]['rt_re'] += int(float(re_val)) if re_val else 0
                    except ValueError: pass
                    if meta['qty'] and not (or_val or re_val):
                        summary_dict[key]['qty'] += meta['qty']
                else:
                    summary_dict[key]['qty'] += meta['qty']
                    
        # Sort and Write
        sorted_keys = sorted(summary_dict.keys())
        row_idx = 3
        for key in sorted_keys:
            company, sec, line, size, method, unit = key
            stats = summary_dict[key]
            
            # 1~5: Group keys
            for i, val in enumerate((company, sec, line, size, method), start=1):
                cell = ws.cell(row=row_idx, column=i, value=val)
                cell.font = self.font_normal
                cell.alignment = self.align_center
                cell.border = self.border_thin
                
            # 6: Count
            cell = ws.cell(row=row_idx, column=6, value=f"{stats['count']} 건")
            cell.font = self.font_normal
            cell.alignment = self.align_center
            cell.border = self.border_thin
            
            # 7: Qty
            cell = ws.cell(row=row_idx, column=7)
            if method == 'RT' and (stats['rt_or'] or stats['rt_re'] or not stats['qty']):
                cell.value = f"OR: {stats['rt_or']} 매 / RE: {stats['rt_re']} 매"
            elif unit == 'm':
                cell.value = f"{stats['qty']:.4f} m"
            else:
                q_val = stats['qty']
                q_txt = str(int(q_val)) if float(q_val).is_integer() else f"{q_val:g}"
                cell.value = f"{q_txt} {unit}"
                
            cell.font = self.font_normal
            cell.alignment = self.align_center
            cell.border = self.border_thin
            
            row_idx += 1

    def _write_sheet(self, ws, method, rows):
        has_non_pipe = any(self._row_meta(r, method)['is_non_pipe'] for r in rows)
        if has_non_pipe:
            headers = ['일자', '순번', '업체', '대상유형', '검사품명', '식별번호', '재질', '치수원문', '결과', '적용규격']
            qty_headers = ['수량', '단위']
        else:
            headers = ['일자', '순번', '업체', '구간(Sec.No)', '라인번호', 'Joint No.', '관경', '용접사', '구간정보', '결과', '규격']
            qty_headers = ['RT_OR 전체(매)', 'RT_RE 불량(매)'] if method == 'RT' else ['물량(m)']
        all_headers = headers + qty_headers
        
        col_widths = {
            'A': 12, 'B': 6, 'C': 16, 'D': 14, 'E': 22, 'F': 16, 'G': 12, 'H': 22, 'I': 14, 'J': 14, 'K': 10, 'L': 10, 'M': 12
        }
        for col_idx, width in enumerate(list(col_widths.values())[:len(all_headers)], start=1):
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = width

        num_cols = len(all_headers)
        col_letter = get_column_letter(num_cols)
        ws.merge_cells(f'A1:{col_letter}1')
        title_cell = ws.cell(row=1, column=1, value="용 역 명 : 경인사무소 비파괴검사용역")
        title_cell.font = Font(name='맑은 고딕', size=14, bold=True)
        title_cell.alignment = Alignment(horizontal='left', vertical='center')
        ws.row_dimensions[1].height = 25

        for col_idx, header_text in enumerate(all_headers, start=1):
            cell = ws.cell(row=2, column=col_idx, value=header_text)
            cell.font = self.font_bold
            cell.alignment = self.align_center
            cell.fill = self.fill_header
            cell.border = self.border_thin

        for row_idx, row_data in enumerate(rows, start=3):
            col_idx = 1
            meta = self._row_meta(row_data, method)
            if has_non_pipe:
                row_values = [
                    row_data.get('일자', ''),
                    row_idx - 2,
                    row_data.get('업체', ''),
                    meta['target_type'],
                    row_data.get('검사품명', '') or row_data.get('라인번호', ''),
                    meta['ident'],
                    meta['material'],
                    row_data.get('치수원문', '') or row_data.get('구간정보', ''),
                    row_data.get('결과', ''),
                    row_data.get('적용규격', '') or row_data.get('규격', ''),
                ]
                for val in row_values:
                    cell = ws.cell(row=row_idx, column=col_idx, value=val)
                    cell.font = self.font_normal
                    cell.alignment = self.align_center
                    cell.border = self.border_thin
                    col_idx += 1
                for qv in (meta['qty'], meta['unit']):
                    cell = ws.cell(row=row_idx, column=col_idx, value=qv)
                    cell.font = self.font_normal
                    cell.alignment = self.align_center
                    cell.border = self.border_thin
                    col_idx += 1
                continue

            col_map = {
                '일자': '일자', '순번': '순번', '업체': '업체', '구간(Sec.No)': '구간',
                '라인번호': '라인번호', 'Joint No.': 'Joint No.', '관경': '관경',
                '용접사': '용접사', '구간정보': '구간정보', '결과': '결과', '규격': '규격'
            }
            for h in headers:
                val = (row_idx - 2) if h == '순번' else row_data.get(col_map.get(h, ''), '')
                cell = ws.cell(row=row_idx, column=col_idx, value=val)
                cell.font = self.font_normal
                cell.alignment = self.align_center
                cell.border = self.border_thin
                col_idx += 1
            if method == 'RT':
                for k in ['RT_OR', 'RT_RE']:
                    val = row_data.get(k, '')
                    try: val = float(val) if '.' in str(val) else int(val)
                    except ValueError: pass
                    cell = ws.cell(row=row_idx, column=col_idx, value=val)
                    cell.font = self.font_normal
                    cell.alignment = self.align_center
                    cell.border = self.border_thin
                    col_idx += 1
            else:
                val = row_data.get(method, '')
                try:
                    fval = float(str(val).replace(',', ''))
                    cell = ws.cell(row=row_idx, column=col_idx, value=fval)
                    cell.number_format = '0.0000'
                except ValueError:
                    cell = ws.cell(row=row_idx, column=col_idx, value=val)
                cell.font = self.font_normal
                cell.alignment = self.align_center
                cell.border = self.border_thin
                col_idx += 1
