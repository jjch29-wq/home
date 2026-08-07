import os

with open('ndt_summary_exporter.py', 'r', encoding='utf-8') as f:
    code = f.read()

old_logic = """        # If no data at all
        if not grouped_data:
            ws = wb.active
            ws.title = "데이터 없음"
            ws['A1'] = "출력할 비파괴검사 결과가 없습니다."
            wb.save(output_path)
            return

        # Create sheets for each method
        is_first = True
        for method, rows in grouped_data.items():
            if is_first:
                ws = wb.active
                ws.title = method
                is_first = False
            else:
                ws = wb.create_sheet(title=method)
                
            self._write_sheet(ws, method, rows)
            
        wb.save(output_path)"""

new_logic = """        # If no data at all
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
        
    def _write_summary_sheet(self, ws, grouped_data):
        headers = ['구간(Sec.No)', '라인번호', '관경', '검사방법', '검사 횟수(Joint 카운트)', '물량 누계']
        col_widths = {'A': 15, 'B': 25, 'C': 12, 'D': 12, 'E': 22, 'F': 25}
        
        for col_idx, width in enumerate(col_widths.values(), start=1):
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = width
            
        for col_idx, header_text in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_idx, value=header_text)
            cell.font = self.font_bold
            cell.alignment = self.align_center
            cell.fill = self.fill_header
            cell.border = self.border_thin
            
        # Aggregate data
        summary_dict = {}
        for method, rows in grouped_data.items():
            for r in rows:
                sec = str(r.get('구간', '')).strip()
                line = str(r.get('라인번호', '')).strip()
                size = str(r.get('관경', '')).strip()
                
                key = (sec, line, size, method)
                if key not in summary_dict:
                    summary_dict[key] = {'count': 0, 'm': 0.0, 'rt_or': 0, 'rt_re': 0}
                    
                summary_dict[key]['count'] += 1
                
                if method == 'RT':
                    or_val = str(r.get('RT_OR', '')).strip()
                    re_val = str(r.get('RT_RE', '')).strip()
                    try: summary_dict[key]['rt_or'] += int(float(or_val)) if or_val else 0
                    except ValueError: pass
                    try: summary_dict[key]['rt_re'] += int(float(re_val)) if re_val else 0
                    except ValueError: pass
                else:
                    m_val = str(r.get(method, '')).strip().replace(',','')
                    try: summary_dict[key]['m'] += float(m_val) if m_val else 0.0
                    except ValueError: pass
                    
        # Sort and Write
        sorted_keys = sorted(summary_dict.keys())
        row_idx = 2
        for key in sorted_keys:
            sec, line, size, method = key
            stats = summary_dict[key]
            
            # 1~4: Group keys
            for i, val in enumerate(key, start=1):
                cell = ws.cell(row=row_idx, column=i, value=val)
                cell.font = self.font_normal
                cell.alignment = self.align_center
                cell.border = self.border_thin
                
            # 5: Count
            cell = ws.cell(row=row_idx, column=5, value=f"{stats['count']} 개")
            cell.font = self.font_normal
            cell.alignment = self.align_center
            cell.border = self.border_thin
            
            # 6: Qty
            cell = ws.cell(row=row_idx, column=6)
            if method == 'RT':
                cell.value = f"OR: {stats['rt_or']} 매 / RE: {stats['rt_re']} 매"
            else:
                cell.value = f"{stats['m']:.4f} m"
                
            cell.font = self.font_normal
            cell.alignment = self.align_center
            cell.border = self.border_thin
            
            row_idx += 1"""

code = code.replace(old_logic, new_logic)

with open('ndt_summary_exporter.py', 'w', encoding='utf-8') as f:
    f.write(code)
print("Updated NDT summary exporter successfully")
