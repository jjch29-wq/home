import os
import openpyxl
from openpyxl.styles import Border, Side, Alignment, Font
import pandas as pd
import re

class DailyWorkReportManager:
    def __init__(self, template_path):
        self.template_path = template_path

    def generate_report(self, data, output_path, custom_mapping=None):
        wb = openpyxl.load_workbook(self.template_path)
        sheet = wb.active
        cmap = custom_mapping if custom_mapping else {}

        def safe_write(coord, value):
            try: sheet[coord] = value
            except: pass

        def safe_merge(range_str):
            try: sheet.merge_cells(range_str)
            except: pass

        def find_title_row(title_text):
            for row in range(1, 100):
                val = sheet.cell(row=row, column=1).value or sheet.cell(row=row, column=2).value
                if val and str(val).strip().startswith(title_text):
                    return row
            return None

        # 1. Date and General Info (Existing)
        report_date = data.get('date', '2026-00-00')
        safe_write('D4', report_date)

        # 2. Section 3: Vehicle Management (Existing but need to ensure it's correct)
        # (Assuming data['vehicles'] is handled in the main script or here)
        
        # [NEW] Section 4, 5, 6 titles and logic
        s3_title = find_title_row("3. ")
        s4_title = find_title_row("4. ")
        s5_title = find_title_row("5. ")
        s6_title = find_title_row("6. ")

        # 3. Section 6: Material Inventory (Dynamic)
        m_data = data.get('materials', {})
        rt_items = [v for k, v in m_data.items() if v.get('is_rt')]
        # Filter only active items
        rt_items = [m for m in rt_items if float(m.get('used', 0)) > 0 or float(m.get('in', 0)) > 0]
        
        num_rt = len(rt_items)
        rt_height = max(3, num_rt)
        extra_rows = rt_height - 3
        
        mat_start = s6_title + 2 if s6_title else 40
        
        if extra_rows > 0:
            sheet.insert_rows(mat_start, extra_rows)
            # Compensation
            buffer_start = 18
            try:
                rows_to_delete = min(extra_rows, 8)
                sheet.delete_rows(buffer_start, rows_to_delete)
                mat_start -= rows_to_delete
            except: pass

        # Clear area
        has_mt = any('MT' in k.upper() for k in m_data)
        has_pt = any('PT' in k.upper() for k in m_data)
        total_rows = rt_height + (2 if has_mt else 0) + (3 if has_pt else 0)
        for r_cl in range(mat_start, mat_start + total_rows + 5):
            for c_cl in ['D', 'F', 'H', 'I', 'K', 'M', 'O', 'Q']:
                safe_write(f"{c_cl}{r_cl}", "")

        # Write RT Items
        for idx, m in enumerate(rt_items):
            r = mat_start + idx
            safe_write(f'D{r}', m.get('name', ''))
            safe_write(f'F{r}', m.get('spec', ''))
            safe_write(f'H{r}', '매')
            safe_write(f'K{r}', m.get('in', 0))
            safe_write(f'M{r}', m.get('used', 0))
            safe_write(f'O{r}', m.get('stock', 0))
            safe_write(f'I{r}', m.get('remarks', ''))
            # Merges for Name/Spec
            safe_merge(f"D{r}:E{r}")
            safe_merge(f"F{r}:G{r}")
            safe_merge(f"K{r}:L{r}")
            safe_merge(f"M{r}:N{r}")
            safe_merge(f"O{r}:P{r}")
            safe_merge(f"Q{r}:S{r}")

        # TYPE Merge
        safe_merge(f'B{mat_start}:C{mat_start + rt_height - 1}')
        safe_write(f'B{mat_start}', 'RT 필름')

        # MT/PT logic (Simplified for restoration)
        curr_r = mat_start + rt_height
        if has_mt:
            safe_merge(f'B{curr_r}:C{curr_r + 1}'); safe_write(f'B{curr_r}', 'MT 시약')
            curr_r += 2
        if has_pt:
            safe_merge(f'B{curr_r}:C{curr_r + 2}'); safe_write(f'B{curr_r}', 'PT 시약')
            curr_r += 3
        
        mat_end = curr_r - 1

        # 4. Styling System
        thin_side = Side(style='thin')
        hair_side = Side(style='hair')
        
        def apply_section_styling(start_r, end_r):
            for r in range(start_r, end_r + 1):
                for c in range(2, 20): # B to S
                    cell = sheet.cell(row=r, column=c)
                    l_s = thin_side if c == 2 else hair_side
                    r_s = thin_side if c == 19 else hair_side
                    t_s = hair_side
                    b_s = hair_side
                    
                    # Top border for header
                    if r == start_r: t_s = thin_side
                    # Bottom border for end of section
                    if r == end_r: b_s = thin_side
                    
                    cell.border = Border(left=l_s, right=r_s, top=t_s, bottom=b_s)
                    cell.alignment = Alignment(horizontal='center', vertical='center')

        # Apply styles to sections
        if s3_title: apply_section_styling(s3_title + 1, s3_title + 4)
        if s4_title: apply_section_styling(s4_title + 1, s4_title + 2)
        if s5_title: apply_section_styling(s5_title + 1, s5_title + 3)
        if s6_title: 
            apply_section_styling(s6_title + 1, mat_end)
            # Final touch: Row 45 Hairline bottom
            for c in range(2, 20):
                if sheet.cell(row=45, column=c).border.bottom == thin_side:
                    sheet.cell(row=45, column=c).border = Border(
                        left=sheet.cell(row=45, column=c).border.left,
                        right=sheet.cell(row=45, column=c).border.right,
                        top=sheet.cell(row=45, column=c).border.top,
                        bottom=hair_side
                    )

        # 5. Visibility and Open Look
        open_rows = [5, 10, 26, 35, 40, 51]
        for r_o in range(5, max(52, sheet.max_row + 1)):
            if r_o in open_rows:
                b_cell = sheet.cell(row=r_o, column=2)
                s_cell = sheet.cell(row=r_o, column=19)
                b_cell.border = Border(left=None, right=b_cell.border.right, top=b_cell.border.top, bottom=b_cell.border.bottom)
                s_cell.border = Border(left=s_cell.border.left, right=None, top=s_cell.border.top, bottom=s_cell.border.bottom)

        # Shrink to Fit
        for r_f in range(38, max(51, mat_end + 1)):
            for c_f in [4, 6, 11, 12, 13]:
                if r_f > 39 and c_f > 6: continue
                sheet.cell(row=r_f, column=c_f).alignment = Alignment(horizontal='center', vertical='center', shrink_to_fit=True)

        wb.save(output_path)
        return output_path
