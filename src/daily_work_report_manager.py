import openpyxl
from openpyxl.styles import Alignment, Border, Side, Font
from openpyxl.cell.cell import MergedCell
import os
import datetime

class DailyWorkReportManager:
    def __init__(self, template_path):
        self.template_path = template_path
        # Default mapping for fallback
        self.default_mapping = {
            'header': {'date': 'F2'},
            'general': {
                'company': 'E5', 'project_name': 'E6', 'standard': 'E7', 'equipment': 'E8',
                'report_no': 'K5', 'inspection_item': 'K6', 'inspector': 'K7', 'inspector_n8': 'N8', 'car_no': 'N9'
            },
            'methods': {
                'RT': {'row': 13}, 'UT': {'row': 14}, 'MT': {'row': 15}, 'PT': {'row': 16},
                'HT': {'row': 17}, 'VT': {'row': 18}, 'LT': {'row': 19}, 'ET': {'row': 20},
                'PAUT': {'row': 21}
            },
            'rtk': {
                'center_miss': 'D34', 'density': 'F34', 'marking_miss': 'H34', 'film_mark': 'J34',
                'handling': 'L34', 'customer_complaint': 'N34', 'etc': 'P34', 'total': 'R34'
            },
            'ot': {
                'row1_name': 'B38', 'row1_company': 'F38', 'row1_method': 'I38', 'row1_hours': 'K38', 'row1_amount': 'N38',
                'row2_name': 'B39', 'row2_company': 'F39', 'row2_method': 'I39', 'row2_hours': 'K39', 'row2_amount': 'N39'
            },
            'materials': {
                'RT T200': 43, 'RT AA400': 44, 'RT Other': 45,
                'MT WHITE': 46, 'MT 7C-BLACK': 47,
                'PT Penetrant': 48, 'PT Cleaner': 49, 'PT Developer': 50
            },
            'vehicles': {
                'row_start': 26, 
                'col_map': {'vehicle_info': 'B', 'mileage': 'H', 'fuel': 'C', 'clean': 'E', 'oil': 'G', 'tire': 'I', 'light': 'K', 'safety': 'M'}
            }
        }

    def generate_report(self, data, output_path, custom_mapping=None):
        if not os.path.exists(self.template_path):
            raise FileNotFoundError(f"Template not found at {self.template_path}")

        wb = openpyxl.load_workbook(self.template_path)
        sheet = wb.active
        mapping = custom_mapping if custom_mapping else self.default_mapping

        def safe_write(cell_coord, value, is_currency=False, is_bold=False):
            if not cell_coord: return
            try:
                cell = sheet[cell_coord]
                if isinstance(cell, MergedCell):
                    for m_range in sheet.merged_cells.ranges:
                        if cell_coord in m_range:
                            cell = sheet.cell(row=m_range.min_row, column=m_range.min_col)
                            break
                cell.value = value
                cell.font = Font(name='맑은 고딕', size=9, bold=is_bold)
                # [FIX] shrinkToFit and wrapText are mutually exclusive in Excel - use wrapText only
                cell.alignment = Alignment(wrapText=True, vertical='center', horizontal='center')
                if is_currency:
                    cell.number_format = '#,##0 "원"'
            except: pass

        # 1. Header & General
        d = data.get('date', datetime.date.today())
        weekdays = ["월", "화", "수", "목", "금", "토", "일"]
        date_str = f"{d.year}년 {d.month}월 {d.day}일 ({weekdays[d.weekday()]})"
        safe_write(mapping.get('header', {}).get('date', 'F2'), date_str)

        gen_map = mapping.get('general', {})
        for key in ['company', 'project_name', 'standard', 'equipment', 'report_no', 'inspection_item', 'inspector', 'car_no', 'inspector_n8']:
            val_key = 'inspector' if key == 'inspector_n8' else key
            safe_write(gen_map.get(key), data.get(val_key, ''))

        # 2. Methods (Dynamic Offset)
        methods_data = data.get('methods', {})
        current_methods = list(methods_data.keys())
        method_offset = max(0, len(current_methods) - 4)
        
        import re as _re
        def fix_c(coord, offset):
            if not coord or not isinstance(coord, str): return coord
            match = _re.match(r"([A-Z]+)([0-9]+)", coord)
            if match:
                col, row = match.groups()
                row_val = int(row)
                if row_val >= 18: return f"{col}{row_val + offset}"
            return coord

        method_base_rows = [13, 14, 15, 16]
        if method_offset > 0:
            sheet.insert_rows(17, method_offset)
            for i in range(method_offset):
                target_r = 18 + method_offset + i
                if target_r <= 25 + method_offset:
                    sheet.row_dimensions[target_r].height = 5
            method_base_rows += list(range(17, 17 + method_offset))

        for idx, m_name in enumerate(current_methods):
            if idx < len(method_base_rows):
                r = method_base_rows[idx]; m = methods_data[m_name]
                safe_write(f'B{r}', m_name); safe_write(f'E{r}', m.get('unit', ''))
                safe_write(f'H{r}', m.get('qty', 0)); safe_write(f'K{r}', m.get('price', 0), is_currency=True)
                safe_write(f'N{r}', m.get('travel', 0), is_currency=True); safe_write(f'Q{r}', m.get('total', 0), is_currency=True)

        # 3. Section 2 (Notes) - Styling & Normalize
        # [NEW] Sync Row 17 height with Row 10 height for consistency
        h10 = sheet.row_dimensions[10].height if sheet.row_dimensions[10].height else 15
        sheet.row_dimensions[17].height = h10
        for m_range in list(sheet.merged_cells.ranges):
            if 18 + method_offset <= m_range.min_row and m_range.max_row <= 25 + method_offset:
                try: sheet.unmerge_cells(m_range.coord)
                except: pass
        try: sheet.merge_cells(f"B{18+method_offset}:S{25+method_offset}")
        except: pass
        
        # 4. Section 4 (RTK)
        rtk_data = data.get('rtk_quality', {}); rtk_map = mapping.get('rtk', {})
        rtk_key_map = {'center_miss':'센터미스','density':'농도','marking_miss':'마킹미스','film_mark':'필름마크','handling':'취급부주의','customer_complaint':'고객불만','etc':'기타'}
        for eng, kor in rtk_key_map.items():
            safe_write(fix_c(rtk_map.get(eng), method_offset), rtk_data.get(kor, 0))
        total_rtk = rtk_data.get('총계', sum([v for v in rtk_data.values() if isinstance(v, (int, float))]))
        safe_write(fix_c(rtk_map.get('total'), method_offset), total_rtk)

        # 5. Section 5 (OT)
        for r_search in range(1, 100):
            if sheet.cell(row=r_search, column=2).value == "검사자":
                ot_header_row = r_search; break
        else: ot_header_row = 38 + method_offset
        ot_list = data.get('ot_status', [])
        for i, ot in enumerate(ot_list[:2]):
            r = ot_header_row + 1 + i
            safe_write(f"B{r}", ot.get('names', '')); safe_write(f"F{r}", ot.get('company', ''))
            safe_write(f"I{r}", ot.get('method', '')); safe_write(f"K{r}", ot.get('ot_hours', ''))
            safe_write(f"N{r}", ot.get('ot_amount', ''), is_currency=True)

        # 6. Section 6 (Materials)
        materials_data = data.get('materials', {}); active_rt = []
        for k, v in materials_data.items():
            if isinstance(v, dict) and (k.startswith('RT ') or k.startswith('RT_ROW_') or v.get('is_rt')):
                item = v.copy(); item['name'] = item.get('name', k); active_rt.append(item)
        
        rt_extra = max(0, len(active_rt) - 3); total_offset = method_offset + rt_extra
        # [OPTIMIZED LAYOUT] Fill the A4 page and fix Category Merges
        total_extra_rows = method_offset + rt_extra + 1
        
        # Increase Section 2 (Notes) height to absorb bottom white space
        # [FIXED GRID MODE] No insertions or deletions to protect Row 42 Title and Row 51 End
        # 1. Category Merges (B:C / 2:3) - STRICTLY FIXED
        def safe_merge(sheet, s_row, s_col, e_row, e_col, value):
            for m_range in list(sheet.merged_cells.ranges):
                if not (e_row < m_range.min_row or s_row > m_range.max_row or 
                        e_col < m_range.min_col or s_col > m_range.max_col):
                    try: sheet.unmerge_cells(m_range.coord)
                    except: pass
            try:
                sheet.merge_cells(start_row=s_row, start_column=s_col, end_row=e_row, end_column=e_col)
                cell = sheet.cell(row=s_row, column=s_col)
                cell.value = value
                cell.alignment = Alignment(horizontal='center', vertical='center')
            except: pass

        safe_merge(sheet, 43, 2, 46, 3, "RT") 
        safe_merge(sheet, 47, 2, 48, 3, "MT")
        safe_merge(sheet, 49, 2, 51, 3, "PT")

        # [FIX] Set column widths for D and E so merged D:E cell is wide enough for long names
        sheet.column_dimensions['D'].width = 12
        sheet.column_dimensions['E'].width = 12

        # 2. Item Name (D:E) and Spec (F:G) Horizontal Merges
        for r in range(43, 52):
            try:
                # Clean up existing merges in D:G range first
                for m_range in list(sheet.merged_cells.ranges):
                    if m_range.min_row == r and m_range.max_row == r and (m_range.min_col >= 4 and m_range.max_col <= 7):
                        sheet.unmerge_cells(m_range.coord)
                
                if r <= 46:
                    # RT Area: Separate Name (D:E) and Spec (F:G)
                    sheet.merge_cells(start_row=r, start_column=4, end_row=r, end_column=5)
                    sheet.merge_cells(start_row=r, start_column=6, end_row=r, end_column=7)
                else:
                    # MT/PT Area: Wide Combined Name/Spec (D:G)
                    sheet.merge_cells(start_row=r, start_column=4, end_row=r, end_column=7)
                    
                # [SPECIFIC ROW 51 MERGE] I:J, K:L, M:N, O:P, Q:S
                if r == 51:
                    for s_c, e_c in [(9,10), (11,12), (13,14), (15,16), (17,19)]:
                        try:
                            # Unmerge first
                            for m_range in list(sheet.merged_cells.ranges):
                                if m_range.min_row == 51 and m_range.max_row == 51 and m_range.min_col == s_c:
                                    sheet.unmerge_cells(m_range.coord)
                            sheet.merge_cells(start_row=51, start_column=s_c, end_row=51, end_column=e_c)
                        except: pass
            except: pass

        # [DYNAMIC STYLE COPY & DATA PROTECTION]
        # 1. Clear the RT Range (43-46) first to remove template remnants
        mat_start = 43
        for r in range(mat_start, 47):
            for col in range(4, 20): # D to S
                cell = sheet.cell(row=r, column=col)
                if not isinstance(cell, MergedCell):
                    cell.value = None

        # Helper: Remove known brand prefixes for display
        def strip_brand(name):
            for prefix in ['Carestream ', 'AGFA ', 'Fuji ', 'Kodak ']:
                if name.startswith(prefix):
                    return name[len(prefix):]
            return name

        # Write RT Data with Standardized Formatting
        for idx, m in enumerate(active_rt):
            r = mat_start + idx
            if r > 46: break
            
            # Strip brand prefix so short name fits in merged cell
            display_name = strip_brand(m.get('name', ''))
            safe_write(f'D{r}', display_name)
            # Apply style AFTER safe_write
            cell_name = sheet.cell(row=r, column=4)
            cell_name.font = Font(name='맑은 고딕', size=9)
            cell_name.alignment = Alignment(horizontal='center', vertical='center', wrapText=True)
            sheet.row_dimensions[r].height = 30
            
            safe_write(f'F{r}', m.get('spec', ''))
            sheet.cell(row=r, column=6).alignment = Alignment(horizontal='center', vertical='center', wrapText=True)
            safe_write(f'H{r}', m.get('unit', '매'))
            
            # Correct Mapping: K=Out, M=Used, O=In (Q is intentionally left blank)
            safe_write(f'K{r}', '-')
            used_val = int(m.get('used', 0))
            in_val = int(m.get('in', 0))
            safe_write(f'M{r}', f"{used_val} 매") # Usage is M
            safe_write(f'O{r}', f"{in_val} 매")   # Incoming is O
            # Stock (Q) removed per request

        # 3. Add same placeholders to empty RT rows (ensure 43-46 are uniform)
        for r in range(43, 47):
            if r >= mat_start + len(active_rt):
                safe_write(f'K{r}', '-'); safe_write(f'M{r}', '매'); safe_write(f'O{r}', '0 매')
            
        synonyms = {
            'MT WHITE': ['MT WHITE', '백색', '백색자분', 'WHITE'],
            'MT 7C-BLACK': ['MT 7C-BLACK', '7C', '자분', 'BLACK', '7C-BLACK'],
            'PT Penetrant': ['PT Penetrant', '침투', '침투액', 'Penetrant'],
            'PT Cleaner': ['PT Cleaner', '세척', '세척액', 'Cleaner'],
            'PT Developer': ['PT Developer', '현상', '현상액', '현상제', 'Developer']
        }
        # Fixed Map matching RT: 43-46, MT: 47-48, PT: 49-51
        mat_map = {'MT WHITE': 47, 'MT 7C-BLACK': 48, 'PT Penetrant': 49, 'PT Cleaner': 50, 'PT Developer': 51}
        
        # Korean Display Names for MT/PT
        display_names = {
            'MT WHITE': '백색페인트', 'MT 7C-BLACK': '흑색자분', 
            'PT Penetrant': '침투액', 'PT Cleaner': '세척액', 'PT Developer': '현상액'
        }
        
        for m_key, r in mat_map.items():
            # ALWAYS write the chemical name to D (merged with E, F, G)
            safe_write(f'D{r}', display_names.get(m_key, m_key)) 
            
            # ALWAYS write the unit to H (All MT/PT use CAN)
            safe_write(f'H{r}', "CAN")
            
            m_d = {}
            for syn in synonyms.get(m_key, []):
                m_d = materials_data.get(syn)
                if m_d: break
            
            if m_d:
                # K=Out, M=Used, O=In (Q is intentionally left blank)
                safe_write(f'K{r}', '-'); safe_write(f'M{r}', int(m_d.get('used', 0)))
                safe_write(f'O{r}', int(m_d.get('in', 0)))

        # [BORDER & ALIGNMENT] Apply THIN border to B42:S51
        thin = Side(style='thin')
        for r in range(42, 52):
            for col in range(2, 20): # B(2) to S(19)
                cell = sheet.cell(row=r, column=col)
                cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
                cell.alignment = Alignment(horizontal='center', vertical='center')

        # 7. Section 3 (Vehicles) & B27 Cleanup
        veh_list = data.get('vehicles', []); v = veh_list[0] if veh_list else {}
        veh_row = 27 + method_offset
        safe_write(f"B{veh_row}", ""); safe_write(f"H{veh_row}", v.get('mileage', ''))
        
        chk_rows = {'out': 29 + total_offset, 'in': 30 + total_offset}
        for rs in range(25+total_offset, 35+total_offset):
            val = sheet.cell(row=rs, column=2).value
            if val == "출차시": chk_rows['out'] = rs
            if val == "입차시": chk_rows['in'] = rs
        
        chk_map = {'exterior':'E','cleanliness':'H','cleaning':'K','locking':'N'}
        WHITE_SQ, BLACK_SQ = "\u25a1", "\u25a0"
        import re
        for rk, row_idx in chk_rows.items():
            for ck, col_let in chk_map.items():
                val = v.get(f"{rk}_{ck}")
                if val:
                    coord = f"{col_let}{row_idx}"; cell = sheet[coord]
                    if isinstance(cell, MergedCell):
                        for mr in sheet.merged_cells.ranges:
                            if coord in mr: cell = sheet.cell(row=mr.min_row, column=mr.min_col); break
                    if cell.value and isinstance(cell.value, str):
                        pattern = f"({WHITE_SQ})(\\s*){re.escape(val)}"
                        if re.search(pattern, cell.value):
                            cell.value = re.sub(pattern, f"{BLACK_SQ}\\2{val}", cell.value)

        sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
        sheet.page_setup.horizontalCentered = True
        sheet.page_setup.verticalCentered = False # Top-aligned for precision
        
        # Exact Margins (Total 0.8" vertical consumption = 57.6pt)
        sheet.page_margins.top = 0.4; sheet.page_margins.bottom = 0.4
        sheet.page_margins.left = 0.3; sheet.page_margins.right = 0.3
        
        # A4 Usable Height at 100% scale is approx 780pt.
        # Let's allocate heights precisely:
        # 1. Fixed Header (1-11): 11 rows * 22pt = 242pt
        for r in range(1, 12): sheet.row_dimensions[r].height = 22
        
        # 2. Title Rows (12, 26, 31, 37, 42): 5 rows * 26pt = 130pt
        # 2. Title Rows (12, 26, 31, 37, 42)
        fixed_titles = [12, 26, 31, 37, 42]
        
        # [ROW 51 FIT OPTIMIZATION]
        # 1. Global Row Height Baseline (Aggressive Slimming)
        for r in range(1, 101):
            try: sheet.row_dimensions[r].height = 14 # Slimmer baseline
            except: pass
        
        # 2. Precision Column Widths (Remain slim)
        slim_widths = {
            'A': 1, 'B': 7, 'C': 1, 'D': 9, 'E': 1, 'F': 9, 'G': 1, 'H': 5, 
            'I': 4, 'J': 4, 'K': 5, 'L': 4, 'M': 5, 'N': 4, 'O': 5, 
            'P': 4, 'Q': 5, 'R': 4, 'S': 10
        }
        for col, width in slim_widths.items():
            sheet.column_dimensions[col].width = width

        # 3. Precision Row Heights (Recalibrated for Row 51 fit)
        # Header (1-11): 11 * 15 = 165pt
        for r in range(1, 12): sheet.row_dimensions[r].height = 15
        # Titles: 5 * 20 = 100pt
        for t_row in [12, 26, 31, 37, 42]:
            try: sheet.row_dimensions[t_row + method_offset].height = 20
            except: pass
        # Section 2 (Notes): 8 * 25 = 200pt
        for b_row in range(18+method_offset, 26+method_offset):
            sheet.row_dimensions[b_row].height = 25
        # Materials: ~9 * 16 = 144pt
        for m_row in range(43 + method_offset, 52 + total_offset):
            sheet.row_dimensions[m_row].height = 16
            
        # 8. Final Print Setup - FIXED 95% SCALE & TIGHT MARGINS
        sheet.sheet_properties.pageSetUpPr.fitToPage = False
        sheet.page_setup.scale = 95
        sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
        sheet.page_setup.horizontalCentered = True
        sheet.page_setup.verticalCentered = True
        
        # Symmetrical Margins with MINIMIZED BOTTOM for extra space
        sheet.page_margins.left = 0.5; sheet.page_margins.right = 0.5
        sheet.page_margins.top = 0.4; sheet.page_margins.bottom = 0.2

        sheet.print_area = f"A1:S{51 + method_offset + rt_extra}"
        
        wb.save(output_path); return output_path
