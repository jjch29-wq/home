import openpyxl
from openpyxl.styles import Alignment, Border, Side, Font
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
                'report_no': 'K5', 'inspection_item': 'K6', 'inspector': 'K7', 'car_no': 'N9'
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
                'RT T200': 43, 'RT AA400': 44, 'MT WHITE': 46, 'MT 7C-BLACK': 47,
                'PT Penetrant': 48, 'PT Cleaner': 49, 'PT Developer': 50
            },
            'vehicles': {
                'row_start': 22,
                'col_map': {'vehicle_info': 'B', 'mileage': 'H', 'fuel': 'C', 'clean': 'E', 'oil': 'G', 'tire': 'I', 'light': 'K', 'safety': 'M'}
            }
        }

    def generate_report(self, data, output_path, custom_mapping=None):
        if not os.path.exists(self.template_path):
            raise FileNotFoundError(f"Template not found at {self.template_path}")

        wb = openpyxl.load_workbook(self.template_path)
        sheet = wb.active
        
        mapping = custom_mapping if custom_mapping else self.default_mapping

        def safe_write(cell_coord, value):
            if not cell_coord: return
            try:
                sheet[cell_coord].value = value
            except:
                pass

        # 1. Header (Date)
        d = data.get('date', datetime.date.today())
        weekdays = ["월", "화", "수", "목", "금", "토", "일"]
        date_str = f"{d.year}년 {d.month}월 {d.day}일 ({weekdays[d.weekday()]})"
        safe_write(mapping.get('header', {}).get('date', 'F2'), date_str)

        # 2. General
        gen_map = mapping.get('general', {})
        for key in ['company', 'project_name', 'standard', 'equipment', 'report_no', 'inspection_item', 'inspector', 'car_no']:
            safe_write(gen_map.get(key), data.get(key, ''))

        # 3. Methods (Dynamic Population)
        method_map = mapping.get('methods', {})
        methods_data = data.get('methods', {})
        
        def extract_row_num(v):
            if not v: return None
            import re
            m = re.search(r'\d+', str(v))
            return int(m.group()) if m else None

        # Collect and sort all possible rows defined for methods to clear them first
        method_rows = []
        for m in method_map.values():
            r = extract_row_num(m.get('row'))
            if r is not None:
                method_rows.append(r)
        method_rows = sorted(list(set(method_rows)))

        for row in method_rows:
            for col in ['A', 'B', 'E', 'H', 'K', 'N', 'Q']: # Clear Name, Unit, Qty, Price, Total, Travel, Fee
                safe_write(f'{col}{row}', '')

        # Fill with active methods sequentially
        if method_rows:
            current_row_idx = 0
            # Ensure we only write methods that have data
            for m_name in methods_data.keys():
                if current_row_idx >= len(method_rows): break # Out of space in template
                
                row = method_rows[current_row_idx]
                m_data = methods_data[m_name]
                
                safe_write(f'B{row}', m_name)               # Method Name to Col B
                safe_write(f'E{row}', m_data.get('unit', '')) # Unit to Col E
                safe_write(f'H{row}', m_data.get('qty', 0))   # Qty to Col H
                safe_write(f'K{row}', m_data.get('price', 0)) # Unit Price to Col K
                safe_write(f'N{row}', m_data.get('travel', 0)) # Travel Cost to Col N
                safe_write(f'Q{row}', m_data.get('total', 0)) # Test Fee to Col Q
                
                current_row_idx += 1

        # 4. RTK
        rtk_map = mapping.get('rtk', {})
        rtk_data = data.get('rtk', {})
        
        # [NEW] Re-map English config keys to Korean data keys
        rtk_key_map = {
            'center_miss': '센터미스', 'density': '농도', 'marking_miss': '마킹미스', 
            'film_mark': '필름마크', 'handling': '취급부주의', 'customer_complaint': '고객불만', 'etc': '기타'
        }
        
        for eng_key, kor_key in rtk_key_map.items():
            safe_write(rtk_map.get(eng_key), rtk_data.get(kor_key, 0))
        
        total_rtk = rtk_data.get('총계', sum([v for v in rtk_data.values() if isinstance(v, (int, float))]))
        safe_write(rtk_map.get('total'), total_rtk)


        # 5. OT
        ot_map = mapping.get('ot', {})
        ot_list = data.get('ot_status', [])
        for i, ot in enumerate(ot_list[:2]):
            idx = i + 1
            safe_write(ot_map.get(f'row{idx}_name'), ot.get('names', ''))
            safe_write(ot_map.get(f'row{idx}_company'), ot.get('company', ''))
            safe_write(ot_map.get(f'row{idx}_method'), ot.get('method', ''))
            safe_write(ot_map.get(f'row{idx}_hours'), str(ot.get('ot_hours', '')))

            safe_write(ot_map.get(f'row{idx}_amount'), ot.get('ot_amount', ''))

        # [NEW] Style the boundary between Row 38 and Row 39 with hair (hairline) lines
        hair_side = Side(style='hair')
        thin_side = Side(style='thin')
        for col_name in ['K', 'L', 'M', 'N', 'O', 'P']:
            cell_addr = f"{col_name}38"
            current_border = sheet[cell_addr].border
            sheet[cell_addr].border = Border(
                left=current_border.left,
                right=current_border.right,
                top=current_border.top,
                bottom=hair_side
            )

        # [NEW] Specific Header Borders
        # E8 Left: Hair
        e8_border = sheet['E8'].border
        sheet['E8'].border = Border(left=hair_side, right=e8_border.right, top=e8_border.top, bottom=e8_border.bottom)
        
        # N6~N9 Left: Thin
        for r_idx in range(6, 10):
            cell_n = f"N{r_idx}"
            cur_b = sheet[cell_n].border
            sheet[cell_n].border = Border(left=thin_side, right=cur_b.right, top=cur_b.top, bottom=cur_b.bottom)


        # [NEW] Clear template placeholder in Q39 (which usually contains '숙박')
        safe_write('Q39', '')








        # 6. Materials
        mat_map = mapping.get('materials', {})
        materials_data = data.get('materials', {})
        
        # [NEW] Multi-Item Support for RT Films (Rows 43, 44, 45)
        # Handle comma-separated list in selected_material
        import re
        sel_mat = data.get('selected_material', '')
        rt_items = []
        if sel_mat:
            # Handle multiple entries separated by comma or semicolon
            rt_items = [s.strip() for s in re.split(r'[,;]', sel_mat) if s.strip()]
            
        # 3. Write up to 3 items to rows 43-45
        for i, item in enumerate(rt_items[:3]):
            row_idx = 43 + i
            if '-' in item:
                parts = item.split('-', 1)
                p1 = re.sub(r'(?i)Carestream\s*', '', parts[0]).strip()
                p2 = parts[1].strip()
            else:
                p1 = re.sub(r'(?i)Carestream\s*', '', item).strip()
                p2 = ""
                
            safe_write(f'D{row_idx}', p1)
            if p2: safe_write(f'F{row_idx}', p2)

        # [NEW] Clear template placeholders for unused rows (if fewer than 3 items)
        for i in range(len(rt_items), 3):
            row_idx = 43 + i
            safe_write(f'D{row_idx}', '')
            safe_write(f'F{row_idx}', '')
            
        # [USER_REQ] Always put '매' in H45 (or ensure it's there)
        safe_write('H45', '매')

        # [NEW] Clear range K43:P50 as requested
        for r in range(43, 51):
            for c_idx in range(11, 17): # K(11) to P(16)
                safe_write(f"{chr(64+c_idx)}{r}", "")





        for mat_name, row in mat_map.items():
            if mat_name in materials_data:

                m_data = materials_data[mat_name]
                safe_write(f'K{row}', m_data.get('init', 0))
                safe_write(f'M{row}', m_data.get('used', 0))
                safe_write(f'O{row}', m_data.get('in', 0))
        
        # 7. Vehicles & Safety (Section 3)
        # Detailed 2x4 Table Implementation (출차/입차)
        veh_map = mapping.get('vehicles', {})
        veh_list = data.get('vehicles', [])
        v = veh_list[0] if veh_list else {}
        veh_row = veh_map.get('row_start', 22) # Vehicle Info Row
        cmap = veh_map.get('col_map', {})
        
        safe_write(f"{cmap.get('vehicle_info')}{veh_row}", v.get('vehicle_info', ''))
        safe_write(f"{cmap.get('mileage')}{veh_row}", v.get('mileage', ''))

        # Detailed Inspection Checklist Mapping
        # Row 29: 출차시 / Row 30: 입차시
        chk_rows = {'out': 29, 'in': 30}
        chk_cols = {
            'exterior': {'양호': 'E', '불량': 'F'},
            'cleanliness': {'양호': 'H', '불량': 'I'},
            'cleaning': {'함': 'K', '안함': 'L'},
            'locking': {'잠금': 'N', '안함': 'O'}
        }
        
        for r_key, row in chk_rows.items():
            for c_key, options in chk_cols.items():
                val = v.get(f"{r_key}_{c_key}")
                if val and val in options:
                    safe_write(f"{options[val]}{row}", 'V')

        wb.save(output_path)
        return output_path

