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

        def safe_write(cell_coord, value, is_currency=False):
            if not cell_coord: return
            try:
                cell = sheet[cell_coord]
                cell.value = value
                cell.font = Font(name='맑은 고딕', size=9)
                # 내용을 중앙 정렬로 설정
                cell.alignment = Alignment(shrinkToFit=True, vertical='center', horizontal='center')
                # 금액단위(원) 및 콤마 처리
                if is_currency:
                    cell.number_format = '#,##0 "원"'
            except:
                pass

        # 1. Header (Date)
        d = data.get('date', datetime.date.today())
        weekdays = ["월", "화", "수", "목", "금", "토", "일"]
        date_str = f"{d.year}년 {d.month}월 {d.day}일 ({weekdays[d.weekday()]})"
        safe_write(mapping.get('header', {}).get('date', 'F2'), date_str)

        # 2. General
        gen_map = mapping.get('general', {})
        for key in ['company', 'project_name', 'standard', 'equipment', 'report_no', 'inspection_item', 'inspector', 'car_no', 'inspector_n8']:
            # If key is inspector_n8, take data from 'inspector' field
            val_key = 'inspector' if key == 'inspector_n8' else key
            safe_write(gen_map.get(key), data.get(val_key, ''))

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
        current_methods = list(methods_data.keys())
        num_methods = len(current_methods)
        
        # 4개 초과 시 행 추가 및 기존 18-25행에서 삭제 (동적 레이아웃 유지)
        method_base_rows = [13, 14, 15, 16]
        if num_methods > 4:
            extra_count = num_methods - 4
            # 17행 위치에 필요한 만큼 행 삽입
            sheet.insert_rows(17, extra_count)
            # 기존 18~25행 영역(이제는 18+extra_count 이후)에서 삽입된 만큼 행 삭제
            # 18행부터 삭제를 시작하여 전체 레이아웃 길이를 맞춤
            for _ in range(extra_count):
                sheet.delete_rows(18 + extra_count) 
            
            # 사용할 행 리스트 업데이트
            method_base_rows += list(range(17, 17 + extra_count))
        
        for idx, m_name in enumerate(current_methods):
            if idx >= len(method_base_rows): break
            
            row = method_base_rows[idx]
            m_data = methods_data[m_name]
            
            safe_write(f'B{row}', m_name)
            safe_write(f'E{row}', m_data.get('unit', ''))
            safe_write(f'H{row}', m_data.get('qty', 0))
            safe_write(f'K{row}', m_data.get('price', 0), is_currency=True)
            safe_write(f'N{row}', m_data.get('travel', 0), is_currency=True)
            safe_write(f'Q{row}', m_data.get('total', 0), is_currency=True)

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
        
        # [NEW] Clear OT placeholders before writing to prevent ghost data from template
        for row_idx in [1, 2]:
            for key_suffix in ['_name', '_company', '_method', '_hours', '_amount']:
                cell_key = f'row{row_idx}{key_suffix}'
                if cell_key in ot_map:
                    safe_write(ot_map[cell_key], '')

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


        # [NEW] Clear template placeholder in Q39
        safe_write('Q39', '')








        # 6. 자재 수행현황 (Materials)
        mat_map = mapping.get('materials', {})
        materials_data = data.get('materials', {})

        # --- Material display name mapping (for D column) ---
        mat_display_names = {
            'RT T200':      ('RT', 'T200'),
            'RT AA400':     ('RT', 'AA400'),
            'RT Other':     ('RT', '기타'),
            'MT WHITE':     ('MT', '백색페인트'),
            'MT 7C-BLACK':  ('MT', '흑색자분'),
            'PT Penetrant': ('PT', '침투제'),
            'PT Cleaner':   ('PT', '세척제'),
            'PT Developer': ('PT', '현상제'),
        }

        # Clear rows 43-50 first (K~Q for quantities, D/F for names)
        for r in range(43, 51):
            for c_idx in range(11, 18):  # K(11) to Q(17)
                safe_write(f"{chr(64 + c_idx)}{r}", None)
            safe_write(f'D{r}', '')
            safe_write(f'F{r}', '')

            # --- [STRONG OVERRIDE] 화학자재 행 번호 기반 명칭 및 단위 강제 고정 ---
            override_names = {
                46: ('백색페인트', ''),
                47: ('흑색자분', ''),
                48: ('침투제', ''),
                49: ('세척제', ''),
                50: ('현상제', ''),
            }
            d_val_override = None
            f_val_override = None
            if r in override_names:
                d_val_override, f_val_override = override_names[r]

            if d_val_override is not None:
                safe_write(f'D{r}', d_val_override)
                safe_write(f'F{r}', f_val_override)

        # Write each material row
        for mat_name, row in mat_map.items():
            row = int(row)
            m_data = materials_data.get(mat_name, {})
            used_val = m_data.get('used', 0)

            # DB에서 온 상세 이름/규격 (주로 RT용)
            d_val = m_data.get('name', '')
            f_val = m_data.get('spec', '')

            # 해당 행이 고정 명칭 대상이 아닐 때만(예: RT 자재) DB 값 기입
            # 고정 명칭 대상(46~50)은 위에서 이미 기입했으므로 건너뜀
            if row not in (46, 47, 48, 49, 50):
                if d_val: safe_write(f'D{row}', d_val)
                if f_val: safe_write(f'F{row}', f_val)

            # 수량이 있을 때만 K/M/O 기입 (사용량이 없으면 행 전체를 비워둠)
            if used_val and used_val > 0:
                safe_write(f'K{row}', '-')                            # K: 하이픈
                safe_write(f'M{row}', used_val)                       # M: 당일사용량
                safe_write(f'O{row}', m_data.get('in', 0))           # O: 반입

        # Unit column H: RT rows '매', Others by row override
        for mat_name, r in mat_map.items():
            row = int(r)
            if row in (46, 47, 48, 49, 50): 
                safe_write(f'H{row}', 'CAN')
            elif row in (43, 44, 45):
                safe_write(f'H{row}', '매')
            else:
                safe_write(f'H{row}', '통')
        
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
                    safe_write(f"{options[val]}{row}", f"{val} ✔")

        wb.save(output_path)
        return output_path

