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

        # 집계 구조
        qty_summary = {}
        ndt_groups = {
            'PAUT': defaultdict(lambda: {'count': 0, 'qty': 0.0}),
            'MT': defaultdict(lambda: {'count': 0, 'qty': 0.0}),
            'RT': defaultdict(lambda: {'count': 0, 'qty': 0.0}),
            'PT': defaultdict(lambda: {'count': 0, 'qty': 0.0})
        }

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

        # 역순으로 채워야 행 삽입 시 위에 있는 인덱스(row number)가 변하지 않음
        # 따라서 아래에서부터 표를 처리합니다.
        
        # 표 위치들을 리스트로 묶어서 역순 처리
        sections = [
            ('pt_2', 'PT'), ('rt_2', 'RT'), ('mt_2', 'MT'), ('paut_2', 'PAUT'),
            ('pt_1', 'PT'), ('rt_1', 'RT'), ('mt_1', 'MT'), ('paut_1', 'PAUT')
        ]
        
        for section_key, method in sections:
            header_row = self.table_markers[section_key]
            start_row = header_row + 2
            
            # TOTAL 행을 찾음 (header_row 이후 빈 행이 아닌 "TOTAL"이 있는 행)
            total_row = start_row
            for r in range(start_row, start_row + 20):
                val = str(ws.cell(row=r, column=2).value or "").strip()
                if "TOTAL" in val.upper() or "총계" in val:
                    total_row = r
                    break
            
            # 데이터 삽입
            data_items = list(ndt_groups[method].items())
            num_items = len(data_items)
            
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
                ws.insert_rows(start_row, rows_to_insert)
                total_row += rows_to_insert
                # 모든 기존 마커 업데이트 (삽입된 행 아래에 있는 마커들만 += rows_to_insert)
                for k, v in self.table_markers.items():
                    if v >= start_row:
                        self.table_markers[k] += rows_to_insert
            
            # 데이터 쓰기
            current_row = start_row
            for idx, ((comp, sec, line, size, spec), vals) in enumerate(data_items):
                count = vals['count']
                qty = vals['qty']
                unit = "매" if method == "RT" else "m"
                
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

                _write_safe(current_row, 17, f"{qty:.4f}" if qty % 1 != 0 else str(int(qty)))
                if method == 'PAUT':
                    _write_safe(current_row, 20, f"{qty:.4f}" if qty % 1 != 0 else str(int(qty)))
                
                # 병합 처리 (템플릿 양식에 맞춤)
                safe_merge(current_row, 4, current_row, 5) # Section
                safe_merge(current_row, 6, current_row, 9) # Line No.
                safe_merge(current_row, 10, current_row, 11) # 관경
                safe_merge(current_row, 12, current_row, 13) # 용접개소
                safe_merge(current_row, 14, current_row, 15) # 규격
                
                if method == 'PAUT':
                    safe_merge(current_row, 18, current_row, 19) # RE'
                else:
                    safe_merge(current_row, 17, current_row, 20) # 길이
                    
                # Line No. 글씨 다 보이도록 높이 조절
                ws.row_dimensions[current_row].height = 45

                # 폰트, 정렬 적용 (병합된 칸 전체에 정렬 속성을 먹여야 엑셀이 줄바꿈을 정상 인식함)
                for c_idx in range(2, 21):
                    c_cell = ws.cell(row=current_row, column=c_idx)
                    c_cell.font = self.font_normal
                    c_cell.alignment = self.align_center
                    c_cell.border = self.border_thin

                current_row += 1
                
            # 빈 행(데이터 없는 행)들 스타일 정리 및 PAUT 병합 처리 준비
            for r in range(current_row, total_row):
                # 빈 행도 기본적으로 다른 칸들도 병합을 유지해야 템플릿 양식이 안 깨짐
                safe_merge(r, 4, r, 5)
                safe_merge(r, 6, r, 9)
                safe_merge(r, 10, r, 11)
                safe_merge(r, 12, r, 13)
                safe_merge(r, 14, r, 15)
                
                if method == 'PAUT':
                    safe_merge(r, 18, r, 19)
                else:
                    safe_merge(r, 17, r, 20)
                for c_idx in range(2, 21):
                    cell = ws.cell(row=r, column=c_idx)
                    cell.border = self.border_thin

            # TOTAL 행 값 쓰기 및 스타일 보정
            total_qty = sum(vals['qty'] for _, vals in data_items)
            
            # TOTAL 행의 17, 18, 19, 20 칸 배경색(파란색)을 동일하게 맞춤 (첫 번째 셀 서식 복사)
            fill_style = ws.cell(row=total_row, column=2).fill
            if method == 'PAUT':
                import copy
                for c_idx in range(17, 21):
                    c = ws.cell(row=total_row, column=c_idx)
                    if fill_style:
                        c.fill = copy.copy(fill_style)

            if total_qty > 0:
                from openpyxl.styles import Alignment as _Align
                shrink_align = _Align(horizontal='center', vertical='center', shrink_to_fit=True)
                if method == 'PAUT':
                    _write_safe(total_row, 17, f"{total_qty:.4f}" if total_qty % 1 != 0 else str(int(total_qty)))
                    ws.cell(row=total_row, column=17).alignment = shrink_align
                    _write_safe(total_row, 20, f"{total_qty:.4f}" if total_qty % 1 != 0 else str(int(total_qty)))
                    ws.cell(row=total_row, column=20).alignment = shrink_align
                else:
                    _write_safe(total_row, 17, f"{total_qty:.4f}" if total_qty % 1 != 0 else str(int(total_qty)))
                    ws.cell(row=total_row, column=17).alignment = shrink_align

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
        ws.column_dimensions['Q'].width = 9.0
        ws.column_dimensions['R'].width = 4.5
        ws.column_dimensions['S'].width = 4.5
        ws.column_dimensions['T'].width = 4.5
        ws.column_dimensions['U'].width = 4.5
        ws.column_dimensions['V'].width = 0.0   # 비고 우측 끝 숨김
        ws.column_dimensions['V'].hidden = True
        ws.column_dimensions['W'].width = 0.0   # 비고 우측 끝 숨김
        ws.column_dimensions['W'].hidden = True

        # 사용자 요청: 모든 표의 좌측(2열), 우측(21열) 외곽선을 굵은 실선으로 복원/강제 설정
        medium = Side(style='medium')
        def apply_outer_borders(start_r, end_r):
            for r in range(start_r, end_r + 1):
                lc = ws.cell(row=r, column=2)
                lc.border = Border(left=medium, right=lc.border.right, top=lc.border.top, bottom=lc.border.bottom)
                rc = ws.cell(row=r, column=21)
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
                            # 사용자가 3페이지 등에서 '=F41' 처럼 수식으로 제목을 끌고 오는 경우 대응
                            import re
                            match = re.search(r'F\$?(\d+)', val_no_spaces, re.IGNORECASE)
                            if match:
                                ref_row = int(match.group(1))
                                # 41, 81, 121 등 40행 간격으로 제목이 위치하므로 이를 수식에서 참조하면 제목 칸으로 간주
                                if ref_row % 40 == 1:
                                    is_title = True
                                    
                        if is_title:
                            contract_name = replacements.get('[[계약명]]', '2026년 중앙지사 열수송관 비파괴검사용역 단가계약')
                            # 무조건 괄호와 줄바꿈이 포함된 포맷으로 강제 덮어쓰기 (수식도 문자열로 덮어씌움)
                            cell.value = f"【{contract_name}】\n월 간 용 역 진 도 보 고 서"
                            
                            from openpyxl.styles import Alignment, Font
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
