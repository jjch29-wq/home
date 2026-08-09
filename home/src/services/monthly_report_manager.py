import openpyxl
from openpyxl.styles import Alignment, Border, Side, Font
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

    def generate_report(self, history_path, year_month, output_path):
        """
        history_path: daily_work_history.json 경로
        year_month: "YYYY-MM" 형식 (예: "2026-08")
        """
        if not os.path.exists(self.template_path):
            raise FileNotFoundError(f"Template not found at {self.template_path}")
            
        if not os.path.exists(history_path):
            raise FileNotFoundError(f"History file not found at {history_path}")

        with open(history_path, 'r', encoding='utf-8') as f:
            history_data = json.load(f)

        # 1. 데이터 필터링 및 집계
        target_dates = sorted([d for d in history_data.keys() if d.startswith(year_month)])
        if not target_dates:
            print(f"No data found for {year_month}")
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
                ws.cell(row=current_row, column=17).value = f"{qty:.4f}" if qty % 1 != 0 else str(int(qty))
                
                # 병합 처리 (템플릿 양식에 맞춤)
                ws.merge_cells(start_row=current_row, start_column=4, end_row=current_row, end_column=5) # Section
                ws.merge_cells(start_row=current_row, start_column=6, end_row=current_row, end_column=9) # Line No.
                ws.merge_cells(start_row=current_row, start_column=10, end_row=current_row, end_column=11) # 관경
                ws.merge_cells(start_row=current_row, start_column=12, end_row=current_row, end_column=13) # 용접개소
                ws.merge_cells(start_row=current_row, start_column=14, end_row=current_row, end_column=15) # 규격
                ws.merge_cells(start_row=current_row, start_column=17, end_row=current_row, end_column=19) # 길이
                
                # 모든 병합된/단일 셀에 스타일 적용 (2~19번 열)
                for c_idx in range(2, 20):
                    cell = ws.cell(row=current_row, column=c_idx)
                    cell.font = self.font_normal
                    cell.alignment = self.align_center
                    cell.border = self.border_thin

                current_row += 1
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
