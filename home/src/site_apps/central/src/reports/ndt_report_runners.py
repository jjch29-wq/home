# -*- coding: utf-8 -*-
"""각 NDT 검사방법별(PMI, PT, MT, RT) 엑셀 성적서 출력 실행 믹스인 (NDT Report Runners Mixin)"""
import os
import io
import re
import copy
import glob
import tempfile
import datetime
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU

thin_side = Side(style='thin')
medium_side = Side(style='medium')
thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)


class NDTReportRunnersMixin:
    def _write_gapji_metadata(self, ws, mode="PMI", first_item=None):
        """Write common report metadata to the cover sheet."""
        if mode == "PAUT":
            # 중앙지사 PAUT 표준 표기
            self.safe_set_value(ws, "G9", "중앙지사 열수송관 공사")
            self.safe_set_value(ws, "AB9", "SIS N 264 KDHC DT Rev.0")
            mapping = [
                ('GAPJI_PROJECT', 'G8'),
                ('GAPJI_CUSTOMER', 'AA4'),
                ('GAPJI_EXAM_DATE', 'AA6'),
                ('GAPJI_REPORT_NO', 'AA5'),
                ('PAUT_SEC', 'M35'),
                ('GAPJI_MATERIAL', 'A35'),
            ]
            for key, default_cell in mapping:
                val = self.paut_sec.get().strip() if key == 'PAUT_SEC' else self.config.get(key, "")
                if key == 'PAUT_SEC' and val:
                    val = 'Sec.' + re.sub(r'^sec\s*\.?\s*', '', val, flags=re.I)
                if val or key == 'PAUT_SEC':
                    self.safe_set_value(ws, default_cell, val)

            # PAUT 원본 양식에 남아 있는 동탄지사 문구를 중앙지사 기준으로 교체한다.
            general_information = {
                "A9": self.paut_contractor.get().strip(),
                "G9": "중앙지사 열수송관 공사",
                "AB9": "SIS N 264 KDHC-JY Rev.0",
                "AA44": "한국지역난방공사 중앙지사",
            }
            gain_value = self.config.get(
                "PAUT_EQUIP_GAIN", "11 + 6dB / 21.4 + 6dB"
            )
            if hasattr(self, "paut_equip_vars") and "gain" in self.paut_equip_vars:
                gain_value = self.paut_equip_vars["gain"].get() or gain_value
            general_information["R38"] = gain_value

            for cell_ref, cell_value in general_information.items():
                self.safe_set_value(ws, cell_ref, cell_value)

            if first_item:
                line_no = first_item.get('Line No.', '')
                if line_no:
                    self.safe_set_value(ws, 'Y35', line_no)

                size_val = first_item.get('Size', '')
                if size_val:
                    self.safe_set_value(ws, 'J35', size_val)

                thk_val = first_item.get('Th\'k(mm)', '')
                if thk_val:
                    self.safe_set_value(ws, 'G35', thk_val)
        elif mode in ("RT", "KOGAS"):
            # RT/KOGAS 갑지 전용 좌표. PMI 기본 좌표를 사용하면 B9가
            # A8:B9 병합영역의 A8로 이동하여 날짜가 부재두께/관경 칸에 들어간다.
            mapping = [
                ('GAPJI_PROJECT', 'Q2'),
                ('GAPJI_CUSTOMER', 'Q1'),
                ('GAPJI_MATERIAL', 'C10'),
                ('GAPJI_EXAM_DATE', 'I28'),
                ('GAPJI_REPORT_NO', 'C5'),
            ]
            for cfg_key, coord in mapping:
                val = self.config.get(cfg_key, "")
                if val:
                    self.safe_set_value(ws, coord, val)
        elif mode == "PT":
            # [NEW] PT 지역난방 템플릿 갑지 전용 좌표 매핑
            mapping = [
                ('GAPJI_PROJECT', 'A9'),   # 공사명 (Project Name)
                ('GAPJI_CUSTOMER', 'P4'),  # 발주처 (Customer)
                ('GAPJI_ITEM', 'F9'),      # 품명 (Item Name)
                ('GAPJI_MATERIAL', 'A11'), # 재질 (Material)
                ('GAPJI_REPORT_NO', 'P6')  # 리포트번호 (Report No.)
                # 검사일자(Date)는 원본 템플릿 상단에 명시적인 고정 칸이 없음
            ]
            for key, default_cell in mapping:
                val = self.config.get(key, "")
                if val:
                    self.safe_set_value(ws, default_cell, val)

            if first_item:
                # 첫 번째 데이터의 도면번호(ISO/Dwg)를 갑지의 Dwg. No 칸(L9)에 자동 기입
                line_no = first_item.get('Line No.', '') or first_item.get('Dwg', '') or first_item.get('ISO', '')
                if line_no:
                    self.safe_set_value(ws, 'L9', line_no)
        elif mode == 'MT':
            for key, cell in [('GAPJI_PROJECT', 'B9'), ('GAPJI_CUSTOMER', 'S4'),
                              ('GAPJI_ITEM', 'J9'), ('GAPJI_MATERIAL', 'B11'),
                              ('GAPJI_REPORT_NO', 'S6')]:
                self.safe_set_value(ws, cell, self.config.get(key, ''))
        else:
            # PMI
            try:
                a11_val = str(ws['A11'].value or '')
            except Exception:
                a11_val = ''
            try:
                l5_val = str(ws['L5'].value or '')
            except Exception:
                l5_val = ''
            if ('PROJECT' in a11_val.upper() or '공사명' in a11_val) and ('CUSTOMER' in l5_val.upper() or '주문주' in l5_val):
                # SITCO standard PMI layout (e.g. K1 / K3 / 롯데 / PROVIDENCE)
                combined_pmi_item = self._get_combined_pmi_item_names() if hasattr(self, '_get_combined_pmi_item_names') else ''
                sitco_pmi_mapping = [
                    ('GAPJI_PROJECT', 'E11', getattr(self, 'gapji_project', None)),
                    ('GAPJI_CUSTOMER', 'N5', getattr(self, 'gapji_customer', None)),
                    ('GAPJI_REPORT_NO', 'N8', getattr(self, 'gapji_report_no', None)),
                    ('GAPJI_ITEM', 'E13', getattr(self, 'gapji_item', None)),
                    ('GAPJI_MATERIAL', 'E17', getattr(self, 'gapji_material', None)),
                ]
                for cfg_key, coord, ui_var in sitco_pmi_mapping:
                    if cfg_key == 'GAPJI_ITEM' and combined_pmi_item:
                        val = combined_pmi_item
                    else:
                        ui_val = ui_var.get().strip() if (ui_var and hasattr(ui_var, 'get')) else ''
                        val = ui_val or self.config.get(cfg_key, "")
                    if not val and cfg_key == 'GAPJI_REPORT_NO' and first_item:
                        val = str(first_item.get('Report No', '') or first_item.get('Report No.', '')).strip()
                    if val:
                        self.safe_set_value(ws, self.config.get(f"{cfg_key}_CELL", coord), val)

                # [SMART] K11 라벨을 감지하여 N11에 도면번호 또는 검사일자 정밀 매핑
                k11_val = str(ws['K11'].value or '').upper()
                if 'DOC' in k11_val or 'DWG' in k11_val or '도면' in k11_val:
                    first_dwg = ""
                    if first_item:
                        first_dwg = first_item.get('Line No.', '') or first_item.get('Dwg', '') or first_item.get('ISO', '')
                    dwg_val = self.config.get('GAPJI_DWG', '') or first_dwg
                    if dwg_val:
                        self.safe_set_value(ws, 'N11', dwg_val)
                else:
                    exam_date_val = self.config.get('GAPJI_EXAM_DATE', '')
                    if exam_date_val:
                        self.safe_set_value(ws, 'N11', exam_date_val)

                # 보고서 작성일자: K42 (또는 J40)
                exam_date = (self.gapji_exam_date.get().strip() if hasattr(self, 'gapji_exam_date') else '') or self.config.get('GAPJI_EXAM_DATE', '')
                if exam_date:
                    if ws.cell(row=42, column=11).value is not None or 'Date of Report' in str(ws.cell(row=40, column=11).value or ''):
                        self.safe_set_value(ws, 'K42', exam_date)
                    elif ws.cell(row=40, column=10).value is not None:
                        self.safe_set_value(ws, 'J40', exam_date)

                # [NEW] 갑지 1. Acceptance Criteria (%) 표(24~39행, A~T열)를 [방안 B] 가로 비교 매트릭스 형태로 자동 갱신
                if hasattr(self, '_apply_pmi_acceptance_criteria'):
                    self._apply_pmi_acceptance_criteria(ws)
            else:
                # PMI templates already contain their own date field. Do not use B9
                # as a fallback because it belongs to the cover logo/header area.
                combined_pmi_item = self._get_combined_pmi_item_names() if hasattr(self, '_get_combined_pmi_item_names') else ''
                mapping = [
                    ('GAPJI_PROJECT', 'B5'),
                    ('GAPJI_CUSTOMER', 'B6'),
                    ('GAPJI_ITEM', 'B7'),
                    ('GAPJI_MATERIAL', 'B8'),
                    ('GAPJI_REPORT_NO', 'B10')
                ]

                for cfg_key, coord in mapping:
                    if cfg_key == 'GAPJI_ITEM' and combined_pmi_item:
                        val = combined_pmi_item
                    else:
                        val = self.config.get(cfg_key, "")
                    if val:
                        self.safe_set_value(ws, self.config.get(f"{cfg_key}_CELL", coord), val)

                # Write the date only when an explicit template coordinate is set.
                exam_date_cell = str(self.config.get('GAPJI_EXAM_DATE_CELL', '')).strip()
                exam_date = self.config.get('GAPJI_EXAM_DATE', '')
                if exam_date_cell and exam_date:
                    self.safe_set_value(ws, exam_date_cell, exam_date)

    def _run_pmi_process(self, final_list, template_path):
        self.save_settings() # Ensure UI -> config sync
        self.log(f"🚀 PMI 성적서 생성 시작 (총 {len(final_list)} 건)...")
        self.progress['value'] = 0
        all_extracted_data = final_list

        # This name is reassigned further below, so Python treats it as local
        # throughout this function. Initialise it before the cover border work.
        thin_side = Side(style='thin')

        # [FIX] data_end_row 전역/설정으로 보강
        data_start_row = int(self.config.get('START_ROW', 17))
        data_end_row = int(self.config.get('DATA_END_ROW', 47))

        try:
            # keep_vba=True forces openpyxl to write a macro-enabled workbook
            # content type. When the selected template is a normal .xlsx file,
            # saving that workbook with an .xlsx extension makes Excel reject it
            # as an extension/format mismatch.
            template_ext = os.path.splitext(template_path)[1].lower()
            keep_vba = template_ext == '.xlsm'
            wb = openpyxl.load_workbook(template_path, keep_vba=keep_vba)
            if len(wb.worksheets) < 1:
                raise ValueError("선택한 템플릿 파일에 시트가 존재하지 않습니다.")

            if len(wb.worksheets) >= 1:
                ws0 = wb.worksheets[0]
                self.add_logos_to_sheet(ws0, is_cover=True, clear_existing=False, mode="PMI")
                self._write_gapji_metadata(ws0)
                self.force_print_settings(ws0, context="COVER") # [NEW] 갑지 전용 여백 적용

                # [DYNAMIC ELEMENTS] Gapji Signature Area Border Stabilization
                # Signature usually ends at row 48. We must ensure no vertical lines extend below it.
                b_start = int(self.config.get('GAPJI_START_ROW', 23))
                b_end = int(self.config.get('GAPJI_DATA_END_ROW', 38))
                sig_end = 48 # Standard signature end row for this template

                # 1. Apply side borders only within the valid range
                for r in range(b_start, sig_end + 1):
                    try:
                        # Left Border (Column A)
                        cell_a = ws0.cell(row=r, column=1); eb = cell_a.border
                        ws0.cell(row=r, column=1).border = Border(left=medium_side, right=eb.right, top=eb.top, bottom=eb.bottom)
                        # Right Border (Column T)
                        cell_t = ws0.cell(row=r, column=20); eb_t = cell_t.border
                        ws0.cell(row=r, column=20).border = Border(left=eb_t.left, right=medium_side, top=eb_t.top, bottom=eb_t.bottom)
                    except: pass

                def _add_border(cell_ref, *, right=None, bottom=None):
                    cell = ws0[cell_ref]
                    eb = cell.border
                    cell.border = Border(
                        left=eb.left,
                        right=right if right is not None else eb.right,
                        top=eb.top,
                        bottom=bottom if bottom is not None else eb.bottom,
                    )

                # [FIX] 결재란 50행(Row 50) 바닥 테두리 완벽 밀폐 (A50~T50 바닥 굵은선 밀폐 및 박스 완성)
                for c in range(1, 21):
                    cell = ws0.cell(row=50, column=c)
                    eb = cell.border
                    l_s = medium_side if c == 1 else (eb.left or None)
                    r_s = medium_side if c == 20 else (thin_side if c == 10 else (eb.right or None))
                    cell.border = Border(left=l_s, right=r_s, top=eb.top, bottom=medium_side)

                # 2. [FIX] 52행 이하 불필요한 고스트 선/병합 정리 (51행 바닥글 행 보존)
                for r in range(52, 70):
                    merged_ranges = [str(rng) for rng in ws0.merged_cells.ranges if rng.min_row >= r]
                    for m_rng in merged_ranges:
                        try: ws0.unmerge_cells(m_rng)
                        except: pass
                    for c in range(1, 26):
                        try: ws0.cell(row=r, column=c).border = Border()
                        except: pass

                # 바닥글(51행)의 좌측 양식 번호와 우측 로고까지 온전히 포함하는 공식 인쇄 영역
                gapji_print_end = int(self.config.get('GAPJI_PRINT_END_ROW', 51))
                if ws0.row_dimensions[gapji_print_end].height is None or ws0.row_dimensions[gapji_print_end].height < 24.0:
                    ws0.row_dimensions[gapji_print_end].height = 24.0
                ws0.print_area = f'A1:T{gapji_print_end}'

                ws0['I35'].border = Border() # [FIX] I35 셀 선 제거
                self.safe_set_value(ws0, 'I35', None) 
                self.apply_custom_dimensions(ws0, "COVER")

            def normalize_pmi_header_stamp(sheet, is_cover=False):
                """Normalize header logo and seal/stamp on the PMI sheet to ensure boundaries are respected.
                - Company logo: Placed within Column K (Cover) or Column J (Data) without clipping.
                - Stamp (직인): Sized to 84x84px (5-10 row box) and aligned flush against Column K (Cover) or Column J (Data).
                """
                try:
                    from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, TwoCellAnchor, AnchorMarker
                    from openpyxl.drawing.xdr import XDRPositiveSize2D
                    from openpyxl.utils.units import pixels_to_EMU

                    for img in getattr(sheet, '_images', []):
                        marker = getattr(img.anchor, '_from', None)
                        if marker is None:
                            continue
                        r, c = marker.row, marker.col
                        aspect = img.width / max(img.height, 1)

                        # 1. Company logo ("서울검사 주식회사") - 상단 외곽선(5행 윗선)을 가리지 않도록 rowOff를 5px(47,625 EMU) 아래로 안착
                        if 3 <= r <= 6 and 3 <= c <= 6 and 3.5 <= aspect <= 9.0:
                            if is_cover:
                                img.width = 280
                                img.height = 42
                                img.anchor = TwoCellAnchor(
                                    _from=AnchorMarker(col=4, colOff=pixels_to_EMU(2), row=4, rowOff=pixels_to_EMU(5)),
                                    to=AnchorMarker(col=10, colOff=pixels_to_EMU(25), row=6, rowOff=pixels_to_EMU(15))
                                )
                            else:
                                is_triple_mode = hasattr(self, 'pmi_layout_mode') and self.pmi_layout_mode.get() == "PLANT_TRIPLE"
                                if is_triple_mode:
                                    img.width = 220
                                    img.height = 38
                                    img.anchor = TwoCellAnchor(
                                        _from=AnchorMarker(col=4, colOff=pixels_to_EMU(2), row=4, rowOff=pixels_to_EMU(5)),
                                        to=AnchorMarker(col=8, colOff=pixels_to_EMU(15), row=6, rowOff=pixels_to_EMU(15))
                                    )
                                else:
                                    img.width = 230
                                    img.height = 40
                                    img.anchor = TwoCellAnchor(
                                        _from=AnchorMarker(col=4, colOff=pixels_to_EMU(2), row=4, rowOff=pixels_to_EMU(5)),
                                        to=AnchorMarker(col=9, colOff=pixels_to_EMU(20), row=6, rowOff=pixels_to_EMU(15))
                                    )

                        # 2. Seal / stamp (rows 2~6, cols 7~12, aspect 0.65~1.35)
                        # [EXACT BOX ALIGNMENT] 5~10행 세로 높이(84px) 정밀 일치 및 우측 끝선 무오차 밀착 (84x84px)
                        elif 2 <= r <= 6 and 7 <= c <= 12 and 0.65 <= aspect <= 1.35:
                            is_triple_mode = hasattr(self, 'pmi_layout_mode') and self.pmi_layout_mode.get() == "PLANT_TRIPLE"
                            if is_cover:
                                img.width = 84
                                img.height = 84
                                img.anchor = OneCellAnchor(
                                    _from=AnchorMarker(col=8, colOff=179300, row=4, rowOff=0),
                                    ext=XDRPositiveSize2D(800100, 800100)
                                )
                            elif is_triple_mode:
                                img.width = 84
                                img.height = 84
                                img.anchor = OneCellAnchor(
                                    _from=AnchorMarker(col=8, colOff=140000, row=4, rowOff=0),
                                    ext=XDRPositiveSize2D(800100, 800100)
                                )
                            else:
                                # 20열 단일행 표준 모드: 회사 정보 박스는 E~L열, M열부터 Customer
                                # 상호명 "株式會社" 위에 정확히 날인되며, L열 우측(M열 Customer 경계선) 침범 없이 마감
                                img.width = 84
                                img.height = 84
                                img.anchor = OneCellAnchor(
                                    _from=AnchorMarker(col=9, colOff=161925, row=4, rowOff=0),
                                    ext=XDRPositiveSize2D(800100, 800100)
                                )

                        # 3. SITCO logo (A5:D10 box, cols 0~3, rows 5~10)
                        # [EXACT BOX CENTERING] A5:D10 박스(가로 약 99px, 세로 84px) 내 상하좌우 완벽한 정중앙 대칭 안착
                        # - 상하: 5행(row 4) 8px 지점부터 10행(row 9) 6px 지점까지 (상하 여백 각 8px 균등)
                        # - 좌우: A열(col 0) 6px 지점부터 D열(col 3) 12px 지점까지 (좌우 여백 각 6px 균등)
                        elif c == 0 and 3 <= r <= 6 and 2.5 <= aspect <= 4.0:
                            img.anchor = TwoCellAnchor(
                                _from=AnchorMarker(col=0, colOff=pixels_to_EMU(6), row=4, rowOff=pixels_to_EMU(8)),
                                to=AnchorMarker(col=3, colOff=pixels_to_EMU(12), row=9, rowOff=pixels_to_EMU(6))
                            )
                except Exception:
                    pass

            if ws0 is not None:
                normalize_pmi_header_stamp(ws0, is_cover=True)

            data_sheet_id = 1 if len(wb.worksheets) >= 2 else 0
            ws = wb.worksheets[data_sheet_id]; ws.title = f"{ws.title[:20]}_001"
            # 을지 기본 설정 (기존 템플릿 로고 보존을 위해 clear_existing=False)
            self.add_logos_to_sheet(ws, is_cover=False, clear_existing=False, mode="PMI")
            normalize_pmi_header_stamp(ws, is_cover=False)

            self.force_print_settings(ws, context="DATA"); self.set_eulji_headers(ws)
            # self.apply_custom_dimensions(ws, "DATA") # [MOVED] To the end of process

            # [NEW] 일반 단일행 템플릿(20열)에서 '플랜트 배관용' 모드 선택 시,
            # 을지 헤더(11~18행) 및 상단 구조를 13열 3단 배관용 구조(사진 3)로 자동 변환
            is_triple_mode = hasattr(self, 'pmi_layout_mode') and self.pmi_layout_mode.get() == "PLANT_TRIPLE"
            if is_triple_mode:
                a11_val = str(ws.cell(11, 1).value or '').strip()
                merges_str = [str(m) for m in ws.merged_cells.ranges]
                if 'Item' in a11_val or ws.max_column > 13 or ws.cell(11, 15).value is not None or 'A11:D18' in merges_str:
                    self.convert_pmi_eulji_to_plant_triple(ws, ws0=ws0)

            try:
                if len(wb.worksheets) >= 2:
                    is_triple_mode = hasattr(self, 'pmi_layout_mode') and self.pmi_layout_mode.get() == "PLANT_TRIPLE"
                    if is_triple_mode:
                        # 13열 플랜트 3단 모드 (K5: 주문주, M5: 고객사, K8: 보고서번호, M8: 번호값)
                        self.safe_set_value(ws, 'K5', f"='{ws0.title}'!L5")
                        self.safe_set_value(ws, 'M5', f"='{ws0.title}'!N5")
                        self.safe_set_value(ws, 'K8', f"='{ws0.title}'!L8")
                        self.safe_set_value(ws, 'M8', f"='{ws0.title}'!N8")
                        ws.column_dimensions['M'].width = 18.0
                    else:
                        # 20열 표준 모드 (M5: 주문주 라벨 보존, O5: 고객사 연동, M8: 보고서번호 라벨 보존, O8: 번호 연동)
                        self.safe_set_value(ws, 'K5', None)  # K열 상호 로고 영역 잔여 텍스트 정리
                        m5_val = ws.cell(5, 13).value
                        if not m5_val or '한국' in str(m5_val):
                            self.safe_set_value(ws, 'M5', '주문주\nCustomer')
                        self.safe_set_value(ws, 'O5', f"='{ws0.title}'!N5")
                        m8_val = ws.cell(8, 13).value
                        if not m8_val or 'SIS' in str(m8_val):
                            self.safe_set_value(ws, 'M8', '보고서 번호\nReport No.')
                        rep_no = (self.gapji_report_no.get().strip() if hasattr(self, 'gapji_report_no') else '') or self.config.get('GAPJI_REPORT_NO', '')
                        if rep_no:
                            self.safe_set_value(ws, 'O8', rep_no)
                        else:
                            self.safe_set_value(ws, 'O8', f"=IF('{ws0.title}'!N8=\"\",\"\",'{ws0.title}'!N8)")

                        pmi_20col_widths = {
                            'H': 4.83, 'I': 4.83, 'J': 4.61, 'K': 4.61, 'L': 4.61, 'M': 4.61,
                            'N': 11.00, 'O': 6.00
                        }
                        for col_l, col_w in pmi_20col_widths.items():
                            ws.column_dimensions[col_l].width = col_w

                        for r_idx in range(5, 11):
                            for c_idx in (13, 15):
                                cell = ws.cell(row=r_idx, column=c_idx)
                                cell.font = Font(name='바탕', size=9, bold=False)
                        ws.cell(row=5, column=13).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                        ws.cell(row=8, column=13).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                        ws.cell(row=5, column=15).alignment = Alignment(horizontal='center', vertical='center', shrink_to_fit=True)
                        ws.cell(row=8, column=15).alignment = Alignment(horizontal='center', vertical='center', shrink_to_fit=True)
            except Exception as e:
                self.log(f"갑지 데이터 복사 실패: {e}")

            materials = "SS304,SS304L,SS316,SS316L,SS321,SS347,SS410,SS430,DUPLEX,MONEL,INCONEL,ER308,ER308L,ER309,ER309L,ER316,ER316L,ER347,ER2209,WP316,WP316L,TP316,TP316L,F316L,A182-F316L,A312-TP316L"
            dv_q = DataValidation(type="list", formula1=f'"{materials}"', allow_blank=True)
            # [NEW] 데이터 기입 전 기존 병합 영역 정리 (A-E 등 중복 병합 방지)
            def clear_merges_in_range(sheet, start_row, end_row):
                merged_ranges = list(sheet.merged_cells.ranges)
                for r in merged_ranges:
                    # 영역이 겹치면 해제 (aggressive clearing)
                    if r.min_row <= end_row and r.max_row >= start_row:
                        try: sheet.unmerge_cells(str(r))
                        except: pass

            ws = wb.worksheets[data_sheet_id]
            used_pmi_sheets = [ws0, ws] if ws0 is not None else [ws]
            def apply_pmi_variable_header(sheet):
                """Apply PMI fields that repeat on every data-sheet page."""
                is_triple = hasattr(self, 'pmi_layout_mode') and self.pmi_layout_mode.get() == "PLANT_TRIPLE"
                if is_triple:
                    cust = self.gapji_customer.get().strip()
                    rep = self.gapji_report_no.get().strip()
                    if cust:
                        self.safe_set_value(sheet, 'M5', cust)
                    if rep:
                        self.safe_set_value(sheet, 'M8', rep)
                    for col_idx, var in enumerate(self.pmi_chem_elements, start=8):
                        self.safe_set_value(
                            sheet, sheet.cell(row=15, column=col_idx).coordinate,
                            var.get().strip(), align='center'
                        )
                    for row_idx, var in enumerate(self.pmi_chem_specs, start=16):
                        val = var.get().strip()
                        self.safe_set_value(sheet, f'H{row_idx}', val if val else "-")
                else:
                    # 20열 표준 모드: M5는 '주문주\nCustomer', M8은 '보고서 번호\nReport No.' 고정 라벨 보존
                    self.safe_set_value(sheet, 'M5', '주문주\nCustomer')
                    self.safe_set_value(sheet, 'M8', '보고서 번호\nReport No.')
                    cust = self.gapji_customer.get().strip()
                    if cust:
                        self.safe_set_value(sheet, 'O5', cust)
                    elif ws0 is not None:
                        self.safe_set_value(sheet, 'O5', f"='{ws0.title}'!N5")
                    rep = (self.gapji_report_no.get().strip() if hasattr(self, 'gapji_report_no') else '') or self.config.get('GAPJI_REPORT_NO', '')
                    if rep:
                        self.safe_set_value(sheet, 'O8', rep)
                    elif ws0 is not None:
                        self.safe_set_value(sheet, 'O8', f"=IF('{ws0.title}'!N8=\"\",\"\",'{ws0.title}'!N8)")

                    # 일반 단일행: 템플릿의 기존 원소 및 기준란 형태를 유지하며 기입
                    # 표준 20열 템플릿: 15행은 'Chemical Composition'(H15:M16), 17행이 원소 행(H17:Cr, I17:Ni)
                    # 플랜트 3단 템플릿: 14행이 'Chemical Composition', 15행이 원소 행(H15:Ni, I15:Cr)
                    elem_row = 17
                    for r_check in (17, 15):
                        row_vals = [str(sheet.cell(row=r_check, column=c).value or '').strip().upper() for c in (8, 9, 10)]
                        if 'CR' in row_vals or 'NI' in row_vals:
                            elem_row = r_check
                            break

                    # 15행이 원소 행이 아닐 때는 H15에 'Chemical Composition'을 확실히 유지 (덮어쓰기 방지)
                    if elem_row != 15:
                        h15_val = str(sheet.cell(15, 8).value or '').strip()
                        if not h15_val or h15_val == '-':
                            self.safe_set_value(sheet, 'H15', "Chemical Composition")
                            sheet.cell(15, 8).alignment = Alignment(horizontal='center', vertical='center')
                            sheet.cell(15, 8).font = Font(name='바탕', size=9, bold=False)

                    if hasattr(self, 'pmi_chem_elements') and self.pmi_chem_elements:
                        for col_idx, var in enumerate(self.pmi_chem_elements, start=8):
                            val = var.get().strip()
                            if val:
                                coord = sheet.cell(row=elem_row, column=col_idx).coordinate
                                self.safe_set_value(sheet, coord, val, align='center')
                                sheet.cell(row=elem_row, column=col_idx).font = Font(name='바탕', size=9, bold=True)
                                sheet.cell(row=elem_row, column=col_idx).alignment = Alignment(horizontal='center', vertical='center', shrink_to_fit=True)

            apply_pmi_variable_header(ws)
            # Preserve a pristine copy of the PMI data-sheet header. Additional
            # pages must restore it after copying a populated worksheet.
            is_triple = hasattr(self, 'pmi_layout_mode') and self.pmi_layout_mode.get() == "PLANT_TRIPLE"
            pmi_start_row = max(19, int(self.config.get('START_ROW', 17))) if is_triple else int(self.config.get('START_ROW', 17))
            header_end_row = max(1, pmi_start_row - 1)
            header_cells = {}
            for header_row in range(1, header_end_row + 1):
                for header_col in range(1, 21):
                    source_cell = ws.cell(row=header_row, column=header_col)
                    header_cells[(header_row, header_col)] = {
                        'value': source_cell.value,
                        'style': copy.copy(source_cell._style),
                        'hyperlink': copy.copy(getattr(source_cell, 'hyperlink', None)),
                        'comment': copy.copy(getattr(source_cell, 'comment', None)),
                    }
            header_merges = [
                str(merged_range)
                for merged_range in ws.merged_cells.ranges
                if merged_range.max_row <= header_end_row
            ]
            header_row_heights = {
                row_idx: ws.row_dimensions[row_idx].height
                for row_idx in range(1, header_end_row + 1)
            }

            # Snapshot the pristine second-sheet data area before writing page 2.
            # Later pages are copied from that sheet after it has been populated,
            # so the snapshot is required to remove copied page-2 values safely.
            pmi_max_cols = 13 if is_triple else 20
            pmi_data_end_row = 45 if is_triple else int(self.config.get('DATA_END_ROW', 47))
            data_cells = {}
            for data_row in range(pmi_start_row, pmi_data_end_row + 1):
                for data_col in range(1, pmi_max_cols + 1):
                    source_cell = ws.cell(row=data_row, column=data_col)
                    data_cells[(data_row, data_col)] = {
                        'value': source_cell.value,
                        'style': copy.copy(source_cell._style),
                    }
            data_row_heights = {
                row_idx: ws.row_dimensions[row_idx].height
                for row_idx in range(pmi_start_row, pmi_data_end_row + 1)
            }

            def restore_pmi_header(sheet):
                for merged_range in list(sheet.merged_cells.ranges):
                    if merged_range.min_row <= header_end_row:
                        try:
                            sheet.unmerge_cells(str(merged_range))
                        except Exception:
                            pass
                for (header_row, header_col), snapshot in header_cells.items():
                    target_cell = sheet.cell(row=header_row, column=header_col)
                    target_cell.value = snapshot['value']
                    target_cell._style = copy.copy(snapshot['style'])
                    target_cell.hyperlink = copy.copy(snapshot['hyperlink'])
                    target_cell.comment = copy.copy(snapshot['comment'])
                for row_idx, row_height in header_row_heights.items():
                    sheet.row_dimensions[row_idx].height = row_height
                for merged_ref in header_merges:
                    sheet.merge_cells(merged_ref)
                if not is_triple:
                    sheet.cell(5, 13).value = '주문주\nCustomer'
                    sheet.cell(8, 13).value = '보고서 번호\nReport No.'
                    sheet.cell(15, 8).value = 'Chemical Composition'

            def reset_pmi_data_area(sheet):
                clear_merges_in_range(sheet, pmi_start_row, pmi_data_end_row)
                for (data_row, data_col), snapshot in data_cells.items():
                    target_cell = sheet.cell(row=data_row, column=data_col)
                    # Preserve the template style only. The template may still
                    # contain sample/previous report values, which must never
                    # appear below the last row on a partially filled page.
                    target_cell.value = None
                    target_cell._style = copy.copy(snapshot['style'])
                for row_idx, row_height in data_row_heights.items():
                    sheet.row_dimensions[row_idx].height = row_height

            reset_pmi_data_area(ws)
            ws.add_data_validation(dv_q)

            non_element_keys = {
                "_status", "selected", "No", "Date", "Report No", "Dwg", "Joint",
                "Loc", "Grade", "Result", "Remarks", "Welder", "order_index", "V", "ST",
                "Ni", "Cr", "Mo"
            }
            custom_elements = [
                k for k in self.column_keys
                if k not in non_element_keys
            ]
            ni_col = self.col_to_num(self.config.get('PMI_COL_NI', '8'))
            cr_col = self.col_to_num(self.config.get('PMI_COL_CR', '9'))
            mo_col = self.col_to_num(self.config.get('PMI_COL_MO', '10'))
            print_elements = [('Ni', ni_col), ('Cr', cr_col), ('Mo', mo_col)]
            if len(custom_elements) > 0:
                print_elements.append((custom_elements[0], 11))
            if len(custom_elements) > 1:
                print_elements.append((custom_elements[1], 12))

            current_row = pmi_start_row
            current_page = 1
            data_ptr = 0
            while data_ptr < len(all_extracted_data):
                rows_left = pmi_data_end_row - current_row + 1
                if rows_left <= 0:
                    current_page += 1
                    ws = self.prepare_next_sheet(wb, data_sheet_id, current_page, mode="PMI")
                    if ws not in used_pmi_sheets:
                        used_pmi_sheets.append(ws)
                    restore_pmi_header(ws)
                    normalize_pmi_header_stamp(ws, is_cover=False)
                    reset_pmi_data_area(ws)
                    current_row = pmi_start_row
                    ws.add_data_validation(dv_q)
                    rows_left = pmi_data_end_row - current_row + 1

                # [NEW] Dynamically print headers for the new elements precisely on the first row of each sheet (플랜트 3단 전용)
                if is_triple and current_row == pmi_start_row and len(print_elements) > 3:
                    h_row = pmi_start_row - 4
                    for r_i in range(max(1, pmi_start_row - 5), pmi_start_row):
                        if str(ws.cell(row=r_i, column=8).value).strip().upper() == "NI":
                            h_row = r_i; break
                    for val_k, c_idx in print_elements[3:]:
                        hdr_cell = ws.cell(row=h_row, column=c_idx)
                        self.safe_set_value(ws, hdr_cell.coordinate, val_k, align='center')
                        hdr_cell.font = Font(name=hdr_cell.font.name if hasattr(hdr_cell.font, 'name') else 'Arial', size=hdr_cell.font.size if hasattr(hdr_cell.font, 'size') else 9, bold=True)
                        hdr_cell = ws.cell(row=h_row, column=c_idx)
                        self.safe_set_value(ws, hdr_cell.coordinate, val_k, align='center')
                        hdr_cell.font = Font(name=hdr_cell.font.name if hasattr(hdr_cell.font, 'name') else 'Arial', size=hdr_cell.font.size if hasattr(hdr_cell.font, 'size') else 9, bold=True)

                # [NEW DYNAMIC ALGORITHM] Determine block size based on layout mode and merge logic
                logical_block_size = 1
                if is_triple:
                    base_item = all_extracted_data[data_ptr]
                    base_joint = base_item.get('visual_group_joint', base_item.get('Joint', ''))
                    base_iso = base_item.get('visual_group_iso', base_item.get('ISO', base_item.get('Dwg', '')))

                    # Check consecutive items
                    for ahead_idx in range(data_ptr + 1, min(len(all_extracted_data), data_ptr + rows_left)):
                        ahead_item = all_extracted_data[ahead_idx]
                        ui_merged_j = ahead_item.get('is_merged_joint') == True
                        name_merged_j = ahead_item.get('visual_group_joint', ahead_item.get('Joint', '')) == base_joint
                        joint_ok = ui_merged_j or name_merged_j

                        ui_merged_i = ahead_item.get('is_merged_iso') == True
                        name_merged_i = ahead_item.get('visual_group_iso', ahead_item.get('ISO', ahead_item.get('Dwg', ''))) == base_iso
                        iso_ok = ui_merged_i or name_merged_i

                        if joint_ok and iso_ok:
                            logical_block_size += 1
                        else:
                            break

                # Cap the block size strictly to the available rows remaining on the sheet
                this_block_size = min(logical_block_size, rows_left)
                actual_block_rows = this_block_size
                batch = all_extracted_data[data_ptr : data_ptr + this_block_size]

                # 테두리 및 서식 적용
                d_height = float(self.config.get('ROW_HEIGHT_DATA', 19.5))
                if is_triple:
                    for r_offset in range(actual_block_rows):
                        r = current_row + r_offset
                        rd = ws.row_dimensions[r]; rd.height = d_height
                        for c in range(1, 14):
                            cell = ws.cell(row=r, column=c)
                            l_s = thin_side; r_s = thin_side
                            t_s = medium_side if r == pmi_start_row else thin_side
                            b_s = medium_side if r == pmi_data_end_row else thin_side

                            if c == 1: l_s = medium_side
                            if c == 13: r_s = medium_side

                            if c <= 5: # A-E 수직 병합 구역 내부 선 제거
                                if 1 < c < 5: r_s = Side(style=None); l_s = Side(style=None)
                                elif c == 1: r_s = Side(style=None)
                                elif c == 5: l_s = Side(style=None)

                                if actual_block_rows > 1:
                                    if r_offset == 0: b_s = Side(style=None)
                                    elif r_offset == actual_block_rows - 1: t_s = Side(style=None)
                                    else: t_s = Side(style=None); b_s = Side(style=None)

                            cell.border = Border(left=l_s, right=r_s, top=t_s, bottom=b_s)

                    # A-E 수평/수직 병합 및 데이터 입력 (블록 전체 범위)
                    self.safe_merge_cells(ws, start_row=current_row, start_column=1, end_row=current_row + actual_block_rows - 1, end_column=5)

                    if actual_block_rows > 1:
                        self.safe_merge_cells(ws, start_row=current_row, start_column=6, end_row=current_row + actual_block_rows - 1, end_column=6)

                    if batch:
                        dwg_col = self.col_to_num(self.config.get('PMI_COL_DWG', '1'))
                        joint_col = self.col_to_num(self.config.get('PMI_COL_JOINT', '6'))

                        if dwg_col >= 1:
                            self.safe_set_value(ws, ws.cell(row=current_row, column=dwg_col).coordinate, batch[0].get('Dwg', ''))
                            ws.cell(row=current_row, column=dwg_col).alignment = Alignment(horizontal='center', vertical='center', shrink_to_fit=True)

                        if joint_col >= 1:
                            self.safe_set_value(ws, ws.cell(row=current_row, column=joint_col).coordinate, batch[0].get('Joint', batch[0].get('No', '')))
                            ws.cell(row=current_row, column=joint_col).alignment = Alignment(horizontal='center', vertical='center')

                    for i, item in enumerate(batch):
                        r = current_row + i
                        loc_col = self.col_to_num(self.config.get('PMI_COL_LOC', '7'))
                        if loc_col >= 1:
                            self.safe_set_value(ws, ws.cell(row=r, column=loc_col).coordinate, item.get('Loc', ''), align='center')

                        for val_key, col_idx in print_elements:
                            if col_idx < 1: continue
                            raw_v = item.get(val_key, 0.0)
                            v = self.to_float(raw_v)
                            cell = ws.cell(row=r, column=col_idx)
                            self.safe_set_value(ws, cell.coordinate, v if v > 0 else "", align='center')
                            cell.font = Font(name='바탕', size=9, bold=False)
                            cell.alignment = Alignment(horizontal='center', vertical='center', shrink_to_fit=True)
                            if v > 0: cell.number_format = '0.00'

                        res_col = self.col_to_num(self.config.get('PMI_COL_RES', '13'))
                        if res_col >= 1:
                            self.safe_set_value(ws, ws.cell(row=r, column=res_col).coordinate, item.get('Grade', ''), align='center')
                            cell_l = ws.cell(row=r, column=res_col); cell_l.font = Font(size=8.5); dv_q.add(cell_l)
                else:
                    # 일반 단일행 (1포인트 = 1행, 표준 SIT-PMI-001 서식)
                    for i, item in enumerate(batch):
                        r = current_row + i
                        rd = ws.row_dimensions[r]; rd.height = d_height
                        for c in range(1, 21):
                            cell = ws.cell(row=r, column=c)
                            l_s = thin_side; r_s = thin_side
                            t_s = medium_side if r == pmi_start_row else thin_side
                            b_s = medium_side if r == pmi_data_end_row else thin_side

                            if c == 1: l_s = medium_side
                            if c == 20: r_s = medium_side

                            if 1 <= c <= 4:
                                if 1 < c < 4: r_s = Side(style=None); l_s = Side(style=None)
                                elif c == 1: r_s = Side(style=None)
                                elif c == 4: l_s = Side(style=None)
                            elif 5 <= c <= 6:
                                if c == 5: r_s = Side(style=None)
                                elif c == 6: l_s = Side(style=None)
                            elif 16 <= c <= 19:
                                if 16 < c < 19: r_s = Side(style=None); l_s = Side(style=None)
                                elif c == 16: r_s = Side(style=None)
                                elif c == 19: l_s = Side(style=None)

                            cell.border = Border(left=l_s, right=r_s, top=t_s, bottom=b_s)

                        # 병합: A:D (Item No / Dwg), E:F (Point / Joint), P:S (Date)
                        self.safe_merge_cells(ws, start_row=r, start_column=1, end_row=r, end_column=4)
                        self.safe_merge_cells(ws, start_row=r, start_column=5, end_row=r, end_column=6)
                        self.safe_merge_cells(ws, start_row=r, start_column=16, end_row=r, end_column=19)

                        # A열(Item No. / Dwg)
                        item_dwg = item.get('Dwg', '') or item.get('Item', '')
                        if item_dwg:
                            self.safe_set_value(ws, f'A{r}', item_dwg, align='center')

                        # E열(Point / Joint)
                        point_val = item.get('Joint', '') or item.get('Loc', '')
                        if point_val:
                            self.safe_set_value(ws, f'E{r}', point_val, align='center')

                        # G열(Data Reading No.)
                        reading_no = item.get('No', '')
                        self.safe_set_value(ws, f'G{r}', reading_no, align='center')

                        # H~M열 (Chemical Composition)
                        for val_key, col_idx in print_elements:
                            if col_idx < 1: continue
                            raw_v = item.get(val_key, 0.0)
                            v = self.to_float(raw_v)
                            cell = ws.cell(row=r, column=col_idx)
                            self.safe_set_value(ws, cell.coordinate, v if v > 0 else "", align='center')
                            cell.font = Font(name='바탕', size=9, bold=False)
                            cell.alignment = Alignment(horizontal='center', vertical='center', shrink_to_fit=True)
                            if v > 0: cell.number_format = '0.00'

                        # N열(Result / ALL MATCH)
                        res_val = item.get('Grade', '') or item.get('Result', '')
                        if res_val:
                            self.safe_set_value(ws, f'N{r}', res_val, align='center')
                            ws.cell(row=r, column=14).font = Font(size=8.5)
                            dv_q.add(ws.cell(row=r, column=14))

                        # O열(Examiner)
                        examiner_val = self.config.get('GAPJI_EXAMINER', '') or 'B.H.JIN'
                        self.safe_set_value(ws, f'O{r}', examiner_val, align='center')
                        ws.cell(row=r, column=15).font = Font(size=8.5)

                        # P열(Date: P~S)
                        date_val = item.get('Date', '') or self.gapji_exam_date.get().strip() or self.config.get('GAPJI_EXAM_DATE', '')
                        self.safe_set_value(ws, f'P{r}', date_val, align='center')
                        ws.cell(row=r, column=16).font = Font(size=8.5)

                        # T열(Remark)
                        remark_val = item.get('Remarks', '')
                        if remark_val:
                            self.safe_set_value(ws, f'T{r}', remark_val, align='center')

                data_ptr += len(batch)
                current_row += actual_block_rows
                self.progress['value'] = 30 + (data_ptr / len(all_extracted_data)) * 65

            # [NEW] 마지막 데이터 기입 직후 다음 행에 BLANK 표시 (병합 적용)
            if is_triple:
                try:
                    last_data_r = (current_row - actual_block_rows) + len(batch)
                    if last_data_r <= pmi_data_end_row:
                        self.safe_merge_cells(ws, start_row=last_data_r, start_column=8, end_row=last_data_r, end_column=12)
                        self.safe_set_value(ws, ws.cell(row=last_data_r, column=8).coordinate, "BLANK")
                        ws.cell(row=last_data_r, column=8).alignment = Alignment(horizontal='center', vertical='center')
                        ws.cell(row=last_data_r, column=8).font = Font(size=9, bold=True)
                except: pass

                # 빈 칸도 3개 행 블록 단위로 테두리/병합 적용
                d_height = float(self.config.get('ROW_HEIGHT_DATA', 19.5))
                while current_row <= pmi_data_end_row:
                    rows_left = pmi_data_end_row - current_row + 1
                    this_block_size = min(3, rows_left)

                    for r_offset in range(this_block_size):
                        r = current_row + r_offset
                        rd = ws.row_dimensions[r]; rd.height = d_height
                        for c in range(1, 14):
                            cell = ws.cell(row=r, column=c)
                            l_s = thin_side; r_s = thin_side
                            t_s = medium_side if r == pmi_start_row else thin_side
                            b_s = medium_side if r == pmi_data_end_row else thin_side

                            if c == 1: l_s = medium_side
                            if c == 13: r_s = medium_side

                            if 1 <= c <= 5:
                                if 1 < c < 5: r_s = Side(style=None); l_s = Side(style=None)
                                elif c == 1: r_s = Side(style=None)
                                elif c == 5: l_s = Side(style=None)

                            cell.border = Border(left=l_s, right=r_s, top=t_s, bottom=b_s)

                        self.safe_merge_cells(ws, start_row=r, start_column=1, end_row=r, end_column=5)

                    dv_q.add(ws.cell(row=current_row, column=13))
                    current_row += this_block_size
            else:
                # 단일행 BLANK 표시 및 잔여 빈 행 서식 유지
                try:
                    last_data_r = current_row
                    if last_data_r <= int(self.config.get('DATA_END_ROW', 47)):
                        self.safe_merge_cells(ws, start_row=last_data_r, start_column=8, end_row=last_data_r, end_column=14)
                        self.safe_set_value(ws, ws.cell(row=last_data_r, column=8).coordinate, "BLANK")
                        ws.cell(row=last_data_r, column=8).alignment = Alignment(horizontal='center', vertical='center')
                        ws.cell(row=last_data_r, column=8).font = Font(size=9, bold=True)
                except: pass

                d_height = float(self.config.get('ROW_HEIGHT_DATA', 19.5))
                while current_row <= int(self.config.get('DATA_END_ROW', 47)):
                    r = current_row
                    rd = ws.row_dimensions[r]; rd.height = d_height
                    for c in range(1, 21):
                        cell = ws.cell(row=r, column=c)
                        l_s = thin_side; r_s = thin_side
                        t_s = medium_side if r == int(self.config.get('START_ROW', 17)) else thin_side
                        b_s = medium_side if r == int(self.config.get('DATA_END_ROW', 47)) else thin_side

                        if c == 1: l_s = medium_side
                        if c == 20: r_s = medium_side

                        if 1 <= c <= 4:
                            if 1 < c < 4: r_s = Side(style=None); l_s = Side(style=None)
                            elif c == 1: r_s = Side(style=None)
                            elif c == 4: l_s = Side(style=None)
                        elif 5 <= c <= 6:
                            if c == 5: r_s = Side(style=None)
                            elif c == 6: l_s = Side(style=None)
                        elif 16 <= c <= 19:
                            if 16 < c < 19: r_s = Side(style=None); l_s = Side(style=None)
                            elif c == 16: r_s = Side(style=None)
                            elif c == 19: l_s = Side(style=None)

                        cell.border = Border(left=l_s, right=r_s, top=t_s, bottom=b_s)

                    self.safe_merge_cells(ws, start_row=r, start_column=1, end_row=r, end_column=4)
                    self.safe_merge_cells(ws, start_row=r, start_column=5, end_row=r, end_column=6)
                    self.safe_merge_cells(ws, start_row=r, start_column=16, end_row=r, end_column=19)
                    current_row += 1

            # [FIX] PMI 템플릿의 미사용 시트(28개 시트 중 초과분) 정리
            remove_unused_report_sheets(wb, used_pmi_sheets)

            # [NEW] 서식 정리 도구
            def clear_borders_in_range(sheet, start_r, end_r, start_c=1, end_c=15):
                for r_idx in range(start_r, end_r + 1):
                    for c_idx in range(start_c, end_c + 1):
                        try: sheet.cell(row=r_idx, column=c_idx).border = Border()
                        except: pass

            # [FORCE] 모든 데이터 시트의 바닥선 및 하단 서식 정리
            data_end_row = int(float(self.config.get('DATA_END_ROW', 47)))
            for idx, s in enumerate(wb.worksheets):
                if idx > 0: # 을지 (데이터 시트)
                    # 데이터 영역 아래쪽(60행까지)의 불필요한 선 제거
                    clear_borders_in_range(s, data_end_row + 1, data_end_row + 15, start_c=1, end_c=pmi_max_cols)

                    if s.max_row >= data_end_row:
                        for c in range(1, pmi_max_cols + 1):
                            cell = s.cell(row=data_end_row, column=c)
                            curr_border = cell.border
                            cell.border = Border(
                                left=curr_border.left,
                                right=curr_border.right,
                                top=curr_border.top,
                                bottom=Side(style='medium')
                            )
                else: # 갑지 (Cover 시트: 결재/서명란 40~50행 내부선 정돈 및 훼손 방지)
                    # - 45행: D~I열(Examined by 서명 밑줄)과 K~T열(보고서 작성일자 구분선)만 thin 바닥선 유지
                    # - 41~44행, 46~49행: 검사자/입회자/검토자 내부의 파편화된 바닥선(47행 잔여선 등) 완전 제거
                    for r_sub in range(41, 50):
                        for c in range(1, 21):
                            cell = s.cell(row=r_sub, column=c)
                            cb = cell.border
                            if r_sub == 45:
                                if (4 <= c <= 9) or (11 <= c <= 20):
                                    b_side = Side(style='thin')
                                else:
                                    b_side = Side(style=None)
                            else:
                                b_side = Side(style=None)
                            cell.border = Border(left=cb.left, right=cb.right, top=cb.top, bottom=b_side)

            # [FORCE] 을지 시트 상단 외곽 테두리 제거 (1행 Top 및 A1-A4 Left, T1-T4 Right)
            for idx, s in enumerate(wb.worksheets):
                if idx > 0: # 을지 시트
                    # 1행 상단 테두리 제거 (T열까지 확장)
                    for c_idx in range(1, 21):
                        cell = s.cell(row=1, column=c_idx)
                        b = cell.border
                        cell.border = Border(left=b.left, right=b.right, top=Side(style=None), bottom=b.bottom)

                    # A1-A4 좌측, M1-M4 우측, T1-T4 우측 테두리 제거
                    for r_idx in range(1, 5):
                        # A열 (좌측)
                        cell_a = s.cell(row=r_idx, column=1)
                        b_a = cell_a.border
                        cell_a.border = Border(left=Side(style=None), right=b_a.right, top=b_a.top, bottom=b_a.bottom)

                        # M열 (우측)
                        cell_m = s.cell(row=r_idx, column=13)
                        b_m = cell_m.border
                        cell_m.border = Border(left=b_m.left, right=Side(style=None), top=b_m.top, bottom=b_m.bottom)

                        # T열 (우측)
                        cell_t = s.cell(row=r_idx, column=20)
                        b_t = cell_t.border
                        cell_t.border = Border(left=b_t.left, right=Side(style=None), top=b_t.top, bottom=b_t.bottom)

                    # [NEW] 페이지 번호 구역 모든 테두리 제거
                    if is_triple:
                        for c_idx in [12, 13]: # L, M
                            s.cell(row=3, column=c_idx).border = Border()
                    else:
                        for c_idx in range(14, 21): # N~T
                            s.cell(row=3, column=c_idx).border = Border()

                    # [NEW] 5행 위쪽 선 진한선으로 (A~M/T열 범위)
                    for c_idx in range(1, pmi_max_cols + 1):
                        cell_5 = s.cell(row=5, column=c_idx)
                        b_5 = cell_5.border
                        cell_5.border = Border(left=b_5.left, right=b_5.right, top=medium_side, bottom=b_5.bottom)

            # [FORCE] 을지 시트의 지정 행 글꼴 크기 통일
            for idx, s in enumerate(wb.worksheets):
                if idx > 0: # 을지 시트
                    for r_idx in range(data_start_row, data_end_row + 1):
                        for c_idx in range(1, pmi_max_cols + 1):
                            cell = s.cell(row=r_idx, column=c_idx)
                            f = cell.font
                            if f:
                                cell.font = Font(name=f.name, size=10, bold=f.bold, italic=f.italic, vertAlign=f.vertAlign, underline=f.underline, strike=f.strike, color=f.color)
                            else:
                                cell.font = Font(name='맑은 고딕', size=10)

            total_p = len(wb.worksheets)
            for p_idx, s in enumerate(wb.worksheets):
                page_num = p_idx + 1
                # [NEW] 최종 저장 직전 사용자의 인쇄 설정 및 행/열 커스텀 설정을 모든 시트에 다시 한번 적용 (최우선순위)
                ctx = "COVER" if p_idx == 0 else "DATA"
                self.force_print_settings(s, context=ctx)
                self.apply_custom_dimensions(s, ctx)
                normalize_pmi_header_stamp(s, is_cover=(p_idx == 0))
                # 원본 템플릿 상호 로고(E5:K7) 이미지 보존 (불필요한 텍스트 덮어쓰기 방지)
                self.safe_set_value(s, 'E5', None)

                try:
                    is_triple_mode = hasattr(self, 'pmi_layout_mode') and self.pmi_layout_mode.get() == "PLANT_TRIPLE"
                    if p_idx == 0: # Gapji (Cover)
                        self.safe_merge_cells(s, 3, 14, 3, 20) # N3:T3
                        cell = s["N3"]
                    else: # Eulji (Data)
                        if is_triple_mode:
                            # [13열 플랜트 3단 모드] 우측 상단 L3:M3 병합
                            try: self.safe_merge_cells(s, 3, 12, 3, 13) # L3:M3
                            except: pass
                            cell = s["L3"] if "L3" in s.merged_cells else s["M3"]
                            self.safe_set_value(s, 'O3', None)
                        else:
                            # [20열 단일행 표준 모드] 우측 상단 N3:T3 병합 (정중앙 쏠림 해결 및 기존 중복 병합 해제)
                            for m_rng in list(s.merged_cells.ranges):
                                if m_rng.min_row <= 3 <= m_rng.max_row and m_rng.max_col >= 12 and m_rng.min_col <= 20:
                                    try: s.unmerge_cells(str(m_rng))
                                    except: pass
                            self.safe_set_value(s, 'L3', None)
                            self.safe_set_value(s, 'M3', None)
                            self.safe_set_value(s, 'N3', None)
                            self.safe_set_value(s, 'O3', None)
                            try: self.safe_merge_cells(s, 3, 14, 3, 20) # N3:T3
                            except: pass
                            cell = s["N3"]

                    cell.alignment = Alignment(horizontal='center', vertical='center', shrink_to_fit=True)
                    cell.font = Font(name='맑은 고딕', size=11, bold=True)
                    # 단어 사이의 간격을 인위적으로 늘림
                    self.safe_set_value(s, cell.coordinate, f"Page    {page_num}    of    {total_p}")
                except Exception as e:
                    self.log(f"페이지 번호 기입 실패 ({p_idx}): {e}")

            now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            output_ext = '.xlsm' if keep_vba else '.xlsx'
            save_path = self._report_save_path(os.path.dirname(template_path), output_ext, mode='PMI')
            if not save_path:
                return
            output_name = os.path.basename(save_path)
            wb.save(save_path)
            # [PRESERVE VML FOOTER] 템플릿의 [페이지 설정] -> [머리글/바닥글] 우측 로고(VML &G)를 모든 시트에 복원
            try:
                try:
                    from site_apps.central.src.reports.mt_footer import restore_template_footer_pictures
                except ImportError:
                    from site_apps.central.src.reports.mt_footer import restore_template_footer_pictures
                restore_template_footer_pictures(template_path, save_path)
                self.log("   [OK] 템플릿 머리글/바닥글 및 바닥글 로고(VML) 복원 완료")
            except Exception as fe:
                self.log(f"   [WARN] 바닥글 로고 복원 스킵 또는 오류: {fe}")
            self._record_report_issue('PMI', wb, save_path, final_list)
            self.progress['value'] = 100; self.log(f"✨ 완료! 저장됨: {output_name}")
            messagebox.showinfo("성공", f"통합 성적서 생성이 완료되었습니다.\n\n경로: {os.path.dirname(save_path)}\n파일명: {output_name}")
        except Exception as e:
            self.log(f"❌ 오류: {e}")
            traceback.print_exc()
            # [NEW] 사용자에게 친절한 팝업 알림 추가
            error_msg = f"성적서 생성 중 오류가 발생했습니다.\n\n원인: {e}\n\n* 주로 양식 파일의 시트 구조나 병합 상태가 맞지 않을 때 발생합니다. 다른 양식을 시도하거나 설정을 확인해주세요."
            messagebox.showerror("생성 실패", error_msg)
            # End of _run_pmi_process logic
        finally:
            if 'wb' in locals() and wb:
                try:
                    if hasattr(wb, 'vba_archive') and wb.vba_archive:
                        wb.vba_archive.close()
                except: pass
                try: wb.close()
                except: pass
            for f in glob.glob(os.path.join(tempfile.gettempdir(), "temp_*.png")):
                try: os.remove(f)
                except: pass

    def _run_pt_process(self, final_list, template_path):
        """PT 성적서 생성 (1-row-per-data 레이아웃 + ISO 병합)"""
        self.save_settings() # Ensure UI -> config sync
        self.log(f"🚀 PT 성적서 생성 시작 (총 {len(final_list)} 건)...")
        self.progress['value'] = 0
        try:
            wb = openpyxl.load_workbook(template_path, keep_vba=False)
            if len(wb.worksheets) < 1:
                raise ValueError("선택한 템플릿 파일에 시트가 존재하지 않습니다.")

            # 갑지 (Cover) - PT 템플릿은 이미 갑지 로고가 포함되어 있으므로 로고 삽입 생략
            ws0 = wb.worksheets[0]
            used_pt_sheets = [ws0]
            # self.add_logos_to_sheet(ws0, is_cover=True, clear_existing=False, mode="PT")  # 템플릿 원본 유지
            self._write_gapji_metadata(ws0, mode="PT")
            thickness_range = format_thickness_range(final_list)
            if thickness_range:
                self.safe_set_value(ws0, 'D11', thickness_range)
            exam_date = first_report_date([{"Date": self.gapji_exam_date.get()}])
            if exam_date:
                # PT cover: Date of Examination comes from the editable
                # 검사일자 field, not from the source file's Request Date.
                # Keep it as an Excel date so the template's Korean date
                # display format is retained.
                self.safe_set_value(ws0, 'J40', exam_date)
            self.force_print_settings(ws0, context="COVER")
            self.apply_custom_dimensions(ws0, "COVER")

            # 을지/갑지 자동 감지 - 헤더가 있는 시트에 데이터 기입
            data_sheet_id = 1 if len(wb.worksheets) >= 2 else 0

            # 헤더 감지 및 시작 행 결정
            start_row = int(self.config.get('PT_START_ROW', 18))
            end_row = int(self.config.get('PT_END_ROW', 38))

            # 스타일 설정
            thin_side = Side(style='thin')
            thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

            def _detect_pt_header(sheet):
                """시트에서 데이터 헤더 행을 찾아 데이터 시작 행 반환"""
                for r in range(1, 45):
                    vals = [str(sheet.cell(row=r, column=c).value).upper() for c in range(1, 22)]
                    s = "".join(vals)
                    if ("JOINT" in s or "WELD" in s) and ("ISO" in s or "DWG" in s or "LINE" in s):
                        return r + 1
                    if "DWG" in s and "RESULT" in s:
                        return r + 1
                    if "NO." in s and "JOINT" in s and ("ACC" in s or "REJ" in s):
                        return r + 1
                return -1

            # 1) 갑지(첫 번째 시트)에서 먼저 헤더 감지
            d_cover = _detect_pt_header(ws0)
            if d_cover > 0:
                data_sheet_id = 0
                ws = ws0
                start_row = d_cover
                self.log(f"[PT] 갑지(1번 시트) {start_row}행에서 데이터 헤더 감지 → 갑지에 기입")
            else:
                # 2) 을지(두 번째 시트)에서 감지
                ws = wb.worksheets[data_sheet_id]
                d_data = _detect_pt_header(ws)
                if d_data > 0:
                    start_row = d_data
                    self.log(f"[PT] 을지(2번 시트) {start_row}행에서 데이터 헤더 감지 → 을지에 기입")
                else:
                    self.log(f"[PT] 헤더 자동 감지 실패 → 기본 start_row={start_row} 사용")

            if ws not in used_pt_sheets:
                used_pt_sheets.append(ws)

            ws.title = f"{ws.title[:20]}_001"
            # 갑지에 데이터가 기록되더라도 을지의 여백/배율/행·열 설정이
            # 갑지로 새어 들어가지 않도록 설정 컨텍스트를 완전히 분리한다.
            if ws is not ws0:
                self.force_print_settings(ws, context="DATA")
                self.apply_custom_dimensions(ws, "DATA")

            # PT 을지 원본(두 번째 시트)의 실제 레이아웃을 저장한다.
            # openpyxl 복제 과정에서 암시적 기본 행/열 크기가 달라지는 것을 막기 위해
            # A~S 열과 1~47행의 유효 크기를 모든 후속 을지에 명시적으로 재적용한다.
            pt_layout_source = wb.worksheets[1] if len(wb.worksheets) >= 2 else ws
            self.force_print_settings(pt_layout_source, context="DATA")
            self.apply_custom_dimensions(pt_layout_source, "DATA")
            pt_default_col_width = pt_layout_source.sheet_format.defaultColWidth or 8.38
            pt_default_row_height = pt_layout_source.sheet_format.defaultRowHeight or 15.0
            pt_column_widths = {}
            for col_idx in range(1, 20):
                col_letter = openpyxl.utils.get_column_letter(col_idx)
                if col_letter in pt_layout_source.column_dimensions:
                    pt_column_widths[col_letter] = pt_layout_source.column_dimensions[col_letter].width
                else:
                    pt_column_widths[col_letter] = pt_default_col_width
            pt_row_heights = {
                row_idx: (
                    pt_layout_source.row_dimensions[row_idx].height
                    if row_idx in pt_layout_source.row_dimensions
                    and pt_layout_source.row_dimensions[row_idx].height is not None
                    else pt_default_row_height
                )
                for row_idx in range(1, 48)
            }

            # [FIX] PT 001 시트에도 고객사, 리포트 번호 명시적 주입
            try:
                self.safe_set_value(ws, 'P4', self.gapji_customer.get().strip())
                self.safe_set_value(ws, 'P6', self.gapji_report_no.get().strip())
            except: pass

            # 갑지는 38행이 Sketch 고정행이므로 37행까지만 데이터 기입
            # 001시트(을지)는 10~38행까지 데이터 기입
            cover_end_row = int(self.config.get('PT_COVER_END_ROW', 37))  # 갑지 마지막 데이터행
            data_end_row = int(self.config.get('PT_END_ROW', 38))          # 001시트 마지막 데이터행
            # 001시트 원본은 항상 index=1 (갑지=0에 데이터를 쓰는 경우에도 복사 원본은 2번째 시트)
            template_sheet_id = 1 if len(wb.worksheets) >= 2 else 0

            # 갑지에 쓰는 경우 end_row를 커버용으로 분리
            on_cover = (data_sheet_id == 0)
            current_end_row = cover_end_row if on_cover else data_end_row

            current_row = start_row
            current_page = 1
            data_ptr = 0

            while data_ptr < len(final_list):
                if current_row > current_end_row:
                    current_page += 1
                    if on_cover and len(wb.worksheets) >= 2:
                        ws = wb.worksheets[1]
                        ws.title = f"{ws.title[:20]}_{current_page:03d}"
                        try:
                            self.force_print_settings(ws, context="DATA")
                            self.apply_custom_dimensions(ws, "DATA")
                        except: pass
                    else:
                        ws = self.prepare_next_sheet(wb, template_sheet_id, current_page, mode="PT")
                    if ws not in used_pt_sheets:
                        used_pt_sheets.append(ws)

                    current_row = int(self.config.get('PT_START_ROW', 10))
                    current_end_row = int(self.config.get('PT_DATA_END_ROW', 38))
                    on_cover = False

                item = final_list[data_ptr]

                # PT 고정 열 하드코딩 (원본 병합셀 보호)
                self.safe_set_value(ws, ws.cell(row=current_row, column=1).coordinate, item.get('Dwg', ''))
                try: self.safe_merge_cells(ws, start_row=current_row, start_column=1, end_row=current_row, end_column=3)
                except: pass
                self.safe_set_value(ws, ws.cell(row=current_row, column=4).coordinate, item.get('Joint', ''))

                res_val = str(item.get('Result', 'Acc')).strip().upper()
                is_acc = (res_val in ['ACC', 'ACCEPT', 'PASS', '합격', 'O', 'OK', 'V', ''])

                if is_acc:
                    self.safe_set_value(ws, ws.cell(row=current_row, column=5).coordinate, 'V')
                    self.safe_set_value(ws, ws.cell(row=current_row, column=8).coordinate, 'NO RECORDABLE INDICATION')
                else:
                    self.safe_set_value(ws, ws.cell(row=current_row, column=6).coordinate, 'V')
                    self.safe_set_value(ws, ws.cell(row=current_row, column=8).coordinate, res_val)

                self.safe_set_value(ws, ws.cell(row=current_row, column=18).coordinate, item.get('Welder', ''))
                self.safe_set_value(ws, ws.cell(row=current_row, column=19).coordinate, item.get('NPS', ''))

                # 스타일링 생략 (병합셀 손상 방지)

                data_ptr += 1
                current_row += 1
                self.progress['value'] = (data_ptr / len(final_list)) * 95

            # 마지막 데이터 다음 빈 행 하나에 TOTAL 요약을 추가한다.
            # 갑지 38행은 Sketch/Attached 고정행이므로 절대 덮어쓰지 않는다.
            pt_print_end_row = int(self.config.get('PT_PRINT_END_ROW', 40))
            summary_end_row = cover_end_row if on_cover else pt_print_end_row
            if current_row > summary_end_row:
                current_page += 1
                ws = self.prepare_next_sheet(wb, template_sheet_id, current_page, mode="PT")
                if ws not in used_pt_sheets:
                    used_pt_sheets.append(ws)
                current_row = int(self.config.get('PT_START_ROW', 10))
            self._write_pt_total_summary(ws, current_row, final_list)

            # The PT template contains a ready-made data sheet.  When all
            # records fit on the cover, that sheet remains unused and must not
            # inflate the page count or appear as a blank second page.
            remove_unused_report_sheets(wb, used_pt_sheets)
            total_p = len(wb.worksheets)
            # 서식 정리 (병합셀 손상 방지로 제거)
            for p_idx, s in enumerate(wb.worksheets):
                page_num = p_idx + 1
                # Keep the current cover metadata consistent on every PT
                # continuation sheet, including template and copied sheets.
                self.safe_set_value(s, 'P6', self.gapji_report_no.get().strip())
                # 최종 저장 직전 모든 PT 시트에 UI 설정을 동일하게 재적용한다.
                # Area가 비어 있으면 force_print_settings의 PT 기본 영역을 사용한다.
                ctx = "COVER" if p_idx == 0 else "DATA"
                self.force_print_settings(s, context=ctx)
                self.apply_custom_dimensions(s, ctx)
                if p_idx > 0:
                    for col_letter, width in pt_column_widths.items():
                        s.column_dimensions[col_letter].width = width
                    for row_idx, height in pt_row_heights.items():
                        s.row_dimensions[row_idx].height = height
                # 페이지 번호 기입
                try:
                    p_text = f"Page    {page_num}    of    {total_p}"
                    if p_idx > 0: self.safe_set_value(s, 'P3', p_text)
                    else: self.safe_set_value(s, 'P3', p_text)
                except: pass

            now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            output_name = os.path.basename(self._report_save_path(os.path.dirname(template_path), '.xlsx', choose_directory=False))
            remembered_dir = str(self.config.get('PT_LAST_SAVE_DIR', '')).strip()
            initial_dir = (
                remembered_dir
                if remembered_dir and os.path.isdir(remembered_dir)
                else os.path.dirname(template_path)
            )
            save_path = filedialog.asksaveasfilename(
                parent=self.root,
                title="PT 성적서 저장 위치 선택",
                initialdir=initial_dir,
                initialfile=output_name,
                defaultextension=".xlsx",
                filetypes=[("Excel Workbook", "*.xlsx")],
            )
            if not save_path:
                self.log("PT 성적서 저장이 취소되었습니다.")
                return

            self.config['PT_LAST_SAVE_DIR'] = os.path.dirname(save_path)
            self.save_settings()
            wb.save(save_path)
            self._record_report_issue('PT', wb, save_path, final_list)
            self.progress['value'] = 100
            self.log(f"✨ PT 완료! 저장됨: {os.path.basename(save_path)}")
            messagebox.showinfo(
                "성공",
                f"PT 성적서 생성이 완료되었습니다.\n경로: {save_path}",
            )
        except Exception as e:
            self.log(f"❌ PT 생성 오류: {e}")
            traceback.print_exc()
            messagebox.showerror("생성 실패", f"PT 성적서 생성 중 오류가 발생했습니다: {e}")
        finally:
            if 'wb' in locals() and wb:
                try:
                    if hasattr(wb, 'vba_archive') and wb.vba_archive:
                        wb.vba_archive.close()
                except: pass
                try: wb.close()
                except: pass
            for f in glob.glob(os.path.join(tempfile.gettempdir(), "temp_*.png")):
                try: os.remove(f)
                except: pass


    def _apply_mt_signatures(self, ws):
        from openpyxl.drawing.image import Image as SignatureImage
        from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
        from openpyxl.drawing.xdr import XDRPositiveSize2D
        from openpyxl.utils.units import pixels_to_EMU
        for name_cell, row in (('D39', 39), ('D42', 42)):
            name = re.sub(r'\s+', '', str(ws[name_cell].value or ''))
            if not name:
                continue
            path = os.path.join(SRC_ROOT, 'signs', name + '.png')
            if not os.path.isfile(path):
                self.log(f'서명 파일 없음: {name}')
                continue
            stream = self._white_background_to_transparent_png(path)
            image = SignatureImage(stream)
            image._pmi_source_stream = stream
            scale = min(48 / image.width, 40 / image.height)
            width, height = image.width * scale, image.height * scale
            image.anchor = OneCellAnchor(
                _from=AnchorMarker(col=5, row=row - 1, colOff=pixels_to_EMU(2),
                                   rowOff=pixels_to_EMU((48 - height) / 2)),
                ext=XDRPositiveSize2D(pixels_to_EMU(width), pixels_to_EMU(height)))
            ws.add_image(image)

    def _run_mt_process(self, final_list, template_path):
        """MT 성적서 생성 (1-row-per-data 레이아웃 + ISO 병합)"""
        self.save_settings() # Ensure UI -> config sync
        self.log(f"🚀 MT 성적서 생성 시작 (총 {len(final_list)} 건)...")
        self.progress['value'] = 0
        try:
            from site_apps.central.src.reports.mt_report_layout import hot_tapping_points
            report_rows = hot_tapping_points(final_list) if self.gapji_item.get().strip().upper() == 'HOT TAPPING' else final_list
            keep_vba = os.path.splitext(template_path)[1].lower() == '.xlsm'
            wb = openpyxl.load_workbook(template_path, keep_vba=keep_vba)
            if len(wb.worksheets) < 1:
                raise ValueError("선택한 템플릿 파일에 시트가 존재하지 않습니다.")

            # 갑지 (Cover)
            ws0 = wb.worksheets[0]
            from site_apps.central.src.reports.mt_report_layout import fit_equipment_labels
            fit_equipment_labels(ws0)
            self._write_gapji_metadata(ws0, mode="MT")
            temperature = self.mt_surface_temp.get().strip()
            if temperature:
                try:
                    temperature = float(temperature)
                except ValueError:
                    messagebox.showerror('온도 확인', '표면온도는 숫자로 입력해 주세요. 예: 25')
                    return
            self.safe_set_value(ws0, 'B21', temperature if temperature != '' else None)
            ws0['B21'].alignment = Alignment(horizontal='center', vertical='center')
            self.safe_set_value(ws0, 'D21', '■')
            self.safe_set_value(ws0, 'D22', '□')
            line_numbers = list(dict.fromkeys(str(item.get('Dwg', '') or '').strip()
                                             for item in final_list if str(item.get('Dwg', '') or '').strip()))
            self.safe_set_value(ws0, 'Q9', '\n'.join(line_numbers))
            ws0['Q9'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            if len(line_numbers) > 1:
                ws0.row_dimensions[9].height = max(ws0.row_dimensions[9].height or 15, 15 * len(line_numbers))
            self.safe_set_value(ws0, 'L39', self.gapji_exam_date.get().strip())
            ws0['L39'].alignment = Alignment(horizontal='center', vertical='center')

            # 을지 (Data)
            data_sheet_id = 1 if len(wb.worksheets) >= 2 else 0
            ws = wb.worksheets[data_sheet_id]
            ws.title = f"{ws.title[:20]}_001"
            # [FIX] Do NOT clear existing images if data sheet is the same as cover

            # [NEW] 지역난방 MT 맞춤형 페이징 및 행 설정
            gapji_start_row = 29
            gapji_end_row = 36
            eulji_start_row = 10
            eulji_end_row = 38

            # 스타일 설정
            thin_side = Side(style='thin')
            thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

            current_row = gapji_start_row
            current_page = 1
            data_ptr = 0
            is_gapji = True
            from site_apps.central.src.reports.mt_report_layout import mt_data_positions
            available = 8 - len(report_rows) if len(report_rows) < 8 else 29 - ((len(report_rows) - 8) % 29)
            summary_index = len(report_rows) + (available if available < 3 else 0)
            positions = mt_data_positions(wb, summary_index + 3,
                lambda page: self.prepare_next_sheet(wb, data_sheet_id, page, mode='MT'))
            while data_ptr < len(report_rows):
                ws, current_row = positions[data_ptr]

                item = report_rows[data_ptr]

                self.safe_set_value(ws, f'B{current_row}', item.get('Joint', ''))

                # MT acceptance check and classification.
                result_val = str(item.get('Result', '')).strip().upper()
                acc_cell = ws.cell(row=current_row, column=6)
                rej_cell = ws.cell(row=current_row, column=8)
                self.safe_set_value(ws, f'K{current_row}', '')
                if result_val in ('합격', 'ACC', 'ACCEPT', 'ACCEPTED'):
                    self.safe_set_value(ws, acc_cell.coordinate, "√")
                    self.safe_set_value(ws, rej_cell.coordinate, "")
                    self.safe_set_value(ws, f'K{current_row}', 1)
                elif result_val in ('불합격', 'REJ', 'REJECT', 'REJECTED'):
                    self.safe_set_value(ws, acc_cell.coordinate, "")
                    self.safe_set_value(ws, rej_cell.coordinate, "√")
                else:
                    self.safe_set_value(ws, acc_cell.coordinate, "")
                    self.safe_set_value(ws, rej_cell.coordinate, "")

                # Interpretation is separate from the welder identity.
                interp_cell = ws.cell(row=current_row, column=12)
                interpretation = item.get('Interpretation', '')
                if not str(interpretation or '').strip() and result_val in ('합격', 'ACC', 'ACCEPT', 'ACCEPTED'):
                    interpretation = '관련지시 없음'
                self.safe_set_value(ws, interp_cell.coordinate, interpretation)

                # [MT] Remarks (Col 21) 에는 비고
                remarks_cell = ws.cell(row=current_row, column=21)
                remarks = str(item.get('Remarks', '') or '').strip()
                welder = str(item.get('Welder', '') or '').strip()
                self.safe_set_value(ws, remarks_cell.coordinate,
                                    '\n'.join(value for value in (remarks, welder) if value))
                remarks_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                if remarks and welder:
                    ws.row_dimensions[current_row].height = max(ws.row_dimensions[current_row].height or 15, 30)

                # 스타일링 (가운데 정렬)
                for c in [1, 2, 6, 8, 11, 12]:
                    cell = ws.cell(row=current_row, column=c)
                    cell.alignment = Alignment(horizontal='center', vertical='center')

                data_ptr += 1
                current_row += 1
                self.progress['value'] = (data_ptr / len(report_rows)) * 95

            marker_sheet, marker_row = positions[summary_index]
            from site_apps.central.src.reports.mt_report_layout import mt_length_totals
            from openpyxl.styles import PatternFill
            original, repair, missing = mt_length_totals(final_list)
            for offset, (label, value) in enumerate((('Original', original), ('Repair', repair), ('Total', original + repair))):
                row = marker_row + offset
                marker_sheet.unmerge_cells(f'L{row}:T{row}')
                marker_sheet.merge_cells(f'L{row}:Q{row}')
                marker_sheet.merge_cells(f'R{row}:T{row}')
                self.safe_set_value(marker_sheet, f'R{row}', label)
                self.safe_set_value(marker_sheet, f'U{row}', '확인 필요' if missing else float(value))
                marker_sheet[f'U{row}'].number_format = '0.0000" M"'
                for col in ('R', 'U'):
                    cell = marker_sheet[f'{col}{row}']
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.font = Font(name='맑은 고딕', size=10, bold=label == 'Total')
                for column in range(18, 25):
                    cell = marker_sheet.cell(row, column)
                    border = copy.copy(cell.border)
                    border.top = Side(style='thin', color='666666')
                    if label == 'Total':
                        border.bottom = Side(style='thin', color='666666')
                        cell.fill = PatternFill('solid', fgColor='E7E6E6')
                    cell.border = border
                marker_sheet.row_dimensions[row].height = max(marker_sheet.row_dimensions[row].height or 15, 24)
            self.safe_set_value(marker_sheet, f'L{marker_row}', '- 이 하 여 백 -')
            marker_sheet[f'L{marker_row}'].alignment = Alignment(horizontal='center', vertical='center')
            marker_sheet[f'L{marker_row}'].font = Font(name='맑은 고딕', size=10, bold=True)
            total_p = len(wb.worksheets)
            for p_idx, s in enumerate(wb.worksheets):
                page_num = p_idx + 1
                s.print_area = 'A1:X44' if p_idx == 0 else 'A1:X39'
                s.page_setup.orientation = 'portrait'
                s.page_setup.paperSize = s.PAPERSIZE_A4
                s.page_setup.scale = None
                s.page_setup.fitToWidth = 1
                s.page_setup.fitToHeight = 1
                s.sheet_properties.pageSetUpPr.fitToPage = True
                self.safe_set_value(s, 'S4', ws0['S4'].value)
                self.safe_set_value(s, 'S6', ws0['S6'].value)
                # Template: Page [V3] of [X3].
                self.safe_set_value(s, 'V3', page_num)
                self.safe_set_value(s, 'X3', total_p)
                for coordinate in ('V3', 'X3'):
                    s[coordinate].alignment = Alignment(horizontal='center', vertical='center')

            now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            output_ext = '.xlsm' if keep_vba else '.xlsx'
            save_path = self._report_save_path(os.path.dirname(template_path), output_ext, mode='MT')
            if not save_path:
                return
            output_name = os.path.basename(save_path)
            self._apply_mt_signatures(ws0)
            stamp_data = self._get_photo_log_stamp()
            if stamp_data:
                from site_apps.central.src.reports.mt_report_layout import apply_header_stamp
                for sheet in wb.worksheets:
                    apply_header_stamp(sheet, stamp_data)
            else:
                self.log('⚠️ MT 직인 삽입 실패: PAUT 양식의 직인을 확인해 주세요.')
            wb.save(save_path)
            from site_apps.central.src.reports.mt_footer import restore_mt_footer_pictures
            restore_mt_footer_pictures(template_path, save_path)
            self._record_report_issue('MT', wb, save_path, final_list)
            self.progress['value'] = 100
            self.log(f"✨ MT 완료! 저장됨: {output_name}")
            messagebox.showinfo("성공", f"MT 성적서 생성이 완료되었습니다.\n경로: {os.path.dirname(save_path)}")
        except Exception as e:
            self.log(f"❌ MT 생성 오류: {e}")
            traceback.print_exc()
            messagebox.showerror("생성 실패", f"MT 성적서 생성 중 오류가 발생했습니다: {e}")
        finally:
            if 'wb' in locals() and wb:
                try:
                    if hasattr(wb, 'vba_archive') and wb.vba_archive:
                        wb.vba_archive.close()
                except: pass
                try: wb.close()
                except: pass
            for f in glob.glob(os.path.join(tempfile.gettempdir(), "temp_*.png")):
                try: os.remove(f)
                except: pass

    def _run_rt_process(self, final_list, template_path, mode="RT"):
        """RT 성적서 생성 (1-row-per-data 레이아웃)"""
        # [CRITICAL] Force synchronize UI variables to config before generation
        self.save_settings() 

        self.log(f"🚀 {mode} 성적서 생성 시작 (총 {len(final_list)} 건)...")
        self.progress['value'] = 0
        try:
            keep_vba = os.path.splitext(template_path)[1].lower() == '.xlsm'
            wb = openpyxl.load_workbook(template_path, keep_vba=keep_vba)
            if len(wb.worksheets) < 1:
                raise ValueError("선택한 템플릿 파일에 시트가 존재하지 않습니다.")

            self.log(f"?? [DEBUG] {mode} 갑지 로고 삽입 시도 중...")
            # Gapji (Cover) Logic
            ws0 = wb.worksheets[0]
            self.add_logos_to_sheet(ws0, is_cover=True, clear_existing=False, mode=mode)
            self._write_gapji_metadata(ws0, mode=mode, first_item=final_list[0] if final_list else None)
            self.force_print_settings(ws0, context="COVER")
            self.apply_custom_dimensions(ws0, f"{mode}_COVER")

            if mode == "RT":
                from site_apps.central.src.reports.rt_report_layout import fit_film_headings
                fit_film_headings(ws0)
                first_item = final_list[0] if final_list else {}
                # Cover inspection/film information.
                cover_size = str(first_item.get("Size", "")).strip()
                cover_thickness = str(first_item.get("T", "")).strip()
                if cover_size and cover_size.lower() != "nan":
                    self.safe_set_value(ws0, "C9", cover_size)
                    size_number_match = re.search(r"\d+(?:\.\d+)?", cover_size)
                    if size_number_match:
                        size_number = float(size_number_match.group(0))
                        self.safe_set_value(
                            ws0, "C26",
                            int(size_number) if size_number.is_integer() else size_number
                        )
                if cover_thickness and cover_thickness.lower() != "nan":
                    self.safe_set_value(ws0, "E9", cover_thickness)
                # Penetrating thickness is the double wall plus 2 mm.
                thickness_match = re.fullmatch(
                    r'\s*(\d+(?:\.\d+)?)\s*(?:mm)?\s*', cover_thickness, re.IGNORECASE
                )
                if thickness_match:
                    penetrating_thickness = float(thickness_match.group(1)) * 2 + 2
                    self.safe_set_value(ws0, "C16", penetrating_thickness)
                    ws0['C16'].number_format = '0.00"mm"'
                else:
                    ws0['C16'] = None
                    self.log('투과 두께 미입력: 첫 번째 검사 항목의 두께를 확인해주세요.')

                raw_film_type = self.rt_film_type.get().strip() or "I"
                base_film_type = re.sub(r"\s*\((?:I|II)\)\s*$", "", raw_film_type, flags=re.IGNORECASE)
                formatted_film_type = (
                    f"{base_film_type}(II)"
                    if base_film_type.upper() == "AA400"
                    else f"{base_film_type}(I)"
                )
                cover_film_values = {
                    "Q9": self.rt_film_brand.get().strip() or "FUJI#80",
                    "V9": formatted_film_type,
                    "Q10": self.rt_film_size.get().strip() or "3⅓ X 12",
                }
                for cell_ref, film_value in cover_film_values.items():
                    self.safe_set_value(ws0, cell_ref, film_value)

                selected_isotope = self.rt_isotope.get().strip() or "Ir-192"
                self.safe_set_value(ws0, "I10", "[V]" if selected_isotope == "Ir-192" else "[ ]")
                self.safe_set_value(ws0, "J10", "Ir-192")
                self.safe_set_value(ws0, "I11", "[V]" if selected_isotope == "Se-75" else "[ ]")
                self.safe_set_value(ws0, "J11", "SE-75")
                self.safe_set_value(ws0, "K10", self.rt_activity.get().strip())
                self.safe_set_value(ws0, "M10", self.rt_activity_unit.get().strip() or "Ci")
                self.safe_set_value(ws0, "C24", self.rt_exposure_time.get().strip())
                self.safe_set_value(ws0, "R20", "SIS-N-106 KDHC-JY")

                # Use the current input, preserving date ranges instead of
                # silently reducing them to their first date.
                exam_date_value = self.gapji_exam_date.get().strip()
                date_parts = None
                if isinstance(exam_date_value, (datetime.datetime, datetime.date)):
                    date_parts = (
                        exam_date_value.year, exam_date_value.month, exam_date_value.day
                    )
                else:
                    date_match = re.fullmatch(
                        r"\s*(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})\s*",
                        str(exam_date_value or "")
                    )
                    if date_match:
                        date_parts = tuple(int(part) for part in date_match.groups())
                formatted_exam_date = str(exam_date_value or '')
                if date_parts:
                    year, month, day = date_parts
                    formatted_exam_date = f"{year} 년   {month:02d} 월   {day:02d} 일"
                self.safe_set_value(ws0, "I28", formatted_exam_date)
                from site_apps.central.src.reports.rt_report_layout import report_date_from_exam_period
                self.safe_set_value(ws0, "I30", report_date_from_exam_period(exam_date_value)
                                    if not date_parts else formatted_exam_date)
                ws0["I28"].alignment = Alignment(
                    horizontal='center', vertical='center', shrink_to_fit=True
                )
                ws0["I30"].alignment = Alignment(
                    horizontal='center', vertical='center', shrink_to_fit=True
                )

                first_sec = str(first_item.get("Sec", "")).strip()
                first_drawing_no = str(first_item.get("Dwg", "")).strip()
                sec_drawing_value = "/".join(
                    value for value in (first_sec, first_drawing_no) if value
                )
                if sec_drawing_value:
                    self.safe_set_value(ws0, "I5", sec_drawing_value)

                # 갑지 검사자/책임검사자 서명란 고정값
                signature_values = {
                    "B32": "대리",
                    "D32": "우명광",
                    "B35": "부장",
                    "D35": "주진철",
                }
                for cell_ref, signature_value in signature_values.items():
                    ws0[cell_ref] = signature_value
                    ws0[cell_ref].alignment = Alignment(
                        horizontal='center', vertical='center',
                        wrap_text=True, shrink_to_fit=True
                    )

                # A8:B9 폭에 맞춰 영문/한글 항목명을 2줄로 모두 표시한다.
                material_size_label = ws0["A8"]
                material_size_label.value = "MAT'L THK./PIPE DIA.\n부재두께 및 관경"
                material_size_label.font = Font(name='돋움체', size=8)
                material_size_label.alignment = Alignment(
                    horizontal='left', vertical='center',
                    wrap_text=True, shrink_to_fit=True
                )

                for cell_ref, label in (
                    ('P1', 'CUSTOMER\n발주자'),
                    ('P2', 'PROJECT\n공사명'),
                    ('P5', 'DRAWING NO.\n도면번호'),
                ):
                    ws0[cell_ref] = label
                    ws0[cell_ref].alignment = Alignment(
                        horizontal='left', vertical='center',
                        wrap_text=True, shrink_to_fit=True
                    )

                # 갑지에서 자동 입력되는 값 영역만 상하좌우 중앙 정렬한다.
                for cell_ref in ("Q1", "Q2", "C5", "I5", "C10", "I28"):
                    ws0[cell_ref].alignment = Alignment(
                        horizontal='center', vertical='center',
                        wrap_text=(cell_ref != "I5"), shrink_to_fit=True
                    )

            # 갑지 제목 영역을 을지(F1:L4)와 같은 4행 높이로 구성한다.
            # 기존 H1:O3 및 K4:L4 병합을 해제한 뒤 H1:O4로 통합한다.
            cover_title_bounds = (8, 1, 15, 4)
            for merged_range in list(ws0.merged_cells.ranges):
                min_c, min_r, max_c, max_r = merged_range.bounds
                if not (max_c < cover_title_bounds[0] or
                        min_c > cover_title_bounds[2] or
                        max_r < cover_title_bounds[1] or
                        min_r > cover_title_bounds[3]):
                    ws0.unmerge_cells(str(merged_range))

            ws0.merge_cells("H1:O4")
            cover_title = ws0["H1"]
            cover_title.value = "방사선투과검사 보고서\nRADIOGRAPHIC EXAMINATION REPORT"
            cover_title.alignment = Alignment(
                horizontal='center', vertical='center',
                wrap_text=True, shrink_to_fit=True
            )
            cover_title.font = Font(name='돋움체', size=14, bold=True)

            # Data Sheet Logic (1-row-per-data)
            data_sheet_id = 1 if len(wb.worksheets) >= 2 else 0
            ws = wb.worksheets[data_sheet_id]
            ws.title = f"{ws.title[:20]}_001"
            # [FIX] Do NOT clear existing images if data sheet is the same as cover
            self.add_logos_to_sheet(ws, is_cover=False, clear_existing=(ws != ws0), mode=mode)
            self.force_print_settings(ws, context="DATA")
            self.apply_custom_dimensions(ws, f"{mode}_DATA")

            # 을지 제목은 원본 병합영역(F1:L4)을 유지하고 한글/영문 2줄로 고정한다.
            if ws != ws0 and mode == "RT":
                title_cell = ws["F1"]
                title_cell.value = "방사선투과검사 보고서\nRADIOGRAPHIC EXAMINATION REPORT"
                title_cell.alignment = Alignment(
                    horizontal='center', vertical='center',
                    wrap_text=True, shrink_to_fit=True
                )
                title_cell.font = Font(name='돋움체', size=14, bold=True)

                # 을지 검사 조건 헤더는 사용자 데이터와 무관한 고정값이다.
                fixed_headers = {
                    "N7": "SENSITIVITY\n식별도",
                    "O7": "DENSITY\n농도",
                    "R7": "TECHNIQUE\n촬영방법",
                }
                for cell_ref, header_text in fixed_headers.items():
                    header_cell = ws[cell_ref]
                    header_cell.value = header_text
                    header_cell.alignment = Alignment(
                        horizontal='center', vertical='center',
                        wrap_text=True, shrink_to_fit=True
                    )
            # RT usually doesn't have the same headers as PMI
            # self.set_eulji_headers(ws) 

            # Configurable or Default RT boundaries
            is_kogas = (mode == "KOGAS")
            start_row = int(self.config.get(f'{mode}_START_ROW', 14 if is_kogas else 11))
            end_row = int(self.config.get(f'{mode}_DATA_END_ROW', 25 if is_kogas else 34))
            block_size = 1

            current_row = start_row
            current_page = 1
            data_ptr = 0
            rt_condition_data_rows = set()
            rt_summary_items = {}
            rt_film_location_counts = {}
            cover_d32_value = ws0["D32"].value if mode == "RT" else None
            if mode == "RT":
                from site_apps.central.src.reports.ndt_history_import import rt_report_quantities
                report_original_count, report_repair_count, fallback_rows = rt_report_quantities(final_list)
                if fallback_rows:
                    self.log(f'수량 주의: {fallback_rows}개 항목에 일보 수량이 없어 구간정보로 계산했습니다. 일보를 다시 불러오면 RT_OR/RT_RE 합계가 적용됩니다.')

            while data_ptr < len(final_list):
                # 페이지 넘김 체크: 가스공사는 2행 블록 기준
                if current_row + block_size - 1 > end_row:
                    current_page += 1
                    ws = self.prepare_next_sheet(wb, data_sheet_id, current_page, mode=mode)
                    current_row = start_row

                item = final_list[data_ptr]
                item_output_row = current_row

                # 각 을지 페이지의 SEC. NO./LINE NO. 값 영역(H5:L6)에는
                # 갑지의 동일 항목 값(I5)을 그대로 복사한다.
                if mode == "RT" and current_row == start_row:
                    self.safe_set_value(ws, "H5", ws0["I5"].value)

                # [DISABLED] Do NOT write headers automatically to preserve template's own headers
                # if current_row == start_row:
                #     try:
                #         h_row = start_row - 1
                #         if h_row >= 1:
                #             h_map = {
                #                 'RT_COL_NO': 'RT_NAME_NO', 'RT_COL_DATE': 'RT_NAME_DATE',
                #                 'RT_COL_DWG': 'RT_NAME_DWG', 'RT_COL_JOINT': 'RT_NAME_JOINT',
                #                 'RT_COL_LOC': 'RT_NAME_LOC', 'RT_COL_THK': 'RT_NAME_THK',
                #                 'RT_COL_MAT': 'RT_NAME_MAT', 'RT_COL_ACC': 'RT_NAME_ACC',
                #                 'RT_COL_REJ': 'RT_NAME_REJ', 'RT_COL_DEG': 'RT_NAME_DEG',
                #                 'RT_COL_RES': 'RT_NAME_RES', 'RT_COL_WELDER': 'RT_NAME_WELDER',
                #                 'RT_COL_REM': 'RT_NAME_REM'
                #             }
                #             for c_key, n_key in h_map.items():
                #                 c_idx = self.col_to_num(self.config.get(c_key, '0'))
                #                 if c_idx >= 1:
                #                     h_default = c_key.replace('RT_COL_', '').capitalize()
                #                     h_text = self.config.get(n_key, h_default)
                #                     self.safe_set_value(ws, ws.cell(row=h_row, column=c_idx).coordinate, h_text)
                #             for d_i in range(1, 16):
                #                 d_c_idx = self.col_to_num(self.config.get(f'RT_COL_D{d_i}', '0'))
                #                 if d_c_idx >= 1:
                #                     d_text = self.config.get(f'RT_NAME_D{d_i}', f'D{d_i}')
                #                     self.safe_set_value(ws, ws.cell(row=h_row, column=d_c_idx).coordinate, d_text)
                #     except: pass

                if is_kogas:
                    # ===== 가스공사 전용: 1행 1세트 블록 주입 (KOGAS_W_COL_... 적용) =====
                    row_top = current_row      # 단일 행

                    no_col = self.col_to_num(self.config.get('KOGAS_W_COL_NO', '1'))
                    if no_col >= 1:
                        self.safe_set_value(ws, ws.cell(row=row_top, column=no_col).coordinate, item.get('No', ''))

                    joint_col = self.col_to_num(self.config.get('KOGAS_W_COL_JOINT', '4'))
                    if joint_col >= 1:
                        self.safe_set_value(ws, ws.cell(row=row_top, column=joint_col).coordinate, item.get('Joint', ''))

                    date_col = self.col_to_num(self.config.get('KOGAS_W_COL_DATE', '2'))
                    if date_col >= 1:
                        self.safe_set_value(ws, ws.cell(row=row_top, column=date_col).coordinate, item.get('Date', ''))

                    dwg_col = self.col_to_num(self.config.get('KOGAS_W_COL_DWG', '3'))
                    if dwg_col >= 1:
                        self.safe_set_value(ws, ws.cell(row=row_top, column=dwg_col).coordinate, item.get('Dwg', ''))

                    mat_col = self.col_to_num(self.config.get('KOGAS_W_COL_MAT', '7'))
                    if mat_col >= 1:
                        self.safe_set_value(ws, ws.cell(row=row_top, column=mat_col).coordinate, item.get('Mat', ''))

                    wld_col = self.col_to_num(self.config.get('KOGAS_W_COL_WELDER', '29'))
                    if wld_col >= 1:
                        self.safe_set_value(ws, ws.cell(row=row_top, column=wld_col).coordinate, item.get('Welder', ''))

                    thk_col = self.col_to_num(self.config.get('KOGAS_W_COL_THK', '6'))
                    if thk_col >= 1:
                        self.safe_set_value(ws, ws.cell(row=row_top, column=thk_col).coordinate, item.get('T', ''))

                    loc_col = self.col_to_num(self.config.get('KOGAS_W_COL_LOC', '0'))
                    if loc_col >= 1:
                        self.safe_set_value(ws, ws.cell(row=row_top, column=loc_col).coordinate, item.get('Loc', ''))

                    size_col = self.col_to_num(self.config.get('KOGAS_W_COL_SIZE', '0'))
                    if size_col >= 1:
                        self.safe_set_value(ws, ws.cell(row=row_top, column=size_col).coordinate, item.get('Size', ''))

                    # D1~D15 결과 - 상단 행에 개별 지정 열 또는 기본 순차 열 기입
                    for d_i in range(1, 16):
                        d_col_idx = self.col_to_num(self.config.get(f'KOGAS_W_COL_D{d_i}', '0'))
                        if d_col_idx >= 1:
                            d_val = item.get(f'D{d_i}', '')
                            if d_val:
                                self.safe_set_value(ws, ws.cell(row=row_top, column=d_col_idx).coordinate, d_val)

                    acc_col = self.col_to_num(self.config.get('KOGAS_W_COL_ACC', '0'))
                    if acc_col >= 1:
                        res_str = str(item.get('Result', '')).upper()
                        acc_val = item.get('Acc', '')
                        if not acc_val:
                            acc_val = "O" if ("ACC" in res_str or "OK" in res_str) else ""
                        self.safe_set_value(ws, ws.cell(row=row_top, column=acc_col).coordinate, acc_val)

                    rej_col = self.col_to_num(self.config.get('KOGAS_W_COL_REJ', '0'))
                    if rej_col >= 1:
                        res_str = str(item.get('Result', '')).upper()
                        rej_val = item.get('Rej', '')
                        if not rej_val:
                            rej_val = "O" if any(x in res_str for x in ["REJ", "NG", "불합격"]) else ""
                        self.safe_set_value(ws, ws.cell(row=row_top, column=rej_col).coordinate, rej_val)

                    deg_col = self.col_to_num(self.config.get('KOGAS_W_COL_DEG', '0'))
                    if deg_col >= 1:
                        deg_val = str(item.get('Deg', '')).strip()
                        if deg_val.lower() == 'nan': deg_val = ''
                        self.safe_set_value(ws, ws.cell(row=row_top, column=deg_col).coordinate, deg_val)

                    current_row += block_size  # 블록 사이즈만큼 전진 (1행)

                else:
                    # ===== 일반 모드: 1행 1데이터 =====
                    no_col = self.col_to_num(self.config.get('RT_COL_NO', '1'))
                    if no_col >= 1: self.safe_set_value(ws, ws.cell(row=current_row, column=no_col).coordinate, item.get('No', ''))

                    date_col = self.col_to_num(self.config.get('RT_COL_DATE', '2'))
                    if date_col >= 1: self.safe_set_value(ws, ws.cell(row=current_row, column=date_col).coordinate, item.get('Date', ''))

                    dwg_col = self.col_to_num(self.config.get('RT_COL_DWG', '3'))
                    if dwg_col >= 1: self.safe_set_value(ws, ws.cell(row=current_row, column=dwg_col).coordinate, item.get('Dwg', ''))

                    joint_col = self.col_to_num(self.config.get('RT_COL_JOINT', '4'))
                    if joint_col >= 1: self.safe_set_value(ws, ws.cell(row=current_row, column=joint_col).coordinate, item.get('Joint', ''))

                    loc_col = self.col_to_num(self.config.get('RT_COL_LOC', '5'))
                    if loc_col >= 1: self.safe_set_value(ws, ws.cell(row=current_row, column=loc_col).coordinate, item.get('Loc', ''))

                    thk_col = self.col_to_num(self.config.get('RT_COL_THK', '6'))
                    if thk_col >= 1: self.safe_set_value(ws, ws.cell(row=current_row, column=thk_col).coordinate, item.get('T', ''))

                    mat_col = self.col_to_num(self.config.get('RT_COL_MAT', '7'))
                    if mat_col >= 1: self.safe_set_value(ws, ws.cell(row=current_row, column=mat_col).coordinate, item.get('Mat', ''))

                    acc_col = self.col_to_num(self.config.get('RT_COL_ACC', '7'))
                    if acc_col >= 1: self.safe_set_value(ws, ws.cell(row=current_row, column=acc_col).coordinate, item.get('Acc', ''))

                    rej_col = self.col_to_num(self.config.get('RT_COL_REJ', '8'))
                    if rej_col >= 1: self.safe_set_value(ws, ws.cell(row=current_row, column=rej_col).coordinate, item.get('Rej', ''))

                    deg_col = self.col_to_num(self.config.get('RT_COL_DEG', '9'))
                    if deg_col >= 1: self.safe_set_value(ws, ws.cell(row=current_row, column=deg_col).coordinate, item.get('Deg', ''))

                    iqi_col = self.col_to_num(self.config.get('RT_COL_IQI', '10'))
                    if iqi_col >= 1: self.safe_set_value(ws, ws.cell(row=current_row, column=iqi_col).coordinate, item.get('IQI', ''))

                    sens_col = self.col_to_num(self.config.get('RT_COL_SENS', '11'))
                    if sens_col >= 1: self.safe_set_value(ws, ws.cell(row=current_row, column=sens_col).coordinate, item.get('Sens', ''))

                    den_col = self.col_to_num(self.config.get('RT_COL_DEN', '12'))
                    if den_col >= 1: self.safe_set_value(ws, ws.cell(row=current_row, column=den_col).coordinate, item.get('Den', ''))

                    for d_i in range(1, 16):
                        d_col_key = f'RT_COL_D{d_i}'
                        d_col_idx = self.col_to_num(self.config.get(d_col_key, str(12 + d_i)))
                        if d_col_idx >= 1:
                            self.safe_set_value(ws, ws.cell(row=current_row, column=d_col_idx).coordinate, item.get(f'D{d_i}', ''))

                    res_col = self.col_to_num(self.config.get('RT_COL_RES', '28'))
                    if res_col >= 1: self.safe_set_value(ws, ws.cell(row=current_row, column=res_col).coordinate, item.get('Result', 'ACC'))

                    wld_col = self.col_to_num(self.config.get('RT_COL_WELDER', '29'))
                    if wld_col >= 1: self.safe_set_value(ws, ws.cell(row=current_row, column=wld_col).coordinate, item.get('Welder', ''))

                    rem_col = self.col_to_num(self.config.get('RT_COL_REM', '30'))
                    if rem_col >= 1: self.safe_set_value(ws, ws.cell(row=current_row, column=rem_col).coordinate, item.get('Remarks', ''))

                    # Empty positions are meaningful: ,N/1 is the second film
                    # and ,,N/1 is the third, directly under the rejected interval.
                    film_locations = [
                        value.strip()
                        for value in re.split(r'[,;\n]', str(item.get('Loc', '')))
                    ]
                    rt_film_location_counts[ws.title] = max(
                        rt_film_location_counts.get(ws.title, 0),
                        min(len(film_locations), 8)
                    )
                    for offset, col_idx in enumerate(range(4, 12)):
                        ws.cell(row=current_row, column=col_idx).value = (
                            film_locations[offset] if offset < len(film_locations) else None
                        )

                    # 갑지 D32 값은 을지 M열의 실제 데이터 행에 반복 입력한다.
                    ws[f"M{current_row}"] = cover_d32_value

                    # 을지 원본의 검사일자 열은 S열이다.
                    ws[f"S{current_row}"] = item.get('Date', '')

                    # 을지 원본의 관경(DIA.) 열은 T열이다.
                    ws[f"T{current_row}"] = item.get('Size', '')

                    # 기본 검사조건은 실제 RT 데이터가 있는 행에만 표시한다.
                    fixed_condition_values = {
                        "N": 0.20,   # SENSITIVITY
                        "O": 1,      # DENSITY
                        "P": "-",
                        "Q": 4.0,
                        "R": "DWSI" # TECHNIQUE
                    }
                    for col_letter, fixed_value in fixed_condition_values.items():
                        ws[f"{col_letter}{current_row}"] = fixed_value
                    rt_condition_data_rows.add((ws.title, current_row))

                    # 을지의 실제 데이터 행은 모든 값을 상하좌우 중앙 정렬한다.
                    for col_idx in range(1, 23):  # A:V
                        ws.cell(row=current_row, column=col_idx).alignment = Alignment(
                            horizontal='center', vertical='center',
                            wrap_text=True, shrink_to_fit=True
                        )

                    current_row += 1

                    # Keep the exact item written to each RT row for the footer summary.
                    rt_summary_items[(ws.title, item_output_row)] = item

                data_ptr += 1
                self.progress['value'] = (data_ptr / len(final_list)) * 95

            # [NEW] 서식 정리 [DISABLED] - 템플릿 선(Border) 보존을 위해 주석 처리
            # rt_data_end_row = int(float(self.config.get('RT_DATA_END_ROW', 31)))
            total_rt_pages = len(wb.worksheets)
            for p_idx, s in enumerate(wb.worksheets):
                # if p_idx > 0: # 을지
                #     for r_idx in range(rt_data_end_row + 1, rt_data_end_row + 15):
                #         for c_idx in range(1, 31):
                #             try: s.cell(row=r_idx, column=c_idx).border = Border()
                #             except: pass

                self.apply_custom_dimensions(s, "DATA" if p_idx > 0 else "COVER")
                # 각 워크시트가 한 페이지이므로 전체 보고서 기준 번호를 고정한다.
                # 셀을 사용하지 않고 우측 인쇄 머리글에 표시해 기존 양식을 보존한다.
                s.oddHeader.right.text = (
                    f"Page    {p_idx + 1}    of    {total_rt_pages}"
                )
                s.oddHeader.right.font = "Arial"
                s.oddHeader.right.size = 9
                s.page_margins.header = 0.55
                if p_idx > 0 and mode == "RT":
                    # A copied/template page may contain old rows beyond the
                    # current page's last item. Clear all unused data cells.
                    for row_idx in range(start_row, end_row + 1):
                        if (s.title, row_idx) not in rt_summary_items:
                            for cell in s.iter_rows(min_row=row_idx, max_row=row_idx,
                                                    min_col=1, max_col=22):
                                for value_cell in cell:
                                    if not isinstance(value_cell, openpyxl.cell.cell.MergedCell):
                                        value_cell.value = None
                    # Repeat the cover's report number on every continuation page.
                    self.safe_set_value(s, "B5", ws0["C5"].value)
                    s["B5"].alignment = Alignment(
                        horizontal='center', vertical='center',
                        wrap_text=True, shrink_to_fit=True
                    )
                    # 갑지의 고정 발주처를 모든 을지 CUSTOMER 영역에 동일 적용한다.
                    self.safe_set_value(s, "O1", ws0["Q1"].value)
                    s["O1"].alignment = Alignment(
                        horizontal='center', vertical='center',
                        wrap_text=True, shrink_to_fit=True
                    )

                    # 갑지의 고정 공사명을 모든 을지 PROJECT 영역에 동일 적용한다.
                    self.safe_set_value(s, "O2", ws0["Q2"].value)
                    s["O2"].alignment = Alignment(
                        horizontal='center', vertical='center',
                        wrap_text=True, shrink_to_fit=True
                    )

                    # 을지 DRAWING NO./도면번호 값은 N/A로 고정한다.
                    self.safe_set_value(s, "O5", "N/A")
                    # Slightly reduce the dense D7 header while preserving all other font styling.
                    d7_cell = s["D7"]
                    d7_font_size = d7_cell.font.sz or 8
                    d7_cell.font = d7_cell.font.copy(size=max(1, d7_font_size - 0.5))
                    s["S8"] = "검사일자"
                    s["S8"].alignment = Alignment(
                        horizontal='center', vertical='center',
                        wrap_text=False, shrink_to_fit=True
                    )
                    s["O5"].alignment = Alignment(
                        horizontal='center', vertical='center',
                        wrap_text=True, shrink_to_fit=True
                    )

                    # Film Location 개수에 따라 D8:K8 순환 구간 헤더를 생성한다.
                    location_count = rt_film_location_counts.get(s.title, 0)
                    for offset, col_idx in enumerate(range(4, 12)):
                        if offset < location_count:
                            next_no = offset + 2 if offset + 1 < location_count else 1
                            s.cell(row=8, column=col_idx).value = f"{offset + 1}-{next_no}"
                        else:
                            s.cell(row=8, column=col_idx).value = None

                    # 템플릿의 미사용 기본값은 제거하고 실제 데이터 행만 남긴다.
                    for row_idx in range(9, 38):
                        if (s.title, row_idx) not in rt_condition_data_rows:
                            for col_letter in "NOPQR":
                                s[f"{col_letter}{row_idx}"] = None

                    # 마지막 데이터 다음 첫 행에는 검사 요약, 두 번째 행에는 이하 여백을 표시한다.
                    sheet_data_rows = sorted(
                        row_idx for sheet_title, row_idx in rt_condition_data_rows
                        if sheet_title == s.title
                    )
                    if sheet_data_rows:
                        # 데이터 양에 맞춰 마지막 데이터 바로 다음 행에 요약을 배치한다.
                        summary_row = sheet_data_rows[-1] + 1
                        blank_notice_row = summary_row + 1
                        print_end_row = int(self.config.get('RT_PRINT_END_ROW', end_row + 2))
                        if blank_notice_row <= print_end_row:
                            sheet_items = [
                                rt_summary_items[(s.title, row_idx)]
                                for row_idx in sheet_data_rows
                                if (s.title, row_idx) in rt_summary_items
                            ]

                            def has_summary_value(value):
                                text_value = str(value).strip()
                                return bool(text_value) and text_value.lower() not in ("nan", "none", "null", "-")

                            # ORIGINAL: number of populated inspection-position cells in D:K.
                            original_count = report_original_count

                            # REPAIR: only values 3 or 4 in D:K on rows checked as rejected in C.
                            repair_count = report_repair_count

                            total_count = original_count + repair_count
                            inspection_count = sum(
                                1 for item in final_list
                                if has_summary_value(item.get("Joint", ""))
                            )
                            sizes = []
                            for item in final_list:
                                size = str(item.get("Size", "")).strip()
                                if size and size.lower() != "nan" and size not in sizes:
                                    sizes.append(size)
                            size_text = "/".join(
                                size if size.upper().endswith("A") else f"{size}A"
                                for size in sizes
                            )

                            # Rebuild the row so it follows the requested nine-part summary layout.
                            for merged_range in list(s.merged_cells.ranges):
                                if (
                                    merged_range.min_row <= summary_row <= merged_range.max_row
                                    and merged_range.max_col >= 1
                                    and merged_range.min_col <= 18  # Only merges intersecting A:R
                                ):
                                    s.unmerge_cells(str(merged_range))
                            summary_parts = [
                                (1, 2, "검 사 개 소"),       # A:B
                                (3, 4, size_text),            # C:D
                                (5, 6, f"{inspection_count} 개소"),  # E:F
                                (7, 8, "ORIGINAL"),          # G:H
                                (9, 10, f"{original_count} 매"),     # I:J
                                (11, 12, "REPAIR"),          # K:L
                                (13, 14, f"{repair_count} 매"),      # M:N
                                (15, 18, f"TOTAL  {total_count} 매"), # O:R
                            ]
                            summary_border_side = Side(style='thin', color='000000')
                            for start_col, end_col, value in summary_parts:
                                if start_col != end_col:
                                    s.merge_cells(
                                        start_row=summary_row, start_column=start_col,
                                        end_row=summary_row, end_column=end_col
                                    )
                                cell = s.cell(row=summary_row, column=start_col)
                                cell.value = value
                                cell.font = Font(name='맑은 고딕', size=9)
                                cell.alignment = Alignment(
                                    horizontal='center', vertical='center', shrink_to_fit=True
                                )
                                # Thin outline for every merged summary block. Together these
                                # form a continuous A:R border with thin vertical separators.
                                for col_idx in range(start_col, end_col + 1):
                                    border_cell = s.cell(row=summary_row, column=col_idx)
                                    old_border = border_cell.border
                                    border_cell.border = Border(
                                        left=(summary_border_side if col_idx == start_col else old_border.left),
                                        right=(summary_border_side if col_idx == end_col else old_border.right),
                                        top=summary_border_side,
                                        bottom=summary_border_side
                                    )

                            # Extend the summary-row outer outline through column V while
                            # preserving the existing S:V merges and their internal layout.
                            for col_idx in range(1, 23):  # A:V
                                outline_cell = s.cell(row=summary_row, column=col_idx)
                                old_border = outline_cell.border
                                outline_cell.border = Border(
                                    left=(summary_border_side if col_idx == 1 else old_border.left),
                                    right=(summary_border_side if col_idx == 22 else old_border.right),
                                    top=summary_border_side,
                                    bottom=summary_border_side
                                )

                            # Merge the entire RT report width for the blank notice.
                            for merged_range in list(s.merged_cells.ranges):
                                if merged_range.min_row <= blank_notice_row <= merged_range.max_row:
                                    s.unmerge_cells(str(merged_range))
                            for col_idx in range(1, 23):  # A:V
                                notice_cell = s.cell(row=blank_notice_row, column=col_idx)
                                notice_cell.value = None
                            s.merge_cells(
                                start_row=blank_notice_row, start_column=1,
                                end_row=blank_notice_row, end_column=22
                            )
                            notice_cell = s.cell(row=blank_notice_row, column=1)
                            notice_cell.alignment = Alignment(
                                horizontal='center', vertical='center'
                            )
                            notice_cell.value = "~이  하   여  백~"
                            notice_cell.font = Font(name='돋움체', size=9)

                    # Standardize the RT data-sheet perimeter without changing inner borders.
                    outer_side = Side(style='thin', color='000000')
                    outer_last_row = 37
                    for col_idx in range(1, 23):  # A:V
                        top_cell = s.cell(row=1, column=col_idx)
                        top_border = top_cell.border
                        top_cell.border = Border(
                            left=top_border.left, right=top_border.right,
                            top=outer_side, bottom=top_border.bottom
                        )

                        bottom_cell = s.cell(row=outer_last_row, column=col_idx)
                        bottom_border = bottom_cell.border
                        bottom_cell.border = Border(
                            left=bottom_border.left, right=bottom_border.right,
                            top=bottom_border.top, bottom=outer_side
                        )

                    for row_idx in range(1, outer_last_row + 1):
                        left_cell = s.cell(row=row_idx, column=1)
                        left_border = left_cell.border
                        left_cell.border = Border(
                            left=outer_side, right=left_border.right,
                            top=left_border.top, bottom=left_border.bottom
                        )

                        right_cell = s.cell(row=row_idx, column=22)
                        right_border = right_cell.border
                        right_cell.border = Border(
                            left=right_border.left, right=outer_side,
                            top=right_border.top, bottom=right_border.bottom
                        )

            # [FINAL CLEANUP] Standardized cleanup for RT
            try:
                ws0 = wb.worksheets[0]
                sig_end = 48
                # [SAFE] Only clear far below the Shooting Sketch (Rows 65+)
                for r in range(65, 80):
                    # Do NOT unmerge on Gapji anymore to be safe
                    # merged_ranges = [str(rng) for rng in ws0.merged_cells.ranges if rng.min_row >= r]
                    for c in range(1, 21):
                        try: ws0.cell(row=r, column=c).border = Border()
                        except: pass
                self.log("   ✅ RT 갑지 하단 잔여 서식 정리 완료 (65-80행)")
            except: pass

            now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            output_ext = '.xlsm' if keep_vba else '.xlsx'
            save_path = self._report_save_path(os.path.dirname(template_path), output_ext, mode=mode)
            if not save_path:
                return
            output_name = os.path.basename(save_path)
            wb.save(save_path)

            # [SURGERY] openpyxl destroys complex drawings (like grouped shapes). 
            # We surgically inject the template's drawing layer back into the generated file.
            self.inject_drawing_layer(template_path, save_path)
            if mode == 'RT':
                from site_apps.central.src.reports.report_signatures import append_rt_signatures
                signatures = []
                stamp_data = self._get_photo_log_stamp()
                if stamp_data:
                    from site_apps.central.src.reports.rt_report_layout import header_stamp_anchor
                    stamp_anchor = header_stamp_anchor(ws0)
                    signatures.append((io.BytesIO(stamp_data), stamp_anchor, 72, 72))
                else:
                    self.log('⚠️ RT 직인 삽입 실패: PAUT 양식의 직인을 확인해 주세요.')
                for cell, anchor in (('D32', 'E32'), ('D35', 'E35')):
                    name = re.sub(r'\s+', '', str(ws0[cell].value or ''))
                    path = os.path.join(SRC_ROOT, 'signs', name + '.png')
                    if os.path.isfile(path):
                        signatures.append((self._white_background_to_transparent_png(path), anchor, 65, 23))
                    else:
                        self.log(f'서명 파일 없음: {name}')
                if signatures:
                    append_rt_signatures(save_path, signatures)
                if stamp_data:
                    for sheet_index in range(1, len(wb.worksheets)):
                        append_rt_signatures(
                            save_path,
                            [(io.BytesIO(stamp_data), header_stamp_anchor(wb.worksheets[sheet_index]), 72, 72)],
                            sheet_index=sheet_index,
                        )

            self._record_report_issue(mode, wb, save_path, final_list)

            self.progress['value'] = 100
            self.log(f"✨ RT 완료! 저장됨: {output_name}")
            messagebox.showinfo("성공", f"RT 성적서 생성이 완료되었습니다.\n경로: {os.path.dirname(save_path)}")
        except Exception as e:
            self.log(f"❌ RT 생성 오류: {e}")
            traceback.print_exc()
            messagebox.showerror("생성 실패", f"RT 성적서 생성 중 오류가 발생했습니다: {e}")
        finally:
            if 'wb' in locals() and wb:
                try:
                    if hasattr(wb, 'vba_archive') and wb.vba_archive:
                        wb.vba_archive.close()
                except: pass
                try: wb.close()
                except: pass
            for f in glob.glob(os.path.join(tempfile.gettempdir(), "temp_*.png")):
                try: os.remove(f)
                except: pass

