# -*- coding: utf-8 -*-
"""보고서 엑셀 시트 서식/로고/인쇄설정/도면레이어 믹스인 (Report Sheet Styler Mixin)"""
import os
import io
import re
import copy
import traceback
from pathlib import Path
import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.utils.cell import coordinate_from_string, coordinate_to_tuple, column_index_from_string

thin_side = Side(style='thin')
medium_side = Side(style='medium')


class ReportSheetStylerMixin:
    def add_logos_to_sheet(self, ws, is_cover=False, clear_existing=True, mode=None):
        # [SUPER SAFE] Identify if this is the first sheet (Cover/Gapji) of the workbook
        try:
            is_first_sheet = (ws == ws.parent.worksheets[0])
        except:
            is_first_sheet = is_cover

        # [FIX] Never clear existing images on the Cover sheet (Gapji) or in RT/PMI mode to preserve template diagrams and logos
        if is_first_sheet or is_cover or mode in ("RT", "PMI"):
            clear_existing = False

        # [NEW] If RT mode and images already exist, skip adding logos to preserve the template perfectly
        if mode == "RT" and getattr(ws, '_images', None):
            self.log(f"   ℹ️ [RT] 시트에 이미 그림(도면)이 있어 추가 로고 삽입을 건너뜁니다.")
            return

        if clear_existing:
            try: ws._images = [] 
            except: pass

        # [ENHANCED] Use provided mode or tracked current_mode
        if mode is None:
            mode = getattr(self, 'current_mode', "PMI")
            try:
                tab_text = self.mode_notebook.tab(self.mode_notebook.select(), "text")
                if "PMI" in tab_text: mode = "PMI"
                elif "RT" in tab_text: mode = "RT"
                elif "PT" in tab_text: mode = "PT"
                elif "MT" in tab_text: mode = "MT"
                elif "PAUT" in tab_text: mode = "PAUT"
            except: pass

        ctx = "COVER" if is_cover else "DATA"
        context = f"{mode}_{ctx}"

        debug_msg = f"?? 로고 엔진 가동: {context} (Mode: {mode}, Sheet: {ws.title})"
        self.log(debug_msg)

        def _get_val(prefix, suffix, default):
            # [FIX] Match ACTUAL UI key format: {prefix}_{mode}_{ctx}{suffix}
            # e.g., SITCO_RT_COVER_X
            full_key = f"{prefix}_{context}{suffix}"
            v = self.config.get(full_key)

            # If value is missing or effectively zero for coordinates/size
            is_coord = suffix in ["_X", "_Y", "_W", "_H"]
            if v is not None and v != "":
                try:
                    if is_coord and float(v) <= 0: pass # Fallback if zero
                    else: return v
                except: pass

            # [FIX] Enhanced Fallback: If DATA is missing, try COVER of the same mode first
            if "_DATA" in context:
                cover_key = f"{prefix}_" + context.replace("_DATA", "_COVER") + suffix
                cv = self.config.get(cover_key)
                if cv is not None and cv != "":
                    try:
                        if is_coord and float(cv) <= 0: pass
                        else: return cv
                    except: pass

            # [FIX] Enhanced Fallback: If current mode settings are missing, try PMI_COVER
            if mode != "PMI":
                pmi_fallback_key = f"{prefix}_PMI_COVER{suffix}"
                pv = self.config.get(pmi_fallback_key)
                if pv is not None and pv != "":
                    try:
                        if is_coord and float(pv) <= 0: pass
                        else: return pv
                    except: pass

            # [FIX] Final Global Fallback (e.g., SITCO_COVER_X)
            global_ctx = "COVER" if "COVER" in context else "DATA"
            global_key = f"{prefix}_{global_ctx}{suffix}"
            gv = self.config.get(global_key, default)
            if is_coord:
                try:
                    if gv is None or gv == "" or float(gv) <= 0: return default
                except: return default
            return gv

        def _get_effective_path(prefix, keywords):
            # 1. [FIX] Match ACTUAL UI key format: {prefix}_{mode}_{ctx}_PATH
            # e.g., SITCO_RT_COVER_PATH
            k1 = f"{prefix}_{context}_PATH"
            path = self.config.get(k1, "")

            # [FIX] Enhanced Fallback: If DATA path is empty, try COVER path of same mode
            if (not path or not str(path).strip()) and "_DATA" in context:
                k2 = f"{prefix}_" + context.replace("_DATA", "_COVER") + "_PATH"
                path = self.config.get(k2, "")

            # [FIX] Enhanced Fallback: If current mode path is empty, try PMI_COVER path
            if (not path or not str(path).strip()) and mode != "PMI":
                k3 = f"{prefix}_PMI_COVER_PATH"
                path = self.config.get(k3, "")

            # [FIX] Final Global Fallback (e.g., SITCO_COVER_PATH)
            if (not path or not str(path).strip()):
                global_ctx = "COVER" if "COVER" in context else "DATA"
                k4 = f"{prefix}_{global_ctx}_PATH"
                path = self.config.get(k4, "")

            # self.log(f"   > '{prefix}' 검색 시작 (Target Context: {context})")

            # 2. If path is valid and exists, use it
            if path and str(path).strip() and os.path.exists(str(path).strip()):
                self.log(f"   [OK] '{prefix}' 수동 지정 경로 발견: {os.path.basename(str(path))}")
                return str(path).strip()

            # 3. Fallback to global path if mode-specific is empty
            global_ctx = "COVER" if "COVER" in context else "DATA"
            # [FIX] Match UI key format for global fallback path
            g_path = self.config.get(f"{global_ctx}_{prefix}_PATH", "")
            if g_path and str(g_path).strip() and os.path.exists(str(g_path).strip()):
                self.log(f"   [OK] '{prefix}' 공용 설정 경로 사용: {os.path.basename(str(g_path))}")
                return str(g_path).strip()

            # 4. If manual path is empty or invalid, try Smart Search in Logo Folder
            # self.log(f"   ? '{prefix}' 설정 없음 -> 기본 폴더 자동 검색 시작...")
            for kw in keywords:
                found = self.find_image_smart(kw)
                if found:
                    self.log(f"   [OK] '{prefix}' 자동 검색 성공: {os.path.basename(found)}")
                    return found

            # self.log(f"   [FAIL] '{prefix}' 로고를 어디에서도 찾을 수 없습니다.")
            return None

        def _has_image_near(cell_ref, row_tol=2, col_tol=3):
            """Return True when the template already has an image at or near cell_ref."""
            try:
                target_r, target_c = coordinate_to_tuple(str(cell_ref))
                for existing_image in getattr(ws, '_images', []):
                    anchor = getattr(existing_image, 'anchor', None)
                    if isinstance(anchor, str):
                        r, c = coordinate_to_tuple(anchor)
                        if abs(r - target_r) <= row_tol and abs(c - target_c) <= col_tol:
                            return True
                        continue
                    marker = getattr(anchor, '_from', None)
                    if marker:
                        r, c = marker.row + 1, marker.col + 1
                        if abs(r - target_r) <= row_tol and abs(c - target_c) <= col_tol:
                            return True
            except Exception:
                pass
            return False

        _has_image_at = _has_image_near

        # 1. SITCO - preserve a logo already embedded in the template.
        sitco_anchor = _get_val("SITCO", "_ANCHOR", "A1")
        p = _get_effective_path("SITCO", ["SITCO"])
        if p:
            if _has_image_near(sitco_anchor, row_tol=2, col_tol=2):
                self.log(f"   [INFO] 'SITCO' 템플릿에 기존 로고가 있어 중복 배치를 건너뜁니다.")
            else:
                self.place_image_freely(ws, p, sitco_anchor, float(_get_val("SITCO", "_W", 100)), float(_get_val("SITCO", "_H", 50)), float(_get_val("SITCO", "_X", 0)), float(_get_val("SITCO", "_Y", 0)))

        # 2. SEOUL - preserve a header/logo already embedded in the template.
        # [USER REQUEST] 갑지와 을지 서울검사 로고 없어도 됨 (PMI 모드에서는 삽입 생략)
        seoul_anchor = _get_val("SEOUL", "_ANCHOR", "E5" if mode == "PMI" else "A1")
        p = None if mode == "PMI" else _get_effective_path("SEOUL", ["서울검사"])
        if p:
            if _has_image_near(seoul_anchor, row_tol=2, col_tol=3):
                self.log(f"   [INFO] 'SEOUL' 템플릿에 기존 로고가 있어 중복 배치를 건너뜁니다.")
            else:
                seoul_w = float(_get_val("SEOUL", "_W", 200 if mode == "PMI" else 100))
                seoul_h = float(_get_val("SEOUL", "_H", 35 if mode == "PMI" else 50))
                seoul_x = float(_get_val("SEOUL", "_X", 0))
                seoul_y = float(_get_val("SEOUL", "_Y", 0))
                # [PMI RULE] 갑지는 K열, 을지는 J열을 넘어가면 안 됨
                if mode == "PMI":
                    seoul_anchor = "E5"
                    max_w = 280.0 if is_cover else 230.0
                    max_h = 42.0 if is_cover else 40.0
                    seoul_w = min(seoul_w, max_w)
                    seoul_h = min(seoul_h, max_h)
                self.place_image_freely(ws, p, seoul_anchor, seoul_w, seoul_h, seoul_x, seoul_y)

        # 3. FOOTER (우측 바닥글: SITCO 서울검사 주식회사 로고)
        if mode == "PMI":
            # [NOTICE] PMI 모드는 템플릿의 [페이지 설정] -> [머리글/바닥글] VML Drawing(&G)을 통해
            # 인쇄 시 용지 최하단 여백 바닥글로 온전하게 복원(restore_template_footer_pictures)되므로,
            # 셀(51행/46행) 위에 임의로 복사 배치하여 인쇄 영역 밖으로 밀려나거나 중복되는 것을 방지합니다.
            self.log("   [INFO] PMI 모드: 템플릿 머리글/바닥글(VML &G) 규격을 사용하므로 셀 내 중복 삽입을 건너뜁니다.")
        else:
            footer_keywords = ["바닥글", "PMI"]
            p = _get_effective_path("FOOTER", footer_keywords)
            if p:
                f_anchor = _get_val("FOOTER", "_ANCHOR", "A1")
                f_w = float(_get_val("FOOTER", "_W", 100))
                f_h = float(_get_val("FOOTER", "_H", 50))
                f_x = float(_get_val("FOOTER", "_X", 0))
                f_y = float(_get_val("FOOTER", "_Y", 0))
                has_existing_footer_img = _has_image_near(f_anchor, row_tol=2, col_tol=4)
                if not has_existing_footer_img:
                    self.place_image_freely(ws, p, f_anchor, f_w, f_h, f_x, f_y)

        # 4. FOOTER_PT (Left: 좌측 바닥글 양식 번호)
        if mode == "PMI":
            f_pt_anchor = "A51" if is_cover else "A46"
            # [PRESERVE TEMPLATE & PROCEDURE FOOTER]
            # 1) 셀에 이미 텍스트(예: '양식 E 501-1', '양식 H 801-1' 등)가 존재하는지 확인
            cell_has_text = False
            for chk_r in ([50, 51, 52] if is_cover else [45, 46, 47]):
                v = ws.cell(row=chk_r, column=1).value
                if v and ('양식' in str(v) or '양 식' in str(v) or 'FORM' in str(v).upper()):
                    cell_has_text = True
                    break
            # 2) oddFooter.left.text에 절차서 번호가 이미 있는지 확인
            has_odd_footer = hasattr(ws, 'oddFooter') and ws.oddFooter and ws.oddFooter.left.text and ('양' in str(ws.oddFooter.left.text) or 'FORM' in str(ws.oddFooter.left.text).upper() or 'H ' in str(ws.oddFooter.left.text) or 'E ' in str(ws.oddFooter.left.text))
            # 3) 기존 이미지가 이미 배치되어 있는지 확인
            has_existing_pt_img = _has_image_near(f_pt_anchor, row_tol=2, col_tol=3)

            is_triple_mode = hasattr(self, 'pmi_layout_mode') and self.pmi_layout_mode.get() == "PLANT_TRIPLE"

            if cell_has_text or has_odd_footer or has_existing_pt_img:
                self.log(f"   [INFO] 'FOOTER_PT' 템플릿 절차서 원본 바닥글이 감지되어 기존 서식을 보존합니다.")
            elif not is_triple_mode and not is_cover:
                # [STANDARD MODE EULJI] 단일행 일반 모드 을지에는 플랜트용 '양식 H 802-2'('바닥글 PMI-1') 이미지를 강제 삽입하지 않음
                self.log(f"   [INFO] 'FOOTER_PT' STANDARD 단일행 을지에는 플랜트 배관용 양식을 주입하지 않습니다.")
            else:
                footer_left_keywords = ["바닥글 PMI갑"] if is_cover else ["바닥글 PMI-1"]
                p = _get_effective_path("FOOTER_PT", footer_left_keywords)
                if p:
                    self.place_image_freely(ws, p, f_pt_anchor, 92.0, 22.0, 0.0, 0.0)
                    r_target = 51 if is_cover else 46
                    if ws.row_dimensions[r_target].height is None or ws.row_dimensions[r_target].height < 24.0:
                        ws.row_dimensions[r_target].height = 24.0
        else:
            footer_left_keywords = ["PMI갑", "PMI-1", "PT"]
            p = _get_effective_path("FOOTER_PT", footer_left_keywords)
            if p: self.place_image_freely(ws, p, _get_val("FOOTER_PT", "_ANCHOR", "A1"), float(_get_val("FOOTER_PT", "_W", 100)), float(_get_val("FOOTER_PT", "_H", 50)), float(_get_val("FOOTER_PT", "_X", 0)), float(_get_val("FOOTER_PT", "_Y", 0)))



    def force_print_settings(self, ws, context="DATA"):
        try:
            # [ENHANCED] Use tracked current_mode
            mode = getattr(self, 'current_mode', "PMI")
            try:
                tab_text = self.mode_notebook.tab(self.mode_notebook.select(), "text")
                if "PMI" in tab_text: mode = "PMI"
                elif "RT" in tab_text: mode = "RT"
                elif "PT" in tab_text: mode = "PT"
                elif "MT" in tab_text: mode = "MT"
                elif "PAUT" in tab_text: mode = "PAUT"
            except: pass

            full_context = f"{mode}_{context}"

            # [NEW] Manual Print Area Override (Priority: Mode-Specific > Global)
            manual_area = self.config.get(f'PRINT_AREA_{full_context}', "").strip()
            if not manual_area:
                manual_area = self.config.get(f'PRINT_AREA_{context}', "").strip()

            if mode == "RT":
                if context == "COVER":
                    ws.print_area = 'A1:V35'
                else:
                    ws.print_area = 'A1:V37'
            elif manual_area:
                ws.print_area = manual_area
            else:
                if context == "COVER":
                    # [NEW] Highly Dynamic print area for Gapji
                    key = f"{mode}_GAPJI_PRINT_END_ROW" if mode != "PMI" else "GAPJI_PRINT_END_ROW"
                    end_r = int(self.config.get(key, 51))

                    if mode == "PAUT" and end_r > 0:
                        ws.print_area = f'A1:AJ{end_r}'
                    elif mode == "PT" and end_r > 0:
                        ws.print_area = f'A1:S{end_r}'
                    elif end_r > 0:
                        ws.print_area = f'A1:T{end_r}'
                else:
                    # [NEW] Highly Dynamic print area for Eulji
                    key = f"{mode}_PRINT_END_ROW" if mode != "PMI" else "PRINT_END_ROW"
                    end_r = int(self.config.get(key, 47))
                    if end_r > 0:
                        if mode == "PAUT":
                            ws.print_area = f'A1:AH{end_r}'
                        elif mode == "PT":
                            ws.print_area = f'A1:S{end_r}'
                        elif mode == "PMI":
                            is_triple = hasattr(self, 'pmi_layout_mode') and self.pmi_layout_mode.get() == "PLANT_TRIPLE"
                            if is_triple:
                                ws.print_area = 'A1:M45'
                            else:
                                if ws.row_dimensions[46].height is None or ws.row_dimensions[46].height < 24.0:
                                    ws.row_dimensions[46].height = 24.0
                                ws.print_area = f'A1:T{end_r}'
                        else:
                            ws.print_area = f'A1:M{end_r}'

            # [FIX] Clear any manual page breaks that might cause mid-page splitting
            try:
                ws.row_breaks = []
                ws.col_breaks = []
            except: pass

            ws.page_setup.paperSize = 9
            # PAUT 갑지와 을지는 모두 A4 세로 방향으로 출력한다.
            if mode in ("PMI", "PAUT", "PT"):
                ws.page_setup.orientation = 'portrait'
            elif context == "DATA" or mode == "RT":
                ws.page_setup.orientation = 'landscape'
            else:
                ws.page_setup.orientation = 'portrait'

            # [FIX] 모드별 설정 우선 읽기 (RT_COVER → COVER 순 폴백)
            default_scale = 95
            if mode == "RT" and context == "COVER": default_scale = 90
            def _cfg_str(key):
                v = self.config.get(key, '')
                return str(v).strip() if v != '' else ''
            scale_val = (_cfg_str(f'PRINT_SCALE_{full_context}') or
                         _cfg_str(f'PRINT_SCALE_{context}') or
                         str(default_scale))
            ws.print_options.horizontalCentered = True
            # Keep cover and data sheets visually aligned within the same A4
            # margins. PMI page sizing is handled by the explicit scale below.
            ws.print_options.verticalCentered = True

            if mode == "RT" and context == "DATA":
                # 을지는 갑지보다 약 4% 크게 출력되므로 고정 87%로 축소한다.
                ws.sheet_properties.pageSetUpPr.fitToPage = False
                ws.page_setup.fitToWidth = None
                ws.page_setup.fitToHeight = None
                ws.page_setup.scale = 87
            elif mode == "PMI" and context == "DATA":
                # [PMI 1페이지 완벽 맞춤] 배관용 및 일반 모드 모두 1시트당 정확히 1페이지로 고정 (4페이지 분할 원천 차단)
                ws.sheet_properties.pageSetUpPr.fitToPage = True
                ws.page_setup.fitToWidth = 1
                ws.page_setup.fitToHeight = 1
                ws.page_setup.scale = None
            elif mode == "PT" and context == "DATA":
                # PT의 S(Scale) 입력값을 실제 Excel 인쇄 배율로 적용한다.
                # fitToPage가 켜져 있으면 scale 값이 무시되므로 함께 해제한다.
                try:
                    pt_scale = max(10, min(400, int(float(scale_val))))
                except (TypeError, ValueError):
                    pt_scale = 95
                ws.sheet_properties.pageSetUpPr.fitToPage = False
                ws.page_setup.fitToWidth = None
                ws.page_setup.fitToHeight = None
                ws.page_setup.scale = pt_scale
            else:
                # 그 외 시트는 기존 한 페이지 맞춤을 유지한다.
                ws.sheet_properties.pageSetUpPr.fitToPage = True
                ws.page_setup.fitToWidth = 1
                ws.page_setup.fitToHeight = (
                    0 if mode == "RT" and context == "COVER" else 1
                )
                if mode in ("PT", "PMI") and context == "COVER":
                    ws.page_setup.scale = None

            def _margin(name, default):
                return float(
                    _cfg_str(f'MARGIN_{full_context}_{name}') or
                    _cfg_str(f'MARGIN_{context}_{name}') or
                    default
                )
            ws.page_margins.top    = _margin('TOP',    0.2)
            ws.page_margins.bottom = _margin('BOTTOM', 0.2)
            if mode == "PMI":
                # [PMI 여백 표준화] 템플릿 원본 규격(0.25인치, 약 6.35mm / 5~7mm)으로 완전 복원하여
                # 4페이지 분할을 원천 차단하고 갑지/을지 1장 맞춤 및 좌우 여백을 완벽 일치시킴
                ws.page_margins.left  = 0.25
                ws.page_margins.right = 0.25
                ws.page_margins.top   = 0.25
                ws.page_margins.bottom = 0.25
                if hasattr(self, 'config'):
                    self.config['MARGIN_PMI_COVER_LEFT'] = 0.25
                    self.config['MARGIN_PMI_COVER_RIGHT'] = 0.25
                    self.config['MARGIN_PMI_DATA_LEFT'] = 0.25
                    self.config['MARGIN_PMI_DATA_RIGHT'] = 0.25
                if hasattr(self, 'setting_vars'):
                    for k in ['MARGIN_PMI_COVER_LEFT', 'MARGIN_PMI_COVER_RIGHT', 'MARGIN_PMI_DATA_LEFT', 'MARGIN_PMI_DATA_RIGHT']:
                        if k in self.setting_vars:
                            try: self.setting_vars[k].set("0.25")
                            except: pass
            else:
                ws.page_margins.left   = _margin('LEFT',   0.5)
                ws.page_margins.right  = _margin('RIGHT',  0.3)
            if mode == "PAUT":
                # Keep the photo log unchanged; both PAUT sheets follow its
                # current margin controls, including restored photo settings.
                for name, default in (("top", 0.5), ("bottom", 0.5),
                                      ("left", 0.4), ("right", 0.4)):
                    var = getattr(self, f"photo_margin_{name}_var", None)
                    try:
                        margin = float(var.get()) if var is not None else default
                    except (TypeError, ValueError):
                        margin = default
                    setattr(ws.page_margins, name, margin)
            if mode == "RT":
                ws.page_margins.top = ws.page_margins.bottom = 0.25
                ws.page_margins.left = ws.page_margins.right = 0.25
                # 표 영역을 침범하지 않는 인쇄 머리글 우측에 페이지 번호를 표시한다.
                # &P: 현재 페이지, &N: 전체 페이지 (Excel header/footer field codes)
                ws.oddHeader.right.text = "Page    &P    of    &N"
                ws.oddHeader.right.font = "Arial"
                ws.oddHeader.right.size = 9
                # 페이지 번호를 표 상단 외곽선 바로 위에 배치한다.
                ws.page_margins.header = 0.55

        except Exception as e:
            print(f"[WARNING] Print settings failed: {e}")
            print(f"[WARNING] Print settings failed: {e}")

    def apply_custom_dimensions(self, ws, context):
        """[NEW] 사용자의 행/열 조절 설정을 파싱하여 적용 (모드별 설정 대응)"""
        try:
            # Determine current mode for key fallback
            mode = getattr(self, 'current_mode', "PMI")
            try:
                tab_text = self.mode_notebook.tab(self.mode_notebook.select(), "text")
                if "PMI" in tab_text: mode = "PMI"
                elif "RT" in tab_text: mode = "RT"
                elif "PT" in tab_text: mode = "PT"
                elif "MT" in tab_text: mode = "MT"
                elif "PAUT" in tab_text: mode = "PAUT"
            except: pass

            full_context = f"{mode}_{context}" if "_" not in context else context

            # Row Adjustment
            row_range_str = (self.config.get(f"CUSTOM_ROWS_{full_context}", "") or 
                             self.config.get(f"CUSTOM_ROWS_{context}", "")).strip()

            self.log(f"!!! [SPY-ENTRY] Sheet: {ws.title}, mode={mode}, ctx={context}, full={full_context}, rows='{row_range_str}'")

            if row_range_str:
                height_val = (self.config.get(f"CUSTOM_ROW_HEIGHT_{full_context}") or 
                              self.config.get(f"CUSTOM_ROW_HEIGHT_{context}", 16.5))
                height = float(height_val)
                self.log(f"!!! [SPY-HEIGHT] Sheet: {ws.title}, rows='{row_range_str}', h={height}")
                self.log(f"[DEBUG] Applying row height {height} to rows {row_range_str} for context {full_context}")
                import re
                for part in row_range_str.split(','):
                    part = part.strip()
                    if not part: continue
                    # [SMART PARSE] 'A19:M45', '19-45', '19:45', '19~45' 등 모든 형식에서 행 번호 숫자만 자동 추출
                    nums = [int(n) for n in re.findall(r'\d+', part)]
                    if len(nums) >= 2:
                        start, end = nums[0], nums[1]
                        if start > end: start, end = end, start
                        for r in range(start, end + 1):
                            ws.row_dimensions[r].height = height
                    elif len(nums) == 1:
                        ws.row_dimensions[nums[0]].height = height

            # Column Adjustment
            col_range_str = (self.config.get(f"CUSTOM_COLS_{full_context}", "") or 
                             self.config.get(f"CUSTOM_COLS_{context}", "")).strip()
            if col_range_str:
                width_val = (self.config.get(f"CUSTOM_COL_WIDTH_{full_context}") or 
                             self.config.get(f"CUSTOM_COL_WIDTH_{context}", 10.0))
                width = float(width_val)
                for part in col_range_str.split(','):
                    part = part.strip()
                    if not part: continue
                    sep = '-' if '-' in part else (':' if ':' in part else ('~' if '~' in part else None))
                    if sep:
                        bits = part.split(sep)
                        if len(bits) == 2:
                            import re
                            start_l = re.sub(r'[^A-Z]', '', bits[0].strip().upper())
                            end_l = re.sub(r'[^A-Z]', '', bits[1].strip().upper())
                            if start_l and end_l:
                                for c in range(column_index_from_string(start_l), column_index_from_string(end_l) + 1):
                                    ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = width
                    else:
                        import re
                        c_clean = re.sub(r'[^A-Z]', '', part.upper())
                        if c_clean:
                            ws.column_dimensions[c_clean].width = width
        except Exception as e:
            print(f"[DEBUG] Custom dimension failed for {context}: {e}")

    def safe_set_value(self, ws, coord, value, align=None):
        """병합된 셀이라도 기준 셀(Top-Left)을 찾아 안전하게 값을 입력합니다. (성능 최적화 버전)"""
        try:
            # 전달받은 인자가 이미 셀 객체인 경우
            if hasattr(coord, 'coordinate'):
                target_cell = coord
                coord_str = target_cell.coordinate
            else:
                target_cell = ws[coord]
                coord_str = coord

            # 해당 셀이 MergedCell(병합 영역의 기준 셀이 아닌 곳)인 경우에만 탐색 수행
            if isinstance(target_cell, MergedCell):
                for m_range in ws.merged_cells.ranges:
                    if coord_str in m_range:
                        target_cell = ws.cell(row=m_range.min_row, column=m_range.min_col)
                        break

            target_cell.value = value
            if align:
                target_cell.alignment = Alignment(horizontal=align, vertical='center')
        except Exception: pass

    def safe_merge_cells(self, ws, start_row, start_column, end_row, end_column):
        """[FIX] Prevent unmerging or conflicting merges on the cover sheet and prevent overlapping merges"""
        try:
            # If it's the cover sheet, be extremely conservative with merging/unmerging
            if ws == ws.parent.worksheets[0]:
                return
        except: pass

        try:
            # 대상 범위와 겹치는 기존 병합 셀이 있으면 unmerge하여 OpenXML 손상(중복 병합 오류) 방지
            for m in list(ws.merged_cells.ranges):
                if not (m.max_row < start_row or m.min_row > end_row or
                        m.max_col < start_column or m.min_col > end_column):
                    if (m.min_row == start_row and m.max_row == end_row and
                        m.min_col == start_column and m.max_col == end_column):
                        return
                    try:
                        ws.unmerge_cells(str(m))
                    except Exception:
                        pass

            ws.merge_cells(start_row=start_row, start_column=start_column, end_row=end_row, end_column=end_column)
        except Exception as e:
            self.log(f"   ⚠️ 셀 병합 실패 ({start_row},{start_column}): {e}")

    def _write_pt_total_summary(self, ws, summary_row, items):
        """마지막 PT 빈 행 하나에 TOTAL/Acc/Rej/BLANK 요약을 기록한다."""
        accepted_values = {'ACC', 'ACCEPT', 'PASS', '합격', 'O', 'OK', 'V', ''}
        accepted_count = sum(
            1 for item in items
            if str(item.get('Result', 'Acc')).strip().upper() in accepted_values
        )
        rejected_count = max(0, len(items) - accepted_count)

        # PT 을지 열 구조: A:C Identification, D Joint, E Acc, F Rej,
        # G Step, H:Q Interpretation, R Welder, S NPS.
        self.safe_merge_cells(ws, summary_row, 1, summary_row, 3)
        self.safe_merge_cells(ws, summary_row, 8, summary_row, 17)

        values = {
            (summary_row, 1): 'TOTAL',
            (summary_row, 5): f'{accepted_count}P',
            (summary_row, 6): rejected_count,
            (summary_row, 8): 'B  L  A  N  K',
        }
        for (row_idx, col_idx), value in values.items():
            cell = ws.cell(row=row_idx, column=col_idx)
            self.safe_set_value(ws, cell.coordinate, value, align='center')
            cell.font = Font(name='바탕', size=9, bold=False)

    def set_eulji_headers(self, ws):
        # [FIX] RT 모드일 때는 헤더를 자동으로 쓰지 않음 (템플릿 보존)
        if getattr(self, 'current_mode', "") == "RT":
            return

        headers = ["NI", "CR", "MO"]
        data_font = Font(size=9); header_row = int(self.config.get('START_ROW', 10))
        for i, val in enumerate(headers):
            col = 8 + i
            cell = ws.cell(row=header_row, column=col)
            self.safe_set_value(ws, cell.coordinate, val, align='center')
            cell.font = data_font

        materials = "SS304,SS304L,SS316,SS316L,SS321,SS347,SS410,SS430,DUPLEX,MONEL,INCONEL,ER308,ER308L,ER309,ER309L,ER316,ER316L,ER347,ER2209,WP316,WP316L,TP316,TP316L,F316L,A182-F316L,A312-TP316L"
        dv_q = DataValidation(type="list", formula1=f'"{materials}"', allow_blank=True)
        ws.add_data_validation(dv_q)
        for r in range(int(self.config.get('START_ROW', 17)), int(self.config.get('DATA_END_ROW', 45)) + 1):
            target_l = ws.cell(row=r, column=13); dv_q.add(target_l) # 12 -> 13 (M)
            target_l.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center'); target_l.font = Font(size=8.5)

        # [NEW] 플랜트 3단 모드에서만 K15에 하이픈 적용 (표준 20열 템플릿의 H15:M16 'Chemical Composition' 훼손 방지)
        is_triple = hasattr(self, 'pmi_layout_mode') and self.pmi_layout_mode.get() == "PLANT_TRIPLE"
        if is_triple:
            self.safe_set_value(ws, 'K15', "-")
            ws['K15'].alignment = Alignment(horizontal='center', vertical='center')
        else:
            # 표준 20열 모드: H15:M16에 'Chemical Composition' 복원 및 보존
            h15_val = str(ws['H15'].value or '').strip()
            if not h15_val or h15_val == '-':
                self.safe_set_value(ws, 'H15', "Chemical Composition")
                ws['H15'].alignment = Alignment(horizontal='center', vertical='center')
                ws['H15'].font = Font(name='바탕', size=9, bold=False)

    def convert_pmi_eulji_to_plant_triple(self, sheet, ws0=None):
        """일반 단일행 20열 PMI 을지 시트를 플랜트 배관용 13열 3단 구조(사진 3)로 동적 변환합니다.

        - 1~10행 상단 헤더: 제목(A1:M1, A2:M2), 고객사(K5:L7, M5:M7), 보고서번호(K8:L10, M8:M10) 재정렬
        - 11~18행 검사 헤더: Drawing No(A11:E18), Joint No(F11:F18), Test Location(G11:G18),
          Method(H11:L12), Analyzer(H13:L13), Chem Comp(H14:L14), 원소행(H15~L15),
          규격치 3행(H16:L16 모재1, H17:L17 용접부, H18:L18 모재2), Type of Material(M11:M18)
        - 14~20열(N~T열): 미사용 컬럼 내용/테두리 클리어
        - 열 너비 및 인쇄영역: 13열 플랜트 배관용 규격(A1:M47)으로 최적화
        """
        medium_side = Side(style='medium')
        thin_side = Side(style='thin')
        from openpyxl.cell.cell import MergedCell

        # 1. 14열(N열) 이상에 걸쳐 있는 병합 및 1~18행 병합 완전 해제
        for m in list(sheet.merged_cells.ranges):
            if m.max_col >= 14 or m.max_row <= 18:
                try: sheet.unmerge_cells(str(m))
                except: pass

        # 2. 열 너비 13열 최적화 (갑지 총합 91.7과 100% 일치시켜 좌우 여백 및 테두리 완벽 일치)
        col_widths = {
            'A': 4.5, 'B': 4.5, 'C': 3.0, 'D': 3.0,
            'E': 5.5, 'F': 7.0, 'G': 8.5, 'H': 6.5, 'I': 7.5, 'J': 7.5,
            'K': 7.5, 'L': 7.5, 'M': 19.2
        }
        for col_letter, w in col_widths.items():
            sheet.column_dimensions[col_letter].width = w
            sheet.column_dimensions[col_letter].hidden = False
        for c_idx in range(14, 26):
            col_l = openpyxl.utils.get_column_letter(c_idx)
            sheet.column_dimensions[col_l].hidden = True
            sheet.column_dimensions[col_l].width = 0

        # 3. 1~10행 상단 헤더 재구성 (첫 번째 사진과 100% 동일한 구조와 외관)
        self.safe_merge_cells(sheet, 1, 1, 1, 13)
        self.safe_merge_cells(sheet, 2, 1, 2, 13)
        self.safe_merge_cells(sheet, 3, 12, 3, 13)

        # A5:D10 (SITCO 로고 영역)
        self.safe_merge_cells(sheet, 5, 1, 10, 4)

        # E8:J8, E9:J9, E10:J10 서울검사 정보 (첫 번째 사진과 완벽 일치: 좌측 정렬, 바탕 9pt)
        co_info = [
            (8, 'SEOUL INSPECTION & TESTING Co., Ltd.'),
            (9, '서울특별시  서초구  바우뫼로  41길  54 (용진빌딩)'),
            (10, 'TEL : (02) 552-1112      FAX : (02) 2058-0720')
        ]
        font_addr = Font(name='바탕', size=9, bold=False)
        for r, txt in co_info:
            self.safe_set_value(sheet, f'E{r}', txt, align='left')
            sheet.cell(r, 5).font = font_addr
            sheet.cell(r, 5).alignment = Alignment(horizontal='left', vertical='center')
            self.safe_merge_cells(sheet, r, 5, r, 10)

        font_label = Font(name='바탕', size=9, bold=False)
        # K5:L7, K8:L10 (주문주 / 보고서 번호 라벨: 2줄 wrap_text)
        self.safe_set_value(sheet, 'K5', '주문주\nCustomer', align='center')
        sheet.cell(5, 11).font = font_label
        sheet.cell(5, 11).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        self.safe_merge_cells(sheet, 5, 11, 7, 12)

        self.safe_set_value(sheet, 'K8', '보고서 번호\nReport No.', align='center')
        sheet.cell(8, 11).font = font_label
        sheet.cell(8, 11).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        self.safe_merge_cells(sheet, 8, 11, 10, 12)

        # M5:M7 (고객사명), M8:M10 (보고서 번호값)
        cust = (self.gapji_customer.get().strip() if hasattr(self, 'gapji_customer') else '') or self.config.get('GAPJI_CUSTOMER', '')
        if cust:
            self.safe_set_value(sheet, 'M5', cust, align='center')
        elif ws0 is not None:
            self.safe_set_value(sheet, 'M5', f"='{ws0.title}'!N5", align='center')
        sheet.cell(5, 13).font = font_label
        sheet.cell(5, 13).alignment = Alignment(horizontal='center', vertical='center', shrink_to_fit=True)
        self.safe_merge_cells(sheet, 5, 13, 7, 13)

        rep = (self.gapji_report_no.get().strip() if hasattr(self, 'gapji_report_no') else '') or self.config.get('GAPJI_REPORT_NO', '')
        if rep:
            self.safe_set_value(sheet, 'M8', rep, align='center')
        elif ws0 is not None:
            self.safe_set_value(sheet, 'M8', f"=IF('{ws0.title}'!N8=\"\",\"\",'{ws0.title}'!N8)", align='center')
        sheet.cell(8, 13).font = font_label
        sheet.cell(8, 13).alignment = Alignment(horizontal='center', vertical='center', shrink_to_fit=True)
        self.safe_merge_cells(sheet, 8, 13, 10, 13)

        # 5~10행 테두리: 첫 번째 사진과 100% 동일 (E~J 내부 가로선 일절 없음)
        # A5:D10 (SITCO)
        for r in range(5, 11):
            for c in range(1, 5):
                cell = sheet.cell(r, c)
                l = medium_side if c == 1 else None
                t = medium_side if r == 5 else None
                b = medium_side if r == 10 else None
                cell.border = Border(left=l, right=None, top=t, bottom=b)

        # E5:J10 (서울검사 상호 + 주소: 내부 가로선 없음!)
        for r in range(5, 11):
            for c in range(5, 11):
                cell = sheet.cell(r, c)
                r_s = thin_side if c == 10 else None
                t = medium_side if r == 5 else None
                b = medium_side if r == 10 else None
                cell.border = Border(left=None, right=r_s, top=t, bottom=b)

        # K5:L10 (주문주/보고서번호 라벨)
        for r in range(5, 11):
            for c in [11, 12]:
                cell = sheet.cell(r, c)
                l = thin_side if c == 11 else None
                r_s = thin_side if c == 12 else None
                t = medium_side if r == 5 else (thin_side if r == 8 else None)
                b = medium_side if r == 10 else (thin_side if r == 7 else None)
                cell.border = Border(left=l, right=r_s, top=t, bottom=b)

        # M5:M10 (고객사/보고서번호값)
        for r in range(5, 11):
            cell = sheet.cell(r, 13)
            l = thin_side
            r_s = medium_side
            t = medium_side if r == 5 else (thin_side if r == 8 else None)
            b = medium_side if r == 10 else (thin_side if r == 7 else None)
            cell.border = Border(left=l, right=r_s, top=t, bottom=b)

        # 4. 11~18행 헤더 변환 (세 번째 사진과 100% 일치)
        for r in range(11, 19):
            for c in range(1, 26):
                cell = sheet.cell(row=r, column=c)
                cell.value = None
                cell.border = Border()

        font_bold = Font(name='바탕', size=9, bold=True)
        font_spec = Font(name='바탕', size=8, bold=False)

        # A11:E18 Drawing No.
        self.safe_set_value(sheet, 'A11', 'Drawing No.', align='center')
        sheet.cell(11, 1).font = font_bold
        self.safe_merge_cells(sheet, 11, 1, 18, 5)

        # F11:F18 Joint No.
        self.safe_set_value(sheet, 'F11', 'Joint No.', align='center')
        sheet.cell(11, 6).font = font_bold
        self.safe_merge_cells(sheet, 11, 6, 18, 6)

        # G11:G18 Test Location
        self.safe_set_value(sheet, 'G11', 'Test Location', align='center')
        sheet.cell(11, 7).font = font_bold
        sheet.cell(11, 7).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        self.safe_merge_cells(sheet, 11, 7, 18, 7)

        # H11:L12 Alloy Verification Method / 합금분석시험방법
        self.safe_set_value(sheet, 'H11', 'Alloy Verification Method\n합금분석시험방법', align='center')
        sheet.cell(11, 8).font = font_bold
        sheet.cell(11, 8).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        self.safe_merge_cells(sheet, 11, 8, 12, 12)

        # H13:L13 Alloy Analyzer재질분석기
        self.safe_set_value(sheet, 'H13', 'Alloy Analyzer재질분석기', align='center')
        sheet.cell(13, 8).font = font_bold
        self.safe_merge_cells(sheet, 13, 8, 13, 12)

        # H14:L14 Chemical Composition
        self.safe_set_value(sheet, 'H14', 'Chemical Composition', align='center')
        sheet.cell(14, 8).font = font_bold
        self.safe_merge_cells(sheet, 14, 8, 14, 12)

        # 15행 원소기호 (H15~L15 개별 셀)
        chem_elems = ["Ni", "Cr", "Mo", "-", "-"]
        if hasattr(self, 'pmi_chem_elements') and self.pmi_chem_elements:
            for idx, v in enumerate(self.pmi_chem_elements[:5]):
                val = v.get().strip()
                if val: chem_elems[idx] = val
        for idx, el in enumerate(chem_elems):
            c_target = 8 + idx
            self.safe_set_value(sheet, sheet.cell(15, c_target).coordinate, el, align='center')
            sheet.cell(15, c_target).font = font_bold

        # 16~18행 규격치 3행 (H16:L16, H17:L17, H18:L18)
        default_specs = ["-", "-", "-"]
        if hasattr(self, 'pmi_chem_specs') and self.pmi_chem_specs:
            for idx, v in enumerate(self.pmi_chem_specs[:3]):
                val = v.get().strip()
                if val: default_specs[idx] = val
        for idx, sp in enumerate(default_specs):
            r_target = 16 + idx
            self.safe_set_value(sheet, f'H{r_target}', sp, align='center')
            sheet.cell(r_target, 8).font = font_spec
            self.safe_merge_cells(sheet, r_target, 8, r_target, 12)

        # M11:M18 Type of Material
        self.safe_set_value(sheet, 'M11', 'Type of Material', align='center')
        sheet.cell(11, 13).font = font_bold
        sheet.cell(11, 13).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        self.safe_merge_cells(sheet, 11, 13, 18, 13)

        # 11~18행 테두리 및 행 높이 설정
        for r in range(11, 19):
            sheet.row_dimensions[r].height = 19.5
            for c in range(1, 14):
                cell = sheet.cell(row=r, column=c)
                t_s = medium_side if r == 11 else thin_side
                b_s = medium_side if r == 18 else thin_side
                l_s = medium_side if c == 1 else thin_side
                r_s = medium_side if c == 13 else thin_side

                # A~E (Drawing No) 내부 선 정리
                if 1 <= c <= 5:
                    if 1 < c < 5: l_s = None; r_s = None
                    elif c == 1: r_s = None
                    elif c == 5: l_s = None
                    if 11 < r < 18: t_s = None; b_s = None
                    elif r == 11: b_s = None
                    elif r == 18: t_s = None
                elif c == 6:
                    if 11 < r < 18: t_s = None; b_s = None
                    elif r == 11: b_s = None
                    elif r == 18: t_s = None
                elif c == 7:
                    if 11 < r < 18: t_s = None; b_s = None
                    elif r == 11: b_s = None
                    elif r == 18: t_s = None
                elif 8 <= c <= 12:
                    if r in (11, 12):
                        if 8 < c < 12: l_s = None; r_s = None
                        elif c == 8: r_s = None
                        elif c == 12: l_s = None
                        if r == 11: b_s = None
                        elif r == 12: t_s = None
                    elif r in (13, 14, 16, 17, 18):
                        if 8 < c < 12: l_s = None; r_s = None
                        elif c == 8: r_s = None
                        elif c == 12: l_s = None
                elif c == 13:
                    if 11 < r < 18: t_s = None; b_s = None
                    elif r == 11: b_s = None
                    elif r == 18: t_s = None

                cell.border = Border(left=l_s, right=r_s, top=t_s, bottom=b_s)

        # 5. 14~25열 잔여 셀 내용 및 테두리 클리어 (1~50행)
        for r in range(1, 52):
            for c in range(14, 26):
                cell = sheet.cell(row=r, column=c)
                if not isinstance(cell, MergedCell):
                    cell.value = None
                    cell.border = Border()

        # 6. 인쇄 영역 및 인쇄 설정 최적화 (배관용 19~45행 27행 9조인트 규격: A1:M45)
        sheet.print_area = "A1:M45"
        sheet.page_setup.orientation = 'portrait'
        sheet.page_setup.paperSize = 9
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
        sheet.page_setup.fitToWidth = 1
        sheet.page_setup.fitToHeight = 1
        sheet.page_setup.scale = None
        sheet.page_margins.left = 0.25
        sheet.page_margins.right = 0.25
        sheet.page_margins.top = 0.25
        sheet.page_margins.bottom = 0.25

    def prepare_next_sheet(self, wb, source_sheet_idx, page_num, mode=None):
        source_sheet = wb.worksheets[source_sheet_idx]; new_sheet = wb.copy_worksheet(source_sheet) 
        current_mode = mode or getattr(self, 'current_mode', '')

        # [FIX] openpyxl's copy_worksheet does not copy images. Manually copy them for ALL modes to preserve template logos.
        try:
            import io
            import copy
            from openpyxl.drawing.image import Image as OpenpyxlImage
            for idx, img in enumerate(source_sheet._images):
                try:
                    img_bytes = img._data()

                    # Fix original image
                    fixed_orig = OpenpyxlImage(io.BytesIO(img_bytes))
                    if hasattr(img, 'anchor'): fixed_orig.anchor = copy.copy(img.anchor)
                    source_sheet._images[idx] = fixed_orig

                    # Add to new sheet
                    new_img = OpenpyxlImage(io.BytesIO(img_bytes))
                    if hasattr(img, 'anchor'): new_img.anchor = copy.copy(img.anchor)
                    new_sheet.add_image(new_img)
                except: pass
        except: pass

        if current_mode == 'MT':
            from openpyxl.cell.cell import MergedCell
            new_sheet.title = f'MT(을)_{page_num:03}'
            for row in new_sheet.iter_rows(min_row=10, max_row=38, max_col=24):
                for cell in row:
                    if not isinstance(cell, MergedCell):
                        cell.value = None
            return new_sheet

        # 병합된 셀 명시적으로 다시 복사 (openpyxl 버그 방지)
        try:
            for merged_cell_range in source_sheet.merged_cells.ranges:
                new_sheet.merge_cells(str(merged_cell_range))
        except: pass

        # 인쇄 설정 완벽 복사 (여백, 페이지 설정, 프린트 옵션)
        try:
            import copy
            new_sheet.page_setup = copy.copy(source_sheet.page_setup)
            new_sheet.page_margins = copy.copy(source_sheet.page_margins)
            new_sheet.print_options = copy.copy(source_sheet.print_options)
            if hasattr(source_sheet, 'sheet_format') and source_sheet.sheet_format:
                new_sheet.sheet_format = copy.copy(source_sheet.sheet_format)
            # 열 너비와 행 높이를 A~Z, 1~60행까지 강제 복사하여 여백 불일치 원천 차단
            from openpyxl.utils import get_column_letter

            def_col_w = source_sheet.sheet_format.defaultColWidth if (hasattr(source_sheet, 'sheet_format') and source_sheet.sheet_format and source_sheet.sheet_format.defaultColWidth) else 8.38
            for i in range(1, 27):
                # copy_worksheet already preserves PMI grouped dimensions such
                # as A:D. Creating separate default-width columns makes copied
                # pages wider than the original data sheet.
                if current_mode == "PMI":
                    continue
                col_letter = get_column_letter(i)
                if col_letter in source_sheet.column_dimensions and source_sheet.column_dimensions[col_letter].width:
                    new_sheet.column_dimensions[col_letter].width = source_sheet.column_dimensions[col_letter].width
                else:
                    new_sheet.column_dimensions[col_letter].width = def_col_w

            def_row_h = source_sheet.sheet_format.defaultRowHeight if (hasattr(source_sheet, 'sheet_format') and source_sheet.sheet_format and source_sheet.sheet_format.defaultRowHeight) else 15.0
            for r in range(1, 60):
                if current_mode == "PMI":
                    continue
                if r in source_sheet.row_dimensions and source_sheet.row_dimensions[r].height:
                    new_sheet.row_dimensions[r].height = source_sheet.row_dimensions[r].height
                else:
                    new_sheet.row_dimensions[r].height = def_row_h
        except: pass

        base_title = source_sheet.title.split('_')[0]; new_sheet.title = f"{base_title[:20]}_{page_num:03d}"
        self.force_print_settings(new_sheet, context="DATA")

        # [FIX] Only add logos if they weren't already copied from the source sheet
        if not new_sheet._images:
            self.add_logos_to_sheet(new_sheet, is_cover=False, mode=current_mode or None)
        self.apply_custom_dimensions(new_sheet, "DATA")
        # [REMOVED] Second column width overwrite removed because we already enforced it powerfully above.
        data_font = Font(size=9); grade_font = Font(size=8.5)
        for r in range(int(self.config.get('START_ROW', 17)), int(self.config.get('DATA_END_ROW', 45)) + 1):
            rd = new_sheet.row_dimensions[r] # [REMOVED] Hardcoded 20.55 override

        # [FIX] RT 모드일 경우 Shooting Sketch(32행~42행)를 보호하기 위해 데이터 종료 행을 조절
        start_row = int(self.config.get('START_ROW', 17))
        end_row = int(self.config.get('DATA_END_ROW', 45))

        if current_mode == "RT":
            start_row = int(self.config.get('RT_START_ROW', 11))
            end_row = int(self.config.get('RT_DATA_END_ROW', 34))
        elif current_mode == "KOGAS":
            start_row = int(self.config.get('KOGAS_START_ROW', 14))
            end_row = int(self.config.get('KOGAS_DATA_END_ROW', 25))
        elif current_mode == "PT":
            start_row = int(self.config.get('PT_START_ROW', 10))
            end_row = int(self.config.get('PT_DATA_END_ROW', 38))

        # Unmerge the copied data area before clearing it. Clearing through
        # safe_set_value while cells are still merged repeatedly targets only
        # the merge anchor and can leave copied values in the later rows.
        merged_to_clear = [rng for rng in new_sheet.merged_cells.ranges if rng.min_row >= start_row and rng.max_row <= end_row]
        # Keep a copy of the string ranges to restore them later
        ranges_to_restore = [str(rng) for rng in merged_to_clear]

        for rng in merged_to_clear:
            try: new_sheet.unmerge_cells(str(rng))
            except: pass

        for r in range(start_row, end_row + 1):
            # Clear up to column 26 (Z) to ensure H~Q and others are fully cleared
            for c in range(1, 26):
                cell = new_sheet.cell(row=r, column=c)
                cell.value = None
                if c < 14:
                    cell.font = grade_font if c == 13 else data_font

        # [FIX] Restore the merges that were unmerged for clearing
        # This fixes the H19~Q38 and A19~C38 merging issues for empty rows
        for rng_str in ranges_to_restore:
            try: new_sheet.merge_cells(rng_str)
            except: pass

        # [FIX] 갑지 데이터 수식으로 연결 (첫번째 시트 참조)
        try:
            ws0 = wb.worksheets[0]
            if len(wb.worksheets) > 1 and current_mode not in ("RT", "KOGAS"):
                if current_mode == "PT":
                    # PT 모드는 P4, P6에 텍스트를 직접 꽂아줌 (수식 오류로 0 표기 방지)
                    self.safe_set_value(new_sheet, 'P4', self.gapji_customer.get().strip())
                    self.safe_set_value(new_sheet, 'P6', self.gapji_report_no.get().strip())
                elif current_mode == "PMI":
                    is_triple_mode = hasattr(self, 'pmi_layout_mode') and self.pmi_layout_mode.get() == "PLANT_TRIPLE"
                    if is_triple_mode:
                        self.safe_set_value(new_sheet, 'K5', f"='{ws0.title}'!L5")
                        self.safe_set_value(new_sheet, 'M5', f"='{ws0.title}'!N5")
                        self.safe_set_value(new_sheet, 'M8', f"='{ws0.title}'!N8")
                        new_sheet.column_dimensions['M'].width = 18.0
                    else:
                        # 20열 표준 단일행 모드: K5 정리, M5/M8 라벨 보존, O5/O8 연동, 열 너비 보존
                        self.safe_set_value(new_sheet, 'K5', None)
                        self.safe_set_value(new_sheet, 'M5', '주문주\nCustomer')
                        self.safe_set_value(new_sheet, 'O5', f"='{ws0.title}'!N5")
                        self.safe_set_value(new_sheet, 'M8', '보고서 번호\nReport No.')
                        self.safe_set_value(new_sheet, 'O8', f"=IF('{ws0.title}'!N8=\"\",\"\",'{ws0.title}'!N8)")
                        pmi_20col_widths = {
                            'H': 4.83, 'I': 4.83, 'J': 4.61, 'K': 4.61, 'L': 4.61, 'M': 4.61,
                            'N': 11.00, 'O': 6.00
                        }
                        for col_l, col_w in pmi_20col_widths.items():
                            new_sheet.column_dimensions[col_l].width = col_w
                else:
                    # 엑셀 수식을 사용하여 갑지의 값이 바뀌면 자동으로 바뀌게 설정
                    self.safe_set_value(new_sheet, 'K5', f"='{ws0.title}'!L5")
                    self.safe_set_value(new_sheet, 'M5', f"='{ws0.title}'!N5")
                    self.safe_set_value(new_sheet, 'M8', f"='{ws0.title}'!N8")
                # [FIX] K5:M10 범위 글씨체 바탕, 크기 9 적용
                for r_idx in range(5, 11):
                    for c_idx in range(11, 14): # K(11) ~ M(13)
                        cell = new_sheet.cell(row=r_idx, column=c_idx)
                        cell.font = Font(name='바탕', size=9, bold=False)
                        # K5, K8는 줄바꿈 적용, 그 외는 셀 크기에 맞춤 적용
                        if (r_idx == 5 or r_idx == 8) and c_idx == 11:
                            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                        else:
                            cell.alignment = Alignment(horizontal='center', vertical='center', shrink_to_fit=True)
        except: pass

        # [NEW] 플랜트 3단 모드에서만 K15 하이픈 및 M열 18.0 적용
        if current_mode == "PMI" and hasattr(self, 'pmi_layout_mode') and self.pmi_layout_mode.get() == "PLANT_TRIPLE":
            self.safe_set_value(new_sheet, 'K15', "-")
            new_sheet['K15'].alignment = Alignment(horizontal='center', vertical='center')
            new_sheet.column_dimensions['M'].width = 18.0
        elif current_mode == "PMI":
            h15_val = str(new_sheet['H15'].value or '').strip()
            if not h15_val or h15_val == '-':
                self.safe_set_value(new_sheet, 'H15', "Chemical Composition")
                new_sheet['H15'].alignment = Alignment(horizontal='center', vertical='center')
                new_sheet['H15'].font = Font(name='바탕', size=9, bold=False)
        return new_sheet

    def inject_drawing_layer(self, template_path, target_path):
        """
        [FINAL V4] Definitive hybrid surgery.
        - Report's workbook.xml, workbook.xml.rels, [Content_Types].xml → correct structure, print area, all pages
        - Template's drawings, media, sheet structure → drawings preserved
        - Report's sheetData injected into template's sheet XMLs → content preserved
        - Blueprint for extra sheets (multipage RT)
        """
        import zipfile, shutil, os, re, traceback
        from site_apps.central.src.reports.worksheet_xml import normalize_print_options

        final_out = target_path + ".final"
        shutil.copy2(template_path, final_out)
        temp_swap = final_out + ".swap"

        try:
            # ── Step 1: 리포트에서 필요한 모든 데이터 수집 ──────────────────
            with zipfile.ZipFile(target_path, 'r') as z_rep:
                rep_names = z_rep.namelist()

                # sharedStrings 읽기
                rep_shared_strings = []
                ss_f = next((f for f in rep_names if f.lower() == 'xl/sharedstrings.xml'), None)
                if ss_f:
                    try:
                        ss_xml = z_rep.read(ss_f).decode('utf-8', errors='ignore')
                        rep_shared_strings = re.findall(r'<t(?:\s[^>]*)?>([^<]*)</t>', ss_xml)
                    except: pass

                # 시트 데이터 추출 (sharedStr → inlineStr 변환)
                processed_sheets = {}
                for f in rep_names:
                    if f.lower().startswith('xl/worksheets/sheet') and f.lower().endswith('.xml'):
                        try:
                            content = z_rep.read(f).decode('utf-8', errors='ignore')
                            def _to_inline(m, _ss=rep_shared_strings):
                                try:
                                    idx = int(m.group(2))
                                    if 0 <= idx < len(_ss):
                                        t = _ss[idx].replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
                                        return f'<c {m.group(1).replace(chr(116)+"=\"s\"","t=\"inlineStr\"")}><is><t>{t}</t></is>'
                                except: pass
                                return m.group(0)
                            content = re.sub(r'<c ([^>]*t="s"[^>]*)>\s*<v>(\d+)</v>', _to_inline, content)
                            processed_sheets[f.lower()] = content
                        except: pass

                # 리포트 인쇄 설정 추출 (force_print_settings 적용값)
                rep_page_settings = {}  # fl_lower → setup/margins/options/header
                rep_scale = None
                for fl_s, c in processed_sheets.items():
                    setup_m  = re.search(r'<pageSetup\b[^>]*/>', c)
                    margin_m = re.search(r'<pageMargins\b[^>]*/>', c)
                    options_m = re.search(r'<printOptions\b[^>]*/>', c)
                    header_m = re.search(
                        r'<headerFooter\b[^>]*>.*?</headerFooter>|<headerFooter\b[^>]*/>',
                        c, flags=re.DOTALL
                    )
                    scale_m  = re.search(r'<pageSetup\b[^>]*\bscale="(\d+)"', c)
                    if scale_m and not rep_scale:
                        rep_scale = scale_m.group(1)
                    rep_page_settings[fl_s] = {
                        'setup':   setup_m.group(0)  if setup_m  else None,
                        'margins': margin_m.group(0) if margin_m else None,
                        'options': options_m.group(0) if options_m else None,
                        'header':  header_m.group(0)  if header_m  else None,
                    }

                # 리포트의 workbook 메타 파일들 (raw bytes)
                rep_workbook_xml  = next((z_rep.read(f) for f in rep_names if f.lower() == 'xl/workbook.xml'), None)
                rep_wb_rels       = next((z_rep.read(f) for f in rep_names if f.lower() == 'xl/_rels/workbook.xml.rels'), None)
                rep_content_types = next((z_rep.read(f) for f in rep_names if f.lower() == '[content_types].xml'), None)
                rep_styles        = next((z_rep.read(f) for f in rep_names if f.lower() == 'xl/styles.xml'), None)
                rep_ss_data       = next((z_rep.read(f) for f in rep_names if f.lower() == 'xl/sharedstrings.xml'), None)

            # ── Step 2: 템플릿에서 드로잉/Blueprint 수집 ────────────────────
            with zipfile.ZipFile(final_out, 'r') as z_tmpl:
                blueprint_xml  = None
                blueprint_rels = None
                template_sheets = sorted(
                    (f for f in z_tmpl.namelist()
                     if re.fullmatch(r'xl/worksheets/sheet\d+\.xml', f, re.IGNORECASE)),
                    key=lambda f: int(re.search(r'sheet(\d+)\.xml', f, re.IGNORECASE).group(1))
                )
                # Continuation pages must inherit the data sheet, never the cover.
                first_sheet_f = (template_sheets[1] if len(template_sheets) > 1
                                 else template_sheets[0] if template_sheets else None)
                if first_sheet_f:
                    blueprint_xml = z_tmpl.read(first_sheet_f).decode('utf-8', errors='ignore')
                    rel_p = f"xl/worksheets/_rels/{os.path.basename(first_sheet_f)}.rels"
                    if rel_p in z_tmpl.namelist():
                        blueprint_rels = z_tmpl.read(rel_p)

                # 템플릿 파일 맵 (소문자 경로 → 파일 내용)
                tmpl_map = {}
                tmpl_names = {}
                for item in z_tmpl.infolist():
                    tmpl_map[item.filename.lower()] = z_tmpl.read(item.filename)
                    tmpl_names[item.filename.lower()] = item.filename

            # ── Step 3: 출력 ZIP 구성 ────────────────────────────────────────
            with zipfile.ZipFile(temp_swap, 'w', compression=zipfile.ZIP_DEFLATED) as z_out:

                # ── 인쇄영역 강제 확정 주입 (시트별 구분: 갑지 vs 을지)
                if rep_workbook_xml:
                    wb_str = rep_workbook_xml.decode('utf-8', errors='ignore')
                    snames = re.findall(r'<sheet\b[^>]*\bname="([^"]+)"', wb_str)

                    # 기존 Print_Area definedName 전체 제거
                    wb_str = re.sub(r'<definedName[^>]*name="_xlnm\.Print_Area"[^/]*/>', '', wb_str)
                    wb_str = re.sub(r'<definedName[^>]*name="_xlnm\.Print_Area"[^>]*>.*?</definedName>\s*',
                                    '', wb_str, flags=re.DOTALL)

                    def _get_norm_pa(pa_str):
                        return re.sub(r'\$?([A-Za-z]+)\$?(\d+)', lambda m2: f'${m2.group(1).upper()}${m2.group(2)}', pa_str)

                    new_dns_list = []
                    for i, s_name in enumerate(snames):
                        if i == 0: # 갑지 (Cover)
                            pa = 'A1:V35'
                        else: # 을지 (Data)
                            pa = 'A1:V37'

                        pa_norm = _get_norm_pa(pa)
                        new_dns_list.append(f'<definedName name="_xlnm.Print_Area" localSheetId="{i}">\'{s_name}\'!{pa_norm}</definedName>')

                    new_dns = ''.join(new_dns_list)
                    if '<definedNames' in wb_str:
                        wb_str = re.sub(r'(<definedNames[^>]*>)', r'\1' + new_dns, wb_str, count=1)
                    elif '</workbook>' in wb_str:
                        wb_str = wb_str.replace('</workbook>', f'<definedNames>{new_dns}</definedNames></workbook>', 1)

                    rep_workbook_xml = wb_str.encode('utf-8')
                    self.log(f"   📐 [인쇄영역] 갑지/을지 개별 영역 주입 완료")

                # [Content_Types].xml: 템플릿 기반 + 리포트의 추가 시트 Override 병합
                # (템플릿 버전: drawing, media 등 포함 / 리포트 버전: 추가 시트 항목 포함)
                merged_ct = tmpl_map.get('[content_types].xml', b'')
                if merged_ct and rep_content_types:
                    tmpl_ct_str = merged_ct.decode('utf-8', errors='ignore')
                    rep_ct_str  = rep_content_types.decode('utf-8', errors='ignore')
                    # 템플릿에 없는 Override 항목을 리포트에서 추가
                    tmpl_parts  = set(re.findall(r'PartName="([^"]+)"', tmpl_ct_str))
                    for m_ov in re.finditer(r'<Override\b[^/]*/>', rep_ct_str):
                        pn_m = re.search(r'PartName="([^"]+)"', m_ov.group(0))
                        if pn_m and pn_m.group(1) not in tmpl_parts:
                            tmpl_ct_str = tmpl_ct_str.replace('</Types>', m_ov.group(0) + '</Types>', 1)
                    merged_ct = tmpl_ct_str.encode('utf-8')
                # The generated workbook determines the output format, even when
                # the drawing template has a different workbook content type.
                if merged_ct and rep_content_types:
                    import xml.etree.ElementTree as ET
                    merged_root = ET.fromstring(merged_ct)
                    report_root = ET.fromstring(rep_content_types)
                    workbook_type = next(
                        node.attrib['ContentType'] for node in report_root
                        if node.attrib.get('PartName') == '/xl/workbook.xml'
                    )
                    for node in merged_root:
                        if node.attrib.get('PartName') == '/xl/workbook.xml':
                            node.set('ContentType', workbook_type)
                    if not rep_ss_data:
                        for node in list(merged_root):
                            if node.attrib.get('PartName', '').lower() == '/xl/sharedstrings.xml':
                                merged_root.remove(node)
                    merged_ct = ET.tostring(merged_root, encoding='utf-8', xml_declaration=True)
                z_out.writestr('[Content_Types].xml', merged_ct if merged_ct else (rep_content_types or b''))

                # 리포트의 workbook 메타 파일 기록
                if rep_workbook_xml:  z_out.writestr('xl/workbook.xml',           rep_workbook_xml)
                if rep_wb_rels:       z_out.writestr('xl/_rels/workbook.xml.rels', rep_wb_rels)
                if rep_styles:        z_out.writestr('xl/styles.xml',              rep_styles)
                if rep_ss_data:       z_out.writestr('xl/sharedStrings.xml',       rep_ss_data)
                self.log("   📋 [workbook] 리포트 workbook 구조 적용")


                written = {'xl/workbook.xml', 'xl/_rels/workbook.xml.rels',
                           '[content_types].xml', 'xl/styles.xml', 'xl/sharedstrings.xml'}

                # 템플릿의 모든 파일 처리 (sheet, drawing, media 등)
                for fl_lower, raw in tmpl_map.items():
                    if fl_lower in written:
                        continue  # 이미 리포트 버전으로 기록됨

                    if fl_lower.startswith('xl/worksheets/sheet') and fl_lower.endswith('.xml'):
                        # 시트 XML: 템플릿 구조 + 리포트 데이터
                        try:
                            tmpl_sheet = raw.decode('utf-8', errors='ignore')
                            # rowBreaks 제거
                            tmpl_sheet = re.sub(r'<rowBreaks[^>]*>.*?</rowBreaks>', '', tmpl_sheet, flags=re.DOTALL)
                            tmpl_sheet = re.sub(r'<rowBreaks[^>]*/>', '', tmpl_sheet)
                            tmpl_sheet = re.sub(r'<colBreaks[^>]*>.*?</colBreaks>', '', tmpl_sheet, flags=re.DOTALL)
                            # pageSetup 전체 교체 (배율 + 방향 + 용지 크기)
                            rep_ps = rep_page_settings.get(fl_lower, {}).get('setup')
                            if rep_ps:
                                if re.search(r'<pageSetup\b[^>]*/>', tmpl_sheet):
                                    tmpl_sheet = re.sub(r'<pageSetup\b[^>]*/>', rep_ps, tmpl_sheet, count=1)
                                else:
                                    # pageSetup 태그가 없으면 sheetData 앞에 삽입
                                    tmpl_sheet = tmpl_sheet.replace('</sheetData>', '</sheetData>' + rep_ps, 1)
                            elif rep_scale:
                                # pageSetup은 없지만 scale만 있으면 기존 방식으로 패치
                                tmpl_sheet = re.sub(r'<pageSetup\b[^>]*/>', lambda m: re.sub(r'\bscale="\d+"', f'scale="{rep_scale}"', m.group(0)) if 'scale=' in m.group(0) else m.group(0).replace('<pageSetup', f'<pageSetup scale="{rep_scale}"', 1), tmpl_sheet)

                            # pageMargins 교체 (여백 설정)
                            rep_pm = rep_page_settings.get(fl_lower, {}).get('margins')
                            if rep_pm:
                                if re.search(r'<pageMargins\b[^>]*/>', tmpl_sheet):
                                    tmpl_sheet = re.sub(r'<pageMargins\b[^>]*/>', rep_pm, tmpl_sheet, count=1)
                                else:
                                    tmpl_sheet = tmpl_sheet.replace('</pageSetup>', '</pageSetup>' + rep_pm)

                            # 가로/세로 페이지 가운데 맞춤(printOptions)도 최종 파일에 보존
                            rep_po = rep_page_settings.get(fl_lower, {}).get('options')
                            if rep_po:
                                if re.search(r'<printOptions\b[^>]*/>', tmpl_sheet):
                                    tmpl_sheet = re.sub(r'<printOptions\b[^>]*/>', rep_po, tmpl_sheet, count=1)
                                elif '<pageMargins' in tmpl_sheet:
                                    tmpl_sheet = tmpl_sheet.replace('<pageMargins', rep_po + '<pageMargins', 1)
                                else:
                                    tmpl_sheet = tmpl_sheet.replace('</sheetData>', '</sheetData>' + rep_po, 1)

                            # 우측 상단 Page &P of &N 인쇄 머리글도 템플릿 병합 후 유지한다.
                            rep_header = rep_page_settings.get(fl_lower, {}).get('header')
                            if rep_header:
                                header_pattern = (
                                    r'<headerFooter\b[^>]*>.*?</headerFooter>|'
                                    r'<headerFooter\b[^>]*/>'
                                )
                                if re.search(header_pattern, tmpl_sheet, flags=re.DOTALL):
                                    tmpl_sheet = re.sub(
                                        header_pattern, rep_header, tmpl_sheet,
                                        count=1, flags=re.DOTALL
                                    )
                                elif re.search(r'<pageSetup\b[^>]*/>', tmpl_sheet):
                                    tmpl_sheet = re.sub(
                                        r'(<pageSetup\b[^>]*/>)', r'\1' + rep_header,
                                        tmpl_sheet, count=1
                                    )
                            # sheetData 주입 + 행 높이 보존
                            if fl_lower in processed_sheets:
                                rc = processed_sheets[fl_lower]
                                s, e = '<sheetData>', '</sheetData>'
                                si, ei = rc.find(s), rc.find(e)
                                ts, te = tmpl_sheet.find(s), tmpl_sheet.find(e)
                                if si != -1 and ei != -1 and ts != -1 and te != -1:
                                    # 1) 템플릿의 행 높이/속성 추출 (ht, customHeight, hidden)
                                    tmpl_sd = tmpl_sheet[ts+len(s):te]
                                    tmpl_row_attrs = {}
                                    for rm in re.finditer(r'<row\b([^>]*)>', tmpl_sd):
                                        ra = rm.group(1)
                                        rid = re.search(r'\br="(\d+)"', ra)
                                        if rid:
                                            tmpl_row_attrs[rid.group(1)] = ra

                                    # 2) 리포트 sheetData 주입
                                    tmpl_sheet = tmpl_sheet[:ts+len(s)] + rc[si+len(s):ei] + tmpl_sheet[te:]

                                    # 3) 템플릿 행 높이 복원 (ht, customHeight 속성만 덮어씀)
                                    def _restore_height(m, _ra=tmpl_row_attrs):
                                        attrs = m.group(1)
                                        rid = re.search(r'\br="(\d+)"', attrs)
                                        if rid and rid.group(1) in _ra:
                                            ta = _ra[rid.group(1)]
                                            for attr in ('ht', 'customHeight', 'hidden'):
                                                # [FIX] If ht or customHeight already exists in attrs (user set it), skip overwriting from template
                                                if attr in ('ht', 'customHeight') and f'{attr}=' in attrs:
                                                    continue

                                                am = re.search(fr'\b{attr}="([^"]+)"', ta)
                                                if am:
                                                    val = f'{attr}="{am.group(1)}"'
                                                    if f'{attr}=' in attrs:
                                                        attrs = re.sub(fr'\b{attr}="[^"]*"', val, attrs)
                                                    else:
                                                        attrs = attrs.rstrip() + f' {val}'
                                        return f'<row {attrs}>'
                                    # 주입된 sheetData 영역에만 적용
                                    ts2 = tmpl_sheet.find(s)
                                    te2 = tmpl_sheet.find(e)
                                    if ts2 != -1 and te2 != -1:
                                        sd_inner = tmpl_sheet[ts2+len(s):te2]
                                        sd_inner = re.sub(r'<row\b([^>]*)>', _restore_height, sd_inner)
                                        tmpl_sheet = tmpl_sheet[:ts2+len(s)] + sd_inner + tmpl_sheet[te2:]

                                    # Keep dynamic merges from the generated workbook.  Copying only
                                    # sheetData leaves the template's old mergeCells block in place.
                                    report_merges = re.search(
                                        r'<mergeCells\b[^>]*>.*?</mergeCells>', rc, flags=re.DOTALL
                                    )
                                    if report_merges:
                                        merge_pattern = r'<mergeCells\b[^>]*>.*?</mergeCells>'
                                        if re.search(merge_pattern, tmpl_sheet, flags=re.DOTALL):
                                            tmpl_sheet = re.sub(
                                                merge_pattern, report_merges.group(0), tmpl_sheet,
                                                count=1, flags=re.DOTALL
                                            )
                                        else:
                                            tmpl_sheet = tmpl_sheet.replace(
                                                '</sheetData>', '</sheetData>' + report_merges.group(0), 1
                                            )
                                    self.log(f"   ✅ [V4] {os.path.basename(fl_lower)} 데이터 주입 (행높이 보존)")
                            z_out.writestr(tmpl_names[fl_lower], tmpl_sheet.encode('utf-8'))

                        except:
                            z_out.writestr(tmpl_names[fl_lower], raw)
                    else:
                        # 나머지 (drawing, media, rels, etc.) → 템플릿 그대로
                        z_out.writestr(tmpl_names[fl_lower], raw)
                    written.add(fl_lower)

                # 추가 시트 (템플릿에 없는 멀티페이지 시트) → Blueprint 적용
                for fl_rep, rc in processed_sheets.items():
                    if fl_rep in written:
                        continue
                    if blueprint_xml:
                        try:
                            s, e = '<sheetData>', '</sheetData>'
                            si, ei = rc.find(s), rc.find(e)
                            bp = re.sub(r'<rowBreaks[^>]*>.*?</rowBreaks>', '', blueprint_xml, flags=re.DOTALL)
                            ts, te = bp.find(s), bp.find(e)
                            if si != -1 and ei != -1 and ts != -1 and te != -1:
                                new_xml = bp[:ts+len(s)] + rc[si+len(s):ei] + bp[te:]
                                # 추가 을지에도 생성된 리포트의 고정 배율과 여백을 적용한다.
                                for tag_name in ('pageSetup', 'pageMargins', 'printOptions'):
                                    report_tag = re.search(
                                        fr'<{tag_name}\b[^>]*/>', rc
                                    )
                                    if not report_tag:
                                        continue
                                    tag_pattern = fr'<{tag_name}\b[^>]*/>'
                                    if re.search(tag_pattern, new_xml):
                                        new_xml = re.sub(
                                            tag_pattern, report_tag.group(0),
                                            new_xml, count=1
                                        )
                                    else:
                                        new_xml = new_xml.replace(
                                            '</sheetData>',
                                            '</sheetData>' + report_tag.group(0), 1
                                        )
                                report_merges = re.search(
                                    r'<mergeCells\b[^>]*>.*?</mergeCells>', rc, flags=re.DOTALL
                                )
                                if report_merges:
                                    merge_pattern = r'<mergeCells\b[^>]*>.*?</mergeCells>'
                                    if re.search(merge_pattern, new_xml, flags=re.DOTALL):
                                        new_xml = re.sub(
                                            merge_pattern, report_merges.group(0), new_xml,
                                            count=1, flags=re.DOTALL
                                        )
                                    else:
                                        new_xml = new_xml.replace(
                                            '</sheetData>', '</sheetData>' + report_merges.group(0), 1
                                        )
                                report_header = re.search(
                                    r'<headerFooter\b[^>]*>.*?</headerFooter>|'
                                    r'<headerFooter\b[^>]*/>',
                                    rc, flags=re.DOTALL
                                )
                                if report_header:
                                    header_pattern = (
                                        r'<headerFooter\b[^>]*>.*?</headerFooter>|'
                                        r'<headerFooter\b[^>]*/>'
                                    )
                                    if re.search(header_pattern, new_xml, flags=re.DOTALL):
                                        new_xml = re.sub(
                                            header_pattern, report_header.group(0), new_xml,
                                            count=1, flags=re.DOTALL
                                        )
                                    elif re.search(r'<pageSetup\b[^>]*/>', new_xml):
                                        new_xml = re.sub(
                                            r'(<pageSetup\b[^>]*/>)',
                                            r'\1' + report_header.group(0), new_xml, count=1
                                        )
                                    else:
                                        new_xml = new_xml.replace(
                                            '</sheetData>',
                                            '</sheetData>' + report_header.group(0), 1
                                        )
                                new_xml = normalize_print_options(new_xml)
                                z_out.writestr(fl_rep, new_xml.encode('utf-8'))
                                if blueprint_rels:
                                    z_out.writestr(f"xl/worksheets/_rels/{os.path.basename(fl_rep)}.rels", blueprint_rels)
                                self.log(f"   📄 [V4] {os.path.basename(fl_rep)} Blueprint 추가")
                                written.add(fl_rep)
                        except: pass

            os.remove(target_path)
            os.rename(temp_swap, target_path)
            if os.path.exists(final_out): os.remove(final_out)
            self.log("   🛡️ [V4 완료] 인쇄영역 + 그림 + 내용 모두 보존")

        except Exception as e:
            for fc in [temp_swap, final_out]:
                if os.path.exists(fc):
                    try: os.remove(fc)
                    except: pass
            self.log(f"   ⚠️ [오류] 드로잉 수술 실패: {e}")
            traceback.print_exc()

    def calculate_rt_shots(self, size_str):
        """배관 구경을 분석하여 [표 1] 기준 촬영 매수 반환"""
        try:
            nums = re.findall(r"[-+]?\d*\.\d+|\d+", str(size_str))
            if not nums: return 1
            val = float(nums[0])
            if val >= 30: return 7
            if val >= 26: return 6
            if val >= 20: return 5
            if val == 18: return 6
            if val >= 14: return 5
            if val >= 10: return 4
            if val >= 6: return 3
            if val >= 3: return 3
            if val <= 2.5: return 2
            return 1
        except: return 1

