import openpyxl
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
from openpyxl.cell.cell import Cell, MergedCell
import os

# [FIX] Patch openpyxl MergedCell to prevent 'read-only' error when assigning to merged cells
if not hasattr(MergedCell, '_patched_value'):
    MergedCell._patched_value = True
    MergedCell.value = property(lambda self: None, lambda self, val: None)

import json
import datetime
from collections import defaultdict
import re
from time import perf_counter

from site_apps.central.src.services.monthly_report_pagination_mixin import MonthlyReportPaginationMixin
from site_apps.central.src.services.monthly_report_sections_mixin import MonthlyReportSectionsMixin

class MonthlyReportManager(MonthlyReportPaginationMixin, MonthlyReportSectionsMixin):
    # Worksheet row-height points, including the common document header.
    # The cumulative continuation pages need headroom for Excel's print fit.
    NDT_PAGE_HEIGHT = 990.0
    CUMULATIVE_NDT_PAGE_HEIGHT = 940.0
    WELDER_NAMES = {
        'W-2024-A-03': '김진철',
        'W-2026-A-03': '김충환',
        'W-2024-A-33': '김태웅',
        'W-2023-A-27': '박은식',
        'W-2024-A-31': '배준형',
        'W-2026-A-04': '서병오',
        'W-2026-A-05': '선성문',
        'W-2024-A-46': '신상원',
        'W-2026-A-06': '이신희',
        'W-2023-A-25': '이종군',
        'W-2024-A-32': '이진우',
        'W-2024-A-28': '임황제',
        'W-2023-A-28': '하선종',
        'W-2024-A-01': '허동욱',
        'W-2026-A-07': '황규창',
        'W-2026-A-08': '황성철',
        'W-2024-A-47': '황수인',
        'W-2024-A-30': '임범철',
        'W-2024-A-07': '강성엽',
        'W-2024-A-10': '박상길',
        'W-2024-A-44': '정경훈',
        'W-2023-A-34': '강대용',
        'W-2024-A-29': '이정철',
        'W-2024-A-17': '이두월',
        'W-2024-A-15': '이정규',
        'W-2024-A-12': '송명섭',
        'W-2024-A-09': '김영기',
        'W-2024-A-08': '김병규',
        'W-2024-A-05': '이일재',
        'W-2026-A-10': '임범준',
        'W-2024-A-04': '주정용',
        'W-2026-A-02': '이봉준',
        'W-2026-A-09': '김종필',
        'W-2024-A-22': '김영준',
        'W-2024-A-36': '노정래',
    }

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

    def generate_report(self, history_path, target_ym, output_path, doc_num="01",
                        create_date=None, owner_report_path=None,
                        report_scope="month", progress=None):

        started = perf_counter()
        last_stage = started
        timings = []

        def record_stage(label):
            nonlocal last_stage
            now = perf_counter()
            timings.append(f"{label} {now - last_stage:.2f}초")
            last_stage = now

        self._history_path = history_path

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

        if os.path.exists(output_path):
            for source_path in (self.template_path, history_path, owner_report_path):
                if source_path and os.path.exists(source_path) and os.path.samefile(output_path, source_path):
                    raise ValueError("보고서 저장 위치가 입력 파일과 같습니다. 다른 파일 이름을 선택해 주세요.")

        if progress is not None:
            progress("작업일보 집계 중…")
        with open(history_path, 'r', encoding='utf-8') as f:
            history_data = json.load(f)

        # 1. 데이터 필터링 및 집계
        # 월간: 선택한 월의 자료만 포함한다. 전체 누적: 최초 기록부터
        # 선택 월 말일까지 포함한다.
        valid_dates = sorted(
            d for d in history_data.keys()
            if re.match(r'^\d{4}-\d{2}-\d{2}$', str(d))
            and str(d)[:7] <= target_ym
        )
        if report_scope == "month":
            target_dates = [d for d in valid_dates if str(d)[:7] == target_ym]
        elif report_scope == "cumulative":
            target_dates = valid_dates
        else:
            raise ValueError(f"지원하지 않는 집계 범위입니다: {report_scope}")

        if not target_dates:
            scope_text = "in" if report_scope == "month" else "through"
            print(f"No data found {scope_text} {target_ym}")
            return None

        # 집계 구조 (1.2 요약용: 업체+구간+라인+관경+규격별 그룹)
        qty_summary = {}
        ndt_groups = {
            'PAUT': defaultdict(lambda: {'count': 0, 'qty': 0.0}),
            'MT': defaultdict(lambda: {'count': 0, 'qty': 0.0}),
            'RT': defaultdict(lambda: {'count': 0, 'qty': 0.0}),
            'PT': defaultdict(lambda: {'count': 0, 'qty': 0.0})
        }
        # 세부 현황용 (2.x 전체 레코드 - 날짜순)
        ndt_details = {'PAUT': [], 'MT': [], 'RT': [], 'PT': []}
        process_photos = []
        # 4.0/1.1 RT 용접사별 불합격 결함 수
        rt_defects_by_welder = defaultdict(lambda: defaultdict(int))
        rt_retests_by_welder = defaultdict(int)
        rt_original_counts_by_welder = defaultdict(int)
        paut_retest_lengths_by_welder = defaultdict(float)
        paut_original_lengths_by_welder = defaultdict(float)

        def split_spec_and_shift(row):
            """신규 분리 필드와 기존 결합 규격을 동일한 형태로 정규화한다."""
            spec = str(row.get('규격', '') or '').strip()
            shift = str(row.get('근무구분', '') or '').strip()
            if not shift:
                for legacy_shift in ('주간', '야간', '재검'):
                    if spec.endswith(legacy_shift):
                        spec = spec[:-len(legacy_shift)].strip()
                        shift = legacy_shift
                        break
            if spec.startswith('31/3'):
                spec = '3 1/3' + spec[len('31/3'):]
            spec = re.sub(r'\s*[Xx×]\s*', ' X ', spec)
            return spec, shift

        # 물량표(qty) 처리를 위해 시작일의 전일누계와 전체 금일작업 합산을 구함
        from site_apps.central.src.rt_quantities import corrected_rt_history
        history_data = corrected_rt_history(history_data)
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
            # 월간 보고서는 해당 월 사진만 사용한다. 전체누적 보고서는
            # 누적 시작일부터 종료 월까지 등록된 공정사진을 모두 포함한다.
            if report_scope == "cumulative" or str(d)[:7] == target_ym:
                ndt_rows = day_data.get('ndt_results', [])
                for stored_photo in day_data.get('process_photos', []):
                    photo = dict(stored_photo)
                    joint_no = str(photo.get('joint_no', '') or '').strip()
                    location = str(photo.get('location', '') or '').strip()
                    section = str(photo.get('section', '') or '').strip()
                    photo_method = str(photo.get('process', '') or '').strip().upper()
                    photo_line = str(photo.get('line_no', '') or '').strip().casefold()
                    matched_row = next(
                        (
                            ndt_row for ndt_row in ndt_rows
                            if str(ndt_row.get('검사방법', '') or '').strip().upper() == photo_method
                            and str(ndt_row.get('라인번호', '') or '').strip().casefold() == photo_line
                            and str(ndt_row.get('Joint No.', '') or '').strip().casefold() == joint_no.casefold()
                        ),
                        None,
                    )
                    if matched_row is None and joint_no:
                        matched_row = next(
                            (
                                ndt_row for ndt_row in ndt_rows
                                if str(ndt_row.get('검사방법', '') or '').strip().upper() == photo_method
                                and str(ndt_row.get('Joint No.', '') or '').strip().casefold() == joint_no.casefold()
                            ),
                            None,
                        )
                    if matched_row is not None:
                        matched_section = str(matched_row.get('구간', '') or '').strip()
                        matched_line = str(matched_row.get('라인번호', '') or '').strip()
                        if matched_line:
                            photo['line_no'] = matched_line
                        if matched_section:
                            photo['section'] = matched_section
                            photo['location'] = matched_section
                        elif not location or location.casefold() == joint_no.casefold():
                            if matched_line:
                                photo['location'] = matched_line
                    process_photos.append(photo)
            for key, val in day_data.get('qty_data', {}).items():
                if key not in qty_summary:
                    qty_summary[key] = {'예상량': val.get('예상량', ''), '전월누계': '0', '금월작업': 0.0, '총누계': '0', '공정률': '', '불량': 0, '불량률': ''}

                try:
                    today_work = float(str(val.get('금일작업', '0')).replace(',', '') or 0)
                    qty_summary[key]['금월작업'] += today_work
                except ValueError:
                    pass

            # NDT 결과 집계
            for row in day_data.get('ndt_results', []):
                method = str(row.get('검사방법', '')).strip().upper()
                if not method or method not in ndt_groups:
                    continue

                if method == 'RT':
                    welder = str(row.get('용접사', '') or '').strip()
                    if welder:
                        for defect_code in self._rt_reject_defects(row):
                            rt_defects_by_welder[welder][defect_code] += 1
                        joint_no = str(row.get('Joint No.', '') or '').strip().upper()
                        film_count = self._rt_film_count(row)
                        from site_apps.central.src.rt_quantities import retest_parent
                        if retest_parent(joint_no):
                            rt_retests_by_welder[welder] += film_count
                        elif joint_no:
                            rt_original_counts_by_welder[welder] += film_count
                elif method == 'PAUT':
                    welder = str(row.get('용접사', '') or '').strip()
                    joint_no = str(row.get('Joint No.', '') or '').strip().upper()
                    inspection_length = self._safe_float(row.get('PAUT'))
                    if welder and inspection_length > 0:
                        if joint_no.endswith('R1'):
                            paut_retest_lengths_by_welder[welder] += inspection_length
                        elif joint_no:
                            paut_original_lengths_by_welder[welder] += inspection_length

                company = str(row.get('업체', '')).strip()
                if not company: continue # 빈 행 무시

                section = str(row.get('구간', '')).strip()
                line_no = str(row.get('라인번호', '')).strip()
                size = str(row.get('관경', '')).strip()
                spec, work_shift = split_spec_and_shift(row)

                # 수량 (길이 또는 매수)
                qty_val = 0.0
                if method == 'PAUT': qty_val = self._safe_float(row.get('PAUT'))
                elif method == 'MT': qty_val = self._safe_float(row.get('MT'))
                elif method == 'PT': qty_val = self._safe_float(row.get('PT'))
                elif method == 'RT':
                    from site_apps.central.src.rt_quantities import rt_counts
                    qty_val, _ = rt_counts(row)

                group_key = (company, section, line_no, size, spec, work_shift)
                ndt_groups[method][group_key]['count'] += 1
                if method == 'RT':
                    from site_apps.central.src.rt_quantities import retest_parent
                    group = ndt_groups[method][group_key]
                    joint_str = str(row.get('Joint No.', '') or '').strip().upper()
                    if retest_parent(joint_str):
                        d_ori, d_re = 0, qty_val
                        group['re'] = group.get('re', 0) + qty_val
                    else:
                        d_ori, d_re = qty_val, 0
                        ndt_groups[method][group_key]['qty'] += qty_val
                else:
                    ndt_groups[method][group_key]['qty'] += qty_val
                    d_ori = qty_val
                    d_re  = 0.0

                # 2.x 세부현황용: 개별 레코드 저장 (날짜 포함)
                if (d_ori + d_re) > 0:
                    ndt_details[method].append({
                        'key': group_key,
                        'joint_no': str(row.get('Joint No.', '') or '').strip(),
                        'ori': d_ori, 're': d_re, 'qty': qty_val
                    })

        # 마지막 날의 총누계, 공정률, 누적 불량 및 불량률을 최종 반영
        for key, val in last_day_data.get('qty_data', {}).items():
            if key in qty_summary:
                qty_summary[key]['총누계'] = val.get('총누계', '0')
                qty_summary[key]['공정률'] = val.get('공정률', '')
                try:
                    qty_summary[key]['불량'] = float(
                        str(val.get('불량', '0')).replace(',', '') or 0
                    )
                except ValueError:
                    qty_summary[key]['불량'] = 0
                if val.get('불량률'):
                    qty_summary[key]['불량률'] = val.get('불량률', '')

        # 2. 엑셀 쓰기 준비
        if progress is not None:
            progress("엑셀 작성 중…")
        record_stage("이력 읽기 및 집계")
        wb = openpyxl.load_workbook(self.template_path)
        record_stage("양식 열기")

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
        self._ensure_valid_print_area(ws)

        # Use the saved report's width-fit scale while choosing page boundaries.
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.page_setup.scale = None

        # The template already has the page 6/7 boundary at row 229. Insert
        # the missing page 5/6 boundary at row 192 without replacing row 229
        # or any later boundary (including the page 7/8 break at row 278).
        from openpyxl.worksheet.pagebreak import Break
        template_breaks = sorted(ws.row_breaks.brk, key=lambda brk: int(brk.id))
        if not any(int(brk.id) == 192 for brk in template_breaks):
            template_breaks.append(Break(id=192))
        ws.row_breaks.brk = sorted(
            template_breaks, key=lambda brk: int(brk.id)
        )

        # Capture the template's compact body-row height before dynamic inserts
        # move row 466. Header rows keep their original heights.
        detail_body_row_height = (
            ws.row_dimensions[466].height
            or ws.sheet_format.defaultRowHeight
            or 15
        )




        # 행 삽입마다 table_markers를 갱신하며 위쪽 표부터 처리합니다.
        # is_detail=True: 2.x 세부(전체 레코드), False: 1.x 요약(그룹)
        # Markers are updated after each insertion. Process top to bottom so
        # earlier summary growth cannot invalidate detail-page boundaries.
        sections = [
            ('paut_1', 'PAUT', False), ('mt_1', 'MT', False), ('rt_1', 'RT', False), ('pt_1', 'PT', False),
            ('paut_2', 'PAUT', True), ('mt_2', 'MT', True), ('rt_2', 'RT', True), ('pt_2', 'PT', True)
        ]
        section_item_counts = {}
        # MT/PT가 없고 RT만 있는 경우 13페이지의 PT 하단 빈 행을
        # RT 확장 공간으로 우선 재사용한다.
        rt_uses_page13_blanks = (
            bool(ndt_details['RT'])
            and not ndt_details['MT']
            and not ndt_details['PT']
        )

        for section_key, method, is_detail in sections:
            header_row = self.table_markers[section_key]
            start_row = header_row + 2

            # TOTAL 행을 찾음 (header_row 이후 빈 행이 아닌 "TOTAL"이 있는 행)
            total_row = start_row
            for r in range(start_row, start_row + 20):
                val = str(ws.cell(row=r, column=2).value or "").strip()
                if "TOTAL" in val.upper() or "총계" in val:
                    total_row = r
                    break

            # 데이터 삽입: 2.x는 전체 레코드, 1.x는 요약 그룹
            if is_detail:
                ws.cell(row=header_row, column=12).value = 'Joint No.'
                data_items = [
                    (rec['key'], {'count': 1, 'joint_no': rec['joint_no'],
                                  'qty': rec['qty'], 'ori': rec['ori'], 're': rec['re']})
                    for rec in ndt_details[method]
                ]
            else:
                data_items = []
                current_company = None
                sub_count = 0
                sub_qty = 0.0
                sub_re = 0.0

                # 업체명 기준으로 1차 정렬
                sorted_groups = sorted(ndt_groups[method].items(), key=lambda x: str(x[0][0]))

                for k, v in sorted_groups:
                    comp = k[0]
                    if current_company is not None and comp != current_company:
                        data_items.append(
                            ((f"[{current_company} 소계]", "", "", "", "", ""), {'count': sub_count, 'qty': sub_qty, 'ori': sub_qty, 're': sub_re, 'is_subtotal': True})
                        )
                        sub_count = 0
                        sub_qty = 0.0
                        sub_re = 0.0

                    current_company = comp
                    data_items.append(
                        (k, {'count': v['count'], 'qty': v['qty'], 'ori': v['qty'], 're': v.get('re', 0), 'is_subtotal': False})
                    )
                    sub_count += v['count']
                    sub_qty += v['qty']
                    sub_re += v.get('re', 0)

                if current_company is not None:
                    data_items.append(
                        ((f"[{current_company} 소계]", "", "", "", "", ""), {'count': sub_count, 'qty': sub_qty, 'ori': sub_qty, 're': sub_re, 'is_subtotal': True})
                    )
            num_items = len(data_items)
            section_item_counts[section_key] = num_items

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
                # 안전한 행 삽입: 아래 시트(13페이지 등) 병합 셀 구조 보존
                self._insert_rows_safely(ws, start_row, rows_to_insert)
                for image in ws._images:
                    anchor = getattr(image, 'anchor', None)
                    if hasattr(anchor, '_from') and anchor._from.row + 1 >= start_row:
                        anchor._from.row += rows_to_insert
                        if hasattr(anchor, 'to'):
                            anchor.to.row += rows_to_insert
                total_row += rows_to_insert
                # 모든 기존 마커 업데이트 (삽입된 행 아래에 있는 마커들만 += rows_to_insert)
                for k, v in self.table_markers.items():
                    if v >= start_row:
                        self.table_markers[k] += rows_to_insert

                if is_detail:
                    for added_row in range(start_row, start_row + rows_to_insert):
                        ws.row_dimensions[added_row].height = detail_body_row_height
                    self._consume_detail_page_whitespace(
                        ws, total_row, rows_to_insert * detail_body_row_height
                    )
            else:
                rows_to_insert = 0

            current_row = start_row
            seq_num = 1
            for idx, ((comp, sec, line, size, spec, work_shift), vals) in enumerate(data_items):
                count = vals['count']
                qty = vals['qty']
                unit = "매" if method == "RT" else "m"
                is_subtotal = vals.get('is_subtotal', False)
                if is_subtotal:
                    unit = ""

                # Clear reusable template values before writing this detail row.
                for clear_col in range(17, 24):
                    clear_cell = ws.cell(row=current_row, column=clear_col)
                    if type(clear_cell).__name__ != 'MergedCell':
                        clear_cell.value = None

                # 셀 위치: 2:업체, 3:순번, 4:Section, 6:Line No., 10:관경, 12:용접개소, 14:규격, 16:단위, 17:길이
                ws.cell(row=current_row, column=2).value = comp
                if is_subtotal:
                    ws.cell(row=current_row, column=3).value = ""
                else:
                    ws.cell(row=current_row, column=3).value = str(seq_num)
                    seq_num += 1
                ws.cell(row=current_row, column=4).value = sec
                ws.cell(row=current_row, column=6).value = line
                ws.cell(row=current_row, column=10).value = size
                ws.cell(row=current_row, column=12).value = vals.get('joint_no', '') if is_detail else count
                ws.cell(row=current_row, column=14).value = spec
                ws.cell(row=current_row, column=16).value = unit

                ws.cell(row=current_row, column=16).value = unit

                ori_v = vals.get('ori', qty)
                re_v  = vals.get('re', 0.0)
                tot_v = ori_v + re_v

                def _fmt(v):
                    return f"{v:.4f}" if v % 1 != 0 else str(int(v))

                _write_safe(current_row, 17, _fmt(ori_v) if ori_v > 0 else '')
                if re_v > 0:
                    _write_safe(current_row, 18, _fmt(re_v))  # RE' 앵커(col18)
                # PAUT/RT 및 세부현황 모두 TOTAL(ORI'+RE') 표시
                if method in ('PAUT', 'RT') or is_detail:
                    _write_safe(current_row, 20, _fmt(tot_v) if tot_v > 0 else '')

                # 병합 처리 (템플릿 양식에 맞춤)
                safe_merge(current_row, 4, current_row, 5) # Section
                safe_merge(current_row, 6, current_row, 9) # Line No.
                safe_merge(current_row, 10, current_row, 11) # 관경
                safe_merge(current_row, 12, current_row, 13) # 용접개소
                safe_merge(current_row, 14, current_row, 15) # 규격

                # PAUT/MT/RT/PT 모두 동일: RE'(col18~19)만 병합, ORI'(17)·TOTAL(20)은 개별 셀
                safe_merge(current_row, 18, current_row, 19) # RE'
                safe_merge(current_row, 20, current_row, 21) # TOTAL
                safe_merge(current_row, 22, current_row, 23) # 비고
                _write_safe(current_row, 22, work_shift)

                # Line No. 글씨 다 보이도록 높이 조절
                if current_row >= 401:
                    ws.row_dimensions[current_row].height = detail_body_row_height

                # 폰트, 정렬 적용 (병합된 칸 전체에 정렬 속성을 먹여야 엑셀이 줄바꿈을 정상 인식함)
                for c_idx in range(2, 24):
                    c_cell = ws.cell(row=current_row, column=c_idx)
                    if is_subtotal:
                        import copy
                        bold_font = copy.copy(self.font_normal)
                        bold_font.bold = True
                        c_cell.font = bold_font

                        align = copy.copy(self.align_center)
                        if c_idx == 2:
                            align.wrap_text = False
                            align.shrink_to_fit = True
                        c_cell.alignment = align
                    else:
                        c_cell.font = self.font_normal
                        c_cell.alignment = self.align_center
                    c_cell.border = self.border_thin
                    self._keep_single_lower_horizontal_border(c_cell)

                # Keep long Line No. values on one line inside the widened F:I area.
                # 병합된 F:I(6~9) 모든 셀에 wrap_text=False를 적용하고,
                # 한글 포함 또는 긴 라인번호는 폰트 크기를 자동 축소하여 줄바꿈 및 글자 잘림을 방지함.
                line_str = str(line or '')
                line_display_w = sum(2 if ord(c) > 127 else 1 for c in line_str)
                if line_display_w >= 28:
                    import copy
                    line_font = copy.copy(self.font_normal)
                    line_font.size = 7.5
                elif line_display_w >= 24:
                    import copy
                    line_font = copy.copy(self.font_normal)
                    line_font.size = 8.5
                else:
                    line_font = self.font_normal

                line_align = Alignment(
                    horizontal='center', vertical='center',
                    wrap_text=False, shrink_to_fit=True
                )
                for col_idx in range(6, 10):
                    c = ws.cell(row=current_row, column=col_idx)
                    c.alignment = line_align
                    c.font = line_font

                # Keep complete Section values such as Sec.12-3 visible in
                # the merged D:E area instead of clipping the final digit.
                sec_align = Alignment(
                    horizontal='center', vertical='center',
                    wrap_text=False, shrink_to_fit=True
                )
                for col_idx in (4, 5):
                    ws.cell(row=current_row, column=col_idx).alignment = sec_align

                if is_detail:
                    detail_align = Alignment(
                        horizontal='center', vertical='center',
                        wrap_text=False, shrink_to_fit=True
                    )
                    for col_idx in (12, 13):
                        ws.cell(row=current_row, column=col_idx).alignment = detail_align

                # 규격은 좁은 N:O 병합 영역에서도 두 줄로 나뉘지 않게 한다.
                spec_align = Alignment(
                    horizontal='center', vertical='center',
                    wrap_text=False, shrink_to_fit=True
                )
                for col_idx in (14, 15):
                    ws.cell(row=current_row, column=col_idx).alignment = spec_align

                current_row += 1

            # Dynamic row operations can move the TOTAL row; locate it again.
            for candidate_row in range(current_row, current_row + 40):
                candidate_value = str(
                    ws.cell(row=candidate_row, column=2).value or ''
                ).strip().upper()
                if 'TOTAL' in candidate_value:
                    total_row = candidate_row
                    break

            if num_items == 0:
                # Keep the method visible while clearly distinguishing "no
                # activity" from a missing or failed data load.
                safe_merge(start_row, 2, start_row, 23)
                no_data_cell = ws.cell(row=start_row, column=2)
                no_data_cell.value = '해당 월 검사실적 없음'
                no_data_cell.font = self.font_normal
                no_data_cell.alignment = self.align_center
                no_data_cell.border = self.border_thin
                if start_row >= 401:
                    ws.row_dimensions[start_row].height = detail_body_row_height
                current_row = start_row + 1

            # 빈 행(데이터 없는 행)들 스타일 정리 및 PAUT 병합 처리 준비
            for r in range(current_row, total_row):
                if r >= 401:
                    ws.row_dimensions[r].height = detail_body_row_height
                # 빈 행도 기본적으로 다른 칸들도 병합을 유지해야 템플릿 양식이 안 깨짐
                safe_merge(r, 4, r, 5)
                safe_merge(r, 6, r, 9)
                safe_merge(r, 10, r, 11)
                safe_merge(r, 12, r, 13)
                safe_merge(r, 14, r, 15)

                # 빈 행도 RE'(col18~19)만 병합 (ORI'·TOTAL 개별 유지)
                safe_merge(r, 18, r, 19)
                safe_merge(r, 20, r, 21) # TOTAL
                safe_merge(r, 22, r, 23) # 비고
                for c_idx in range(2, 24):
                    cell = ws.cell(row=r, column=c_idx)
                    cell.border = self.border_thin
                    self._keep_single_lower_horizontal_border(cell)

            # TOTAL 행 값 쓰기 및 스타일 보정
            total_ori = sum(vals.get('ori', vals['qty']) for _, vals in data_items if not vals.get('is_subtotal', False))
            total_re = sum(vals.get('re', 0.0) for _, vals in data_items if not vals.get('is_subtotal', False))
            total_qty = total_ori + total_re

            # TOTAL 행의 17, 18, 19, 20 칸 배경색(파란색)을 동일하게 맞춤 (첫 번째 셀 서식 복사)
            fill_style = ws.cell(row=total_row, column=2).fill
            if method == 'PAUT':
                import copy
                for c_idx in range(17, 21):
                    c = ws.cell(row=total_row, column=c_idx)
                    if fill_style:
                        c.fill = copy.copy(fill_style)

            safe_merge(total_row, 22, total_row, 23) # 비고
            for c_idx in range(2, 24):
                self._keep_single_lower_horizontal_border(
                    ws.cell(row=total_row, column=c_idx)
                )
            if total_row >= 401:
                ws.row_dimensions[total_row].height = detail_body_row_height

            if num_items == 0:
                _write_safe(total_row, 17, 0)
                _write_safe(total_row, 20, 0)
                ws.cell(row=total_row, column=17).alignment = self.align_center
                ws.cell(row=total_row, column=20).alignment = self.align_center

            if total_qty > 0:
                from openpyxl.styles import Alignment as _Align
                shrink_align = _Align(horizontal='center', vertical='center', shrink_to_fit=True)
                if method == 'PAUT':
                    safe_merge(total_row, 20, total_row, 21)
                    _write_safe(total_row, 17, f"{total_ori:.4f}" if total_ori % 1 != 0 else str(int(total_ori)))
                    ws.cell(row=total_row, column=17).alignment = shrink_align
                    if total_re > 0:
                        _write_safe(total_row, 18, f"{total_re:.4f}" if total_re % 1 != 0 else str(int(total_re)))
                        ws.cell(row=total_row, column=18).alignment = shrink_align
                    _write_safe(total_row, 20, f"{total_qty:.4f}" if total_qty % 1 != 0 else str(int(total_qty)))
                    ws.cell(row=total_row, column=20).alignment = shrink_align
                else:
                    _write_safe(total_row, 17, f"{total_ori:.4f}" if total_ori % 1 != 0 else str(int(total_ori)))
                    ws.cell(row=total_row, column=17).alignment = shrink_align
                    safe_merge(total_row, 18, total_row, 19)
                    if total_re > 0:
                        _write_safe(
                            total_row,
                            18,
                            f"{total_re:.4f}"
                            if total_re % 1 != 0 else str(int(total_re)),
                        )
                        ws.cell(row=total_row, column=18).alignment = shrink_align
                    safe_merge(total_row, 20, total_row, 21)
                    _write_safe(
                        total_row,
                        20,
                        f"{total_qty:.4f}"
                        if total_qty % 1 != 0 else str(int(total_qty)),
                    )
                    ws.cell(row=total_row, column=20).alignment = shrink_align
                    if is_detail:
                        safe_merge(total_row, 18, total_row, 19)
                        if total_re > 0:
                            _write_safe(
                                total_row,
                                18,
                                f"{total_re:.4f}"
                                if total_re % 1 != 0 else str(int(total_re)),
                            )
                            ws.cell(row=total_row, column=18).alignment = shrink_align

            # When either a 1.2 summary table or a 2.x detail table overflows,
            # continue it on a real new page. Repeat both the seven-row document
            # header and the subsection/table header.  Repeat as many times as
            # necessary; following MT/RT/PT/4.0 content moves down as a block.
            # The template contains fixed manual breaks for placeholder row
            # counts, including breaks on TOTAL rows. Once real records are
            # written, those boundaries can force an early split even when the
            # calculation below allows the following table on the same page.
            self._remove_stale_table_breaks(ws, start_row, total_row)

            remaining_start = start_row
            remaining_count = num_items
            if not is_detail:
                self._compact_ndt_summary_page(ws)
            previous_section_breaks = sorted(
                int(brk.id)
                for brk in ws.row_breaks.brk
                if int(brk.id) < header_row
            )
            document_header_row = (
                previous_section_breaks[-1] + 1
                if previous_section_breaks else 1
            )
            if section_key == 'paut_2':
                self._compact_blank_spacers_before_table(
                    ws, document_header_row, header_row
                )
            if section_key == 'rt_2':
                self._compact_rt_leading_gap(ws, header_row)
            while remaining_count > 0:
                # Include TOTAL in height checks, allowing it to split alone.
                if is_detail and remaining_start != start_row:
                    continuation_row = self._detail_continuation_overflow_row(
                        ws, remaining_start, remaining_count
                    )
                else:
                    continuation_row = self._detail_page_overflow_row(
                        ws, remaining_start, remaining_count
                    )
                if continuation_row is None:
                    break

                rows_on_current_page = continuation_row - remaining_start
                # The first table attempt can legitimately have no room when
                # earlier tables already consumed the current page.  In that
                # case the normal continuation insertion below moves the whole
                # table to a fresh page.  If it still cannot fit one record on
                # that fresh page, the repeated headings really are too tall.
                if rows_on_current_page <= 0 and remaining_start != start_row:
                    raise ValueError(
                        '검사 표의 제목 높이가 인쇄 가능 영역을 초과합니다. '
                        '양식의 제목 행 높이와 인쇄 여백을 확인해 주세요.'
                    )
                if section_key == 'rt_2' and rt_uses_page13_blanks:
                    inserted_rows = self._insert_page13_rt_continuation(
                        ws, continuation_row, document_header_row, header_row
                    )
                else:
                    inserted_rows = self._insert_table_continuation(
                        ws, continuation_row, document_header_row, header_row
                    )
                from openpyxl.worksheet.pagebreak import Break
                if not any(
                    int(brk.id) == continuation_row - 1
                    for brk in ws.row_breaks.brk
                ):
                    ws.row_breaks.append(Break(id=continuation_row - 1))

                for key, marker_row in self.table_markers.items():
                    if marker_row >= continuation_row:
                        self.table_markers[key] += inserted_rows

                remaining_count -= rows_on_current_page
                remaining_start = continuation_row + inserted_rows

        # When the three method-detail tables on this page have no records,
        # distribute all unused vertical space across MT->RT, RT->PT, and
        # PT->page-bottom.  Including the final gap prevents the lower half of
        # the page from remaining empty while the tables crowd the top.
        if all(section_item_counts.get(key, 0) == 0 for key in ('mt_2', 'rt_2', 'pt_2')):
            method_keys = ('mt_2', 'rt_2', 'pt_2')
            total_rows = []
            for index, current_key in enumerate(method_keys):
                current_marker = self.table_markers[current_key]
                search_end = (
                    self.table_markers[method_keys[index + 1]]
                    if index + 1 < len(method_keys)
                    else ws.max_row + 1
                )
                current_total = None
                for row in range(current_marker + 2, search_end):
                    if 'TOTAL' in str(ws.cell(row=row, column=2).value or '').upper():
                        current_total = row
                        break
                total_rows.append(current_total)

            later_breaks = sorted(
                int(brk.id) for brk in ws.row_breaks.brk
                if int(brk.id) >= self.table_markers['pt_2']
            )
            page_end = later_breaks[0] if later_breaks else None

            if all(row is not None for row in total_rows) and page_end is not None:
                gap_limits = (
                    (total_rows[0] + 1, self.table_markers['rt_2']),
                    (total_rows[1] + 1, self.table_markers['pt_2']),
                    (total_rows[2] + 1, page_end + 1),
                )
                default_height = ws.sheet_format.defaultRowHeight or 15
                gap_rows = []
                for gap_start, gap_end in gap_limits:
                    blank_spacer_rows = []
                    for spacer_row in range(gap_start, gap_end):
                        is_blank = all(
                            not str(ws.cell(row=spacer_row, column=col).value or '').strip()
                            for col in range(1, 24)
                        )
                        if is_blank:
                            blank_spacer_rows.append(spacer_row)
                    gap_rows.append(blank_spacer_rows)

                # Equalize the number of spacer rows instead of assigning
                # different row heights to each gap.  This keeps every visible
                # grid cell the same height while preserving equal spacing.
                total_spacers = sum(len(rows) for rows in gap_rows)
                base_count, remainder = divmod(total_spacers, 3)
                desired_counts = [
                    base_count + (1 if index < remainder else 0)
                    for index in range(3)
                ]

                first_delta = desired_counts[0] - len(gap_rows[0])
                rt_marker = self.table_markers['rt_2']
                rt_title_row = rt_marker - 1
                if first_delta > 0:
                    self._insert_rows_safely(ws, rt_title_row, first_delta)
                elif first_delta < 0:
                    self._delete_rows_safely(
                        ws, rt_title_row + first_delta, -first_delta
                    )
                if first_delta:
                    self.table_markers['rt_2'] += first_delta
                    self.table_markers['pt_2'] += first_delta
                    total_rows[1] += first_delta
                    total_rows[2] += first_delta
                    page_end += first_delta

                second_delta = desired_counts[1] - len(gap_rows[1])
                pt_marker = self.table_markers['pt_2']
                pt_title_row = pt_marker - 1
                if second_delta > 0:
                    self._insert_rows_safely(ws, pt_title_row, second_delta)
                elif second_delta < 0:
                    self._delete_rows_safely(
                        ws, pt_title_row + second_delta, -second_delta
                    )
                if second_delta:
                    self.table_markers['pt_2'] += second_delta
                    total_rows[2] += second_delta
                    page_end += second_delta

                third_delta = desired_counts[2] - len(gap_rows[2])
                if third_delta > 0:
                    self._insert_rows_safely(ws, page_end, third_delta)
                    page_end += third_delta
                elif third_delta < 0:
                    self._delete_rows_safely(
                        ws, total_rows[2] + 1, -third_delta
                    )
                    page_end += third_delta

                uniform_gaps = (
                    (total_rows[0] + 1, self.table_markers['rt_2'] - 1),
                    (total_rows[1] + 1, self.table_markers['pt_2'] - 1),
                    (total_rows[2] + 1, page_end + 1),
                )
                for gap_start, gap_end in uniform_gaps:
                    for spacer_row in range(gap_start, gap_end):
                        ws.row_dimensions[spacer_row].height = default_height

        # 마지막으로 맨 위의 1.1 물량표 처리
        qty_header_row = self.table_markers['qty']
        # Monthly report terminology only; column positions and data mappings
        # remain unchanged (I:J = prior-month cumulative, K:L = this month).
        ws.cell(row=qty_header_row, column=9).value = '\uc804\uc6d4 \ub204\uacc4'
        ws.cell(row=qty_header_row, column=11).value = '\uae08\uc6d4 \uc791\uc5c5'
        qty_start_row = self.table_markers['qty'] + 1
        paut_spec_rows = [
            ('PAUT_300A이상', '300A이상', '129'),
            ('PAUT_300A이상-야간', '300A이상-야간', '624'),
            ('PAUT_250A', '250A', '4'),
            ('PAUT_250A-야간', '250A-야간', '1'),
            ('PAUT_200A', '200A', '4'),
            ('PAUT_200A-야간', '200A-야간', '2'),
            ('PAUT_150A~125A', '150A~125A', '1'),
            ('PAUT_150A~125A-야간', '150A~125A-야간', '1'),
            ('PAUT_100A이하', '100A이하', '1'),
            ('PAUT_100A이하-야간', '100A이하-야간', '1'),
            ('PAUT_소계', '소계', '768'),
        ]
        import copy as _copy
        rt_start_offset = next(
            (
                offset for offset in range(1, 15)
                if str(ws.cell(qty_start_row + offset, 2).value or '').strip().upper() == 'RT'
            ),
            6,
        )
        extra_paut_rows = len(paut_spec_rows) - rt_start_offset
        if extra_paut_rows > 0:
            paut_top_b_style = _copy.copy(ws.cell(qty_start_row, 2)._style)
            paut_top_c_style = _copy.copy(ws.cell(qty_start_row, 3)._style)
            spec_styles = [
                _copy.copy(ws.cell(qty_start_row + 2, col)._style)
                for col in range(1, 24)
            ]
            sub_styles = [
                _copy.copy(ws.cell(qty_start_row + rt_start_offset - 1, col)._style)
                for col in range(1, 24)
            ]
            insert_at = qty_start_row + rt_start_offset - 1
            self._insert_rows_safely(ws, insert_at, extra_paut_rows)
            for image in ws._images:
                anchor = getattr(image, 'anchor', None)
                if hasattr(anchor, '_from') and anchor._from.row + 1 >= insert_at:
                    anchor._from.row += extra_paut_rows
                    if hasattr(anchor, 'to'):
                        anchor.to.row += extra_paut_rows
            for k, v in self.table_markers.items():
                if k != 'qty' and v >= insert_at:
                    self.table_markers[k] += extra_paut_rows
            for offset, (spec_key, spec_label, default_exp) in enumerate(paut_spec_rows):
                r_idx = qty_start_row + offset
                row_styles = sub_styles if offset == len(paut_spec_rows) - 1 else spec_styles
                for col_idx in range(2, 24):
                    ws.cell(r_idx, col_idx)._style = _copy.copy(row_styles[col_idx - 1])
                for sc, ec in ((4, 6), (7, 8), (9, 10), (11, 12), (13, 14),
                               (15, 16), (17, 18), (19, 20), (21, 23)):
                    safe_merge(r_idx, sc, r_idx, ec)
                for clear_col in (7, 9, 11, 13, 15, 17, 19, 21):
                    _write_safe(r_idx, clear_col, None)
                _write_safe(r_idx, 4, spec_label)
                if spec_key not in qty_summary:
                    qty_summary[spec_key] = {
                        '예상량': default_exp, '전월누계': '0', '금월작업': 0.0,
                        '총누계': '0', '공정률': '', '불량': 0, '불량률': '',
                    }
            ws.cell(qty_start_row, 2)._style = paut_top_b_style
            ws.cell(qty_start_row, 3)._style = paut_top_c_style
            safe_merge(qty_start_row, 2, qty_start_row + len(paut_spec_rows) - 1, 3)
            _write_safe(qty_start_row, 2, 'PAUT')
            ws.cell(qty_start_row, 2).alignment = self.align_center

        rt_base = qty_start_row + len(paut_spec_rows)
        qty_mapping = {
            **{key: qty_start_row + idx for idx, (key, _, _) in enumerate(paut_spec_rows)},
            'RT_150A~100A': rt_base,
            'RT_150A~100A-야간': rt_base + 1,
            'RT_80A이하': rt_base + 2,
            'RT_80A이하-야간': rt_base + 3,
            'RT_소계': rt_base + 4,
            'MT_전체(주간)': rt_base + 5,
            'MT_전체(야간)': rt_base + 6,
            'PT_전체(주간)': rt_base + 7,
            'PT_전체(야간)': rt_base + 8,
        }
        for r_idx in range(qty_start_row, qty_start_row + len(qty_mapping)):
            ws.row_dimensions[r_idx].height = 12.4

        # [FIX] Dynamically recalculate subtotals (소계) to ensure accuracy even if daily JSON records are missing subtotal values
        paut_keys = ['PAUT_300A이상', 'PAUT_300A이상-야간', 'PAUT_250A', 'PAUT_250A-야간', 'PAUT_200A', 'PAUT_200A-야간', 'PAUT_150A~125A', 'PAUT_150A~125A-야간', 'PAUT_100A이하', 'PAUT_100A이하-야간', 'PAUT_불가']
        rt_keys = ['RT_150A~100A', 'RT_150A~100A-야간', 'RT_80A이하', 'RT_80A이하-야간', 'RT_불가']

        for sub_key, keys_to_sum in [('PAUT_소계', paut_keys), ('RT_소계', rt_keys)]:
            if sub_key not in qty_summary:
                qty_summary[sub_key] = {'예상량': '0', '전월': '0', '전월누계': '0', '금월작업': 0.0, '총누계': '0', '공정률': '', '불량': 0, '불량률': ''}

            qty_summary[sub_key]['금월작업'] = 0.0
            total_expected = 0.0
            total_prev = 0.0
            total_accum = 0.0
            total_defect = 0.0

            for k in keys_to_sum:
                if k in qty_summary:
                    try:
                        qty_summary[sub_key]['금월작업'] += float(qty_summary[k].get('금월작업', 0))
                        total_expected += float(str(qty_summary[k].get('예상량', '0')).replace(',', '') or 0)

                        # 전월은 '전월' 대신 '전일누계'에서 올 수 있음.
                        prev_val = qty_summary[k].get('전월', '0')
                        if prev_val == '0' and '전일누계' in qty_summary[k]:
                            prev_val = qty_summary[k]['전일누계']
                        total_prev += float(str(prev_val).replace(',', '') or 0)

                        total_accum += float(str(qty_summary[k].get('총누계', '0')).replace(',', '') or 0)
                        total_defect += float(str(qty_summary[k].get('불량', '0')).replace(',', '') or 0)
                    except ValueError:
                        pass

            qty_summary[sub_key]['예상량'] = str(total_expected)
            qty_summary[sub_key]['전월'] = str(total_prev)
            qty_summary[sub_key]['전월누계'] = str(total_prev)
            qty_summary[sub_key]['총누계'] = str(total_accum)
            qty_summary[sub_key]['불량'] = str(total_defect)
            # 실적이 전혀 없는 공정은 0.0%로 표시하지 않는다. 0.0은
            # 실제 실적값처럼 보일 수 있으므로 템플릿의 빈칸을 유지한다.
            if total_expected > 0 and total_accum > 0:
                qty_summary[sub_key]['공정률'] = str(round((total_accum / total_expected) * 100, 1))
            else:
                qty_summary[sub_key]['공정률'] = ''
            if total_accum > 0 and total_defect > 0:
                qty_summary[sub_key]['불량률'] = str(round((total_defect / total_accum) * 100, 1))
            else:
                qty_summary[sub_key]['불량률'] = ''


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
        # Rebalance widths so Line No. fits on one line without widening the page.
        for col in ('F', 'G', 'H', 'I'):
            ws.column_dimensions[col].width = 6.5
        for col in ('J', 'K'):
            ws.column_dimensions[col].width = 3.5
        for col in ('L', 'M'):
            ws.column_dimensions[col].width = 3.0
        for col in ('N', 'O'):
            ws.column_dimensions[col].width = 3.5
        ws.column_dimensions['P'].width = 5.0

        ws.column_dimensions['Q'].width = 9.0
        ws.column_dimensions['R'].width = 4.5
        ws.column_dimensions['S'].width = 4.5
        ws.column_dimensions['T'].width = 4.5
        ws.column_dimensions['U'].width = 4.5
        ws.column_dimensions['V'].width = 0.0   # 비고 우측 끝 숨김
        ws.column_dimensions['V'].width = 2.25
        ws.column_dimensions['V'].hidden = False
        ws.column_dimensions['W'].width = 0.0   # 비고 우측 끝 숨김
        ws.column_dimensions['W'].width = 2.25
        ws.column_dimensions['W'].hidden = False

        # 사용자 요청: 모든 표의 좌측(2열), 우측(21열) 외곽선을 굵은 실선으로 복원/강제 설정
        import copy
        medium = Side(style='medium')
        def apply_outer_borders(start_r, end_r):
            # Preserve all internal thin lines and replace only the four
            # perimeter edges with one consistent medium outline.
            for r in range(start_r, end_r + 1):
                lc = ws.cell(row=r, column=2)
                left_border = copy.copy(lc.border)
                left_border.left = copy.copy(medium)
                lc.border = left_border

                rc = ws.cell(row=r, column=23)
                right_border = copy.copy(rc.border)
                right_border.right = copy.copy(medium)
                rc.border = right_border

            for c in range(2, 24):
                tc = ws.cell(row=start_r, column=c)
                top_border = copy.copy(tc.border)
                top_border.top = copy.copy(medium)
                tc.border = top_border

                bc = ws.cell(row=end_r, column=c)
                bottom_border = copy.copy(bc.border)
                bottom_border.bottom = copy.copy(medium)
                bc.border = bottom_border

        # 1. 상단 물량표 외곽선 교정
        if self.table_markers['qty'] > 0:
            qty_last_row = self.table_markers['qty'] + len(qty_mapping)
            apply_outer_borders(
                self.table_markers['qty'],
                qty_last_row,
            )

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
            # Walk populated cells in row order without materializing blank cells.
            text_cells = sorted(
                (position, cell) for position, cell in sheet._cells.items()
                if isinstance(cell.value, str) and cell.value
            )
            for _, cell in text_cells:
                val_str = str(cell.value)

                if isinstance(cell.value, str):
                    # Normalize the current central-branch contract
                    # wording even when the template already contains
                    # a non-placeholder 2026 title.
                    central_branch = '2026\ub144 \uc911\uc559\uc9c0\uc0ac'
                    ndt_contract = '\ube44\ud30c\uad34\uac80\uc0ac\uc6a9\uc5ed \ub2e8\uac00\uacc4\uc57d'
                    heat_pipe = '\uc5f4\ubc30\uad00'
                    heat_transport_pipe = '\uc5f4\uc218\uc1a1\uad00'
                    if (
                        central_branch in val_str
                        and ndt_contract in val_str
                        and heat_pipe in val_str
                    ):
                        cell.value = val_str.replace(
                            heat_pipe, heat_transport_pipe
                        )
                        val_str = str(cell.value)
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
                    # 수식으로 제목을 참조하는 경우 (예: =F41, =F$41, =$F$41 등)
                    match = re.search(r'\$?F\$?(\d+)', val_no_spaces, re.IGNORECASE)
                    if match:
                        ref_row = int(match.group(1))
                        # 참조된 셀의 실제 값이 제목 키워드를 포함하는지 확인
                        ref_val = str(sheet.cell(row=ref_row, column=6).value or '').replace(' ', '').replace('\xa0', '').replace('\n', '')
                        if ('단가계약' in ref_val or '계약명' in ref_val) and ('월간용역' in ref_val or '보고서' in ref_val):
                            is_title = True
                        # 참조 셀에 값이 없는 경우 행 번호 패턴으로 폴백 (41, 81, 121... 등 40행 간격)
                        elif ref_row % 40 == 1:
                            is_title = True

                if is_title:
                    contract_name = replacements.get('[[계약명]]', '2026년 중앙지사 열수송관 비파괴검사용역 단가계약')
                    # 무조건 괄호와 줄바꿈이 포함된 포맷으로 강제 덮어쓰기 (수식도 문자열로 덮어씌움)
                    cell.value = f"【{contract_name}】\n월 간 용 역 진 도 보 고 서"

                    # Alignment and Font are imported at module level.
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
        # Dynamic continuation pages change the section total, so recalculate
        # the 3.0 page labels after all row and page operations are complete.
        self._write_welder_defect_summary(
            ws, rt_defects_by_welder, rt_retests_by_welder,
            rt_original_counts_by_welder, '1.1', ('방사선투과검사', 'RT'),
        )
        self._write_welder_defect_summary(
            ws, {}, paut_retest_lengths_by_welder,
            paut_original_lengths_by_welder, '1.2',
            ('위상배열초음파탐상검사', 'PAUT'),
        )
        self._populate_safety_training_log(ws, target_ym, ndt_groups, target_dates, report_scope)
        self._populate_owner_report_rows(ws, target_ym, owner_report_path)
        self._fill_repeated_header_document_numbers(ws, doc_num)
        self._fix_ndt_section_labels(ws)
        self._ensure_cover_cell_elements(ws, create_date)
        self._move_mt_criteria_to_page6(ws)
        self._merge_pages7_and8(ws)
        # Each continuation already has a height-based boundary. Adding blank
        # rows to match the preceding page's row count can create extra pages.
        self._normalize_ndt_header_logos(ws)
        self._normalize_all_report_header_logos(ws)
        self._compact_ndt_method_spacers(ws)
        self._ensure_pt_continuation_header(ws)
        self._compact_weld_defect_trailing_spacers(ws)
        self._collapse_blank_rows_before_report_headers(ws)
        self._normalize_repeated_report_header_text(ws, doc_num)
        self._fit_long_section_headers(ws)
        self._fit_inspector_qualification_rows(ws)
        self._rebalance_toc_page_layout(ws)
        self._align_report_breaks_to_repeated_headers(ws)
        self._renumber_ndt_status_pages(ws)
        self._keep_iuc_euc_headers_single_line(ws)
        self._clear_broken_reference_placeholders(ws)
        self._widen_weld_defect_table(ws)

        # Preserve the visible right edge of merged template boxes and headers.
        for sheet in wb.worksheets:
            self._repair_merged_right_borders(sheet)
        self._fix_pt_summary_bottom_border(ws)

        # Prevent blank right-side pages caused by the widened V:W remarks area.
        # Keep vertical pagination automatic while fitting the report to one
        # printed page horizontally.
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.page_setup.scale = None

        for visible_row in (570, 571):
            ws.row_dimensions[visible_row].hidden = False
            if not ws.row_dimensions[visible_row].height:
                ws.row_dimensions[visible_row].height = 1.0

        # Rows 570-571 belong to the visible report frame and must remain
        # present even when blank; hiding them suppresses the outer border.
        # Only the original final trailing spacer (row 572) may be collapsed.
        for blank_row in range(572, 573):
            is_blank = all(
                not str(ws.cell(row=blank_row, column=col).value or '').strip()
                for col in range(1, min(ws.max_column, 23) + 1)
            )
            if is_blank:
                ws.row_dimensions[blank_row].height = 0
                ws.row_dimensions[blank_row].hidden = True

        self._trim_trailing_blank_print_pages(ws)
        self._complete_ndt_quantity_totals(ws)
        self._annotate_ndt_overruns(ws)
        self._layout_ndt_pages(ws, capacity_points=self.CUMULATIVE_NDT_PAGE_HEIGHT)
        # Build photo pages only after every dynamic 3.0 row operation. Images
        # are drawing anchors rather than cells; creating them earlier makes
        # later row deletion/insertion detach them from their page headers.
        record_stage("표 작성 및 페이지 구성")
        self._populate_process_photo_pages(ws, process_photos, doc_num=doc_num)
        record_stage("공정사진 페이지 작성")
        self._renumber_ndt_status_pages(ws)
        self._normalize_ndt_header_logos(ws)
        self._normalize_all_report_header_logos(
            ws, reference_toc=True
        )
        self._normalize_report_header_top_borders(ws)
        self._fill_repeated_header_document_numbers(ws, doc_num)

        self._fit_photo_page_trailing_rows(ws)
        self._trim_print_area_after_last_process_photo(ws)
        self._space_photo_page_caption(ws)
        self._insert_other_report_top_spacers(ws)
        self._place_general_equipment_table_image(ws)
        self._bridge_blank_detail_row_borders(ws)
        self._collapse_blank_detail_row_gaps(ws)
        self._fix_report_section_header_labels(ws)
        # Normalize after every row-height and top-spacer adjustment.
        self._normalize_all_report_header_logos(ws, reference_toc=True)
        self._center_enlarge_report_logos(ws)
        self._ensure_valid_print_area(ws)
        self._fit_blank_photo_page_overflow(ws)
        if progress is not None:
            progress("파일 저장 중…")
        record_stage("최종 페이지 정리")
        result = self._save_workbook_with_fallback(wb, output_path)
        record_stage("파일 저장")
        print(f"[월간진도보고서] 총 {perf_counter() - started:.2f}초 | " + " | ".join(timings))
        return result
