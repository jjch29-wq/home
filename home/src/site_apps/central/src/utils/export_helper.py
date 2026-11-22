import os
import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

def to_float(val):
    if val is None:
        return 0.0
    try:
        clean = str(val).replace(',', '').replace('₩', '').replace('%', '').strip()
        return float(clean or 0)
    except:
        return 0.0


def export_budget_estimation_to_excel(filepath, summary_data, labor_data, material_data, expense_data,
                                       actual_labor_data=None, actual_material_data=None, actual_expense_data=None,
                                       base_salaries=None):
    """
    Export the detailed budget estimation to a formatted Excel file.
    Prefers the official company standard 15-sheet template if available.
    Falls back to the full-width portrait 2-sheet format if template is absent.
    """
    template_path = os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 
        'templates', 
        '공사실행예산_정산서_신규양식.xlsx'
    )
    if os.path.exists(template_path):
        try:
            return export_budget_to_standard_template(
                filepath, summary_data, labor_data, material_data, expense_data,
                actual_labor_data=actual_labor_data,
                actual_material_data=actual_material_data,
                actual_expense_data=actual_expense_data,
                template_path=template_path
            )
        except Exception as e:
            import traceback
            traceback.print_exc()
            return _export_budget_legacy_excel(
                filepath, summary_data, labor_data, material_data, expense_data,
                actual_labor_data=actual_labor_data,
                actual_material_data=actual_material_data,
                actual_expense_data=actual_expense_data,
                base_salaries=base_salaries
            )
    else:
        return _export_budget_legacy_excel(
            filepath, summary_data, labor_data, material_data, expense_data,
            actual_labor_data=actual_labor_data,
            actual_material_data=actual_material_data,
            actual_expense_data=actual_expense_data,
            base_salaries=base_salaries
        )


def export_budget_to_standard_template(filepath, summary_data, labor_data, material_data, expense_data,
                                       actual_labor_data=None, actual_material_data=None, actual_expense_data=None,
                                       template_path=None):
    """
    Fill the company standard 15-sheet execution budget & settlement template.
    Preserves all formulas, styles, links, and calculation chains.
    """
    import openpyxl
    wb = openpyxl.load_workbook(template_path)

    def _to_f(val):
        try:
            return float(str(val).replace(',', '').replace('₩', '').replace('%', '').strip() or 0)
        except:
            return 0.0

    def _parse_dates(period_str):
        if not period_str or '~' not in str(period_str):
            return None, None
        try:
            parts = str(period_str).split('~')
            s_date = datetime.datetime.strptime(parts[0].strip(), '%Y.%m.%d').date()
            e_date = datetime.datetime.strptime(parts[1].strip(), '%Y.%m.%d').date()
            return s_date, e_date
        except Exception:
            return None, None

    # 1. Sheet: 공사정보
    if '공사정보' in wb.sheetnames:
        ws_info = wb['공사정보']
        site_raw = summary_data.get('site', '중앙지사')
        site_title = "한국지역난방공사 중앙지사 열수송관 비파괴검사" if "중앙" in site_raw else f"{site_raw} 비파괴검사"
        ws_info['C4'] = site_title
        ws_info['C5'] = "한국지역난방공사"
        ws_info['C6'] = "중앙지사 : 부장 주진철"

        p_start, p_end = _parse_dates(summary_data.get('period', ''))
        if p_start and p_end:
            ws_info['C8'] = p_start
            ws_info['C9'] = p_end

        a_start, a_end = _parse_dates(summary_data.get('actual_period', ''))
        if a_start and a_end:
            ws_info['D8'] = a_start
            ws_info['D9'] = a_end
        elif p_start and p_end:
            ws_info['D8'] = p_start
            ws_info['D9'] = p_end

        rev = _to_f(summary_data.get('revenue', 0))
        if rev > 0:
            ws_info['C14'] = rev
            ws_info['C15'] = 1.0  # 지분율 100% (단독수행)
            ws_info['C17'] = f"도급금액 ₩{rev:,.0f} (한국지역난방공사 중앙지사)"

        a_rev = _to_f(summary_data.get('actual_revenue', 0))
        if a_rev > 0:
            ws_info['D14'] = a_rev
            ws_info['D15'] = 1.0
            ws_info['D17'] = f"정산(기성) 누계 ₩{a_rev:,.0f}"
        elif rev > 0:
            ws_info['D14'] = rev
            ws_info['D15'] = 1.0
            ws_info['D17'] = "정산 진행중"

    # 실제 2026년 중앙지사 사전실행예산 확정 투입 인원 및 13개월 월별 투입 계획
    planned_workers = [
        ("주진철", "부장", "경인", "자사", "PAUT/MT/PT"),
        ("권인경", "차장", "경인", "자사", "PAUT/MT/PT"),
        ("우명광", "대리", "경인", "자사", "PAUT/MT/PT"),
        ("박상태", "계장", "경인", "자사", "MT/PT")
    ]
    planned_monthly_days = [
        [3, 5, 15, 15, 20, 10, 5, 20, 10, 15, 16, 16, 10], # 주진철 부장: 160일
        [3, 5, 15, 12, 10, 10, 5, 10, 10, 10, 10, 10, 10], # 권인경 차장: 120일
        [0, 1, 15, 15, 16, 20, 10, 20, 20, 13, 10, 10, 10], # 우명광 대리: 160일
        [0, 1, 10, 10, 15, 20, 15, 9, 8, 8, 8, 8, 8]        # 박상태 계장: 120일
    ]

    # 중앙지사 전체 투입 가능 인력 명단 (사후 실적 정산용)
    actual_workers = [
        ("주진철", "부장", "경인", "자사", "PAUT/MT/PT"),
        ("권인경", "차장", "경인", "자사", "PAUT/MT/PT"),
        ("우명광", "대리", "경인", "자사", "PAUT/MT/PT"),
        ("박상태", "계장", "경인", "자사", "MT/PT"),
        ("장승대", "계장", "경인", "자사", "PAUT/MT/PT"),
        ("김진환", "주임", "경인", "자사", "PAUT/MT/PT"),
        ("김성렬", "주임", "경인", "자사", "PAUT/MT/PT"),
        ("박광복", "부장", "경인", "자사", "PAUT/MT/PT"),
        ("주영광", "과장", "경인", "자사", "PAUT/MT/PT")
    ]

    rank_to_row = {
        '이사': 42,
        '부장': 43,
        '차장': 44,
        '과장': 45,
        '대리': 46,
        '계장': 47,
        '주임': 48,
        '기사': 49,
        '여사원': 49
    }

    # 2. Sheet: 인원투입계획 (사전 - 13개월 월별 및 8월 일자별 투입 계획)
    if '인원투입계획' in wb.sheetnames:
        ws_p_labor = wb['인원투입계획']
        ws_p_labor['M5'] = '8월 (+)'
        ws_p_labor.sheet_properties.outlinePr.summaryRight = False
        ws_p_labor.sheet_view.showOutlineSymbols = True
        ws_p_labor.column_dimensions['F'].width = 18.0
        if 'M' in ws_p_labor.column_dimensions:
            ws_p_labor.column_dimensions['M'].min = 13
            ws_p_labor.column_dimensions['M'].max = 13
            ws_p_labor.column_dimensions['M'].width = 8.0
            ws_p_labor.column_dimensions['M'].collapsed = True
        aug_daily_inputs = [
            {0: 1, 1: 1, 2: 1}, # 주진철: 8-5(수), 8-6(목), 8-7(금) = 3일
            {0: 1, 1: 1, 2: 1}, # 권인경: 8-5(수), 8-6(목), 8-7(금) = 3일
            {},                 # 우명광: 0일
            {}                  # 박상태: 0일
        ]
        for idx, w_info in enumerate(planned_workers):
            r_w = 7 + idx
            ws_p_labor[f'B{r_w}'] = w_info[0]
            ws_p_labor[f'C{r_w}'] = w_info[1]
            ws_p_labor[f'D{r_w}'] = w_info[2]
            ws_p_labor[f'E{r_w}'] = w_info[3]
            ws_p_labor[f'F{r_w}'] = w_info[4]
            ws_p_labor[f'G{r_w}'] = f'=IF(B{r_w}="","",M{r_w}+SUM(Q{r_w}:AB{r_w}))'
            ws_p_labor[f'J{r_w}'] = f'=IF(B{r_w}="","",G{r_w})'
            ws_p_labor[f'K{r_w}'] = 0

            # M열 (8월 요약): =IF(B{r_w}="","",SUM(N{r_w}:P{r_w}))
            ws_p_labor[f'M{r_w}'] = f'=IF(B{r_w}="","",SUM(N{r_w}:P{r_w}))'

            # N~P열 (8월 실제 투입 3개 일자: 8-5, 8-6, 8-7)
            w_aug = aug_daily_inputs[idx] if idx < len(aug_daily_inputs) else {}
            for d_idx in range(3):
                c_col = 14 + d_idx
                ws_p_labor.cell(r_w, c_col).value = w_aug.get(d_idx, None)

            # Q~AB열 (9월 ~ 내년 8월, 12개 월: Col 17 ~ 28)
            if idx < len(planned_monthly_days):
                for m_idx, m_days in enumerate(planned_monthly_days[idx][1:]):
                    ws_p_labor.cell(r_w, 17 + m_idx).value = m_days

            for c_col in range(29, 65):
                ws_p_labor.cell(r_w, c_col).value = None

        for r_clear in range(7 + len(planned_workers), 37):
            ws_p_labor[f'B{r_clear}'] = ""
            ws_p_labor[f'C{r_clear}'] = ""
            ws_p_labor[f'D{r_clear}'] = ""
            ws_p_labor[f'E{r_clear}'] = ""
            ws_p_labor[f'F{r_clear}'] = ""
            ws_p_labor[f'G{r_clear}'] = f'=IF(B{r_clear}="","",M{r_clear}+SUM(Q{r_clear}:AB{r_clear}))'
            ws_p_labor[f'J{r_clear}'] = ""
            ws_p_labor[f'K{r_clear}'] = ""
            ws_p_labor[f'M{r_clear}'] = f'=IF(B{r_clear}="","",SUM(N{r_clear}:P{r_clear}))'
            for c_col in range(14, 65):
                ws_p_labor.cell(r_clear, c_col).value = None

        # 기본 사전 확정 투입 공수 (부장 160일, 차장 120일, 대리 160일, 계장 120일 = 총 560공수)
        default_planned_labor = {
            '부장': {'personnel': 1, 'period': 160},
            '차장': {'personnel': 1, 'period': 120},
            '대리': {'personnel': 1, 'period': 160},
            '계장': {'personnel': 1, 'period': 120}
        }
        active_p_labor = dict(default_planned_labor)
        if labor_data:
            for rank, vals in labor_data.items():
                if rank in rank_to_row and _to_f(vals.get('period')) > 0:
                    active_p_labor[rank] = vals

        for rank, r_idx in rank_to_row.items():
            if rank in active_p_labor:
                vals = active_p_labor[rank]
                p = _to_f(vals.get('personnel'))
                d = _to_f(vals.get('period'))
                ws_p_labor[f'E{r_idx}'] = p if p > 0 else ""
                ws_p_labor[f'F{r_idx}'] = d if d > 0 else ""
            else:
                ws_p_labor[f'E{r_idx}'] = ""
                ws_p_labor[f'F{r_idx}'] = ""

    # 3. Sheet: 경비예상지출항목 (사전)
    if '경비예상지출항목' in wb.sheetnames:
        ws_p_exp = wb['경비예상지출항목']
        # 기본 사전 확정 특별근무 시간 (연장 240h, 야간 240h, 휴일 248h = 총 4,020,000원)
        ot_ext = _to_f(labor_data.get('연장근무', {}).get('period', 0)) if labor_data else 0
        ot_ngt = _to_f(labor_data.get('야간근무', {}).get('period', 0)) if labor_data else 0
        ot_hol = _to_f(labor_data.get('휴일근무', {}).get('period', 0)) if labor_data else 0

        ws_p_exp['D6'] = ot_ext if ot_ext > 0 else 240
        ws_p_exp['D7'] = ot_ngt if ot_ngt > 0 else 240
        ws_p_exp['D8'] = ot_hol if ot_hol > 0 else 248

        for r_mat in range(14, 20):
            ws_p_exp[f'D{r_mat}'] = 0
        if material_data:
            for idx, m_item in enumerate(material_data[:6]):
                r_mat = 14 + idx
                ws_p_exp[f'B{r_mat}'] = m_item.get('name', '')
                ws_p_exp[f'C{r_mat}'] = m_item.get('spec', '')
                ws_p_exp[f'D{r_mat}'] = _to_f(m_item.get('qty', 0))
                ws_p_exp[f'F{r_mat}'] = _to_f(m_item.get('price', 0))
                ws_p_exp[f'I{r_mat}'] = '○'

        for r_clr in range(25, 44):
            ws_p_exp[f'D{r_clr}'] = 0
            ws_p_exp[f'E{r_clr}'] = ""
            ws_p_exp[f'G{r_clr}'] = 1.0

        if expense_data:
            for exp_item in expense_data.get('site_expense', []):
                cat = str(exp_item.get('cat', '')).strip()
                cont = str(exp_item.get('cont', '')).strip()
                ppl = _to_f(exp_item.get('ppl', 0))
                qty = _to_f(exp_item.get('qty', 0))
                price = _to_f(exp_item.get('price', 0))

                if '차량' in cat or '차량' in cont:
                    ws_p_exp['D31'] = ppl if ppl > 0 else 1
                    ws_p_exp['E31'] = qty if qty > 0 else 1
                    ws_p_exp['F31'] = price
                    ws_p_exp['G31'] = 1.0
                elif '소모품' in cat or '소모품' in cont:
                    ws_p_exp['D32'] = ppl if ppl > 0 else 1
                    ws_p_exp['E32'] = qty if qty > 0 else 1
                    ws_p_exp['F32'] = price
                    ws_p_exp['G32'] = 1.0
                elif '복리' in cat or '복리' in cont:
                    ws_p_exp['D36'] = ppl if ppl > 0 else 1
                    ws_p_exp['E36'] = qty if qty > 0 else 1
                    ws_p_exp['F36'] = price
                    ws_p_exp['G36'] = 1.0
                elif '식비' in cat or '식비' in cont:
                    ws_p_exp['D27'] = ppl if ppl > 0 else 1
                    ws_p_exp['E27'] = qty if qty > 0 else 1
                    ws_p_exp['F27'] = price
                    ws_p_exp['G27'] = 1.0
                elif '숙박' in cat or '숙박' in cont:
                    ws_p_exp['D25'] = ppl if ppl > 0 else 1
                    ws_p_exp['E25'] = qty if qty > 0 else 1
                    ws_p_exp['F25'] = price
                    ws_p_exp['G25'] = 1.0

            for r_clr in range(49, 53):
                ws_p_exp[f'D{r_clr}'] = 0
                ws_p_exp[f'E{r_clr}'] = ""
                ws_p_exp[f'G{r_clr}'] = 1.0
            for idx, r_item in enumerate(expense_data.get('rental', [])[:4]):
                r_row = 49 + idx
                cat = r_item.get('cat', '')
                spec = r_item.get('spec', '')
                qty = _to_f(r_item.get('qty', 0))
                period = _to_f(r_item.get('period', 0))
                price = _to_f(r_item.get('price', 0))
                if cat or qty > 0:
                    ws_p_exp[f'B{r_row}'] = cat
                    ws_p_exp[f'C{r_row}'] = spec
                    ws_p_exp[f'D{r_row}'] = qty
                    ws_p_exp[f'E{r_row}'] = period if period > 0 else 1
                    ws_p_exp[f'F{r_row}'] = price
                    ws_p_exp[f'G{r_row}'] = 1.0

            for r_clr in range(58, 64):
                ws_p_exp[f'D{r_clr}'] = 0
                ws_p_exp[f'E{r_clr}'] = ""
                ws_p_exp[f'G{r_clr}'] = 1.0
            for idx, o_item in enumerate(expense_data.get('outsource', [])[:6]):
                o_row = 58 + idx
                cat = o_item.get('cat', '')
                work = o_item.get('work', '')
                count = _to_f(o_item.get('count', 0))
                price = _to_f(o_item.get('price', 0))
                if cat or count > 0:
                    ws_p_exp[f'B{o_row}'] = cat
                    ws_p_exp[f'C{o_row}'] = work
                    ws_p_exp[f'D{o_row}'] = count
                    ws_p_exp[f'E{o_row}'] = 1
                    ws_p_exp[f'F{o_row}'] = price
                    ws_p_exp[f'G{o_row}'] = 1.0

    # 4. Sheet: 장비감가상각 (사전)
    if '장비감가상각' in wb.sheetnames:
        ws_p_dep = wb['장비감가상각']
        dep_list = expense_data.get('depreciation', []) if expense_data else []
        default_p_dep = [
            {'item': "PAUT 장비", 'spec': "", 'qty': 1, 'days': 120, 'rate': 44444},
            {'item': "PAUT SCANNER (MANUAL)", 'spec': "", 'qty': 1, 'days': 120, 'rate': 5556},
            {'item': "PAUT SCANNER (COBRA)", 'spec': "", 'qty': 1, 'days': 120, 'rate': 16667},
            {'item': "YOKE", 'spec': "", 'qty': 1, 'days': 10, 'rate': 222},
            {'item': "현장용 탑차(5년간 보험비 포함)", 'spec': "현장별 차량기입시 탑차 구분 기입", 'qty': 1, 'days': 30, 'rate': 16667},
            {'item': "스타렉스(5년간 보험비 포함)", 'spec': "현장별 차량기입시 스타렉스 구분 기입", 'qty': 1, 'days': 160, 'rate': 16667}
        ]
        active_p_dep = dep_list if dep_list else default_p_dep
        start_d = datetime.date(2026, 8, 5)
        end_d = datetime.date(2027, 8, 5)

        for idx in range(8):
            d_row = 5 + idx
            ws_p_dep[f'B{d_row}'] = idx + 1
            if idx < len(active_p_dep):
                d_item = active_p_dep[idx]
                item_name = d_item.get('item', '')
                spec = d_item.get('spec', '')
                qty = _to_f(d_item.get('qty', 1))
                days = _to_f(d_item.get('days', 0))

                ws_p_dep[f'C{d_row}'] = item_name
                ws_p_dep[f'D{d_row}'] = qty if qty > 0 else 1
                ws_p_dep[f'E{d_row}'] = start_d
                ws_p_dep[f'F{d_row}'] = end_d
                ws_p_dep[f'G{d_row}'] = days if days > 0 else ""
                ws_p_dep[f'H{d_row}'] = f'=IFERROR(VLOOKUP(C{d_row},장비마스터,5,FALSE),0)'
                ws_p_dep[f'I{d_row}'] = f'=ROUND(N(D{d_row})*G{d_row}*H{d_row}, 0)'
                ws_p_dep[f'J{d_row}'] = spec
            else:
                ws_p_dep[f'C{d_row}'] = ""
                ws_p_dep[f'D{d_row}'] = ""
                ws_p_dep[f'E{d_row}'] = ""
                ws_p_dep[f'F{d_row}'] = ""
                ws_p_dep[f'G{d_row}'] = ""
                ws_p_dep[f'H{d_row}'] = f'=IF(C{d_row}="","",IFERROR(VLOOKUP(C{d_row},장비마스터,5,FALSE),0))'
                ws_p_dep[f'I{d_row}'] = f'=IF(OR(C{d_row}="",G{d_row}=""),"",ROUND(N(D{d_row})*G{d_row}*H{d_row}, 0))'
                ws_p_dep[f'J{d_row}'] = ""

    # 5. Sheet: 인원투입계획_사후
    if '인원투입계획_사후' in wb.sheetnames:
        ws_a_labor = wb['인원투입계획_사후']
        ws_a_labor.column_dimensions['F'].width = 18.0
        for idx, w_info in enumerate(actual_workers):
            r_w = 7 + idx
            ws_a_labor[f'B{r_w}'] = w_info[0]
            ws_a_labor[f'C{r_w}'] = w_info[1]
            ws_a_labor[f'D{r_w}'] = w_info[2]
            ws_a_labor[f'E{r_w}'] = w_info[3]
            ws_a_labor[f'F{r_w}'] = w_info[4]
            for c_col in range(13, 58):
                ws_a_labor.cell(r_w, c_col).value = None

        for r_clear in range(7 + len(actual_workers), 37):
            ws_a_labor[f'B{r_clear}'] = ""
            ws_a_labor[f'C{r_clear}'] = ""
            ws_a_labor[f'D{r_clear}'] = ""
            ws_a_labor[f'E{r_clear}'] = ""
            ws_a_labor[f'F{r_clear}'] = ""
            for c_col in range(13, 58):
                ws_a_labor.cell(r_clear, c_col).value = None

        target_a_labor = actual_labor_data if actual_labor_data else labor_data
        if target_a_labor:
            for rank, r_idx in rank_to_row.items():
                if rank in target_a_labor:
                    vals = target_a_labor[rank]
                    p = _to_f(vals.get('personnel'))
                    d = _to_f(vals.get('period'))
                    ws_a_labor[f'E{r_idx}'] = p if p > 0 else ""
                    ws_a_labor[f'F{r_idx}'] = d if d > 0 else ""

    # 6. Sheet: 경비지출항목_사후
    if '경비지출항목_사후' in wb.sheetnames:
        ws_a_exp = wb['경비지출항목_사후']
        target_a_labor = actual_labor_data if actual_labor_data else labor_data
        if target_a_labor:
            ws_a_exp['D6'] = _to_f(target_a_labor.get('연장근무', {}).get('period', 0))
            ws_a_exp['D7'] = _to_f(target_a_labor.get('야간근무', {}).get('period', 0))
            ws_a_exp['D8'] = _to_f(target_a_labor.get('휴일근무', {}).get('period', 0))

        target_a_mat = actual_material_data if actual_material_data else material_data
        for r_mat in range(14, 20):
            ws_a_exp[f'D{r_mat}'] = 0
        if target_a_mat:
            for idx, m_item in enumerate(target_a_mat[:6]):
                r_mat = 14 + idx
                ws_a_exp[f'B{r_mat}'] = m_item.get('name', '')
                ws_a_exp[f'C{r_mat}'] = m_item.get('spec', '')
                ws_a_exp[f'D{r_mat}'] = _to_f(m_item.get('qty', 0))
                ws_a_exp[f'F{r_mat}'] = _to_f(m_item.get('price', 0))
                ws_a_exp[f'I{r_mat}'] = '○'

        for r_clr in range(25, 43):
            ws_a_exp[f'D{r_clr}'] = 0
            ws_a_exp[f'E{r_clr}'] = ""
            ws_a_exp[f'G{r_clr}'] = 1.0

        target_a_exp = actual_expense_data if actual_expense_data else expense_data
        if target_a_exp:
            for exp_item in target_a_exp.get('site_expense', []):
                cat = str(exp_item.get('cat', '')).strip()
                cont = str(exp_item.get('cont', '')).strip()
                ppl = _to_f(exp_item.get('ppl', 0))
                qty = _to_f(exp_item.get('qty', 0))
                price = _to_f(exp_item.get('price', 0))

                if '차량' in cat or '차량' in cont:
                    ws_a_exp['D31'] = ppl if ppl > 0 else 1
                    ws_a_exp['E31'] = qty if qty > 0 else 1
                    ws_a_exp['F31'] = price
                    ws_a_exp['G31'] = 1.0
                elif '소모품' in cat or '소모품' in cont:
                    ws_a_exp['D32'] = ppl if ppl > 0 else 1
                    ws_a_exp['E32'] = qty if qty > 0 else 1
                    ws_a_exp['F32'] = price
                    ws_a_exp['G32'] = 1.0
                elif '복리' in cat or '복리' in cont:
                    ws_a_exp['D36'] = ppl if ppl > 0 else 1
                    ws_a_exp['E36'] = qty if qty > 0 else 1
                    ws_a_exp['F36'] = price
                    ws_a_exp['G36'] = 1.0
                elif '식비' in cat or '식비' in cont:
                    ws_a_exp['D27'] = ppl if ppl > 0 else 1
                    ws_a_exp['E27'] = qty if qty > 0 else 1
                    ws_a_exp['F27'] = price
                    ws_a_exp['G27'] = 1.0
                elif '숙박' in cat or '숙박' in cont:
                    ws_a_exp['D25'] = ppl if ppl > 0 else 1
                    ws_a_exp['E25'] = qty if qty > 0 else 1
                    ws_a_exp['F25'] = price
                    ws_a_exp['G25'] = 1.0

            for r_clr in range(48, 52):
                ws_a_exp[f'D{r_clr}'] = 0
                ws_a_exp[f'E{r_clr}'] = ""
                ws_a_exp[f'G{r_clr}'] = 1.0
            for idx, r_item in enumerate(target_a_exp.get('rental', [])[:4]):
                r_row = 48 + idx
                cat = r_item.get('cat', '')
                spec = r_item.get('spec', '')
                qty = _to_f(r_item.get('qty', 0))
                period = _to_f(r_item.get('period', 0))
                price = _to_f(r_item.get('price', 0))
                if cat or qty > 0:
                    ws_a_exp[f'B{r_row}'] = cat
                    ws_a_exp[f'C{r_row}'] = spec
                    ws_a_exp[f'D{r_row}'] = qty
                    ws_a_exp[f'E{r_row}'] = period if period > 0 else 1
                    ws_a_exp[f'F{r_row}'] = price
                    ws_a_exp[f'G{r_row}'] = 1.0

            for r_clr in range(57, 63):
                ws_a_exp[f'D{r_clr}'] = 0
                ws_a_exp[f'E{r_clr}'] = ""
                ws_a_exp[f'G{r_clr}'] = 1.0
            for idx, o_item in enumerate(target_a_exp.get('outsource', [])[:6]):
                o_row = 57 + idx
                cat = o_item.get('cat', '')
                work = o_item.get('work', '')
                count = _to_f(o_item.get('count', 0))
                price = _to_f(o_item.get('price', 0))
                if cat or count > 0:
                    ws_a_exp[f'B{o_row}'] = cat
                    ws_a_exp[f'C{o_row}'] = work
                    ws_a_exp[f'D{o_row}'] = count
                    ws_a_exp[f'E{o_row}'] = 1
                    ws_a_exp[f'F{o_row}'] = price
                    ws_a_exp[f'G{o_row}'] = 1.0

    # 7. Sheet: 장비감가상각_사후
    if '장비감가상각_사후' in wb.sheetnames:
        ws_a_dep = wb['장비감가상각_사후']
        target_a_exp = actual_expense_data if actual_expense_data else expense_data
        dep_a_list = target_a_exp.get('depreciation', []) if target_a_exp else []
        for idx in range(8):
            d_row = 5 + idx
            ws_a_dep[f'B{d_row}'] = idx + 1
            if idx < len(dep_a_list):
                d_item = dep_a_list[idx]
                item_name = d_item.get('item', '')
                spec = d_item.get('spec', '')
                qty = _to_f(d_item.get('qty', 1))
                days = _to_f(d_item.get('days', 0))

                ws_a_dep[f'C{d_row}'] = item_name
                ws_a_dep[f'D{d_row}'] = qty if qty > 0 else 1
                if days > 0:
                    ws_a_dep[f'G{d_row}'] = days
                ws_a_dep[f'H{d_row}'] = f'=IFERROR(VLOOKUP(C{d_row},장비감가상각!장비마스터,5,FALSE),0)'
                ws_a_dep[f'I{d_row}'] = f'=ROUND(N(D{d_row})*G{d_row}*H{d_row}, 0)'
                ws_a_dep[f'J{d_row}'] = spec
            else:
                ws_a_dep[f'C{d_row}'] = ""
                ws_a_dep[f'D{d_row}'] = ""
                ws_a_dep[f'E{d_row}'] = ""
                ws_a_dep[f'F{d_row}'] = ""
                ws_a_dep[f'G{d_row}'] = f'=IF(AND(ISNUMBER(E{d_row}),ISNUMBER(F{d_row}),F{d_row}>=E{d_row}),F{d_row}-E{d_row}+1,0)'
                ws_a_dep[f'H{d_row}'] = f'=IF(C{d_row}="","",IFERROR(VLOOKUP(C{d_row},장비감가상각!장비마스터,5,FALSE),0))'
                ws_a_dep[f'I{d_row}'] = f'=IF(OR(C{d_row}="",G{d_row}=""),"",ROUND(N(D{d_row})*G{d_row}*H{d_row}, 0))'
                ws_a_dep[f'J{d_row}'] = ""

    # 8. Sheet: 실적배분
    if '실적배분' in wb.sheetnames:
        ws_dist = wb['실적배분']
        ws_dist['F2'] = site_title
        ws_dist['F4'] = a_rev if a_rev > 0 else rev

    wb.save(filepath)
    return True


def _export_budget_legacy_excel(
    filepath,
    summary_data,
    labor_data,
    material_data,
    expense_data,
    actual_labor_data=None,
    actual_material_data=None,
    actual_expense_data=None,
    base_salaries=None
):
    """
    공사실행예산서 엑셀 내보내기 (상단 총괄 비교 + 하단 상세 일체형 2개 시트):
    - 시트 1: [사전예산서 (Planned)] - 상단 결재란 + 사전예산 요약 집계표 + 하단 상세 내역
    - 시트 2: [사후정산서 (Actuals)] - 상단 결재란 + 계획 vs 실적 vs 차액 총괄 비교표 + 하단 실적 상세 내역
    """
    wb = Workbook()
    
    # ----------------------------------------------------
    # 스타일 상수 정의
    # ----------------------------------------------------
    title_font = Font(name='맑은 고딕', size=16, bold=True)
    info_title_font = Font(name='맑은 고딕', size=11, bold=True)
    section_font = Font(name='맑은 고딕', size=11, bold=True)
    header_font = Font(name='맑은 고딕', size=10, bold=True)
    bold_font = Font(name='맑은 고딕', size=10, bold=True)
    normal_font = Font(name='맑은 고딕', size=10)
    ref_title_font = Font(name='맑은 고딕', size=10, bold=True, color='B71C1C')

    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left_align = Alignment(horizontal='left', vertical='center')
    right_align = Alignment(horizontal='right', vertical='center')

    thin_border_line = Side(border_style="thin", color="000000")
    border_all = Border(top=thin_border_line, left=thin_border_line, right=thin_border_line, bottom=thin_border_line)

    fill_header = PatternFill(start_color='D9D9D9', end_color='D9D9D9', fill_type='solid')      # 헤더 기본 회색
    fill_sub_header = PatternFill(start_color='E8ECEF', end_color='E8ECEF', fill_type='solid')  # 서브헤더
    fill_subtotal = PatternFill(start_color='FFF9C4', end_color='FFF9C4', fill_type='solid')    # 소계 연노랑
    fill_grand_total = PatternFill(start_color='FFE082', end_color='FFE082', fill_type='solid') # 총계 강조노랑
    fill_summary_hdr = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid') # 총괄 연하늘
    fill_ref_hdr = PatternFill(start_color='FCE4EC', end_color='FCE4EC', fill_type='solid')     # 기준급여 연분홍

    site_name = summary_data.get('site', '')
    period_str = summary_data.get('period', '')
    act_period_str = summary_data.get('actual_period', '') or period_str

    default_base_salaries = {
        "이사": 55250000, "부장": 55250000, "차장": 47670000, "과장": 41170000,
        "대리": 37920000, "계장": 34670000, "주임": 31420000, "기사": 29250000
    }
    if base_salaries and isinstance(base_salaries, dict):
        current_base_salaries = {**default_base_salaries, **base_salaries}
    else:
        current_base_salaries = default_base_salaries

    summary_items_def = [
        ("계약금액(Revenue) (원)", 'revenue', 'actual_revenue', 'diff_revenue'),
        ("매출금액(UnitPrice) (원)", 'unitprice', 'actual_unitprice', 'diff_unitprice'),
        ("실행 노무비(Labor) (원)", 'labor', 'actual_labor', 'diff_labor'),
        ("실행 재료비(Material) (원)", 'material', 'actual_material', 'diff_material'),
        ("실행 경비(Expense) (원)", 'expense', 'actual_expense', 'diff_expense'),
        ("실행 외주비(Outsource) (원)", 'outsource', 'actual_outsource', 'diff_outsource'),
        ("영업이익(Profit) (원)", 'profit', 'actual_profit', 'diff_profit'),
        ("이익률(Margin) (%)", 'margin', 'actual_margin', 'diff_margin'),
    ]

    def apply_box_border(ws, start_row, start_col, end_row, end_col):
        for r in range(start_row, end_row + 1):
            for c in range(start_col, end_col + 1):
                ws.cell(row=r, column=c).border = border_all

    # ====================================================
    # 공통 상단부 (타이틀, 결재란, 현장기본정보, 1. 수입(매출)) 작성
    # ====================================================
    def write_sheet_top(ws, title_text, is_actual=False):
        ws.views.sheetView[0].showGridLines = True
        
        # 인쇄 설정: 세로보기(Portrait), A4 용지, 너비 1페이지 맞춤, 가로 가운데 정렬
        ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.print_options.horizontalCentered = True
        ws.page_margins.left = 0.3
        ws.page_margins.right = 0.3
        ws.page_margins.top = 0.5
        ws.page_margins.bottom = 0.5

        # 1행: 빈 행
        ws.row_dimensions[1].height = 15

        # 2행: 타이틀 (A2:G2 병합)
        ws.merge_cells('A2:G2')
        c_title = ws['A2']
        c_title.value = title_text
        c_title.font = title_font
        c_title.alignment = left_align
        ws.row_dimensions[2].height = 30

        # 결재란 (H2:L7)
        # 2행: 직급 헤더
        approval_headers = ["소장", "담당", "상무", "본부장", "사장"]
        start_appr_col = 8  # H열 = 8
        for idx, h in enumerate(approval_headers):
            col_idx = start_appr_col + idx
            h_cell = ws.cell(row=2, column=col_idx, value=h)
            h_cell.font = bold_font
            h_cell.alignment = center_align
            h_cell.border = border_all

        # 3~7행: 서명란 세로 병합 (H3:H7, I3:I7, J3:J7, K3:K7, L3:L7)
        for idx in range(len(approval_headers)):
            col_idx = start_appr_col + idx
            ws.merge_cells(start_row=3, start_column=col_idx, end_row=7, end_column=col_idx)
            apply_box_border(ws, 3, col_idx, 7, col_idx)

        for r_h in range(3, 8):
            ws.row_dimensions[r_h].height = 16

        # 현장 기본정보 (Row 4 ~ 7)
        full_site_name = site_name
        if not full_site_name:
            full_site_name = "2026년 중앙지사 열수송관 비파괴검사 단가계약"
        elif "열수송관" not in full_site_name and "비파괴" not in full_site_name:
            full_site_name = f"2026년 {full_site_name} 열수송관 비파괴검사 단가계약"

        # 4행: 현 장 명 (A4:G4 병합)
        ws.merge_cells('A4:G4')
        c_site = ws.cell(row=4, column=1, value=f"현 장 명 : {full_site_name}")
        c_site.font = info_title_font
        c_site.alignment = left_align
        ws.row_dimensions[4].height = 20

        # 5행: 공사기간 및 일수 (A5:D5 병합, E5 일수)
        cur_period = act_period_str if is_actual else period_str
        days_str = ""
        if cur_period and '~' in cur_period:
            try:
                s_str, e_str = cur_period.split('~')
                s_d = datetime.datetime.strptime(s_str.strip().replace(' ', ''), "%Y.%m.%d")
                e_d = datetime.datetime.strptime(e_str.strip().replace(' ', ''), "%Y.%m.%d")
                days_str = f"{(e_d - s_d).days + 1} 일"
            except: pass
        if not cur_period:
            cur_period = "2026. 08. 04 ~ 2027. 08. 05"
            days_str = "365 일"

        p_label = "공사기간(실적) :" if is_actual else "공사기간 :"
        ws.merge_cells('A5:D5')
        c_prd = ws.cell(row=5, column=1, value=f"{p_label} {cur_period}")
        c_prd.font = info_title_font
        c_prd.alignment = left_align

        if days_str:
            c_days = ws.cell(row=5, column=5, value=days_str)
            c_days.font = bold_font
            c_days.alignment = right_align
        ws.row_dimensions[5].height = 20

        # 6행: 계약업체 (A6:G6 병합)
        ws.merge_cells('A6:G6')
        c_com = ws.cell(row=6, column=1, value="계약업체 : 한국지역난방공사 중앙지사")
        c_com.font = info_title_font
        c_com.alignment = left_align
        ws.row_dimensions[6].height = 20

        # 7행: 경인사무소 (A7:G7 병합)
        ws.merge_cells('A7:G7')
        c_off = ws.cell(row=7, column=1, value="경인사무소 : 주진철")
        c_off.font = info_title_font
        c_off.alignment = left_align
        ws.row_dimensions[7].height = 20

        # 8행: 공백
        ws.row_dimensions[8].height = 12

        # 9행: 1. 수입(매출) 섹션 타이틀 (A9:D9 병합)
        ws.merge_cells('A9:D9')
        c_rev_lbl = ws.cell(row=9, column=1, value="1. 수입(매출)")
        c_rev_lbl.font = section_font
        c_rev_lbl.alignment = left_align
        ws.row_dimensions[9].height = 22

        # 10행: 계약 금액 / 산출근거 헤더 (B:D 병합으로 넉넉한 금액 공간 확보, E:L 산출근거)
        ws.merge_cells('B10:D10')
        c_b10 = ws.cell(row=10, column=2, value="계약 금액")
        c_b10.font = bold_font
        c_b10.fill = fill_header
        c_b10.alignment = center_align
        apply_box_border(ws, 10, 2, 10, 4)

        ws.merge_cells('E10:L10')
        c_c10 = ws.cell(row=10, column=5, value="산출근거")
        c_c10.font = bold_font
        c_c10.fill = fill_header
        c_c10.alignment = center_align
        apply_box_border(ws, 10, 5, 10, 12)
        ws.row_dimensions[10].height = 22

        # 11행: 계약금액 수치 및 산출근거 내용
        rev_val = summary_data.get('actual_revenue' if is_actual and summary_data.get('actual_revenue') else 'revenue', '')
        rev_num = to_float(rev_val)
        if rev_num == 0 and not rev_val:
            rev_num = 288268000.0  # 기본 계약금액

        ws.merge_cells('B11:D11')
        c_b11 = ws.cell(row=11, column=2, value=rev_num)
        c_b11.font = bold_font
        c_b11.alignment = center_align
        c_b11.number_format = '"₩"#,##0'
        apply_box_border(ws, 11, 2, 11, 4)

        ws.merge_cells('E11:L11')
        c_c11 = ws.cell(row=11, column=5, value="계약금액/부가세별도")
        c_c11.font = normal_font
        c_c11.alignment = center_align
        apply_box_border(ws, 11, 5, 11, 12)
        ws.row_dimensions[11].height = 22

        # 12행: 공백
        ws.row_dimensions[12].height = 15

    # ====================================================
    # 상세 내역 (인건비, 재료비, 경비) 작성 헬퍼 (A~L열 풀-위드 100% 채움)
    # ====================================================
    def write_details_section(ws, start_row, cost_header_name, labor_d, mat_d, exp_d):
        cur_row = start_row
        ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=12)
        c_sec = ws.cell(row=cur_row, column=1, value="3. 상세 산출 내역")
        c_sec.font = section_font
        c_sec.alignment = left_align
        ws.row_dimensions[cur_row].height = 24

        # ----------------------------------------------------
        # 1) 정시근무 + 우측 기준급여 참조표 (A~G열 근무표 + H~L열 고정금액)
        # ----------------------------------------------------
        cur_row += 1
        ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=7)
        c_sub1 = ws.cell(row=cur_row, column=1, value="1) 정시근무 (240일/년)")
        c_sub1.font = bold_font
        c_sub1.alignment = left_align

        # 우측 기준급여 타이틀 (H~L열)
        ws.merge_cells(start_row=cur_row, start_column=8, end_row=cur_row, end_column=12)
        c_ref_title = ws.cell(row=cur_row, column=8, value="고정금액(변경불가)")
        c_ref_title.font = ref_title_font
        c_ref_title.fill = fill_ref_hdr
        c_ref_title.alignment = center_align
        apply_box_border(ws, cur_row, 8, cur_row, 12)
        ws.row_dimensions[cur_row].height = 22

        cur_row += 1
        # 좌측 근무표 헤더 (A~G열)
        headers_t1 = [("구분", 1, 1), ("직급", 2, 2), ("투입인원(명)", 3, 3), ("투입일수/인(일)", 4, 4), ("단가/일", 5, 5), (cost_header_name, 6, 7)]
        for h, c_s, c_e in headers_t1:
            if c_s != c_e:
                ws.merge_cells(start_row=cur_row, start_column=c_s, end_row=cur_row, end_column=c_e)
            cell = ws.cell(row=cur_row, column=c_s, value=h)
            cell.font = header_font
            cell.fill = fill_header
            cell.alignment = center_align
            apply_box_border(ws, cur_row, c_s, cur_row, c_e)

        # 우측 기준급여 헤더 (H~I열: 직급, J~L열: 기준급여)
        ws.merge_cells(start_row=cur_row, start_column=8, end_row=cur_row, end_column=9)
        c_rh1 = ws.cell(row=cur_row, column=8, value="직급")
        c_rh1.font = header_font
        c_rh1.fill = fill_header
        c_rh1.alignment = center_align
        apply_box_border(ws, cur_row, 8, cur_row, 9)

        ws.merge_cells(start_row=cur_row, start_column=10, end_row=cur_row, end_column=12)
        c_rh2 = ws.cell(row=cur_row, column=10, value="기준급여")
        c_rh2.font = header_font
        c_rh2.fill = fill_header
        c_rh2.alignment = center_align
        apply_box_border(ws, cur_row, 10, cur_row, 12)
        ws.row_dimensions[cur_row].height = 30

        ranks = ["이사", "부장", "차장", "과장", "대리", "계장", "주임", "기사"]
        ref_ranks = ["부장", "차장", "과장", "대리", "계장", "주임", "기사"]
        
        start_t1_row = cur_row + 1
        t1_personnel_sum = 0.0
        t1_days_sum = 0.0
        t1_cost_sum = 0.0

        for idx, rank in enumerate(ranks):
            r = start_t1_row + idx
            ws.row_dimensions[r].height = 20

            c_rk = ws.cell(row=r, column=2, value=rank)
            c_rk.font = normal_font
            c_rk.border = border_all
            c_rk.alignment = center_align

            r_data = (labor_d or {}).get(rank, {})
            p = to_float(r_data.get('personnel'))
            d = to_float(r_data.get('period'))
            u = to_float(r_data.get('unit_price'))
            if u == 0:
                base_sal = current_base_salaries.get(rank, 0)
                u = round(base_sal / 240) if base_sal else 0

            amt = round(p * d * u) if (p > 0 and d > 0) else 0

            t1_personnel_sum += p
            t1_days_sum += d
            t1_cost_sum += amt

            c_p = ws.cell(row=r, column=3, value=p if p > 0 else "")
            c_p.font = normal_font
            c_p.border = border_all
            c_p.alignment = center_align

            c_d = ws.cell(row=r, column=4, value=d if d > 0 else "")
            c_d.font = normal_font
            c_d.border = border_all
            c_d.alignment = center_align

            c_u = ws.cell(row=r, column=5, value=u)
            c_u.font = normal_font
            c_u.border = border_all
            c_u.alignment = right_align
            c_u.number_format = '#,##0'

            ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=7)
            c_amt = ws.cell(row=r, column=6, value=amt)
            c_amt.font = normal_font
            c_amt.alignment = right_align
            c_amt.number_format = '#,##0'
            apply_box_border(ws, r, 6, r, 7)

            if idx < len(ref_ranks):
                ref_rk = ref_ranks[idx]
                ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=9)
                c_rrk = ws.cell(row=r, column=8, value=ref_rk)
                c_rrk.font = normal_font
                c_rrk.alignment = center_align
                apply_box_border(ws, r, 8, r, 9)

                sal_val = current_base_salaries.get(ref_rk, 0)
                ws.merge_cells(start_row=r, start_column=10, end_row=r, end_column=12)
                c_rsal = ws.cell(row=r, column=10, value=sal_val)
                c_rsal.font = normal_font
                c_rsal.alignment = right_align
                c_rsal.number_format = '#,##0'
                apply_box_border(ws, r, 10, r, 12)
            else:
                apply_box_border(ws, r, 8, r, 12)

        end_t1_row = start_t1_row + len(ranks) - 1
        ws.merge_cells(start_row=start_t1_row, start_column=1, end_row=end_t1_row, end_column=1)
        c_m1 = ws.cell(row=start_t1_row, column=1, value="정시근무\n(240일/년)")
        c_m1.font = bold_font
        c_m1.fill = fill_sub_header
        c_m1.alignment = center_align
        for r_fill in range(start_t1_row, end_t1_row + 1):
            ws.cell(row=r_fill, column=1).border = border_all

        sub_t1_row = end_t1_row + 1
        ws.row_dimensions[sub_t1_row].height = 22
        ws.merge_cells(start_row=sub_t1_row, start_column=1, end_row=sub_t1_row, end_column=2)
        c_s1_lbl = ws.cell(row=sub_t1_row, column=1, value="소계")
        c_s1_lbl.font = bold_font
        c_s1_lbl.fill = fill_subtotal
        c_s1_lbl.alignment = center_align
        apply_box_border(ws, sub_t1_row, 1, sub_t1_row, 2)

        c_s1_p = ws.cell(row=sub_t1_row, column=3, value=f"{t1_personnel_sum:g} 명")
        c_s1_p.font = bold_font
        c_s1_p.fill = fill_subtotal
        c_s1_p.alignment = center_align
        c_s1_p.border = border_all

        c_s1_d = ws.cell(row=sub_t1_row, column=4, value=f"{t1_days_sum:g} 일")
        c_s1_d.font = bold_font
        c_s1_d.fill = fill_subtotal
        c_s1_d.alignment = center_align
        c_s1_d.border = border_all

        c_s1_e = ws.cell(row=sub_t1_row, column=5, value="")
        c_s1_e.fill = fill_subtotal
        c_s1_e.border = border_all

        ws.merge_cells(start_row=sub_t1_row, start_column=6, end_row=sub_t1_row, end_column=7)
        c_s1_amt = ws.cell(row=sub_t1_row, column=6, value=t1_cost_sum)
        c_s1_amt.font = bold_font
        c_s1_amt.fill = fill_subtotal
        c_s1_amt.alignment = right_align
        c_s1_amt.number_format = '#,##0'
        apply_box_border(ws, sub_t1_row, 6, sub_t1_row, 7)

        # ----------------------------------------------------
        # 2) 특별근무 (A~L열 12개 열에 고르게 6분배)
        # ----------------------------------------------------
        cur_row = sub_t1_row + 2
        ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=12)
        c_sub2 = ws.cell(row=cur_row, column=1, value="2) 특별근무")
        c_sub2.font = bold_font
        c_sub2.alignment = left_align
        ws.row_dimensions[cur_row].height = 22

        cur_row += 1
        headers_t2 = [
            ("구분", 1, 2),
            ("형태", 3, 4),
            ("투입인원(명)", 5, 6),
            ("투입시간/인(시간)", 7, 8),
            ("단가", 9, 10),
            (cost_header_name, 11, 12)
        ]
        for h, c_s, c_e in headers_t2:
            ws.merge_cells(start_row=cur_row, start_column=c_s, end_row=cur_row, end_column=c_e)
            cell = ws.cell(row=cur_row, column=c_s, value=h)
            cell.font = header_font
            cell.fill = fill_header
            cell.alignment = center_align
            apply_box_border(ws, cur_row, c_s, cur_row, c_e)
        ws.row_dimensions[cur_row].height = 30

        special_types = ["연장근무", "야간근무", "휴일근무"]
        default_ot_rates = {"연장근무": 4000, "야간근무": 5000, "휴일근무": 7500}
        start_t2_row = cur_row + 1
        t2_personnel_sum = 0.0
        t2_hours_sum = 0.0
        t2_cost_sum = 0.0

        for idx, stype in enumerate(special_types):
            r = start_t2_row + idx
            ws.row_dimensions[r].height = 20

            ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
            c_st = ws.cell(row=r, column=3, value=stype)
            c_st.font = normal_font
            c_st.alignment = center_align
            apply_box_border(ws, r, 3, r, 4)

            s_data = (labor_d or {}).get(stype, {})
            p = to_float(s_data.get('personnel'))
            h = to_float(s_data.get('period'))
            u = to_float(s_data.get('unit_price', default_ot_rates.get(stype, 0)))
            if u == 0:
                u = default_ot_rates.get(stype, 0)

            amt = round(p * h * u) if (p > 0 and h > 0) else 0

            t2_personnel_sum += p
            t2_hours_sum += h
            t2_cost_sum += amt

            ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=6)
            c_p = ws.cell(row=r, column=5, value=p if p > 0 else "")
            c_p.font = normal_font
            c_p.alignment = center_align
            apply_box_border(ws, r, 5, r, 6)

            ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=8)
            c_h = ws.cell(row=r, column=7, value=h if h > 0 else "")
            c_h.font = normal_font
            c_h.alignment = center_align
            apply_box_border(ws, r, 7, r, 8)

            ws.merge_cells(start_row=r, start_column=9, end_row=r, end_column=10)
            c_u = ws.cell(row=r, column=9, value=u)
            c_u.font = normal_font
            c_u.alignment = right_align
            c_u.number_format = '#,##0'
            apply_box_border(ws, r, 9, r, 10)

            ws.merge_cells(start_row=r, start_column=11, end_row=r, end_column=12)
            c_amt = ws.cell(row=r, column=11, value=amt)
            c_amt.font = normal_font
            c_amt.alignment = right_align
            c_amt.number_format = '#,##0'
            apply_box_border(ws, r, 11, r, 12)

        end_t2_row = start_t2_row + len(special_types) - 1
        ws.merge_cells(start_row=start_t2_row, start_column=1, end_row=end_t2_row, end_column=2)
        c_m2 = ws.cell(row=start_t2_row, column=1, value="특별근무")
        c_m2.font = bold_font
        c_m2.fill = fill_sub_header
        c_m2.alignment = center_align
        apply_box_border(ws, start_t2_row, 1, end_t2_row, 2)

        sub_t2_row = end_t2_row + 1
        ws.row_dimensions[sub_t2_row].height = 22
        ws.merge_cells(start_row=sub_t2_row, start_column=1, end_row=sub_t2_row, end_column=4)
        c_s2_lbl = ws.cell(row=sub_t2_row, column=1, value="소계")
        c_s2_lbl.font = bold_font
        c_s2_lbl.fill = fill_subtotal
        c_s2_lbl.alignment = center_align
        apply_box_border(ws, sub_t2_row, 1, sub_t2_row, 4)

        ws.merge_cells(start_row=sub_t2_row, start_column=5, end_row=sub_t2_row, end_column=6)
        c_s2_p = ws.cell(row=sub_t2_row, column=5, value=f"{t2_personnel_sum:g} 명")
        c_s2_p.font = bold_font
        c_s2_p.fill = fill_subtotal
        c_s2_p.alignment = center_align
        apply_box_border(ws, sub_t2_row, 5, sub_t2_row, 6)

        ws.merge_cells(start_row=sub_t2_row, start_column=7, end_row=sub_t2_row, end_column=8)
        c_s2_h = ws.cell(row=sub_t2_row, column=7, value=f"{t2_hours_sum:g} 시간")
        c_s2_h.font = bold_font
        c_s2_h.fill = fill_subtotal
        c_s2_h.alignment = center_align
        apply_box_border(ws, sub_t2_row, 7, sub_t2_row, 8)

        ws.merge_cells(start_row=sub_t2_row, start_column=9, end_row=sub_t2_row, end_column=10)
        c_s2_e = ws.cell(row=sub_t2_row, column=9, value="")
        c_s2_e.fill = fill_subtotal
        apply_box_border(ws, sub_t2_row, 9, sub_t2_row, 10)

        ws.merge_cells(start_row=sub_t2_row, start_column=11, end_row=sub_t2_row, end_column=12)
        c_s2_amt = ws.cell(row=sub_t2_row, column=11, value=t2_cost_sum)
        c_s2_amt.font = bold_font
        c_s2_amt.fill = fill_subtotal
        c_s2_amt.alignment = right_align
        c_s2_amt.number_format = '#,##0'
        apply_box_border(ws, sub_t2_row, 11, sub_t2_row, 12)

        grand_labor_row = sub_t2_row + 1
        ws.row_dimensions[grand_labor_row].height = 24
        ws.merge_cells(start_row=grand_labor_row, start_column=1, end_row=grand_labor_row, end_column=10)
        c_gl_lbl = ws.cell(row=grand_labor_row, column=1, value="인건비 총 합계 (정시근무 + 특별근무)")
        c_gl_lbl.font = bold_font
        c_gl_lbl.fill = fill_grand_total
        c_gl_lbl.alignment = center_align
        apply_box_border(ws, grand_labor_row, 1, grand_labor_row, 10)

        ws.merge_cells(start_row=grand_labor_row, start_column=11, end_row=grand_labor_row, end_column=12)
        c_gl_val = ws.cell(row=grand_labor_row, column=11, value=t1_cost_sum + t2_cost_sum)
        c_gl_val.font = bold_font
        c_gl_val.fill = fill_grand_total
        c_gl_val.alignment = right_align
        c_gl_val.number_format = '#,##0'
        apply_box_border(ws, grand_labor_row, 11, grand_labor_row, 12)

        # ----------------------------------------------------
        # 3) 재료비 상세 (A~L열 풀-위드)
        # ----------------------------------------------------
        cur_row = grand_labor_row + 2
        ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=12)
        c_sub3 = ws.cell(row=cur_row, column=1, value="2) 재료비 상세 (Material Cost Detail)")
        c_sub3.font = bold_font
        c_sub3.alignment = left_align
        ws.row_dimensions[cur_row].height = 22

        cur_row += 1
        headers_mat = [
            ("품목", 1, 3),
            ("사양", 4, 5),
            ("수량", 6, 7),
            ("규격", 8, 8),
            ("단가", 9, 10),
            (cost_header_name, 11, 12)
        ]
        for h, c_s, c_e in headers_mat:
            if c_s != c_e:
                ws.merge_cells(start_row=cur_row, start_column=c_s, end_row=cur_row, end_column=c_e)
            cell = ws.cell(row=cur_row, column=c_s, value=h)
            cell.font = header_font
            cell.fill = fill_header
            cell.alignment = center_align
            apply_box_border(ws, cur_row, c_s, cur_row, c_e)
        ws.row_dimensions[cur_row].height = 26

        mat_total = 0.0
        if mat_d:
            for item in mat_d:
                qty = to_float(item.get('qty'))
                price = to_float(item.get('price'))
                amt = round(qty * price) if qty > 0 else 0
                mat_total += amt

                cur_row += 1
                ws.row_dimensions[cur_row].height = 20
                ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=3)
                c_m_name = ws.cell(row=cur_row, column=1, value=item.get('name', ''))
                c_m_name.font = normal_font
                c_m_name.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                apply_box_border(ws, cur_row, 1, cur_row, 3)

                ws.merge_cells(start_row=cur_row, start_column=4, end_row=cur_row, end_column=5)
                c_spec = ws.cell(row=cur_row, column=4, value=item.get('spec', ''))
                c_spec.font = normal_font
                c_spec.alignment = center_align
                apply_box_border(ws, cur_row, 4, cur_row, 5)

                ws.merge_cells(start_row=cur_row, start_column=6, end_row=cur_row, end_column=7)
                c_q = ws.cell(row=cur_row, column=6, value=qty if qty > 0 else "")
                c_q.font = normal_font
                c_q.alignment = center_align
                if qty > 0: c_q.number_format = '#,##0.##'
                apply_box_border(ws, cur_row, 6, cur_row, 7)

                c_unit = ws.cell(row=cur_row, column=8, value=item.get('unit', ''))
                c_unit.font = normal_font
                c_unit.alignment = center_align
                c_unit.border = border_all

                ws.merge_cells(start_row=cur_row, start_column=9, end_row=cur_row, end_column=10)
                c_pr = ws.cell(row=cur_row, column=9, value=price)
                c_pr.font = normal_font
                c_pr.alignment = right_align
                c_pr.number_format = '#,##0'
                apply_box_border(ws, cur_row, 9, cur_row, 10)

                ws.merge_cells(start_row=cur_row, start_column=11, end_row=cur_row, end_column=12)
                c_mamt = ws.cell(row=cur_row, column=11, value=amt)
                c_mamt.font = normal_font
                c_mamt.alignment = right_align
                c_mamt.number_format = '#,##0'
                apply_box_border(ws, cur_row, 11, cur_row, 12)

        cur_row += 1
        ws.row_dimensions[cur_row].height = 24
        ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=10)
        c_mat_sum_lbl = ws.cell(row=cur_row, column=1, value="재료비 합계")
        c_mat_sum_lbl.font = bold_font
        c_mat_sum_lbl.fill = fill_grand_total
        c_mat_sum_lbl.alignment = center_align
        apply_box_border(ws, cur_row, 1, cur_row, 10)

        ws.merge_cells(start_row=cur_row, start_column=11, end_row=cur_row, end_column=12)
        c_mat_sum = ws.cell(row=cur_row, column=11, value=mat_total)
        c_mat_sum.font = bold_font
        c_mat_sum.fill = fill_grand_total
        c_mat_sum.alignment = right_align
        c_mat_sum.number_format = '#,##0'
        apply_box_border(ws, cur_row, 11, cur_row, 12)

        # ----------------------------------------------------
        # 4) 경비 및 외주 상세 (A~L열 풀-위드)
        # ----------------------------------------------------
        cur_row += 2
        ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=12)
        c_sub4 = ws.cell(row=cur_row, column=1, value="3) 경비 및 이익 상세 (Expense & Profit Detail)")
        c_sub4.font = bold_font
        c_sub4.alignment = left_align
        ws.row_dimensions[cur_row].height = 22

        exp_d = exp_d or {}

        def write_sub_expense_table(sub_title, col_headers, row_items, calc_func):
            nonlocal cur_row
            cur_row += 1
            ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=12)
            c_st = ws.cell(row=cur_row, column=1, value=sub_title)
            c_st.font = bold_font
            c_st.alignment = left_align
            ws.row_dimensions[cur_row].height = 20

            cur_row += 1
            exp_col_ranges = [(1, 4), (5, 6), (7, 8), (9, 10), (11, 12)]
            for (c_s, c_e), h in zip(exp_col_ranges, col_headers):
                ws.merge_cells(start_row=cur_row, start_column=c_s, end_row=cur_row, end_column=c_e)
                cell = ws.cell(row=cur_row, column=c_s, value=h)
                cell.font = header_font
                cell.fill = fill_sub_header
                cell.alignment = center_align
                apply_box_border(ws, cur_row, c_s, cur_row, c_e)
            ws.row_dimensions[cur_row].height = 26

            sub_sum = 0.0
            for item in row_items:
                cur_row += 1
                ws.row_dimensions[cur_row].height = 19
                row_vals, amt = calc_func(item)
                sub_sum += amt
                for (c_s, c_e), val in zip(exp_col_ranges, row_vals):
                    ws.merge_cells(start_row=cur_row, start_column=c_s, end_row=cur_row, end_column=c_e)
                    c = ws.cell(row=cur_row, column=c_s, value=val)
                    c.font = normal_font
                    if isinstance(val, (int, float)):
                        c.alignment = right_align
                        c.number_format = '#,##0'
                    elif c_s == 1:
                        c.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                    else:
                        c.alignment = center_align
                    apply_box_border(ws, cur_row, c_s, cur_row, c_e)

            cur_row += 1
            ws.row_dimensions[cur_row].height = 20
            ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=10)
            c_sub_lbl = ws.cell(row=cur_row, column=1, value=f"{sub_title} 소계")
            c_sub_lbl.font = bold_font
            c_sub_lbl.fill = fill_subtotal
            c_sub_lbl.alignment = center_align
            apply_box_border(ws, cur_row, 1, cur_row, 10)

            ws.merge_cells(start_row=cur_row, start_column=11, end_row=cur_row, end_column=12)
            c_sub_val = ws.cell(row=cur_row, column=11, value=sub_sum)
            c_sub_val.font = bold_font
            c_sub_val.fill = fill_subtotal
            c_sub_val.alignment = right_align
            c_sub_val.number_format = '#,##0'
            apply_box_border(ws, cur_row, 11, cur_row, 12)
            return sub_sum

        def calc_site_exp(item):
            cont = item.get('cont', '')
            ppl = to_float(item.get('ppl'))
            qty = to_float(item.get('qty'))
            pr = to_float(item.get('price'))
            amt = round(qty * pr)
            return [cont, ppl if ppl>0 else "", qty if qty>0 else "", pr, amt], amt

        sum_site_exp = write_sub_expense_table(
            "① 현장경비",
            ["항목(내용)", "투입인원", "수량/일수", "단가", cost_header_name],
            exp_d.get('site_expense', []),
            calc_site_exp
        )

        def calc_rental(item):
            spec = item.get('spec', '')
            qty = to_float(item.get('qty'))
            period = to_float(item.get('period'))
            pr = to_float(item.get('price'))
            amt = round(qty * period * pr)
            return [spec, qty if qty>0 else "", period if period>0 else "", pr, amt], amt

        sum_rental = write_sub_expense_table(
            "② 렌탈비",
            ["사양", "수량", "기간(일수)", "단가", cost_header_name],
            exp_d.get('rental', []),
            calc_rental
        )

        def calc_outsource(item):
            work = item.get('work', '')
            cnt = to_float(item.get('count'))
            note = item.get('note', '')
            pr = to_float(item.get('price'))
            amt = round(cnt * pr)
            return [work, cnt if cnt>0 else "", note, pr, amt], amt

        sum_outsource = write_sub_expense_table(
            "③ 외주비",
            ["작업내용", "인원/횟수", "비고", "단가", cost_header_name],
            exp_d.get('outsource', []),
            calc_outsource
        )

        def calc_deprec(item):
            name = item.get('item', '')
            qty = to_float(item.get('qty'))
            days = to_float(item.get('days'))
            rate = to_float(item.get('rate'))
            amt = round(qty * days * rate)
            return [name, qty if qty>0 else "", days if days>0 else "", rate, amt], amt

        sum_deprec = write_sub_expense_table(
            "④ 감가상각비",
            ["장비명", "대수", "사용일수", "상각율", cost_header_name],
            exp_d.get('depreciation', []),
            calc_deprec
        )

        cur_row += 1
        ws.row_dimensions[cur_row].height = 24
        ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=10)
        c_exp_total_lbl = ws.cell(row=cur_row, column=1, value="경비 및 외주 총 합계")
        c_exp_total_lbl.font = bold_font
        c_exp_total_lbl.fill = fill_grand_total
        c_exp_total_lbl.alignment = center_align
        apply_box_border(ws, cur_row, 1, cur_row, 10)

        tot_exp_all = sum_site_exp + sum_rental + sum_outsource + sum_deprec
        ws.merge_cells(start_row=cur_row, start_column=11, end_row=cur_row, end_column=12)
        c_exp_total_val = ws.cell(row=cur_row, column=11, value=tot_exp_all)
        c_exp_total_val.font = bold_font
        c_exp_total_val.fill = fill_grand_total
        c_exp_total_val.alignment = right_align
        c_exp_total_val.number_format = '#,##0'
        apply_box_border(ws, cur_row, 11, cur_row, 12)

        # 열 너비 설정 (세로보기 A4 인쇄 시 1페이지 너비 및 텍스트 가독성 최적화: 총 12열 풀-위드)
        ws.column_dimensions['A'].width = 11
        ws.column_dimensions['B'].width = 11
        ws.column_dimensions['C'].width = 12.5
        ws.column_dimensions['D'].width = 13.5
        ws.column_dimensions['E'].width = 11
        ws.column_dimensions['F'].width = 11
        ws.column_dimensions['G'].width = 11
        ws.column_dimensions['H'].width = 9.5
        ws.column_dimensions['I'].width = 9.5
        ws.column_dimensions['J'].width = 10
        ws.column_dimensions['K'].width = 10
        ws.column_dimensions['L'].width = 10

    # ====================================================
    # 시트 1: [사전예산서 (Planned)]
    # ====================================================
    ws1 = wb.active
    ws1.title = "사전예산서 (Planned)"
    write_sheet_top(ws1, "공사실행예산서(사전원가계산)", is_actual=False)

    # 2. 상단: 사전예산 요약 집계표 (Row 13부터 시작, A13:L13 병합)
    cur_r = 13
    ws1.merge_cells('A13:L13')
    c_p_sum_lbl = ws1.cell(row=cur_r, column=1, value="2. 사전예산 요약 집계")
    c_p_sum_lbl.font = section_font
    c_p_sum_lbl.alignment = left_align
    ws1.row_dimensions[cur_r].height = 24

    cur_r += 1
    # B:D(항목), E:G(사전예산 금액), H:L(비고/산출근거)
    headers_plan_sum = [("항목", 2, 4), ("사전예산 (계획) 금액", 5, 7), ("비고 / 산출근거", 8, 12)]
    for h, c_s, c_e in headers_plan_sum:
        ws1.merge_cells(start_row=cur_r, start_column=c_s, end_row=cur_r, end_column=c_e)
        cell = ws1.cell(row=cur_r, column=c_s, value=h)
        cell.font = header_font
        cell.fill = fill_summary_hdr
        cell.alignment = center_align
        apply_box_border(ws1, cur_r, c_s, cur_r, c_e)
    ws1.row_dimensions[cur_r].height = 22

    for label, plan_k, _, _ in summary_items_def:
        cur_r += 1
        ws1.row_dimensions[cur_r].height = 20
        ws1.merge_cells(start_row=cur_r, start_column=2, end_row=cur_r, end_column=4)
        c_lbl = ws1.cell(row=cur_r, column=2, value=label)
        c_lbl.font = bold_font
        c_lbl.alignment = left_align
        apply_box_border(ws1, cur_r, 2, cur_r, 4)

        val_plan = summary_data.get(plan_k, '')
        num_plan = to_float(val_plan)

        ws1.merge_cells(start_row=cur_r, start_column=5, end_row=cur_r, end_column=7)
        c_val = ws1.cell(row=cur_r, column=5, value=num_plan if '이익률' not in label else str(val_plan))
        c_val.font = normal_font
        c_val.alignment = right_align
        if '이익률' not in label: c_val.number_format = '#,##0'
        apply_box_border(ws1, cur_r, 5, cur_r, 7)

        ws1.merge_cells(start_row=cur_r, start_column=8, end_row=cur_r, end_column=12)
        c_note = ws1.cell(row=cur_r, column=8, value="계약금액/부가세별도" if "계약금액" in label else "")
        c_note.font = normal_font
        c_note.alignment = center_align
        apply_box_border(ws1, cur_r, 8, cur_r, 12)

    # 3. 하단: 사전예산 상세 내역
    write_details_section(ws1, cur_r + 2, "사전원가가액", labor_data, material_data, expense_data)

    # ====================================================
    # 시트 2: [사후정산서 (Actuals)]
    # ====================================================
    ws2 = wb.create_sheet(title="사후정산서 (Actuals)")
    write_sheet_top(ws2, "공사실행예산서(사후정산계산)", is_actual=True)

    # 2. 상단: [계획 vs 실적 vs 차액] 총괄 비교 집계표 (Row 13부터 시작, A13:L13 병합)
    cur_r2 = 13
    ws2.merge_cells('A13:L13')
    c_a_sum_lbl = ws2.cell(row=cur_r2, column=1, value="2. 실행예산 총괄 비교 집계 (계획 vs 실적 vs 차액)")
    c_a_sum_lbl.font = section_font
    c_a_sum_lbl.alignment = left_align
    ws2.row_dimensions[cur_r2].height = 24

    cur_r2 += 1
    # 헤더: B~D열(항목), E~F열(사전예산), G~H열(사후원가), I~J열(잔여예산), K~L열(달성률/증감)
    headers_act_sum = [
        ("항목", 2, 4),
        ("사전예산 (계획)", 5, 6),
        ("사후원가 (실적)", 7, 8),
        ("잔여예산 (차액)", 9, 10),
        ("달성률 / 증감", 11, 12)
    ]
    for h, c_s, c_e in headers_act_sum:
        ws2.merge_cells(start_row=cur_r2, start_column=c_s, end_row=cur_r2, end_column=c_e)
        cell = ws2.cell(row=cur_r2, column=c_s, value=h)
        cell.font = header_font
        cell.fill = fill_summary_hdr
        cell.alignment = center_align
        apply_box_border(ws2, cur_r2, c_s, cur_r2, c_e)
    ws2.row_dimensions[cur_r2].height = 22

    for label, plan_k, act_k, diff_k in summary_items_def:
        cur_r2 += 1
        ws2.row_dimensions[cur_r2].height = 20
        ws2.merge_cells(start_row=cur_r2, start_column=2, end_row=cur_r2, end_column=4)
        c_lbl = ws2.cell(row=cur_r2, column=2, value=label)
        c_lbl.font = bold_font
        c_lbl.alignment = left_align
        apply_box_border(ws2, cur_r2, 2, cur_r2, 4)

        val_plan = summary_data.get(plan_k, '')
        val_act = summary_data.get(act_k, '')
        val_diff = summary_data.get(diff_k, '')

        num_plan = to_float(val_plan)
        num_act = to_float(val_act)
        num_diff = to_float(val_diff)

        ws2.merge_cells(start_row=cur_r2, start_column=5, end_row=cur_r2, end_column=6)
        c_p = ws2.cell(row=cur_r2, column=5, value=num_plan if '이익률' not in label else str(val_plan))
        c_p.font = normal_font
        c_p.alignment = right_align
        if '이익률' not in label: c_p.number_format = '#,##0'
        apply_box_border(ws2, cur_r2, 5, cur_r2, 6)

        ws2.merge_cells(start_row=cur_r2, start_column=7, end_row=cur_r2, end_column=8)
        c_a = ws2.cell(row=cur_r2, column=7, value=num_act if '이익률' not in label else str(val_act))
        c_a.font = normal_font
        c_a.alignment = right_align
        if '이익률' not in label: c_a.number_format = '#,##0'
        apply_box_border(ws2, cur_r2, 7, cur_r2, 8)

        ws2.merge_cells(start_row=cur_r2, start_column=9, end_row=cur_r2, end_column=10)
        c_d = ws2.cell(row=cur_r2, column=9, value=num_diff if '이익률' not in label else str(val_diff))
        c_d.font = normal_font
        c_d.alignment = right_align
        if '이익률' not in label: c_d.number_format = '#,##0'
        apply_box_border(ws2, cur_r2, 9, cur_r2, 10)

        rate_str = "-"
        if '이익률' not in label and num_plan > 0:
            rate = (num_act / num_plan) * 100
            rate_str = f"{rate:.1f}%"
        elif '이익률' in label:
            rate_str = str(val_act)
            
        ws2.merge_cells(start_row=cur_r2, start_column=11, end_row=cur_r2, end_column=12)
        c_r = ws2.cell(row=cur_r2, column=11, value=rate_str)
        c_r.font = normal_font
        c_r.alignment = center_align
        apply_box_border(ws2, cur_r2, 11, cur_r2, 12)

    # 3. 하단: 사후원가 실적 상세 내역
    write_details_section(
        ws2,
        cur_r2 + 2,
        "사후원가가액",
        actual_labor_data if actual_labor_data is not None else labor_data,
        actual_material_data if actual_material_data is not None else material_data,
        actual_expense_data if actual_expense_data is not None else expense_data
    )

    wb.save(filepath)
    return True
