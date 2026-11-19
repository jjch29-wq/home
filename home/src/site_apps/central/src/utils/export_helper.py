import os
import datetime
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

def export_budget_estimation_to_excel(filepath, summary_data, labor_data, material_data, expense_data,
                                       actual_labor_data=None, actual_material_data=None, actual_expense_data=None):
    """
    Export the detailed budget estimation to a formatted Excel file.
    Prefers the official company standard 15-sheet template if available.
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
            return _export_budget_legacy_excel(filepath, summary_data, labor_data, material_data, expense_data)
    else:
        return _export_budget_legacy_excel(filepath, summary_data, labor_data, material_data, expense_data)


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

    central_workers = [
        ("주진철", "부장", "중앙", "자사", "PAUT/MT/PT"),
        ("우명광", "대리", "중앙", "자사", "PAUT/MT/PT"),
        ("장승대", "계장", "중앙", "자사", "PAUT/MT/PT"),
        ("김진환", "주임", "중앙", "자사", "PAUT/MT/PT"),
        ("김성렬", "주임", "중앙", "자사", "PAUT/MT/PT"),
        ("박광복", "부장", "중앙", "자사", "PAUT/MT/PT"),
        ("주영광", "과장", "중앙", "자사", "PAUT/MT/PT"),
        ("박상태", "계장", "중앙", "자사", "PAUT/MT/PT")
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

    # 2. Sheet: 인원투입계획 (사전)
    if '인원투입계획' in wb.sheetnames:
        ws_p_labor = wb['인원투입계획']
        for idx, w_info in enumerate(central_workers):
            r_w = 7 + idx
            ws_p_labor[f'B{r_w}'] = w_info[0]
            ws_p_labor[f'C{r_w}'] = w_info[1]
            ws_p_labor[f'D{r_w}'] = w_info[2]
            ws_p_labor[f'E{r_w}'] = w_info[3]
            ws_p_labor[f'F{r_w}'] = w_info[4]
            for c_col in range(13, 58):
                ws_p_labor.cell(r_w, c_col).value = None

        for r_clear in range(7 + len(central_workers), 37):
            ws_p_labor[f'B{r_clear}'] = ""
            ws_p_labor[f'C{r_clear}'] = ""
            ws_p_labor[f'D{r_clear}'] = ""
            ws_p_labor[f'E{r_clear}'] = ""
            ws_p_labor[f'F{r_clear}'] = ""
            for c_col in range(13, 58):
                ws_p_labor.cell(r_clear, c_col).value = None

        if labor_data:
            for rank, r_idx in rank_to_row.items():
                if rank in labor_data:
                    vals = labor_data[rank]
                    p = _to_f(vals.get('personnel'))
                    d = _to_f(vals.get('period'))
                    ws_p_labor[f'E{r_idx}'] = p if p > 0 else ""
                    ws_p_labor[f'F{r_idx}'] = d if d > 0 else ""

    # 3. Sheet: 경비예상지출항목 (사전)
    if '경비예상지출항목' in wb.sheetnames:
        ws_p_exp = wb['경비예상지출항목']
        if labor_data:
            ws_p_exp['D6'] = _to_f(labor_data.get('연장근무', {}).get('period', 0))
            ws_p_exp['D7'] = _to_f(labor_data.get('야간근무', {}).get('period', 0))
            ws_p_exp['D8'] = _to_f(labor_data.get('휴일근무', {}).get('period', 0))

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
    if '장비감가상각' in wb.sheetnames and expense_data:
        ws_p_dep = wb['장비감가상각']
        for idx, d_item in enumerate(expense_data.get('depreciation', [])[:8]):
            d_row = 5 + idx
            item_name = d_item.get('item', '')
            spec = d_item.get('spec', '')
            qty = _to_f(d_item.get('qty', 0))
            days = _to_f(d_item.get('days', 0))
            rate = _to_f(d_item.get('rate', 0))
            if item_name:
                ws_p_dep[f'B{d_row}'] = item_name
                ws_p_dep[f'C{d_row}'] = spec
                ws_p_dep[f'D{d_row}'] = qty
                ws_p_dep[f'G{d_row}'] = days
                ws_p_dep[f'H{d_row}'] = rate

    # 5. Sheet: 인원투입계획_사후
    if '인원투입계획_사후' in wb.sheetnames:
        ws_a_labor = wb['인원투입계획_사후']
        for idx, w_info in enumerate(central_workers):
            r_w = 7 + idx
            ws_a_labor[f'B{r_w}'] = w_info[0]
            ws_a_labor[f'C{r_w}'] = w_info[1]
            ws_a_labor[f'D{r_w}'] = w_info[2]
            ws_a_labor[f'E{r_w}'] = w_info[3]
            ws_a_labor[f'F{r_w}'] = w_info[4]
            for c_col in range(13, 58):
                ws_a_labor.cell(r_w, c_col).value = None

        for r_clear in range(7 + len(central_workers), 37):
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

        for r_clr in range(25, 44):
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

            for r_clr in range(49, 53):
                ws_a_exp[f'D{r_clr}'] = 0
                ws_a_exp[f'E{r_clr}'] = ""
                ws_a_exp[f'G{r_clr}'] = 1.0
            for idx, r_item in enumerate(target_a_exp.get('rental', [])[:4]):
                r_row = 49 + idx
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

            for r_clr in range(58, 64):
                ws_a_exp[f'D{r_clr}'] = 0
                ws_a_exp[f'E{r_clr}'] = ""
                ws_a_exp[f'G{r_clr}'] = 1.0
            for idx, o_item in enumerate(target_a_exp.get('outsource', [])[:6]):
                o_row = 58 + idx
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
        if target_a_exp:
            for idx, d_item in enumerate(target_a_exp.get('depreciation', [])[:8]):
                d_row = 5 + idx
                item_name = d_item.get('item', '')
                spec = d_item.get('spec', '')
                qty = _to_f(d_item.get('qty', 0))
                days = _to_f(d_item.get('days', 0))
                rate = _to_f(d_item.get('rate', 0))
                if item_name:
                    ws_a_dep[f'B{d_row}'] = item_name
                    ws_a_dep[f'C{d_row}'] = spec
                    ws_a_dep[f'D{d_row}'] = qty
                    ws_a_dep[f'G{d_row}'] = days
                    ws_a_dep[f'H{d_row}'] = rate

    # 8. Sheet: 실적배분
    if '실적배분' in wb.sheetnames:
        ws_dist = wb['실적배분']
        ws_dist['F2'] = site_title
        ws_dist['F4'] = a_rev if a_rev > 0 else rev

    wb.save(filepath)
    return True


def _export_budget_legacy_excel(filepath, summary_data, labor_data, material_data, expense_data):
    """
    Fallback legacy single-sheet export when standard template is absent.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "공사실행예산서"
    
    # Styles
    title_font = Font(name='맑은 고딕', size=16, bold=True)
    header_font = Font(name='맑은 고딕', size=11, bold=True)
    bold_font = Font(name='맑은 고딕', size=10, bold=True)
    normal_font = Font(name='맑은 고딕', size=10)
    
    center_align = Alignment(horizontal='center', vertical='center')
    right_align = Alignment(horizontal='right', vertical='center')
    
    thin = Side(border_style="thin", color="000000")
    border_all = Border(top=thin, left=thin, right=thin, bottom=thin)
    
    header_fill = PatternFill(start_color='E0E0E0', end_color='E0E0E0', fill_type='solid')
    sub_fill = PatternFill(start_color='F5F5F5', end_color='F5F5F5', fill_type='solid')
    yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
    blue_fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')

    def to_float(val):
        try:
            return float(str(val).replace(',', '').replace('₩', '').replace('%', '').strip() or 0)
        except:
            return 0.0

    # 1. Title
    ws.merge_cells('A1:G1')
    c = ws['A1']
    c.value = f"[{summary_data.get('site', '')}] 공사실행예산서 (사정원가)"
    c.font = Font(name='맑은 고딕', size=18, bold=True)
    c.alignment = center_align
    ws.row_dimensions[1].height = 30
    
    # Write metadata (Print Date)
    ws.merge_cells('A2:B2')
    ws['A2'] = "출력일시: " + datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    ws['A2'].font = normal_font
    ws['A2'].alignment = Alignment(horizontal='left', vertical='top')

    # 1.5 Approval Box (결재란)
    approval_headers = ["소장", "담당", "상무", "본부장", "사장"]
    start_col = 3 # Column C
    
    for i, h in enumerate(approval_headers):
        col_letter = get_column_letter(start_col + i)
        
        # Header cell (Row 2)
        h_cell = ws[f"{col_letter}2"]
        h_cell.value = h
        h_cell.font = bold_font
        h_cell.alignment = center_align
        h_cell.border = border_all
        
        # Stamp cell (Row 3)
        s_cell = ws[f"{col_letter}3"]
        s_cell.border = border_all
        s_cell.alignment = center_align
        
    ws.row_dimensions[2].height = 20
    ws.row_dimensions[3].height = 60
    
    # Helper function to write a table
    def write_table(start_row, title, headers, data_rows, total_label=None, total_val=None):
        # Table Title
        ws.cell(row=start_row, column=1, value=title).font = header_font
        ws.cell(row=start_row, column=1).fill = header_fill
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=len(headers))
        for col in range(1, len(headers)+1):
            ws.cell(row=start_row, column=col).border = border_all
            ws.cell(row=start_row, column=col).alignment = center_align
            
        # Headers
        r = start_row + 1
        for col, h in enumerate(headers, 1):
            c = ws.cell(row=r, column=col, value=h)
            c.font = bold_font
            c.fill = sub_fill
            c.border = border_all
            c.alignment = center_align
                
        # Data Rows
        r += 1
        for row_data in data_rows:
            for col, val in enumerate(row_data, 1):
                c = ws.cell(row=r, column=col, value=val)
                c.font = normal_font
                c.border = border_all
                if isinstance(val, (int, float)) or (isinstance(val, str) and val.replace(',', '').replace('.', '').replace('-', '').isdigit()):
                    c.number_format = '#,##0'
                    c.alignment = right_align
                    if isinstance(val, str):
                        try: c.value = float(val.replace(',', ''))
                        except: pass
                else:
                    c.alignment = center_align
            r += 1
            
        # Total Row
        if total_label is not None:
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=len(headers)-1)
            c = ws.cell(row=r, column=1, value=total_label)
            c.font = bold_font
            c.fill = yellow_fill
            c.border = border_all
            c.alignment = center_align
            for col in range(2, len(headers)):
                ws.cell(row=r, column=col).border = border_all
                
            c_tot = ws.cell(row=r, column=len(headers), value=total_val)
            c_tot.font = bold_font
            c_tot.fill = yellow_fill
            c_tot.border = border_all
            c_tot.alignment = right_align
            c_tot.number_format = '#,##0'
            try: c_tot.value = float(str(total_val).replace(',', '').replace('₩', '').strip())
            except: pass
            r += 1
            
        return r + 1

    # Set Global Column Widths to ensure all text and large numbers fit perfectly
    # Set Global Column Widths
    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 12
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 20
    ws.column_dimensions['F'].width = 20
    ws.column_dimensions['G'].width = 12

    # Print / Page Setup: Fit to one page wide, Landscape, A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = 9


    # 1.8 Site Metadata Info
    site_info_row = 5
    info_font = Font(name='맑은 고딕', size=11, bold=True)
    
    ws.cell(row=site_info_row, column=1, value="현 장 명 :").font = info_font
    ws.cell(row=site_info_row, column=2, value=f"{summary_data.get('site', '')} 중 비파괴검사").font = info_font
    
    ws.cell(row=site_info_row+1, column=1, value="공사기간 :").font = info_font
    period_str = summary_data.get('period', '')
    ws.cell(row=site_info_row+1, column=2, value=period_str).font = info_font
    
    days_diff_str = ""
    if period_str and "~" in period_str:
        try:
            start_str, end_str = period_str.split("~")
            start_date = datetime.datetime.strptime(start_str.strip(), "%Y.%m.%d")
            end_date = datetime.datetime.strptime(end_str.strip(), "%Y.%m.%d")
            days_diff = (end_date - start_date).days + 1
            days_diff_str = f"{days_diff}"
        except Exception:
            pass

    ws.cell(row=site_info_row+1, column=5, value=days_diff_str).font = info_font
    ws.cell(row=site_info_row+1, column=5).alignment = Alignment(horizontal='right', vertical='center')
    ws.cell(row=site_info_row+1, column=6, value="일").font = info_font
    ws.cell(row=site_info_row+1, column=6).alignment = Alignment(horizontal='left', vertical='center')
    
    ws.cell(row=site_info_row+2, column=1, value="계약업체 :").font = info_font
    ws.cell(row=site_info_row+2, column=2, value="").font = info_font
    
    ws.cell(row=site_info_row+3, column=1, value="경인사무소 :").font = info_font
    ws.cell(row=site_info_row+3, column=2, value="").font = info_font

    current_row = site_info_row + 5

    # 2. Summary Table
    ws.cell(row=current_row, column=1, value="1) 사전예산 요약").font = header_font
    ws.cell(row=current_row, column=1).fill = blue_fill
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=2)
    for col in range(1, 3):
        ws.cell(row=current_row, column=col).border = border_all
        ws.cell(row=current_row, column=col).alignment = center_align
    current_row += 1

    summary_items = [
        ("계약금액(Revenue)", summary_data.get('revenue', 0)),
        ("매출금액(UnitPrice)", summary_data.get('unitprice', 0)),
        ("실행 노무비(Labor)", summary_data.get('labor', 0)),
        ("실행 재료비(Material)", summary_data.get('material', 0)),
        ("실행 경비(Expense)", summary_data.get('expense', 0)),
        ("실행 외주비(Outsource)", summary_data.get('outsource', 0)),
        ("영업이익(Profit)", summary_data.get('profit', 0)),
        ("이익률(Margin)", summary_data.get('margin', '0%'))
    ]
    
    for k, v in summary_items:
        ws.cell(row=current_row, column=1, value=k).font = bold_font
        ws.cell(row=current_row, column=1).fill = sub_fill
        ws.cell(row=current_row, column=1).border = border_all
        
        if k == "이익률(Margin)" and '%' not in str(v):
            v = f"{v}%"
            
        c_val = ws.cell(row=current_row, column=2, value=v)
        c_val.font = normal_font
        c_val.border = border_all
        if '%' in str(v):
            c_val.alignment = right_align
        else:
            c_val.number_format = '#,##0'
            c_val.alignment = right_align
            try: c_val.value = float(str(v).replace(',', '').replace('₩', '').strip())
            except: pass
        current_row += 1
        
    current_row += 2

    # 3. Labor Data
    if labor_data:
        # Table 1: Regular
        reg_rows = []
        reg_total = 0
        regular_keys = ['부장', '차장', '과장', '대리', '사원', '초급', '중급', '고급', '특급']
        for rank in regular_keys:
            if rank in labor_data:
                vals = labor_data[rank]
                p = to_float(vals.get('personnel'))
                d = to_float(vals.get('period'))
                u = to_float(vals.get('unit_price'))
                amt = p * d * u
                if p > 0 or d > 0:
                    reg_rows.append([rank, p, d, u, amt])
                    reg_total += amt
        
        if reg_rows:
            current_row = write_table(
                current_row, "2) 노무비 - 정시근무", 
                ["직급", "투입인원(명)", "투입기간(일)", "단가", "사전원가가액"], 
                reg_rows, "정시근무 합계", reg_total
            )

        # Table 2: Special
        sp_rows = []
        sp_total = 0
        special_keys = ['연장근무', '야간근무', '휴일근무']
        for stype in special_keys:
            if stype in labor_data:
                vals = labor_data[stype]
                p = to_float(vals.get('personnel'))
                h = to_float(vals.get('period'))
                u = to_float(vals.get('unit_price'))
                amt = p * h * u
                if p > 0 or h > 0:
                    sp_rows.append([stype, p, h, u, amt])
                    sp_total += amt
                
        if sp_rows:
            current_row = write_table(
                current_row, "2) 노무비 - 특별근무", 
                ["형태", "투입인원(명)", "투입시간(시간)", "단가", "사전원가가액"], 
                sp_rows, "특별근무 합계", sp_total
            )

    # 4. Material Data
    if material_data:
        mat_rows = []
        mat_total = 0
        for item in material_data:
            qty = to_float(item.get('qty'))
            price = to_float(item.get('price'))
            amt = qty * price
            if qty > 0:
                mat_rows.append([item.get('name'), item.get('spec'), qty, item.get('unit'), price, amt])
                mat_total += amt
        if mat_rows:
            current_row = write_table(
                current_row, "3) 재료비", 
                ["품목", "사양", "수량", "규격", "단가", "사전원가가액"], 
                mat_rows, "재료비 합계", mat_total
            )
            
    # 5. Expense Data
    if expense_data:
        # Site Expense
        exp_rows = []
        exp_total = 0
        for item in expense_data.get('site_expense', []):
            cont = item.get('cont', '')
            ppl = to_float(item.get('ppl'))
            qty = to_float(item.get('qty'))
            price = to_float(item.get('price'))
            amt = qty * price
            if qty > 0 or ppl > 0:
                exp_rows.append([cont, ppl, qty, price, amt])
                exp_total += amt
        if exp_rows:
            current_row = write_table(
                current_row, "4) 경비 - 현장경비", 
                ["항목(내용)", "투입인원", "수량/일수", "단가", "사전원가가액"], 
                exp_rows, "현장경비 합계", exp_total
            )
            
        # Rental
        rent_rows = []
        rent_total = 0
        for item in expense_data.get('rental', []):
            spec = item.get('spec', '')
            qty = to_float(item.get('qty'))
            period = to_float(item.get('period'))
            price = to_float(item.get('price'))
            amt = qty * period * price
            if qty > 0 or period > 0:
                rent_rows.append([spec, qty, period, price, amt])
                rent_total += amt
        if rent_rows:
            current_row = write_table(
                current_row, "4) 경비 - 렌탈비", 
                ["사양", "수량", "기간(일수)", "단가", "사전원가가액"], 
                rent_rows, "렌탈비 합계", rent_total
            )
            
        # Outsource
        out_rows = []
        out_total = 0
        for item in expense_data.get('outsource', []):
            work = item.get('work', '')
            count = to_float(item.get('count'))
            price = to_float(item.get('price'))
            amt = count * price
            if count > 0:
                out_rows.append([work, count, "", price, amt])
                out_total += amt
        if out_rows:
            current_row = write_table(
                current_row, "4) 경비 - 외주비", 
                ["작업내용", "인원/횟수", "비고", "단가", "사전원가가액"], 
                out_rows, "외주비 합계", out_total
            )
            
        # Depreciation
        dep_rows = []
        dep_total = 0
        for item in expense_data.get('depreciation', []):
            name = item.get('item', '')
            qty = to_float(item.get('qty'))
            days = to_float(item.get('days'))
            rate = to_float(item.get('rate'))
            amt = qty * days * rate
            if qty > 0 or days > 0:
                dep_rows.append([name, qty, days, rate, amt])
                dep_total += amt
        if dep_rows:
            current_row = write_table(
                current_row, "4) 경비 - 감가상각", 
                ["장비명", "대수", "사용일수", "상각율", "사전원가가액"], 
                dep_rows, "감가상각 합계", dep_total
            )

    wb.save(filepath)
    return True
