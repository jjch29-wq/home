# Excel exporter for central-site NDT billing statements.
import json
import os
from datetime import datetime
from time import perf_counter
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import win32com.client as win32
from pywintypes import com_error

from site_apps.central.src.data_sync import get_history_path
from site_apps.central.src.services.billing_totals import grouped_supply_amount


def export_to_excel(self):
    target_records = []
    if not self.records:
        if not messagebox.askyesno("기록 없음", "출력할 작업 기록(금회 기성)이 없습니다. 계약 내역만 출력하시겠습니까?"):
            return
    else:
        selected_items = self.tree.selection()
        if selected_items:
            if not messagebox.askyesno("부분 출력", f"선택된 {len(selected_items)}개의 기록만 기성 청구 내역서로 출력하시겠습니까?"):
                return
            indices = []
            for item in selected_items:
                tags = self.tree.item(item, 'tags')
                try:
                    indices.append(int(tags[0]) if tags else self.tree.index(item))
                except (ValueError, IndexError):
                    indices.append(self.tree.index(item))
            target_records = [self.records[i] for i in indices if 0 <= i < len(self.records)]
        else:
            if not messagebox.askyesno("전체 출력", "선택된 항목이 없습니다. 전체 기록을 기성 청구 내역서로 출력하시겠습니까?"):
                return
            target_records = self.records

    dates_all = sorted([r["date"] for r in target_records])

    user_period = getattr(self, "billing_period_var", None)
    user_period_val = user_period.get().strip() if user_period else ""

    billed_nrs = set()

    if user_period_val:
        global_period = user_period_val
    elif dates_all:
        start_date = dates_all[0].replace("-", ".")
        end_date = dates_all[-1].replace("-", ".")
        global_period = f"{start_date} ~ {end_date}" if start_date != end_date else start_date
    else:
        global_period = "기간 없음"

    round_val = self.round_var.get()
    default_name = f"제{round_val}회_기성청구내역서_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=default_name, filetypes=[("Excel File", "*.xlsx")], title="정식 기성청구 엑셀 양식으로 저장")
    if not filepath: return

    export_started = perf_counter()
    stage_started = export_started
    stage_times = []

    def record_stage(label):
        nonlocal stage_started
        now = perf_counter()
        stage_times.append(f"{label} {now - stage_started:.2f}초")
        stage_started = now

    progress_win = None
    excel = None

    def freeze_header(sheet, rows):
        try:
            sheet.Activate()
            window = wb.Windows(1)
            window.FreezePanes = False
            window.SplitColumn = 0
            window.SplitRow = rows
            window.FreezePanes = True
            window.ScrollRow = 1
            window.ScrollColumn = 1
        except com_error as error:
            print(f"[기성청구 엑셀] 상단 {rows}행 틀 고정 설정 실패: {error}", flush=True)

    try:
        try:
            top_owner = self.winfo_toplevel()
            progress_win = tk.Toplevel(top_owner)
            progress_win.title("기성청구 엑셀 생성 중")
            progress_win.geometry("360x105")
            progress_win.resizable(False, False)
            progress_win.transient(top_owner)
            ttk.Label(
                progress_win,
                text="기성청구 내역서 엑셀 파일을 생성하고 있습니다...\n잠시만 기다려 주세요.",
                justify=tk.CENTER,
                font=("맑은 고딕", 10, "bold"),
            ).pack(expand=True, pady=20)
            progress_win.update()
        except Exception:
            progress_win = None

        record_stage("안내창")
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            excel.ScreenUpdating = False
            excel.EnableEvents = False
        except Exception:
            pass

        wb = excel.Workbooks.Add()
        try:
            excel.Calculation = -4135  # xlCalculationManual
        except Exception:
            pass
        ws = wb.ActiveSheet
        ws.Name = "기성청구내역서"
        record_stage("Excel 시작 및 통합문서 준비")

        # --- 상단 기본 정보 ---
        ws.Range("A1:O2").Merge()
        ws.Range("A1").Value = f"제 {round_val} 회 비파괴검사기술용역 기성청구 내역서"
        ws.Range("A1").Font.Size = 20
        ws.Range("A1").Font.Bold = True
        ws.Range("A1").HorizontalAlignment = -4108
        ws.Range("A1").VerticalAlignment = -4108

        ws.Range("A4:J4").Merge()
        ws.Range("A4").Value = "공 사 명 :  2026년 중앙지사 열수송관 비파괴검사용역 단가계약"
        ws.Range("A4").Font.Bold = True
        ws.Range("A4").Font.Size = 14
        ws.Range("A4").HorizontalAlignment = -4131 # xlLeft

        ws.Range("K4:L4").Merge()
        ws.Range("K4").Value = "청구 기간 :"
        ws.Range("K4").Font.Bold = True
        ws.Range("K4").Font.Size = 11
        ws.Range("K4").HorizontalAlignment = -4152 # xlRight

        ws.Range("M4:O4").Merge()
        ws.Range("M4").Value = global_period
        ws.Range("M4").Font.Size = 11
        ws.Range("M4").HorizontalAlignment = -4131 # xlLeft

        # --- 기성 요약 테이블 ---
        ws.Cells(6, 1).Value = "공종"
        ws.Range(ws.Cells(6, 1), ws.Cells(7, 1)).Merge()
        ws.Cells(6, 2).Value = "규격"
        ws.Range(ws.Cells(6, 2), ws.Cells(7, 3)).Merge()
        ws.Cells(6, 4).Value = "단위"
        ws.Range(ws.Cells(6, 4), ws.Cells(7, 4)).Merge()

        ws.Cells(6, 5).Value = "단가"
        ws.Range(ws.Cells(6, 5), ws.Cells(7, 5)).Merge()

        groups = ["계약", "전회까지기성", "금회기성", "누계기성", "잔액"]
        for idx, g_name in enumerate(groups):
            start_col = 6 + idx * 2
            ws.Cells(6, start_col).Value = g_name
            ws.Range(ws.Cells(6, start_col), ws.Cells(6, start_col + 1)).Merge()
            ws.Cells(7, start_col).Value = "수량"
            ws.Cells(7, start_col + 1).Value = "금액"

        for c in range(1, 16):
            ws.Cells(6, c).Font.Bold = True
            ws.Cells(7, c).Font.Bold = True
            ws.Cells(6, c).HorizontalAlignment = -4108
            ws.Cells(7, c).HorizontalAlignment = -4108
            ws.Cells(6, c).Interior.Color = 14277081
            ws.Cells(7, c).Interior.Color = 14277081

        extra_items_total = sum([self.equip_cost_var.get(), self.safety_cost_var.get(), self.travel_cost_var.get(), self.print_cost_var.get(), self.liability_cost_var.get()])

        # 계약서 원본 순서: 검사 규격별로 일반/야간을 나란히 배치한다.
        material_order = [
            "PAUT_300A 이상", "PAUT_250A", "PAUT_200A",
            "PAUT_150A-125A", "PAUT_100A 이하",
            'RT_3 1/3 x 12"', 'RT_3 1/3 x 6"', "MT_MT", "PT_PT"
        ]
        preferred_categories = [
            f"열배관_{work_time}_{material}"
            for material in material_order
            for work_time in ("일반", "야간")
            if f"열배관_{work_time}_{material}" in self.contract_vars
        ]
        # 업체별 기성요약에 집계되는 실제 청구 항목이 기성청구내역서에서
        # 누락되지 않도록, 계약서 기본 순서에 없는 계약 항목도 모두 포함한다.
        # (플랜트/추가 규격 등 현장별로 확장된 항목 대응)
        categories = preferred_categories + [
            cat for cat in self.contract_vars
            if cat not in preferred_categories
        ]
        # 금액이 존재하는 실비 항목만 내역서 표에 출력 (0원인 항목 숨김 처리)
        for cat, k in [("장비손료", "equip"), ("안전관리비", "safety"), ("주재비 및 출장여비", "travel"), ("도서인쇄비", "print")]:
            if self.get_int(self.exp_vars[k]["budget"]) > 0 or self.get_int(self.exp_vars[k]["prev"]) > 0 or self.exp_vars[k]["curr"].get() > 0:
                categories.append(cat)
        categories.extend(["기타실비 소계", "엔지니어링 손해배상공제료", "공 급 가 액", "부가가치세", "합        계"])

        ws.Range(ws.Cells(6, 1), ws.Cells(6 + len(categories) + 2, 15)).Borders.LineStyle = 1

        row = 8
        ws.Cells(row, 1).Value = "□ 비파괴검사용역비"
        ws.Range(ws.Cells(row, 1), ws.Cells(row, 3)).Merge()
        ws.Cells(row, 1).Font.Bold = True
        ws.Cells(row, 1).HorizontalAlignment = -4131
        row += 1

        data_rows = []
        extra_rows = []
        subtotal_row = 0
        liability_row = 0
        total_row = 0
        contract_item_no = 0

        for cat in categories:
            c_qty, p_qty, cur_qty, tot_qty, rem_qty = "", "", "", "", ""
            c_amt, p_amt, cur_amt, tot_amt, rem_amt = 0, 0, 0, 0, 0

            if cat == "공 급 가 액":
                c_amt = self.get_int(self.total_contract_var)
                p_amt = self.get_int(self.total_prev_var)
                cur_amt = sum(r["subtotal"] for r in target_records) + extra_items_total
                total_row = row
            elif cat in ["부가가치세", "합        계"]:
                pass
            elif cat == "기타실비 소계":
                c_amt = sum(self.get_int(self.exp_vars[k]["budget"]) for k in ["equip", "safety", "travel", "print"])
                p_amt = sum(self.get_int(self.exp_vars[k]["prev"]) for k in ["equip", "safety", "travel", "print"])
                cur_amt = sum(self.exp_vars[k]["curr"].get() for k in ["equip", "safety", "travel", "print"])
                subtotal_row = row
            elif cat in ["장비손료", "안전관리비", "주재비 및 출장여비", "도서인쇄비", "엔지니어링 손해배상공제료"]:
                if cat == "엔지니어링 손해배상공제료": liability_row = row
                else: extra_rows.append(row)
                k = ""
                if cat == "장비손료": k = "equip"
                elif cat == "안전관리비": k = "safety"
                elif cat == "주재비 및 출장여비": k = "travel"
                elif cat == "엔지니어링 손해배상공제료": k = "liability"
                else: k = "print"
                c_amt = self.get_int(self.exp_vars[k]["budget"])
                p_amt = self.get_int(self.exp_vars[k]["prev"])
                cur_amt = self.exp_vars[k]["curr"].get()
            else:
                c_qty = self.get_float(self.contract_vars[cat]["c_qty"])
                p_qty = self.get_float(self.contract_vars[cat]["p_qty"])
                c_amt = self.get_int(self.contract_vars[cat]["contract"])
                p_amt = self.get_int(self.contract_vars[cat]["prev"])

                cur_qty = 0.0
                cur_amt = 0
                for r in target_records:
                    loc = "플랜트(관리소)" if "관리소" in r["loc"] or "플랜트" in r.get("loc_type", r["loc"]) else "열배관"
                    t_time = r.get("work_time", "일반")
                    mat = f"{r['ndt_type']}_{r['material_type']}"
                    key = f"{loc}_{t_time}_{mat}"
                    if key == cat:
                        cur_qty += r["qty"]
                        cur_amt += r["subtotal"]

                data_rows.append(row)

            if cat in ["공 급 가 액", "부가가치세", "합        계", "기타실비 소계", "장비손료", "안전관리비", "주재비 및 출장여비", "도서인쇄비", "엔지니어링 손해배상공제료"]:
                ws.Cells(row, 1).Value = cat
                ws.Range(ws.Cells(row, 1), ws.Cells(row, 5)).Merge()
                ws.Cells(row, 1).HorizontalAlignment = -4108

                if cat == "합        계":
                    ws.Range(ws.Cells(row, 1), ws.Cells(row, 15)).Interior.Color = 10066329
                    ws.Range(ws.Cells(row, 1), ws.Cells(row, 15)).Font.Color = 16777215
                    ws.Range(ws.Cells(row, 1), ws.Cells(row, 15)).Font.Bold = True
                elif cat in ["공 급 가 액", "기타실비 소계"]:
                    ws.Range(ws.Cells(row, 1), ws.Cells(row, 15)).Interior.Color = 15987699
                    ws.Cells(row, 1).Font.Bold = True
            else:
                parts = cat.split('_')
                loc = parts[0]
                t_time = parts[1]
                m_key = '_'.join(parts[2:])
                unit = "매" if m_key.startswith("RT") else "M"
                contract_item_no += 1
                if m_key.startswith("PAUT"):
                    work_name = "위상배열초음파검사(PAUT)"
                    spec = m_key.removeprefix("PAUT_")
                elif m_key.startswith("RT"):
                    work_name = "방사선투과검사(RT)"
                    spec = m_key.removeprefix("RT_")
                elif m_key.startswith("MT"):
                    work_name = "자분탐상검사(MT)"
                    spec = ""
                else:
                    work_name = "액체침투탐상검사(PT)"
                    spec = ""
                if t_time == "야간":
                    spec = f"{spec}, 야간" if spec else "야간"

                ws.Cells(row, 1).Value = f"{contract_item_no})    {work_name}"
                ws.Cells(row, 2).Value = spec
                ws.Range(ws.Cells(row, 2), ws.Cells(row, 3)).Merge()
                ws.Cells(row, 4).Value = unit
                unit_price = self.contract_vars[cat].get("c_price", 0)
                if unit_price == 0:
                    unit_price = c_amt / c_qty if c_qty > 0 else (cur_amt / cur_qty if cur_qty > 0 else (p_amt / p_qty if p_qty > 0 else 0))
                if unit_price > 0:
                    ws.Cells(row, 5).Value = int(unit_price)
                    ws.Cells(row, 5).NumberFormat = "#,##0"
                else:
                    ws.Cells(row, 5).Value = 0
                    ws.Cells(row, 5).NumberFormat = '#,##0;-#,##0;"-"'
                ws.Range(ws.Cells(row, 1), ws.Cells(row, 5)).HorizontalAlignment = -4108

            num_fmt = '#,##0;-#,##0;"-"'
            float_fmt = '#,##0.0000;-#,##0.0000;"-"'

            if cat == "공 급 가 액":
                ws.Cells(row, 7).Value = c_amt
                ws.Cells(row, 7).NumberFormat = num_fmt
                ws.Cells(row, 7).Font.Bold = True

                for col, l in zip([9, 11], ['I', 'K']):
                    f1 = f"SUM({l}{data_rows[0]}:{l}{data_rows[-1]})" if data_rows else "0"
                    f2 = f"{l}{subtotal_row}" if subtotal_row else "0"
                    f3 = f"{l}{liability_row}" if liability_row else "0"
                    ws.Cells(row, col).Formula = f"={f1}+{f2}+{f3}"
                    ws.Cells(row, col).NumberFormat = num_fmt
                    ws.Cells(row, col).Font.Bold = True

                ws.Cells(row, 13).Formula = f"=I{row}+K{row}"
                ws.Cells(row, 13).NumberFormat = num_fmt
                ws.Cells(row, 13).Font.Bold = True

                ws.Cells(row, 15).Formula = f"=G{row}-M{row}"
                ws.Cells(row, 15).NumberFormat = num_fmt
                ws.Cells(row, 15).Font.Bold = True

                for col in [6, 8, 10, 12, 14]:
                    ws.Cells(row, col).Value = 0
                    ws.Cells(row, col).NumberFormat = num_fmt

                self.supply_row = row

            elif cat == "부가가치세":
                ws.Cells(row, 7).Formula = f"=TRUNC(G{self.supply_row}*10%,0)"
                ws.Cells(row, 9).Formula = f"=TRUNC(I{self.supply_row}*10%,0)"
                ws.Cells(row, 11).Formula = f"=TRUNC(K{self.supply_row}*10%,0)"
                ws.Cells(row, 13).Formula = f"=I{row}+K{row}"
                ws.Cells(row, 15).Formula = f"=G{row}-M{row}"
                for col in [7, 9, 11, 13, 15]: ws.Cells(row, col).NumberFormat = num_fmt
                for col in [6, 8, 10, 12, 14]:
                    ws.Cells(row, col).Value = 0
                    ws.Cells(row, col).NumberFormat = num_fmt

            elif cat == "합        계":
                cols_idx = [7, 9, 11, 13, 15]
                cols_let = ['G', 'I', 'K', 'M', 'O']
                for col, let in zip(cols_idx, cols_let):
                    ws.Cells(row, col).Formula = f"={let}{self.supply_row}+{let}{self.supply_row+1}"
                    ws.Cells(row, col).NumberFormat = num_fmt
                    ws.Cells(row, col).Font.Bold = True
                for col in [6, 8, 10, 12, 14]:
                    ws.Cells(row, col).Value = 0
                    ws.Cells(row, col).NumberFormat = num_fmt
            elif cat == "기타실비 소계":
                if extra_rows:
                    for col, l in zip([7, 9, 11, 13, 15], ['G', 'I', 'K', 'M', 'O']):
                        ws.Cells(row, col).Formula = f"=SUM({l}{extra_rows[0]}:{l}{extra_rows[-1]})"
                        ws.Cells(row, col).NumberFormat = num_fmt
                        ws.Cells(row, col).Font.Bold = True
                else:
                    for col in [7, 9, 11, 13, 15]:
                        ws.Cells(row, col).Value = 0
                        ws.Cells(row, col).NumberFormat = num_fmt
                for col in [6, 8, 10, 12, 14]:
                    ws.Cells(row, col).Value = 0
                    ws.Cells(row, col).NumberFormat = num_fmt

            elif cat in ["장비손료", "안전관리비", "주재비 및 출장여비", "도서인쇄비", "엔지니어링 손해배상공제료"]:
                ws.Cells(row, 7).Value = c_amt
                ws.Cells(row, 9).Value = p_amt
                ws.Cells(row, 11).Value = cur_amt
                ws.Cells(row, 13).Formula = f"=I{row}+K{row}"
                ws.Cells(row, 15).Formula = f"=G{row}-M{row}"
                for col in [7, 9, 11, 13, 15]:
                    ws.Cells(row, col).NumberFormat = num_fmt
                for col in [6, 8, 10, 12, 14]:
                    ws.Cells(row, col).Value = 0
                    ws.Cells(row, col).NumberFormat = num_fmt

            else:
                is_float = (unit == "M")
                fmt_qty = float_fmt if is_float else num_fmt

                ws.Cells(row, 6).Value = round(float(c_qty), 4) if c_qty else 0
                ws.Cells(row, 7).Formula = f"=TRUNC(F{row}*E{row})"

                ws.Cells(row, 8).Value = round(float(p_qty), 4) if p_qty else 0
                ws.Cells(row, 9).Formula = f"=TRUNC(H{row}*E{row})"

                ws.Cells(row, 10).Value = round(float(cur_qty), 4) if cur_qty else 0
                ws.Cells(row, 11).Formula = f"=TRUNC(J{row}*E{row})"

                ws.Cells(row, 12).Formula = f"=H{row}+J{row}"
                ws.Cells(row, 13).Formula = f"=I{row}+K{row}"

                ws.Cells(row, 14).Formula = f"=F{row}-L{row}"
                ws.Cells(row, 15).Formula = f"=G{row}-M{row}"

                for col in [6, 8, 10, 12, 14]: ws.Cells(row, col).NumberFormat = fmt_qty
                for col in [7, 9, 11, 13, 15]: ws.Cells(row, col).NumberFormat = num_fmt

            row += 1

        # 전체 행 높이 일괄 적용 (A4 가로 1장에 위아래 여백 없이 시원하게 꽉 차도록 최적 배정)
        ws.Rows(1).RowHeight = 22 # 제목 윗부분 (A1:O2 병합)
        ws.Rows(2).RowHeight = 22 # 제목 아랫부분
        ws.Rows(3).RowHeight = 8  # 공백
        ws.Rows(4).RowHeight = 24 # 서브 타이틀 (공사명 및 청구기간)
        ws.Rows(5).RowHeight = 8  # 공백
        ws.Rows(6).RowHeight = 25 # 헤더 1 (공종, 규격, 계약 등)
        ws.Rows(7).RowHeight = 25 # 헤더 2 (수량, 금액 등)
        for r in range(8, row):
            ws.Rows(r).RowHeight = 29.5 # 데이터 행 (29.5pt로 확장하여 가로 축소 배율에서도 하단 백지 없이 꽉 채움)

        # --- 세부 내역 테이블 (기성청구 내역서 하단에서 제거됨, 수량 및 금액 계산 로직만 유지) ---
        total_mat = total_lab = total_ovr = total_tech = total_sub = 0

        categories_det = ["PAUT", "RT_B", "RT_A", "RT_A2", "UT", "PT", "MT"]
        for g_type in categories_det:
            if g_type.startswith("RT"):
                if g_type == "RT_B":
                    group_records = [r for r in target_records if r["ndt_type"] == "RT" and "17" in r["material_type"]]
                elif g_type == "RT_A":
                    group_records = [r for r in target_records if r["ndt_type"] == "RT" and "12" in r["material_type"]]
                else:
                    group_records = [r for r in target_records if r["ndt_type"] == "RT" and "6" in r["material_type"]]
            else:
                group_records = [r for r in target_records if r["ndt_type"] == g_type]

            if not group_records: continue

            sub_mat = sub_lab = sub_ovr = sub_tech = sub_sub = 0

            aggregated = {}
            for r in group_records:
                key = (r["loc"], r["ndt_type"], r["material_type"], r["work_time"], r["unit"], r["corr"])
                if key not in aggregated:
                    aggregated[key] = {
                        "date_list": [], "qty": 0.0, "adjusted_qty": 0.0,
                        "mat_cost": 0, "lab_cost": 0, "overhead": 0, "tech": 0, "subtotal": 0
                    }
                aggregated[key]["date_list"].append(r["date"])
                aggregated[key]["qty"] += r["qty"]
                aggregated[key]["adjusted_qty"] += r["adjusted_qty"]
                aggregated[key]["mat_cost"] += r["mat_cost"]
                aggregated[key]["lab_cost"] += r["lab_cost"]
                aggregated[key]["overhead"] += r["overhead"]
                aggregated[key]["tech"] += r["tech"]
                aggregated[key]["subtotal"] += r["subtotal"]

            for key, data in aggregated.items():
                loc_val = "플랜트(관리소)" if "관리소" in key[0] or "플랜트" in key[0] else "열배관"
                t_time = key[3]
                mat_raw = key[2]
                ndt_type = key[1]

                mat = ""
                if ndt_type == "RT":
                    if "17" in mat_raw: mat = "RT_B"
                    elif "12" in mat_raw: mat = "RT_A"
                    elif "6" in mat_raw: mat = "RT_A2"
                else:
                    mat = ndt_type

                cat_key = f"{loc_val}_{t_time}_{mat}"
                c_price = self.contract_vars.get(cat_key, {}).get("c_price", 0)

                if c_price > 0:
                    exact_subtotal = int(data["qty"] * c_price)
                    data["subtotal"] = exact_subtotal

                sub_mat += data["mat_cost"]; sub_lab += data["lab_cost"]
                sub_ovr += data["overhead"]; sub_tech += data["tech"]
                sub_sub += data["subtotal"]

            total_mat += sub_mat; total_lab += sub_lab; total_ovr += sub_ovr
            total_tech += sub_tech; total_sub += sub_sub

        # (세부 내역에는 실비 정산 및 부가세 항목 생략 - 갑지 및 기성내역서에만 포함)

        # --- 내역서 페이지 여백 및 A4 1장 맞춤 설정 ---
        try:
            excel.PrintCommunication = False
        except Exception:
            pass
        ws.PageSetup.PrintArea = f"$A$1:$O${row - 1}"
        ws.PageSetup.PaperSize = 9  # xlPaperA4
        ws.PageSetup.Orientation = 2 # xlLandscape
        ws.PageSetup.Zoom = False
        ws.PageSetup.FitToPagesWide = 1
        ws.PageSetup.FitToPagesTall = 1
        ws.PageSetup.LeftMargin = 20
        ws.PageSetup.RightMargin = 20
        ws.PageSetup.TopMargin = 20
        ws.PageSetup.BottomMargin = 20
        ws.PageSetup.CenterHorizontally = True # 페이지 가로 가운데 정렬
        ws.PageSetup.CenterVertically = True
        try:
            excel.PrintCommunication = True
        except Exception:
            pass

        # --- 열 너비 최적화: 공종 글자 잘림 방지 및 A4 가로 꽉 찬 배율 확보 ---
        ws.Columns(1).ColumnWidth = 27 # 공종 ("위상배열초음파검사(PAUT)" 완벽 표시)
        ws.Columns(1).ShrinkToFit = True
        ws.Columns(2).ColumnWidth = 6.5 # 규격 (병합 좌측)
        ws.Columns(3).ColumnWidth = 11  # 규격 (병합 우측)
        ws.Columns(4).ColumnWidth = 5   # 단위
        ws.Columns(5).ColumnWidth = 11  # 단가

        # 수량 열(F, H, J, L, N)과 금액 열(G, I, K, M, O)
        for c in [6, 8, 10, 12, 14]:
            ws.Columns(c).ColumnWidth = 8.8 # 수량
        for c in [7, 9, 11, 13, 15]:
            ws.Columns(c).ColumnWidth = 12.0 # 금액

        billing_last_row = row - 1

        freeze_header(ws, 7)

        # 셀 내부 글자 잘림 원천 차단 (셀 크기 맞춤)
        ws.Range(f"A6:O{row-1}").ShrinkToFit = True

        # 표 외곽 테두리선 굵게 적용 (청구서 갑지와 동일한 완성도)
        tbl_box = ws.Range(f"A6:O{row-1}")
        for edge in (7, 8, 9, 10): # xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight
            tbl_box.Borders(edge).LineStyle = 1
            tbl_box.Borders(edge).Weight = 3
        # --- 표지 (청구서 갑지) 생성 ---
        record_stage("기성청구내역서 작성")
        ws_cover = wb.Sheets.Add(ws)
        ws_cover.Name = "청구서(갑지)"

        # --- 페이지 가로 모드 및 여백 설정 ---
        try:
            excel.PrintCommunication = False
        except Exception:
            pass
        ws_cover.PageSetup.PaperSize = 9  # xlPaperA4
        ws_cover.PageSetup.Orientation = 2 # xlLandscape
        ws_cover.PageSetup.LeftMargin = 20
        ws_cover.PageSetup.RightMargin = 20
        ws_cover.PageSetup.TopMargin = 20
        ws_cover.PageSetup.BottomMargin = 20
        ws_cover.PageSetup.CenterHorizontally = True
        ws_cover.PageSetup.CenterVertically = True
        ws_cover.PageSetup.Zoom = False
        ws_cover.PageSetup.FitToPagesWide = 1
        ws_cover.PageSetup.FitToPagesTall = 1
        try:
            excel.PrintCommunication = True
        except Exception:
            pass

        # --- 제목 ---
        ws_cover.Range("A2:D5").Merge()
        ws_cover.Range("A2").Value = "청 구 서"
        ws_cover.Range("A2").Font.Size = 36
        ws_cover.Range("A2").Font.Bold = True
        ws_cover.Range("A2").HorizontalAlignment = -4108
        ws_cover.Range("A2").VerticalAlignment = -4108

        # --- 결재란 (우측 상단) ---
        ws_cover.Range("E2").Value = "담 당"
        ws_cover.Range("F2").Value = "검 토"
        ws_cover.Range("G2").Value = "승 인"

        for col_name in ["E", "F", "G"]:
            cell = ws_cover.Range(f"{col_name}2")
            cell.HorizontalAlignment = -4108
            cell.VerticalAlignment = -4108
            cell.Interior.Color = 15132390
            cell.Borders.LineStyle = 1
            cell.Font.Bold = True

            sig_range = ws_cover.Range(f"{col_name}3:{col_name}5")
            sig_range.Merge()
            sig_range.Borders.LineStyle = 1

        ws_cover.Range("A7:B7").Merge()
        ws_cover.Range("A7").Value = "건 명 :"
        ws_cover.Range("A7").Font.Size = 18
        ws_cover.Range("A7").Font.Bold = True
        ws_cover.Range("A7").HorizontalAlignment = -4152 # xlRight

        ws_cover.Range("C7:G7").Merge()
        ws_cover.Range("C7").Value = f"제{round_val}회 비파괴검사기술용역 기성청구"
        ws_cover.Range("C7").Font.Size = 18
        ws_cover.Range("C7").Font.Bold = True

        ws_cover.Range("A9:B9").Merge()
        ws_cover.Range("A9").Value = "청구금액 :"
        ws_cover.Range("A9").Font.Size = 20
        ws_cover.Range("A9").Font.Bold = True
        ws_cover.Range("A9").HorizontalAlignment = -4152 # xlRight

        ws_cover.Range("C9:G9").Merge()

        # [FIX] 내역서의 '단수 조정(Fraction Adjustment)' 로직과 동일하게 total_sub 재계산
        total_sub = grouped_supply_amount(target_records, self.contract_vars)

        grand_total = total_sub + extra_items_total
        vat = int(grand_total * 0.1)
        grand_total_with_vat = grand_total + vat
        ws_cover.Range("C9").Value = f"일금 {grand_total_with_vat:,}원정 (VAT포함)"
        ws_cover.Range("C9").Font.Bold = True
        ws_cover.Range("C9").Font.Size = 20

        # --- 공사명 (row 11) ---
        def _cover_row(row, label, value, label_size=16, val_size=16):
            ws_cover.Range(f"A{row}:B{row}").Merge()
            ws_cover.Range(f"A{row}").Value = label
            ws_cover.Range(f"A{row}").Font.Size = label_size
            ws_cover.Range(f"A{row}").Font.Bold = True
            ws_cover.Range(f"A{row}").HorizontalAlignment = -4108 # xlCenter
            ws_cover.Range(f"A{row}").VerticalAlignment = -4108
            ws_cover.Range(f"A{row}").Interior.Color = 15132390
            ws_cover.Range(f"C{row}:G{row}").Merge()
            ws_cover.Range(f"C{row}").Value = " " + value # Add slight padding
            ws_cover.Range(f"C{row}").Font.Size = val_size
            ws_cover.Range(f"C{row}").HorizontalAlignment = -4131 # xlLeft
            ws_cover.Range(f"C{row}").VerticalAlignment = -4108
            for c_idx in range(1, 8):
                ws_cover.Cells(row, c_idx).Borders.LineStyle = 1
            ws_cover.Rows(row).RowHeight = 30

        _cover_row(11, "공 사 명", "2026년 중앙지사 열수송관 비파괴검사용역 단가계약")
        _cover_row(12, "계 약 기 간", "2026.08.05 ~ 2027.08.05")

        # --- 5단계 자금 흐름 표 ---
        c_sup = self.get_int(self.total_contract_var)
        c_vat = int(c_sup * 0.1)
        c_tot = c_sup + c_vat

        p_sup = self.get_int(self.total_prev_var)
        p_vat = int(p_sup * 0.1)
        p_tot = p_sup + p_vat

        cur_sup = grand_total
        cur_vat = vat
        cur_tot = grand_total_with_vat

        cum_sup = p_sup + cur_sup
        cum_vat = p_vat + cur_vat
        cum_tot = p_tot + cur_tot

        rem_sup = c_sup - cum_sup
        rem_vat = c_vat - cum_vat
        rem_tot = c_tot - cum_tot

        table_start = 14
        table_headers = ["계약금액", "전회누계", "금회청구", "총누계", "잔여금액"]

        ws_cover.Range(f"A{table_start}:B{table_start}").Merge()
        ws_cover.Range(f"A{table_start}").Value = "구 분"
        ws_cover.Range(f"C{table_start}:G{table_start}").Value = [table_headers]
        cell = ws_cover.Range(f"A{table_start}:G{table_start}")
        cell.Font.Bold = True
        cell.Font.Size = 14
        cell.HorizontalAlignment = -4108
        cell.VerticalAlignment = -4108
        cell.Interior.Color = 10066329
        cell.Font.Color = 16777215
        cell.Borders.LineStyle = 1
        ws_cover.Rows(table_start).RowHeight = 28

        rows_data = [
            ("공급가액", [c_sup, p_sup, cur_sup, cum_sup, rem_sup]),
            ("부가가치세", [c_vat, p_vat, cur_vat, cum_vat, rem_vat]),
            ("합     계", [c_tot, p_tot, cur_tot, cum_tot, rem_tot])
        ]

        for r_idx, (label, vals) in enumerate(rows_data):
            row_idx = table_start + 1 + r_idx
            ws_cover.Range(f"A{row_idx}:B{row_idx}").Merge()
            cell_lbl = ws_cover.Range(f"A{row_idx}")
            cell_lbl.Value = label
            cell_lbl.Font.Bold = True
            cell_lbl.Font.Size = 14
            cell_lbl.HorizontalAlignment = -4108
            cell_lbl.VerticalAlignment = -4108
            cell_lbl.Interior.Color = 15132390
            for col_i in [1, 2]: ws_cover.Cells(row_idx, col_i).Borders.LineStyle = 1

            value_range = ws_cover.Range(f"C{row_idx}:G{row_idx}")
            value_range.Value = [[f"{v:,}" for v in vals]]
            value_range.Font.Size = 14
            value_range.HorizontalAlignment = -4152
            value_range.VerticalAlignment = -4108
            value_range.Borders.LineStyle = 1
            ws_cover.Rows(row_idx).RowHeight = 28

        ws_cover.Range("A21:G21").Merge()
        ws_cover.Range("A21").Value = "위와 같이 기성대금을 청구합니다."
        ws_cover.Range("A21").Font.Size = 16
        ws_cover.Range("A21").HorizontalAlignment = -4108

        ws_cover.Range("A26:G26").Merge()
        ws_cover.Range("A26").Value = f"{datetime.now().strftime('%Y년 %m월 %d일')}"
        ws_cover.Range("A26").Font.Size = 16
        ws_cover.Range("A26").HorizontalAlignment = -4108

        ws_cover.Range("A29:G29").Merge()
        ws_cover.Range("A29").Value = "청구인 : 서울검사(주) (인)"
        ws_cover.Range("A29").Font.Size = 18
        ws_cover.Range("A29").Font.Bold = True
        ws_cover.Range("A29").HorizontalAlignment = -4108 # xlCenter

        ws_cover.Range("A32:G32").Merge()
        ws_cover.Range("A32").Value = "한국지역난방공사 중앙지사 귀하"
        ws_cover.Range("A32").Font.Size = 22
        ws_cover.Range("A32").Font.Bold = True
        ws_cover.Range("A32").HorizontalAlignment = -4108 # xlCenter

        ws_cover.Columns(1).ColumnWidth = 20
        ws_cover.Columns(2).ColumnWidth = 20
        ws_cover.Columns(3).ColumnWidth = 26
        ws_cover.Columns(4).ColumnWidth = 26
        ws_cover.Columns(5).ColumnWidth = 26
        ws_cover.Columns(6).ColumnWidth = 26
        ws_cover.Columns(7).ColumnWidth = 26

        # --- 전체 외곽선 (A1 ~ G33) 굵게 설정 ---
        outer_range = ws_cover.Range("A1:G33")
        # Match the cover's proportions at the same printed width.
        try:
            detail_width = ws.Range(f"A1:O{billing_last_row}").Width
            cover_ratio = outer_range.Height / outer_range.Width
            a4_print_ratio = (595 - 40) / (842 - 40)
            target_height = detail_width * min(cover_ratio, a4_print_ratio * 0.97)
            header_height = ws.Range("A1:O7").Height
            data_rows = billing_last_row - 7
            if data_rows > 0:
                data_height = max(14.0, min(40.0, (target_height - header_height) / data_rows))
                ws.Rows(f"8:{billing_last_row}").RowHeight = data_height
        except com_error as error:
            print(f"내역서 행 높이 조정 실패: {error}")
        for edge in (7, 8, 9, 10): # xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight
            outer_range.Borders(edge).LineStyle = 1
            outer_range.Borders(edge).Weight = 4 # xlThick

        # --- 업체별 기성요약 시트 생성 ---
        record_stage("청구서 표지 작성")
        ws_summary = wb.Sheets.Add(None, ws)
        ws_summary.Name = "업체별 기성요약"

        ws_summary.Range("A1:G2").Merge()
        ws_summary.Range("A1").Value = f"제 {round_val} 회 비파괴검사기술용역 업체별 기성요약"
        ws_summary.Range("A1").Font.Size = 16
        ws_summary.Range("A1").Font.Bold = True
        ws_summary.Range("A1").HorizontalAlignment = -4108
        ws_summary.Range("A1").VerticalAlignment = -4108

        # --- 실제 시공업체 매핑 ---
        # target_records의 'company'가 '한국지역난방공사' 등으로 일괄 지정되어 있으므로,
        # daily_work_history.json의 ndt_results에서 해당 일자/검사방법의 실제 업체를 찾아 매핑합니다.
        from site_apps.central.src.rt_quantities import corrected_rt_history, rt_counts
        history_path = get_history_path()
        history_data = {}
        history_data_corr = {}
        history_warning = None
        if target_records:
            try:
                with open(history_path, 'r', encoding='utf-8') as f:
                    history_data = json.load(f)
                if not isinstance(history_data, dict) or any(
                    not isinstance(day, dict) for day in history_data.values()
                ):
                    raise ValueError("작업일보 이력 형식이 올바르지 않습니다.")
                history_data_corr = corrected_rt_history(history_data)
            except (OSError, ValueError, TypeError, AttributeError, ArithmeticError) as error:
                history_data = {}
                history_warning = (
                    "작업일보 이력을 읽지 못해 업체별 요약의 섹션·라인번호를 "
                    "연결하지 못했습니다. 청구 금액은 기성 기록을 기준으로 저장했습니다.\n"
                    f"{error}"
                )

        from site_apps.central.src.services.billing_summary import build_company_summaries
        company_summaries = build_company_summaries(target_records, history_data, self.contract_vars)

        # One common header remains visible while company bands scroll.
        headers_sum_list = ["섹션", "라인번호", "규격", "개소", "단가", "검사수량", "금액"]
        hdr_sum_rng = ws_summary.Range("A4:G4")
        hdr_sum_rng.Value = [headers_sum_list]
        hdr_sum_rng.Font.Bold = True
        hdr_sum_rng.Interior.Color = 14277081
        hdr_sum_rng.HorizontalAlignment = -4108
        hdr_sum_rng.Borders.LineStyle = 1
        sum_row = 5
        contract_category_order = [
            cat for cat in categories if cat in self.contract_vars
        ]
        contract_item_numbers = {
            cat: idx for idx, cat in enumerate(contract_category_order, start=1)
        }
        company_subtotal_rows = []
        companies = sorted(company_summaries) or ["미지정"]

        for comp in companies:
            ws_summary.Cells(sum_row, 1).Value = f"■ 업체명 : {comp}"
            ws_summary.Cells(sum_row, 1).Font.Bold = True
            ws_summary.Cells(sum_row, 1).Font.Size = 12
            sum_row += 1

            work_summary = company_summaries.get(comp, {})

            start_data_row = sum_row

            # [FIX] 정렬: 규격(검사방법) -> 섹션 -> 라인번호 (검사 방법별로 그룹화)
            sorted_keys = sorted(work_summary.keys(), key=lambda x: (x[2], x[0], x[1]))

            from itertools import groupby
            for unit_val, keys_group in groupby(sorted_keys, key=lambda x: x[3]):
                keys_group = list(keys_group)

                start_data_row = sum_row
                summary_rows_batch = []
                for key in keys_group:
                    data = work_summary[key]
                    sec, l_no, spec, unit, c_price = key
                    calculated_amt = data.get("calculated_amt", 0)

                    if data["qty"] == 0 and calculated_amt == 0:
                        continue
                    qty_cell = round(data["qty"], 4) if unit == "M" else int(data["qty"])
                    places_cell = int(data["places"]) if data["places"] is not None else ""
                    summary_rows_batch.append((
                        sec,
                        l_no,
                        f"{spec} ({unit})",
                        places_cell,
                        int(c_price),
                        qty_cell,
                        calculated_amt,
                    ))

                if not summary_rows_batch:
                    continue

                end_data_row = start_data_row + len(summary_rows_batch) - 1
                sum_data_rng = ws_summary.Range(
                    ws_summary.Cells(start_data_row, 1),
                    ws_summary.Cells(end_data_row, 7),
                )
                sum_data_rng.Value = summary_rows_batch
                sum_data_rng.Borders.LineStyle = 1
                ws_summary.Range(ws_summary.Cells(start_data_row, 1), ws_summary.Cells(end_data_row, 2)).HorizontalAlignment = -4108
                ws_summary.Range(ws_summary.Cells(start_data_row, 3), ws_summary.Cells(end_data_row, 3)).HorizontalAlignment = -4131
                ws_summary.Range(ws_summary.Cells(start_data_row, 4), ws_summary.Cells(end_data_row, 4)).HorizontalAlignment = -4108
                ws_summary.Range(ws_summary.Cells(start_data_row, 5), ws_summary.Cells(end_data_row, 5)).NumberFormat = "#,##0"
                ws_summary.Range(ws_summary.Cells(start_data_row, 6), ws_summary.Cells(end_data_row, 6)).NumberFormat = (
                    '#,##0.0000;-#,##0.0000;"-"' if unit_val == "M" else '#,##0;-#,##0;"-"'
                )
                ws_summary.Range(ws_summary.Cells(start_data_row, 7), ws_summary.Cells(end_data_row, 7)).NumberFormat = '#,##0;-#,##0;"-"'
                try:
                    ws_summary.Range(ws_summary.Cells(start_data_row, 2), ws_summary.Cells(end_data_row, 2)).ShrinkToFit = True
                except Exception:
                    pass
                sum_row = end_data_row + 1

                # 소계 렌더링
                sub_rng = ws_summary.Range(ws_summary.Cells(sum_row, 1), ws_summary.Cells(sum_row, 7))
                ws_summary.Range(ws_summary.Cells(sum_row, 1), ws_summary.Cells(sum_row, 5)).Merge()
                ws_summary.Cells(sum_row, 1).Value = f"소 계 ({unit_val})"
                ws_summary.Cells(sum_row, 1).HorizontalAlignment = -4108
                sub_rng.Font.Bold = True
                sub_rng.Interior.Color = 15987699
                sub_rng.Borders.LineStyle = 1

                qty_sum_formula = f"=SUM(F{start_data_row}:F{sum_row-1})" if sum_row > start_data_row else "0"
                ws_summary.Cells(sum_row, 6).Formula = qty_sum_formula
                ws_summary.Cells(sum_row, 6).NumberFormat = (
                    '#,##0.0000;-#,##0.0000;"-"' if unit_val == "M" else '#,##0;-#,##0;"-"'
                )

                sum_formula = f"=SUM(G{start_data_row}:G{sum_row-1})" if sum_row > start_data_row else "0"
                ws_summary.Cells(sum_row, 7).Formula = sum_formula
                ws_summary.Cells(sum_row, 7).NumberFormat = '#,##0;-#,##0;"-"'

                company_subtotal_rows.append(sum_row)
                sum_row += 1

            sum_row += 2

        # --- 전체 공급가액 / 부가가치세 / 합계 ---
        supply_row = sum_row
        vat_row = sum_row + 1
        grand_total_row = sum_row + 2

        supply_formula = (
            "=" + "+".join(f"G{row_no}" for row_no in company_subtotal_rows)
            if company_subtotal_rows else "=0"
        )
        total_rows = [
            (supply_row, "공 급 가 액", supply_formula),
            (vat_row, "부가가치세", f"=TRUNC(G{supply_row}*10%,0)"),
            (grand_total_row, "합        계", f"=G{supply_row}+G{vat_row}"),
        ]
        for total_summary_row, label, formula in total_rows:
            ws_summary.Range(
                ws_summary.Cells(total_summary_row, 1),
                ws_summary.Cells(total_summary_row, 6)
            ).Merge()
            ws_summary.Cells(total_summary_row, 1).Value = label
            ws_summary.Cells(total_summary_row, 1).HorizontalAlignment = -4108
            ws_summary.Cells(total_summary_row, 1).Font.Bold = True
            ws_summary.Cells(total_summary_row, 7).Formula = formula
            ws_summary.Cells(total_summary_row, 7).NumberFormat = '#,##0;-#,##0;"-"'
            ws_summary.Cells(total_summary_row, 7).Font.Bold = True
            for col in range(1, 8):
                ws_summary.Cells(total_summary_row, col).Borders.LineStyle = 1
            ws_summary.Rows(total_summary_row).RowHeight = 26

        # VAT 포함 최종 합계 행을 기존 총합계 강조 형식으로 표시한다.
        for col in range(1, 8):
            ws_summary.Cells(grand_total_row, col).Interior.Color = 10066329
            ws_summary.Cells(grand_total_row, col).Font.Color = 16777215
            ws_summary.Cells(grand_total_row, col).Font.Size = 13
        ws_summary.Rows(grand_total_row).RowHeight = 28

        ws_summary.Columns(1).ColumnWidth = 14
        ws_summary.Columns(2).ColumnWidth = 40
        ws_summary.Columns(3).ColumnWidth = 44
        ws_summary.Columns(4).ColumnWidth = 8
        ws_summary.Columns(5).ColumnWidth = 12
        ws_summary.Columns(6).ColumnWidth = 12
        ws_summary.Columns(7).ColumnWidth = 15
        try:
            ws_summary.Columns(2).ShrinkToFit = True
        except Exception:
            pass
        try:
            excel.PrintCommunication = False
        except Exception:
            pass
        ws_summary.PageSetup.Orientation = 2 # xlLandscape
        ws_summary.PageSetup.Zoom = False
        ws_summary.PageSetup.FitToPagesWide = 1
        ws_summary.PageSetup.FitToPagesTall = False
        ws_summary.PageSetup.LeftMargin = 20
        ws_summary.PageSetup.RightMargin = 20
        ws_summary.PageSetup.TopMargin = 20
        ws_summary.PageSetup.BottomMargin = 20
        ws_summary.PageSetup.CenterHorizontally = True
        ws_summary.PageSetup.RightHeader = "\n\nPage &P of &N&KFFFFFF" + " "*10 + "."
        ws_summary.PageSetup.PrintTitleRows = "$1:$4"
        try:
            excel.PrintCommunication = True
        except Exception:
            pass
        freeze_header(ws_summary, 4)

        # --- 업체별 수량내역 시트 생성 ---
        record_stage("업체별 기성요약 작성")
        ws_cont = wb.Sheets.Add(None, ws_summary)
        ws_cont.Name = "업체별 수량내역"

        ws_cont.Range("A1:K2").Merge()
        ws_cont.Range("A1").Value = f"제 {round_val} 회 기성청구 업체별 수량내역"
        ws_cont.Range("A1").Font.Size = 16
        ws_cont.Range("A1").Font.Bold = True
        ws_cont.Range("A1").HorizontalAlignment = -4108
        ws_cont.Range("A1").VerticalAlignment = -4108

        headers_cont = ["No.", "업체명", "검사방법", "구간", "라인번호", "Joint No.", "관경", "두께", "용접사", "수량", "결과"]
        hdr_cont_rng = ws_cont.Range("A4:K4")
        hdr_cont_rng.Value = [headers_cont]
        hdr_cont_rng.Font.Bold = True
        hdr_cont_rng.Interior.Color = 14277081
        hdr_cont_rng.HorizontalAlignment = -4108
        hdr_cont_rng.Borders.LineStyle = 1

        ws_cont.Columns(1).ColumnWidth = 8
        ws_cont.Columns(2).ColumnWidth = 20
        ws_cont.Columns(3).ColumnWidth = 10
        ws_cont.Columns(4).ColumnWidth = 15
        ws_cont.Columns(5).ColumnWidth = 38
        ws_cont.Columns(6).ColumnWidth = 10
        ws_cont.Columns(7).ColumnWidth = 10
        ws_cont.Columns(8).ColumnWidth = 10
        ws_cont.Columns(9).ColumnWidth = 12
        ws_cont.Columns(10).ColumnWidth = 12
        ws_cont.Columns(11).ColumnWidth = 10
        try:
            ws_cont.Columns(5).ShrinkToFit = True
        except Exception:
            pass

        cont_row = 5

        target_dates = set(r["date"] for r in target_records)
        all_ndt_results = []
        source_rows = {}
        legacy_dates = set()
        for record in target_records:
            if record.get('source') == 'work_log':
                source_rows.setdefault(record['date'], set()).update(record.get('history_rows', []))
            else:
                legacy_dates.add(record['date'])
        for t_date in target_dates:
            if t_date in history_data_corr and "ndt_results" in history_data_corr[t_date]:
                all_ndt_results.extend(
                    detail for index, detail in enumerate(history_data_corr[t_date]['ndt_results'])
                    if t_date in legacy_dates or index in source_rows.get(t_date, set())
                )

        all_ndt_results.sort(key=lambda x: (str(x.get("업체", "")), str(x.get("검사방법", "")), str(x.get("구간", "")), str(x.get("라인번호", ""))))

        total_points = 0
        total_meters = 0.0

        from itertools import groupby

        idx_cont = 1
        for (comp, method), group_iter in groupby(all_ndt_results, key=lambda x: (str(x.get("업체", "")).strip(), str(x.get("검사방법", "")).strip())):
            if not comp: continue
            group = list(group_iter)

            grp_hdr_rng = ws_cont.Range(ws_cont.Cells(cont_row, 1), ws_cont.Cells(cont_row, 11))
            ws_cont.Cells(cont_row, 1).Value = f"■ {comp} - {method}"
            grp_hdr_rng.Merge()
            ws_cont.Cells(cont_row, 1).Font.Bold = True
            ws_cont.Cells(cont_row, 1).Interior.Color = 15132390
            ws_cont.Cells(cont_row, 1).HorizontalAlignment = -4108
            grp_hdr_rng.Borders.LineStyle = 1
            cont_row += 1

            start_group_row = cont_row
            is_length_inspection = any(
                name in method for name in ("PAUT", "UT", "MT", "PT")
            )
            unit_str = "M" if is_length_inspection else "매"
            format_str = '#,##0.0000;-#,##0.0000;"-"' if unit_str == "M" else '#,##0;-#,##0;"-"'

            cont_rows_batch = []
            for r in group:
                if not str(r.get("업체", "")).strip() and not str(r.get("Joint No.", "")).strip():
                    continue

                if "PAUT" in method:
                    qty = r.get("PAUT", "")
                elif "RT" in method:
                    qty = rt_counts(r)[0]
                elif "MT" in method:
                    qty = r.get("MT", "")
                elif "PT" in method:
                    qty = r.get("PT", "")
                else:
                    qty = r.get(method, "")

                try:
                    q_val = float(qty) if str(qty).strip() else 0.0
                except ValueError:
                    q_val = 0.0

                if is_length_inspection:
                    total_meters += q_val
                else:
                    total_points += q_val

                # [FIX] 수량을 숫자로 저장해야 SUBTOTAL 수식이 작동함
                q_cell = q_val if q_val != 0.0 else (qty if str(qty).strip() else "")
                cont_rows_batch.append((
                    idx_cont,
                    comp,
                    method,
                    r.get("구간", ""),
                    r.get("라인번호", ""),
                    r.get("Joint No.", ""),
                    r.get("관경", ""),
                    r.get("두께", ""),
                    r.get("용접사", ""),
                    q_cell,
                    r.get("결과", ""),
                ))
                idx_cont += 1

            if cont_rows_batch:
                end_group_row = start_group_row + len(cont_rows_batch) - 1
                grp_data_rng = ws_cont.Range(
                    ws_cont.Cells(start_group_row, 1),
                    ws_cont.Cells(end_group_row, 11),
                )
                grp_data_rng.Value = cont_rows_batch
                grp_data_rng.Borders.LineStyle = 1
                grp_data_rng.HorizontalAlignment = -4108
                ws_cont.Range(
                    ws_cont.Cells(start_group_row, 10),
                    ws_cont.Cells(end_group_row, 10),
                ).NumberFormat = format_str
                try:
                    ws_cont.Range(
                        ws_cont.Cells(start_group_row, 5),
                        ws_cont.Cells(end_group_row, 5),
                    ).ShrinkToFit = True
                except Exception:
                    pass
                cont_row = end_group_row + 1
                subtotal_formula = f'=SUBTOTAL(9, J{start_group_row}:J{end_group_row})'
            else:
                subtotal_formula = "0"

            sub_cont_rng = ws_cont.Range(ws_cont.Cells(cont_row, 1), ws_cont.Cells(cont_row, 11))
            ws_cont.Cells(cont_row, 1).Value = f"{comp} {method} 소계"
            ws_cont.Range(ws_cont.Cells(cont_row, 1), ws_cont.Cells(cont_row, 9)).Merge()
            ws_cont.Cells(cont_row, 1).HorizontalAlignment = -4108
            sub_cont_rng.Font.Bold = True
            sub_cont_rng.Interior.Color = 14277081
            sub_cont_rng.Borders.LineStyle = 1

            ws_cont.Cells(cont_row, 10).Formula = subtotal_formula
            ws_cont.Cells(cont_row, 10).NumberFormat = format_str

            ws_cont.Cells(cont_row, 11).Value = unit_str
            ws_cont.Cells(cont_row, 11).HorizontalAlignment = -4108

            cont_row += 1

        if all_ndt_results:
            tot_len_rng = ws_cont.Range(ws_cont.Cells(cont_row, 1), ws_cont.Cells(cont_row, 11))
            ws_cont.Cells(cont_row, 1).Value = "총 누적 물량 (길이 검사)"
            ws_cont.Range(ws_cont.Cells(cont_row, 1), ws_cont.Cells(cont_row, 9)).Merge()
            ws_cont.Cells(cont_row, 1).HorizontalAlignment = -4108
            ws_cont.Cells(cont_row, 1).Interior.Color = 10066329
            ws_cont.Cells(cont_row, 1).Font.Color = 16777215
            tot_len_rng.Font.Bold = True
            tot_len_rng.Borders.LineStyle = 1

            ws_cont.Cells(cont_row, 10).Value = total_meters
            ws_cont.Cells(cont_row, 10).NumberFormat = '#,##0.0000;-#,##0.0000;"-"'

            ws_cont.Cells(cont_row, 11).Value = "M"
            ws_cont.Cells(cont_row, 11).HorizontalAlignment = -4108
            cont_row += 1

            tot_pt_rng = ws_cont.Range(ws_cont.Cells(cont_row, 1), ws_cont.Cells(cont_row, 11))
            ws_cont.Cells(cont_row, 1).Value = "총 누적 물량 (포인트 검사)"
            ws_cont.Range(ws_cont.Cells(cont_row, 1), ws_cont.Cells(cont_row, 9)).Merge()
            ws_cont.Cells(cont_row, 1).HorizontalAlignment = -4108
            ws_cont.Cells(cont_row, 1).Interior.Color = 10066329
            ws_cont.Cells(cont_row, 1).Font.Color = 16777215
            tot_pt_rng.Font.Bold = True
            tot_pt_rng.Borders.LineStyle = 1

            ws_cont.Cells(cont_row, 10).Value = int(total_points)
            ws_cont.Cells(cont_row, 10).NumberFormat = '#,##0;-#,##0;"-"'

            ws_cont.Cells(cont_row, 11).Value = "매/Point"
            ws_cont.Cells(cont_row, 11).HorizontalAlignment = -4108
            cont_row += 1

        # 헤더 자동 필터 및 페이지 설정/틀 고정을 데이터 입력 완료 후 일괄 적용
        try:
            ws_cont.Range("A4:K4").AutoFilter(Field=1)
        except Exception:
            pass
        try:
            excel.PrintCommunication = False
        except Exception:
            pass
        ws_cont.PageSetup.Orientation = 2 # xlLandscape
        ws_cont.PageSetup.Zoom = False
        ws_cont.PageSetup.FitToPagesWide = 1
        ws_cont.PageSetup.FitToPagesTall = False
        ws_cont.PageSetup.LeftMargin = 20
        ws_cont.PageSetup.RightMargin = 20
        ws_cont.PageSetup.TopMargin = 20
        ws_cont.PageSetup.BottomMargin = 20
        ws_cont.PageSetup.CenterHorizontally = True
        ws_cont.PageSetup.PrintTitleRows = "$1:$4"
        ws_cont.PageSetup.RightHeader = "\n\nPage &P of &N&KFFFFFF" + " "*12 + "."
        try:
            excel.PrintCommunication = True
        except Exception:
            pass

        freeze_header(ws_cont, 5)

        ws_cover.Select()
        record_stage("업체별 수량내역 작성")
        try:
            excel.Calculation = -4105  # xlCalculationAutomatic
            excel.Calculate()
            excel.ScreenUpdating = True
            excel.EnableEvents = True
        except Exception:
            pass

        filepath = filepath.replace("/", "\\")
        record_stage("계산 및 저장 준비")
        wb.SaveAs(filepath)
        record_stage("파일 저장")
        wb.Close()
        excel.Quit()
        record_stage("Excel 종료")
        excel = None
        print(f"[기성청구 엑셀] 총 {perf_counter() - export_started:.2f}초 | " + " | ".join(stage_times))
        if progress_win is not None:
            try:
                progress_win.destroy()
                progress_win = None
            except Exception:
                pass

        if history_warning:
            messagebox.showwarning("저장 완료 · 작업일보 연결 실패", f"{history_warning}\n\n{filepath}")
        else:
            messagebox.showinfo("저장 완료", f"엑셀 기성청구 내역서가 성공적으로 생성되었습니다.\n{filepath}")
        os.startfile(filepath)

    except Exception as e:
        if progress_win is not None:
            try:
                progress_win.destroy()
                progress_win = None
            except Exception:
                pass
        messagebox.showerror("저장 오류", f"엑셀 파일 생성 중 오류가 발생했습니다.\n{str(e)}")
        if excel is not None:
            try: excel.Quit()
            except: pass
    finally:
        if progress_win is not None:
            try:
                progress_win.destroy()
            except Exception:
                pass
