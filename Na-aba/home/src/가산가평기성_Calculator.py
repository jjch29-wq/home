import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import os
import win32com.client as win32

# --- 상수 및 단가 정의 ---
MATERIAL_COST = {
    'RT (B필름: 3⅓"x17")': 3379,
    'RT (A필름: 3⅓"x12")': 2540,
    'RT (A/2필름: 3⅓"x6")': 1515,
    "UT": 1115,
    "PT": 3971
}

# 엑셀 기준 보정물량(매/M)당 인건비 단가 (단위: 원)
LABOR_COST = {
    "일반": {"RT": 34863, "UT": 25000, "PT": 20000},
    "야간": {"RT": 49240, "UT": 37500, "PT": 30000},
    "휴일": {"RT": 49313, "UT": 37500, "PT": 30000}
}

class NDTCalculator(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("비파괴검사 기성 산출 계산기 (가산~가평)")
        self.geometry("1100x950")  
        self.configure(padx=20, pady=20)
        self.style = ttk.Style(self)
        self.style.theme_use("clam")
        
        self.records = [] # 저장된 기록 목록
        
        self.create_widgets()
        
    def create_widgets(self):
        # 1. 상단 프레임 (기본 정보)
        info_frame = ttk.Frame(self)
        info_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(info_frame, text="• 검사일자:", font=("Arial", 10, "bold")).pack(side=tk.LEFT)
        self.date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        ttk.Entry(info_frame, textvariable=self.date_var, width=15).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(info_frame, text="• 작업구간 (Joint No 등):", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=(20, 5))
        self.loc_var = tk.StringVar(value="")
        ttk.Entry(info_frame, textvariable=self.loc_var, width=30).pack(side=tk.LEFT, padx=5)

        # 1. 검사 종류
        ttk.Label(self, text="1. 검사 종류", font=("Arial", 11, "bold")).pack(anchor=tk.W, pady=(10, 5))
        self.ndt_type_var = tk.StringVar(value="RT")
        type_frame = ttk.Frame(self)
        type_frame.pack(fill=tk.X, pady=5)
        for t in ["RT", "UT", "PT"]:
            ttk.Radiobutton(type_frame, text=t, value=t, variable=self.ndt_type_var, command=self.update_dynamic_ui).pack(side=tk.LEFT, padx=10)
            
        # 2. 작업 형태
        ttk.Label(self, text="2. 작업 형태 (근무 시간대)", font=("Arial", 11, "bold")).pack(anchor=tk.W, pady=(10, 5))
        self.work_time_var = tk.StringVar(value="일반")
        time_frame = ttk.Frame(self)
        time_frame.pack(fill=tk.X, pady=5)
        for t in ["일반", "야간", "휴일"]:
            ttk.Radiobutton(time_frame, text=t, value=t, variable=self.work_time_var).pack(side=tk.LEFT, padx=10)

        # 3. 사용 자재
        self.material_lbl = ttk.Label(self, text="3. 사용 자재 (RT 필름 규격)", font=("Arial", 11, "bold"))
        self.material_lbl.pack(anchor=tk.W, pady=(10, 5))
        self.material_var = tk.StringVar(value='RT (B필름: 3⅓"x17")')
        self.material_combo = ttk.Combobox(self, textvariable=self.material_var, values=['RT (B필름: 3⅓"x17")', 'RT (A필름: 3⅓"x12")', 'RT (A/2필름: 3⅓"x6")'], state="readonly")
        self.material_combo.pack(fill=tk.X, pady=5)
        
        # --- 동적 UI 영역 ---
        self.dynamic_frame = ttk.LabelFrame(self, text="4. 보정계수 조건 선택", padding=10)
        self.dynamic_frame.pack(fill=tk.X, pady=(10, 5))
        
        # 4-1. 방사선원 (RT)
        self.source_frame = ttk.Frame(self.dynamic_frame)
        ttk.Label(self.source_frame, text="• 방사선원 :", width=15).pack(side=tk.LEFT)
        self.source_var = tk.StringVar(value="Ir-192 또는 Se-75 (1.0)")
        self.source_combo = ttk.Combobox(self.source_frame, textvariable=self.source_var, state="readonly", width=35)
        self.source_combo['values'] = ["Ir-192 또는 Se-75 (1.0)", "X-ray 발생장치 (1.3)"]
        self.source_combo.pack(side=tk.LEFT)
        
        # 4-2. 관경 (UT, PT)
        self.pipe_frame = ttk.Frame(self.dynamic_frame)
        ttk.Label(self.pipe_frame, text="• 관경(구경) :", width=15).pack(side=tk.LEFT)
        self.pipe_var = tk.StringVar()
        self.pipe_combo = ttk.Combobox(self.pipe_frame, textvariable=self.pipe_var, state="readonly", width=35)
        self.pipe_combo.pack(side=tk.LEFT)
        
        # 4-3. 두께 (RT, UT)
        self.thickness_frame = ttk.Frame(self.dynamic_frame)
        ttk.Label(self.thickness_frame, text="• 투과/모재두께 :", width=15).pack(side=tk.LEFT)
        self.thickness_var = tk.StringVar()
        self.thickness_combo = ttk.Combobox(self.thickness_frame, textvariable=self.thickness_var, state="readonly", width=35)
        self.thickness_combo.pack(side=tk.LEFT)
        
        # 5. 검사 물량
        ttk.Label(self, text="5. 실검사 물량 (RT: 매 / UT,PT: Meter)", font=("Arial", 11, "bold")).pack(anchor=tk.W, pady=(10, 5))
        self.quantity_var = tk.DoubleVar(value=10.0)
        ttk.Entry(self, textvariable=self.quantity_var).pack(fill=tk.X, pady=5)
        
        # 6. 적용 요율 (제경비, 기술료)
        rate_frame = ttk.Frame(self)
        rate_frame.pack(fill=tk.X, pady=(15, 5))
        ttk.Label(rate_frame, text="6. 적용 요율 (%)", font=("Arial", 11, "bold")).pack(side=tk.LEFT, padx=(0, 20))
        
        ttk.Label(rate_frame, text="제경비율:").pack(side=tk.LEFT)
        self.overhead_rate_var = tk.DoubleVar(value=80.0)
        ttk.Entry(rate_frame, textvariable=self.overhead_rate_var, width=8).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(rate_frame, text="기술료율:").pack(side=tk.LEFT, padx=(15, 0))
        self.tech_fee_rate_var = tk.DoubleVar(value=5.86)
        ttk.Entry(rate_frame, textvariable=self.tech_fee_rate_var, width=8).pack(side=tk.LEFT, padx=5)
        
        # 버튼 영역
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill=tk.X, pady=15)
        
        ttk.Button(btn_frame, text="금액 계산하기", command=self.calculate).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5, ipady=8)
        ttk.Button(btn_frame, text="기록 목록에 추가", command=self.add_to_record).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5, ipady=8)
        
        # 7. 단일 계산 결과 출력 영역
        ttk.Label(self, text="[ 단일 계산 결과 ]", font=("Arial", 11, "bold")).pack(anchor=tk.W, pady=(0, 5))
        self.result_text = tk.Text(self, height=5, width=50, state=tk.DISABLED, font=("Consolas", 11))
        self.result_text.pack(fill=tk.X, expand=False)
        
        # 8. 누적 기록 테이블
        lbl_frame = ttk.Frame(self)
        lbl_frame.pack(fill=tk.X, pady=(15, 5))
        ttk.Label(lbl_frame, text="[ 일일 작업 기록 목록 ]", font=("Arial", 11, "bold")).pack(side=tk.LEFT)
        ttk.Button(lbl_frame, text="엑셀 파일로 저장", command=self.export_to_excel).pack(side=tk.RIGHT)
        ttk.Button(lbl_frame, text="기록 초기화", command=self.clear_records).pack(side=tk.RIGHT, padx=5)

        columns = ("date", "loc", "type", "time", "mat", "qty", "unit", "corr", "adj_qty", "overhead", "tech", "total_amt")
        self.tree = ttk.Treeview(self, columns=columns, show="headings", height=8)
        
        self.tree.heading("date", text="일자")
        self.tree.heading("loc", text="구간/위치")
        self.tree.heading("type", text="종류")
        self.tree.heading("time", text="형태")
        self.tree.heading("mat", text="자재")
        self.tree.heading("qty", text="실물량")
        self.tree.heading("unit", text="단위")
        self.tree.heading("corr", text="보정계수")
        self.tree.heading("adj_qty", text="보정물량")
        self.tree.heading("overhead", text="제경비(원)")
        self.tree.heading("tech", text="기술료(원)")
        self.tree.heading("total_amt", text="공급가액(원)")
        
        self.tree.column("date", width=80, anchor="center")
        self.tree.column("loc", width=120, anchor="w")
        self.tree.column("type", width=40, anchor="center")
        self.tree.column("time", width=40, anchor="center")
        self.tree.column("mat", width=90, anchor="center")
        self.tree.column("qty", width=50, anchor="e")
        self.tree.column("unit", width=40, anchor="center")
        self.tree.column("corr", width=50, anchor="e")
        self.tree.column("adj_qty", width=50, anchor="e")
        self.tree.column("overhead", width=70, anchor="e")
        self.tree.column("tech", width=70, anchor="e")
        self.tree.column("total_amt", width=90, anchor="e")
        
        self.tree.pack(fill=tk.BOTH, expand=True)
        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)
        
        self.update_dynamic_ui()

    def update_dynamic_ui(self, *args):
        ndt_type = self.ndt_type_var.get()
        if ndt_type == "RT":
            self.material_combo.config(values=['RT (B필름: 3⅓"x17")', 'RT (A필름: 3⅓"x12")', 'RT (A/2필름: 3⅓"x6")'], state="readonly")
            if not self.material_var.get().startswith("RT"):
                self.material_var.set('RT (B필름: 3⅓"x17")')
        else:
            self.material_combo.config(values=[ndt_type], state="disabled")
            self.material_var.set(ndt_type)
            
        self.source_frame.pack_forget()
        self.pipe_frame.pack_forget()
        self.thickness_frame.pack_forget()
        
        if ndt_type == "RT":
            self.source_frame.pack(fill=tk.X, pady=2)
            self.thickness_frame.pack(fill=tk.X, pady=2)
            self.thickness_combo['values'] = ["15mm 이하 (1.0)", "15mm 초과 ~ 25mm 이하 (1.4)", "25mm 초과 ~ 40mm 이하 (2.2)"]
            self.thickness_var.set("15mm 이하 (1.0)")
        elif ndt_type == "UT":
            self.pipe_frame.pack(fill=tk.X, pady=2)
            self.thickness_frame.pack(fill=tk.X, pady=2)
            self.pipe_combo['values'] = ["250mm 초과 [10인치 이상] (1.0)", "200~250mm [8인치] (1.2)", "150~200mm [6인치] (1.4)", "100~150mm [4인치] (1.7)", "100mm 이하 [3인치 이하] (2.0)"]
            self.pipe_var.set("250mm 초과 [10인치 이상] (1.0)")
            self.thickness_combo['values'] = ["15mm 이하 (1.0)", "15mm 초과 ~ 50mm 이하 (1.2)"]
            self.thickness_var.set("15mm 이하 (1.0)")
        elif ndt_type == "PT":
            self.pipe_frame.pack(fill=tk.X, pady=2)
            self.pipe_combo['values'] = ["150mm 초과 [6인치 이상] (1.2)", "150mm 이하 [4인치 이하] (1.4)"]
            self.pipe_var.set("150mm 초과 [6인치 이상] (1.2)")

    def get_correction_factor(self):
        ndt_type = self.ndt_type_var.get()
        factor = 1.0
        if ndt_type == "RT":
            if "1.3" in self.source_var.get(): factor *= 1.3
            if "1.4" in self.thickness_var.get(): factor *= 1.4
            elif "2.2" in self.thickness_var.get(): factor *= 2.2
        elif ndt_type == "UT":
            pipe_val = self.pipe_var.get()
            if "1.2" in pipe_val: factor *= 1.2
            elif "1.4" in pipe_val: factor *= 1.4
            elif "1.7" in pipe_val: factor *= 1.7
            elif "2.0" in pipe_val: factor *= 2.0
            if "1.2" in self.thickness_var.get(): factor *= 1.2
        elif ndt_type == "PT":
            if "1.2" in self.pipe_var.get(): factor *= 1.2
            elif "1.4" in self.pipe_var.get(): factor *= 1.4
        return factor

    def _do_calculate(self):
        date_str = self.date_var.get()
        loc_str = self.loc_var.get()
        ndt_type = self.ndt_type_var.get()
        work_time = self.work_time_var.get()
        material_type = self.material_var.get()
        qty = float(self.quantity_var.get())
        
        # 단위 판별
        unit_str = "매" if ndt_type == "RT" else "M"
        
        overhead_rate = float(self.overhead_rate_var.get()) / 100.0
        tech_fee_rate = float(self.tech_fee_rate_var.get()) / 100.0
        
        corr = self.get_correction_factor()
        adjusted_qty = qty * corr
        
        mat_unit_cost = MATERIAL_COST.get(material_type, 0)
        total_mat_cost = int(qty * mat_unit_cost)
        
        lab_unit_cost = LABOR_COST[work_time][ndt_type]
        total_lab_cost = int(adjusted_qty * lab_unit_cost)
        
        overhead_cost = int(total_lab_cost * overhead_rate)
        tech_fee = int((total_lab_cost + overhead_cost) * tech_fee_rate)
        
        subtotal = total_mat_cost + total_lab_cost + overhead_cost + tech_fee
        vat = int(subtotal * 0.1)
        total_amount = subtotal + vat
        
        return {
            "date": date_str,
            "loc": loc_str,
            "ndt_type": ndt_type,
            "work_time": work_time,
            "material_type": material_type,
            "qty": qty,
            "unit": unit_str,
            "corr": corr,
            "adjusted_qty": adjusted_qty,
            "mat_cost": total_mat_cost,
            "lab_cost": total_lab_cost,
            "overhead": overhead_cost,
            "tech": tech_fee,
            "subtotal": subtotal,
            "vat": vat,
            "total_amount": total_amount
        }

    def calculate(self):
        try:
            res = self._do_calculate()
            self.result_text.config(state=tk.NORMAL)
            self.result_text.delete(1.0, tk.END)
            
            mat_unit = MATERIAL_COST.get(res['material_type'], 0)
            lab_unit = LABOR_COST[res['work_time']][res['ndt_type']]
            
            txt = (f"▶ [현재 입력값] 일자: {res['date']} | 구간: {res['loc']}\n"
                   f"▶ [적용 기준] 보정계수: {res['corr']:.2f} | 재료비 단가: {mat_unit:,}원 | 인건비 단가: {lab_unit:,}원\n"
                   f"▶ [공급 가액] {res['subtotal']:,} 원 (재료비 {res['mat_cost']:,} + 인건비 {res['lab_cost']:,} + 제경비 {res['overhead']:,} + 기술료 {res['tech']:,})\n"
                   f"▶ [최종 금액] 총 청구액 {res['total_amount']:,} 원 (부가세 {res['vat']:,}원 포함)\n")
            
            self.result_text.insert(tk.END, txt)
            self.result_text.config(state=tk.DISABLED)
            return res
        except ValueError:
            messagebox.showerror("입력 오류", "숫자를 정확히 입력해주세요.")
            return None

    def on_tree_select(self, event):
        selected_items = self.tree.selection()
        if not selected_items:
            return
            
        item = selected_items[0]
        idx = self.tree.index(item)
        
        if idx < 0 or idx >= len(self.records):
            return
            
        res = self.records[idx]
        
        mat_unit = MATERIAL_COST.get(res['material_type'], 0)
        lab_unit = LABOR_COST[res['work_time']][res['ndt_type']]
        
        txt = (f"▶ [선택된 기록] 일자: {res['date']} | 구간: {res['loc']}\n"
               f"▶ [적용 기준] 보정계수: {res['corr']:.2f} | 재료비 단가: {mat_unit:,}원 | 인건비 단가: {lab_unit:,}원\n"
               f"▶ [공급 가액] {res['subtotal']:,} 원 (재료비 {res['mat_cost']:,} + 인건비 {res['lab_cost']:,} + 제경비 {res['overhead']:,} + 기술료 {res['tech']:,})\n"
               f"▶ [최종 금액] 총 청구액 {res['total_amount']:,} 원 (부가세 {res['vat']:,}원 포함)\n")
        
        self.result_text.config(state=tk.NORMAL)
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, txt)
        self.result_text.config(state=tk.DISABLED)

    def add_to_record(self):
        res = self.calculate()
        if res:
            self.records.append(res)
            self.tree.insert("", tk.END, values=(
                res["date"],
                res["loc"],
                res["ndt_type"], 
                res["work_time"], 
                res["material_type"], 
                f"{res['qty']:.1f}", 
                res["unit"],
                f"{res['corr']:.2f}", 
                f"{res['adjusted_qty']:.2f}", 
                f"{res['overhead']:,}",
                f"{res['tech']:,}",
                f"{res['subtotal']:,}" # 공급가액 소계 출력
            ))
            
    def clear_records(self):
        self.records = []
        for item in self.tree.get_children():
            self.tree.delete(item)

    def export_to_excel(self):
        if not self.records:
            messagebox.showwarning("기록 없음", "저장할 작업 기록이 없습니다.")
            return
            
        default_name = f"기성청구내역서_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel File", "*.xlsx")],
            title="정식 기성청구 엑셀 양식으로 저장"
        )
        
        if not filepath:
            return
            
        try:
            # 엑셀 애플리케이션 실행
            excel = win32.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            wb = excel.Workbooks.Add()
            ws = wb.ActiveSheet
            ws.Name = "기성청구내역서"
            
            # --- 상단 기본 정보 ---
            ws.Range("A1:O2").Merge()
            ws.Range("A1").Value = "비파괴검사기술용역 기성청구 내역서"
            ws.Range("A1").Font.Size = 20
            ws.Range("A1").Font.Bold = True
            ws.Range("A1").HorizontalAlignment = -4108 # xlCenter
            ws.Range("A1").VerticalAlignment = -4108
            
            ws.Range("A4:B4").Merge()
            ws.Range("A4").Value = "공 사 명 :"
            ws.Range("A4").Font.Bold = True
            ws.Range("C4:J4").Merge()
            ws.Range("C4").Value = "가산~가평 천연가스 공급시설 건설공사 비파괴검사 기술용역"
            
            ws.Range("L4:M4").Merge()
            ws.Range("L4").Value = "청구일자 :"
            ws.Range("L4").Font.Bold = True
            ws.Range("L4").HorizontalAlignment = -4152 # xlRight
            ws.Range("N4:O4").Merge()
            ws.Range("N4").Value = datetime.now().strftime('%Y년 %m월 %d일')
            
            # 총 합계 계산용
            total_sum = sum(r["total_amount"] for r in self.records)
            
            ws.Range("A5:B5").Merge()
            ws.Range("A5").Value = "총청구액 :"
            ws.Range("A5").Font.Bold = True
            ws.Range("C5:F5").Merge()
            ws.Range("C5").Value = f"₩ {total_sum:,}"
            ws.Range("C5").Font.Bold = True
            ws.Range("C5").Font.Size = 12
            
            # --- 테이블 헤더 ---
            headers = ["No.", "검사일자", "작업구간", "검사종류", "규격/자재", "근무형태", "실물량", "단위", "보정계수", "환산물량", 
                       "재료비", "직접인건비", "제경비", "기술료", "공급가액소계"]
            
            start_row = 7
            for col, h in enumerate(headers, start=1):
                cell = ws.Cells(start_row, col)
                cell.Value = h
                cell.Font.Bold = True
                cell.Interior.Color = 14277081 # Light Gray
                cell.HorizontalAlignment = -4108
                cell.Borders.LineStyle = 1
            
            # 열 너비 설정
            ws.Columns(1).ColumnWidth = 4  # No
            ws.Columns(2).ColumnWidth = 11 # 일자
            ws.Columns(3).ColumnWidth = 20 # 구간
            ws.Columns(4).ColumnWidth = 9  # 종류
            ws.Columns(5).ColumnWidth = 22 # 규격
            ws.Columns(6).ColumnWidth = 9  # 형태
            ws.Columns(7).ColumnWidth = 8  # 실물량
            ws.Columns(8).ColumnWidth = 5  # 단위
            ws.Columns(9).ColumnWidth = 9  # 계수
            ws.Columns(10).ColumnWidth = 9 # 환산
            ws.Columns(11).ColumnWidth = 11 # 재료비
            ws.Columns(12).ColumnWidth = 12 # 인건비
            ws.Columns(13).ColumnWidth = 11 # 제경비
            ws.Columns(14).ColumnWidth = 11 # 기술료
            ws.Columns(15).ColumnWidth = 13 # 공급가액
            
            # --- 그룹별 데이터 채우기 ---
            current_row = start_row + 1
            total_mat = total_lab = total_ovr = total_tech = total_sub = 0
            idx = 1
            
            for g_type in ["RT", "UT", "PT"]:
                group_records = [r for r in self.records if r["ndt_type"] == g_type]
                if not group_records:
                    continue
                
                sub_mat = sub_lab = sub_ovr = sub_tech = sub_sub = 0
                
                # 그룹별 레코드 작성
                for r in group_records:
                    ws.Cells(current_row, 1).Value = idx
                    ws.Cells(current_row, 2).Value = r["date"]
                    ws.Cells(current_row, 3).Value = r["loc"]
                    ws.Cells(current_row, 4).Value = r["ndt_type"]
                    ws.Cells(current_row, 5).Value = r["material_type"]
                    ws.Cells(current_row, 6).Value = r["work_time"]
                    ws.Cells(current_row, 7).Value = r["qty"]
                    ws.Cells(current_row, 8).Value = r["unit"]
                    ws.Cells(current_row, 9).Value = r["corr"]
                    ws.Cells(current_row, 10).Value = r["adjusted_qty"]
                    
                    ws.Cells(current_row, 11).Value = r["mat_cost"]
                    ws.Cells(current_row, 12).Value = r["lab_cost"]
                    ws.Cells(current_row, 13).Value = r["overhead"]
                    ws.Cells(current_row, 14).Value = r["tech"]
                    ws.Cells(current_row, 15).Value = r["subtotal"]
                    
                    # 테두리 및 정렬
                    for c in range(1, 16):
                        cell = ws.Cells(current_row, c)
                        cell.Borders.LineStyle = 1
                        if c <= 4 or c == 6 or c == 8:
                            cell.HorizontalAlignment = -4108 # 가운데
                        elif c == 5 or c == 3:
                            cell.HorizontalAlignment = -4131 # 왼쪽
                        else:
                            if c >= 11:
                                cell.NumberFormat = "#,##0" # 통화 포맷
                            else:
                                cell.NumberFormat = "0.00"
                    
                    sub_mat += r["mat_cost"]
                    sub_lab += r["lab_cost"]
                    sub_ovr += r["overhead"]
                    sub_tech += r["tech"]
                    sub_sub += r["subtotal"]
                    idx += 1
                    current_row += 1
                
                # 소계 (Subtotal) 행 작성
                ws.Range(ws.Cells(current_row, 1), ws.Cells(current_row, 10)).Merge()
                ws.Cells(current_row, 1).Value = f"[{g_type}] 검사 소계"
                ws.Cells(current_row, 1).HorizontalAlignment = -4108
                ws.Cells(current_row, 1).Font.Bold = True
                
                ws.Cells(current_row, 11).Value = sub_mat
                ws.Cells(current_row, 12).Value = sub_lab
                ws.Cells(current_row, 13).Value = sub_ovr
                ws.Cells(current_row, 14).Value = sub_tech
                ws.Cells(current_row, 15).Value = sub_sub
                
                for c in range(1, 16):
                    cell = ws.Cells(current_row, c)
                    cell.Borders.LineStyle = 1
                    cell.Font.Bold = True
                    cell.Interior.Color = 15987699 # Light Blue
                    if c >= 11:
                        cell.NumberFormat = "#,##0"
                
                total_mat += sub_mat
                total_lab += sub_lab
                total_ovr += sub_ovr
                total_tech += sub_tech
                total_sub += sub_sub
                current_row += 1
                
            # --- 전체 합계 행 ---
            ws.Range(ws.Cells(current_row, 1), ws.Cells(current_row, 10)).Merge()
            ws.Cells(current_row, 1).Value = "총   합   계"
            ws.Cells(current_row, 1).HorizontalAlignment = -4108
            ws.Cells(current_row, 1).Font.Bold = True
            
            ws.Cells(current_row, 11).Value = total_mat
            ws.Cells(current_row, 12).Value = total_lab
            ws.Cells(current_row, 13).Value = total_ovr
            ws.Cells(current_row, 14).Value = total_tech
            ws.Cells(current_row, 15).Value = total_sub
            
            for c in range(1, 16):
                cell = ws.Cells(current_row, c)
                cell.Borders.LineStyle = 1
                cell.Font.Bold = True
                cell.Interior.Color = 14277081
                if c >= 11:
                    cell.NumberFormat = "#,##0"
                    
            # --- 부가세 및 총 청구액 ---
            current_row += 1
            ws.Range(ws.Cells(current_row, 1), ws.Cells(current_row, 14)).Merge()
            ws.Cells(current_row, 1).Value = "부가가치세 (10%)"
            ws.Cells(current_row, 1).HorizontalAlignment = -4152 # xlRight
            ws.Cells(current_row, 15).Value = int(total_sub * 0.1)
            ws.Cells(current_row, 15).NumberFormat = "#,##0"
            for c in range(1, 16):
                ws.Cells(current_row, c).Borders.LineStyle = 1
            
            current_row += 1
            ws.Range(ws.Cells(current_row, 1), ws.Cells(current_row, 14)).Merge()
            ws.Cells(current_row, 1).Value = "최 종 총 기 성 청 구 액"
            ws.Cells(current_row, 1).HorizontalAlignment = -4152
            ws.Cells(current_row, 1).Font.Bold = True
            ws.Cells(current_row, 1).Font.Size = 12
            
            total_final = total_sub + int(total_sub * 0.1)
            ws.Cells(current_row, 15).Value = total_final
            ws.Cells(current_row, 15).NumberFormat = "#,##0"
            ws.Cells(current_row, 15).Font.Bold = True
            ws.Cells(current_row, 15).Font.Size = 12
            
            for c in range(1, 16):
                cell = ws.Cells(current_row, c)
                cell.Borders.LineStyle = 1
                cell.Interior.Color = 13434879 # Light Yellow
                
            # 파일 저장
            filepath = filepath.replace("/", "\\")
            wb.SaveAs(filepath)
            wb.Close()
            excel.Quit()
            
            messagebox.showinfo("저장 완료", f"그룹화된 엑셀 기성청구 내역서가 성공적으로 생성되었습니다.\n{filepath}")
            os.startfile(filepath)
            
        except Exception as e:
            messagebox.showerror("저장 오류", f"엑셀 파일 생성 중 오류가 발생했습니다.\n{str(e)}")
            try:
                excel.Quit()
            except:
                pass

if __name__ == "__main__":
    app = NDTCalculator()
    app.mainloop()
