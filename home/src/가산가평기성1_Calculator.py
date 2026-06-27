import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import os
import json
import win32com.client as win32

CONFIG_FILE = "config.json"

DEFAULT_CONFIG = {
    "MATERIAL_COST": {
        'RT (B필름: 3⅓"x17")': 3379,
        'RT (A필름: 3⅓"x12")': 2540,
        'RT (A/2필름: 3⅓"x6")': 1515,
        "UT": 1115,
        "PT": 3971
    },
    "LABOR_COST": {
        "일반": {"RT": 34863, "UT": 25000, "PT": 20000},
        "야간": {"RT": 49240, "UT": 37500, "PT": 30000},
        "휴일": {"RT": 49313, "UT": 37500, "PT": 30000}
    }
}

def load_config():
    if not os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=4)
        return DEFAULT_CONFIG
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except:
        return DEFAULT_CONFIG

def save_config(config):
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=4)

# 글로벌 단가 변수
CONFIG = load_config()
MATERIAL_COST = CONFIG["MATERIAL_COST"]
LABOR_COST = CONFIG["LABOR_COST"]

class NDTCalculator(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("비파괴검사 기성 산출 계산기 (가산~가평)")
        self.geometry("1150x1000")  
        self.configure(padx=20, pady=20)
        self.style = ttk.Style(self)
        self.style.theme_use("clam")
        
        self.records = [] # 저장된 기록 목록
        
        self.create_menu()
        
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        self.tab1 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab1, text="일일 작업 내역 관리 (단가계산)")
        self.create_widgets(self.tab1)
        
        self.load_progress_history()
        
        self.tab2 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab2, text="월간 공식 기성서류 생성")
        self.create_progress_docs_tab(self.tab2)
        
    def create_menu(self):
        menubar = tk.Menu(self)
        
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="작업 불러오기 (Load)", command=self.load_project)
        file_menu.add_command(label="작업 저장하기 (Save)", command=self.save_project)
        file_menu.add_separator()
        file_menu.add_command(label="단가 설정 (Settings)", command=self.open_settings)
        file_menu.add_separator()
        file_menu.add_command(label="종료", command=self.quit)
        
        menubar.add_cascade(label="파일 (File)", menu=file_menu)
        self.config(menu=menubar)

    def create_widgets(self, parent):
        main_pane = ttk.PanedWindow(parent, orient=tk.HORIZONTAL)
        main_pane.pack(fill=tk.BOTH, expand=True)
        
        left_frame = ttk.Frame(main_pane)
        right_frame = ttk.Frame(main_pane, padding=(10, 0, 0, 0))
        
        main_pane.add(left_frame, weight=3)
        main_pane.add(right_frame, weight=1)
        
        # --- LEFT FRAME (기존 검사 입력부) ---
        info_frame = ttk.Frame(left_frame)
        info_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(info_frame, text="• 검사일자:", font=("Arial", 10, "bold")).pack(side=tk.LEFT)
        self.date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        ttk.Entry(info_frame, textvariable=self.date_var, width=15).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(info_frame, text="• 작업구간 (Joint No 등):", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=(20, 5))
        self.loc_var = tk.StringVar(value="")
        ttk.Entry(info_frame, textvariable=self.loc_var, width=30).pack(side=tk.LEFT, padx=5)

        ttk.Label(left_frame, text="1. 검사 종류", font=("Arial", 11, "bold")).pack(anchor=tk.W, pady=(10, 5))
        self.ndt_type_var = tk.StringVar(value="RT")
        type_frame = ttk.Frame(left_frame)
        type_frame.pack(fill=tk.X, pady=5)
        for t in ["RT", "UT", "PT"]:
            ttk.Radiobutton(type_frame, text=t, value=t, variable=self.ndt_type_var, command=self.update_dynamic_ui).pack(side=tk.LEFT, padx=10)
            
        ttk.Label(left_frame, text="2. 작업 형태 (근무 시간대)", font=("Arial", 11, "bold")).pack(anchor=tk.W, pady=(10, 5))
        self.work_time_var = tk.StringVar(value="일반")
        time_frame = ttk.Frame(left_frame)
        time_frame.pack(fill=tk.X, pady=5)
        for t in ["일반", "야간", "휴일"]:
            ttk.Radiobutton(time_frame, text=t, value=t, variable=self.work_time_var).pack(side=tk.LEFT, padx=10)

        self.material_lbl = ttk.Label(left_frame, text="3. 사용 자재 (RT 필름 규격)", font=("Arial", 11, "bold"))
        self.material_lbl.pack(anchor=tk.W, pady=(10, 5))
        self.material_var = tk.StringVar(value='RT (B필름: 3⅓"x17")')
        self.material_combo = ttk.Combobox(left_frame, textvariable=self.material_var, values=['RT (B필름: 3⅓"x17")', 'RT (A필름: 3⅓"x12")', 'RT (A/2필름: 3⅓"x6")'], state="readonly")
        self.material_combo.pack(fill=tk.X, pady=5)
        
        self.dynamic_frame = ttk.LabelFrame(left_frame, text="4. 보정계수 조건 선택", padding=10)
        self.dynamic_frame.pack(fill=tk.X, pady=(10, 5))
        
        self.source_frame = ttk.Frame(self.dynamic_frame)
        ttk.Label(self.source_frame, text="• 방사선원 :", width=15).pack(side=tk.LEFT)
        self.source_var = tk.StringVar(value="Ir-192 또는 Se-75 (1.0)")
        self.source_combo = ttk.Combobox(self.source_frame, textvariable=self.source_var, state="readonly", width=35)
        self.source_combo['values'] = ["Ir-192 또는 Se-75 (1.0)", "X-ray 발생장치 (1.3)"]
        self.source_combo.pack(side=tk.LEFT)
        
        self.pipe_frame = ttk.Frame(self.dynamic_frame)
        ttk.Label(self.pipe_frame, text="• 관경(구경) :", width=15).pack(side=tk.LEFT)
        self.pipe_var = tk.StringVar()
        self.pipe_combo = ttk.Combobox(self.pipe_frame, textvariable=self.pipe_var, state="readonly", width=35)
        self.pipe_combo.pack(side=tk.LEFT)
        
        self.thickness_frame = ttk.Frame(self.dynamic_frame)
        ttk.Label(self.thickness_frame, text="• 투과/모재두께 :", width=15).pack(side=tk.LEFT)
        self.thickness_var = tk.StringVar()
        self.thickness_combo = ttk.Combobox(self.thickness_frame, textvariable=self.thickness_var, state="readonly", width=35)
        self.thickness_combo.pack(side=tk.LEFT)
        
        ttk.Label(left_frame, text="5. 실검사 물량 (RT: 매 / UT,PT: Meter)", font=("Arial", 11, "bold")).pack(anchor=tk.W, pady=(10, 5))
        self.quantity_var = tk.DoubleVar(value=10.0)
        ttk.Entry(left_frame, textvariable=self.quantity_var).pack(fill=tk.X, pady=5)
        
        rate_frame = ttk.Frame(left_frame)
        rate_frame.pack(fill=tk.X, pady=(15, 5))
        ttk.Label(rate_frame, text="6. 적용 요율 (%)", font=("Arial", 11, "bold")).pack(side=tk.LEFT, padx=(0, 20))
        
        ttk.Label(rate_frame, text="제경비율:").pack(side=tk.LEFT)
        self.overhead_rate_var = tk.DoubleVar(value=80.0)
        ttk.Entry(rate_frame, textvariable=self.overhead_rate_var, width=8).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(rate_frame, text="기술료율:").pack(side=tk.LEFT, padx=(15, 0))
        self.tech_fee_rate_var = tk.DoubleVar(value=5.86)
        ttk.Entry(rate_frame, textvariable=self.tech_fee_rate_var, width=8).pack(side=tk.LEFT, padx=5)
        
        btn_frame = ttk.Frame(left_frame)
        btn_frame.pack(fill=tk.X, pady=15)
        
        ttk.Button(btn_frame, text="금액 계산하기", command=self.calculate).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5, ipady=8)
        ttk.Button(btn_frame, text="기록 목록에 추가", command=self.add_to_record).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5, ipady=8)
        
        ttk.Label(left_frame, text="[ 단일 계산 결과 ]", font=("Arial", 11, "bold")).pack(anchor=tk.W, pady=(0, 5))
        self.result_text = tk.Text(left_frame, height=5, width=50, state=tk.DISABLED, font=("Consolas", 11))
        self.result_text.pack(fill=tk.X, expand=False)
        
        # --- RIGHT FRAME (기타 경비 및 실비 정산) ---
        exp_frame = ttk.LabelFrame(right_frame, text="기타 경비 및 실비 정산 (월간)", padding=10)
        exp_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(exp_frame, text="엑셀 청구서 하단에 합산될\n추가 실비정산 금액을 입력하세요.\n(세액 미포함 금액)").pack(pady=(0, 10))
        
        ttk.Label(exp_frame, text="장비손료 (원):").pack(anchor=tk.W)
        self.equip_cost_var = tk.IntVar(value=0)
        ttk.Entry(exp_frame, textvariable=self.equip_cost_var).pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(exp_frame, text="안전관리비 (원):").pack(anchor=tk.W)
        self.safety_cost_var = tk.IntVar(value=0)
        ttk.Entry(exp_frame, textvariable=self.safety_cost_var).pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(exp_frame, text="주재비 및 출장여비 (원):").pack(anchor=tk.W)
        self.travel_cost_var = tk.IntVar(value=0)
        ttk.Entry(exp_frame, textvariable=self.travel_cost_var).pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(exp_frame, text="도서인쇄비 (원):").pack(anchor=tk.W)
        self.print_cost_var = tk.IntVar(value=0)
        ttk.Entry(exp_frame, textvariable=self.print_cost_var).pack(fill=tk.X, pady=(0, 10))
        
        # --- BOTTOM FRAME (누적 테이블) ---
        bottom_frame = ttk.Frame(parent)
        bottom_frame.pack(fill=tk.BOTH, expand=True, pady=(15, 0))
        
        lbl_frame = ttk.Frame(bottom_frame)
        lbl_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(lbl_frame, text="[ 일일 작업 기록 목록 ]", font=("Arial", 11, "bold")).pack(side=tk.LEFT)
        ttk.Button(lbl_frame, text="엑셀 파일로 저장", command=self.export_to_excel).pack(side=tk.RIGHT)
        ttk.Button(lbl_frame, text="기록 초기화", command=self.clear_records).pack(side=tk.RIGHT, padx=5)

        columns = ("date", "loc", "type", "time", "mat", "qty", "unit", "corr", "adj_qty", "overhead", "tech", "total_amt")
        self.tree = ttk.Treeview(bottom_frame, columns=columns, show="headings", height=8)
        
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
                res["date"], res["loc"], res["ndt_type"], res["work_time"], 
                res["material_type"], f"{res['qty']:.1f}", res["unit"],
                f"{res['corr']:.2f}", f"{res['adjusted_qty']:.2f}", 
                f"{res['overhead']:,}", f"{res['tech']:,}", f"{res['subtotal']:,}"
            ))
            
    def clear_records(self):
        self.records = []
        for item in self.tree.get_children():
            self.tree.delete(item)
            
    def save_project(self):
        try:
            filepath = filedialog.asksaveasfilename(defaultextension=".ndt", filetypes=[("NDT Project", "*.ndt")], title="작업 저장하기")
            if not filepath: return
            
            data = {
                "records": self.records,
                "expenses": {
                    "equip": self.equip_cost_var.get(),
                    "safety": self.safety_cost_var.get(),
                    "travel": self.travel_cost_var.get(),
                    "print": self.print_cost_var.get()
                }
            }
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
            messagebox.showinfo("저장 완료", "작업이 성공적으로 저장되었습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"저장 중 오류가 발생했습니다: {e}")

    def load_project(self):
        try:
            filepath = filedialog.askopenfilename(filetypes=[("NDT Project", "*.ndt")], title="작업 불러오기")
            if not filepath: return
            
            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)
                
            self.clear_records()
            self.records = data.get("records", [])
            for res in self.records:
                self.tree.insert("", tk.END, values=(
                    res["date"], res["loc"], res["ndt_type"], res["work_time"], 
                    res["material_type"], f"{res['qty']:.1f}", res["unit"],
                    f"{res['corr']:.2f}", f"{res['adjusted_qty']:.2f}", 
                    f"{res['overhead']:,}", f"{res['tech']:,}", f"{res['subtotal']:,}"
                ))
            
            ex = data.get("expenses", {})
            self.equip_cost_var.set(ex.get("equip", 0))
            self.safety_cost_var.set(ex.get("safety", 0))
            self.travel_cost_var.set(ex.get("travel", 0))
            self.print_cost_var.set(ex.get("print", 0))
            
            messagebox.showinfo("불러오기 완료", "작업을 성공적으로 불러왔습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"불러오기 중 오류가 발생했습니다: {e}")

    def open_settings(self):
        top = tk.Toplevel(self)
        top.title("단가 설정 (Settings)")
        top.geometry("450x550")
        top.configure(padx=20, pady=20)
        
        ttk.Label(top, text="[재료비 단가 설정]", font=("Arial", 11, "bold")).pack(anchor=tk.W, pady=(0, 5))
        mat_vars = {}
        for k, v in MATERIAL_COST.items():
            f = ttk.Frame(top)
            f.pack(fill=tk.X, pady=2)
            ttk.Label(f, text=k, width=25).pack(side=tk.LEFT)
            var = tk.IntVar(value=v)
            ttk.Entry(f, textvariable=var, width=15).pack(side=tk.RIGHT)
            mat_vars[k] = var
            
        ttk.Label(top, text="[인건비 단가 설정]", font=("Arial", 11, "bold")).pack(anchor=tk.W, pady=(20, 5))
        lab_vars = {}
        for w_time in ["일반", "야간", "휴일"]:
            ttk.Label(top, text=f"■ {w_time}", font=("Arial", 9, "bold")).pack(anchor=tk.W, pady=(5, 2))
            lab_vars[w_time] = {}
            for t in ["RT", "UT", "PT"]:
                f = ttk.Frame(top)
                f.pack(fill=tk.X, pady=2)
                ttk.Label(f, text=f"{w_time} - {t}", width=25).pack(side=tk.LEFT)
                var = tk.IntVar(value=LABOR_COST[w_time][t])
                ttk.Entry(f, textvariable=var, width=15).pack(side=tk.RIGHT)
                lab_vars[w_time][t] = var
                
        def save_and_close():
            global MATERIAL_COST, LABOR_COST
            for k, var in mat_vars.items():
                MATERIAL_COST[k] = var.get()
            for w_time in lab_vars:
                for t, var in lab_vars[w_time].items():
                    LABOR_COST[w_time][t] = var.get()
            CONFIG["MATERIAL_COST"] = MATERIAL_COST
            CONFIG["LABOR_COST"] = LABOR_COST
            save_config(CONFIG)
            messagebox.showinfo("저장 완료", "새로운 단가가 저장되었습니다.")
            top.destroy()
            
        ttk.Button(top, text="단가 저장하기", command=save_and_close).pack(pady=20, ipady=5, fill=tk.X)

    def export_to_excel(self):
        if not self.records:
            messagebox.showwarning("기록 없음", "저장할 작업 기록이 없습니다.")
            return
            
        default_name = f"기성청구내역서_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=default_name, filetypes=[("Excel File", "*.xlsx")], title="정식 기성청구 엑셀 양식으로 저장")
        if not filepath: return
            
        try:
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
            ws.Range("A1").HorizontalAlignment = -4108
            ws.Range("A1").VerticalAlignment = -4108
            
            ws.Range("A4:B4").Merge()
            ws.Range("A4").Value = "공 사 명 :"
            ws.Range("A4").Font.Bold = True
            ws.Range("C4:J4").Merge()
            ws.Range("C4").Value = "가산~가평 천연가스 공급시설 건설공사 비파괴검사 기술용역"
            
            ws.Range("L4:M4").Merge()
            ws.Range("L4").Value = "청구일자 :"
            ws.Range("L4").Font.Bold = True
            ws.Range("L4").HorizontalAlignment = -4152
            ws.Range("N4:O4").Merge()
            ws.Range("N4").Value = datetime.now().strftime('%Y년 %m월 %d일')
            
            headers = ["No.", "검사일자", "작업구간", "검사종류", "규격/자재", "근무형태", "실물량", "단위", "보정계수", "환산물량", 
                       "재료비", "직접인건비", "제경비", "기술료", "공급가액소계"]
            
            start_row = 6
            for col, h in enumerate(headers, start=1):
                cell = ws.Cells(start_row, col)
                cell.Value = h
                cell.Font.Bold = True
                cell.Interior.Color = 14277081
                cell.HorizontalAlignment = -4108
                cell.Borders.LineStyle = 1
            
            ws.Columns(1).ColumnWidth = 4
            ws.Columns(2).ColumnWidth = 11
            ws.Columns(3).ColumnWidth = 20
            ws.Columns(4).ColumnWidth = 9
            ws.Columns(5).ColumnWidth = 22
            ws.Columns(6).ColumnWidth = 9
            ws.Columns(7).ColumnWidth = 8
            ws.Columns(8).ColumnWidth = 5
            ws.Columns(9).ColumnWidth = 9
            ws.Columns(10).ColumnWidth = 9
            ws.Columns(11).ColumnWidth = 11
            ws.Columns(12).ColumnWidth = 12
            ws.Columns(13).ColumnWidth = 11
            ws.Columns(14).ColumnWidth = 11
            ws.Columns(15).ColumnWidth = 13
            
            current_row = start_row + 1
            total_mat = total_lab = total_ovr = total_tech = total_sub = 0
            idx = 1
            
            for g_type in ["RT", "UT", "PT"]:
                group_records = [r for r in self.records if r["ndt_type"] == g_type]
                if not group_records: continue
                
                sub_mat = sub_lab = sub_ovr = sub_tech = sub_sub = 0
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
                    
                    for c in range(1, 16):
                        cell = ws.Cells(current_row, c)
                        cell.Borders.LineStyle = 1
                        if c <= 4 or c == 6 or c == 8: cell.HorizontalAlignment = -4108
                        elif c == 5 or c == 3: cell.HorizontalAlignment = -4131
                        else: cell.NumberFormat = "#,##0" if c >= 11 else "0.00"
                    
                    sub_mat += r["mat_cost"]; sub_lab += r["lab_cost"]; sub_ovr += r["overhead"]
                    sub_tech += r["tech"]; sub_sub += r["subtotal"]
                    idx += 1; current_row += 1
                
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
                    cell.Interior.Color = 15987699
                    if c >= 11: cell.NumberFormat = "#,##0"
                
                total_mat += sub_mat; total_lab += sub_lab; total_ovr += sub_ovr
                total_tech += sub_tech; total_sub += sub_sub
                current_row += 1
                
            # --- 전체 합계 ---
            ws.Range(ws.Cells(current_row, 1), ws.Cells(current_row, 10)).Merge()
            ws.Cells(current_row, 1).Value = "검사 비용 합계"
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
                if c >= 11: cell.NumberFormat = "#,##0"
                    
            # --- 실비 정산 추가 ---
            current_row += 1
            extra_items = [
                ("장비손료", self.equip_cost_var.get()),
                ("안전관리비", self.safety_cost_var.get()),
                ("주재비 및 출장여비", self.travel_cost_var.get()),
                ("도서인쇄비", self.print_cost_var.get())
            ]
            
            total_extra = 0
            for name, val in extra_items:
                if val > 0:
                    ws.Range(ws.Cells(current_row, 1), ws.Cells(current_row, 14)).Merge()
                    ws.Cells(current_row, 1).Value = f"+ {name}"
                    ws.Cells(current_row, 1).HorizontalAlignment = -4152
                    ws.Cells(current_row, 15).Value = val
                    ws.Cells(current_row, 15).NumberFormat = "#,##0"
                    for c in range(1, 16): ws.Cells(current_row, c).Borders.LineStyle = 1
                    total_extra += val
                    current_row += 1
            
            # --- 총 공급가액 (검사합계 + 실비) ---
            grand_subtotal = total_sub + total_extra
            
            ws.Range(ws.Cells(current_row, 1), ws.Cells(current_row, 14)).Merge()
            ws.Cells(current_row, 1).Value = "공급가액 총액"
            ws.Cells(current_row, 1).HorizontalAlignment = -4152
            ws.Cells(current_row, 1).Font.Bold = True
            ws.Cells(current_row, 15).Value = grand_subtotal
            ws.Cells(current_row, 15).NumberFormat = "#,##0"
            ws.Cells(current_row, 15).Font.Bold = True
            for c in range(1, 16): ws.Cells(current_row, c).Borders.LineStyle = 1
            
            # --- 부가세 및 최종 청구액 ---
            current_row += 1
            ws.Range(ws.Cells(current_row, 1), ws.Cells(current_row, 14)).Merge()
            ws.Cells(current_row, 1).Value = "+ 부가가치세 (10%)"
            ws.Cells(current_row, 1).HorizontalAlignment = -4152
            vat_val = int(grand_subtotal * 0.1)
            ws.Cells(current_row, 15).Value = vat_val
            ws.Cells(current_row, 15).NumberFormat = "#,##0"
            for c in range(1, 16): ws.Cells(current_row, c).Borders.LineStyle = 1
            
            current_row += 1
            ws.Range(ws.Cells(current_row, 1), ws.Cells(current_row, 14)).Merge()
            ws.Cells(current_row, 1).Value = "최 종 기 성 청 구 액"
            ws.Cells(current_row, 1).HorizontalAlignment = -4152
            ws.Cells(current_row, 1).Font.Bold = True
            ws.Cells(current_row, 1).Font.Size = 12
            
            total_final = grand_subtotal + vat_val
            ws.Cells(current_row, 15).Value = total_final
            ws.Cells(current_row, 15).NumberFormat = "#,##0"
            ws.Cells(current_row, 15).Font.Bold = True
            ws.Cells(current_row, 15).Font.Size = 12
            
            for c in range(1, 16):
                cell = ws.Cells(current_row, c)
                cell.Borders.LineStyle = 1
                cell.Interior.Color = 13434879
                
            filepath = filepath.replace("/", "\\")
            wb.SaveAs(filepath)
            wb.Close()
            excel.Quit()
            
            messagebox.showinfo("저장 완료", f"엑셀 기성청구 내역서가 성공적으로 생성되었습니다.\n{filepath}")
            os.startfile(filepath)
            
        except Exception as e:
            messagebox.showerror("저장 오류", f"엑셀 파일 생성 중 오류가 발생했습니다.\n{str(e)}")
            try: excel.Quit()
            except: pass

    # ---------------------------------------------------------
    # 기성서류(Progress Docs) 자동화 탭 관련 메서드
    # ---------------------------------------------------------
    def load_progress_history(self):
        self.history_file = "progress_history.json"
        self.history_data = {
            "contract": {"RT_qty": 0, "RT_amt": 0, "UT_qty": 0, "UT_amt": 0, "PT_qty": 0, "PT_amt": 0, "total_amt": 0},
            "prev": {"RT_qty": 0, "RT_amt": 0, "UT_qty": 0, "UT_amt": 0, "PT_qty": 0, "PT_amt": 0, "total_amt": 0}
        }
        if os.path.exists(self.history_file):
            try:
                with open(self.history_file, 'r', encoding='utf-8') as f:
                    self.history_data = json.load(f)
            except: pass

    def save_progress_history(self):
        with open(self.history_file, 'w', encoding='utf-8') as f:
            json.dump(self.history_data, f, ensure_ascii=False, indent=4)

    def create_progress_docs_tab(self, parent):
        # 도급(계약) 정보 프레임
        contract_frame = ttk.LabelFrame(parent, text="1. 도급(계약) 물량 및 금액 설정", padding=10)
        contract_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.c_vars = {}
        row_idx = 0
        ttk.Label(contract_frame, text="총 도급 금액(원):").grid(row=row_idx, column=0, sticky="w")
        self.c_vars["total_amt"] = tk.DoubleVar(value=self.history_data["contract"].get("total_amt", 0))
        ttk.Entry(contract_frame, textvariable=self.c_vars["total_amt"]).grid(row=row_idx, column=1, padx=5, pady=2)
        row_idx += 1
        
        for t in ["RT", "UT", "PT"]:
            ttk.Label(contract_frame, text=f"[{t}] 도급 수량:").grid(row=row_idx, column=0, sticky="w")
            self.c_vars[f"{t}_qty"] = tk.DoubleVar(value=self.history_data["contract"].get(f"{t}_qty", 0))
            ttk.Entry(contract_frame, textvariable=self.c_vars[f"{t}_qty"]).grid(row=row_idx, column=1, padx=5, pady=2)
            
            ttk.Label(contract_frame, text=f"[{t}] 도급 금액(원):").grid(row=row_idx, column=2, sticky="w", padx=(20,0))
            self.c_vars[f"{t}_amt"] = tk.DoubleVar(value=self.history_data["contract"].get(f"{t}_amt", 0))
            ttk.Entry(contract_frame, textvariable=self.c_vars[f"{t}_amt"]).grid(row=row_idx, column=3, padx=5, pady=2)
            row_idx += 1
            
        ttk.Button(contract_frame, text="도급 정보 저장", command=self.save_contract_info).grid(row=row_idx, column=0, columnspan=4, pady=10)

        # 상태 표시 프레임 (전회 기성 / 금회 기성)
        status_frame = ttk.LabelFrame(parent, text="2. 기성 물량 상태 (전회 누적 & 금회 집계)", padding=10)
        status_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.status_text = tk.Text(status_frame, height=12, state=tk.DISABLED, font=("Consolas", 11))
        self.status_text.pack(fill=tk.BOTH, expand=True)
        
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(btn_frame, text="금회 기성 집계하기 (1탭 데이터 반영)", command=self.update_progress_status).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5, ipady=10)
        ttk.Button(btn_frame, text="공식 기성서류 엑셀 생성", command=self.generate_progress_excel).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5, ipady=10)
        ttk.Button(btn_frame, text="★ 기성 확정 (금회 기성을 전회 기성으로 이관)", command=self.confirm_progress).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5, ipady=10)
        
        self.update_progress_status()

    def save_contract_info(self):
        for k, var in self.c_vars.items():
            self.history_data["contract"][k] = var.get()
        self.save_progress_history()
        messagebox.showinfo("저장", "도급(계약) 정보가 저장되었습니다.")
        self.update_progress_status()

    def get_current_progress(self):
        # 1탭의 self.records 데이터를 바탕으로 금회 기성을 집계
        cur = {"RT_qty": 0, "RT_amt": 0, "UT_qty": 0, "UT_amt": 0, "PT_qty": 0, "PT_amt": 0, "total_amt": 0}
        
        for r in self.records:
            t = r["ndt_type"]
            cur[f"{t}_qty"] += r["qty"]
            cur[f"{t}_amt"] += r["subtotal"]
            cur["total_amt"] += r["subtotal"]
            
        extra_amt = self.equip_cost_var.get() + self.safety_cost_var.get() + self.travel_cost_var.get() + self.print_cost_var.get()
        cur["total_amt"] += extra_amt
        return cur

    def update_progress_status(self):
        cur = self.get_current_progress()
        prev = self.history_data["prev"]
        c = self.history_data["contract"]
        
        lines = []
        lines.append("[ 총 기성액 요약 ]")
        lines.append(f" - 도급 총액 : {c['total_amt']:,.0f} 원")
        lines.append(f" - 전회 기성 : {prev['total_amt']:,.0f} 원")
        lines.append(f" - 금회 기성 : {cur['total_amt']:,.0f} 원 (1탭 실비정산 포함)")
        
        total_accum = prev['total_amt'] + cur['total_amt']
        rate = (total_accum / c['total_amt'] * 100) if c['total_amt'] > 0 else 0
        lines.append(f" - 누계 기성 : {total_accum:,.0f} 원 (기성률: {rate:.2f}%)")
        lines.append(f" - 잔여 기성 : {c['total_amt'] - total_accum:,.0f} 원\n")
        
        lines.append("[ 항목별 수량 요약 ]")
        for t in ["RT", "UT", "PT"]:
            c_qty, c_amt = c.get(f"{t}_qty", 0), c.get(f"{t}_amt", 0)
            p_qty, p_amt = prev.get(f"{t}_qty", 0), prev.get(f"{t}_amt", 0)
            u_qty, u_amt = cur.get(f"{t}_qty", 0), cur.get(f"{t}_amt", 0)
            a_qty = p_qty + u_qty
            lines.append(f" [{t}] 도급: {c_qty:,.1f} | 전회: {p_qty:,.1f} | 금회: {u_qty:,.1f} | 누계: {a_qty:,.1f} | 잔여: {c_qty - a_qty:,.1f}")
        
        self.status_text.config(state=tk.NORMAL)
        self.status_text.delete(1.0, tk.END)
        self.status_text.insert(tk.END, "\n".join(lines))
        self.status_text.config(state=tk.DISABLED)

    def confirm_progress(self):
        if not self.records and self.get_current_progress()['total_amt'] == 0:
            messagebox.showwarning("오류", "금회 기성 데이터가 없습니다. (1탭에 데이터 없음)")
            return
            
        cur = self.get_current_progress()
        ans = messagebox.askyesno("기성 확정", f"금회 기성({cur['total_amt']:,.0f}원)을 전회 기성으로 누적하시겠습니까?\n이 작업은 되돌릴 수 없습니다.")
        if ans:
            for k in self.history_data["prev"].keys():
                self.history_data["prev"][k] += cur[k]
            self.save_progress_history()
            messagebox.showinfo("완료", "기성이 확정되어 누적되었습니다.\n새로운 달의 청구를 위해 1탭의 일일 기록을 초기화하세요.")
            self.update_progress_status()

    def generate_progress_excel(self):
        self.update_progress_status()
        cur = self.get_current_progress()
        prev = self.history_data["prev"]
        c = self.history_data["contract"]
        
        if c['total_amt'] == 0:
            messagebox.showwarning("경고", "총 도급 금액이 0원입니다. 1번 프레임에서 도급액을 설정해주세요.")
            return

        default_name = f"공식_기성서류_{datetime.now().strftime('%Y%m')}.xlsx"
        filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=default_name, filetypes=[("Excel File", "*.xlsx")], title="기성서류 생성")
        if not filepath: return
        
        try:
            excel = win32.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            wb = excel.Workbooks.Add()
            
            # --- 시트 1: 기성부분 검사조서 ---
            ws1 = wb.ActiveSheet
            ws1.Name = "기성부분 검사조서"
            
            ws1.Range("A2:I3").Merge()
            ws1.Range("A2").Value = "기성부분 검사조서"
            ws1.Range("A2").Font.Size = 22
            ws1.Range("A2").Font.Bold = True
            ws1.Range("A2").HorizontalAlignment = -4108
            
            ws1.Range("A6:B6").Merge()
            ws1.Range("A6").Value = "공 사 명 :"
            ws1.Range("C6:H6").Merge()
            ws1.Range("C6").Value = "가산~가평 천연가스 공급시설 건설공사 비파괴검사 기술용역"
            
            ws1.Range("A8:B8").Merge()
            ws1.Range("A8").Value = "도급금액 :"
            ws1.Range("C8:E8").Merge()
            ws1.Range("C8").Value = f"\\ {c['total_amt']:,.0f}"
            ws1.Range("C8").Font.Bold = True
            
            total_accum = prev['total_amt'] + cur['total_amt']
            ws1.Range("A9:B9").Merge()
            ws1.Range("A9").Value = "전회기성액 :"
            ws1.Range("C9:E9").Merge()
            ws1.Range("C9").Value = f"\\ {prev['total_amt']:,.0f}"
            
            ws1.Range("A10:B10").Merge()
            ws1.Range("A10").Value = "금회기성액 :"
            ws1.Range("C10:E10").Merge()
            ws1.Range("C10").Value = f"\\ {cur['total_amt']:,.0f}"
            ws1.Range("C10").Font.Bold = True
            
            ws1.Range("A11:B11").Merge()
            ws1.Range("A11").Value = "누계기성액 :"
            ws1.Range("C11:E11").Merge()
            ws1.Range("C11").Value = f"\\ {total_accum:,.0f}"
            
            ws1.Range("A12:B12").Merge()
            ws1.Range("A12").Value = "잔여금액 :"
            ws1.Range("C12:E12").Merge()
            ws1.Range("C12").Value = f"\\ {c['total_amt'] - total_accum:,.0f}"
            
            rate = (total_accum / c['total_amt'] * 100) if c['total_amt'] > 0 else 0
            ws1.Range("G8:H8").Merge()
            ws1.Range("G8").Value = "기성공정률 :"
            ws1.Cells(8, 9).Value = f"{rate:.2f} %"
            ws1.Cells(8, 9).Font.Bold = True
            
            ws1.Range("A15:I15").Merge()
            ws1.Range("A15").Value = "위 공사(용역)의 기성부분 검사 결과, 위와 같이 공정이 진척되었음을 인정함."
            ws1.Range("A15").HorizontalAlignment = -4108
            
            ws1.Range("G18:I18").Merge()
            ws1.Range("G18").Value = datetime.now().strftime('%Y년 %m월 %d일')
            ws1.Range("G18").HorizontalAlignment = -4152
            
            ws1.Range("F21:G21").Merge()
            ws1.Range("F21").Value = "검사원 :"
            ws1.Range("H21:I21").Merge()
            ws1.Range("H21").Value = "(인)"
            
            ws1.Columns("A:B").ColumnWidth = 10
            ws1.Columns("C:E").ColumnWidth = 15
            
            # --- 시트 2: 기성 내역서 ---
            ws2 = wb.Sheets.Add(After=ws1)
            ws2.Name = "기성 내역서"
            
            ws2.Range("A1:K2").Merge()
            ws2.Range("A1").Value = "기 성 내 역 서"
            ws2.Range("A1").Font.Size = 20
            ws2.Range("A1").Font.Bold = True
            ws2.Range("A1").HorizontalAlignment = -4108
            ws2.Range("A1").VerticalAlignment = -4108
            
            headers = ["공종/품명", "규격", "단위", "도급수량", "도급금액", "전회수량", "전회금액", "금회수량", "금회금액", "누계금액", "잔여금액"]
            for col, h in enumerate(headers, 1):
                ws2.Cells(4, col).Value = h
                ws2.Cells(4, col).Interior.Color = 14277081
                ws2.Cells(4, col).Font.Bold = True
                ws2.Cells(4, col).HorizontalAlignment = -4108
                ws2.Cells(4, col).Borders.LineStyle = 1
                
            ws2.Columns(1).ColumnWidth = 15
            ws2.Columns(2).ColumnWidth = 12
            ws2.Columns(3).ColumnWidth = 8
            for c_idx in range(4, 12): ws2.Columns(c_idx).ColumnWidth = 14
            
            r_idx = 5
            for t in ["RT", "UT", "PT"]:
                ws2.Cells(r_idx, 1).Value = t
                ws2.Cells(r_idx, 2).Value = "식" if t == "RT" else "m"
                ws2.Cells(r_idx, 3).Value = "식" if t == "RT" else "m"
                
                c_qty, c_amt = c.get(f"{t}_qty", 0), c.get(f"{t}_amt", 0)
                p_qty, p_amt = prev.get(f"{t}_qty", 0), prev.get(f"{t}_amt", 0)
                u_qty, u_amt = cur.get(f"{t}_qty", 0), cur.get(f"{t}_amt", 0)
                a_qty, a_amt = p_qty + u_qty, p_amt + u_amt
                
                ws2.Cells(r_idx, 4).Value = c_qty
                ws2.Cells(r_idx, 5).Value = c_amt
                ws2.Cells(r_idx, 6).Value = p_qty
                ws2.Cells(r_idx, 7).Value = p_amt
                ws2.Cells(r_idx, 8).Value = u_qty
                ws2.Cells(r_idx, 9).Value = u_amt
                ws2.Cells(r_idx, 10).Value = a_amt
                ws2.Cells(r_idx, 11).Value = c_amt - a_amt
                
                for col in range(1, 12):
                    cell = ws2.Cells(r_idx, col)
                    cell.Borders.LineStyle = 1
                    if col >= 4: cell.NumberFormat = "#,##0.0" if col in [4,6,8] else "#,##0"
                r_idx += 1
                
            # 합계 행
            ws2.Range(ws2.Cells(r_idx, 1), ws2.Cells(r_idx, 4)).Merge()
            ws2.Cells(r_idx, 1).Value = "합 계"
            ws2.Cells(r_idx, 1).HorizontalAlignment = -4108
            ws2.Cells(r_idx, 1).Font.Bold = True
            
            ws2.Cells(r_idx, 5).Value = c['total_amt']
            ws2.Cells(r_idx, 7).Value = prev['total_amt']
            ws2.Cells(r_idx, 9).Value = cur['total_amt']
            ws2.Cells(r_idx, 10).Value = total_accum
            ws2.Cells(r_idx, 11).Value = c['total_amt'] - total_accum
            
            for col in range(1, 12):
                cell = ws2.Cells(r_idx, col)
                cell.Borders.LineStyle = 1
                cell.Font.Bold = True
                cell.Interior.Color = 13434879
                if col >= 5: cell.NumberFormat = "#,##0"
                
            filepath = filepath.replace("/", "\\")
            wb.SaveAs(filepath)
            wb.Close()
            excel.Quit()
            
            messagebox.showinfo("완료", f"기성서류 엑셀이 생성되었습니다.\n{filepath}")
            os.startfile(filepath)
            
        except Exception as e:
            messagebox.showerror("오류", f"엑셀 생성 중 오류가 발생했습니다: {e}")
            try: excel.Quit()
            except: pass

if __name__ == "__main__":
    app = NDTCalculator()
    app.mainloop()

