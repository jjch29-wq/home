import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import os
import json
import win32com.client as win32

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(SCRIPT_DIR, "config.json")

DEFAULT_CONFIG = {
    "MATERIAL_COST": {
        'RT (B필름: 3⅓"x17")': 3379,
        'RT (A필름: 3⅓"x12")': 2540,
        'RT (A/2필름: 3⅓"x6")': 1515,
        "UT": 1115,
        "PT": 3971
    },
    "LABOR_COST": {
        "수송배관(주배관)": {
            "일반": {"RT": 45126, "UT": 43734, "PT": 38466},
            "야간": {"RT": 67689, "UT": 65601, "PT": 53457},
            "휴일": {"RT": 67689, "UT": 65601, "PT": 53429}
        },
        "플랜트(관리소)": {
            "일반": {"RT": 40438, "UT": 43734, "PT": 49220},
            "야간": {"RT": 60657, "UT": 65601, "PT": 70142},
            "휴일": {"RT": 60657, "UT": 65601, "PT": 68720}
        }
    },
    "CONTRACT_QTY": {
        "수송배관(주배관)": {
            "RT_B": 19125,
            "RT_A": 0,
            "RT_A2": 0,
            "UT": 319.02,
            "PT": 319.01
        },
        "플랜트(관리소)": {
            "RT_B": 1243,
            "RT_A": 2464,
            "RT_A2": 1704,
            "UT": 0,
            "PT": 19.62
        }
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

# 글로벌 변수
CONFIG = load_config()
MATERIAL_COST = CONFIG["MATERIAL_COST"]
LABOR_COST = CONFIG["LABOR_COST"]
CONTRACT_QTY = CONFIG.get("CONTRACT_QTY", DEFAULT_CONFIG["CONTRACT_QTY"])

class NDTCalculator(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("비파괴검사 기성 산출 계산기 (가산~가평)")
        self.geometry("1150x800")  
        self.configure(padx=10, pady=10)
        self.style = ttk.Style(self)
        self.style.theme_use("clam")
        
        self.records = [] # 저장된 기록 목록
        
        self.create_menu()
        self.create_widgets()
        
    def create_menu(self):
        menubar = tk.Menu(self)
        
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="작업 불러오기 (Load)", command=self.load_project)
        file_menu.add_command(label="작업 저장하기 (Save)", command=self.save_project)
        file_menu.add_separator()
        file_menu.add_command(label="계약 현황 확인 (Contract Status)", command=self.show_contract_status)
        file_menu.add_command(label="단가 설정 (Settings)", command=self.open_settings)
        file_menu.add_separator()
        file_menu.add_command(label="종료", command=self.quit)
        
        menubar.add_cascade(label="파일 (File)", menu=file_menu)
        self.config(menu=menubar)

    def create_widgets(self):
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        tab_work = ttk.Frame(self.notebook)
        self.notebook.add(tab_work, text="1. 일일 작업 기록 및 목록")
        
        tab_billing = ttk.Frame(self.notebook)
        self.notebook.add(tab_billing, text="2. 기성 계약관리")
        # --- TAB 1: WORK (입력 폼 및 목록 사이드바이사이드) ---
        work_pane = ttk.PanedWindow(tab_work, orient=tk.HORIZONTAL)
        work_pane.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        left_frame = ttk.Frame(work_pane)
        work_pane.add(left_frame, weight=1)
        
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
            
        ttk.Label(left_frame, text="2. 작업 구분 (구간 및 시간대)", font=("Arial", 11, "bold")).pack(anchor=tk.W, pady=(10, 5))
        
        self.loc_type_var = tk.StringVar(value="수송배관(주배관)")
        self.work_time_var = tk.StringVar(value="일반")
        
        type_time_frame = ttk.Frame(left_frame)
        type_time_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(type_time_frame, text="구간:").pack(side=tk.LEFT)
        for t in ["수송배관(주배관)", "플랜트(관리소)"]:
            ttk.Radiobutton(type_time_frame, text=t, value=t, variable=self.loc_type_var).pack(side=tk.LEFT, padx=5)
            
        ttk.Label(type_time_frame, text="  |  시간:").pack(side=tk.LEFT)
        for t in ["일반", "야간", "휴일"]:
            ttk.Radiobutton(type_time_frame, text=t, value=t, variable=self.work_time_var).pack(side=tk.LEFT, padx=5)

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
        
        # --- TAB 2: BILLING (계약 및 실비 정산) ---
        billing_container = ttk.Frame(tab_billing)
        billing_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        contract_frame = ttk.LabelFrame(billing_container, text="항목별 계약 및 전회 기성 (세액 미포함)", padding=10)
        contract_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        ttk.Button(contract_frame, text="프로젝트 총 계약수량 자동입력", command=self.auto_load_contract_qty).pack(fill=tk.X, pady=(0, 10))
        
        def format_currency(*args, var=None):
            try:
                val = var.get().replace(',', '')
                if val:
                    formatted = f"{int(val):,}"
                    if var.get() != formatted:
                        var.set(formatted)
            except ValueError:
                pass

        def get_int(var):
            try: return int(str(var.get()).replace(',', ''))
            except: return 0

        def get_float(var):
            try: return float(str(var.get()).replace(',', ''))
            except: return 0.0

        def format_qty(*args, var=None):
            try:
                val = var.get().replace(',', '')
                if not val or val == '.' or val.endswith('.'): return
                if '.' not in val:
                    formatted = f"{int(val):,}"
                    if var.get() != formatted:
                        var.set(formatted)
            except ValueError:
                pass

        self.get_int = get_int
        self.get_float = get_float
        
        self.contract_vars = {}
        default_qtys = {"RT": "24,536", "UT": "319.02", "PT": "338.63"}
        
        for idx, t in enumerate(["RT", "UT", "PT"]):
            f = ttk.Frame(contract_frame)
            f.pack(fill=tk.X, pady=2)
            ttk.Label(f, text=f"[{t}]", width=8).grid(row=0, column=0, rowspan=2, sticky=tk.W)
            
            c_qty = tk.StringVar(value=default_qtys[t])
            c_price = tk.IntVar(value=0)
            c_var = tk.StringVar(value="0")
            p_qty = tk.StringVar(value="0")
            p_var = tk.StringVar(value="0")
            
            c_qty.trace_add("write", lambda *a, v=c_qty: format_qty(var=v))
            p_qty.trace_add("write", lambda *a, v=p_qty: format_qty(var=v))
            
            unit = "매" if t == "RT" else "M"
            ttk.Label(f, text="계약 물량:").grid(row=0, column=1, sticky=tk.W)
            ttk.Entry(f, textvariable=c_qty, width=10).grid(row=0, column=2, padx=2)
            ttk.Label(f, text=unit).grid(row=0, column=3, sticky=tk.W, padx=(0, 15))
            
            ttk.Label(f, text="전회 물량:").grid(row=1, column=1, sticky=tk.W, pady=2)
            ttk.Entry(f, textvariable=p_qty, width=10).grid(row=1, column=2, padx=2)
            ttk.Label(f, text=unit).grid(row=1, column=3, sticky=tk.W)
            ttk.Separator(contract_frame, orient='horizontal').pack(fill=tk.X, pady=3)
            
            self.contract_vars[t] = {"c_qty": c_qty, "c_price": c_price, "contract": c_var, "p_qty": p_qty, "prev": p_var}

        f = ttk.Frame(contract_frame)
        f.pack(fill=tk.X, pady=2)
        ttk.Label(f, text="[프로젝트 총액]", width=12, font=("Arial", 9, "bold")).grid(row=0, column=0, rowspan=2, sticky=tk.W)
        ttk.Label(f, text="계약 총액:").grid(row=0, column=1, sticky=tk.W)
        self.total_contract_var = tk.StringVar(value="2,628,702,818")
        self.total_contract_var.trace_add("write", lambda *a, v=self.total_contract_var: format_currency(var=v))
        ttk.Entry(f, textvariable=self.total_contract_var, width=15).grid(row=0, column=2, padx=2)
        ttk.Label(f, text="원").grid(row=0, column=3)
        
        ttk.Label(f, text="전회 총액:").grid(row=1, column=1, sticky=tk.W, pady=2)
        self.total_prev_var = tk.StringVar(value="0")
        self.total_prev_var.trace_add("write", lambda *a, v=self.total_prev_var: format_currency(var=v))
        ttk.Entry(f, textvariable=self.total_prev_var, width=15).grid(row=1, column=2, padx=2)
        ttk.Label(f, text="원").grid(row=1, column=3)
        
        exp_frame = ttk.LabelFrame(billing_container, text="기타 경비 및 실비 정산 (월간)", padding=10)
        exp_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.exp_budget_var = tk.StringVar(value="72,215,000")
        self.exp_prev_var = tk.StringVar(value="0")
        self.exp_rem_var = tk.StringVar(value="72,215,000")
        
        self.exp_budget_var.trace_add("write", lambda *a, v=self.exp_budget_var: format_currency(var=v))
        self.exp_prev_var.trace_add("write", lambda *a, v=self.exp_prev_var: format_currency(var=v))
        
        bf = ttk.Frame(exp_frame)
        bf.pack(fill=tk.X, pady=2)
        ttk.Label(bf, text="경비 총 예산:", font=("Arial", 9, "bold")).pack(side=tk.LEFT)
        ttk.Entry(bf, textvariable=self.exp_budget_var, width=12).pack(side=tk.LEFT, padx=5)
        ttk.Label(bf, text="원").pack(side=tk.LEFT)
        
        ttk.Label(bf, text="전회 청구액:", font=("Arial", 9, "bold")).pack(side=tk.LEFT, padx=(10, 0))
        ttk.Entry(bf, textvariable=self.exp_prev_var, width=12).pack(side=tk.LEFT, padx=5)
        ttk.Label(bf, text="원").pack(side=tk.LEFT)
        
        rf = ttk.Frame(exp_frame)
        rf.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(rf, text="▶ 경비 잔액 (예산 - 전회 - 금회): ", font=("Arial", 9, "bold")).pack(side=tk.LEFT)
        self.lbl_exp_rem = ttk.Label(rf, textvariable=self.exp_rem_var, font=("Arial", 9, "bold"), foreground="blue")
        self.lbl_exp_rem.pack(side=tk.LEFT)
        ttk.Label(rf, text=" 원", font=("Arial", 9, "bold")).pack(side=tk.LEFT)
        
        ttk.Separator(exp_frame, orient='horizontal').pack(fill=tk.X, pady=5)
        ttk.Label(exp_frame, text="▼ 금회 청구액 (세액 미포함 금액)").pack(anchor=tk.W, pady=(0, 10))
        
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
        
        def update_exp_rem(*args):
            try:
                budget = self.get_int(self.exp_budget_var)
                prev = self.get_int(self.exp_prev_var)
                curr = sum([self.equip_cost_var.get(), self.safety_cost_var.get(), self.travel_cost_var.get(), self.print_cost_var.get()])
                rem = budget - prev - curr
                self.exp_rem_var.set(f"{rem:,}")
                if rem < 0:
                    self.lbl_exp_rem.config(foreground="red")
                else:
                    self.lbl_exp_rem.config(foreground="blue")
            except:
                pass
                
        self.exp_budget_var.trace_add("write", update_exp_rem)
        self.exp_prev_var.trace_add("write", update_exp_rem)
        self.equip_cost_var.trace_add("write", update_exp_rem)
        self.safety_cost_var.trace_add("write", update_exp_rem)
        self.travel_cost_var.trace_add("write", update_exp_rem)
        self.print_cost_var.trace_add("write", update_exp_rem)
        update_exp_rem()
        
        # --- RIGHT FRAME (누적 테이블, TAB 1에 배치) ---
        bottom_frame = ttk.Frame(work_pane)
        work_pane.add(bottom_frame, weight=2)
        
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
        
        tree_scroll = ttk.Scrollbar(bottom_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)
        
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
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
        
        loc_type = getattr(self, "loc_type_var", None)
        loc_type_val = loc_type.get() if loc_type else "수송배관(주배관)"
        
        if loc_type_val in LABOR_COST:
            lab_unit_cost = LABOR_COST[loc_type_val][work_time][ndt_type]
        else:
            lab_unit_cost = LABOR_COST.get(work_time, {}).get(ndt_type, 0)
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
            loc_type = getattr(self, "loc_type_var", None)
            loc_type_val = loc_type.get() if loc_type else "수송배관(주배관)"
            if loc_type_val in LABOR_COST:
                lab_unit = LABOR_COST[loc_type_val][res['work_time']][res['ndt_type']]
            else:
                lab_unit = LABOR_COST.get(res['work_time'], {}).get(res['ndt_type'], 0)
            
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
                "contract": {
                    t: {
                        "c_qty": self.get_float(v["c_qty"]),
                        "contract": self.get_int(v["contract"]),
                        "p_qty": self.get_float(v["p_qty"]),
                        "prev": self.get_int(v["prev"])
                    } for t, v in self.contract_vars.items()
                },
                "expenses": {
                    "equip": self.equip_cost_var.get(),
                    "safety": self.safety_cost_var.get(),
                    "travel": self.travel_cost_var.get(),
                    "print": self.print_cost_var.get(),
                    "budget": self.get_int(self.exp_budget_var),
                    "prev": self.get_int(self.exp_prev_var)
                }
            }
            data["total_amt"] = {"contract": self.get_int(self.total_contract_var), "prev": self.get_int(self.total_prev_var)}
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
            messagebox.showinfo("저장 완료", "작업이 성공적으로 저장되었습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"저장 중 오류가 발생했습니다: {e}")

    def auto_load_contract_qty(self):
        pipe = CONTRACT_QTY.get("수송배관(주배관)", {})
        plant = CONTRACT_QTY.get("플랜트(관리소)", {})
        
        rt_total = (pipe.get("RT_B", 0) + pipe.get("RT_A", 0) + pipe.get("RT_A2", 0) +
                    plant.get("RT_B", 0) + plant.get("RT_A", 0) + plant.get("RT_A2", 0))
        ut_total = pipe.get("UT", 0) + plant.get("UT", 0)
        pt_total = pipe.get("PT", 0) + plant.get("PT", 0)
        
        self.contract_vars["RT"]["c_qty"].set(f"{int(rt_total):,}")
        self.contract_vars["UT"]["c_qty"].set(f"{ut_total:,.2f}")
        self.contract_vars["PT"]["c_qty"].set(f"{pt_total:,.2f}")
        
        messagebox.showinfo("불러오기 완료", "프로젝트 전체 총 계약 물량이 자동으로 입력되었습니다.\n\nRT 항목의 금액은 엑셀 내역서에서 확인 후 직접 입력하시고, UT와 PT는 단가를 입력하시면 자동으로 금액이 산출됩니다.")

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
            
            cont = data.get("contract", {})
            for t, v in self.contract_vars.items():
                if t in cont:
                    cq = cont[t].get("c_qty", 0.0)
                    pq = cont[t].get("p_qty", 0.0)
                    v["c_qty"].set(f"{int(cq):,}" if cq.is_integer() else f"{cq:,.2f}")
                    v["contract"].set(f"{cont[t].get('contract', 0):,}")
                    v["p_qty"].set(f"{int(pq):,}" if pq.is_integer() else f"{pq:,.2f}")
                    v["prev"].set(f"{cont[t].get('prev', 0):,}")
            
            ex = data.get("expenses", {})
            self.equip_cost_var.set(ex.get("equip", 0))
            self.safety_cost_var.set(ex.get("safety", 0))
            self.travel_cost_var.set(ex.get("travel", 0))
            self.print_cost_var.set(ex.get("print", 0))
            self.exp_budget_var.set(f"{ex.get('budget', 72215000):,}")
            self.exp_prev_var.set(f"{ex.get('prev', 0):,}")
            
            tot = data.get("total_amt", {})
            self.total_contract_var.set(f"{tot.get('contract', 2628702818):,}")
            self.total_prev_var.set(f"{tot.get('prev', 0):,}")
            
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
            ws.Range("A4").Font.Size = 12
            ws.Range("C4:J4").Merge()
            ws.Range("C4").Value = "가산~가평 천연가스 공급시설 건설공사 비파괴검사 기술용역"
            ws.Range("C4").Font.Bold = True
            ws.Range("C4").Font.Size = 14
            
            ws.Range("L4:M4").Merge()
            ws.Range("L4").Value = "청구일자 :"
            ws.Range("L4").Font.Bold = True
            ws.Range("L4").Font.Size = 11
            ws.Range("L4").HorizontalAlignment = -4152
            ws.Range("N4:O4").Merge()
            ws.Range("N4").Value = datetime.now().strftime('%Y년 %m월 %d일')
            ws.Range("N4").Font.Size = 11
            
            # --- 기성 요약 테이블 ---
            ws.Cells(6, 1).Value = "구분"
            ws.Range(ws.Cells(6, 1), ws.Cells(7, 1)).Merge()
            
            groups = ["계약", "전회까지기성", "금회기성", "누계기성", "잔액"]
            for idx, g_name in enumerate(groups):
                start_col = 2 + idx * 2
                ws.Cells(6, start_col).Value = g_name
                ws.Range(ws.Cells(6, start_col), ws.Cells(6, start_col + 1)).Merge()
                ws.Cells(7, start_col).Value = "수량"
                ws.Cells(7, start_col + 1).Value = "금액"
                
            for c in range(1, 12):
                ws.Cells(6, c).Font.Bold = True
                ws.Cells(7, c).Font.Bold = True
                ws.Cells(6, c).HorizontalAlignment = -4108
                ws.Cells(7, c).HorizontalAlignment = -4108
                ws.Cells(6, c).Interior.Color = 14277081
                ws.Cells(7, c).Interior.Color = 14277081
            
            ws.Range(ws.Cells(6, 1), ws.Cells(12, 11)).Borders.LineStyle = 1
            
            extra_items_total = sum([self.equip_cost_var.get(), self.safety_cost_var.get(), self.travel_cost_var.get(), self.print_cost_var.get()])
            
            categories = ["RT", "UT", "PT", "기타실비", "총 계"]
            for i, cat in enumerate(categories):
                row = 8 + i
                ws.Cells(row, 1).Value = cat
                ws.Cells(row, 1).HorizontalAlignment = -4108
                if cat == "총 계":
                    ws.Range(ws.Cells(row, 1), ws.Cells(row, 11)).Interior.Color = 15987699
                    ws.Cells(row, 1).Font.Bold = True
                
                if cat == "총 계":
                    c_qty, p_qty, cur_qty, tot_qty, rem_qty = "", "", "", "", ""
                    c_amt = self.get_int(self.total_contract_var)
                    p_amt = self.get_int(self.total_prev_var)
                    cur_amt = sum(r["subtotal"] for r in self.records) + extra_items_total
                    tot_amt = p_amt + cur_amt
                    rem_amt = c_amt - tot_amt
                elif cat == "기타실비":
                    c_qty, p_qty, cur_qty, tot_qty, rem_qty = "", "", "", "", ""
                    c_amt = self.get_int(self.exp_budget_var)
                    p_amt = self.get_int(self.exp_prev_var)
                    cur_amt = extra_items_total
                    tot_amt = p_amt + cur_amt
                    rem_amt = c_amt - tot_amt
                else:
                    c_qty = self.get_float(self.contract_vars[cat]["c_qty"])
                    p_qty = self.get_float(self.contract_vars[cat]["p_qty"])
                    cur_qty = sum(r["adjusted_qty"] for r in self.records if r["ndt_type"] == cat)
                    tot_qty = p_qty + cur_qty
                    rem_qty = c_qty - tot_qty
                    c_amt = p_amt = tot_amt = rem_amt = ""
                    cur_amt = sum(r["subtotal"] for r in self.records if r["ndt_type"] == cat)
                
                vals = [c_qty, c_amt, p_qty, p_amt, cur_qty, cur_amt, tot_qty, tot_amt, rem_qty, rem_amt]
                for col_idx, val in enumerate(vals, start=2):
                    v_cell = ws.Cells(row, col_idx)
                    if val == "" or val == 0 or val == 0.0:
                        v_cell.Value = "-"
                        v_cell.HorizontalAlignment = -4108
                    else:
                        v_cell.Value = val
                        try:
                            if isinstance(val, (int, float)):
                                if val == int(val):
                                    v_cell.NumberFormat = "#,##0"
                                else:
                                    v_cell.NumberFormat = "#,##0.00"
                        except:
                            pass
                        
                    if cat == "총 계":
                        v_cell.Font.Bold = True

            # --- 세부 내역 테이블 ---
            headers = ["No.", "검사일자", "작업구간", "검사종류", "규격/자재", "근무형태", "실물량", "단위", "보정계수", "환산물량", 
                       "재료비", "직접인건비", "제경비", "기술료", "공급가액소계"]
            
            start_row = 15
            for col, h in enumerate(headers, start=1):
                cell = ws.Cells(start_row, col)
                cell.Value = h
                cell.Font.Bold = True
                cell.Interior.Color = 14277081
                cell.HorizontalAlignment = -4108
                cell.Borders.LineStyle = 1
            
            ws.Columns(1).ColumnWidth = 9
            ws.Columns(2).ColumnWidth = 11
            ws.Columns(3).ColumnWidth = 16
            ws.Columns(4).ColumnWidth = 11
            ws.Columns(5).ColumnWidth = 24
            ws.Columns(6).ColumnWidth = 11
            ws.Columns(7).ColumnWidth = 14
            ws.Columns(8).ColumnWidth = 9
            ws.Columns(9).ColumnWidth = 14
            ws.Columns(10).ColumnWidth = 11
            ws.Columns(11).ColumnWidth = 16
            ws.Columns(12).ColumnWidth = 14
            ws.Columns(13).ColumnWidth = 14
            ws.Columns(14).ColumnWidth = 14
            ws.Columns(15).ColumnWidth = 16
            
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

    def show_contract_status(self):
        win = tk.Toplevel(self)
        win.title("가산~가평 총 계약 수량")
        win.geometry("500x350")
        
        main_frame = ttk.Frame(win, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="📄 프로젝트 총 계약 수량 (실검사 물량 기준)", font=("Arial", 13, "bold")).pack(pady=(0, 15))
        
        tree = ttk.Treeview(main_frame, columns=("Type", "Pipeline", "Plant", "Total"), show="headings", height=8)
        tree.heading("Type", text="검사 종류")
        tree.heading("Pipeline", text="수송배관 (주배관)")
        tree.heading("Plant", text="플랜트 (관리소)")
        tree.heading("Total", text="총계")
        
        tree.column("Type", width=120, anchor=tk.CENTER)
        tree.column("Pipeline", width=100, anchor=tk.E)
        tree.column("Plant", width=100, anchor=tk.E)
        tree.column("Total", width=100, anchor=tk.E)
        
        pipe = CONTRACT_QTY.get("수송배관(주배관)", {})
        plant = CONTRACT_QTY.get("플랜트(관리소)", {})
        
        items = [
            ("RT (B필름: 17\")", pipe.get("RT_B", 0), plant.get("RT_B", 0), "매"),
            ("RT (A필름: 12\")", pipe.get("RT_A", 0), plant.get("RT_A", 0), "매"),
            ("RT (A/2필름: 6\")", pipe.get("RT_A2", 0), plant.get("RT_A2", 0), "매"),
            ("UT", pipe.get("UT", 0), plant.get("UT", 0), "M"),
            ("PT", pipe.get("PT", 0), plant.get("PT", 0), "M")
        ]
        
        for name, p_val, pl_val, unit in items:
            total = round(p_val + pl_val, 2)
            tree.insert("", tk.END, values=(
                name,
                f"{p_val:,.2f}".rstrip('0').rstrip('.') + f" {unit}",
                f"{pl_val:,.2f}".rstrip('0').rstrip('.') + f" {unit}",
                f"{total:,.2f}".rstrip('0').rstrip('.') + f" {unit}"
            ))
            
        tree.pack(fill=tk.BOTH, expand=True)
        ttk.Button(main_frame, text="닫기", command=win.destroy).pack(pady=15)

    def open_settings(self):
        settings_win = tk.Toplevel(self)
        settings_win.title("단가 설정")
        settings_win.geometry("450x650")
        
        main_frame = ttk.Frame(settings_win)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        canvas = tk.Canvas(main_frame)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        scrollbar.pack(side="right", fill="y")
        
        self.setting_entries = {}
        
        row = 0
        ttk.Label(scrollable_frame, text="[ 재료비 단가 (원) ]", font=("Arial", 12, "bold")).grid(row=row, column=0, columnspan=2, pady=10, sticky="w")
        row += 1
        for key, val in MATERIAL_COST.items():
            ttk.Label(scrollable_frame, text=key).grid(row=row, column=0, padx=10, pady=2, sticky="w")
            var = tk.StringVar(value=str(val))
            ttk.Entry(scrollable_frame, textvariable=var, width=15).grid(row=row, column=1, padx=10, pady=2)
            self.setting_entries[("MATERIAL_COST", key)] = var
            row += 1
            
        ttk.Label(scrollable_frame, text="[ 인건비 단가 (원) ]", font=("Arial", 12, "bold")).grid(row=row, column=0, columnspan=2, pady=(20, 10), sticky="w")
        row += 1
        for loc_type, times in LABOR_COST.items():
            ttk.Label(scrollable_frame, text=f"- {loc_type}", font=("Arial", 10, "bold")).grid(row=row, column=0, columnspan=2, pady=5, sticky="w")
            row += 1
            for work_time, ndts in times.items():
                for ndt_type, val in ndts.items():
                    ttk.Label(scrollable_frame, text=f"{work_time}검사 - {ndt_type}").grid(row=row, column=0, padx=30, pady=2, sticky="w")
                    var = tk.StringVar(value=str(val))
                    ttk.Entry(scrollable_frame, textvariable=var, width=15).grid(row=row, column=1, padx=10, pady=2)
                    self.setting_entries[("LABOR_COST", loc_type, work_time, ndt_type)] = var
                    row += 1
                    
        def save_and_close():
            try:
                for key_tuple, var in self.setting_entries.items():
                    if key_tuple[0] == "MATERIAL_COST":
                        MATERIAL_COST[key_tuple[1]] = int(var.get().replace(",", ""))
                    elif key_tuple[0] == "LABOR_COST":
                        LABOR_COST[key_tuple[1]][key_tuple[2]][key_tuple[3]] = int(var.get().replace(",", ""))
                
                save_config({"MATERIAL_COST": MATERIAL_COST, "LABOR_COST": LABOR_COST})
                messagebox.showinfo("저장 완료", "단가 설정이 파일(config.json)에 정상적으로 저장되었습니다.\n변경된 단가는 다음 계산부터 바로 적용됩니다.", parent=settings_win)
                settings_win.destroy()
            except ValueError:
                messagebox.showerror("오류", "모든 단가는 숫자 형식이어야 합니다.", parent=settings_win)
                
        ttk.Button(settings_win, text="저장 및 닫기", command=save_and_close).pack(pady=10)

if __name__ == "__main__":
    app = NDTCalculator()
    app.mainloop()
