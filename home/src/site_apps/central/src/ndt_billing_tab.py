import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import os
import json
import win32com.client as win32
from tkcalendar import DateEntry
from site_apps.central.src.services.ndt_calculator import calculate_billing

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(SCRIPT_DIR, "config.json")

DEFAULT_CONFIG = {
    "MATERIAL_COST": {
        "PAUT_300A 이상": 37559,
        "PAUT_250A": 37559,
        "PAUT_200A": 37559,
        "PAUT_150A-125A": 38229,
        "PAUT_100A 이하": 38229,
        "RT_3 1/3 x 12\"": 9274,
        "RT_3 1/3 x 6\"": 8044,
        "MT_MT": 411,
        "PT_PT": 1177
    },
    "LABOR_COST": {
        "열배관": {
            "일반": {
                "PAUT_300A 이상": 65200, "PAUT_250A": 76163, "PAUT_200A": 87348, "PAUT_150A-125A": 98560, "PAUT_100A 이하": 92362,
                "RT_3 1/3 x 12\"": 51885, "RT_3 1/3 x 6\"": 51885,
                "MT_MT": 29771, "PT_PT": 30847
            },
            "야간": {
                "PAUT_300A 이상": 97800, "PAUT_250A": 114243, "PAUT_200A": 131021, "PAUT_150A-125A": 147841, "PAUT_100A 이하": 138544,
                "RT_3 1/3 x 12\"": 77833, "RT_3 1/3 x 6\"": 77833,
                "MT_MT": 44659, "PT_PT": 46272
            }
        }
    },
    "CONTRACT_QTY": {
        "열배관": {
            "일반": {
                "PAUT_300A 이상": 129,
                "PAUT_250A": 4,
                "PAUT_200A": 4,
                "PAUT_150A-125A": 1,
                "PAUT_100A 이하": 1,
                "RT_3 1/3 x 12\"": 293,
                "RT_3 1/3 x 6\"": 105,
                "MT_MT": 26,
                "PT_PT": 26
            },
            "야간": {
                "PAUT_300A 이상": 624,
                "PAUT_250A": 1,
                "PAUT_200A": 2,
                "PAUT_150A-125A": 1,
                "PAUT_100A 이하": 1,
                "RT_3 1/3 x 12\"": 43,
                "RT_3 1/3 x 6\"": 49,
                "MT_MT": 1,
                "PT_PT": 1
            }
        }
    },
}

def load_config():
    if not os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=4)
        return DEFAULT_CONFIG
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            config = json.load(f)
            # 2026 신규 단가계약 스키마가 없는 경우(과거 config.json) DEFAULT로 덮어쓰기
            if "PAUT_300A 이상" not in config.get("MATERIAL_COST", {}):
                with open(CONFIG_FILE, 'w', encoding='utf-8') as fw:
                    json.dump(DEFAULT_CONFIG, fw, ensure_ascii=False, indent=4)
                return DEFAULT_CONFIG
                
            # Migrate old config to district heating schema
            if "열배관" in config.get("LABOR_COST", {}):
                config["LABOR_COST"]["열배관"] = config["LABOR_COST"].pop("열배관")
                config["LABOR_COST"].pop("플랜트(관리소)", None)
                with open(CONFIG_FILE, 'w', encoding='utf-8') as fw:
                    json.dump(config, fw, ensure_ascii=False, indent=4)
                    
            return config
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

class NDTCalculatorTab(ttk.Frame):
    def __init__(self, parent, main_app=None):
        super().__init__(parent)
        self.main_app = main_app
        # self.title("비파괴검사 기성 산출 계산기 (가산~가평)")
        # self.geometry("1150x800")  
        self.configure(padding=10)
        self.style = ttk.Style()
        self.style.theme_use("clam")
        
        self.records = [] # 저장된 기록 목록
        
        # self.create_menu()
        self.create_widgets()
        # self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.load_billing_records()
        
    def load_billing_records(self):
        global CONFIG
        self.records = CONFIG.get("BILLING_RECORDS", [])
        
        if hasattr(self, 'tree'):
            for item in self.tree.get_children():
                self.tree.delete(item)
                
            for res in self.records:
                self.tree.insert("", tk.END, values=(
                    res.get("date", ""), res.get("loc", ""), res.get("ndt_type", ""), res.get("work_time", ""), 
                    res.get("material_type", ""), f"{res.get('qty', 0):.1f}", res.get("unit", ""),
                    f"{res.get('corr', 1):.2f}", f"{res.get('adjusted_qty', 0):.2f}", 
                    f"{res.get('mat_cost', 0):,}", f"{res.get('lab_cost', 0):,}",
                    f"{res.get('overhead', 0):,}", f"{res.get('tech', 0):,}", f"{res.get('subtotal', 0):,}"
                ))
            if hasattr(self, 'update_qty_summary'):
                self.update_qty_summary()

    def save_billing_records(self):
        try:
            from config_manager import save_config
            global CONFIG
            CONFIG["BILLING_RECORDS"] = self.records
            save_config(CONFIG)
        except Exception as e:
            print(f"Error saving billing records: {e}")

    def save_ui_state(self):
        try:
            widths = {}
            for col in self.tree["columns"]:
                widths[col] = self.tree.column(col, "width")
            global CONFIG
            CONFIG["TREE_WIDTHS"] = widths
            
            if hasattr(self, 'work_pane'):
                try:
                    # tk.PanedWindow는 sash_coord(n)로 sash 위치를 반환
                    sash_x, sash_y = self.work_pane.sash_coord(0)
                    CONFIG["SASH_POS"] = int(sash_x)
                except:
                    pass
                
            save_config(CONFIG)
        except:
            pass
        
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
        # self.config(menu=menubar)

    def create_widgets(self):
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        tab_work = ttk.Frame(self.notebook)
        self.notebook.add(tab_work, text="1. 일일 작업 기록 및 목록")
        
        tab_billing = ttk.Frame(self.notebook)
        self.notebook.add(tab_billing, text="2. 기성 계약관리")
        # --- TAB 1: WORK (입력 폼 및 목록 사이드바이사이드) ---
        self.work_pane = tk.PanedWindow(tab_work, orient=tk.HORIZONTAL, sashrelief=tk.RAISED, sashwidth=5, bg="#b0b0b0")
        self.work_pane.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        left_outer_frame = ttk.Frame(self.work_pane)
        self.work_pane.add(left_outer_frame, stretch="always")
        
        left_canvas = tk.Canvas(left_outer_frame, highlightthickness=0)
        left_scroll = ttk.Scrollbar(left_outer_frame, orient="vertical", command=left_canvas.yview)
        left_canvas.configure(yscrollcommand=left_scroll.set)
        
        left_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        left_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        left_frame = ttk.Frame(left_canvas)
        left_window = left_canvas.create_window((0, 0), window=left_frame, anchor="nw")
        
        left_frame.bind("<Configure>", lambda e: left_canvas.configure(scrollregion=left_canvas.bbox("all")))
        left_canvas.bind("<Configure>", lambda e: left_canvas.itemconfig(left_window, width=e.width))
        
        def _on_mousewheel(event):
            try:
                # Scroll only when mouse is within the outer frame
                if str(event.widget).startswith(str(left_outer_frame)):
                    delta = event.delta
                    if event.num == 4: delta = 120
                    if event.num == 5: delta = -120
                    left_canvas.yview_scroll(int(-1*(delta/120)), "units")
            except:
                pass
                
        # Safe binding for mouse wheel
        left_outer_frame.bind_all("<MouseWheel>", _on_mousewheel, add='+')
        left_outer_frame.bind_all("<Button-4>", _on_mousewheel, add='+')
        left_outer_frame.bind_all("<Button-5>", _on_mousewheel, add='+')
        
        info_frame1 = ttk.Frame(left_frame)
        info_frame1.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(info_frame1, text="• 검사일자:", font=("Arial", 10, "bold")).pack(side=tk.LEFT)
        self.date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        self.date_entry = DateEntry(info_frame1, textvariable=self.date_var, width=13, date_pattern='yyyy-mm-dd', background='darkblue', foreground='white', borderwidth=2)
        self.date_entry.pack(side=tk.LEFT, padx=5)
        
        info_frame2 = ttk.Frame(left_frame)
        info_frame2.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(info_frame2, text="• 작업구간 (Joint No 등):", font=("Arial", 10, "bold")).pack(side=tk.LEFT)
        self.loc_var = tk.StringVar(value="")
        ttk.Entry(info_frame2, textvariable=self.loc_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        ttk.Label(left_frame, text="1. 검사 종류", font=("Arial", 11, "bold")).pack(anchor=tk.W, pady=(10, 5))
        self.ndt_type_var = tk.StringVar(value="PAUT")
        type_frame = ttk.Frame(left_frame)
        type_frame.pack(fill=tk.X, pady=5)
        for t in ["PAUT", "RT", "MT", "PT"]:
            ttk.Radiobutton(type_frame, text=t, value=t, variable=self.ndt_type_var, command=self.update_dynamic_ui).pack(side=tk.LEFT, padx=10)
            
        ttk.Label(left_frame, text="2. 작업 구분 (시간대)", font=("Arial", 11, "bold")).pack(anchor=tk.W, pady=(10, 5))
        
        self.loc_type_var = tk.StringVar(value="열배관")
        self.work_time_var = tk.StringVar(value="일반")
        
        type_time_frame2 = ttk.Frame(left_frame)
        type_time_frame2.pack(fill=tk.X, pady=2)
        ttk.Label(type_time_frame2, text="시간:").pack(side=tk.LEFT)
        for t in ["일반", "야간"]:
            ttk.Radiobutton(type_time_frame2, text=t, value=t, variable=self.work_time_var).pack(side=tk.LEFT, padx=5)

        self.material_lbl = ttk.Label(left_frame, text="3. 세부 규격 및 조건", font=("Arial", 11, "bold"))
        self.material_lbl.pack(anchor=tk.W, pady=(10, 5))
        self.material_var = tk.StringVar(value='RT (B필름: 3⅓"x17")')
        self.material_combo = ttk.Combobox(left_frame, textvariable=self.material_var, values=['RT (B필름: 3⅓"x17")', 'RT (A필름: 3⅓"x12")', 'RT (A/2필름: 3⅓"x6")'], state="readonly")
        self.material_combo.pack(fill=tk.X, pady=5)
        
        self.dynamic_frame = ttk.LabelFrame(left_frame, text="4. 보정계수 조건 선택", padding=10)
        self.dynamic_frame.pack(fill=tk.X, pady=(10, 5))
        
        self.source_frame = ttk.Frame(self.dynamic_frame)
        ttk.Label(self.source_frame, text="• 방사선원 :", width=14).pack(side=tk.LEFT)
        self.source_var = tk.StringVar(value="Se-75 (1.0)")
        self.source_combo = ttk.Combobox(self.source_frame, textvariable=self.source_var, state="readonly")
        self.source_combo['values'] = ["Ir-192 또는 Se-75 (1.0)", "X-ray 발생장치 (1.3)"]
        self.source_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        self.pipe_frame = ttk.Frame(self.dynamic_frame)
        ttk.Label(self.pipe_frame, text="• 관경(구경) :", width=14).pack(side=tk.LEFT)
        self.pipe_var = tk.StringVar()
        self.pipe_combo = ttk.Combobox(self.pipe_frame, textvariable=self.pipe_var, state="readonly")
        self.pipe_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        self.thickness_frame = ttk.Frame(self.dynamic_frame)
        ttk.Label(self.thickness_frame, text="• 투과/모재두께 :", width=14).pack(side=tk.LEFT)
        self.thickness_var = tk.StringVar()
        self.thickness_combo = ttk.Combobox(self.thickness_frame, textvariable=self.thickness_var, state="readonly")
        self.thickness_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Label(left_frame, text="5. 실검사 물량 (RT: 매 / UT,PT: Meter)", font=("Arial", 11, "bold")).pack(anchor=tk.W, pady=(10, 5))
        self.quantity_var = tk.DoubleVar(value=10.0)
        ttk.Entry(left_frame, textvariable=self.quantity_var).pack(fill=tk.X, pady=5)
        
        rate_outer_frame = ttk.Frame(left_frame)
        rate_outer_frame.pack(fill=tk.X, pady=(15, 5))
        ttk.Label(rate_outer_frame, text="6. 적용 요율 (%)", font=("Arial", 11, "bold")).pack(anchor=tk.W, pady=(0, 5))
        
        rate_frame = ttk.Frame(rate_outer_frame)
        rate_frame.pack(fill=tk.X)
        ttk.Label(rate_frame, text="제경비율:").pack(side=tk.LEFT)
        self.overhead_rate_var = tk.DoubleVar(value=110.0)
        ttk.Entry(rate_frame, textvariable=self.overhead_rate_var, width=8).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(rate_frame, text="기술료율:").pack(side=tk.LEFT, padx=(10, 0))
        self.tech_fee_rate_var = tk.DoubleVar(value=20.0)
        ttk.Entry(rate_frame, textvariable=self.tech_fee_rate_var, width=8).pack(side=tk.LEFT, padx=5)
        
        btn_frame = ttk.Frame(left_frame)
        btn_frame.pack(fill=tk.X, pady=15)
        
        ttk.Button(btn_frame, text="금액 계산하기", command=self.calculate).pack(side=tk.TOP, expand=True, fill=tk.X, pady=2, ipady=4)
        ttk.Button(btn_frame, text="기록 목록에 추가", command=self.add_to_record).pack(side=tk.TOP, expand=True, fill=tk.X, pady=2, ipady=4)
        
        ttk.Label(left_frame, text="[ 단일 계산 결과 ]", font=("Arial", 11, "bold")).pack(anchor=tk.W, pady=(0, 5))
        self.result_text = tk.Text(left_frame, height=10, width=25, state=tk.DISABLED, font=("Consolas", 11), wrap=tk.WORD)
        self.result_text.pack(fill=tk.X, expand=False)
        
        # --- TAB 2: BILLING (계약 및 실비 정산) ---
        billing_container = ttk.Frame(tab_billing)
        billing_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        round_frame = ttk.Frame(billing_container)
        round_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.round_var = tk.IntVar(value=1)
        ttk.Label(round_frame, text="기성 청구 회차: 제", font=("Arial", 11, "bold")).pack(side=tk.LEFT)
        ttk.Entry(round_frame, textvariable=self.round_var, width=5, justify="center", font=("Arial", 11, "bold")).pack(side=tk.LEFT, padx=5)
        ttk.Label(round_frame, text="회", font=("Arial", 11, "bold")).pack(side=tk.LEFT)
        
        ttk.Button(round_frame, text="다음 회차로 이월하기 (전회 누적 & 금회 초기화)", command=self.carry_over_round).pack(side=tk.RIGHT)
        ttk.Button(round_frame, text="이전 백업 불러오기 (.ndt)", command=self.load_project).pack(side=tk.RIGHT, padx=10)
        ttk.Button(round_frame, text="✨ 엑셀 보고서 생성기 열기", command=self.open_report_hub).pack(side=tk.RIGHT, padx=5)
        
        content_frame = ttk.Frame(billing_container)
        content_frame.pack(fill=tk.BOTH, expand=True)
        content_frame.columnconfigure(0, weight=2)  # Left frame takes 2/3
        content_frame.columnconfigure(1, weight=1)  # Right frame takes 1/3
        content_frame.rowconfigure(0, weight=1)
        
        contract_frame = ttk.LabelFrame(content_frame, text="항목별 계약 및 전회 기성 (세액 미포함)", padding=10)
        contract_frame.grid(row=0, column=0, sticky='nsew', padx=(0, 10))
        
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
        
        contract_canvas = tk.Canvas(contract_frame, highlightthickness=0, height=200)
        contract_scrollbar = ttk.Scrollbar(contract_frame, orient="vertical", command=contract_canvas.yview)
        contract_canvas.configure(yscrollcommand=contract_scrollbar.set)
        
        contract_canvas.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        contract_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        contract_inner_frame = ttk.Frame(contract_canvas)
        contract_window = contract_canvas.create_window((0, 0), window=contract_inner_frame, anchor="nw")
        
        contract_inner_frame.bind("<Configure>", lambda e: contract_canvas.configure(scrollregion=contract_canvas.bbox("all")))
        contract_canvas.bind("<Configure>", lambda e: contract_canvas.itemconfig(contract_window, width=e.width))
        
        self.contract_vars = {}
        
        headers = ["구간", "시간", "항목", "단위", "계약수량", "전회수량", "금회수량", "잔여수량", "계약단가", "계약금액", "전회금액"]
        for i, h in enumerate(headers):
            ttk.Label(contract_inner_frame, text=h, font=("Arial", 9, "bold"), anchor="center").grid(row=0, column=i, padx=5, pady=2, sticky="ew")
            
        locations = ["열배관"]
        times = ["일반", "야간"]
        materials = [
            ("PAUT_300A 이상", "PAUT 300A 이상"),
            ("PAUT_250A", "PAUT 250A"),
            ("PAUT_200A", "PAUT 200A"),
            ("PAUT_150A-125A", "PAUT 150A-125A"),
            ("PAUT_100A 이하", "PAUT 100A 이하"),
            ("RT_3 1/3 x 12\"", 'RT 3 1/3 x 12"'),
            ("RT_3 1/3 x 6\"", 'RT 3 1/3 x 6"'),
            ("MT_MT", "MT"),
            ("PT_PT", "PT")
        ]
        
        row_idx = 1
        for loc in locations:
            for t_time in times:
                for m_key, m_name in materials:
                    unit = "매" if m_key.startswith("RT") else "M"
                    full_key = f"{loc}_{t_time}_{m_key}"
                    
                    ttk.Label(contract_inner_frame, text=loc).grid(row=row_idx, column=0, sticky="w", padx=2)
                    ttk.Label(contract_inner_frame, text=t_time).grid(row=row_idx, column=1, sticky="w", padx=2)
                    ttk.Label(contract_inner_frame, text=m_name).grid(row=row_idx, column=2, sticky="w", padx=2)
                    ttk.Label(contract_inner_frame, text=unit).grid(row=row_idx, column=3, padx=2)
                    
                    c_qty = tk.StringVar(value="0")
                    p_qty = tk.StringVar(value="0")
                    curr_qty = tk.StringVar(value="0")
                    rem_qty = tk.StringVar(value="0")
                    
                    c_var = tk.StringVar(value="0")
                    p_var = tk.StringVar(value="0")
                    c_price_var = tk.StringVar(value="0")
                    
                    unit_cost = 0
                    try:
                        lab_unit = LABOR_COST[loc][t_time].get(m_key, 0)
                        mat_unit = MATERIAL_COST.get(m_key, 0)
                        
                        oh = int(lab_unit * float(self.overhead_rate_var.get()) / 100.0)
                        tech = int((lab_unit + oh) * float(self.tech_fee_rate_var.get()) / 100.0)
                        
                        unit_cost = mat_unit + lab_unit + oh + tech
                    except Exception as e:
                        print(e)
                        pass
                    c_qty.trace_add("write", lambda *a, v=c_qty: format_qty(var=v))
                    p_qty.trace_add("write", lambda *a, v=p_qty: format_qty(var=v))
                    
                    def update_rem_qty(*args, k=full_key):
                        try:
                            c = self.get_float(self.contract_vars[k]["c_qty"])
                            p = self.get_float(self.contract_vars[k]["p_qty"])
                            cur = self.get_float(self.contract_vars[k]["curr_qty"])
                            rem = c - p - cur
                            formatted_rem = f"{int(rem):,}" if rem.is_integer() else f"{rem:,.2f}"
                            self.contract_vars[k]["rem_qty"].set(formatted_rem)
                            
                            if rem < 0:
                                self.contract_vars[k]["lbl_rem"].config(foreground="red")
                            else:
                                self.contract_vars[k]["lbl_rem"].config(foreground="blue")
                        except:
                            pass
                            
                    c_qty.trace_add("write", update_rem_qty)
                    p_qty.trace_add("write", update_rem_qty)
                    curr_qty.trace_add("write", update_rem_qty)
                    
                    ttk.Entry(contract_inner_frame, textvariable=c_qty, width=8).grid(row=row_idx, column=4, padx=1, pady=1)
                    ttk.Entry(contract_inner_frame, textvariable=p_qty, width=8).grid(row=row_idx, column=5, padx=1, pady=1)
                    ttk.Label(contract_inner_frame, textvariable=curr_qty, width=6, anchor="e", foreground="green").grid(row=row_idx, column=6, padx=1, pady=1)
                    lbl_rem = ttk.Label(contract_inner_frame, textvariable=rem_qty, width=8, anchor="e", font=("Arial", 9, "bold"))
                    lbl_rem.grid(row=row_idx, column=7, padx=1, pady=1)
                    
                    ttk.Label(contract_inner_frame, textvariable=c_price_var, anchor="e", width=12).grid(row=row_idx, column=8, padx=2, pady=1)
                    ttk.Entry(contract_inner_frame, textvariable=c_var, width=12).grid(row=row_idx, column=9, padx=1, pady=1)
                    ttk.Entry(contract_inner_frame, textvariable=p_var, width=12).grid(row=row_idx, column=10, padx=1, pady=1)
                    
                    c_var.trace_add("write", lambda *a, v=c_var: format_currency(var=v))
                    p_var.trace_add("write", lambda *a, v=p_var: format_currency(var=v))
                    
                    self.contract_vars[full_key] = {
                        "c_qty": c_qty, "p_qty": p_qty, "curr_qty": curr_qty, "rem_qty": rem_qty, "lbl_rem": lbl_rem,
                        "c_price": unit_cost, "c_price_var": c_price_var, "contract": c_var, "prev": p_var
                    }
                    row_idx += 1
                    
        # Total Contract Amount Label
        self.total_contract_amt_var = tk.StringVar(value="총 계약금액: 0 원")
        ttk.Label(contract_frame, textvariable=self.total_contract_amt_var, font=("Arial", 11, "bold"), foreground="blue").pack(side=tk.BOTTOM, pady=10, anchor="e")
        
        def update_total_contract_amt(*args):
            total = 0
            for k, v in self.contract_vars.items():
                total += self.get_int(v["contract"])
                
            # 기타 경비 및 실비 정산 예산 합산
            if hasattr(self, 'exp_vars'):
                for k, v in self.exp_vars.items():
                    total += self.get_int(v["budget"])
            else:
                total += 86944000 # exp_vars 초기화 전 기본 실비 합계 (단수조정 제거 후 순수 합계)
                
            self.total_contract_amt_var.set(f"총 계약금액 (부가세 별도): {total:,} 원")
            
        for k, v in self.contract_vars.items():
            v["contract"].trace_add("write", update_total_contract_amt)
            
        def _on_contract_mousewheel(event):
            try:
                if str(event.widget).startswith(str(contract_frame)):
                    delta = event.delta
                    if event.num == 4: delta = 120
                    if event.num == 5: delta = -120
                    contract_canvas.yview_scroll(int(-1*(delta/120)), "units")
            except:
                pass
        contract_frame.bind_all("<MouseWheel>", _on_contract_mousewheel, add='+')
        contract_frame.bind_all("<Button-4>", _on_contract_mousewheel, add='+')
        contract_frame.bind_all("<Button-5>", _on_contract_mousewheel, add='+')

        f = ttk.Frame(contract_frame)
        f.pack(fill=tk.X, pady=2)
        ttk.Label(f, text="[프로젝트 총액]", width=12, font=("Arial", 9, "bold")).grid(row=0, column=0, rowspan=2, sticky=tk.W)
        ttk.Label(f, text="계약 총액:").grid(row=0, column=1, sticky=tk.W)
        self.total_contract_var = tk.StringVar(value="288,268,000")
        self.total_contract_var.trace_add("write", lambda *a, v=self.total_contract_var: format_currency(var=v))
        ttk.Entry(f, textvariable=self.total_contract_var, width=15).grid(row=0, column=2, padx=2)
        ttk.Label(f, text="원").grid(row=0, column=3)
        
        ttk.Label(f, text="전회 총액:").grid(row=1, column=1, sticky=tk.W, pady=2)
        self.total_prev_var = tk.StringVar(value="0")
        self.total_prev_var.trace_add("write", lambda *a, v=self.total_prev_var: format_currency(var=v))
        ttk.Entry(f, textvariable=self.total_prev_var, width=15).grid(row=1, column=2, padx=2)
        ttk.Label(f, text="원").grid(row=1, column=3)
        
        exp_frame = ttk.LabelFrame(content_frame, text="기타 경비 및 실비 정산 (월간)", padding=10)
        exp_frame.grid(row=0, column=1, sticky='nsew')
        
        hf = ttk.Frame(exp_frame)
        hf.pack(fill=tk.X, pady=2)
        ttk.Label(hf, text="항목", width=18, font=("Arial", 9, "bold")).grid(row=0, column=0, padx=2)
        ttk.Label(hf, text="책정예산", width=12, font=("Arial", 9, "bold")).grid(row=0, column=1, padx=2)
        ttk.Label(hf, text="전회청구액", width=12, font=("Arial", 9, "bold")).grid(row=0, column=2, padx=2)
        ttk.Label(hf, text="금회청구액", width=12, font=("Arial", 9, "bold")).grid(row=0, column=3, padx=2)
        ttk.Label(hf, text="잔여예산", width=12, font=("Arial", 9, "bold")).grid(row=0, column=4, padx=2)
        
        self.exp_vars = {}
        items = [
            ("equip", "원자력 안전부담금", 343000),
            ("safety", "안전관리비 (미사용)", 0),
            ("travel", "주재비 (미사용)", 0),
            ("print", "도서인쇄비 (미사용)", 0),
            ("liability", "손해배상공제 수수료", 1613000)
        ]
        
        for i, (k, name, budget) in enumerate(items, start=1):
            f = ttk.Frame(exp_frame)
            f.pack(fill=tk.X, pady=2)
            
            ttk.Label(f, text=name, width=18).grid(row=0, column=0, padx=2)
            
            b_var = tk.StringVar(value=f"{budget:,}")
            p_var = tk.StringVar(value="0")
            c_var = tk.IntVar(value=0)
            r_var = tk.StringVar(value=f"{budget:,}")
            
            def update_rem(*args, kv=k):
                try:
                    b = self.get_float(self.exp_vars[kv]["budget"])
                    p = self.get_float(self.exp_vars[kv]["prev"])
                    c = float(self.exp_vars[kv]["curr"].get())
                    rem = b - p - c
                    self.exp_vars[kv]["rem"].set(f"{int(rem):,}" if rem.is_integer() else f"{rem:,.2f}")
                    if rem < 0: self.exp_vars[kv]["lbl"].config(foreground="red")
                    else: self.exp_vars[kv]["lbl"].config(foreground="blue")
                except: pass
                
            b_var.trace_add("write", lambda *a, v=b_var: format_currency(var=v))
            p_var.trace_add("write", lambda *a, v=p_var: format_currency(var=v))
            
            b_var.trace_add("write", update_rem)
            p_var.trace_add("write", update_rem)
            c_var.trace_add("write", update_rem)
            
            ttk.Entry(f, textvariable=b_var, width=12).grid(row=0, column=1, padx=2)
            ttk.Entry(f, textvariable=p_var, width=12).grid(row=0, column=2, padx=2)
            ttk.Entry(f, textvariable=c_var, width=12).grid(row=0, column=3, padx=2)
            lbl = ttk.Label(f, textvariable=r_var, width=12, anchor="e", font=("Arial", 9, "bold"), foreground="blue")
            lbl.grid(row=0, column=4, padx=2)
            
            if k == "equip":
                btn = ttk.Button(f, text="계산기", width=6, command=lambda v=c_var: self.open_equip_calculator(v))
                btn.grid(row=0, column=5, padx=2)
            
            self.exp_vars[k] = {"budget": b_var, "prev": p_var, "curr": c_var, "rem": r_var, "lbl": lbl}
            
        self.equip_cost_var = self.exp_vars["equip"]["curr"]
        self.safety_cost_var = self.exp_vars["safety"]["curr"]
        self.travel_cost_var = self.exp_vars["travel"]["curr"]
        self.print_cost_var = self.exp_vars["print"]["curr"]
        self.liability_cost_var = self.exp_vars["liability"]["curr"]

        
        # --- RIGHT FRAME (누적 테이블, TAB 1에 배치) ---
        bottom_frame = ttk.Frame(self.work_pane)
        self.work_pane.add(bottom_frame, stretch="always")
        
        lbl_frame = ttk.Frame(bottom_frame)
        lbl_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(lbl_frame, text="[ 일일 작업 기록 목록 ]", font=("Arial", 11, "bold")).pack(side=tk.LEFT)
        ttk.Button(lbl_frame, text="기성청구", command=self.export_to_excel).pack(side=tk.RIGHT)
        ttk.Button(lbl_frame, text="기록 초기화", command=self.clear_records).pack(side=tk.RIGHT, padx=5)
        ttk.Button(lbl_frame, text="일일 장부에서 연동", command=self.import_from_daily_db).pack(side=tk.RIGHT, padx=5)
        ttk.Button(lbl_frame, text="선택 삭제", command=self.delete_selected_records).pack(side=tk.RIGHT)

        tree_container = ttk.Frame(bottom_frame)
        tree_container.pack(fill=tk.BOTH, expand=True)

        columns = ("date", "loc", "type", "time", "mat", "qty", "unit", "mat_cost", "lab_cost", "overhead", "tech", "total_amt")
        self.tree = ttk.Treeview(tree_container, columns=columns, show="headings", height=8)
        
        self.tree.heading("date", text="일자", anchor="center")
        self.tree.heading("loc", text="구간/위치", anchor="center")
        self.tree.heading("type", text="종류", anchor="center")
        self.tree.heading("time", text="형태", anchor="center")
        self.tree.heading("mat", text="자재", anchor="center")
        self.tree.heading("qty", text="실물량", anchor="center")
        self.tree.heading("unit", text="단위", anchor="center")
        self.tree.heading("mat_cost", text="재료비(원)", anchor="center")
        self.tree.heading("lab_cost", text="인건비(원)", anchor="center")
        self.tree.heading("overhead", text="제경비(원)", anchor="center")
        self.tree.heading("tech", text="기술료(원)", anchor="center")
        self.tree.heading("total_amt", text="공급가액(원)", anchor="center")
        
        default_widths = {
            "date": 80, "loc": 120, "type": 40, "time": 40, "mat": 90, 
            "qty": 40, "unit": 40,
            "mat_cost": 70, "lab_cost": 70, "overhead": 60, "tech": 60, "total_amt": 80
        }
        saved_widths = CONFIG.get("TREE_WIDTHS", {})
        
        self.tree.column("date", width=saved_widths.get("date", default_widths["date"]), anchor="center")
        self.tree.column("loc", width=saved_widths.get("loc", default_widths["loc"]), anchor="w")
        self.tree.column("type", width=saved_widths.get("type", default_widths["type"]), anchor="center")
        self.tree.column("time", width=saved_widths.get("time", default_widths["time"]), anchor="center")
        self.tree.column("mat", width=saved_widths.get("mat", default_widths["mat"]), anchor="center")
        self.tree.column("qty", width=saved_widths.get("qty", default_widths["qty"]), anchor="center")
        self.tree.column("unit", width=saved_widths.get("unit", default_widths["unit"]), anchor="center")
        self.tree.column("mat_cost", width=saved_widths.get("mat_cost", default_widths["mat_cost"]), anchor="center")
        self.tree.column("lab_cost", width=saved_widths.get("lab_cost", default_widths["lab_cost"]), anchor="center")
        self.tree.column("overhead", width=saved_widths.get("overhead", default_widths["overhead"]), anchor="center")
        self.tree.column("tech", width=saved_widths.get("tech", default_widths["tech"]), anchor="center")
        self.tree.column("total_amt", width=saved_widths.get("total_amt", default_widths["total_amt"]), anchor="center")
        
        tree_scroll = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
        tree_hscroll = ttk.Scrollbar(tree_container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=tree_scroll.set, xscrollcommand=tree_hscroll.set)
        
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        tree_hscroll.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)
        self.tree.bind('<Delete>', self.delete_selected_records)
        self.tree.bind('<BackSpace>', self.delete_selected_records)
        
        # 저장된 탭 영역(Sash) 너비 복원
        # 저장된 탭 영역(Sash) 너비 복원 (안정화 후 적용을 위한 타이머 방식)
        self._sash_timer = None
        def on_configure(event):
            if getattr(self, "_sash_restored", False):
                return
                
            if self._sash_timer:
                self.after_cancel(self._sash_timer)
                
            def do_restore():
                try:
                    sash_pos = int(CONFIG.get("SASH_POS", 450))
                    # 창이 완전히 렌더링 된 이후에 복원
                    self.work_pane.sash_place(0, sash_pos, 0)
                    self._sash_restored = True
                except:
                    pass
                    
            # 화면 크기 변경(Configure) 이벤트가 멈추고 200ms 뒤에 한 번만 실행
            self._sash_timer = self.after(200, do_restore)
                
        self.work_pane.bind("<Configure>", on_configure)
        
        # 마우스로 드래그해서 놓을 때 즉시 저장
        def save_sash_on_release(event):
            try:
                sash_x, _ = self.work_pane.sash_coord(0)
                CONFIG["SASH_POS"] = int(sash_x)
                save_config(CONFIG)
                print(f"[DEBUG] save_sash_on_release SUCCESS! Saved at {sash_x}")
            except Exception as e:
                print(f"[DEBUG] save_sash_on_release ERROR: {e}")
        self.work_pane.bind("<ButtonRelease-1>", save_sash_on_release)
                
        self.update_dynamic_ui()

    def open_equip_calculator(self, target_var):
        top = tk.Toplevel(self)
        top.title("장비손료 실비 정산 계산기")
        top.geometry("380x250")
        top.transient(self)
        top.grab_set()

        ttk.Label(top, text="[ 장비손료 산출식 : 계 * 장비투입일수 / 20 ]", font=("Arial", 10, "bold")).pack(pady=10)

        f = ttk.Frame(top)
        f.pack(fill=tk.BOTH, expand=True, padx=20)

        rt_rate = 396378
        ut_rate = 219790
        cr_rate = 1542240

        rt_days = tk.StringVar(value="0")
        ut_days = tk.StringVar(value="0")
        cr_days = tk.StringVar(value="0")
        
        total_var = tk.StringVar(value="0")

        def calc_total(*args):
            try:
                rt = float(rt_days.get() or 0)
                ut = float(ut_days.get() or 0)
                cr = float(cr_days.get() or 0)
                amt = int(rt_rate * rt / 20.0) + int(ut_rate * ut / 20.0) + int(cr_rate * cr / 20.0)
                total_var.set(f"{amt:,}")
            except:
                total_var.set("0")

        rt_days.trace_add("write", calc_total)
        ut_days.trace_add("write", calc_total)
        cr_days.trace_add("write", calc_total)

        ttk.Label(f, text="RT 투입일수:").grid(row=0, column=0, pady=5, sticky="w")
        ttk.Entry(f, textvariable=rt_days, width=10).grid(row=0, column=1, pady=5, padx=5)
        ttk.Label(f, text="일").grid(row=0, column=2, sticky="w")

        ttk.Label(f, text="UT 투입일수:").grid(row=1, column=0, pady=5, sticky="w")
        ttk.Entry(f, textvariable=ut_days, width=10).grid(row=1, column=1, pady=5, padx=5)
        ttk.Label(f, text="일").grid(row=1, column=2, sticky="w")

        ttk.Label(f, text="크롤러 투입일수:").grid(row=2, column=0, pady=5, sticky="w")
        ttk.Entry(f, textvariable=cr_days, width=10).grid(row=2, column=1, pady=5, padx=5)
        ttk.Label(f, text="일").grid(row=2, column=2, sticky="w")

        ttk.Label(f, text="합계 금액:").grid(row=3, column=0, pady=15, sticky="w")
        ttk.Label(f, textvariable=total_var, font=("Arial", 11, "bold"), foreground="blue").grid(row=3, column=1, columnspan=2, pady=15, sticky="w")

        def apply():
            try:
                val = int(total_var.get().replace(",", ""))
                target_var.set(val)
                top.destroy()
            except:
                pass

        btn_f = ttk.Frame(top)
        btn_f.pack(pady=10)
        ttk.Button(btn_f, text="적용", command=apply).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_f, text="취소", command=top.destroy).pack(side=tk.LEFT)

    def update_dynamic_ui(self, *args):
        ndt_type = self.ndt_type_var.get()
        if ndt_type == "PAUT":
            self.material_combo.config(values=["300A 이상", "250A", "200A", "150A-125A", "100A 이하"], state="readonly")
            if self.material_var.get() not in ["300A 이상", "250A", "200A", "150A-125A", "100A 이하"]:
                self.material_var.set("300A 이상")
        elif ndt_type == "RT":
            self.material_combo.config(values=['3 1/3 x 12"', '3 1/3 x 6"'], state="readonly")
            if self.material_var.get() not in ['3 1/3 x 12"', '3 1/3 x 6"']:
                self.material_var.set('3 1/3 x 12"')
        elif ndt_type == "MT":
            self.material_combo.config(values=["MT"], state="disabled")
            self.material_var.set("MT")
        elif ndt_type == "PT":
            self.material_combo.config(values=["PT"], state="disabled")
            self.material_var.set("PT")
            
        self.source_frame.pack_forget()
        self.pipe_frame.pack_forget()
        self.thickness_frame.pack_forget()

    def get_correction_factor(self):
        # 2026 단가계약은 고정 단가를 주로 사용하므로, 
        # 보정계수(source, pipe, thickness)가 별도로 지정되지 않으면 기본 1.0 적용.
        return 1.0

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
        
        key = f"{ndt_type}_{material_type}"
        loc_type = getattr(self, "loc_type_var", None)
        loc_type_val = loc_type.get() if loc_type else "열배관"

        costs = calculate_billing(
            quantity=qty,
            adjusted_quantity=adjusted_qty,
            material_key=key,
            ndt_type=ndt_type,
            location=loc_type_val,
            work_time=work_time,
            material_costs=MATERIAL_COST,
            labor_costs=LABOR_COST,
            overhead_rate=overhead_rate,
            technical_fee_rate=tech_fee_rate,
        )
        total_mat_cost = costs["mat_cost"]
        total_lab_cost = costs["lab_cost"]
        overhead_cost = costs["overhead"]
        tech_fee = costs["tech"]
        subtotal = costs["subtotal"]
        vat = costs["vat"]
        total_amount = costs["total_amount"]
        
        display_loc = f"[{loc_type_val}] {loc_str}".strip() if loc_str else f"[{loc_type_val}]"
        
        return {
            "date": date_str,
            "loc": display_loc,
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
            
            key = f"{res['ndt_type']}_{res['material_type']}"
            mat_unit = MATERIAL_COST.get(key, MATERIAL_COST.get(res['material_type'], 0))
            loc_type = getattr(self, "loc_type_var", None)
            loc_type_val = loc_type.get() if loc_type else "열배관"
            
            if loc_type_val in LABOR_COST:
                lab_unit = LABOR_COST[loc_type_val][res['work_time']].get(key, LABOR_COST[loc_type_val][res['work_time']].get(res['ndt_type'], 0))
            else:
                lab_unit = LABOR_COST.get(res['work_time'], {}).get(key, LABOR_COST.get(res['work_time'], {}).get(res['ndt_type'], 0))
            
            txt = (f"▶ [현재 입력값] 일자: {res['date']} | 구간: {res['loc']}\n"
                   f"▶ [적용 기준] 재료비 단가: {mat_unit:,}원 | 인건비 단가: {lab_unit:,}원\n"
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
        key = f"{res['ndt_type']}_{res['material_type']}"
        mat_unit = MATERIAL_COST.get(key, MATERIAL_COST.get(res['material_type'], 0))
        
        if "[플랜트(관리소)]" in res['loc']:
            loc_type_val = "플랜트(관리소)"
        else:
            loc_type_val = "열배관"
            
        if loc_type_val in LABOR_COST:
            lab_unit = LABOR_COST[loc_type_val].get(res['work_time'], {}).get(key, LABOR_COST[loc_type_val].get(res['work_time'], {}).get(res['ndt_type'], 0))
        else:
            lab_unit = LABOR_COST.get(res['work_time'], {}).get(key, LABOR_COST.get(res['work_time'], {}).get(res['ndt_type'], 0))
        
        txt = (f"▶ [선택된 기록] 일자: {res['date']} | 구간: {res['loc']}\n"
               f"▶ [적용 기준] 재료비 단가: {mat_unit:,}원 | 인건비 단가: {lab_unit:,}원\n"
               f"▶ [공급 가액] {res['subtotal']:,} 원 (재료비 {res['mat_cost']:,} + 인건비 {res['lab_cost']:,} + 제경비 {res['overhead']:,} + 기술료 {res['tech']:,})\n"
               f"▶ [최종 금액] 총 청구액 {res['total_amount']:,} 원 (부가세 {res['vat']:,}원 포함)\n")
        
        self.result_text.config(state=tk.NORMAL)
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, txt)
        self.result_text.config(state=tk.DISABLED)

    def import_from_daily_db(self):
        try:
            import pandas as pd
            import os
            db_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'data', 'Material_Inventory.xlsx')
            if not os.path.exists(db_path):
                messagebox.showerror("오류", "DB 파일을 찾을 수 없습니다.")
                return
            df = pd.read_excel(db_path, sheet_name='DailyUsage')
            
            if '검사방법' not in df.columns:
                messagebox.showinfo("안내", "데이터베이스 형식이 맞지 않습니다.")
                return
                
            # [NEW] 현장 탭의 조회 기간 및 현장 필터를 그대로 적용
            if hasattr(self, 'main_app') and self.main_app:
                try:
                    start_str = self.main_app.ent_daily_start_date.get().strip()
                    end_str = self.main_app.ent_daily_end_date.get().strip()
                    site_filter = self.main_app.cb_daily_filter_site.get().strip()
                    
                    if start_str or end_str:
                        df['Date'] = pd.to_datetime(df['Date'])
                        if start_str:
                            df = df[df['Date'] >= pd.to_datetime(start_str)]
                        if end_str:
                            df = df[df['Date'] <= (pd.to_datetime(end_str) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1))]
                            
                    if site_filter and site_filter != "전체":
                        df = df[df['Site'] == site_filter]
                except Exception as e:
                    print(f"Date filter applying error: {e}")
                
            ndt_df = df[df['검사방법'].isin(['RT', 'UT', 'PT', 'PAUT'])]
            if ndt_df.empty:
                messagebox.showinfo("안내", "연동할 NDT 작업 기록이 없습니다.")
                return
                
            if not messagebox.askyesno("데이터 연동", f"일일 작업보에서 {len(ndt_df)}건의 NDT 기록을 가져오시겠습니까?\n(기존 기록은 지워집니다)"): return
            
            self.clear_records()
            count = 0
            for _, row in ndt_df.iterrows():
                date_str = str(row.get('Date', ''))[:10]
                loc_str = str(row.get('Site', ''))
                ndt_type = str(row.get('검사방법', 'RT'))
                work_time = str(row.get('작업형태', '일반'))
                if work_time not in ['일반', '야간', '휴일']: work_time = '일반'
                material_type = str(row.get('Material', ''))
                qty = float(row.get('검사량', 0.0) if not pd.isna(row.get('검사량')) else 0.0)
                if qty == 0.0: qty = float(row.get('Usage', 0.0) if not pd.isna(row.get('Usage')) else 0.0)
                
                self.date_var.set(date_str)
                if '관리소' in loc_str or '플랜트' in loc_str:
                    self.loc_type_var.set('플랜트(관리소)')
                else:
                    self.loc_type_var.set('열배관')
                self.loc_var.set(loc_str)
                self.ndt_type_var.set(ndt_type)
                self.work_time_var.set(work_time)
                
                # trigger update_dynamic_ui to populate material values correctly
                self.update_dynamic_ui()
                
                if material_type and material_type != 'nan': 
                    self.material_var.set(material_type)
                self.quantity_var.set(qty)
                
                cond1 = str(row.get('조건1', ''))
                if cond1 and cond1 != 'nan':
                    if ndt_type == 'RT': self.source_var.set(cond1)
                    else: self.pipe_var.set(cond1)
                
                cond2 = str(row.get('조건2', ''))
                if cond2 and cond2 != 'nan': self.thickness_var.set(cond2)
                
                self.add_to_record(auto_save=False)
                count += 1
                
            self.save_billing_records()
            messagebox.showinfo("연동 완료", f"성공적으로 {count}건을 연동했습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"데이터 연동 중 오류가 발생했습니다: {e}")

    def update_qty_summary(self):
        for k in self.contract_vars:
            self.contract_vars[k]["curr_qty"].set("0")
            
        for rec in self.records:
            loc = "플랜트(관리소)" if "관리소" in rec["loc"] or "플랜트" in rec.get("loc_type", rec["loc"]) else "열배관"
            t_time = rec.get("work_time", "일반")
            ndt_type = rec["ndt_type"]
            mat = ""
            mat = f"{rec['ndt_type']}_{rec['material_type']}"
                
            if mat:
                key = f"{loc}_{t_time}_{mat}"
                if key in self.contract_vars:
                    cur_val = self.get_float(self.contract_vars[key]["curr_qty"])
                    new_val = cur_val + rec["qty"]
                    self.contract_vars[key]["curr_qty"].set(f"{int(new_val):,}" if float(new_val).is_integer() else f"{new_val:,.2f}")
                    
        self.export_billing_data()

    def export_billing_data(self):
        try:
            import json
            import os
            data_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'data')
            os.makedirs(data_dir, exist_ok=True)
            export_path = os.path.join(data_dir, 'billing_export.json')
            
            export_data = {}
            for key, var_dict in self.contract_vars.items():
                c_qty = self.get_float(var_dict["c_qty"])
                p_qty = self.get_float(var_dict["p_qty"])
                cur_qty = self.get_float(var_dict["curr_qty"])
                export_data[key] = {
                    "contract_qty": c_qty,
                    "prev_qty": p_qty,
                    "current_qty": cur_qty
                }
            with open(export_path, 'w', encoding='utf-8') as f:
                json.dump(export_data, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"Failed to export billing data: {e}")

    def add_to_record(self, auto_save=True):
        res = self.calculate()
        if res:
            self.records.append(res)
            self.tree.insert("", tk.END, values=(
                res["date"], res["loc"], res["ndt_type"], res["work_time"], 
                res["material_type"], f"{res['qty']:.1f}", res["unit"],
                f"{res.get('mat_cost', 0):,}", f"{res.get('lab_cost', 0):,}",
                f"{res['overhead']:,}", f"{res['tech']:,}", f"{res['subtotal']:,}"
            ))
            self.update_qty_summary()
            if auto_save:
                self.save_billing_records()
            
    def clear_records(self):
        self.records = []
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.update_qty_summary()
        self.save_billing_records()
        
    def delete_selected_records(self, event=None):
        selected_items = self.tree.selection()
        if not selected_items:
            if event is None: # 버튼 클릭으로 호출된 경우에만 경고창
                messagebox.showwarning("선택 오류", "삭제할 항목을 먼저 선택해주세요.")
            return
            
        if not messagebox.askyesno("선택 삭제", f"선택한 {len(selected_items)}개의 기록을 영구히 삭제하시겠습니까?"):
            return
            
        indices = sorted([self.tree.index(item) for item in selected_items], reverse=True)
        for i in indices:
            self.records.pop(i)
            
        for item in selected_items:
            self.tree.delete(item)
            
        self.update_qty_summary()
        self.save_billing_records()
            
    def carry_over_round(self):
        selected_items = self.tree.selection()
        if selected_items:
            msg = f"선택한 {len(selected_items)}개의 작업 기록만 '전회'로 누적하고 지우시겠습니까?\n(선택되지 않은 기록은 남습니다.)\n\n※ 데이터 안전을 위해 이월 전 현재 상태를 파일로 먼저 저장해야 합니다."
            is_partial = True
        else:
            msg = "선택된 항목이 없습니다.\n전체 기록(금회 물량/금액 전체)을 '전회'로 누적하고 초기화하시겠습니까?\n\n※ 데이터 안전을 위해 이월 전 현재 상태를 파일로 먼저 저장해야 합니다."
            is_partial = False
            
        if not messagebox.askyesno("다음 회차로 이월", msg):
            return
            
        # 강제 백업 로직 추가
        current_round = self.round_var.get()
        default_backup_name = f"제{current_round}회_마감기록_{datetime.now().strftime('%Y%m%d_%H%M')}.ndt"
        filepath = filedialog.asksaveasfilename(defaultextension=".ndt", initialfile=default_backup_name, filetypes=[("NDT Project", "*.ndt")], title="[안전장치] 이월 전 현재 상태 백업 저장")
        
        if not filepath:
            messagebox.showwarning("이월 취소", "저장이 취소되어 이월 작업을 중단합니다.")
            return
            
        try:
            data = {
                "round": self.round_var.get(),
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
                    "equip": self.exp_vars["equip"]["curr"].get(),
                    "equip_budget": self.get_int(self.exp_vars["equip"]["budget"]),
                    "equip_prev": self.get_int(self.exp_vars["equip"]["prev"]),
                    "safety": self.exp_vars["safety"]["curr"].get(),
                    "safety_budget": self.get_int(self.exp_vars["safety"]["budget"]),
                    "safety_prev": self.get_int(self.exp_vars["safety"]["prev"]),
                    "travel": self.exp_vars["travel"]["curr"].get(),
                    "travel_budget": self.get_int(self.exp_vars["travel"]["budget"]),
                    "travel_prev": self.get_int(self.exp_vars["travel"]["prev"]),
                    "print": self.exp_vars["print"]["curr"].get(),
                    "print_budget": self.get_int(self.exp_vars["print"]["budget"]),
                    "print_prev": self.get_int(self.exp_vars["print"]["prev"]),
                    "liability": self.exp_vars["liability"]["curr"].get(),
                    "liability_budget": self.get_int(self.exp_vars["liability"]["budget"]),
                    "liability_prev": self.get_int(self.exp_vars["liability"]["prev"])
                }
            }
            data["total_amt"] = {"contract": self.get_int(self.total_contract_var), "prev": self.get_int(self.total_prev_var)}
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
        except Exception as e:
            messagebox.showerror("저장 오류", f"저장 중 오류가 발생하여 이월을 중단합니다: {e}")
            return
            
        if is_partial:
            indices = [self.tree.index(item) for item in selected_items]
            target_records = [self.records[i] for i in indices]
        else:
            target_records = self.records
            
        for cat, v in self.contract_vars.items():
            p_qty = self.get_float(v["p_qty"])
            p_amt = self.get_int(v["prev"])
            
            c_qty = 0.0
            cur_amt = 0
            
            for r in target_records:
                loc = "플랜트(관리소)" if "관리소" in r["loc"] or "플랜트" in r.get("loc_type", r["loc"]) else "열배관"
                t_time = r.get("work_time", "일반")
                mat = ""
                mat = f"{r['ndt_type']}_{r['material_type']}"
                    
                key = f"{loc}_{t_time}_{mat}"
                if key == cat:
                    c_qty += r["qty"]
                    cur_amt += r["subtotal"]
                
            new_p_qty = p_qty + c_qty
            formatted_qty = f"{int(new_p_qty):,}" if float(new_p_qty).is_integer() else f"{new_p_qty:,.2f}"
            v["p_qty"].set(formatted_qty)
            v["prev"].set(f"{p_amt + cur_amt:,}")
            
        exp_curr = sum([self.equip_cost_var.get(), self.safety_cost_var.get(), self.travel_cost_var.get(), self.print_cost_var.get(), self.liability_cost_var.get()])
        
        for k in ["equip", "safety", "travel", "print", "liability"]:
            p = self.get_int(self.exp_vars[k]["prev"])
            c = self.exp_vars[k]["curr"].get()
            self.exp_vars[k]["prev"].set(f"{p + c:,}")
            self.exp_vars[k]["curr"].set(0)
        
        total_sub_cur = sum(r["subtotal"] for r in target_records) + exp_curr
        prev_total = self.get_int(self.total_prev_var)
        self.total_prev_var.set(f"{prev_total + total_sub_cur:,}")

        if is_partial:
            for item in reversed(selected_items):
                idx = self.tree.index(item)
                del self.records[idx]
                self.tree.delete(item)
            self.update_qty_summary()
        else:
            self.clear_records()
            
        self.round_var.set(self.round_var.get() + 1)
        messagebox.showinfo("이월 완료", f"제 {self.round_var.get()} 회차 기성으로 이월되었습니다.")
            
    def save_project(self):
        try:
            filepath = filedialog.asksaveasfilename(defaultextension=".ndt", filetypes=[("NDT Project", "*.ndt")], title="작업 저장하기")
            if not filepath: return
            
            data = {
                "round": self.round_var.get(),
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
                    "equip": self.exp_vars["equip"]["curr"].get(),
                    "equip_budget": self.get_int(self.exp_vars["equip"]["budget"]),
                    "equip_prev": self.get_int(self.exp_vars["equip"]["prev"]),
                    "safety": self.exp_vars["safety"]["curr"].get(),
                    "safety_budget": self.get_int(self.exp_vars["safety"]["budget"]),
                    "safety_prev": self.get_int(self.exp_vars["safety"]["prev"]),
                    "travel": self.exp_vars["travel"]["curr"].get(),
                    "travel_budget": self.get_int(self.exp_vars["travel"]["budget"]),
                    "travel_prev": self.get_int(self.exp_vars["travel"]["prev"]),
                    "print": self.exp_vars["print"]["curr"].get(),
                    "print_budget": self.get_int(self.exp_vars["print"]["budget"]),
                    "print_prev": self.get_int(self.exp_vars["print"]["prev"]),
                    "liability": self.exp_vars["liability"]["curr"].get(),
                    "liability_budget": self.get_int(self.exp_vars["liability"]["budget"]),
                    "liability_prev": self.get_int(self.exp_vars["liability"]["prev"])
                }
            }
            data["total_amt"] = {"contract": self.get_int(self.total_contract_var), "prev": self.get_int(self.total_prev_var)}
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
            messagebox.showinfo("저장 완료", "작업이 성공적으로 저장되었습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"저장 중 오류가 발생했습니다: {e}")

    def auto_load_contract_qty(self):
        global CONFIG, MATERIAL_COST, LABOR_COST
        CONFIG = load_config()
        MATERIAL_COST = CONFIG["MATERIAL_COST"]
        LABOR_COST = CONFIG["LABOR_COST"]
        # 강제로 소스코드의 완벽한 마스터 물량(DEFAULT_CONFIG)을 불러오도록 수정
        contract_qtys = DEFAULT_CONFIG.get("CONTRACT_QTY", {})
        
        debug_msg = []
        updated_count = 0
        
        for loc in contract_qtys:
            for t_time in contract_qtys[loc]:
                if isinstance(contract_qtys[loc][t_time], dict):
                    for mat, val in contract_qtys[loc][t_time].items():
                        key = f"{loc}_{t_time}_{mat}"
                        if key in self.contract_vars:
                            formatted = f"{int(val):,}" if float(val).is_integer() else f"{float(val):,.2f}"
                            self.contract_vars[key]["c_qty"].set(formatted)
                            updated_count += 1
                            if updated_count <= 3:
                                debug_msg.append(f"{key}: {formatted}")
                            
                            # Recalculate unit cost dynamically based on latest config and UI rates
                            try:
                                lab_unit = LABOR_COST[loc][t_time].get(mat, 0)
                                mat_unit = MATERIAL_COST.get(mat, 0)
                                oh = int(lab_unit * float(self.overhead_rate_var.get()) / 100.0)
                                tech = int((lab_unit + oh) * float(self.tech_fee_rate_var.get()) / 100.0)
                                unit_cost = mat_unit + lab_unit + oh + tech
                                
                                self.contract_vars[key]["c_price"] = unit_cost
                                self.contract_vars[key]["c_price_var"].set(f"{int(unit_cost):,}")
                            except Exception as e:
                                unit_cost = self.contract_vars[key].get("c_price", 0)
                                self.contract_vars[key].get("c_price_var", tk.StringVar()).set(f"{int(unit_cost):,}")
                                
                            amt = float(val) * unit_cost
                            self.contract_vars[key]["contract"].set(f"{int(amt):,}")
                else:
                    # Backward compatibility if config.json hasn't been updated or has old format
                    mat = t_time
                    val = contract_qtys[loc][t_time]
                    key = f"{loc}_일반_{mat}"
                    if key in self.contract_vars:
                        formatted = f"{int(val):,}" if float(val).is_integer() else f"{float(val):,.2f}"
                        self.contract_vars[key]["c_qty"].set(formatted)
                        try:
                            lab_unit = LABOR_COST[loc]["일반"].get(mat, 0)
                            mat_unit = MATERIAL_COST.get(mat, 0)
                            oh = int(lab_unit * float(self.overhead_rate_var.get()) / 100.0)
                            tech = int((lab_unit + oh) * float(self.tech_fee_rate_var.get()) / 100.0)
                            unit_cost = mat_unit + lab_unit + oh + tech
                            
                            self.contract_vars[key]["c_price"] = unit_cost
                        except Exception:
                            unit_cost = self.contract_vars[key].get("c_price", 0)
                            
                        amt = float(val) * unit_cost
                        self.contract_vars[key]["contract"].set(f"{int(amt):,}")
                    
        msg = f"총 {updated_count}개의 항목이 업데이트 되었습니다.\n\n[업데이트 샘플]\n" + "\n".join(debug_msg)
        self.export_billing_data()
        messagebox.showinfo("불러오기 완료", msg)

    def load_project(self):
        try:
            filepaths = filedialog.askopenfilenames(filetypes=[("NDT Project", "*.ndt")], title="작업 불러오기 (여러 파일 선택 시 병합됨)")
            if not filepaths: return
            
            self.clear_records()
            self.records = []
            
            # 다중 파일 선택 시 순서 보장을 위해 정렬
            filepaths = sorted(filepaths)
            
            latest_data = None
            max_round = -1
            
            for filepath in filepaths:
                with open(filepath, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    
                current_records = data.get("records", [])
                self.records.extend(current_records)
                
                for res in current_records:
                    self.tree.insert("", tk.END, values=(
                        res["date"], res["loc"], res["ndt_type"], res["work_time"], 
                        res["material_type"], f"{res['qty']:.1f}", res["unit"],
                        f"{res.get('mat_cost', 0):,}", f"{res.get('lab_cost', 0):,}",
                        f"{res['overhead']:,}", f"{res['tech']:,}", f"{res['subtotal']:,}"
                    ))
                    
                curr_round = data.get("round", 1)
                if curr_round > max_round:
                    max_round = curr_round
                    latest_data = data
            
            if not latest_data: return
            
            self.round_var.set(latest_data.get("round", max_round))
            cont = latest_data.get("contract", {})
            
            # 하위 호환성 (과거 저장 파일에 'RT' 하나만 있는 경우 RT_B에 모두 몰아넣음)
            if "RT" in cont and "RT_B" not in cont:
                cont["RT_B"] = cont.pop("RT")
                
            for t, v in self.contract_vars.items():
                if t in cont:
                    cq = cont[t].get("c_qty", 0.0)
                    pq = cont[t].get("p_qty", 0.0)
                    v["c_qty"].set(f"{int(cq):,}" if cq.is_integer() else f"{cq:,.2f}")
                    v["contract"].set(f"{cont[t].get('contract', 0):,}")
                    v["p_qty"].set(f"{int(pq):,}" if pq.is_integer() else f"{pq:,.2f}")
                    v["prev"].set(f"{cont[t].get('prev', 0):,}")
            
            self.update_qty_summary()
            
            ex = latest_data.get("expenses", {})
            self.exp_vars["equip"]["budget"].set(f"{ex.get('equip_budget', 41000000):,}")
            self.exp_vars["equip"]["prev"].set(f"{ex.get('equip_prev', 0):,}")
            self.exp_vars["equip"]["curr"].set(ex.get("equip", 0))
            
            self.exp_vars["safety"]["budget"].set(f"{ex.get('safety_budget', 28507303):,}")
            self.exp_vars["safety"]["prev"].set(f"{ex.get('safety_prev', 0):,}")
            self.exp_vars["safety"]["curr"].set(ex.get("safety", 0))
            
            self.exp_vars["travel"]["budget"].set(f"{ex.get('travel_budget', 2226096):,}")
            self.exp_vars["travel"]["prev"].set(f"{ex.get('travel_prev', 0):,}")
            self.exp_vars["travel"]["curr"].set(ex.get("travel", 0))
            
            self.exp_vars["print"]["budget"].set(f"{ex.get('print_budget', 481600):,}")
            self.exp_vars["print"]["prev"].set(f"{ex.get('print_prev', 0):,}")
            self.exp_vars["print"]["curr"].set(ex.get("print", 0))
            
            self.exp_vars["liability"]["budget"].set(f"{ex.get('liability_budget', 14729000):,}")
            self.exp_vars["liability"]["prev"].set(f"{ex.get('liability_prev', 0):,}")
            self.exp_vars["liability"]["curr"].set(ex.get("liability", 0))
            
            tot = latest_data.get("total_amt", {})
            self.total_contract_var.set(f"{tot.get('contract', 288268000):,}")
            self.total_prev_var.set(f"{tot.get('prev', 0):,}")
            
            msg = f"총 {len(filepaths)}개의 작업 파일에서 {len(self.records)}개의 기록을 성공적으로 병합하여 불러왔습니다."
            messagebox.showinfo("불러오기 완료", msg)
        except Exception as e:
            messagebox.showerror("오류", f"불러오기 중 오류가 발생했습니다: {e}")

    def open_report_hub(self):
        import subprocess
        import sys
        import os
        hub_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "문서_통합_관리_허브.py")
        if os.path.exists(hub_path):
            subprocess.Popen([sys.executable, hub_path])
        else:
            messagebox.showerror("오류", "문서_통합_관리_허브.py 파일을 찾을 수 없습니다.")

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
        target_records = []
        if not self.records:
            if not messagebox.askyesno("기록 없음", "출력할 작업 기록(금회 기성)이 없습니다. 계약 내역만 출력하시겠습니까?"):
                return
        else:
            selected_items = self.tree.selection()
            if selected_items:
                if not messagebox.askyesno("부분 출력", f"선택된 {len(selected_items)}개의 기록만 기성 청구 내역서로 출력하시겠습니까?"):
                    return
                indices = [self.tree.index(item) for item in selected_items]
                target_records = [self.records[i] for i in indices]
            else:
                if not messagebox.askyesno("전체 출력", "선택된 항목이 없습니다. 전체 기록을 기성 청구 내역서로 출력하시겠습니까?"):
                    return
                target_records = self.records
                
        dates_all = sorted([r["date"] for r in target_records])
        if dates_all:
            start_date = dates_all[0].replace("-", ".")
            end_date = dates_all[-1].replace("-", ".")
            global_period = f"{start_date} ~ {end_date}" if start_date != end_date else start_date
        else:
            global_period = "기간 없음"
            
        round_val = self.round_var.get()
        default_name = f"제{round_val}회_기성청구내역서_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
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
            ws.Range("A1").Value = f"제 {round_val} 회 비파괴검사기술용역 기성청구 내역서"
            ws.Range("A1").Font.Size = 20
            ws.Range("A1").Font.Bold = True
            ws.Range("A1").HorizontalAlignment = -4108
            ws.Range("A1").VerticalAlignment = -4108
            
            ws.Range("A4:B4").Merge()
            ws.Range("A4").Value = "공 사 명 :"
            ws.Range("A4").Font.Bold = True
            ws.Range("A4").Font.Size = 12
            ws.Range("C4:J4").Merge()
            ws.Range("C4").Value = "2026년 중앙지사 열수송관 비파괴검사용역 단가계약"
            ws.Range("C4").Font.Bold = True
            ws.Range("C4").Font.Size = 14
            
            ws.Range("K4:L4").Merge()
            ws.Range("K4").Value = "청구 기간 :"
            ws.Range("K4").Font.Bold = True
            ws.Range("K4").Font.Size = 11
            ws.Range("K4").HorizontalAlignment = -4152
            ws.Range("M4:O4").Merge()
            ws.Range("M4").Value = global_period
            ws.Range("M4").Font.Size = 11
            ws.Range("M4").HorizontalAlignment = -4108
            
            # --- 기성 요약 테이블 ---
            ws.Cells(6, 1).Value = "구간"
            ws.Range(ws.Cells(6, 1), ws.Cells(7, 1)).Merge()
            ws.Cells(6, 2).Value = "시간"
            ws.Range(ws.Cells(6, 2), ws.Cells(7, 2)).Merge()
            ws.Cells(6, 3).Value = "항목"
            ws.Range(ws.Cells(6, 3), ws.Cells(7, 3)).Merge()
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
            
            categories = list(self.contract_vars.keys()) + ["장비손료", "안전관리비", "주재비 및 출장여비", "도서인쇄비", "기타실비 소계", "엔지니어링 손해배상공제료", "총 계"]
            
            ws.Range(ws.Cells(6, 1), ws.Cells(6 + len(categories) + 1, 15)).Borders.LineStyle = 1
            
            row = 8
            data_rows = []
            extra_rows = []
            subtotal_row = 0
            liability_row = 0
            total_row = 0
            
            for cat in categories:
                c_qty, p_qty, cur_qty, tot_qty, rem_qty = "", "", "", "", ""
                c_amt, p_amt, cur_amt, tot_amt, rem_amt = 0, 0, 0, 0, 0
                
                if cat == "총 계":
                    c_amt = self.get_int(self.total_contract_var)
                    p_amt = self.get_int(self.total_prev_var)
                    cur_amt = sum(r["subtotal"] for r in target_records) + extra_items_total
                    total_row = row
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
                
                if cat in ["총 계", "기타실비 소계", "장비손료", "안전관리비", "주재비 및 출장여비", "도서인쇄비", "엔지니어링 손해배상공제료"]:
                    ws.Cells(row, 1).Value = cat
                    ws.Range(ws.Cells(row, 1), ws.Cells(row, 5)).Merge()
                    ws.Cells(row, 1).HorizontalAlignment = -4108
                    if cat in ["총 계", "기타실비 소계"]:
                        ws.Range(ws.Cells(row, 1), ws.Cells(row, 15)).Interior.Color = 15987699
                        ws.Cells(row, 1).Font.Bold = True
                else:
                    parts = cat.split('_')
                    loc = parts[0]
                    t_time = parts[1]
                    m_key = '_'.join(parts[2:])
                    unit = "매" if m_key.startswith("RT") else "M"
                    ws.Cells(row, 1).Value = loc
                    ws.Cells(row, 2).Value = t_time
                    ws.Cells(row, 3).Value = m_key
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
                float_fmt = '#,##0.00;-#,##0.00;"-"'
                
                if cat == "총 계":
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
                    
                    ws.Cells(row, 6).Value = round(float(c_qty), 2) if c_qty else 0
                    ws.Cells(row, 7).Formula = f"=TRUNC(F{row}*E{row})"
                    
                    ws.Cells(row, 8).Value = round(float(p_qty), 2) if p_qty else 0
                    ws.Cells(row, 9).Formula = f"=TRUNC(H{row}*E{row})"
                    
                    ws.Cells(row, 10).Value = round(float(cur_qty), 2) if cur_qty else 0
                    ws.Cells(row, 11).Formula = f"=TRUNC(J{row}*E{row})"
                    
                    ws.Cells(row, 12).Formula = f"=H{row}+J{row}"
                    ws.Cells(row, 13).Formula = f"=I{row}+K{row}"
                    
                    ws.Cells(row, 14).Formula = f"=F{row}-L{row}"
                    ws.Cells(row, 15).Formula = f"=G{row}-M{row}"
                    
                    for col in [6, 8, 10, 12, 14]: ws.Cells(row, col).NumberFormat = fmt_qty
                    for col in [7, 9, 11, 13, 15]: ws.Cells(row, col).NumberFormat = num_fmt
                    
                row += 1

            # --- 세부 내역 테이블 ---
            headers = ["No.", "검사일자", "작업구간", "검사종류", "규격/자재", "근무형태", "실물량", "단위", "보정계수", "환산물량", 
                       "재료비", "직접인건비", "제경비", "기술료", "공급가액소계"]
            
            start_row = row + 2
            for col, h in enumerate(headers, start=1):
                cell = ws.Cells(start_row, col)
                cell.Value = h
                cell.Font.Bold = True
                cell.Interior.Color = 14277081
                cell.HorizontalAlignment = -4108
                cell.Borders.LineStyle = 1
            
            ws.Columns(1).ColumnWidth = 15
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
            
            categories_det = ["RT_B", "RT_A", "RT_A2", "UT", "PT"]
            display_names_det = {
                "RT_B": "RT (B필름)", "RT_A": "RT (A필름)", "RT_A2": "RT (A/2필름)",
                "UT": "UT", "PT": "PT"
            }
            
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
                    dates = sorted(list(set(data["date_list"])))
                    if len(dates) == 1:
                        date_str = dates[0]
                    else:
                        date_str = f"{dates[0]} ~ {dates[-1]}"
                    
                    ws.Cells(current_row, 1).Value = idx
                    ws.Cells(current_row, 2).Value = date_str
                    ws.Cells(current_row, 3).Value = key[0].split()[0] if isinstance(key[0], str) and key[0] else key[0]
                    ws.Cells(current_row, 4).Value = key[1]
                    ws.Cells(current_row, 5).Value = key[2]
                    ws.Cells(current_row, 6).Value = key[3]
                    ws.Cells(current_row, 7).Value = data["qty"]
                    ws.Cells(current_row, 8).Value = key[4]
                    ws.Cells(current_row, 9).Value = key[5]
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
                        adj_tech = exact_subtotal - data["mat_cost"] - data["lab_cost"] - data["overhead"]
                        
                        ws.Cells(current_row, 10).Value = round(data["adjusted_qty"], 2)
                        ws.Cells(current_row, 11).Value = data["mat_cost"]
                        ws.Cells(current_row, 12).Value = data["lab_cost"]
                        ws.Cells(current_row, 13).Value = data["overhead"]
                        ws.Cells(current_row, 14).Value = adj_tech
                        ws.Cells(current_row, 15).Value = exact_subtotal
                        
                        # Update the data dictionary so that sub_mat/sub_lab/sub_sub sums use the adjusted values
                        data["tech"] = adj_tech
                        data["subtotal"] = exact_subtotal
                    else:
                        ws.Cells(current_row, 10).Value = round(data["adjusted_qty"], 2)
                        ws.Cells(current_row, 11).Value = data["mat_cost"]
                        ws.Cells(current_row, 12).Value = data["lab_cost"]
                        ws.Cells(current_row, 13).Value = data["overhead"]
                        ws.Cells(current_row, 14).Value = data["tech"]
                        ws.Cells(current_row, 15).Value = data["subtotal"]
                    
                    unit_str = key[4]
                    for c in range(1, 16):
                        cell = ws.Cells(current_row, c)
                        cell.Borders.LineStyle = 1
                        if c <= 4 or c == 6 or c == 8: cell.HorizontalAlignment = -4108
                        elif c == 5 or c == 3: cell.HorizontalAlignment = -4131
                        elif c == 7: 
                            cell.NumberFormat = '#,##0.000;-#,##0.000;"-"' if unit_str == "M" else '#,##0;-#,##0;"-"'
                        elif c == 9: 
                            cell.NumberFormat = "0.0"
                        elif c == 10: 
                            cell.NumberFormat = '#,##0.00;-#,##0.00;"-"' if unit_str == "M" else '#,##0;-#,##0;"-"'
                        elif c >= 11: 
                            cell.NumberFormat = "#,##0"
                    
                    sub_mat += data["mat_cost"]; sub_lab += data["lab_cost"]
                    sub_ovr += data["overhead"]; sub_tech += data["tech"]
                    sub_sub += data["subtotal"]
                    idx += 1; current_row += 1
                
                ws.Range(ws.Cells(current_row, 1), ws.Cells(current_row, 10)).Merge()
                ws.Cells(current_row, 1).Value = f"[{display_names_det[g_type]}] 검사 소계"
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
            
            extra_items_map = []
            if len(extra_rows) >= 4:
                extra_items_map = [
                    ("장비손료", extra_rows[0]),
                    ("안전관리비", extra_rows[1]),
                    ("주재비 및 출장여비", extra_rows[2]),
                    ("도서인쇄비", extra_rows[3]),
                    ("엔지니어링 손해배상공제료", liability_row)
                ]
            else:
                extra_items_map = [
                    ("장비손료", self.equip_cost_var.get()),
                    ("안전관리비", self.safety_cost_var.get()),
                    ("주재비 및 출장여비", self.travel_cost_var.get()),
                    ("도서인쇄비", self.print_cost_var.get()),
                    ("엔지니어링 손해배상공제료", self.liability_cost_var.get())
                ]
            
            for item in extra_items_map:
                name = item[0]
                val_or_row = item[1]
                ws.Range(ws.Cells(current_row, 1), ws.Cells(current_row, 14)).Merge()
                ws.Cells(current_row, 1).Value = f"'+ {name}"
                ws.Cells(current_row, 1).HorizontalAlignment = -4152
                
                if val_or_row > 0 and val_or_row < 1000:
                    ws.Cells(current_row, 15).Formula = f"=K{val_or_row}"
                else:
                    ws.Cells(current_row, 15).Value = val_or_row if val_or_row > 0 else 0
                    
                ws.Cells(current_row, 15).NumberFormat = '#,##0;-#,##0;"-"'
                for c in range(1, 16): ws.Cells(current_row, c).Borders.LineStyle = 1
                current_row += 1
            
            # --- 총 공급가액 (검사합계 + 실비) ---
            ws.Range(ws.Cells(current_row, 1), ws.Cells(current_row, 14)).Merge()
            ws.Cells(current_row, 1).Value = "공급가액 총액"
            ws.Cells(current_row, 1).HorizontalAlignment = -4152
            ws.Cells(current_row, 1).Font.Bold = True
            if total_row > 0:
                ws.Cells(current_row, 15).Formula = f"=K{total_row}"
            else:
                ws.Cells(current_row, 15).Value = 0
            ws.Cells(current_row, 15).NumberFormat = '#,##0;-#,##0;"-"'
            ws.Cells(current_row, 15).Font.Bold = True
            for c in range(1, 16): ws.Cells(current_row, c).Borders.LineStyle = 1
            
            # --- 부가세 및 최종 청구액 ---
            current_row += 1
            ws.Range(ws.Cells(current_row, 1), ws.Cells(current_row, 14)).Merge()
            ws.Cells(current_row, 1).Value = "'+ 부가가치세 (10%)"
            ws.Cells(current_row, 1).HorizontalAlignment = -4152
            ws.Cells(current_row, 15).Formula = f"=TRUNC(O{current_row-1}*0.1)"
            ws.Cells(current_row, 15).NumberFormat = '#,##0;-#,##0;"-"'
            for c in range(1, 16): ws.Cells(current_row, c).Borders.LineStyle = 1
            
            current_row += 1
            ws.Range(ws.Cells(current_row, 1), ws.Cells(current_row, 14)).Merge()
            ws.Cells(current_row, 1).Value = "최 종 기 성 청 구 액"
            ws.Cells(current_row, 1).HorizontalAlignment = -4152
            ws.Cells(current_row, 1).Font.Bold = True
            ws.Cells(current_row, 1).Font.Size = 12
            
            ws.Cells(current_row, 15).Formula = f"=O{current_row-2}+O{current_row-1}"
            ws.Cells(current_row, 15).NumberFormat = '#,##0;-#,##0;"-"'
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
        tree.heading("Pipeline", text="수송배관 (열배관)")
        tree.heading("Plant", text="플랜트 (관리소)")
        tree.heading("Total", text="총계")
        
        tree.column("Type", width=120, anchor=tk.CENTER)
        tree.column("Pipeline", width=100, anchor=tk.E)
        tree.column("Plant", width=100, anchor=tk.E)
        tree.column("Total", width=100, anchor=tk.E)
        
        pipe = CONTRACT_QTY.get("열배관", {})
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

if __name__ == '__main__':
    root = tk.Tk()
    app = NDTCalculatorTab(root)
    app.pack(fill=tk.BOTH, expand=True)
    root.mainloop()
