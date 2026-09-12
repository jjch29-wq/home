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
                unit_price = res.get("unit_price", 0)
                if unit_price == 0 and res.get("qty", 0) > 0:
                    unit_price = int(res.get("subtotal", 0) / res.get("qty"))
                    
                self.tree.insert("", tk.END, values=(
                    res.get("date", ""), res.get("company", ""), res.get("loc", ""), res.get("ndt_type", ""), res.get("work_time", ""), 
                    res.get("material_type", ""), f"{res.get('qty', 0):.1f}", res.get("unit", ""),
                    f"{unit_price:,}", f"{res.get('subtotal', 0):,}"
                ))
            if hasattr(self, 'update_qty_summary'):
                self.update_qty_summary()

    def save_billing_records(self):
        try:
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
        
        ttk.Label(info_frame1, text="• 업체명:", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=(10, 0))
        self.company_var = tk.StringVar(value="")
        ttk.Entry(info_frame1, textvariable=self.company_var, width=15).pack(side=tk.LEFT, padx=5)
        
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
        
        # ttk.Label(round_frame, text="  |  기성청구 기간:", font=("Arial", 11, "bold")).pack(side=tk.LEFT, padx=(15, 5))
        # self.billing_period_var = tk.StringVar(value="")
        # ttk.Entry(round_frame, textvariable=self.billing_period_var, width=25, font=("Arial", 11)).pack(side=tk.LEFT)
        
        ttk.Button(round_frame, text="다음 회차로 이월하기 (전회 누적 & 금회 초기화)", command=self.carry_over_round).pack(side=tk.RIGHT)
        ttk.Button(round_frame, text="이전 백업 불러오기 (.ndt)", command=self.load_project).pack(side=tk.RIGHT, padx=10)
        # ttk.Button(round_frame, text="✨ 엑셀 보고서 생성기 열기", command=self.open_report_hub).pack(side=tk.RIGHT, padx=5)
        
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
                    
                    val = 0
                    if isinstance(CONTRACT_QTY.get(loc), dict) and isinstance(CONTRACT_QTY[loc].get(t_time), dict):
                        val = CONTRACT_QTY[loc][t_time].get(m_key, 0)
                    else:
                        flat_key = f"{m_key}_야간" if t_time == "야간" else m_key
                        val = CONTRACT_QTY.get(flat_key, 0)

                    formatted_val = f"{int(val):,}" if float(val).is_integer() else f"{float(val):,.2f}"
                    c_qty = tk.StringVar(value=formatted_val)
                    p_qty = tk.StringVar(value="0")
                    curr_qty = tk.StringVar(value="0")
                    
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

                    amt = float(val) * unit_cost
                    c_var = tk.StringVar(value=f"{int(amt):,}")
                    p_var = tk.StringVar(value="0")
                    c_price_var = tk.StringVar(value=f"{int(unit_cost):,}")
                    rem_qty = tk.StringVar(value=formatted_val)
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
            
        update_total_contract_amt()
            
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
        ttk.Label(lbl_frame, text="[ 월별 작업 기록 목록 ]", font=("Arial", 11, "bold")).pack(side=tk.LEFT)
        
        ttk.Label(lbl_frame, text="  |  기성청구 기간: ", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=(10, 2))
        
        self.billing_start_date = tk.StringVar(value=datetime.now().strftime('%Y-%m-01'))
        self.billing_end_date = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        
        try:
            self.ent_billing_start = DateEntry(lbl_frame, textvariable=self.billing_start_date, width=12, date_pattern='yyyy-mm-dd', background='darkblue', foreground='white', borderwidth=2)
            self.ent_billing_start.pack(side=tk.LEFT)
        except Exception:
            self.ent_billing_start = ttk.Entry(lbl_frame, textvariable=self.billing_start_date, width=12)
            self.ent_billing_start.pack(side=tk.LEFT)
            
        ttk.Label(lbl_frame, text=" ~ ", font=("Arial", 10, "bold")).pack(side=tk.LEFT)
        
        try:
            self.ent_billing_end = DateEntry(lbl_frame, textvariable=self.billing_end_date, width=12, date_pattern='yyyy-mm-dd', background='darkblue', foreground='white', borderwidth=2)
            self.ent_billing_end.pack(side=tk.LEFT)
        except Exception:
            self.ent_billing_end = ttk.Entry(lbl_frame, textvariable=self.billing_end_date, width=12)
            self.ent_billing_end.pack(side=tk.LEFT)
            
        ttk.Button(lbl_frame, text="선택", command=self.import_from_daily_db).pack(side=tk.LEFT, padx=(5, 0))
            
        ttk.Button(lbl_frame, text="기성청구", command=self.export_to_excel).pack(side=tk.RIGHT)
        ttk.Button(lbl_frame, text="기록 초기화", command=self.clear_records).pack(side=tk.RIGHT, padx=5)
        ttk.Button(lbl_frame, text="일일 장부에서 연동", command=self.import_from_daily_db).pack(side=tk.RIGHT, padx=5)
        ttk.Button(lbl_frame, text="선택 삭제", command=self.delete_selected_records).pack(side=tk.RIGHT)

        tree_container = ttk.Frame(bottom_frame)
        tree_container.pack(fill=tk.BOTH, expand=True)

        columns = ("date", "company", "loc", "type", "time", "mat", "qty", "unit", "unit_price", "total_amt")
        self.tree = ttk.Treeview(tree_container, columns=columns, show="headings", height=8)
        
        self.tree.heading("date", text="일자", anchor="center")
        self.tree.heading("company", text="업체명", anchor="center")
        self.tree.heading("loc", text="구간/위치", anchor="center")
        self.tree.heading("type", text="종류", anchor="center")
        self.tree.heading("time", text="형태", anchor="center")
        self.tree.heading("mat", text="자재", anchor="center")
        self.tree.heading("qty", text="실물량", anchor="center")
        self.tree.heading("unit", text="단위", anchor="center")
        self.tree.heading("unit_price", text="단가(원)", anchor="center")
        self.tree.heading("total_amt", text="공급가액(원)", anchor="center")
        
        default_widths = {
            "date": 80, "company": 80, "loc": 120, "type": 40, "time": 40, "mat": 90, 
            "qty": 40, "unit": 40,
            "unit_price": 90, "total_amt": 90
        }
        saved_widths = CONFIG.get("TREE_WIDTHS", {})
        
        self.tree.column("date", width=saved_widths.get("date", default_widths["date"]), anchor="center")
        self.tree.column("company", width=saved_widths.get("company", default_widths["company"]), anchor="center")
        self.tree.column("loc", width=saved_widths.get("loc", default_widths["loc"]), anchor="w")
        self.tree.column("type", width=saved_widths.get("type", default_widths["type"]), anchor="center")
        self.tree.column("time", width=saved_widths.get("time", default_widths["time"]), anchor="center")
        self.tree.column("mat", width=saved_widths.get("mat", default_widths["mat"]), anchor="center")
        self.tree.column("qty", width=saved_widths.get("qty", default_widths["qty"]), anchor="center")
        self.tree.column("unit", width=saved_widths.get("unit", default_widths["unit"]), anchor="center")
        self.tree.column("unit_price", width=saved_widths.get("unit_price", default_widths["unit_price"]), anchor="center")
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
        company_str = self.company_var.get()
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
            "company": company_str,
            "loc": display_loc,
            "ndt_type": ndt_type,
            "work_time": work_time,
            "material_type": material_type,
            "qty": qty,
            "unit": unit_str,
            "corr": corr,
            "adjusted_qty": adjusted_qty,
            "unit_price": costs["unit_price"],
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
                    start_str = getattr(self, "billing_start_date", self.main_app.ent_daily_start_date).get().strip()
                    end_str = getattr(self, "billing_end_date", self.main_app.ent_daily_end_date).get().strip()
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
                site_str = str(row.get('Site', ''))
                item_name = str(row.get('검사품명', ''))
                if item_name and item_name != 'nan':
                    loc_str = f"{site_str}_{item_name}"
                else:
                    loc_str = site_str
                company_str = str(row.get('업체명', ''))
                if not company_str or company_str == 'nan': company_str = ''
                ndt_type = str(row.get('검사방법', 'RT'))
                work_time = str(row.get('작업형태', '일반'))
                if work_time not in ['일반', '야간', '휴일']: work_time = '일반'
                material_type = str(row.get('Material', ''))
                qty = float(row.get('검사량', 0.0) if not pd.isna(row.get('검사량')) else 0.0)
                if qty == 0.0: qty = float(row.get('Usage', 0.0) if not pd.isna(row.get('Usage')) else 0.0)
                
                self.date_var.set(date_str)
                self.company_var.set(company_str)
                if '관리소' in loc_str or '플랜트' in loc_str:
                    self.loc_type_var.set('플랜트(관리소)')
                else:
                    self.loc_type_var.set('열배관')
                self.loc_var.set(loc_str)
                self.ndt_type_var.set(ndt_type)
                self.work_time_var.set(work_time)
                
                # [NEW] PAUT의 경우 관경을 우선적으로 확인하여 자동 맵핑
                pipe_size_str = ""
                for k in row.keys():
                    if 'Inch' in str(k) or '관경' in str(k):
                        pipe_size_str = str(row.get(k, '')).strip()
                        break
                if ndt_type == 'PAUT' and pipe_size_str and pipe_size_str != 'nan':
                    import re
                    m = re.search(r'(\d+)', pipe_size_str)
                    if m:
                        p_val = int(m.group(1))
                        if p_val >= 300: material_type = "300A 이상"
                        elif p_val == 250: material_type = "250A"
                        elif p_val == 200: material_type = "200A"
                        elif 125 <= p_val <= 150: material_type = "150A-125A"
                        else: material_type = "100A 이하"

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
            
            unit_price = res.get("unit_price", 0)
            if unit_price == 0 and res.get("qty", 0) > 0:
                unit_price = int(res.get("subtotal", 0) / res.get("qty"))
                
            self.tree.insert("", tk.END, values=(
                res["date"], res.get("company", ""), res["loc"], res["ndt_type"], res["work_time"], 
                res["material_type"], f"{res['qty']:.1f}", res["unit"],
                f"{unit_price:,}", f"{res['subtotal']:,}"
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
                
            # 전체 행 높이 일괄 적용 (표시를 시원하게 하여 갑지와 밸런스 맞춤)
            ws.Rows(1).RowHeight = 45 # 제목
            ws.Rows(4).RowHeight = 25 # 서브 타이틀
            ws.Rows(6).RowHeight = 25 # 헤더 1
            ws.Rows(7).RowHeight = 25 # 헤더 2
            for r in range(8, row):
                ws.Rows(r).RowHeight = 25 # 데이터 행

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
            ws.PageSetup.Orientation = 2 # xlLandscape
            ws.PageSetup.Zoom = False
            ws.PageSetup.FitToPagesWide = 1
            ws.PageSetup.FitToPagesTall = 1
            ws.PageSetup.LeftMargin = 20
            ws.PageSetup.RightMargin = 20
            ws.PageSetup.TopMargin = 20
            ws.PageSetup.BottomMargin = 20
            ws.PageSetup.CenterHorizontally = True # 페이지 가로 가운데 정렬
            ws.PageSetup.CenterVertically = True # 페이지 세로 가운데 정렬 (위아래 여백 동일하게)
            
            # --- 열 너비 자동 맞춤 및 여백 넉넉하게 확장 ---
            ws.Columns("A:O").AutoFit()
            
            # 너무 좁은 열들은 최소 너비 확보하여 시원하게 표시
            if ws.Columns(1).ColumnWidth < 8: ws.Columns(1).ColumnWidth = 8
            if ws.Columns(2).ColumnWidth < 8: ws.Columns(2).ColumnWidth = 8
            if ws.Columns(3).ColumnWidth < 20: ws.Columns(3).ColumnWidth = 20
            if ws.Columns(4).ColumnWidth < 6: ws.Columns(4).ColumnWidth = 6
            if ws.Columns(5).ColumnWidth < 12: ws.Columns(5).ColumnWidth = 12
            
            # 금액/수량 열(F~O) 넓게 설정하여 A4 가로폭 채우기
            for c in range(6, 16):
                if ws.Columns(c).ColumnWidth < 12:
                    ws.Columns(c).ColumnWidth = 12
            # --- 표지 (청구서 갑지) 생성 ---
            ws_cover = wb.Sheets.Add(ws)
            ws_cover.Name = "청구서(갑지)"
            
            # --- 페이지 가로 모드 및 여백 설정 ---
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
            cell = ws_cover.Range(f"A{table_start}")
            cell.Value = "구 분"
            cell.Font.Bold = True
            cell.Font.Size = 14
            cell.HorizontalAlignment = -4108
            cell.VerticalAlignment = -4108
            cell.Interior.Color = 10066329
            cell.Font.Color = 16777215
            for col_i in [1, 2]: ws_cover.Cells(table_start, col_i).Borders.LineStyle = 1
            
            for col_i, h in enumerate(table_headers, start=3):
                cell = ws_cover.Cells(table_start, col_i)
                cell.Value = h
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
                
                for c_idx, v in enumerate(vals, start=3):
                    cell_val = ws_cover.Cells(row_idx, c_idx)
                    cell_val.Value = f"{v:,}"
                    cell_val.Font.Size = 14
                    cell_val.HorizontalAlignment = -4152
                    cell_val.VerticalAlignment = -4108
                    cell_val.Borders.LineStyle = 1
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
            for edge in (7, 8, 9, 10): # xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight
                outer_range.Borders(edge).LineStyle = 1
                outer_range.Borders(edge).Weight = 4 # xlThick
            
            # --- 업체별 기성요약 시트 생성 ---
            ws_summary = wb.Sheets.Add(None, ws)
            ws_summary.Name = "업체별 기성요약"
            
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
            
            ws_summary.Range("A1:G2").Merge()
            ws_summary.Range("A1").Value = f"제 {round_val} 회 비파괴검사기술용역 업체별 기성요약"
            ws_summary.Range("A1").Font.Size = 16
            ws_summary.Range("A1").Font.Bold = True
            ws_summary.Range("A1").HorizontalAlignment = -4108
            ws_summary.Range("A1").VerticalAlignment = -4108
            
            # --- 실제 시공업체 매핑 ---
            # target_records의 'company'가 '한국지역난방공사' 등으로 일괄 지정되어 있으므로,
            # daily_work_history.json의 ndt_results에서 해당 일자/검사방법의 실제 업체를 찾아 매핑합니다.
            import json
            history_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'daily_work_history.json')
            history_data = {}
            if os.path.exists(history_path):
                try:
                    with open(history_path, 'r', encoding='utf-8') as f:
                        history_data = json.load(f)
                except: pass

            global_comps = {}
            for d, d_data in history_data.items():
                if "ndt_results" in d_data:
                    for nr in d_data["ndt_results"]:
                        c = str(nr.get("업체", "")).strip()
                        if c and c != "미지정" and "지역난방공사" not in c:
                            global_comps[c] = global_comps.get(c, 0) + 1
            global_fallback = max(global_comps.items(), key=lambda x: x[1])[0] if global_comps else "미지정"

            for r in target_records:
                t_date = r["date"]
                n_type = r["ndt_type"]
                
                actual_comp = r.get("company", "미지정")
                if not actual_comp or str(actual_comp).strip() == "" or "지역난방공사" in str(actual_comp):
                    actual_comp = "미지정"
                
                if t_date in history_data and "ndt_results" in history_data[t_date]:
                    comps = {}
                    for nr in history_data[t_date]["ndt_results"]:
                        nr_type = str(nr.get("검사방법", "")).strip()
                        if not nr_type:
                            continue
                        if nr_type.startswith(n_type) or n_type.startswith(nr_type):
                            c = nr.get("업체", "미지정")
                            if str(c).strip() == "": c = "미지정"
                            comps[c] = comps.get(c, 0) + 1
                    if comps:
                        best_comp = max(comps.items(), key=lambda x: x[1])[0]
                        if best_comp != "미지정":
                            actual_comp = best_comp
                
                if actual_comp == "미지정" and r.get("company") and str(r.get("company")).strip() != "" and "지역난방공사" not in str(r.get("company")):
                    actual_comp = r.get("company")
                    
                if actual_comp == "미지정" and global_fallback != "미지정":
                    actual_comp = global_fallback
                    
                r["actual_company"] = actual_comp

            sum_row = 4
            contract_category_order = [
                cat for cat in categories if cat in self.contract_vars
            ]
            contract_item_numbers = {
                cat: idx for idx, cat in enumerate(contract_category_order, start=1)
            }
            company_subtotal_rows = []
            companies = sorted(list(set(r.get("actual_company", "미지정") for r in target_records)))
            if not companies:
                companies = ["미지정"]
                
            for comp in companies:
                ws_summary.Cells(sum_row, 1).Value = f"■ 업체명 : {comp}"
                ws_summary.Cells(sum_row, 1).Font.Bold = True
                ws_summary.Cells(sum_row, 1).Font.Size = 12
                sum_row += 1
                
                comp_records = [r for r in target_records if r.get("actual_company", "미지정") == comp]
                
                # --- 통합 기성요약 표 (섹션 -> 라인번호 -> 규격) ---
                work_summary = {}
                for r in comp_records:
                    t_date = str(r.get("date", ""))
                    n_type = str(r.get("ndt_type", "")).strip()
                    m_type = str(r.get("material_type", "")).strip()
                    w_time = str(r.get("work_time", "")).strip()
                    r_qty = r.get("qty", 0.0)
                    r_amt = r.get("subtotal", 0)
                    
                    if n_type == "PAUT":
                        spec = f"위상배열초음파검사(PAUT) {m_type.removeprefix('PAUT_')}"
                    elif n_type == "RT":
                        spec = f"방사선투과검사(RT) {m_type.removeprefix('RT_')}"
                    elif n_type == "MT":
                        spec = "자분탐상검사(MT)"
                    else:
                        spec = "액체침투탐상검사(PT)"
                    if w_time == "야간":
                        spec += " 야간"
                        
                    unit = r.get("unit", "")
                    
                    loc_type = "플랜트(관리소)" if "관리소" in r.get("loc", "") or "플랜트" in r.get("loc_type", r.get("loc", "")) else "열배관"
                    cat_key = f"{loc_type}_{w_time}_{n_type}_{m_type}"
                    c_price = self.contract_vars.get(cat_key, {}).get("c_price", 0)
                    if c_price == 0 and r_qty > 0:
                        c_price = r_amt / r_qty
                    
                    matched_results = []
                    if t_date in history_data:
                        for nr in history_data[t_date].get("ndt_results", []):
                            if id(nr) in billed_nrs: continue
                            method = str(nr.get("검사방법", "")).strip()
                            c = str(nr.get("업체", "")).strip()
                            pipe_size = str(nr.get("관경", "")).strip()
                            
                            size_matched = True
                            if m_type and pipe_size:
                                import re
                                m1 = re.search(r'(\d+)', pipe_size)
                                m2_all = re.findall(r'(\d+)', m_type)
                                if m1 and m2_all:
                                    p_val = int(m1.group(1))
                                    t_vals = [int(x) for x in m2_all]
                                    if "이상" in m_type:
                                        size_matched = (p_val >= t_vals[0])
                                    elif "미만" in m_type:
                                        size_matched = (p_val < t_vals[0])
                                    elif "이하" in m_type:
                                        size_matched = (p_val <= t_vals[0])
                                    elif len(t_vals) >= 2:
                                        min_val = min(t_vals)
                                        max_val = max(t_vals)
                                        size_matched = (min_val <= p_val <= max_val)
                                    else:
                                        size_matched = (p_val == t_vals[0])
                                        
                            if method and (method.startswith(n_type) or n_type.startswith(method)) and size_matched:
                                if c == comp or c == "미지정" or not c:
                                    matched_results.append(nr)
                                    billed_nrs.add(id(nr))
                                    
                    if not matched_results:
                        continue  # 작업/감독일보에 매칭되는 내역이 없으면 기성청구에서 제외
                    else:
                        sub_groups = {}
                        for nr in matched_results:
                            sec = str(nr.get("구간", "")).strip() or "미지정"
                            l_no = str(nr.get("라인번호", "")).strip() or "미지정"
                            j_no = str(nr.get("Joint No.", "")).strip()
                            method = str(nr.get("검사방법", "")).strip()
                            
                            sg_key = (sec, l_no)
                            if sg_key not in sub_groups:
                                sub_groups[sg_key] = {"joints": set(), "rows": 0, "length": 0.0}
                            
                            if j_no:
                                sub_groups[sg_key]["joints"].add(j_no)
                            sub_groups[sg_key]["rows"] += 1
                            
                            if "PAUT" in method or "UT" in method:
                                try: sub_groups[sg_key]["length"] += float(nr.get("PAUT", 0) or 0)
                                except: pass
                            elif "RT" in method:
                                try: sub_groups[sg_key]["length"] += float(nr.get("RT", 0) or 0)
                                except: pass
                            elif "PT" in method:
                                try: sub_groups[sg_key]["length"] += float(nr.get("PT", 0) or 0)
                                except: pass
                            elif "MT" in method:
                                try: sub_groups[sg_key]["length"] += float(nr.get("MT", 0) or 0)
                                except: pass
                                
                        total_length = sum(sg["length"] for sg in sub_groups.values())
                        total_places = sum(max(len(sg["joints"]), sg["rows"]) for sg in sub_groups.values())
                        
                        sg_items = list(sub_groups.items())
                        for i, (sg_key, sg_data) in enumerate(sg_items):
                            sec, l_no = sg_key
                            group_key = (sec, l_no, spec, unit, c_price)
                            if group_key not in work_summary:
                                work_summary[group_key] = {"places": 0, "qty": 0.0}
                                
                            places = max(len(sg_data["joints"]), sg_data["rows"])
                            work_summary[group_key]["places"] += places
                            
                            my_qty = sg_data["length"]
                            work_summary[group_key]["qty"] += my_qty

                # 테이블 헤더 렌더링
                headers_sum = {
                    1: "섹션", 2: "라인번호", 3: "규격", 4: "개소", 5: "단가",
                    6: "검사수량", 7: "금액"
                }
                for col, h in headers_sum.items():
                    cell = ws_summary.Cells(sum_row, col)
                    cell.Value = h
                    cell.Font.Bold = True
                    cell.Interior.Color = 14277081
                    cell.HorizontalAlignment = -4108
                    cell.Borders.LineStyle = 1
                sum_row += 1
                start_data_row = sum_row
                
                # 정렬: 섹션 -> 라인번호 -> 규격
                sorted_keys = sorted(work_summary.keys(), key=lambda x: (x[0], x[1], x[2]))
                
                # 단수 조정(Fraction Adjustment): 규격별로 수량을 합산하여 총 금액을 구한 뒤, 각 행에 분배하여 1원 단위 오차 방지
                spec_totals = {}
                for key in sorted_keys:
                    sec, l_no, spec, unit, c_price = key
                    qty = work_summary[key]["qty"]
                    spec_key = (spec, c_price)
                    if spec_key not in spec_totals:
                        spec_totals[spec_key] = {"qty": 0.0, "total_amt": 0, "allocated_amt": 0, "rows": []}
                    spec_totals[spec_key]["qty"] += qty
                    spec_totals[spec_key]["rows"].append(key)
                
                for spec_key, totals in spec_totals.items():
                    totals["total_amt"] = int(totals["qty"] * spec_key[1])
                    rows = totals["rows"]
                    for i, key in enumerate(rows):
                        row_qty = work_summary[key]["qty"]
                        if i == len(rows) - 1:
                            row_amt = totals["total_amt"] - totals["allocated_amt"]
                        else:
                            row_amt = int(row_qty * spec_key[1])
                            totals["allocated_amt"] += row_amt
                        work_summary[key]["calculated_amt"] = row_amt
                
                for key in sorted_keys:
                    data = work_summary[key]
                    sec, l_no, spec, unit, c_price = key
                    calculated_amt = data.get("calculated_amt", 0)
                    
                    if data["qty"] == 0 and calculated_amt == 0: continue
                    
                    ws_summary.Cells(sum_row, 1).Value = sec
                    ws_summary.Cells(sum_row, 2).Value = l_no
                    ws_summary.Cells(sum_row, 3).Value = f"{spec} ({unit})"
                    ws_summary.Cells(sum_row, 4).Value = int(data["places"])
                    
                    ws_summary.Cells(sum_row, 5).Value = int(c_price)
                    ws_summary.Cells(sum_row, 5).NumberFormat = "#,##0"
                    
                    if unit == "M":
                        ws_summary.Cells(sum_row, 6).Value = round(data["qty"], 4)
                        ws_summary.Cells(sum_row, 6).NumberFormat = '#,##0.0000;-#,##0.0000;"-"'
                    else:
                        ws_summary.Cells(sum_row, 6).Value = int(data["qty"])
                        ws_summary.Cells(sum_row, 6).NumberFormat = '#,##0;-#,##0;"-"'
                        
                    ws_summary.Cells(sum_row, 7).Value = calculated_amt
                    ws_summary.Cells(sum_row, 7).NumberFormat = '#,##0;-#,##0;"-"'
                    
                    for col in range(1, 8):
                        cell = ws_summary.Cells(sum_row, col)
                        cell.Borders.LineStyle = 1
                        if col in (1, 2, 4): 
                            cell.HorizontalAlignment = -4108
                        elif col == 3: 
                            cell.HorizontalAlignment = -4131
                    
                    sum_row += 1
                    
                # 소계 렌더링
                ws_summary.Range(ws_summary.Cells(sum_row, 1), ws_summary.Cells(sum_row, 5)).Merge()
                ws_summary.Cells(sum_row, 1).Value = "소 계"
                ws_summary.Cells(sum_row, 1).HorizontalAlignment = -4108
                ws_summary.Cells(sum_row, 1).Font.Bold = True
                ws_summary.Cells(sum_row, 1).Interior.Color = 15987699
                

                qty_sum_formula = f"=SUM(F{start_data_row}:F{sum_row-1})" if sum_row > start_data_row else "0"
                ws_summary.Cells(sum_row, 6).Formula = qty_sum_formula
                ws_summary.Cells(sum_row, 6).NumberFormat = '#,##0.0000;-#,##0.0000;"-"'
                ws_summary.Cells(sum_row, 6).Font.Bold = True
                ws_summary.Cells(sum_row, 6).Interior.Color = 15987699
                sum_formula = f"=SUM(G{start_data_row}:G{sum_row-1})" if sum_row > start_data_row else "0"
                ws_summary.Cells(sum_row, 7).Formula = sum_formula
                ws_summary.Cells(sum_row, 7).NumberFormat = '#,##0;-#,##0;"-"'
                ws_summary.Cells(sum_row, 7).Font.Bold = True
                ws_summary.Cells(sum_row, 7).Interior.Color = 15987699
                
                for col in range(1, 8):
                    ws_summary.Cells(sum_row, col).Borders.LineStyle = 1

                company_subtotal_rows.append(sum_row)
                    
                sum_row += 3
            
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

            ws_summary.Columns(1).ColumnWidth = 34
            ws_summary.Columns(2).ColumnWidth = 25
            ws_summary.Columns(3).ColumnWidth = 35
            ws_summary.Columns(4).ColumnWidth = 8
            ws_summary.Columns(5).ColumnWidth = 12
            ws_summary.Columns(6).ColumnWidth = 12
            ws_summary.Columns(7).ColumnWidth = 15
            
            # --- 업체별 수량내역 시트 생성 ---
            ws_cont = wb.Sheets.Add(None, ws_summary)
            ws_cont.Name = "업체별 수량내역"
            
            # --- 페이지 가로 모드 및 폭 1장 맞춤 설정 ---
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
            
            ws_cont.Range("A1:K2").Merge()
            ws_cont.Range("A1").Value = f"제 {round_val} 회 기성청구 업체별 수량내역"
            ws_cont.Range("A1").Font.Size = 16
            ws_cont.Range("A1").Font.Bold = True
            ws_cont.Range("A1").HorizontalAlignment = -4108
            ws_cont.Range("A1").VerticalAlignment = -4108
            
            headers_cont = ["No.", "업체명", "검사방법", "구간", "라인번호", "Joint No.", "관경", "두께", "용접사", "수량", "결과"]
            for col, h in enumerate(headers_cont, start=1):
                cell = ws_cont.Cells(4, col)
                cell.Value = h
                cell.Font.Bold = True
                cell.Interior.Color = 14277081
                cell.HorizontalAlignment = -4108
                cell.Borders.LineStyle = 1
                
            ws_cont.Columns(1).ColumnWidth = 8
            ws_cont.Columns(2).ColumnWidth = 20
            ws_cont.Columns(3).ColumnWidth = 10
            ws_cont.Columns(4).ColumnWidth = 15
            ws_cont.Columns(5).ColumnWidth = 35
            ws_cont.Columns(6).ColumnWidth = 10
            ws_cont.Columns(7).ColumnWidth = 10
            ws_cont.Columns(8).ColumnWidth = 10
            ws_cont.Columns(9).ColumnWidth = 12
            ws_cont.Columns(10).ColumnWidth = 12
            ws_cont.Columns(11).ColumnWidth = 10
            
            # 헤더에 자동 필터 적용 (지원 여부에 따라 선택적 적용)
            try:
                ws_cont.Range("A4:K4").AutoFilter(Field=1)
            except:
                pass
            
            cont_row = 5
            
            import json
            history_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'daily_work_history.json')
            history_data = {}
            if os.path.exists(history_path):
                try:
                    with open(history_path, 'r', encoding='utf-8') as f:
                        history_data = json.load(f)
                except: pass
                
            target_dates = set(r["date"] for r in target_records)
            all_ndt_results = []
            for t_date in target_dates:
                if t_date in history_data and "ndt_results" in history_data[t_date]:
                    all_ndt_results.extend(history_data[t_date]["ndt_results"])
                    
            all_ndt_results.sort(key=lambda x: (str(x.get("업체", "")), str(x.get("검사방법", "")), str(x.get("구간", "")), str(x.get("라인번호", ""))))
            
            total_points = 0
            total_meters = 0.0
            
            data_start_row = cont_row  # 데이터 시작 행 기록
            idx_cont = 1
            for r in all_ndt_results:
                if not str(r.get("업체", "")).strip() and not str(r.get("Joint No.", "")).strip():
                    continue
                    
                ws_cont.Cells(cont_row, 1).Value = idx_cont
                ws_cont.Cells(cont_row, 2).Value = r.get("업체", "")
                ws_cont.Cells(cont_row, 3).Value = r.get("검사방법", "")
                ws_cont.Cells(cont_row, 4).Value = r.get("구간", "")
                ws_cont.Cells(cont_row, 5).Value = r.get("라인번호", "")
                ws_cont.Cells(cont_row, 6).Value = r.get("Joint No.", "")
                ws_cont.Cells(cont_row, 7).Value = r.get("관경", "")
                ws_cont.Cells(cont_row, 8).Value = r.get("두께", "")
                ws_cont.Cells(cont_row, 9).Value = r.get("용접사", "")
                
                m_type = str(r.get("검사방법", "")).strip()
                if "PAUT" in m_type:
                    qty = r.get("PAUT", "")
                elif "RT" in m_type:
                    qty = r.get("RT_OR", "")
                    if not qty: qty = r.get("RT_RE", "")
                    if not qty: qty = "1"
                elif "MT" in m_type:
                    qty = r.get("MT", "")
                elif "PT" in m_type:
                    qty = r.get("PT", "")
                else:
                    qty = r.get(m_type, "")
                    
                try:
                    q_val = float(qty) if str(qty).strip() else 0.0
                except ValueError:
                    q_val = 0.0
                    
                if "PAUT" in m_type or "UT" in m_type:
                    total_meters += q_val
                else:
                    total_points += q_val
                
                # [FIX] 수량을 숫자로 저장해야 SUBTOTAL 수식이 작동함
                ws_cont.Cells(cont_row, 10).Value = q_val if q_val != 0.0 else (qty if str(qty).strip() else "")
                ws_cont.Cells(cont_row, 10).NumberFormat = '#,##0.0000;-#,##0.0000;"-"'
                ws_cont.Cells(cont_row, 11).Value = r.get("결과", "")
                
                for c in range(1, 12):
                    cell = ws_cont.Cells(cont_row, c)
                    cell.Borders.LineStyle = 1
                    cell.HorizontalAlignment = -4108
                
                idx_cont += 1
                cont_row += 1
            
            data_end_row = cont_row - 1  # 데이터 마지막 행
                
            if all_ndt_results:
                ws_cont.Cells(cont_row, 1).Value = "총 누적 물량"
                ws_cont.Range(ws_cont.Cells(cont_row, 1), ws_cont.Cells(cont_row, 9)).Merge()
                ws_cont.Cells(cont_row, 1).HorizontalAlignment = -4108
                ws_cont.Cells(cont_row, 1).Font.Bold = True
                ws_cont.Cells(cont_row, 1).Interior.Color = 14277081
                
                # [FIX] 필터 시 자동 합산: SUBTOTAL(103)=COUNTA(가시행), SUBTOTAL(9)=SUM(가시행)
                # 셀 10: 개소 수 (필터된 행 수)
                ws_cont.Cells(cont_row, 10).Formula = (
                    f'=TEXT(SUBTOTAL(9,J{data_start_row}:J{data_end_row}),"0.0000")'
                )
                ws_cont.Cells(cont_row, 10).Font.Bold = True
                ws_cont.Cells(cont_row, 10).Font.Color = 255
                ws_cont.Cells(cont_row, 10).HorizontalAlignment = -4108
                ws_cont.Cells(cont_row, 10).Interior.Color = 14277081
                ws_cont.Cells(cont_row, 10).Borders.LineStyle = 1
                
                # 셀 11: 검사량 합계 (필터된 수량 합)
                ws_cont.Cells(cont_row, 11).Formula = (
                    f'=""'
                )
                ws_cont.Cells(cont_row, 11).Font.Bold = True
                ws_cont.Cells(cont_row, 11).Font.Color = 255
                ws_cont.Cells(cont_row, 11).HorizontalAlignment = -4108
                ws_cont.Cells(cont_row, 11).Interior.Color = 14277081
                ws_cont.Cells(cont_row, 11).Borders.LineStyle = 1
                
                for c in range(1, 10):
                    ws_cont.Cells(cont_row, c).Borders.LineStyle = 1
                    if c < 10:
                        ws_cont.Cells(cont_row, c).Interior.Color = 14277081
                
                cont_row += 1
            
            ws_cover.Select()
            
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
