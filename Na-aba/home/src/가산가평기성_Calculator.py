import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime, timedelta
import os
import json
import win32com.client as win32
from tkcalendar import DateEntry
import pandas as pd
import re
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

class WorkerCompositeWidget(ttk.Frame):
    """
    Composite widget for Worker selection: [Name] with Autocomplete
    """
    def __init__(self, parent, enable_autocomplete=False, user_list=None, **kwargs):
        super().__init__(parent)
        
        # Worker Name selection
        name_width = kwargs.pop('width', 15)
        self.cb_name = ttk.Combobox(self, width=name_width, **kwargs)
        self.cb_name.pack(side='left', fill='x', expand=True)
        
    def get(self):
        """Return clean name"""
        return self.cb_name.get().strip()

    def set(self, value):
        """Set name, cleaning off any (Shift) prefixes if present"""
        if not value:
            self.cb_name.set("")
            return
            
        import re
        # Progressively migrate: if data still has (Shift) prefix, strip it for the name field
        match = re.match(r"\((주간|야간|휴일|주야간)\)\s*(.*)", str(value))
        if match:
            self.cb_name.set(match.group(2).strip())
        else:
            self.cb_name.set(str(value).strip())

    def bind(self, sequence=None, func=None, add=None):
        self.cb_name.bind(sequence, func, add)

    def current(self, newindex=None):
        return self.cb_name.current(newindex)
        
    def config(self, **kwargs):
        self.cb_name.config(**kwargs)

    def __setitem__(self, key, value):
        self.cb_name[key] = value

    def __getitem__(self, key):
        return self.cb_name[key]

class WorkerDataGroup(ttk.Frame):
    """
    Unified widget for a worker's record: [Name] [Shift] [WorkTime] [OT]
    """
    def __init__(self, parent, worker_index, users_list, time_list=None, enable_autocomplete=False, **kwargs):
        super().__init__(parent, padding=2) # Reduced padding for compact layout
        self.worker_index = worker_index
        
        # 1. Name selection (WorkerCompositeWidget now handles only name)
        self.composite = WorkerCompositeWidget(
            self, width=12, values=users_list, 
            enable_autocomplete=enable_autocomplete, 
            user_list=users_list
        )
        self.composite.pack(side='left', padx=(0, 2))
        self.cb_name = self.composite.cb_name
        
        # 3. Work Time (Changed to Combobox for mouse selection)
        ttk.Label(self, text="시간:").pack(side='left', padx=(5, 0))
        self.ent_worktime = ttk.Combobox(self, width=16, values=time_list or [])
        self.ent_worktime.pack(side='left', padx=(0, 2))
        self.ent_worktime.set("") # Default empty

    def get_worker(self): return self.composite.get()
    def set_worker(self, val): self.composite.set(val)
    
    def get_time(self): 
        """Return time string"""
        return self.ent_worktime.get().strip()
        
    def set_time(self, val):
        """Set time widget"""
        if not val:
            self.ent_worktime.set("")
        else:
            self.ent_worktime.set(str(val).strip())

    def bind_name(self, seq, func): self.cb_name.bind(seq, func)
    def bind_time(self, seq, func): 
        self.ent_worktime.bind(seq, func)
        if 'FocusOut' in seq or 'Return' in seq:
            self.ent_worktime.bind('<<ComboboxSelected>>', func, add='+')

    def update_time_list(self, new_list):
        """Refresh the combobox values with a new list"""
        if hasattr(self, 'ent_worktime'):
            self.ent_worktime['values'] = new_list

class NDTCalculator(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("비파괴검사 기성 산출 계산기 (가산~가평)")
        self.geometry("1150x800")  
        self.configure(padx=10, pady=10)
        self.style = ttk.Style(self)
        self.style.theme_use("clam")
        
        self.records = [] # 저장된 기록 목록
        self.worker_records = [] # 3번 탭 전용 독립 작업자 기록 (Date, Name, Shift, WorkTime, OT)
        
        self.users = ["", "부장 주진철", "대리 우명광", "주임 김진환", "계장 장승대", "주임 김성렬", "부장 박광복", "과장 주영광"]
        def format_time(i):
            return f"{i:02d}:00" if i <= 24 else f"익일{i-24:02d}:00"
            
        self.time_list = [""] + [f"09:00~{format_time(i)}" for i in range(18, 34)]
        
        # Load any custom times previously saved
        custom_times = CONFIG.get("CUSTOM_TIMES", [])
        for ct in custom_times:
            if ct not in self.time_list:
                self.time_list.append(ct)
        
        self.create_menu()
        self.create_widgets()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        
    def on_closing(self):
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
        self.destroy()
        
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
        
        tab_worker = ttk.Frame(self.notebook)
        self.notebook.add(tab_worker, text="3. 작업자별 실적 요약")
        
        # --- TAB 1: WORK (입력 폼 및 목록 사이드바이사이드) ---
        self.work_pane = tk.PanedWindow(tab_work, orient=tk.HORIZONTAL, sashrelief=tk.RAISED, sashwidth=5, bg="#b0b0b0")
        self.work_pane.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        left_frame = ttk.Frame(self.work_pane)
        self.work_pane.add(left_frame, stretch="always")
        
        info_frame = ttk.Frame(left_frame)
        info_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(info_frame, text="• 검사일자:", font=("Arial", 10, "bold")).pack(side=tk.LEFT)
        self.date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        self.date_entry = DateEntry(info_frame, textvariable=self.date_var, width=13, date_pattern='yyyy-mm-dd', background='darkblue', foreground='white', borderwidth=2)
        self.date_entry.pack(side=tk.LEFT, padx=5)

        ttk.Label(left_frame, text="1. 검사 종류", font=("Arial", 11, "bold")).pack(anchor=tk.W, pady=(10, 5))
        self.ndt_type_var = tk.StringVar(value="RT")
        type_frame = ttk.Frame(left_frame)
        type_frame.pack(fill=tk.X, pady=5)
        for t in ["RT", "UT", "PT"]:
            ttk.Radiobutton(type_frame, text=t, value=t, variable=self.ndt_type_var, command=self.update_dynamic_ui).pack(side=tk.LEFT, padx=10)
            
        ttk.Label(left_frame, text="2. 작업 구분", font=("Arial", 11, "bold")).pack(anchor=tk.W, pady=(10, 5))
        
        self.loc_type_var = tk.StringVar(value="수송배관(주배관)")
        self.work_time_var = tk.StringVar(value="일반")
        
        type_frame = ttk.Frame(left_frame)
        type_frame.pack(fill=tk.X, pady=(5, 2))
        
        ttk.Label(type_frame, text="구간:").pack(side=tk.LEFT)
        for t in ["수송배관(주배관)", "플랜트(관리소)"]:
            ttk.Radiobutton(type_frame, text=t, value=t, variable=self.loc_type_var).pack(side=tk.LEFT, padx=5)
            
        time_frame = ttk.Frame(left_frame)
        time_frame.pack(fill=tk.X, pady=(2, 5))
        
        ttk.Label(time_frame, text="시간:").pack(side=tk.LEFT)
        for t in ["일반", "야간", "휴일"]:
            ttk.Radiobutton(time_frame, text=t, value=t, variable=self.work_time_var).pack(side=tk.LEFT, padx=5)

        self.material_lbl = ttk.Label(left_frame, text="3. 사용 자재", font=("Arial", 11, "bold"))
        self.material_lbl.pack(anchor=tk.W, pady=(10, 5))
        self.material_var = tk.StringVar(value='RT (B필름: 3⅓"x17")')
        self.material_combo = ttk.Combobox(left_frame, textvariable=self.material_var, values=['RT (B필름: 3⅓"x17")', 'RT (A필름: 3⅓"x12")', 'RT (A/2필름: 3⅓"x6")'], state="readonly", width=25)
        self.material_combo.pack(fill=tk.X, pady=5)
        
        self.dynamic_frame = ttk.LabelFrame(left_frame, text="4. 보정계수 조건 선택", padding=10)
        self.dynamic_frame.pack(fill=tk.X, pady=(10, 5))
        
        self.source_frame = ttk.Frame(self.dynamic_frame)
        ttk.Label(self.source_frame, text="• 방사선원 :", width=15).pack(side=tk.LEFT)
        self.source_var = tk.StringVar(value="Ir-192 또는 Se-75 (1.0)")
        self.source_combo = ttk.Combobox(self.source_frame, textvariable=self.source_var, state="readonly", width=20)
        self.source_combo['values'] = ["Ir-192 또는 Se-75 (1.0)", "X-ray 발생장치 (1.3)"]
        self.source_combo.pack(side=tk.LEFT)
        
        self.pipe_frame = ttk.Frame(self.dynamic_frame)
        ttk.Label(self.pipe_frame, text="• 관경(구경) :", width=15).pack(side=tk.LEFT)
        self.pipe_var = tk.StringVar()
        self.pipe_combo = ttk.Combobox(self.pipe_frame, textvariable=self.pipe_var, state="readonly", width=20)
        self.pipe_combo.pack(side=tk.LEFT)
        
        self.thickness_frame = ttk.Frame(self.dynamic_frame)
        ttk.Label(self.thickness_frame, text="• 투과/모재두께 :", width=15).pack(side=tk.LEFT)
        self.thickness_var = tk.StringVar()
        self.thickness_combo = ttk.Combobox(self.thickness_frame, textvariable=self.thickness_var, state="readonly", width=20)
        self.thickness_combo.pack(side=tk.LEFT)
        
        ttk.Label(left_frame, text="5. 실검사 물량 (매/m)", font=("Arial", 11, "bold")).pack(anchor=tk.W, pady=(10, 5))
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
        self.result_text = tk.Text(left_frame, height=5, width=30, state=tk.DISABLED, font=("Consolas", 11))
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
        
        content_frame = ttk.Frame(billing_container)
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        contract_frame = ttk.LabelFrame(content_frame, text="항목별 계약 및 전회 기성 (세액 미포함)", padding=10)
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
        
        items = [
            ("RT_B", 'RT (B필름)', "20,368"),
            ("RT_A", 'RT (A필름)', "2,464"),
            ("RT_A2", 'RT (A/2필름)', "1,704"),
            ("UT", "UT", "319.02"),
            ("PT", "PT", "338.63")
        ]
        
        for key, display_name, default_val in items:
            f = ttk.Frame(contract_frame)
            f.pack(fill=tk.X, pady=2)
            
            lbl_f = ttk.Frame(f)
            lbl_f.pack(side=tk.LEFT, padx=(0, 5))
            ttk.Label(lbl_f, text=f"[{display_name}]", width=12, font=("Arial", 9, "bold")).pack(anchor=tk.W)
            
            grid_f = ttk.Frame(f)
            grid_f.pack(side=tk.LEFT, fill=tk.X, expand=True)
            
            c_qty = tk.StringVar(value=default_val)
            p_qty = tk.StringVar(value="0")
            curr_qty = tk.StringVar(value="0")
            rem_qty = tk.StringVar(value=default_val)
            
            # backward compatibility
            c_price = tk.IntVar(value=0)
            c_var = tk.StringVar(value="0")
            p_var = tk.StringVar(value="0")
            
            c_qty.trace_add("write", lambda *a, v=c_qty: format_qty(var=v))
            p_qty.trace_add("write", lambda *a, v=p_qty: format_qty(var=v))
            
            def update_rem_qty(*args, k=key):
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
            
            unit = "매" if key.startswith("RT") else "M"
            
            ttk.Label(grid_f, text="계약:").grid(row=0, column=0, sticky=tk.E, padx=2)
            ttk.Entry(grid_f, textvariable=c_qty, width=8).grid(row=0, column=1, padx=2)
            ttk.Label(grid_f, text=unit).grid(row=0, column=2, sticky=tk.W, padx=(0, 10))
            
            ttk.Label(grid_f, text="전회:").grid(row=0, column=3, sticky=tk.E, padx=2)
            ttk.Entry(grid_f, textvariable=p_qty, width=8).grid(row=0, column=4, padx=2)
            ttk.Label(grid_f, text=unit).grid(row=0, column=5, sticky=tk.W, padx=(0, 10))
            
            ttk.Label(grid_f, text="금회:").grid(row=0, column=6, sticky=tk.E, padx=2)
            ttk.Label(grid_f, textvariable=curr_qty, width=6, anchor="e", foreground="green").grid(row=0, column=7, padx=2)
            ttk.Label(grid_f, text=unit).grid(row=0, column=8, sticky=tk.W, padx=(0, 10))
            
            ttk.Label(grid_f, text="잔여:").grid(row=0, column=9, sticky=tk.E, padx=2)
            lbl_rem = ttk.Label(grid_f, textvariable=rem_qty, width=8, anchor="e", font=("Arial", 9, "bold"))
            lbl_rem.grid(row=0, column=10, padx=2)
            ttk.Label(grid_f, text=unit).grid(row=0, column=11, sticky=tk.W)
            
            c_var.trace_add("write", lambda *a, v=c_var: format_currency(var=v))
            p_var.trace_add("write", lambda *a, v=p_var: format_currency(var=v))
            
            ttk.Label(grid_f, text="계약금액:").grid(row=1, column=0, sticky=tk.E, padx=2, pady=(2, 0))
            ttk.Entry(grid_f, textvariable=c_var, width=12).grid(row=1, column=1, columnspan=2, sticky=tk.W, padx=2, pady=(2, 0))
            
            ttk.Label(grid_f, text="전회금액:").grid(row=1, column=3, sticky=tk.E, padx=2, pady=(2, 0))
            ttk.Entry(grid_f, textvariable=p_var, width=12).grid(row=1, column=4, columnspan=2, sticky=tk.W, padx=2, pady=(2, 0))
            
            ttk.Separator(contract_frame, orient='horizontal').pack(fill=tk.X, pady=3)
            
            self.contract_vars[key] = {
                "c_qty": c_qty, "p_qty": p_qty, "curr_qty": curr_qty, "rem_qty": rem_qty, "lbl_rem": lbl_rem,
                "c_price": c_price, "contract": c_var, "prev": p_var
            }

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
        
        exp_frame = ttk.LabelFrame(content_frame, text="기타 경비 및 실비 정산 (월간)", padding=10)
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
        bottom_frame = ttk.Frame(self.work_pane)
        self.work_pane.add(bottom_frame, stretch="always")
        
        lbl_frame = ttk.Frame(bottom_frame)
        lbl_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(lbl_frame, text="[ 일일 작업 기록 목록 ]", font=("Arial", 11, "bold")).pack(side=tk.LEFT)
        ttk.Button(lbl_frame, text="기성청구", command=self.export_to_excel).pack(side=tk.RIGHT)
        ttk.Button(lbl_frame, text="기록 초기화", command=self.clear_records).pack(side=tk.RIGHT, padx=5)
        ttk.Button(lbl_frame, text="선택 삭제", command=self.delete_selected_records).pack(side=tk.RIGHT)

        columns = ("date", "loc", "type", "time", "mat", "qty", "unit", "corr", "adj_qty", "mat_cost", "lab_cost", "overhead", "tech", "total_amt")
        self.tree = ttk.Treeview(bottom_frame, columns=columns, show="headings", height=8)
        
        self.tree.heading("date", text="일자", anchor="center")
        self.tree.heading("loc", text="구간", anchor="center")
        self.tree.heading("type", text="종류", anchor="center")
        self.tree.heading("time", text="형태", anchor="center")
        self.tree.heading("mat", text="자재", anchor="center")
        self.tree.heading("qty", text="실물량", anchor="center")
        self.tree.heading("unit", text="단위", anchor="center")
        self.tree.heading("corr", text="보정계수", anchor="center")
        self.tree.heading("adj_qty", text="보정물량", anchor="center")
        self.tree.heading("mat_cost", text="재료비(원)", anchor="center")
        self.tree.heading("lab_cost", text="인건비(원)", anchor="center")
        self.tree.heading("overhead", text="제경비(원)", anchor="center")
        self.tree.heading("tech", text="기술료(원)", anchor="center")
        self.tree.heading("total_amt", text="공급가액(원)", anchor="center")
        
        default_widths = {
            "date": 80, "loc": 100, "type": 40, "time": 40, "mat": 90, 
            "qty": 40, "unit": 40, "corr": 50, "adj_qty": 50,
            "mat_cost": 70, "lab_cost": 70, "overhead": 60, "tech": 60, "total_amt": 80
        }
        saved_widths = CONFIG.get("TREE_WIDTHS", {})
        
        for col in columns:
            w = saved_widths.get(col, default_widths.get(col, 80))
            is_last = (col == columns[-1])
            self.tree.column(col, width=w, stretch=tk.YES if is_last else tk.NO, anchor="center" if col != "loc" else "w")
        
        tree_scroll = ttk.Scrollbar(bottom_frame, orient="vertical", command=self.tree.yview)
        tree_xscroll = ttk.Scrollbar(bottom_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=tree_scroll.set, xscrollcommand=tree_xscroll.set)
        
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        tree_xscroll.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)
        self.tree.bind('<Delete>', self.delete_selected_records)
        self.tree.bind('<BackSpace>', self.delete_selected_records)
        
        # --- TAB 3: WORKER SUMMARY ---
        self.create_worker_summary_tab(tab_worker)
        
        # 저장된 탭 영역(Sash) 너비 복원
        def restore_sash(event=None):
            if getattr(self, "_sash_restored", False):
                return
            if "SASH_POS" not in CONFIG:
                return
            try:
                sash_pos = int(CONFIG["SASH_POS"])
                if sash_pos > 280:
                    sash_pos = 280
                self.work_pane.sash_place(0, sash_pos, 0)
                self._sash_restored = True
            except:
                pass
                
        # tk.PanedWindow는 Configure 시 사이즈가 확정된 후 복원
        self.work_pane.bind("<Configure>", restore_sash)
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
        
        display_loc = f"[{loc_type_val}]"
        
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
        
        if "[플랜트(관리소)]" in res['loc']:
            loc_type_val = "플랜트(관리소)"
        else:
            loc_type_val = "수송배관(주배관)"
            
        if loc_type_val in LABOR_COST:
            lab_unit = LABOR_COST[loc_type_val].get(res['work_time'], {}).get(res['ndt_type'], 0)
        else:
            lab_unit = LABOR_COST.get(res['work_time'], {}).get(res['ndt_type'], 0)
        
        txt = (f"▶ [선택된 기록] 일자: {res['date']} | 구간: {res['loc']}\n"
               f"▶ [적용 기준] 보정계수: {res['corr']:.2f} | 재료비 단가: {mat_unit:,}원 | 인건비 단가: {lab_unit:,}원\n"
               f"▶ [공급 가액] {res['subtotal']:,} 원 (재료비 {res['mat_cost']:,} + 인건비 {res['lab_cost']:,} + 제경비 {res['overhead']:,} + 기술료 {res['tech']:,})\n"
               f"▶ [최종 금액] 총 청구액 {res['total_amount']:,} 원 (부가세 {res['vat']:,}원 포함)\n")
        
        self.result_text.config(state=tk.NORMAL)
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, txt)
        self.result_text.config(state=tk.DISABLED)

    def update_qty_summary(self):
        totals = {"RT_B": 0.0, "RT_A": 0.0, "RT_A2": 0.0, "UT": 0.0, "PT": 0.0}
        for rec in self.records:
            if rec["ndt_type"] == "RT":
                if "17" in rec["material_type"]:
                    totals["RT_B"] += rec["qty"]
                elif "12" in rec["material_type"]:
                    totals["RT_A"] += rec["qty"]
                elif "6" in rec["material_type"]:
                    totals["RT_A2"] += rec["qty"]
            elif rec["ndt_type"] in totals:
                totals[rec["ndt_type"]] += rec["qty"]
                
        for k, v in totals.items():
            if k in self.contract_vars:
                formatted = f"{int(v):,}" if v.is_integer() else f"{v:,.2f}"
                self.contract_vars[k]["curr_qty"].set(formatted)

    def add_to_record(self):
        res = self.calculate()
        if res:
            self.records.append(res)
            self.tree.insert("", tk.END, values=(
                res["date"], res["loc"], res["ndt_type"], res["work_time"], 
                res["material_type"], f"{res['qty']:.1f}", res["unit"],
                f"{res['corr']:.2f}", f"{res['adjusted_qty']:.2f}", 
                f"{res.get('mat_cost', 0):,}", f"{res.get('lab_cost', 0):,}",
                f"{res['overhead']:,}", f"{res['tech']:,}", f"{res['subtotal']:,}"
            ))
            self.update_qty_summary()
            
    def clear_records(self):
        self.records = []
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.update_qty_summary()
        
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
            if cat.startswith("RT"):
                if cat == "RT_B":
                    c_qty = sum(r["qty"] for r in target_records if r["ndt_type"] == "RT" and "17" in r["material_type"])
                    cur_amt = sum(r["subtotal"] for r in target_records if r["ndt_type"] == "RT" and "17" in r["material_type"])
                elif cat == "RT_A":
                    c_qty = sum(r["qty"] for r in target_records if r["ndt_type"] == "RT" and "12" in r["material_type"])
                    cur_amt = sum(r["subtotal"] for r in target_records if r["ndt_type"] == "RT" and "12" in r["material_type"])
                else:
                    c_qty = sum(r["qty"] for r in target_records if r["ndt_type"] == "RT" and "6" in r["material_type"])
                    cur_amt = sum(r["subtotal"] for r in target_records if r["ndt_type"] == "RT" and "6" in r["material_type"])
            else:
                c_qty = sum(r["qty"] for r in target_records if r["ndt_type"] == cat)
                cur_amt = sum(r["subtotal"] for r in target_records if r["ndt_type"] == cat)
                
            new_p_qty = p_qty + c_qty
            formatted_qty = f"{int(new_p_qty):,}" if new_p_qty.is_integer() else f"{new_p_qty:,.2f}"
            v["p_qty"].set(formatted_qty)
            v["prev"].set(f"{p_amt + cur_amt:,}")
            
        exp_curr = sum([self.equip_cost_var.get(), self.safety_cost_var.get(), self.travel_cost_var.get(), self.print_cost_var.get()])
        exp_prev = self.get_int(self.exp_prev_var)
        self.exp_prev_var.set(f"{exp_prev + exp_curr:,}")
        
        self.equip_cost_var.set(0)
        self.safety_cost_var.set(0)
        self.travel_cost_var.set(0)
        self.print_cost_var.set(0)
        
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
                "worker_records": getattr(self, 'worker_records', []),
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
        
        rt_b = pipe.get("RT_B", 0) + plant.get("RT_B", 0)
        rt_a = pipe.get("RT_A", 0) + plant.get("RT_A", 0)
        rt_a2 = pipe.get("RT_A2", 0) + plant.get("RT_A2", 0)
        ut_total = pipe.get("UT", 0) + plant.get("UT", 0)
        pt_total = pipe.get("PT", 0) + plant.get("PT", 0)
        
        self.contract_vars["RT_B"]["c_qty"].set(f"{int(rt_b):,}")
        self.contract_vars["RT_A"]["c_qty"].set(f"{int(rt_a):,}")
        self.contract_vars["RT_A2"]["c_qty"].set(f"{int(rt_a2):,}")
        self.contract_vars["UT"]["c_qty"].set(f"{ut_total:,.2f}")
        self.contract_vars["PT"]["c_qty"].set(f"{pt_total:,.2f}")
        
        messagebox.showinfo("불러오기 완료", "프로젝트 전체 총 계약 물량이 자동으로 입력되었습니다.")

    def load_project(self):
        try:
            filepaths = filedialog.askopenfilenames(filetypes=[("NDT Project", "*.ndt")], title="작업 불러오기 (여러 파일 선택 시 병합됨)")
            if not filepaths: return
            
            self.clear_records()
            self.records = []
            if hasattr(self, 'worker_records'):
                self.worker_records.clear()
            else:
                self.worker_records = []
            
            # 다중 파일 선택 시 순서 보장을 위해 정렬
            filepaths = sorted(filepaths)
            
            latest_data = None
            max_round = -1
            
            for filepath in filepaths:
                with open(filepath, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    
                current_records = data.get("records", [])
                self.records.extend(current_records)
                
                current_worker_records = data.get("worker_records", [])
                if hasattr(self, 'worker_records'):
                    self.worker_records.extend(current_worker_records)
                
                for res in current_records:
                    self.tree.insert("", tk.END, values=(
                        res["date"], res.get("worker", ""), res["loc"], res["ndt_type"], res["work_time"], 
                        res["material_type"], f"{res['qty']:.1f}", res["unit"],
                        f"{res['corr']:.2f}", f"{res['adjusted_qty']:.2f}", 
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
            self.equip_cost_var.set(ex.get("equip", 0))
            self.safety_cost_var.set(ex.get("safety", 0))
            self.travel_cost_var.set(ex.get("travel", 0))
            self.print_cost_var.set(ex.get("print", 0))
            self.exp_budget_var.set(f"{ex.get('budget', 72215000):,}")
            self.exp_prev_var.set(f"{ex.get('prev', 0):,}")
            
            tot = latest_data.get("total_amt", {})
            self.total_contract_var.set(f"{tot.get('contract', 2628702818):,}")
            self.total_prev_var.set(f"{tot.get('prev', 0):,}")
            
            if hasattr(self, 'update_worker_summary'):
                self.update_worker_summary()
            
            msg = f"총 {len(filepaths)}개의 작업 파일에서 {len(self.records)}개의 기성 기록과 {len(getattr(self, 'worker_records', []))}개의 작업자 기록을 성공적으로 병합하여 불러왔습니다."
            messagebox.showinfo("불러오기 완료", msg)
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
            messagebox.showwarning("기록 없음", "출력할 작업 기록이 없습니다.")
            return
            
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
            ws.Range("C4").Value = "가산~가평 천연가스 공급시설 건설공사 비파괴검사 기술용역"
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
            
            ws.Range(ws.Cells(6, 1), ws.Cells(14, 11)).Borders.LineStyle = 1
            
            extra_items_total = sum([self.equip_cost_var.get(), self.safety_cost_var.get(), self.travel_cost_var.get(), self.print_cost_var.get()])
            
            categories = ["RT_B", "RT_A", "RT_A2", "UT", "PT", "기타실비", "총 계"]
            display_names = {
                "RT_B": "RT (B필름)", "RT_A": "RT (A필름)", "RT_A2": "RT (A/2필름)",
                "UT": "UT", "PT": "PT", "기타실비": "기타실비", "총 계": "총 계"
            }
            
            for i, cat in enumerate(categories):
                row = 8 + i
                ws.Cells(row, 1).Value = display_names[cat]
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
                    
                    c_amt_val = self.get_int(self.contract_vars[cat]["contract"])
                    p_amt_val = self.get_int(self.contract_vars[cat]["prev"])
                    
                    if cat.startswith("RT"):
                        if cat == "RT_B":
                            cur_qty = sum(r["adjusted_qty"] for r in target_records if r["ndt_type"] == "RT" and "17" in r["material_type"])
                            cur_amt = sum(r["subtotal"] for r in target_records if r["ndt_type"] == "RT" and "17" in r["material_type"])
                        elif cat == "RT_A":
                            cur_qty = sum(r["adjusted_qty"] for r in target_records if r["ndt_type"] == "RT" and "12" in r["material_type"])
                            cur_amt = sum(r["subtotal"] for r in target_records if r["ndt_type"] == "RT" and "12" in r["material_type"])
                        else:
                            cur_qty = sum(r["adjusted_qty"] for r in target_records if r["ndt_type"] == "RT" and "6" in r["material_type"])
                            cur_amt = sum(r["subtotal"] for r in target_records if r["ndt_type"] == "RT" and "6" in r["material_type"])
                    else:
                        cur_qty = sum(r["adjusted_qty"] for r in target_records if r["ndt_type"] == cat)
                        cur_amt = sum(r["subtotal"] for r in target_records if r["ndt_type"] == cat)
                        
                    tot_qty = p_qty + cur_qty
                    rem_qty = c_qty - tot_qty
                    
                    tot_amt = p_amt_val + cur_amt
                    rem_amt = c_amt_val - tot_amt
                    c_amt = c_amt_val
                    p_amt = p_amt_val
                
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
            
            start_row = 17
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
                    ws.Cells(current_row, 3).Value = key[0]
                    ws.Cells(current_row, 4).Value = key[1]
                    ws.Cells(current_row, 5).Value = key[2]
                    ws.Cells(current_row, 6).Value = key[3]
                    ws.Cells(current_row, 7).Value = data["qty"]
                    ws.Cells(current_row, 8).Value = key[4]
                    ws.Cells(current_row, 9).Value = key[5]
                    ws.Cells(current_row, 10).Value = data["adjusted_qty"]
                    ws.Cells(current_row, 11).Value = data["mat_cost"]
                    ws.Cells(current_row, 12).Value = data["lab_cost"]
                    ws.Cells(current_row, 13).Value = data["overhead"]
                    ws.Cells(current_row, 14).Value = data["tech"]
                    ws.Cells(current_row, 15).Value = data["subtotal"]
                    
                    for c in range(1, 16):
                        cell = ws.Cells(current_row, c)
                        cell.Borders.LineStyle = 1
                        if c <= 4 or c == 6 or c == 8: cell.HorizontalAlignment = -4108
                        elif c == 5 or c == 3: cell.HorizontalAlignment = -4131
                        else: cell.NumberFormat = "#,##0" if c >= 11 else "0.00"
                    
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

    def create_worker_summary_tab(self, parent_frame):
        # Use PanedWindow for side-by-side
        paned = ttk.PanedWindow(parent_frame, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Left Panel: Input Form
        left_frame = ttk.LabelFrame(paned, text="작업자 일일 기록 입력", padding=10)
        paned.add(left_frame, weight=1)
        
        # Date selection
        date_frame = ttk.Frame(left_frame)
        date_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(date_frame, text="작업일자:").pack(side=tk.LEFT, padx=(0, 5))
        self.worker_input_date = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        DateEntry(date_frame, textvariable=self.worker_input_date, width=12, date_pattern='yyyy-mm-dd', background='darkblue', foreground='white', borderwidth=2).pack(side=tk.LEFT)
        
        ttk.Button(date_frame, text="첫 줄 기준 일괄적용", command=self.apply_first_row_to_all).pack(side=tk.RIGHT, padx=5)
        
        # 10 WorkerDataGroup rows
        self.worker_entries = []
        for i in range(10):
            wdg = WorkerDataGroup(left_frame, i, self.users, self.time_list, enable_autocomplete=True)
            wdg.pack(fill=tk.X, pady=2)
            
            # Bind custom time entry
            wdg.ent_worktime.bind('<FocusOut>', lambda e, w=wdg: self.check_and_save_time(w.ent_worktime.get()))
            wdg.ent_worktime.bind('<Return>', lambda e, w=wdg: self.check_and_save_time(w.ent_worktime.get()))
            
            self.worker_entries.append(wdg)
            
        # Buttons
        btn_frame = ttk.Frame(left_frame)
        btn_frame.pack(fill=tk.X, pady=10)
        ttk.Button(btn_frame, text="기록 저장", command=self.save_worker_records).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="입력폼 초기화", command=self.clear_worker_form).pack(side=tk.LEFT, padx=5)
        
        # Right Panel: Summary
        right_frame = ttk.Frame(paned)
        paned.add(right_frame, weight=3)
        
        # Filter frame
        filter_frame = ttk.LabelFrame(right_frame, text="기간 조회 및 요약", padding=10)
        filter_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(filter_frame, text="시작일:").pack(side=tk.LEFT, padx=5)
        self.worker_start_date = tk.StringVar(value=(datetime.now() - timedelta(days=30)).strftime('%Y-%m-%d'))
        DateEntry(filter_frame, textvariable=self.worker_start_date, width=12, date_pattern='yyyy-mm-dd', background='darkblue', foreground='white', borderwidth=2).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(filter_frame, text="종료일:").pack(side=tk.LEFT, padx=5)
        self.worker_end_date = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        DateEntry(filter_frame, textvariable=self.worker_end_date, width=12, date_pattern='yyyy-mm-dd', background='darkblue', foreground='white', borderwidth=2).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(filter_frame, text="조회하기", command=self.update_worker_summary).pack(side=tk.LEFT, padx=20)
        ttk.Button(filter_frame, text="전체 삭제", command=self.clear_all_worker_records).pack(side=tk.RIGHT, padx=5)
        
        # Treeview
        tree_frame = ttk.Frame(right_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        columns = (
            "worker", "days", "ot_day", "ot_night", "ot_holiday", "total_ot",
            "amt_day", "amt_night", "amt_holiday", "total_amt"
        )
        self.worker_tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        
        self.worker_tree.heading("worker", text="작업자")
        self.worker_tree.heading("days", text="출근일수")
        self.worker_tree.heading("ot_day", text="연장(시간)")
        self.worker_tree.heading("ot_night", text="야간(시간)")
        self.worker_tree.heading("ot_holiday", text="휴일(시간)")
        self.worker_tree.heading("total_ot", text="총OT(시간)")
        self.worker_tree.heading("amt_day", text="연장(금액)")
        self.worker_tree.heading("amt_night", text="야간(금액)")
        self.worker_tree.heading("amt_holiday", text="휴일(금액)")
        self.worker_tree.heading("total_amt", text="총OT(금액)")
        
        for col in columns:
            is_last = (col == columns[-1])
            self.worker_tree.column(col, width=80, stretch=tk.YES if is_last else tk.NO, anchor="center" if col in ("worker", "days", "ot_day", "ot_night", "ot_holiday", "total_ot") else "e")
            
        scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.worker_tree.yview)
        scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.worker_tree.xview)
        self.worker_tree.configure(yscrollcommand=scroll.set, xscrollcommand=scroll_x.set)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.worker_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.worker_tree.bind('<Double-1>', self.on_worker_double_click)

    def check_and_save_time(self, val):
        val = str(val).strip()
        if not val or val in self.time_list:
            return
            
        self.time_list.append(val)
        
        # Save to config.json
        global CONFIG
        custom_times = CONFIG.get("CUSTOM_TIMES", [])
        if val not in custom_times:
            custom_times.append(val)
            CONFIG["CUSTOM_TIMES"] = custom_times
            save_config(CONFIG)
            
        # Update all active comboboxes
        for wdg in getattr(self, 'worker_entries', []):
            wdg.update_time_list(self.time_list)

    def apply_first_row_to_all(self):
        if not self.worker_entries: return
        
        first_row = self.worker_entries[0]
        time_val = first_row.ent_worktime.get()
        
        applied_count = 0
        for wdg in self.worker_entries[1:]:
            if wdg.get_worker():
                wdg.ent_worktime.set(time_val)
                applied_count += 1
                
        if applied_count > 0:
            messagebox.showinfo("일괄적용", f"첫 번째 줄의 시간 정보가 {applied_count}명의 작업자에게 일괄 적용되었습니다.")
        else:
            messagebox.showinfo("알림", "일괄 적용할 다른 작업자(이름)가 선택되지 않았습니다.\n두 번째 줄부터 이름을 먼저 선택해주세요.")

    def save_worker_records(self):
        input_date = self.worker_input_date.get()
        added = 0
        for wdg in self.worker_entries:
            w = wdg.get_worker()
            if not w: continue
            
            time_val = wdg.ent_worktime.get()
            
            self.worker_records.append({
                "date": input_date,
                "worker": w,
                "work_time": time_val
            })
            added += 1
            
        if added > 0:
            messagebox.showinfo("저장 완료", f"{added}명의 작업자 기록이 추가되었습니다.")
            self.clear_worker_form()
            self.update_worker_summary()
        else:
            messagebox.showwarning("입력 없음", "저장할 작업자가 입력되지 않았습니다.")

    def clear_worker_form(self):
        for wdg in self.worker_entries:
            wdg.set_worker("")
            wdg.ent_worktime.set("")

    def clear_all_worker_records(self):
        if messagebox.askyesno("전체 삭제", "저장된 모든 작업자 기록을 삭제하시겠습니까?\n(1번 탭의 기성 검사 기록은 유지됩니다)"):
            self.worker_records.clear()
            self.update_worker_summary()

    def _calculate_ot_from_worktime(self, worktime_value, date_str):
        try:
            if not worktime_value or "~" not in worktime_value:
                return 0.0, 0.0, 0.0
                
            d_obj = datetime.strptime(date_str, '%Y-%m-%d')
            weekday = d_obj.weekday()
            is_holiday = weekday >= 5
            is_friday = (weekday == 4)
            
            clean_val = str(worktime_value).strip()
            start_str, end_str = clean_val.split("~")
            
            sh, sm = map(int, start_str.strip().split(':'))
            start_f = sh + sm / 60.0
            
            end_str = end_str.strip()
            if "익일" in end_str:
                eh, em = map(int, end_str.replace("익일", "").split(':'))
                end_f = eh + 24 + em / 60.0
            else:
                eh, em = map(int, end_str.split(':'))
                end_f = eh + em / 60.0
                if end_f < start_f: end_f += 24
                
            total_duration = end_f - start_f
            if total_duration <= 0:
                return 0.0, 0.0, 0.0
                
            if is_holiday:
                return 0.0, 0.0, total_duration
            else:
                ot_day = 0.0
                ot_night = 0.0
                ot_holiday = 0.0
                
                if end_f > 18:
                    ot_start = max(start_f, 18.0)
                    
                    # 저녁 연장 (18:00 ~ 22:00)
                    evening_end = min(end_f, 22.0)
                    ot_day += max(0, evening_end - ot_start)
                    
                    # 야간 연장 (22:00 ~ 24:00)
                    night_start = max(ot_start, 22.0)
                    night_end = min(end_f, 24.0)
                    ot_night += max(0, night_end - night_start)
                    
                    # 심야/새벽 (24:00 ~ )
                    dawn_start = max(ot_start, 24.0)
                    dawn_hours = max(0, end_f - dawn_start)
                    
                    if is_friday:
                        ot_holiday += dawn_hours
                    else:
                        ot_night += dawn_hours
                        
                return ot_day, ot_night, ot_holiday
        except:
            return 0.0, 0.0, 0.0

    def update_worker_summary(self):
        for item in self.worker_tree.get_children():
            self.worker_tree.delete(item)
            
        start_str = self.worker_start_date.get()
        end_str = self.worker_end_date.get()
        
        summary = {}
        for rec in self.worker_records:
            r_date = rec.get("date", "")
            if start_str <= r_date <= end_str:
                w = rec.get("worker", "").strip()
                if not w: w = "미지정"
                
                if w not in summary:
                    summary[w] = {
                        "days": set(),
                        "ot_day": 0.0, "ot_night": 0.0, "ot_holiday": 0.0
                    }
                
                summary[w]["days"].add(r_date)
                
                # 1. 자동 OT 계산 로직 (Material Master Manager 기준)
                time_val = rec.get("work_time", "")
                ot_d, ot_n, ot_h = self._calculate_ot_from_worktime(time_val, r_date)
                
                # 2. 과거 수동 기록 호환성 보장 (시간 문자열이 없거나 파싱 안될 때)
                if ot_d == 0 and ot_n == 0 and ot_h == 0:
                    shift = rec.get("shift", "")
                    try: manual_ot = float(rec.get("ot", 0) or 0)
                    except: manual_ot = 0.0
                    
                    if shift == "야간":
                        ot_n = manual_ot
                    elif shift == "휴일" or shift == "주야간":
                        ot_h = manual_ot
                    else:
                        ot_d = manual_ot
                        
                summary[w]["ot_day"] += ot_d
                summary[w]["ot_night"] += ot_n
                summary[w]["ot_holiday"] += ot_h
                    
        # Apply Rates (Day: 4000, Night: 5000, Holiday: 7500)
        tot_days = tot_ot_day = tot_ot_night = tot_ot_holiday = tot_ot = 0
        tot_amt_day = tot_amt_night = tot_amt_holiday = tot_amt = 0
        
        self._last_worker_summary = summary
        
        for w in sorted(summary.keys()):
            s = summary[w]
            days = len(s["days"])
            ot_d = s["ot_day"]
            ot_n = s["ot_night"]
            ot_h = s["ot_holiday"]
            tot = ot_d + ot_n + ot_h
            
            amt_d = ot_d * 4000
            amt_n = ot_n * 5000
            amt_h = ot_h * 7500
            total_a = amt_d + amt_n + amt_h
            
            tot_days += days
            tot_ot_day += ot_d; tot_ot_night += ot_n; tot_ot_holiday += ot_h; tot_ot += tot
            tot_amt_day += amt_d; tot_amt_night += amt_n; tot_amt_holiday += amt_h; tot_amt += total_a
            
            self.worker_tree.insert("", tk.END, values=(
                w, f"{days}일", f"{ot_d:.1f}", f"{ot_n:.1f}", f"{ot_h:.1f}", f"{tot:.1f}",
                f"{int(amt_d):,}", f"{int(amt_n):,}", f"{int(amt_h):,}", f"{int(total_a):,}"
            ))
            
        if summary:
            self.worker_tree.insert("", tk.END, values=(
                "[총계]", f"{tot_days}일", f"{tot_ot_day:.1f}", f"{tot_ot_night:.1f}", f"{tot_ot_holiday:.1f}", f"{tot_ot:.1f}",
                f"{int(tot_amt_day):,}", f"{int(tot_amt_night):,}", f"{int(tot_amt_holiday):,}", f"{int(tot_amt):,}"
            ), tags=('total',))
            self.worker_tree.tag_configure('total', background='#e6f2ff', font=('Arial', 10, 'bold'))

    def on_worker_double_click(self, event):
        item = self.worker_tree.focus()
        if not item:
            return
        
        values = self.worker_tree.item(item, 'values')
        if not values:
            return
            
        worker_name = values[0]
        if worker_name == "[총계]":
            return
            
        summary = getattr(self, '_last_worker_summary', {})
        if worker_name in summary:
            days = sorted(list(summary[worker_name]["days"]))
            days_str = "\n".join(days)
            messagebox.showinfo(f"{worker_name} 출근일 목록", f"총 {len(days)}일 출근하셨습니다.\n\n[상세 출근일]\n{days_str}")
if __name__ == "__main__":
    app = NDTCalculator()
    app.mainloop()
