from utils.helpers import NAN_PATTERN, DOT_ZERO_PATTERN, MARKER_PATTERN
from views.components import *
from utils.helpers import MARKER_PATTERN
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import math
import datetime
import json
from tkcalendar import DateEntry
import traceback

def setup_daily_usage_tab_impl(self):
    """Setup the daily usage entry tab"""
    # Top frame for entry form
    # We use a canvas or large frame to allow free movement? 
    # Actually, we keep the entry_frame as the parent but use grid for initial layout
    # The user can then move them out of grid into place
    
    # Main Frame (No longer PanedWindow since we separated tabs)
    self.daily_usage_paned = ttk.Frame(self.tab_daily_usage)
    self.daily_usage_paned.pack(fill='both', expand=True, padx=5, pady=5)  # Reduced padding
    
    self.daily_usage_sash_locked = getattr(self, 'daily_usage_sash_locked', False)
    
    entry_frame = ttk.LabelFrame(self.daily_usage_paned, text="현장별 일일 사용량 기입")
    entry_frame.pack(fill='both', expand=True) # Changed from add to pack
    
    # Header area with two rows to prevent buttons from being hidden on small screens
    header_container = ttk.Frame(entry_frame)
    header_container.pack(fill='x', padx=2, pady=1)
    
    row1 = ttk.Frame(header_container)
    row1.pack(fill='x', pady=1)
    row2 = ttk.Frame(header_container)
    row2.pack(fill='x', pady=1)
    
    # Row 1: Primary Actions
    self.btn_daily_save = ttk.Button(row1, text="💾 저장", style='Action.TButton', command=self.add_daily_usage_entry, width=8)
    self.btn_daily_save.pack(side='left', padx=2)
    
    ttk.Button(row1, text="🧹 초기화", command=self.clear_daily_usage_form_all, width=10).pack(side='left', padx=2)
    
    btn_ndt_map = ttk.Button(row1, text="🧪 NDT 품목 매핑", command=self.open_ndt_product_map_dialog)
    btn_ndt_map.pack(side='left', padx=5)

    btn_sync = ttk.Button(row1, text="🔄 작업자 일괄 적용", command=self.sync_worker_times, width=20)
    btn_sync.pack(side='left', padx=5)

    btn_add_vehicle = ttk.Button(row1, text="🚙 추가 차량 점검", command=self.add_vehicle_inspection_box, width=15)
    btn_add_vehicle.pack(side='left', padx=5)

    self.btn_daily_report = ttk.Button(row1, text="📄 작업일보 출력", command=self.export_daily_work_report, width=15)
    self.btn_daily_report.pack(side='left', padx=5)

    btn_report_map = ttk.Button(row1, text="⚙️ 출력 설정", command=self.open_report_mapping_dialog)
    btn_report_map.pack(side='left', padx=5)

    # Sash lock button disabled since UI is separated
    # self.btn_sash_lock = ttk.Button(row1, text="🔒 경계 고정됨" if self.daily_usage_sash_locked else "🔓 경계 고정", command=self.toggle_sash_lock)
    # self.btn_sash_lock.pack(side='right', padx=5)

    # Row 2: Secondary / Tool Actions
    btn_save_sess = ttk.Button(row2, text="💾 세션 저장", command=self.save_form_session, width=12)
    btn_save_sess.pack(side='left', padx=5)
    
    btn_load_sess = ttk.Button(row2, text="📂 세션 불러오기", command=self.load_form_session, width=15)
    btn_load_sess.pack(side='left', padx=5)
    
    btn_load_prev = ttk.Button(row2, text="⏮️ 전일 데이터 불러오기", command=self.load_previous_day_data, width=20)
    btn_load_prev.pack(side='left', padx=5)
    
    # [UI REVISION] Removed dynamic add buttons (Vehicle, Checklist, Memo) 
    # as they are now permanently embedded at the bottom of the entry tab.
    if self.daily_usage_sash_locked:
        try:
            self.style.configure("SashLock.TButton", foreground="red")
            self.btn_sash_lock.configure(style="SashLock.TButton")
        except Exception: pass

    # [STABILITY FIX] Use a Canvas to hold the entry form for draggable support.
    # [REVISION] Restored scrollbar to satisfy the user request.
    canvas_parent = ttk.Frame(entry_frame)
    canvas_parent.pack(fill='both', expand=True, padx=2, pady=1)
    
    self.entry_canvas = tk.Canvas(canvas_parent, highlightthickness=0, bg=self.theme_bg)
    entry_vsb = ttk.Scrollbar(canvas_parent, orient="vertical", command=self.entry_canvas.yview)
    entry_vsb.pack(side='right', fill='y')
    self.entry_canvas.configure(yscrollcommand=entry_vsb.set)
    self.entry_canvas.pack(side='left', fill='both', expand=True)
    
    self.entry_inner_frame = ttk.Frame(self.entry_canvas)
    # Use a window to place the frame inside the canvas
    self.entry_canvas_window = self.entry_canvas.create_window((0, 0), window=self.entry_inner_frame, anchor='nw')
    
    def _on_entry_config(e):
        # Update canvas window width
        target_w = max(1100, e.width)
        
        # Use winfo_reqheight() so the frame can naturally expand when new widgets (like NDT) appear.
        # If it's forced to a fixed height, bottom elements like the treeview will be hidden.
        req_h = self.entry_inner_frame.winfo_reqheight()
        target_h = max(e.height, req_h)
        
        self.entry_canvas.itemconfig(self.entry_canvas_window, width=target_w, height=target_h)
        self._ensure_canvas_scroll_region()
    
    self.entry_inner_frame.bind("<Configure>", lambda e: self._ensure_canvas_scroll_region())
    self.entry_canvas.bind("<Configure>", _on_entry_config)
    
    # [NOTE] Mousewheel scrolling is handled by the global handler in __init__

    
    # Explicitly fix all possible grid rows to weight 0
    for r in range(100):
        self.entry_inner_frame.grid_rowconfigure(r, weight=0)
    self.entry_inner_frame.columnconfigure(0, weight=1)
    
    # 1. Unified Master Form Panel summerly
    self.master_form_panel = ttk.LabelFrame(self.entry_inner_frame, text="일일 검사 및 사용량 기록")
    self.master_form_panel.grid(row=0, column=0, sticky='nsew', padx=5, pady=5)

    
    # Configure columns inside the master panel (Reduced minsize)
    self.master_form_panel.columnconfigure(0, weight=0, minsize=350)
    self.master_form_panel.columnconfigure(1, weight=1, minsize=550)
    
            # Inner content for the basic form
    form_content = ttk.Frame(self.master_form_panel, padding=10)
    form_content.grid(row=0, column=0, rowspan=10, sticky='nw')
    
    for c in range(4): form_content.columnconfigure(c, weight=0)

    # Row 0: 업체명, 현장명
    ttk.Label(form_content, text="업체명:").grid(row=0, column=0, padx=(5, 0), pady=1, sticky='e')
    co_container = ttk.Frame(form_content)
    co_container.grid(row=0, column=1, padx=(2, 10), pady=1, sticky='w')
    self.cb_daily_company = ttk.Combobox(co_container, width=25, values=self.companies)
    self.cb_daily_company.pack(side='left')
    btn_company_mgr = tk.Button(co_container, text="⚙️ 관리", font=('Malgun Gothic', 8), bd=0, bg=self.theme_bg, fg='blue', cursor='hand2',
                               command=lambda: self.open_list_management_dialog('companies', target_cb=self.cb_daily_company))
    btn_company_mgr.pack(side='left', padx=(2, 0))

    ttk.Label(form_content, text="현장명:").grid(row=0, column=2, padx=(5, 0), pady=1, sticky='e')
    site_container = ttk.Frame(form_content)
    site_container.grid(row=0, column=3, padx=(2, 5), pady=1, sticky='w')
    self.cb_daily_site = ttk.Combobox(site_container, width=25, values=self.sites)
    self.cb_daily_site.pack(side='left')
    btn_site_mgr = tk.Button(site_container, text="⚙️ 관리", font=('Malgun Gothic', 8), bd=0, bg=self.theme_bg, fg='blue', cursor='hand2',
                            command=lambda: self.open_list_management_dialog('sites', target_cb=self.cb_daily_site))
    btn_site_mgr.pack(side='left', padx=(2, 0))

    # Row 1: 날짜, 장비명
    ttk.Label(form_content, text="날짜:").grid(row=1, column=0, padx=(5, 0), pady=1, sticky='e')
    from tkcalendar import DateEntry
    self.ent_daily_date = DateEntry(form_content, width=15, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd', locale='ko_KR', state='readonly')
    self.ent_daily_date.grid(row=1, column=1, padx=(2, 10), pady=1, sticky='w')

    ttk.Label(form_content, text="장비명:").grid(row=1, column=2, padx=(5, 0), pady=1, sticky='e')
    equip_container = ttk.Frame(form_content)
    equip_container.grid(row=1, column=3, padx=(2, 5), pady=1, sticky='w')
    
    self.ent_daily_equip = ttk.Entry(equip_container, width=15)
    self.ent_daily_equip.pack(side='left', fill='x', expand=True)
    
    # Link to the old variable name to avoid breaking other parts of the code
    self.cb_daily_equip = self.ent_daily_equip 
    
    # [MODERN] Place the search button INSIDE the entry widget at the right end
    btn_equip_search = tk.Button(self.ent_daily_equip, text="🔍", font=('Arial', 8), 
                                bd=0, bg='white', cursor='hand2',
                                command=self.open_equipment_search_dialog)
    btn_equip_search.place(relx=1.0, x=-2, rely=0.5, anchor='e', width=18, height=18)
    
    btn_equip_mgr = tk.Button(equip_container, text="⚙️ 관리", font=('Malgun Gothic', 8), bd=0, bg=self.theme_bg, fg='blue', cursor='hand2',
                             command=lambda: self.open_list_management_dialog('equipments', target_cb=self.ent_daily_equip))
    btn_equip_mgr.pack(side='left', padx=(2, 0))

    # Row 2: 품목명
    ttk.Label(form_content, text="품목명:").grid(row=2, column=0, padx=(5, 0), pady=1, sticky='e')
    mat_container = ttk.Frame(form_content)
    mat_container.grid(row=2, column=1, columnspan=3, padx=(2, 5), pady=1, sticky='w')
    
    self.cb_daily_material = ttk.Entry(mat_container, width=35)
    self.cb_daily_material.pack(side='left')
    
    # [MODERN] Place the search button INSIDE the entry
    btn_material_search = tk.Button(self.cb_daily_material, text="🔍", font=('Arial', 8), 
                                   bd=0, bg='white', cursor='hand2',
                                   command=lambda: self.open_material_search_dialog(target_form='daily_usage'))
    btn_material_search.place(relx=1.0, x=-2, rely=0.5, anchor='e', width=18, height=18)
    
    btn_material_mgr = tk.Button(mat_container, text="⚙️ 관리", font=('Malgun Gothic', 8), bd=0, bg=self.theme_bg, fg='blue', cursor='hand2',
                                command=lambda: self.open_list_management_dialog('materials', target_cb=self.cb_daily_material))
    btn_material_mgr.pack(side='left', padx=(2, 0))

    # Row 3: 방법, 검사품명
    ttk.Label(form_content, text="방법:").grid(row=3, column=0, padx=(5, 0), pady=1, sticky='e')
    self.cb_daily_test_method = ttk.Combobox(form_content, width=15, values=[' ', 'RT', 'PAUT', 'UT', 'MT', 'PT', 'ETC'])
    self.cb_daily_test_method.grid(row=3, column=1, padx=(2, 10), pady=1, sticky='w')
    
    ttk.Label(form_content, text="검사품명:").grid(row=3, column=2, padx=(5, 0), pady=1, sticky='e')
    insp_container = ttk.Frame(form_content)
    insp_container.grid(row=3, column=3, padx=(2, 5), pady=1, sticky='w')
    self.ent_daily_inspection_item = ttk.Entry(insp_container, width=15)
    self.ent_daily_inspection_item.pack(side='left')
    
    btn_insp_item_mgr = tk.Button(insp_container, text="⚙️ 관리", font=('Malgun Gothic', 8), bd=0, bg=self.theme_bg, fg='blue', cursor='hand2',
                                 command=lambda: self.open_list_management_dialog('test_items', target_cb=self.ent_daily_inspection_item))
    btn_insp_item_mgr.pack(side='left', padx=(2, 0))

    # Row 4: 수량, 단위
    ttk.Label(form_content, text="수량:").grid(row=4, column=0, padx=(5, 0), pady=1, sticky='e')
    self.ent_daily_test_amount = ttk.Entry(form_content, width=15)
    self.ent_daily_test_amount.grid(row=4, column=1, padx=(2, 10), pady=1, sticky='w')
    
    ttk.Label(form_content, text="단위:").grid(row=4, column=2, padx=(5, 0), pady=1, sticky='e')
    unit_container = ttk.Frame(form_content)
    unit_container.grid(row=4, column=3, padx=(2, 5), pady=1, sticky='w')
    self.cb_daily_unit = ttk.Combobox(unit_container, width=12, values=self.daily_units)
    self.cb_daily_unit.pack(side='left')
    
    btn_unit_mgr = tk.Button(unit_container, text="⚙️ 관리", font=('Malgun Gothic', 8), bd=0, bg=self.theme_bg, fg='blue', cursor='hand2',
                            command=lambda: self.open_list_management_dialog('daily_units', target_cb=self.cb_daily_unit))
    btn_unit_mgr.pack(side='left', padx=(2, 0))

    # Row 5: 단가, 출장비
    ttk.Label(form_content, text="단가:").grid(row=5, column=0, padx=(5, 0), pady=1, sticky='e')
    self.ent_daily_unit_price = ttk.Entry(form_content, width=15)
    self.ent_daily_unit_price.grid(row=5, column=1, padx=(2, 10), pady=1, sticky='w')

    ttk.Label(form_content, text="출장비:").grid(row=5, column=2, padx=(5, 0), pady=1, sticky='e')
    self.ent_daily_travel_cost = ttk.Entry(form_content, width=15)
    self.ent_daily_travel_cost.grid(row=5, column=3, padx=(2, 5), pady=1, sticky='w')

    # Row 6: 적용코드, 성적서번호
    ttk.Label(form_content, text="적용코드:").grid(row=6, column=0, padx=(5, 0), pady=1, sticky='e')
    app_container = ttk.Frame(form_content)
    app_container.grid(row=6, column=1, padx=(2, 10), pady=1, sticky='w')
    self.ent_daily_applied_code = ttk.Entry(app_container, width=20)
    self.ent_daily_applied_code.pack(side='left')
    
    btn_app_code_mgr = tk.Button(app_container, text="⚙️ 관리", font=('Malgun Gothic', 8), bd=0, bg=self.theme_bg, fg='blue', cursor='hand2',
                                command=lambda: self.open_list_management_dialog('applied_codes', target_cb=self.ent_daily_applied_code))
    btn_app_code_mgr.pack(side='left', padx=(2, 0))

    ttk.Label(form_content, text="성적서번호:").grid(row=6, column=2, padx=(5, 0), pady=1, sticky='e')
    self.ent_daily_report_no = ttk.Entry(form_content, width=18)
    self.ent_daily_report_no.grid(row=6, column=3, padx=(2, 5), pady=1, sticky='w')

    # Row 7: 비고, 일식
    ttk.Label(form_content, text="비고:").grid(row=7, column=0, padx=(5, 0), pady=1, sticky='e')
    self.ent_daily_note = ttk.Entry(form_content, width=15)
    self.ent_daily_note.grid(row=7, column=1, padx=(2, 10), pady=1, sticky='w')

    ttk.Label(form_content, text="일식:").grid(row=7, column=2, padx=(5, 0), pady=1, sticky='e')
    self.ent_daily_meal_cost = ttk.Entry(form_content, width=15)
    self.ent_daily_meal_cost.grid(row=7, column=3, padx=(2, 5), pady=1, sticky='w')

            # Row 8: 검사비
    ttk.Label(form_content, text="검사비:").grid(row=8, column=0, padx=(5, 0), pady=1, sticky='e')
    self.ent_daily_test_fee = ttk.Entry(form_content, width=15)
    self.ent_daily_test_fee.grid(row=8, column=1, padx=(2, 10), pady=1, sticky='w')

    # --- NDT 자동 산출 패널 ---
    self.ndt_calc_frame = ttk.LabelFrame(form_content, text="NDT 상세 조건 및 자동 계산", padding=5)
    self.ndt_calc_frame.grid(row=9, column=0, columnspan=4, sticky='ew', pady=(5,0))
    self.ndt_calc_frame.grid_remove() # 기본 숨김

    self.ndt_work_time_var = tk.StringVar(value="일반")
    self.ndt_loc_type_var = tk.StringVar(value="열배관")
    self.ndt_source_var = tk.StringVar(value="Se-75 (1.0)")
    self.ndt_thickness_var = tk.StringVar(value="조건없음 (1.0)")
    self.ndt_pipe_var = tk.StringVar(value="250mm 초과 [10인치 이상] (1.0)")
    self.ndt_overhead_var = tk.DoubleVar(value=110.0)
    self.ndt_tech_var = tk.DoubleVar(value=20.0)
    self.ndt_ori_joint_var = tk.StringVar(value="")
    self.ndt_ori_qty_var = tk.StringVar(value="")
    self.ndt_rep_joint_var = tk.StringVar(value="")
    self.ndt_rep_qty_var = tk.StringVar(value="")
    self.ndt_rej_joint_var = tk.StringVar(value="")
    self.ndt_report_pipe_var = tk.StringVar(value="")

    row0 = ttk.Frame(self.ndt_calc_frame)
    row0.pack(fill='x', pady=2)
    ttk.Label(row0, text="구분:").pack(side='left')
    ttk.Combobox(row0, textvariable=self.ndt_loc_type_var, values=["열배관", "플랜트(관리소)"], width=18, state="readonly").pack(side='left', padx=2)
    ttk.Label(row0, text="  작업형태:").pack(side='left', padx=(5,0))
    for t in ["일반", "야간", "휴일"]:
        ttk.Radiobutton(row0, text=t, value=t, variable=self.ndt_work_time_var).pack(side='left', padx=2)

    row1 = ttk.Frame(self.ndt_calc_frame)
    row1.pack(fill='x', pady=2)
    ttk.Label(row1, text="조건1:").pack(side='left', padx=(0,0))
    self.cb_ndt_cond1 = ttk.Combobox(row1, textvariable=self.ndt_source_var, width=22, state='readonly')
    self.cb_ndt_cond1.pack(side='left', padx=2)
    ttk.Label(row1, text="  조건2:").pack(side='left', padx=(5,0))
    self.cb_ndt_cond2 = ttk.Combobox(row1, textvariable=self.ndt_thickness_var, width=22, state='readonly')
    self.cb_ndt_cond2.pack(side='left', padx=2)

    row2 = ttk.Frame(self.ndt_calc_frame)
    row2.pack(fill='x', pady=2)
    ttk.Label(row2, text="보고서용 관경(Inch):").pack(side='left', padx=(0,0))
    self.cb_ndt_report_pipe = ttk.Combobox(row2, textvariable=self.ndt_report_pipe_var, width=10)
    self.cb_ndt_report_pipe.pack(side='left', padx=2)
    ttk.Label(row2, text="  제경비율(%):").pack(side='left', padx=(5,0))
    ttk.Entry(row2, textvariable=self.ndt_overhead_var, width=5).pack(side='left')
    ttk.Label(row2, text=" 기술료율(%):").pack(side='left', padx=(5,0))
    ttk.Entry(row2, textvariable=self.ndt_tech_var, width=5).pack(side='left')
    
    row3 = ttk.Frame(self.ndt_calc_frame)
    row3.pack(fill='x', pady=2)
    ttk.Label(row3, text="[ORI] 조인트:").pack(side='left', padx=(0,0))
    ttk.Entry(row3, textvariable=self.ndt_ori_joint_var, width=5).pack(side='left', padx=2)
    ttk.Label(row3, text=" 물량:").pack(side='left', padx=(0,0))
    ttk.Entry(row3, textvariable=self.ndt_ori_qty_var, width=5).pack(side='left', padx=2)
    ttk.Label(row3, text="  [REP] 조인트:").pack(side='left', padx=(5,0))
    ttk.Entry(row3, textvariable=self.ndt_rep_joint_var, width=5).pack(side='left', padx=2)
    ttk.Label(row3, text=" 물량:").pack(side='left', padx=(0,0))
    ttk.Entry(row3, textvariable=self.ndt_rep_qty_var, width=5).pack(side='left', padx=2)
    
    row4 = ttk.Frame(self.ndt_calc_frame)
    row4.pack(fill='x', pady=2)
    ttk.Label(row4, text="당일 불량(REJ)수:").pack(side='left', padx=(0,0))
    ttk.Entry(row4, textvariable=self.ndt_rej_joint_var, width=5).pack(side='left', padx=2)

    def _calculate_ndt_fee():
        try:
            from ndt_billing_tab import MATERIAL_COST, LABOR_COST
            ndt_type = self.cb_daily_test_method.get().strip()
            work_time = self.ndt_work_time_var.get()
            qty_str = self.ent_daily_test_amount.get().replace(',','')
            if not qty_str: return
            qty = float(qty_str)
            
            factor = 1.0
            if ndt_type == "RT":
                if "1.3" in self.ndt_source_var.get(): factor *= 1.3
                if "1.4" in self.ndt_thickness_var.get(): factor *= 1.4
                elif "2.2" in self.ndt_thickness_var.get(): factor *= 2.2
            elif ndt_type == "UT":
                pipe_val = self.ndt_pipe_var.get()
                if "1.2" in pipe_val: factor *= 1.2
                elif "1.4" in pipe_val: factor *= 1.4
                elif "1.7" in pipe_val: factor *= 1.7
                elif "2.0" in pipe_val: factor *= 2.0
                if "1.2" in self.ndt_thickness_var.get(): factor *= 1.2
            elif ndt_type == "PT":
                if "1.2" in self.ndt_pipe_var.get(): factor *= 1.2
                elif "1.4" in self.ndt_pipe_var.get(): factor *= 1.4
            
            adj_qty = qty * factor
            mat_type = self.cb_daily_material.get().strip().lower()
            mat_unit_cost = 0
            if ndt_type == "RT":
                if "b" in mat_type or "17" in mat_type:
                    mat_unit_cost = MATERIAL_COST.get('RT (B필름: 3⅓"x17")', 8867)
                elif "a/2" in mat_type or 'a/2' in mat_type or "6" in mat_type:
                    mat_unit_cost = MATERIAL_COST.get('RT (A/2필름: 3⅓"x6")', 7003)
                elif "a" in mat_type or "12" in mat_type:
                    mat_unit_cost = MATERIAL_COST.get('RT (A필름: 3⅓"x12")', 8025)
                else:
                    mat_unit_cost = MATERIAL_COST.get('RT (B필름: 3⅓"x17")', 8867)
            elif ndt_type == "UT":
                mat_unit_cost = MATERIAL_COST.get('UT', 1112)
            elif ndt_type == "PT":
                mat_unit_cost = MATERIAL_COST.get('PT', 3974)
            elif ndt_type == "PAUT":
                # Cond1 for PAUT is like "300A 이상 (1.0)". We need the string before the "(" to match the key
                # The key in MATERIAL_COST is "PAUT_300A 이상"
                paut_cond = self.ndt_pipe_var.get().split('(')[0].strip()
                mat_key = f"PAUT_{paut_cond}"
                mat_unit_cost = MATERIAL_COST.get(mat_key, 0)
            
            total_mat = int(qty * mat_unit_cost)
            
            loc_type_val = self.ndt_loc_type_var.get().strip()
            
            if loc_type_val in LABOR_COST:
                if ndt_type == "PAUT":
                    lab_key = f"PAUT_{self.ndt_pipe_var.get().split('(')[0].strip()}"
                    lab_unit = LABOR_COST[loc_type_val].get(work_time, {}).get(lab_key, 0)
                else:
                    lab_unit = LABOR_COST[loc_type_val].get(work_time, {}).get(ndt_type, 0)
            else:
                if ndt_type == "PAUT":
                    lab_key = f"PAUT_{self.ndt_pipe_var.get().split('(')[0].strip()}"
                    lab_unit = LABOR_COST.get(work_time, {}).get(lab_key, 0)
                else:
                    lab_unit = LABOR_COST.get(work_time, {}).get(ndt_type, 0)
                
            total_lab = int(adj_qty * lab_unit)
            
            import math
            try:
                overhead_rate = float(self.ndt_overhead_var.get()) / 100
                if math.isnan(overhead_rate): overhead_rate = 0.0
            except:
                overhead_rate = 0.0
            try:
                tech_rate = float(self.ndt_tech_var.get()) / 100
                if math.isnan(tech_rate): tech_rate = 0.0
            except:
                tech_rate = 0.0
                
            overhead = int(total_lab * overhead_rate)
            tech = int((total_lab + overhead) * tech_rate)
            
            subtotal = total_mat + total_lab + overhead + tech
            unit_price = subtotal / qty if qty > 0 else 0
            
            self.ent_daily_unit_price.delete(0, tk.END)
            self.ent_daily_unit_price.insert(0, f"{unit_price:,.0f}")
            
            self.ent_daily_test_fee.delete(0, tk.END)
            self.ent_daily_test_fee.insert(0, f"{subtotal:,.0f}")
            
            self._last_ndt_factor = factor
            self._last_ndt_adj_qty = adj_qty
            self._last_ndt_mat_cost = total_mat
            self._last_ndt_lab_cost = total_lab
            self._last_ndt_overhead = overhead
            self._last_ndt_tech = tech
            
        except Exception as e:
            import traceback; traceback.print_exc()

    ttk.Button(row2, text="자동 산출", command=_calculate_ndt_fee).pack(side='right', padx=5)

    
    # Restore focus transitions and defaults
    self.ent_daily_inspection_item.insert(0, "Piping")
    self.ent_daily_inspection_item.bind('<Return>', lambda e: self.ent_daily_test_amount.focus_set())
    self.ent_daily_test_amount.bind('<Return>', lambda e: self.cb_daily_unit.focus_set())
    self.cb_daily_unit.set('매')
    self.ndt_work_time_var.set("일반")
    self.ndt_source_var.set("Se-75 (1.0)")
    self.ndt_thickness_var.set("조건없음 (1.0)")
    self.ndt_pipe_var.set("250mm 초과 [10인치 이상] (1.0)")
    self.ndt_overhead_var.set(110.0)
    self.ndt_tech_var.set(20.0)
    self.ndt_calc_frame.grid_remove()
    self.cb_daily_unit.bind('<Return>', lambda e: self.ent_daily_unit_price.focus_set())
    self.cb_daily_unit.bind('<<ComboboxSelected>>', lambda e: self.ent_daily_unit_price.focus_set())
    self.ent_daily_unit_price.bind('<Return>', lambda e: self.ent_daily_applied_code.focus_set())
    self.ent_daily_applied_code.insert(0, "KS")
    self.ent_daily_applied_code.bind('<Return>', lambda e: self.ent_daily_report_no.focus_set())
    self.ent_daily_report_no.bind('<Return>', lambda e: self.ent_daily_note.focus_set())
    self.ent_daily_note.bind('<Return>', lambda e: self.ent_daily_meal_cost.focus_set())
    self.ent_daily_meal_cost.insert(0, "0")
    self.ent_daily_meal_cost.bind('<Return>', lambda e: self.ent_daily_test_fee.focus_set())
    
    def on_method_select_focus(e):
        self.root.after(10, self.ent_daily_inspection_item.focus_set)
    self.cb_daily_test_method.bind('<<ComboboxSelected>>', on_method_select_focus, add='+')
    self.cb_daily_test_method.bind('<Return>', on_method_select_focus, add='+')
    
    def on_method_change_auto_unit_logic(e):
        method = self.cb_daily_test_method.get().strip()
        print(f"[DEBUG] on_method_change_auto_unit_logic triggered. method='{method}'")
        unit_map = {'RT': '매', 'UT': 'P,M,I/D', 'MT': 'P,M,I/D', 'PT': 'P,M,I/D', 'PAUT': 'M,I/D'}
        if method in unit_map: self.cb_daily_unit.set(unit_map[method])
        
        if method in ["MT", "PT"]:
            try:
                self.ndt_frame.grid()
                self.empty_guide_frame.grid_remove() # Hide guide
            except:
                pass
        else:
            try:
                self.ndt_frame.grid_remove()
            except:
                pass

        
        if method in ["RT", "UT", "PT", "PAUT", "MT"]:
            try:
                self.ndt_calc_frame.grid(row=9, column=0, columnspan=4, sticky='ew', pady=(5,0))
                self.ndt_calc_frame.lift()
                self.root.after(50, self._ensure_canvas_scroll_region)
            except Exception as ex:
                print(f"Error in grid: {ex}")
            if method == "RT":
                self.rtk_grid.grid() # [NEW] Show RTK
                self.empty_guide_frame.grid_remove() # Hide guide
                self.rtk_grid.lift() # [FIX] Prevent overlay click-blocking
                self.master_form_panel.update_idletasks()
                self.cb_ndt_cond1.config(textvariable=self.ndt_source_var, values=["Se-75 (1.0)", "Ir-192 (1.0)"])
                self.cb_ndt_cond2.config(textvariable=self.ndt_thickness_var, values=["조건없음 (1.0)"])
            elif method == "UT":
                self.rtk_grid.grid_remove() # [NEW] Hide RTK
                self.empty_guide_frame.grid() # Show guide
                self.cb_ndt_cond1.config(textvariable=self.ndt_pipe_var, values=["250mm 초과 [10인치 이상] (1.0)", "200~250mm [8인치] (1.2)", "150~200mm [6인치] (1.4)", "100~150mm [4인치] (1.7)", "100mm 이하 [3인치 이하] (2.0)"])
                self.cb_ndt_cond2.config(textvariable=self.ndt_thickness_var, values=["조건없음 (1.0)"])
            elif method in ["PT", "MT"]:
                self.rtk_grid.grid_remove() # [NEW] Hide RTK
                # guide already hidden above
                self.cb_ndt_cond1.config(textvariable=self.ndt_pipe_var, values=["조건없음 (1.0)"])
                self.cb_ndt_cond2.config(textvariable=self.ndt_thickness_var, values=["조건없음 (1.0)"])
            elif method == "PAUT":
                self.rtk_grid.grid_remove() # [NEW] Hide RTK
                self.empty_guide_frame.grid() # Show guide
                self.cb_ndt_cond1.config(textvariable=self.ndt_pipe_var, values=["300A 이상 (1.0)", "250A (1.0)", "200A (1.0)", "150A-125A (1.0)", "100A 이하 (1.0)"])
                self.cb_ndt_cond2.config(textvariable=self.ndt_thickness_var, values=["조건없음 (1.0)"])
        else:
            self.ndt_calc_frame.grid_remove()
            self.rtk_grid.grid_remove() # [NEW] Hide RTK
            try: self.empty_guide_frame.grid_remove() 
            except: pass
        return "break" 
    self.cb_daily_test_method.bind('<<ComboboxSelected>>', on_method_change_auto_unit_logic, add='+')
    self.cb_daily_test_method.bind('<KeyRelease>', on_method_change_auto_unit_logic, add='+')
    
    # [ROBUST_AUTOCOMPLETE] Use standardized suggest system for Entry Form
    self._bind_combobox_word_suggest(self.cb_daily_site, lambda: sorted(list(set(self.sites))))
    self._bind_combobox_word_suggest(self.cb_daily_material, lambda: self._get_material_candidates(include_all=False))
    self._bind_combobox_word_suggest(self.cb_daily_equip, lambda: self._get_equipment_candidates(include_all=False))
    self._bind_combobox_word_suggest(self.ent_daily_inspection_item, lambda: self._get_inspection_item_candidates())
    self._bind_combobox_word_suggest(self.ent_daily_applied_code, lambda: self._get_applied_code_candidates())
    
    def _get_pipe_candidates():
        if not hasattr(self, 'daily_usage_df') or self.daily_usage_df.empty: return []
        if '관경(Inch)' not in self.daily_usage_df.columns: return []
        return sorted(list(set(str(x) for x in self.daily_usage_df['관경(Inch)'].dropna() if str(x).strip())))
    self._bind_combobox_word_suggest(self.cb_ndt_report_pipe, _get_pipe_candidates)

    # Category definitions for focus flow and loops
    ndt_materials = self.ndt_materials_all
    rtk_cats = ["센터미스", "농도", "마킹미스", "필름마크", "취급부주의", "고객불만", "기타", "총계"]



    # Row 1: NDT with Multi-Company Support
    # NDT 자재 소모량 프레임 (버튼을 타이틀 위치로 이동)
    lbl_frame = ttk.Frame(self.master_form_panel)
    ttk.Label(lbl_frame, text="NDT 자재 소모량 (회사별)", font=('Arial', 9, 'bold')).pack(side='left', padx=(0, 10))
    ttk.Button(lbl_frame, text="+ 회사 추가", command=self.add_ndt_company_section, width=12).pack(side='left', padx=2)
    ttk.Button(lbl_frame, text="- 마지막 회사 삭제", command=self.remove_last_ndt_company, width=15).pack(side='left', padx=2)
    
    self.ndt_frame = ttk.LabelFrame(self.master_form_panel, labelwidget=lbl_frame)
    self.ndt_frame.grid(row=1, column=1, padx=5, pady=2, sticky='new')
    self.ndt_frame.grid_remove() # [NEW] Hide NDT frame by default
    
    # Container for company-specific NDT sections
    self.ndt_company_container = ttk.Frame(self.ndt_frame)
    self.ndt_company_container.pack(fill='x', expand=True, padx=5, pady=2)
    
    # Store NDT entries by company index: {0: {mat: entry, ...}, 1: {...}}
    self.ndt_company_entries = []
    
    # Buttons are now placed inline inside add_ndt_company_section
    
    # Add default first company section
    self.add_ndt_company_section()

    # Removed draggable container logic

    # Row 2: RTK
    self.rtk_grid = ttk.LabelFrame(self.master_form_panel, text="RTK 분류")
    self.rtk_grid.grid(row=1, column=1, padx=5, pady=2, sticky='new')
    self.rtk_grid.grid_remove() # [NEW] Hide RTK by default
    
    for c in range(6): self.rtk_grid.columnconfigure(c, weight=1, uniform="ndt_rtk")
    self.rtk_entries = {}
    for i, cat in enumerate(rtk_cats):
        r = i // 3; col = (i % 3) * 2
        ttk.Label(self.rtk_grid, text=f"{cat}:", font=('Arial', 8)).grid(row=r, column=col, padx=1, pady=1, sticky='w')
        e = ttk.Entry(self.rtk_grid, width=6)
        e.grid(row=r, column=col+1, padx=1, pady=1, sticky='ew')
        self.rtk_entries[cat] = e
        
        # Focus transition to next RTK entry
        if i + 1 < len(rtk_cats) - 1: # Skip "총계" (the last one)
            next_cat = rtk_cats[i+1]
            e.bind('<Return>', lambda e, nc=next_cat: self.rtk_entries[nc].focus_set())
        elif cat != "총계": # From last editable RTK to save button container (approx)
            # We don't have a direct handle to the button, but we can focus the first worker's name
            e.bind('<Return>', lambda e: self.cb_daily_user.focus_set())
        
        # Bind auto-calculation
        if cat != "총계":
            e.bind('<KeyRelease>', lambda e: self.calculate_rtk_total())
    
    self.rtk_entries["총계"].config(state='readonly')
    # Removed draggable container logic

    # [NEW] PAUT/UT 빈 공간 채우기용 통합 안내 패널
    self.empty_guide_frame = ttk.LabelFrame(self.master_form_panel, text="PAUT / UT 검사 안내")
    self.empty_guide_frame.grid(row=1, column=1, padx=5, pady=2, sticky='new')
    self.empty_guide_frame.grid_remove() # Hide by default
    
    guide_lbl = ttk.Label(self.empty_guide_frame, text="💡 PAUT 및 UT 검사는 NDT 자재 및 RTK 입력이 필요하지 않습니다.\n(특이사항은 하단의 메모를 활용해 주세요.)", justify='center', padding=10)
    guide_lbl.pack(fill='both', expand=True)

    # [LAYOUT FIX] Removed obsolete minsize row configurations


    # [MIGRATION] Convert existing times to "익일" format and sort
    self.worktimes = self._migrate_worktimes(self.worktimes if hasattr(self, 'worktimes') else [])
    
    # Create a single container for all workers inside the master panel summerly
    workers_box_frame = ttk.Frame(self.master_form_panel)
    workers_box_frame.grid(row=0, column=1, rowspan=1, sticky='nsew', padx=5, pady=5)
    workers_box_frame.columnconfigure(0, weight=1)
    workers_box_frame.rowconfigure(1, weight=1)
    
    worker_hdr = ttk.Frame(workers_box_frame)
    worker_hdr.grid(row=0, column=0, sticky='ew', pady=(0, 2))
    ttk.Label(worker_hdr, text="작업자 기록", font=('Malgun Gothic', 9, 'bold')).pack(side='left')
    ttk.Button(worker_hdr, text="⚙", width=2, 
               command=lambda: self.open_list_management_dialog('users')).pack(side='right')

    workers_inner = ttk.Frame(workers_box_frame)
    workers_inner.grid(row=1, column=0, sticky='nsew')
    # Configure inner grid for workers (2 columns)
    for c in range(2): workers_inner.grid_columnconfigure(c, weight=1)
    
    def setup_worker_group(idx, row, col):
        # Each worker group is no longer individually draggable
        # Instead, they are sub-frames within the main workers_inner
        group_frame = ttk.LabelFrame(workers_inner, text=f"작업자 {idx}")
        group_frame.grid(row=row, column=col, padx=2, pady=1, sticky='ew')
        
        group = WorkerDataGroup(
            group_frame, worker_index=idx, users_list=getattr(self, 'users', []),
            enable_autocomplete=True,
            time_list=self.worktimes
        )
        group.pack(fill='x', expand=True, padx=1, pady=1)

        # Store global references for all workers (1-10)
        if idx == 1:
            self.cb_daily_user = group.composite
            self.ent_worktime1 = group.ent_worktime
            self.ent_ot1 = group.ent_ot
        
        # Always set index-specific attributes for sync/data access
        setattr(self, f'worker_group{idx}', group)
        setattr(self, f'cb_daily_user{idx}', group.composite)
        setattr(self, f'ent_worktime{idx}', group.ent_worktime)
        setattr(self, f'ent_ot{idx}', group.ent_ot)
        
        # Bindings for auto-save & auto-OT
        group.bind_name('<FocusOut>', lambda e: self.auto_save_to_list(e, group.cb_name, self.users, 'users'))
        group.bind_name('<Return>', lambda e: self.auto_save_to_list(e, group.cb_name, self.users, 'users'))
        
        group.bind_time('<FocusOut>', lambda e: self.auto_save_worktime(e, group.ent_worktime, 'worktimes'))
        group.bind_time('<Return>', lambda e: self.auto_save_worktime(e, group.ent_worktime, 'worktimes'))
        
        group.bind_ot('<FocusOut>', lambda e: self.auto_save_ot(e, group.ent_ot, 'ot_times'))
        group.bind_ot('<Return>', lambda e: self.auto_save_ot(e, group.ent_ot, 'ot_times'))
        
        return group_frame, group

    # 5 rows x 2 cols = 10 workers rearranged inside workers_inner
    for i in range(1, 6): setup_worker_group(i, i-1, 0)
    for i in range(6, 11): setup_worker_group(i, i-6, 1)

    # Removed draggable worker box logic

    # Bindings & Finalization
    calc_trigger = lambda e: self.update_daily_test_fee_calc()
    
    def on_qty_change(e):
        self.update_daily_test_fee_calc()
        
    # Handle Date Changes globally for recalculations
    def on_date_change(e=None):
        # Recalculate OT for all workers when date changes (weekend vs weekday rates)
        for i in range(1, 11):
            group = getattr(self, f'worker_group{i}', None)
            if group:
                self.calculate_and_update_ot(group.ent_worktime.get(), group.ent_worktime)

    self.ent_daily_date.bind('<<DateEntrySelected>>', on_date_change, add='+')
    
    # Initial call to set last known date for change detection
    try:
        self._last_daily_date = self.ent_daily_date.get_date()
    except:
        self._last_daily_date = None
    
    # [V19.2_AUTO_LOAD] Populate the list automatically on startup/setup
    self.root.after(1500, self.update_daily_usage_view)
    
    self.ent_daily_test_amount.bind('<KeyRelease>', on_qty_change)
    
    # [NEW] Add comma auto-formatting and Focus Transition on FocusOut/Return
    cost_entries = [
        self.ent_daily_test_amount,
        self.ent_daily_unit_price, 
        self.ent_daily_travel_cost, 
        self.ent_daily_test_fee
    ]
    
    for i, ent in enumerate(cost_entries):
        ent.bind('<KeyRelease>', calc_trigger, add='+')
        ent.bind('<FocusOut>', lambda e, widget=ent: self.format_entry_with_commas(e, widget), add='+')
        
        # Define focus transition
        def on_return(e, current_idx=i):
            # Format current entry first
            self.format_entry_with_commas(e, cost_entries[current_idx])
            # Move focus to next entry if possible
            if current_idx + 1 < len(cost_entries):
                cost_entries[current_idx + 1].focus_set()
            return "break"
            
        ent.bind('<Return>', on_return)
    
    self.update_material_combo()
    
    # --- Bottom Dashboard (Recent Entries + Fixed Panels) ---
    self.bottom_dashboard = ttk.PanedWindow(self.entry_inner_frame, orient=tk.HORIZONTAL)
    self.bottom_dashboard.grid(row=1, column=0, sticky='nsew', padx=5, pady=(10, 5))
    self.entry_inner_frame.grid_rowconfigure(1, weight=1)

    # 1. Left: Recent Entries Mini-table
    self.recent_frame = ttk.LabelFrame(self.bottom_dashboard, text="오늘의 입력 내역 (최근 기록)")
    self.bottom_dashboard.add(self.recent_frame, weight=9)
    
    # Create Treeview for recent entries
    columns = ("id", "date", "site", "loc_type", "method", "inspection_item", "material", "qty", "worker", "insp_type", "joint_count", "rej_count", "pipe_size")
    self.tv_recent = ttk.Treeview(self.recent_frame, columns=columns, show='headings', height=9)
    self.tv_recent['displaycolumns'] = ("date", "site", "loc_type", "method", "inspection_item", "material", "qty", "worker", "insp_type", "joint_count", "rej_count", "pipe_size")
    
    self.tv_recent.heading("date", text="날짜")
    self.tv_recent.heading("site", text="현장명")
    self.tv_recent.heading("loc_type", text="구분")
    self.tv_recent.heading("method", text="검사방법")
    self.tv_recent.heading("inspection_item", text="검사품명")
    self.tv_recent.heading("material", text="품목명")
    self.tv_recent.heading("qty", text="수량")
    self.tv_recent.heading("worker", text="작업자(첫번째)")
    self.tv_recent.heading("insp_type", text="검사구분")
    self.tv_recent.heading("joint_count", text="조인트수")
    self.tv_recent.heading("rej_count", text="불량수")
    self.tv_recent.heading("pipe_size", text="관경(Inch)")
    
    self.tv_recent.column("date", width=80, minwidth=80, stretch=False, anchor='center')
    self.tv_recent.column("site", width=100, minwidth=100, stretch=False, anchor='center')
    self.tv_recent.column("loc_type", width=100, minwidth=100, stretch=False, anchor='center')
    self.tv_recent.column("method", width=60, minwidth=60, stretch=False, anchor='center')
    self.tv_recent.column("inspection_item", width=80, minwidth=80, stretch=False, anchor='center')
    self.tv_recent.column("material", width=120, minwidth=120, stretch=False, anchor='center')
    self.tv_recent.column("qty", width=50, minwidth=50, stretch=False, anchor='center')
    self.tv_recent.column("worker", width=80, minwidth=80, stretch=False, anchor='center')
    self.tv_recent.column("insp_type", width=60, minwidth=60, stretch=False, anchor='center')
    self.tv_recent.column("joint_count", width=60, minwidth=60, stretch=False, anchor='center')
    self.tv_recent.column("rej_count", width=60, minwidth=60, stretch=False, anchor='center')
    self.tv_recent.column("pipe_size", width=80, minwidth=80, stretch=False, anchor='center')
    
    # Bind click to load record and delete
    self.tv_recent.bind('<<TreeviewSelect>>', self.on_recent_record_click)
    self.tv_recent.bind('<Delete>', self.delete_recent_entry)
    self.tv_recent.bind('<ButtonRelease-1>', lambda e: self.save_tab_config())
    
    # Right-click menu for deletion
    self.recent_menu = tk.Menu(self.tv_recent, tearoff=0)
    self.recent_menu.add_command(label="삭제", command=self.delete_recent_entry)
    
    def show_recent_menu(event):
        item = self.tv_recent.identify_row(event.y)
        if item:
            self.tv_recent.selection_set(item)
            self.recent_menu.tk_popup(event.x_root, event.y_root)
            
    self.tv_recent.bind("<Button-3>", show_recent_menu)
    
    recent_hsb = ttk.Scrollbar(self.recent_frame, orient="horizontal", command=self.tv_recent.xview)
    recent_hsb.pack(side='bottom', fill='x')
    
    self.tv_recent.pack(side='left', fill='both', expand=True, padx=2, pady=2)
    
    recent_vsb = ttk.Scrollbar(self.recent_frame, orient="vertical", command=self.tv_recent.yview)
    recent_vsb.pack(side='right', fill='y')
    self.tv_recent.configure(yscrollcommand=recent_vsb.set, xscrollcommand=recent_hsb.set)
    
    # 2. Middle: Vehicle Inspection (Fixed)
    self.fixed_vehicle_frame = ttk.LabelFrame(self.bottom_dashboard, text="차량점검 (상시 패널)")
    self.bottom_dashboard.add(self.fixed_vehicle_frame, weight=9)
    
    header_f = ttk.Frame(self.fixed_vehicle_frame)
    header_f.pack(fill='x', padx=2, pady=(2, 0))
    
    # [NEW] Add Manage Button safely using pack instead of place
    btn_manage = ttk.Button(header_f, text="⚙️ 차량 목록 설정", cursor='hand2')
    btn_manage.pack(side='right')
    btn_manage.config(command=lambda: self.open_list_management_dialog('차량 목록 관리', getattr(self, 'vehicles', []), 'vehicles'))

    self.fixed_vehicle_widget = VehicleInspectionWidget(self.fixed_vehicle_frame, theme_bg=self.theme_bg, vehicle_list=getattr(self, 'vehicles', []))
    self.fixed_vehicle_widget.pack(fill='both', expand=True, padx=2, pady=2)
    # Register the vehicle widget to be saved along with standard entry
    self.vehicle_widget = self.fixed_vehicle_widget

    # 3. Right: Memo (Fixed)
    self.fixed_memo_frame = ttk.LabelFrame(self.bottom_dashboard, text="메모 (상시 패널)")
    self.bottom_dashboard.add(self.fixed_memo_frame, weight=2)
    self.fixed_memo_text = tk.Text(self.fixed_memo_frame, wrap='word', height=5, width=10, font=('Arial', 10), bg=self.theme_bg, highlightthickness=0)
    self.fixed_memo_text.pack(fill='both', expand=True, padx=2, pady=2)
    # Store for data retrieval
    self.main_memo_text = self.fixed_memo_text
    # Store for data retrieval
    self.main_memo_text = self.fixed_memo_text

    # No longer clamp entry_inner_frame height
    pass


def update_daily_usage_view_impl(self):
    """Update the daily usage treeview with smarter filters and shift classification"""
    import re
    marker_pattern = MARKER_PATTERN
    
    # [FIX] Aggressively hide any active suggestion window
    filter_widgets = [
        getattr(self, 'cb_daily_filter_site', None),
        getattr(self, 'cb_daily_filter_material', None),
        getattr(self, 'cb_daily_filter_equipment', None),
        getattr(self, 'cb_daily_filter_worker', None),
        getattr(self, 'cb_daily_filter_vehicle', None)
    ]
    for widget in filter_widgets:
        if widget and hasattr(widget, '_suggestion_win'):
            widget._suggestion_win.hide()

    # [REMOVED] focus_set here stoles focus from the entry form after saving
    # if hasattr(self, 'daily_usage_tree'):
    #     self.daily_usage_tree.focus_set()

    # Clear current view
    for item in self.daily_usage_tree.get_children():
        self.daily_usage_tree.delete(item)
    
    # Get filter values
    start_date_str = self.ent_daily_start_date.get().strip()
    end_date_str = self.ent_daily_end_date.get().strip()
    filter_site = self.cb_daily_filter_site.get().strip() if hasattr(self, 'cb_daily_filter_site') else '전체'
    filter_company = self.cb_daily_filter_company.get().strip() if hasattr(self, 'cb_daily_filter_company') else '전체'
    filter_material = self.cb_daily_filter_material.get().strip() if hasattr(self, 'cb_daily_filter_material') else '전체'
    filter_equipment = self.cb_daily_filter_equipment.get().strip() if hasattr(self, 'cb_daily_filter_equipment') else '전체'
    filter_worker = self.cb_daily_filter_worker.get().strip() if hasattr(self, 'cb_daily_filter_worker') else '전체'
    filter_vehicle = self.cb_daily_filter_vehicle.get().strip() if hasattr(self, 'cb_daily_filter_vehicle') else '전체'
    filter_shift = self.cb_daily_filter_shift.get().strip() if hasattr(self, 'cb_daily_filter_shift') else '전체'
    
    # Ensure default values if empty
    if not filter_site: filter_site = '전체'
    if not filter_material: filter_material = '전체'
    if not filter_equipment: filter_equipment = '전체'
    if not filter_worker: filter_worker = '전체'
    if not filter_shift: filter_shift = '전체'
    
    # (Dropdown population moved to refresh_inquiry_filters for stability)
    
    # [V20_FORCE_FIRST_LOAD] Ensure 2025 records are seen by resetting start date on first load
    if not hasattr(self, '_daily_usage_first_load_done'):
        self._daily_usage_first_load_done = True
        if hasattr(self, 'ent_daily_start_date'):
            # Force to 2024 to catch 2025-04 records
            self.ent_daily_start_date.set_date(datetime.datetime(2024, 1, 1))
            start_date_str = "2024-01-01"
            
    # Parse dates
    try:
        start_date = datetime.datetime.strptime(start_date_str, '%Y-%m-%d') if start_date_str else None
        end_date = datetime.datetime.strptime(end_date_str, '%Y-%m-%d') if end_date_str else None
        if end_date:
            # Include the entire end date
            end_date = end_date + datetime.timedelta(days=1) - datetime.timedelta(seconds=1)
    except ValueError:
        # Silently handle or log if needed, avoid blocking the entire view
        start_date = datetime.datetime(2024, 1, 1)
        end_date = datetime.datetime.now()
    
    # Filter data
    filtered_df = self.daily_usage_df.copy()
    # [CRITICAL] Normalize columns to ensure detection (site_pairs) matches data lookups
    filtered_df.columns = [str(c).strip().replace(' ', '') for c in filtered_df.columns]
    
    print(f"DEBUG: [Daily Usage] Total records in DB: {len(filtered_df)}")
    
    # [V17_CRASH_PROOF_DATE] Maximum resilience to prevent Exit Code 1
    def robust_date_parse(val):
        if val is None or pd.isna(val) or str(val).strip() == '': return pd.NaT
        try:
            # 1. Handle already parsed dates
            if isinstance(val, (pd.Timestamp, datetime.datetime)):
                return val.replace(tzinfo=None)
            if isinstance(val, datetime.date):
                return pd.Timestamp(val)
                
            # 2. Handle Numeric/Excel Serial
            if isinstance(val, (int, float)):
                if 30000 < val < 60000:
                    return pd.to_datetime(val, unit='D', origin='1899-12-30').round('min').replace(tzinfo=None)
            
            # 3. Handle Strings
            s_val = str(val).strip()
            if s_val.replace('.','').isdigit():
                num = float(s_val)
                if 30000 < num < 60000:
                    return pd.to_datetime(num, unit='D', origin='1899-12-30').round('min').replace(tzinfo=None)
            
            # Standard pandas parse with dot-to-hyphen cleanup
            clean_val = s_val.replace('.', '-').replace(' ', '')
            d = pd.to_datetime(clean_val, errors='coerce')
            if pd.notna(d): return d.replace(tzinfo=None)
            
            # Final raw parse attempt
            d = pd.to_datetime(s_val, errors='coerce')
            if pd.notna(d): return d.replace(tzinfo=None)
        except:
            pass
        return pd.NaT

    # Apply robust parsing
    filtered_df['Date'] = filtered_df['Date'].apply(robust_date_parse)
    
    # [V14_REVERTED_STILL] Sort by Date (Descending)
    sort_date_col = 'Date'
    if 'EntryTime' in filtered_df.columns:
        filtered_df['EntryTime'] = filtered_df['EntryTime'].apply(robust_date_parse)
        # Use EntryTime as tie-breaker (Newest to Oldest)
        filtered_df = filtered_df.sort_values(by=['Date', 'EntryTime'], ascending=[False, False], na_position='last')
    else:
        filtered_df = filtered_df.sort_values(by=['Date'], ascending=[False], na_position='last')
        
    # [V16_INDEX_PRESERVE] Removed reset_index to keep original mapping to self.daily_usage_df
    # filtered_df = filtered_df.reset_index(drop=True)
    
    if start_date is not None:
        filtered_df = filtered_df[filtered_df['Date'] >= start_date]
    
    if end_date is not None:
        filtered_df = filtered_df[filtered_df['Date'] <= end_date]
    
    if filter_site != '전체' and 'Site' in filtered_df.columns:
        # [V18_LOOSE_FILTER] Use partial match for site to avoid missing records with minor spacing/naming differences
        filtered_df = filtered_df[filtered_df['Site'].astype(str).str.contains(filter_site, case=False, na=False, regex=False)]
    
    if filter_company != '전체' and '업체명' in filtered_df.columns:
        filtered_df = filtered_df[filtered_df['업체명'].astype(str).str.contains(filter_company, case=False, na=False, regex=False)]
        
    if filter_material != '전체' and 'MaterialID' in filtered_df.columns:
        # Also allow partial match on display name
        def check_mat(mid):
            d_name = self.get_material_display_name(mid)
            return filter_material.lower() in d_name.lower()
        filtered_df = filtered_df[filtered_df['MaterialID'].apply(check_mat)]
        
    if filter_worker != '전체':
        # Check all 10 worker columns
        worker_cols = ['User'] + [f'User{i}' for i in range(2, 11)]
        mask = pd.Series([False] * len(filtered_df), index=filtered_df.index)
        for col in worker_cols:
            if col in filtered_df.columns:
                mask |= filtered_df[col].astype(str).str.contains(filter_worker, case=False, na=False, regex=False)
        filtered_df = filtered_df[mask]
        # Support partial match for equipment
        filtered_df = filtered_df[filtered_df['장비명'].astype(str).str.contains(filter_equipment, na=False, case=False, regex=False)]
    
    if filter_vehicle != '전체' and '차량번호' in filtered_df.columns:
        # Support partial match for vehicle number
        filtered_df = filtered_df[filtered_df['차량번호'].astype(str).str.contains(filter_vehicle, na=False, case=False, regex=False)]
    
    # Filter by material if specified
    if filter_material != '전체':
        # Robust filtering: map every MaterialID in the current set to its display name 
        # and filter by matching the user's selection string. 
        # This correctly handles both master-linked and orphaned/deleted materials.
        def get_disp_name(mid): return self.get_material_display_name(mid)
        filtered_df = filtered_df[filtered_df['MaterialID'].apply(get_disp_name) == filter_material]

    # Filter by worker or shift (Smarter & Marker-Insensitive)
    if filter_worker != '전체' or filter_shift != '전체':
        worker_cols = ['User', 'User2', 'User3', 'User4', 'User5', 'User6', 'User7', 'User8', 'User9', 'User10']
        time_cols = ['WorkTime', 'WorkTime2', 'WorkTime3', 'WorkTime4', 'WorkTime5', 'WorkTime6', 'WorkTime7', 'WorkTime8', 'WorkTime9', 'WorkTime10']
        
        # Normalize filter text: remove spaces and lowercase for maximum flexibility
        fw_clean = filter_worker.replace(' ', '').lower()
        if fw_clean == '': fw_clean = '전체' # Handle empty input as showing all

        def row_matches(row):
            if fw_clean == '전체' and filter_shift == '전체':
                return True
            
            # Pre-clean filter text for comparison
            f_worker = fw_clean
            f_shift = f"({filter_shift})" if filter_shift != '전체' else None
            
            for i in range(len(worker_cols)):
                # Securely get worker and time data
                w_col = worker_cols[i]
                t_col = time_cols[i]
                
                if w_col not in row: continue
                w_val_raw = str(row[w_col]).strip()
                
                if not w_val_raw or w_val_raw.lower() in ['nan', '0.0', 'none', '']:
                    continue
                    
                t_val = str(row.get(t_col, '')).strip()
                
                # Worker match (Substring match on cleaned names)
                w_match = True
                if f_worker != '전체':
                    w_val_clean = w_val_raw.replace(' ', '').lower()
                    # Also handle records where marker is still in User column
                    w_val_clean = marker_pattern.sub('', w_val_clean).strip()
                    w_match = f_worker in w_val_clean
                
                # Shift match (Look for marker in WorkTime column)
                s_match = True
                if f_shift:
                    s_match = f_shift in t_val
                
                if w_match and s_match:
                    return True
            return False

        # Apply robust filtering
        filtered_df = filtered_df[filtered_df.apply(row_matches, axis=1)]
    
    # [V10] Sorting is now handled at the beginning of the function for better reliability
    
    # Define RTK categories
    rtk_categories = ["센터미스", "농도", "마킹미스", "필름마크", "취급부주의", "고객불만", "기타", "총계"]
    
    # Display entries and calculate totals
    total_rtk = [0.0] * len(rtk_categories)
    total_ndt = [0.0] * 7
    total_test_amount = 0.0
    method_totals = {}
    total_unit_price = 0.0
    total_travel_cost = 0.0
    total_meal_cost = 0.0
    total_test_fee = 0.0
    total_film_count = 0.0 # [NEW] Added for film count
    total_ot_hours = 0.0
    total_ot_amount = 0
    total_work_hours = 0.0 # Added for Total Working Time
    total_mileage = 0.0    # Sum of individual entries if applicable
    min_mileage = float('inf')
    max_mileage = float('-inf')
    total_indiv_ot_hours = [0.0] * 10
    total_indiv_ot_amounts = [0] * 10
    
    has_equip = False
    has_method = False
    has_note = False
    has_product = False
    has_entry_time = False
    has_vehicle_no = False
    has_mileage = False
    has_veh_insp = False
    has_veh_note = False
    has_co_code = False
    has_company = False
    has_applied_code = False # [NEW]
    has_insp_item = False    # [NEW]
    has_report_no = False    # [NEW]
    has_meal_cost = False    # [NEW]
    has_unit = False         # [NEW]
    
    
    current_date = None
    
    # Seen sets for deduping
    seen_entry_times = set()
    seen_contents = set()
    
    for idx, row in filtered_df.iterrows():
        entry = row.to_dict()
        
        # Metadata formatting
        usage_date = self._safe_format_datetime(entry.get('Date', ''), '%Y-%m-%d')
        if not usage_date: usage_date = "Unknown"
        
        mat_id = entry.get('MaterialID', '')
        mat_name = self.get_material_display_name(mat_id)
        
        # Re-calculate workers/worktime early for deduping
        import re as _re
        def clean_s(v): return self.clean_nan(v)
        raw_workers = []
        for j in range(1, 11):
            u_k = 'User' if j == 1 else f'User{j}'
            u_v = clean_s(entry.get(u_k, ''))
            if u_v: raw_workers.append(u_v)
        
        c_workers = self.format_worker_summary(raw_workers)
        c_worktime = clean_s(entry.get('WorkTime', '')) if raw_workers else ""

        # Timestamp-based Key (Safety) - Define early for use in deduplication keys
        e_t_raw = entry.get('EntryTime', '')
        try:
            if isinstance(e_t_raw, (pd.Timestamp, datetime.datetime)):
                t_key = e_t_raw.strftime('%Y-%m-%d %H:%M:%S')
            else:
                t_key = str(e_t_raw).split('.')[0].strip()
        except:
            t_key = str(e_t_raw).strip()

        # Content-based Unique Key (Date, Site, WorkTime, Method, EntryTime, ItemInfo)
        # [FIX] Include EntryTime (t_key) to ensure separate saves are never merged visually.
        # Also include Material, Item, and Code for maximum granularity within a single save if needed.
        n_site = str(entry.get('Site', '')).strip()
        n_date = usage_date
        n_method = str(entry.get('검사방법', '')).strip()
        n_insp = str(entry.get('검사품명', '')).strip()
        n_code = str(entry.get('적용코드', '')).strip()
        
        content_key = (n_date, n_site, c_worktime, n_method, t_key, mat_id, n_insp, n_code)
        
        # Record is duplicate split IF (Timestamp matches) AND (Content matches)
        # [FIX] Changed to OR but because t_key is in content_key, separate saves will always be unique.
        is_duplicate_split = (t_key and t_key in seen_entry_times) or (content_key in seen_contents)
        
        if t_key: seen_entry_times.add(t_key)
        seen_contents.add(content_key)

        # Metadata formatting (re-mapped for display later)
        entry_time_display = self._safe_format_datetime(entry.get('EntryTime', ''), '%Y-%m-%d %H:%M:%S')

        # Worker extraction and WorkTime determination (for display)
        consolidated_workers = c_workers
        display_worktime = ""
        
        # Worker filtering logic
        if filter_worker != '전체':
            consolidated_workers = filter_worker
            f_w_c = filter_worker.replace(' ', '').lower()
            for j in range(1, 11):
                u_k = 'User' if j == 1 else f'User{j}'
                u_v_raw = clean_s(entry.get(u_k, ''))
                u_v_c = marker_pattern.sub('', u_v_raw.replace(' ', '').lower()).strip()
                if u_v_c == f_w_c:
                    wt_k = 'WorkTime' if j == 1 else f'WorkTime{j}'
                    display_worktime = clean_s(entry.get(wt_k, ''))
                    break
        else:
            display_worktime = c_worktime

        # Numeric calculations
        def to_f_local(v):
            try:
                if pd.isna(v) or str(v).lower() == 'nan': return 0.0
                return float(str(v).replace(',', ''))
            except: return 0.0

        q_val = to_f_local(entry.get('검사량', entry.get('수량', 0.0)))
        p_val = to_f_local(entry.get('단가', 0.0))
        t_val_cost = to_f_local(entry.get('출장비', 0.0))
        m_val_cost = to_f_local(entry.get('일식', 0.0))
        f_val_cost = to_f_local(entry.get('검사비', 0.0))
        
        # Sum totals (Guarded by duplicate check to prevent double-counting)
        if not is_duplicate_split:
            total_test_amount += q_val
            method = str(entry.get('검사방법', '')).strip().upper()
            if method:
                method_totals[method] = method_totals.get(method, 0.0) + q_val
            total_unit_price += p_val
            total_travel_cost += t_val_cost
            total_meal_cost += m_val_cost
            total_test_fee += f_val_cost
            total_film_count += to_f_local(entry.get('FilmCount', 0.0))

        # Cumulative mileage (Always sum)
        milk = to_f_local(entry.get('주행거리', entry.get('거리', 0)))
        total_mileage += milk
        if milk > 0.001:
            min_mileage = min(min_mileage, milk) if min_mileage != float('inf') else milk
            max_mileage = max(max_mileage, milk) if max_mileage != float('-inf') else milk

        # OT hours and amounts calculation
        row_ot_hours = 0.0
        row_ot_amount = 0
        row_ots = []
        
        # Exact column name matching for each worker slot
        site_pairs = []
        all_keys = set(entry.keys())
        for j in range(1, 11):
            # j=1: 'User', j=2: 'User2', etc.
            uk = ('User' if j == 1 else f'User{j}') if ('User' if j == 1 else f'User{j}') in all_keys else None
            wk = ('WorkTime' if j == 1 else f'WorkTime{j}') if ('WorkTime' if j == 1 else f'WorkTime{j}') in all_keys else None
            ok = ('OT' if j == 1 else f'OT{j}') if ('OT' if j == 1 else f'OT{j}') in all_keys else None
            if uk:
                site_pairs.append((uk, wk, ok))
        
        for i in range(1, 11):
            if i <= len(site_pairs):
                uk, wk, ok = site_pairs[i-1]
                uv = clean_s(entry.get(uk, ''))
                
                if not uv or (filter_worker != '전체' and uv != filter_worker):
                    row_ots.append("")
                    continue

                ots = str(entry.get(ok, '')).strip()
                wts = str(entry.get(wk, '')).strip()
                
                if ots and ots not in ('nan', '0.0', '0'):
                    try:
                        if '(' in ots and '원)' in ots:
                            h_p = float(ots.split('시간')[0])
                            a_p = int(_re.sub(r'[^0-9]', '', ots.split('(')[1].split('원')[0]))
                        elif ots.replace(',', '').isdigit():
                            a_p = int(ots.replace(',', ''))
                            h_p, _ = self._calculate_ot_from_worktime(wts, pd.to_datetime(entry.get('Date', datetime.datetime.now())))
                        else:
                            a_p = 0
                            h_p = self._parse_ot_hours(ots)
                            
                        # Activity-based hours: Max of all workers in this row
                        row_ot_hours = max(row_ot_hours, h_p)
                        # Cost-based amounts: Always sum across all workers
                        row_ot_amount += a_p
                        
                        if not is_duplicate_split:
                            # Update global individual totals (primarily for column data presence/amounts)
                            total_ot_amount += a_p
                            total_indiv_ot_hours[i-1] += h_p
                            total_indiv_ot_amounts[i-1] += a_p
                            
                        row_ots.append(f"{a_p:,}")
                    except:
                        row_ots.append(ots)
                else:
                    row_ots.append("")  # Empty when no OT data (not '0')
            else: row_ots.append("")
        
        # Trim trailing empty OT slots so inactive workers don't create visible columns
        while row_ots and row_ots[-1] == "":
            row_ots.pop()
        # Re-pad to 10 with empty strings (the Treeview expects 10 slots)
        while len(row_ots) < 10:
            row_ots.append("")

        # Global OT Hours: Sum of per-activity maximums
        if not is_duplicate_split:
            total_ot_hours += row_ot_hours

        # Global Work Hours (ONLY for primary rows)
        if not is_duplicate_split and display_worktime and '~' in str(display_worktime):
            try:
                cwt = marker_pattern.sub('', str(display_worktime)).strip()
                if '~' in cwt:
                    pts = cwt.split('~')
                    sh, sm = map(int, pts[0].split(':'))
                    eh, em = map(int, pts[1].split(':'))
                    sm_t = sh * 60 + sm
                    em_t = eh * 60 + em
                    if em_t < sm_t: em_t += 1440
                    total_work_hours += (em_t - sm_t) / 60.0
            except: pass

        # RTK values
        def robust_to_f(v):
            if pd.isna(v) or str(v).lower() in ('nan', 'none', ''): return 0.0
            try:
                cl = _re.sub(r'[^0-9\.\-]', '', str(v))
                return float(cl) if cl else 0.0
            except: return 0.0

        rtk_vals = []
        row_rtk_sum = 0.0
        for i, cat in enumerate(['센터미스', '농도', '마킹미스', '필름마크', '취급부주의', '고객불만', '기타']):
            v = robust_to_f(entry.get(f'RTK_{cat}', 0))
            rtk_vals.append(f"{v:.1f}" if abs(v) > 0.001 else "")
            row_rtk_sum += v
            if not is_duplicate_split: total_rtk[i] += v
        rtk_vals.append(f"{row_rtk_sum:.1f}" if abs(row_rtk_sum) > 0.001 else "")
        if not is_duplicate_split: total_rtk[7] += row_rtk_sum
        
        # NDT values
        ndt_vals = []
        for i, mat in enumerate(["형광자분", "흑색자분", "백색페인트", "침투제", "세척제", "현상제", "형광침투제"]):
            v = robust_to_f(entry.get(f'NDT_{mat}'))
            if v == 0:
                if mat == "흑색자분": v = robust_to_f(entry.get('NDT_자분'))
                elif mat == "백색페인트": v = robust_to_f(entry.get('NDT_페인트'))
                elif mat == "형광침투제": v = robust_to_f(entry.get('NDT_형광'))
            ndt_vals.append(f"{v:.1f}" if abs(v) > 0.001 else "")
            if not is_duplicate_split: total_ndt[i] += v

        # Remarks and visibility checks
        note_raw = str(entry.get('Note', '')).strip() if not pd.isna(entry.get('Note')) else ""
        display_remark = f"[{consolidated_workers}] {note_raw}" if consolidated_workers and note_raw else note_raw
        
        def has_content(v):
            if pd.isna(v) or str(v).lower() in ('nan', 'none', '') or str(v).strip() in ('', '-', '0', '0.0', '0시간'): return False
            return True
        
        if has_content(entry.get('장비명')): has_equip = True
        if has_content(entry.get('검사방법')): has_method = True
        if has_content(note_raw): has_note = True
        if has_content(entry.get('MaterialID')): has_product = True
        if has_content(entry.get('EntryTime')): has_entry_time = True
        if has_content(entry.get('차량번호')): has_vehicle_no = True
        if has_content(entry.get('주행거리')) or has_content(entry.get('거리')): has_mileage = True
        if has_content(entry.get('차량점검')): has_veh_insp = True
        if has_content(entry.get('차량비고')): has_veh_note = True
        if has_content(entry.get('회사코드')): has_co_code = True
        if has_content(entry.get('업체명')): has_company = True
        if has_content(entry.get('적용코드')): has_applied_code = True
        if has_content(entry.get('검사품명')): has_insp_item = True
        if has_content(entry.get('성적서번호')): has_report_no = True
        if has_content(entry.get('일식')): has_meal_cost = True
        if has_content(entry.get('Unit')): has_unit = True
        
        # [VISUAL DEDUPING] If duplicate split row, clear numeric values for cleaner view/export
        disp_q = f"{q_val:.1f}"
        disp_p = f"{p_val:,.0f}"
        disp_t = f"{t_val_cost:,.0f}"
        disp_f = f"{f_val_cost:,.0f}"
        disp_m = f"{m_val_cost:,.0f}"
        disp_oth = f"{row_ot_hours:.1f}"
        disp_ota = f"{row_ot_amount:,}"

        disp_row_ots = row_ots
        disp_rtk = rtk_vals
        disp_ndt = ndt_vals
        
        if is_duplicate_split:
            disp_q = ""
            disp_p = ""
            disp_t = ""
            disp_m = ""
            disp_f = ""
            disp_oth = ""
            disp_ota = ""
            disp_row_ots = [""] * 10
            disp_rtk = [""] * 8
            disp_ndt = [""] * 7
        
        v_tuple = (
            usage_date,
            entry.get('업체명', ''),
            entry.get('적용코드', ''),
            entry.get('Site', ''),
            entry.get('구분', ''),
            entry.get('검사품명', ''),
            entry.get('성적서번호', ''),
            consolidated_workers, # Header: '작업자' (Index 6)
            display_worktime,     # Header: '작업시간' (Index 7)
            *disp_row_ots,        # OT1..OT10 (Index 8..17)
            entry.get('장비명', ''), # Header: '장비명' (Index 18)
            entry.get('검사방법', ''),# Header: '검사방법' (Index 19)
            entry.get('회사코드', ''), # Index 20
            disp_q,               # Index 21 (수량)
            entry.get('Unit', ''), # Index 22 (단위) [NEW]
            disp_p,               # Index 23 (단가)
            disp_t,               # Index 23 (출장비)
            disp_m,               # Index 24 (일식)
            disp_f,               # Index 25 (검사비)
            disp_oth,             # Index 26 (OT시간)
            disp_ota,             # Index 27 (OT금액)
            mat_name,             # Index 28 (품목명)
            *disp_rtk,            # Index 29..36
            *disp_ndt,            # Index 37..43
            display_remark,       # Index 44
            str(entry.get('EntryTime', '')),
            self.clean_nan(entry.get('차량번호', '')),
            self.clean_nan(entry.get('주행거리', '')),
            self.clean_nan(entry.get('차량점검', '')),
            self.clean_nan(entry.get('차량비고', '')),
            ", ".join(raw_workers)
        )


        v_list = list(v_tuple)
        tree_cols = self.daily_usage_tree['columns']
        while len(v_list) < len(tree_cols): v_list.append("")
        
        for c_idx, c_name in enumerate(tree_cols):
            if c_name in ['작업형태', '조건1', '조건2', '보정계수', '제경비', '기술료', '환산물량', '재료비', '인건비', '검사구분', '조인트수', '불량수', '관경(Inch)']:
                val = entry.get(c_name, '')
                if c_name in ['보정계수', '환산물량'] and val:
                    try: 
                        f_val = float(val)
                        val = f"{int(f_val)}" if f_val.is_integer() else f"{f_val}"
                    except: pass
                elif c_name in ['재료비', '인건비', '제경비', '기술료'] and val:
                    try: val = f"{int(float(val)):,}"
                    except: pass
                v_list[c_idx] = val
        v_tuple = tuple(v_list)
        self.daily_usage_tree.insert('', tk.END, values=v_tuple, tags=(str(idx),))
        
    # Insert last daily subtotal and final total row if data exists
    if not filtered_df.empty:
        
        # Final overall total
        self.daily_usage_tree.tag_configure('total', background='#E8F4F8', font=('Arial', 9, 'bold'))
        
        total_rtk_ready = [f"{v:.1f}" if abs(v) > 0.001 else "" for v in total_rtk]
        while len(total_rtk_ready) < 8: total_rtk_ready.append("")
        
        total_ndt_ready = [f"{v:.1f}" if abs(v) > 0.001 else "" for v in total_ndt]
        while len(total_ndt_ready) < 7: total_ndt_ready.append("")

        total_values = [
            '--- 전체 누계 ---',
            '', # 업체명 (1)
            '', # 적용코드 (2)
            '', # 현장 (3)
            '', # 구분 (4)
            '', # 검사품명 (5)
            '', # 성적서번호 (6)
            '', # 작업자 (6)
            f"{total_work_hours:.1f} Hrs" if total_work_hours > 0.001 else "", # 작업시간 (7)
            # Individual OT Totals (Simplified: Amount only)
            *[f"{a:,}" if a > 0.001 else "" for a in total_indiv_ot_amounts],
            '', # 장비명
            '', # 검사방법
            '', # 회사코드
            f"{total_test_amount:.1f}" if total_test_amount > 0.001 else "",
            '', # 단위 (Unit)
            '', # 단가 (Unit Price is not summed)
            f"{total_travel_cost:,.0f}" if total_travel_cost > 0.001 else "",
            f"{total_meal_cost:,.0f}" if total_meal_cost > 0.001 else "", # Added back missing index
            f"{total_test_fee:,.0f}" if total_test_fee > 0.001 else "",
            f"{total_ot_hours:.1f}" if total_ot_hours > 0.001 else "", # OT시간 합계
            f"{total_ot_amount:,}" if total_ot_amount > 0.001 else "",  # OT금액 합계
            '', # 품목명
            *total_rtk_ready,
            *total_ndt_ready,
            '',   # 비고
            '',   # 입력시간
            '',   # 차량번호
            f"누계: {total_mileage:,.1f} km" if total_mileage > 0.001 else "0 km",   # 주행거리 합계
            '',   # 차량점검
            '',   # 차량비고
            ''    # (Full작업자)
        ]
        # [DEFENSIVE] Ensure Total row matches header count exactly
        while len(total_values) < len(self.daily_usage_tree['columns']): total_values.append("")
        self.daily_usage_tree.tag_configure('total', background='#E8F4F8', font=('Arial', 12, 'bold'))
        self.daily_usage_tree.insert('', tk.END, values=total_values, tags=('total',))
        
        # --- Dynamic Column Hiding ---
        # Mandatory cols (always show to maintain core row identity)
        # [USER REQUEST] Force '센터미스' and '농도' to be mandatory/always shown
        mandatory_cols = ['날짜', '현장', '작업자']
        
        # Use a slightly more robust threshold for data presence
        # Also handle potential string '0.0' leftovers and common empty markers
        def is_active(val):
            if val is None: return False
            s = str(val).strip().lower()
            # Explicitly list all known 'zero' or 'empty' string representations
            if s in ('', '0', '0.0', '0.00', 'nan', 'none', '-', '0.0시간', '0.0(0원)', '0.0 (0원)', '0시간', '0원', '0(0원)', '0 (0원)'):
                return False
            try:
                # Robust cleaning: remove everything except numbers, dots, and minus
                clean_s = re.sub(r'[^0-9\.\-]', '', s)
                if not clean_s: return False # Only markers, but no number? Treat as empty.
                v = float(clean_s)
                return abs(v) > 0.001 # Slightly wider epsilon
            except:
                # If it's a non-numeric string (like a note), return True if not empty
                return bool(s)
        
        dynamic_col_status = {
            '작업시간': is_active(total_work_hours),
            '수량': is_active(total_test_amount),
            '단가': is_active(total_unit_price),
            '출장비': is_active(total_travel_cost),
            '검사비': is_active(total_test_fee),
            'OT시간': is_active(total_ot_hours),
            'OT금액': is_active(total_ot_amount),
            '장비명': has_equip,
            '검사방법': has_method,
            '비고': has_note,
            '차량번호': has_vehicle_no,
            '주행거리': has_mileage,
            '차량점검': has_veh_insp,
            '차량비고': has_veh_note,
            '품목명': has_product, # Now content-based
            '입력시간': has_entry_time, # Now content-based
            '회사코드': has_co_code,
            '업체명': has_company,
            '적용코드': has_applied_code,
            '검사품명': has_insp_item,
            '성적서번호': has_report_no,
            '일식': has_meal_cost,
            '단위': has_unit,
            '수량': is_active(total_test_amount), # [FIX] Ensure these are explicitly handled
            '단가': is_active(total_unit_price),
            '출장비': is_active(total_travel_cost),
            '검사비': is_active(total_test_fee)
        }
        
        # Individual OT columns
        for i in range(1, 11):
            col_name = f'OT{i}'
            dynamic_col_status[col_name] = is_active(total_indiv_ot_amounts[i-1])
        
        # RTK Columns
        rtk_col_names = ["센터미스", "농도", "마킹미스", "필름마크", "취급부주의", "고객불만", "기타"]
        for i, col_name in enumerate(rtk_col_names):
            active_status = is_active(total_rtk[i])
            # [USER REQUEST] Force hide '마킹미스' and others if they have no significant data
            if col_name in ["마킹미스", "필름마크", "취급부주의", "고객불만", "기타"]:
                dynamic_col_status[col_name] = active_status
            else:
                # '센터미스' and '농도' should follow mandatory_cols but we safeguard here too
                dynamic_col_status[col_name] = active_status
        
        # Diagnostic Log (visible in terminal for aid)
        # print(f"DEBUG RTK Totals: {dict(zip(rtk_col_names, total_rtk[:7]))}, Total: {total_rtk[7]}")
        
        dynamic_col_status['RTK총계'] = is_active(total_rtk[7])
        
        # NDT Columns
        ndt_col_names = ["형광자분", "흑색자분", "백색페인트", "침투제", "세척제", "현상제", "형광침투제"]
        for i, col_name in enumerate(ndt_col_names):
            dynamic_col_status[col_name] = is_active(total_ndt[i])
            
        # --- BUILD FINAL VISIBILITY (MERGED LOGIC) ---
        all_cols = list(self.daily_usage_tree['columns'])
        # manual_set: columns user explicitly hid via column manager (empty = no exclusions)
        manual_visible_list = getattr(self, 'manual_visible_cols', [])
        if not manual_visible_list:
            manual_visible_list = all_cols
        manual_hidden = set(all_cols) - set(manual_visible_list)
        
        # [SAFETY] Core columns that should almost never be hidden unless user is very specific
        # [REFINED] Minimal mandatory columns to allow smarter auto-hiding of empty fields
        mandatory_cols = ['날짜', '현장', '작업자']
        
        final_visible = []
        # We iterate in ALL_COLS order to maintain the original column sequence
        for col in all_cols:
            if col == '(Full작업자)': continue
            
            # 1. Mandatory override: Always show core columns
            if col in mandatory_cols:
                final_visible.append(col)
                continue
            
            # 2. SMART HIDING: For columns we track data status for
            if col in dynamic_col_status:
                # Show ONLY if it has data AND isn't manually hidden
                if dynamic_col_status[col] and col not in manual_hidden:
                    final_visible.append(col)
                # Skip untracked fallback for this column (i.e. if NO data, it's HIDDEN)
                continue
            
            # 3. Fallback for untracked columns (e.g. manually added custom columns)
            if col not in manual_hidden:
                final_visible.append(col)

        # [STABILITY] Clear the Treeview's displayed columns
        self.daily_usage_tree['displaycolumns'] = final_visible
        

        
        # Header renames (if any)

        # Ensure stretch=False and minwidth is relaxed for ALL displayed columns
        # This is critical for making all columns, especially the last one, resizable
        for col in final_visible:
            self.daily_usage_tree.column(col, stretch=False, minwidth=20)

        # Re-apply saved column widths if available
        if hasattr(self, 'tab_config') and 'daily_usage_col_widths' in self.tab_config:
            saved_widths = self.tab_config['daily_usage_col_widths']
            for col in final_visible:
                if col in saved_widths:
                    try:
                        self.daily_usage_tree.column(col, width=int(saved_widths[col]), stretch=False)
                        # Enforce minimums for high-precision cols to prevent truncation
                        if col == '날짜':
                            self.daily_usage_tree.column(col, width=max(int(saved_widths[col]), 160), stretch=False)
                        elif col == '입력시간':
                            self.daily_usage_tree.column(col, width=max(int(saved_widths[col]), 300), stretch=False)
                    except: pass
        
        # Ensure Total Row stays at bottom
        self.daily_usage_tree.detach(self.daily_usage_tree.get_children()[-1])
        self.daily_usage_tree.insert('', tk.END, values=total_values, tags=('total',))
    else:
        # If empty, show only mandatory columns to keep the view clean
        all_cols = self.daily_usage_tree['columns']
        final_visible = ['날짜', '현장', '작업자']
        self.daily_usage_tree['displaycolumns'] = final_visible
        
        # Standard column setup for display
        for col in final_visible:
            self.daily_usage_tree.column(col, stretch=False, minwidth=20)
            
        # Re-apply saved column widths if available
        if hasattr(self, 'tab_config') and 'daily_usage_col_widths' in self.tab_config:
            saved_widths = self.tab_config['daily_usage_col_widths']
            for col in final_visible:
                if col in saved_widths:
                    try:
                        self.daily_usage_tree.column(col, width=int(saved_widths[col]), stretch=False)
                    except: pass

    # KPI Update [NEW]
    # KPI Update [NEW]
    if hasattr(self, 'lbl_kpi_summary'):
        if filtered_df.empty:
            self.lbl_kpi_summary.config(text="조회된 데이터가 없습니다.")
        else:
            method_strs = []
            # Sort methods for consistent display (RT, UT, MT, PT, etc.)
            for m in sorted(method_totals.keys()):
                q = method_totals[m]
                if q > 0:
                    method_strs.append(f"{m} 합계: {q:,.1f}")
            method_text = "  |  ".join(method_strs) if method_strs else "검사물량 없음"
            
            kpi_text = (f"총 레코드 수: {len(filtered_df):,.0f}건  |  "
                        f"총 작업시간: {total_work_hours:,.1f} 시간  |  "
                        f"총 OT 시간: {total_ot_hours:,.1f} 시간  |  "
                        f"{method_text}  |  "
                        f"총 주행거리: {total_mileage:,.0f} km")
            self.lbl_kpi_summary.config(text=kpi_text)

    self.update_recent_entries_view()


