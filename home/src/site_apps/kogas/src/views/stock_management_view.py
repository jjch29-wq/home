from site_apps.kogas.src.utils.helpers import NAN_PATTERN, DOT_ZERO_PATTERN, MARKER_PATTERN
from site_apps.kogas.src.views.components import *
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import math
import datetime
import json
from tkcalendar import DateEntry
import traceback

def setup_stock_tab_impl(self):
    # Control Frame (Vertical container for both rows)
    control_frame = ttk.Frame(self.tab_stock)
    control_frame.pack(fill='x', padx=5, pady=5)
    
    # Row 1: Action Buttons
    action_row = ttk.Frame(control_frame)
    action_row.pack(fill='x', side='top', pady=(0, 5))
    
    btn_refresh = ttk.Button(action_row, text="재고 새로고침", command=self.update_stock_view)
    btn_refresh.pack(side='left', padx=5)
    
    btn_alert = ttk.Button(action_row, text="재주문 필요 항목 보기", command=self.show_low_stock)
    btn_alert.pack(side='left', padx=5)
    
    btn_delete = ttk.Button(action_row, text="품목 삭제", command=self.delete_selected_material)
    btn_delete.pack(side='left', padx=5)
    
    btn_edit = ttk.Button(action_row, text="품목 수정", command=self.open_edit_material_dialog)
    btn_edit.pack(side='left', padx=5)
    
    btn_export = ttk.Button(action_row, text="엑셀 내보내기", command=self.export_stock_to_excel)
    btn_export.pack(side='left', padx=5)
    
    btn_select_all = ttk.Button(action_row, text="전체 선택", command=self.select_all_stock)
    btn_select_all.pack(side='left', padx=5)
    
    # [NEW] Popout Button
    btn_popout_stock = ttk.Button(action_row, text="🔍 팝업창으로 열기", command=self.open_detached_stock_view)
    btn_popout_stock.pack(side='right', padx=5)
    
    # Row 2: Search and Filter Frame
    filter_row = ttk.Frame(control_frame)
    filter_row.pack(fill='x', side='top')
    
    filter_frame = ttk.LabelFrame(filter_row, text="검색 필터")
    filter_frame.pack(fill='x', padx=5, pady=2)
    
    # Row 0 of Filter Frame (Grid)
    ttk.Label(filter_frame, text="회사:").grid(row=0, column=0, padx=2, pady=2, sticky='e')
    self.cb_filter_co = ttk.Combobox(filter_frame, width=15)
    self.cb_filter_co.grid(row=0, column=1, padx=2, pady=2)
    
    ttk.Label(filter_frame, text="분류:").grid(row=0, column=2, padx=2, pady=2, sticky='e')
    self.cb_filter_class = ttk.Combobox(filter_frame, width=15)
    self.cb_filter_class.grid(row=0, column=3, padx=2, pady=2)
    
    ttk.Label(filter_frame, text="제조사:").grid(row=0, column=4, padx=2, pady=2, sticky='e')
    self.cb_filter_mfr = ttk.Combobox(filter_frame, width=15)
    self.cb_filter_mfr.grid(row=0, column=5, padx=2, pady=2)
    
    ttk.Label(filter_frame, text="품목명:").grid(row=0, column=6, padx=2, pady=2, sticky='e')
    self.cb_filter_name = ttk.Combobox(filter_frame, width=25)
    self.cb_filter_name.grid(row=0, column=7, padx=2, pady=2)
    
    # Row 1 of Filter Frame
    ttk.Label(filter_frame, text="S/N:").grid(row=1, column=0, padx=2, pady=2, sticky='e')
    self.cb_filter_sn = ttk.Combobox(filter_frame, width=20)
    self.cb_filter_sn.grid(row=1, column=1, padx=2, pady=2)
    
    ttk.Label(filter_frame, text="모델명:").grid(row=1, column=2, padx=2, pady=2, sticky='e')
    self.cb_filter_model = ttk.Combobox(filter_frame, width=20)
    self.cb_filter_model.grid(row=1, column=3, padx=2, pady=2)
    
    ttk.Label(filter_frame, text="관리품번:").grid(row=1, column=4, padx=2, pady=2, sticky='e')
    self.cb_filter_eq = ttk.Combobox(filter_frame, width=20)
    self.cb_filter_eq.grid(row=1, column=5, padx=2, pady=2)
    
    # [UX IMPROVEMENT] 장비 포함 전체 보기 체크박스 추가
    self.chk_show_all_equipment_var = tk.BooleanVar(value=False)
    self.chk_show_all_equipment = ttk.Checkbutton(
        filter_frame, text="☑ 장비 포함 전체 보기", 
        variable=self.chk_show_all_equipment_var, 
        command=self.update_stock_view
    )
    self.chk_show_all_equipment.grid(row=0, column=8, columnspan=2, padx=15, pady=2, sticky='w')

    # [NEW] Bind actions to all stock filter comboboxes for consistency and focus handling
    stock_filters = [
        self.cb_filter_co, self.cb_filter_class, self.cb_filter_mfr,
        self.cb_filter_name, self.cb_filter_sn, self.cb_filter_model,
        self.cb_filter_eq
    ]

    def on_stock_filter_action(e):
        self.update_stock_view()
        # [NEW] Move focus to the next filter in sequence for "sideways" navigation
        try:
            current_idx = stock_filters.index(e.widget)
            if current_idx + 1 < len(stock_filters):
                stock_filters[current_idx + 1].focus_set()
            else:
                # Move to the General Search entry after all dropdowns
                self.search_entry.focus_set()
        except:
            # Fallback to result list if sequence fails
            self.stock_tree.focus_set()

    for combo in stock_filters:
        combo.bind("<Return>", on_stock_filter_action)
        combo.bind('<<ComboboxSelected>>', lambda e: self.update_stock_view())
    
    ttk.Label(filter_frame, text="검색어:").grid(row=1, column=6, padx=2, pady=2, sticky='e')
    self.search_var = tk.StringVar()
    self.search_var.trace_add('write', lambda *args: self.update_stock_view())
    self.search_entry = ttk.Entry(filter_frame, textvariable=self.search_var, width=20)
    self.search_entry.grid(row=1, column=7, padx=2, pady=2)
    
    # [NEW] Special handler for Search Entry to move to Results List
    def on_search_enter(e):
        self.update_stock_view()
        self.stock_tree.focus_set()
        if self.stock_tree.get_children():
            # Highlight first item for immediate keyboard control
            self.stock_tree.selection_set(self.stock_tree.get_children()[0])
    self.search_entry.bind("<Return>", on_search_enter)
    
    # Reset Filters Button
    btn_reset = ttk.Button(filter_frame, text="♻️ 필터 초기화", command=self.reset_stock_filters)
    btn_reset.grid(row=1, column=8, padx=10, pady=2)
    
    # Treeview for Stock with Scrollbars
    tree_frame = ttk.Frame(self.tab_stock)
    tree_frame.pack(expand=True, fill='both', padx=5, pady=5)
    
    # Scrollbars
    vsb = ttk.Scrollbar(tree_frame, orient="vertical")
    hsb = ttk.Scrollbar(tree_frame, orient="horizontal")
    
    columns = ('No.', '회사코드', '관리품번', '품목명', 'SN', '창고', '모델명', '규격', '품목군코드', '공급업체', '제조사', '제조국', '가격', '원가', '관리단위', '입고수량', '사용수량', '재고수량', '재고하한', '상태/위치')
    self.stock_tree = ttk.Treeview(tree_frame, columns=columns, show='headings', 
                                  yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    
    vsb.config(command=self.stock_tree.yview)
    hsb.config(command=self.stock_tree.xview)
    
    # Column configuration
    col_widths = [60, 80, 100, 180, 90, 90, 120, 120, 90, 120, 120, 80, 80, 80, 80, 80, 80, 80, 80, 130]
    for col, width in zip(columns, col_widths):
        self.stock_tree.heading(col, text=col, command=lambda _col=col: self.treeview_sort_column(self.stock_tree, _col, False))
        # Change stretch=True to stretch=False to allow fixed user-defined widths
        self.stock_tree.column(col, width=width, minwidth=50, stretch=False, anchor='center')
    
    # Bind double-click
    self.stock_tree.bind('<Double-1>', lambda e: self.open_edit_material_dialog())
    
    # Auto-save column widths when user resizes columns
    self.stock_tree.bind('<ButtonRelease-1>', lambda e: self.save_tab_config())
    self.enable_tree_column_drag(self.stock_tree, context_menu_handler=lambda e: self._show_generic_tree_heading_context_menu(e, self.stock_tree))
    
    # Grid layout
    self.stock_tree.grid(row=0, column=0, sticky='nsew')
    vsb.grid(row=0, column=1, sticky='ns')
    hsb.grid(row=1, column=0, sticky='ew')
    
    tree_frame.grid_rowconfigure(0, weight=1)
    tree_frame.grid_columnconfigure(0, weight=1)
    
    # Initial view update will be triggered after update_registration_combos 
    # to ensure filters are properly initialized to "전체" first.
    self.refresh_ui_for_list_change('daily_units')


def update_stock_view_impl(self):
    active_trees = []
    if hasattr(self, 'stock_tree') and self.stock_tree.winfo_exists():
        active_trees.append(self.stock_tree)
    if hasattr(self, 'detached_windows') and 'stock' in self.detached_windows:
        dt = self.detached_windows['stock'].get('tree')
        if dt and dt.winfo_exists():
            active_trees.append(dt)
            
    if not active_trees:
        return
        
    # Clear current views
    for tree in active_trees:
        for item in tree.get_children():
            tree.delete(item)
    
    search_term = self.search_var.get().lower() if hasattr(self, 'search_var') else ''
    
    # Helper function to safely get value and replace NaN
    def safe_get(val, default=''):
        cleaned = self.clean_nan(val)
        if cleaned == '' and str(val).lower() in ['nan', 'none', 'null', 'nan.0', '0.0', '-0.0', '']:
            return default
        return val if pd.notna(val) else default
    
    def to_f(val):
        if pd.isna(val) or val is None: return 0.0
        try:
            s = str(val).replace(',', '').strip()
            if not s: return 0.0
            return float(s)
        except:
            return 0.0
    
    # [OPTIMIZATION] Pre-calculate stock lookup to avoid O(N*M) performance hit
    stock_in_lookup = {}
    stock_out_lookup = {}
    if not self.transactions_df.empty:
        temp_trans = self.transactions_df.copy()
        # Use standardized normalization
        temp_trans['NormID'] = temp_trans['MaterialID'].apply(self.normalize_id)
        # [CRITICAL] Exclude "자동 차감" (Automatic Deduction) to avoid double-counting with Daily Usage sheet
        mask = ~temp_trans['Note'].astype(str).str.contains('자동 차감', na=False)
        
        in_mask = mask & (temp_trans['Type'] == 'IN')
        out_mask = mask & (temp_trans['Type'] == 'OUT')
        
        # Group by normalized ID and sum quantities
        stock_in_lookup = temp_trans[in_mask].groupby('NormID')['Quantity'].sum().to_dict()
        stock_out_lookup = temp_trans[out_mask].groupby('NormID')['Quantity'].sum().to_dict()

    # [NEW] Pre-calculate Daily Usage subtraction
    daily_usage_lookup = {}
    daily_name_lookup = {}
    
    # [NEW] Get a set of consumable MaterialIDs from master for robust deduction
    consumable_ids = set()
    if not self.materials_df.empty:
        for _, m in self.materials_df.iterrows():
            # Check if the master item itself is a consumable
            m_name = str(m.get('품목명', '')).strip()
            if self._is_consumable_material(m_name, ''):
                c_id = self.normalize_id(m.get('MaterialID'))
                if c_id: consumable_ids.add(c_id)
    
    if hasattr(self, 'daily_usage_df') and not self.daily_usage_df.empty:
        temp_daily = self.daily_usage_df.copy()
        temp_daily['NormID'] = temp_daily['MaterialID'].apply(self.normalize_id)
        
        def _f(v):
            if pd.isna(v) or v is None: return 0.0
            try: return float(str(v).replace(',', '').strip()) if str(v).strip() else 0.0
            except: return 0.0
            
        # [ROBUST] Support multiple column names for usage and film counts
        temp_daily['TotalUsage'] = temp_daily.apply(lambda r: 
            _f(r.get('Usage', r.get('검사량', r.get('수량', r.get('Quantity', 0))))) + 
            _f(r.get('FilmCount', r.get('매수', 0))), axis=1)
        
        # [NEW] Filter: Only include consumables in the site-usage deduction lookup
        # This ensures durable equipment (PAUT, MT-Yoke) reported in the site tab doesn't decrease stock
        # [ROBUST] Use MaterialID as primary signal, fallback to 품목명/장비명 keywords
        temp_daily['IsConsumable'] = temp_daily.apply(lambda r: 
            (r['NormID'] in consumable_ids) or 
            self._is_consumable_material(
                str(r.get('품목명', r.get('장비명', ''))).strip(), 
                str(r.get('검사방법', '')).strip()
            ), axis=1)
        
        temp_consumable = temp_daily[temp_daily['IsConsumable']]
        
        daily_usage_lookup = temp_consumable.groupby('NormID')['TotalUsage'].sum().to_dict()
        
        # [NEW] Also build a name-based lookup for fallback
        temp_consumable['NormName'] = temp_consumable.apply(lambda r: str(r.get('품목명', r.get('장비명', ''))).strip(), axis=1)
        daily_name_lookup = temp_consumable.groupby('NormName')['TotalUsage'].sum().to_dict()
        print(f"DEBUG: Daily Consumable Lookup built: {daily_name_lookup}")
        
        # [NEW] Pre-calculate NDT chemical totals by name (Enhanced with fuzzy matching and JSON parsing)
        ndt_name_lookup = {}
        ndt_keys = ['세척제', '침투제', '현상제', '백색페인트', '흑색자분', '형광자분', '형광침투제', '자분페인트']
        for k in ndt_keys: ndt_name_lookup[k] = 0.0

        for _, row in temp_daily.iterrows():
            # 1. Check direct columns (fuzzy match)
            for col in temp_daily.columns:
                for k in ndt_keys:
                    if k in str(col):
                        ndt_name_lookup[k] += _f(row.get(col, 0))
            
            # 2. Check ndt_data JSON field
            nj_raw = row.get('ndt_data', row.get('ndtdata', ''))
            if nj_raw and isinstance(nj_raw, str) and nj_raw.strip().startswith('{'):
                try:
                    import json
                    ndt_json = json.loads(nj_raw)
                    for k in ndt_keys:
                        # Search in JSON keys
                        for jk, jv in ndt_json.items():
                            if k in str(jk):
                                ndt_name_lookup[k] += _f(jv)
                except: pass
            elif isinstance(nj_raw, dict):
                for k in ndt_keys:
                    for jk, jv in nj_raw.items():
                        if k in str(jk):
                            ndt_name_lookup[k] += _f(jv)
        
        # Filter out zero values for cleaner lookup
        ndt_name_lookup = {k: v for k, v in ndt_name_lookup.items() if v > 0}
        print(f"DEBUG: Daily Usage lookup built. NDT totals: {ndt_name_lookup}")

    # Calculate current stock
    stock_summary = []
    
    row_idx = 1 # [UX IMPROVEMENT] 순차 번호 (No.) 부여
    
    # Pre-setup tags for equipment status highlighting
    self.stock_tree.tag_configure('deployed', background='#FFF9C4') # Light Yellow for "Field"
    self.stock_tree.tag_configure('in_stock', background='') # Default
    
    for _, mat in self.materials_df.iterrows():
        if mat.get('Active', 1) == 0:
            continue
        
        mat_id = mat['MaterialID']

        # [REFINED] Skip non-consumables OR auto-registered items that are NOT consumables.
        # We ALLOW auto-registered items if they are confirmed as consumables (like films/drugs).
        spec = str(mat.get('규격', '')).strip()
        mat_name_str = str(mat.get('품목명', mat.get('ǰ', ''))).strip()
        is_consumable = self._is_consumable_material(mat_name_str, '')
        
        # [UX IMPROVEMENT] '장비 포함 전체 보기' 체크 해제 상태일 때만 비소모성 자재 숨김
        if not getattr(self, 'chk_show_all_equipment_var', None) or not self.chk_show_all_equipment_var.get():
            if not is_consumable or (spec == "자동등록" and not is_consumable):
                continue

        str_mat_id = self.normalize_id(mat_id)
        
        # Use optimized lookup
        in_qty = stock_in_lookup.get(str_mat_id, 0.0)
        out_qty = abs(stock_out_lookup.get(str_mat_id, 0.0))
        daily_qty = daily_usage_lookup.get(str_mat_id, 0.0)
        
        mat_name_str = str(mat.get('품목명', mat.get('ǰ', ''))).strip()
        # [NEW] Fallback to name-based lookup if ID lookup is 0
        if daily_qty == 0 and 'daily_name_lookup' in locals():
            daily_qty = daily_name_lookup.get(mat_name_str, 0.0)
        
        # Get stored quantity
        val = mat.get('수량', 0)
        try: stored_qty = float(str(val).replace(',', '')) if pd.notna(val) else 0.0
        except: stored_qty = 0.0
        
        # [FINAL_STOCK_CALC] Current Stock = Master + In/Out - Site Usage
        ndt_usage = 0.0
        mat_name_raw = self.clean_nan(mat.get('품목명', mat.get('ǰ', '')))
        model_name_raw = self.clean_nan(mat.get('모델명', mat.get('𵨸', '')))
        
        # Combine Item Name + Model Name for robust keyword searching (e.g. 'PT약품 세척제')
        combined_name = (mat_name_raw + " " + model_name_raw).replace(' ', '')
        
        # If this is an NDT item, find matching usage by name
        if hasattr(self, 'ndt_name_lookup') or 'ndt_name_lookup' in locals():
            lookup_source = ndt_name_lookup if 'ndt_name_lookup' in locals() else self.ndt_name_lookup
            
            # [REFINED] Only match specific keywords from combined name.
            for ndt_key, ndt_val in lookup_source.items():
                if ndt_key in combined_name:
                    ndt_usage = ndt_val
                    # If we found a specific NDT chemical (e.g. via model name), 
                    # ensure we don't also subtract generic inspection quantity
                    daily_qty = 0.0 
                    break
        
        total_incoming = stored_qty + in_qty
        total_used = out_qty + daily_qty + ndt_usage
        current_stock = total_incoming - total_used
        
        # Debug for specific items the user is watching (Films and NDT drugs)
        if any(k in combined_name.upper() for k in ['세척제', '침투제', '현상제', '백색', '흑색', '자분', 'CARESTREAM', 'MX125']):
            raw_val = mat.get('수량', 'MISSING')
            print(f"DEBUG: Stock Calc for '{mat_name_raw} ({model_name_raw})': RawQty='{raw_val}', Master={stored_qty}, In={in_qty}, Out={out_qty}, Daily={daily_qty}, NDT={ndt_usage}, Final={current_stock}")
        
        # --- Dynamic Location/Status Tracking ---
        status_location = "관내 (창고)"
        row_tag = 'in_stock'
        
        if current_stock <= 0:
            # Look for the last "OUT" transaction to determine where it went
            if not self.transactions_df.empty:
                # Optimized last transaction check
                mask = self.transactions_df['MaterialID'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True) == str_mat_id
                relevant_trans = self.transactions_df[mask].sort_values(by='Date', ascending=False)
                
                if not relevant_trans.empty:
                    last_op = relevant_trans.iloc[0]
                    if last_op['Type'] == 'OUT':
                        site_name = self.clean_nan(last_op.get('Site', ''))
                        status_location = f"현장: {site_name}" if site_name else "출고됨"
                        row_tag = 'deployed'
        
        # Default row display
        stock_summary.append({
            'data': (
                str(row_idx),
                safe_get(mat.get('회사코드', ''), ''),
                safe_get(mat.get('관리품번', ''), ''),
                safe_get(mat.get('품목명', ''), ''),
                safe_get(mat.get('SN', ''), ''),
                safe_get(mat.get('창고', ''), ''),
                safe_get(mat.get('모델명', ''), ''),
                safe_get(mat.get('규격', ''), ''),
                safe_get(mat.get('품목군코드', ''), ''),
                safe_get(mat.get('공급업체', ''), ''),
                safe_get(mat.get('제조사', ''), ''),
                safe_get(mat.get('제조국', ''), ''),
                f"{to_f(mat.get('가격', 0)):,.0f}",
                f"{to_f(mat.get('원가', 0)):,.0f}",
                safe_get(mat.get('관리단위', 'EA'), 'EA'),
                f"{to_f(total_incoming):g}",
                f"{to_f(total_used):g}",
                f"{to_f(current_stock):g}",
                f"{to_f(mat.get('재고하한', 0)):g}",
                status_location,
                mat_id # [UX IMPROVEMENT] Hidden real MaterialID at values[-1]
            ),
            'tag': row_tag
        })
        row_idx += 1
    
    # [NEW] Sort by Item Name (Index 3) then Model Name (Index 6)
    stock_summary.sort(key=lambda x: (str(x['data'][3]), str(x['data'][6])))
    
    # Filter by search term and dropdowns
    filter_co = str(self.cb_filter_co.get()).strip() if self.cb_filter_co.get() else "전체"
    filter_class = str(self.cb_filter_class.get()).strip() if self.cb_filter_class.get() else "전체"
    filter_mfr = str(self.cb_filter_mfr.get()).strip() if self.cb_filter_mfr.get() else "전체"
    filter_name = str(self.cb_filter_name.get()).strip() if self.cb_filter_name.get() else "전체"
    filter_sn = str(self.cb_filter_sn.get()).strip() if self.cb_filter_sn.get() else "전체"
    filter_model = str(self.cb_filter_model.get()).strip() if self.cb_filter_model.get() else "전체"
    filter_eq = str(self.cb_filter_eq.get()).strip() if self.cb_filter_eq.get() else "전체"
    
    final_row_idx = 1
    for row_obj in stock_summary:
        row = row_obj['data']
        # Dropdown Filters
        if filter_co != "전체" and str(row[1]) != filter_co: continue
        if filter_class != "전체" and str(row[8]) != filter_class: continue
        if filter_mfr != "전체" and str(row[10]) != filter_mfr: continue
        if filter_name != "전체" and str(row[3]) != filter_name: continue
        if filter_sn != "전체" and str(row[4]) != filter_sn: continue
        if filter_model != "전체" and str(row[6]) != filter_model: continue
        if filter_eq != "전체" and str(row[2]) != filter_eq: continue
        
        # General Search Term
        if search_term:
            row_str = ' '.join(str(x).lower() for x in row)
            if search_term not in row_str:
                continue
        
        # [UX IMPROVEMENT] 튜플을 리스트로 변환 후 화면에 보이는 순서대로 순차 번호 부여
        final_row = list(row)
        final_row[0] = str(final_row_idx)
        
        for tree in active_trees:
            tree.insert('', tk.END, values=final_row, tags=(row_obj['tag'],))
        final_row_idx += 1


