from site_apps.gyeongin.src.utils.helpers import NAN_PATTERN, DOT_ZERO_PATTERN, MARKER_PATTERN
from site_apps.gyeongin.src.views.components import *
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
    
    btn_refresh = ttk.Button(action_row, text="재고 새로고침", command=self.reload_and_update_stock_view)
    btn_refresh.pack(side='left', padx=5)
    
    btn_alert = ttk.Button(action_row, text="재주문 필요 항목 보기", command=self.show_low_stock)
    btn_alert.pack(side='left', padx=5)
    
    btn_delete = ttk.Button(action_row, text="품목 삭제", command=self.delete_selected_material)
    btn_delete.pack(side='left', padx=5)
    
    btn_edit = ttk.Button(action_row, text="품목 수정", command=self.open_edit_material_dialog)
    btn_edit.pack(side='left', padx=5)
    
    from site_apps.gyeongin.src.views.stock_history import show_material_history
    ttk.Button(action_row, text="품목 이력", command=lambda: show_material_history(self)).pack(side='left', padx=5)

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
    
    ttk.Label(filter_frame, text="자재명:").grid(row=0, column=6, padx=2, pady=2, sticky='e')
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
    
    # [UX IMPROVEMENT] 장비 포함 전체 보기 체크박스 추가 (기본 활성화: 자재와 장비 함께 표시)
    self.chk_show_all_equipment_var = tk.BooleanVar(value=True)
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
    
    columns = ('No.', '회사코드', '관리품번', '자재명', '장비명', 'SN', '창고', '모델명', '규격', '품목군코드', '공급업체', '제조사', '제조국', '가격', '원가', '관리단위', '입고수량', '사용수량', '재고수량', '재고하한', '상태/위치')
    self.stock_tree = ttk.Treeview(tree_frame, columns=columns, show='headings', 
                                  yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    
    vsb.config(command=self.stock_tree.yview)
    hsb.config(command=self.stock_tree.xview)
    
    # Column configuration
    col_widths = [50, 70, 80, 160, 160, 90, 80, 120, 100, 80, 100, 100, 70, 70, 70, 65, 70, 70, 70, 70, 110]
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
    
    from site_apps.gyeongin.src.models.stock_calculation import build_stock_lookups, stock_breakdown
    stock_lookups = build_stock_lookups(self)

    # Calculate current stock
    stock_summary = []
    
    row_idx = 1 # [UX IMPROVEMENT] 순차 번호 (No.) 부여
    
    # Pre-setup tags for equipment status highlighting
    self.stock_tree.tag_configure('deployed', background='#FFF9C4') # Light Yellow for "Field"
    self.stock_tree.tag_configure('in_stock', background='') # Default
    
    for _, mat in self.materials_df.iterrows():
        mat_id = mat['MaterialID']
        str_mat_id = self.normalize_id(mat_id)
        
        spec = str(mat.get('규격', '')).strip()
        mat_name_raw = self.clean_nan(mat.get('품목명', mat.get('ǰ', '')))
        model_name_raw = self.clean_nan(mat.get('모델명', mat.get('𵨸', '')))
        combined_name = (mat_name_raw + " " + model_name_raw).replace(' ', '')
        
        # [REFINED] Check consumable status using combined name or item name
        is_consumable = self._is_consumable_material(combined_name or mat_name_raw, '')

        amounts = stock_breakdown(self, mat, stock_lookups)
        in_qty, out_qty = amounts['in_qty'], amounts['out_qty']
        daily_qty, ndt_usage = amounts['daily_qty'], amounts['ndt_usage']
        stored_qty = amounts['stored_qty']
        total_incoming, total_used = amounts['total_incoming'], amounts['total_used']
        current_stock = amounts['current_stock']

        # [CRITICAL FIX] 실제 소비/입출고 내역이 있는 품목은 휴면/비활성화 상태여도 반드시 표시
        has_activity = (in_qty > 0 or out_qty > 0 or daily_qty > 0 or ndt_usage > 0)

        # [UX IMPROVEMENT] '장비 포함 전체 보기' 체크 해제 상태일 때만 비소모성 자재 숨김
        show_all = getattr(self, 'chk_show_all_equipment_var', None) and self.chk_show_all_equipment_var.get()
        if not show_all:
            if not is_consumable or (spec == "자동등록" and not is_consumable):
                continue

        # 활동 내역이 없는 비활성화 품목만 건너뜀
        if mat.get('Active', 1) == 0 and not has_activity:
            continue
        
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
        
        # [NEW] 자재명 / 장비명 컬럼 분리 표기
        mat_col = mat_name_raw if is_consumable else "-"
        eq_col = "-" if is_consumable else mat_name_raw

        # Default row display
        stock_summary.append({
            'data': (
                str(row_idx),
                safe_get(mat.get('회사코드', ''), ''),
                safe_get(mat.get('관리품번', ''), ''),
                mat_col,
                eq_col,
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
    
    # [NEW] Sort by 자재명 (Index 3) then 장비명 (Index 4) then 모델명 (Index 7)
    stock_summary.sort(key=lambda x: (str(x['data'][3]), str(x['data'][4]), str(x['data'][7])))
    
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
        if filter_eq != "전체" and str(row[2]) != filter_eq: continue
        if filter_name != "전체" and str(row[3]) != filter_name and str(row[4]) != filter_name: continue
        if filter_sn != "전체" and str(row[5]) != filter_sn: continue
        if filter_model != "전체" and str(row[7]) != filter_model: continue
        if filter_class != "전체" and str(row[9]) != filter_class: continue
        if filter_mfr != "전체" and str(row[11]) != filter_mfr: continue
        
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


