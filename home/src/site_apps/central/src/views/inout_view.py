# -*- coding: utf-8 -*-
"""
중앙지사 자재관리 입출고 관리(Inout) 탭 화면 및 트랜잭션 관리 Mixin
"""

import json
import logging
import os
import re
import sys
import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd

logger = logging.getLogger(__name__)

class InoutViewMixin:
    """입출고 탭 UI 배치, 자재 등록, 입출고 내역 추가/수정/삭제 관리 Mixin"""
    def setup_inout_tab(self):
        # Main PanedWindow (Vertical): Top (Registration) vs Bottom (History)
        self.inout_paned = ttk.Panedwindow(self.tab_inout, orient='vertical')
        self.inout_paned.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Horizontal PanedWindow for Registration Sections (Side-by-Side)
        self.inout_top_paned = ttk.Panedwindow(self.inout_paned, orient='horizontal')
        self.inout_paned.add(self.inout_top_paned, weight=1)
        
        # Save sash position on adjustment
        self.inout_paned.bind("<ButtonRelease-1>", lambda e: self.save_tab_config())
        
        # Side 1: New Material Registration
        reg_container = ttk.Frame(self.inout_top_paned)
        self.inout_top_paned.add(reg_container, weight=1)
        
        # Frame for Registration (Directly packed)
        reg_frame = ttk.LabelFrame(reg_container, text="📦 자재 / 장비 신규 등록")
        reg_frame.pack(fill='both', expand=True, padx=10, pady=5)
        reg_frame.grid_columnconfigure(1, weight=1)
        reg_frame.grid_columnconfigure(3, weight=1)
        
        # Row 0
        ttk.Label(reg_frame, text="회사코드:").grid(row=0, column=0, padx=5, pady=2, sticky='w')
        self.cb_co_code = ttk.Combobox(reg_frame, width=15)
        self.cb_co_code.grid(row=0, column=1, padx=5, pady=2, sticky='ew')
        self.cb_co_code.bind('<Return>', lambda e: self.cb_eq_code.focus_set())
        
        ttk.Label(reg_frame, text="설비코드:").grid(row=0, column=2, padx=5, pady=2, sticky='w')
        self.cb_eq_code = ttk.Combobox(reg_frame, width=20)
        self.cb_eq_code.grid(row=0, column=3, padx=5, pady=2, sticky='ew')
        self.cb_eq_code.bind('<Return>', lambda e: self.cb_item_name.focus_set())
        
        # Row 1: 자재명 / 장비명 (현재 재고 현황 탭과 통일)
        ttk.Label(reg_frame, text="자재명:").grid(row=1, column=0, padx=5, pady=2, sticky='w')
        item_name_container = ttk.Frame(reg_frame)
        item_name_container.grid(row=1, column=1, padx=5, pady=2, sticky='ew')
        
        self.cb_item_name = ttk.Combobox(item_name_container, width=15)
        self.cb_item_name.pack(side='left', fill='x', expand=True)
        self.cb_item_name.bind('<Return>', lambda e: self.cb_eq_name.focus_set())
        self.cb_item_name.bind('<<ComboboxSelected>>', lambda e: self.cb_eq_name.set('') if hasattr(self, 'cb_eq_name') else None)
        
        # [MODERN] Place the search button INSIDE the combobox
        btn_pref_search = tk.Button(self.cb_item_name, text="🔍", font=('Arial', 8), 
                                    bd=0, bg='white', cursor='hand2',
                                    command=self.open_preferred_item_search_dialog)
        # Offset x=-20 to avoid overlapping the combobox arrow
        btn_pref_search.place(relx=1.0, x=-20, rely=0.5, anchor='e', width=18, height=18)
        
        ttk.Label(reg_frame, text="장비명:").grid(row=1, column=2, padx=5, pady=2, sticky='w')
        self.cb_eq_name = ttk.Combobox(reg_frame, width=20)
        self.cb_eq_name.grid(row=1, column=3, padx=5, pady=2, sticky='ew')
        self.cb_eq_name.bind('<Return>', lambda e: self.cb_model.focus_set())
        self.cb_eq_name.bind('<<ComboboxSelected>>', lambda e: self.cb_item_name.set(''))
        
        # Row 2
        ttk.Label(reg_frame, text="모델명:").grid(row=2, column=0, padx=5, pady=2, sticky='w')
        self.cb_model = ttk.Combobox(reg_frame, width=15)
        self.cb_model.grid(row=2, column=1, padx=5, pady=2, sticky='ew')
        self.cb_model.bind('<Return>', lambda e: self.ent_sn.focus_set())
        
        ttk.Label(reg_frame, text="SN번호:").grid(row=2, column=2, padx=5, pady=2, sticky='w')
        self.ent_sn = ttk.Entry(reg_frame, width=20)
        self.ent_sn.grid(row=2, column=3, padx=5, pady=2, sticky='ew')
        self.ent_sn.bind('<Return>', lambda e: self.cb_class.focus_set())
        
        # Row 3
        ttk.Label(reg_frame, text="분류:").grid(row=3, column=0, padx=5, pady=2, sticky='w')
        self.cb_class = ttk.Combobox(reg_frame, width=15)
        self.cb_class.grid(row=3, column=1, padx=5, pady=2, sticky='ew')
        self.cb_class.bind('<Return>', lambda e: self.cb_spec.focus_set())
        
        ttk.Label(reg_frame, text="규격:").grid(row=3, column=2, padx=5, pady=2, sticky='w')
        self.cb_spec = ttk.Combobox(reg_frame, width=20)
        self.cb_spec.grid(row=3, column=3, padx=5, pady=2, sticky='ew')
        self.cb_spec.bind('<Return>', lambda e: self.cb_unit.focus_set())
        
        # Row 4
        ttk.Label(reg_frame, text="단위:").grid(row=4, column=0, padx=5, pady=2, sticky='w')
        self.cb_unit = ttk.Combobox(reg_frame, width=15)
        self.cb_unit.grid(row=4, column=1, padx=5, pady=2, sticky='ew')
        self.cb_unit.bind('<Return>', lambda e: self.cb_supplier.focus_set())
        
        ttk.Label(reg_frame, text="공급업자:").grid(row=4, column=2, padx=5, pady=2, sticky='w')
        self.cb_supplier = ttk.Combobox(reg_frame, width=20)
        self.cb_supplier.grid(row=4, column=3, padx=5, pady=2, sticky='ew')
        self.cb_supplier.bind('<Return>', lambda e: self.cb_mfr.focus_set())
        
        # Row 5
        ttk.Label(reg_frame, text="제조사:").grid(row=5, column=0, padx=5, pady=2, sticky='w')
        self.cb_mfr = ttk.Combobox(reg_frame, width=15)
        self.cb_mfr.grid(row=5, column=1, padx=5, pady=2, sticky='ew')
        self.cb_mfr.bind('<Return>', lambda e: self.cb_origin.focus_set())
        
        ttk.Label(reg_frame, text="제조국:").grid(row=5, column=2, padx=5, pady=2, sticky='w')
        self.cb_origin = ttk.Combobox(reg_frame, width=20)
        self.cb_origin.grid(row=5, column=3, padx=5, pady=2, sticky='ew')
        self.cb_origin.bind('<Return>', lambda e: self.ent_reorder.focus_set())
        
        # Row 6
        ttk.Label(reg_frame, text="초기재고:").grid(row=6, column=0, padx=5, pady=2, sticky='w')
        self.ent_init = ttk.Entry(reg_frame, width=15)
        self.ent_init.grid(row=6, column=1, padx=5, pady=2, sticky='ew')
        self.ent_init.insert(0, "0")
        self.ent_init.bind('<Return>', lambda e: self.ent_reorder.focus_set())
        
        ttk.Label(reg_frame, text="재주문 수준:").grid(row=6, column=2, padx=5, pady=2, sticky='w')
        self.ent_reorder = ttk.Entry(reg_frame, width=20)
        self.ent_reorder.grid(row=6, column=3, padx=5, pady=2, sticky='ew')
        self.ent_reorder.insert(0, "0")
        self.ent_reorder.bind('<Return>', lambda e: self.ent_price.focus_set())
        
        # Row 7
        ttk.Label(reg_frame, text="단가(가격):").grid(row=7, column=0, padx=5, pady=2, sticky='w')
        self.ent_price = ttk.Entry(reg_frame, width=15)
        self.ent_price.grid(row=7, column=1, padx=5, pady=2, sticky='ew')
        self.ent_price.insert(0, "0")
        self.ent_price.bind('<Return>', lambda e: self.ent_cost.focus_set())
        
        ttk.Label(reg_frame, text="원가:").grid(row=7, column=2, padx=5, pady=2, sticky='w')
        self.ent_cost = ttk.Entry(reg_frame, width=20)
        self.ent_cost.grid(row=7, column=3, padx=5, pady=2, sticky='ew')
        self.ent_cost.insert(0, "0")
        self.ent_cost.bind('<Return>', lambda e: btn_reg.focus_set())
        
        # Row 8
        btn_reg = ttk.Button(reg_frame, text="➕ 자재 / 장비 등록", command=self.register_material, style="Primary.TButton")
        btn_reg.grid(row=8, column=0, columnspan=4, pady=10)
        

        
        # Side 2: Transaction Entry
        trans_reg_container = ttk.Frame(self.inout_top_paned)
        self.inout_top_paned.add(trans_reg_container, weight=1)
        
        # Frame for In/Out Transaction (Directly packed)
        trans_frame = ttk.LabelFrame(trans_reg_container, text="🚚 자재 / 장비 입출고 기록")
        trans_frame.pack(fill='both', expand=True, padx=10, pady=5)
        trans_frame.grid_columnconfigure(1, weight=1)
        trans_frame.grid_columnconfigure(3, weight=1)

        # Bottom frame for history (Occupies full width at bottom)
        history_container = ttk.Frame(self.inout_paned)
        self.inout_paned.add(history_container, weight=2)
        
        ttk.Label(trans_frame, text="자재/장비 선택:").grid(row=0, column=0, padx=5, pady=2, sticky='w')
        mat_sel_container = ttk.Frame(trans_frame)
        mat_sel_container.grid(row=0, column=1, padx=5, pady=2, columnspan=3, sticky='ew')
        
        self.cb_material = ttk.Combobox(mat_sel_container, width=45, font=('Malgun Gothic', 10))
        self.cb_material.pack(side='left', fill='x', expand=True)
        self.cb_material.bind('<<ComboboxSelected>>', self.on_material_selected)
        self.cb_material.bind('<Return>', lambda e: self.cb_type.focus_set())
        
        # [MODERN] Place the search button INSIDE the combobox
        btn_mat_search = tk.Button(self.cb_material, text="🔍", font=('Arial', 8), 
                                    bd=0, bg='white', cursor='hand2',
                                    command=self.open_material_search_dialog)
        # Offset x=-20 to avoid overlapping the combobox arrow
        btn_mat_search.place(relx=1.0, x=-20, rely=0.5, anchor='e', width=18, height=18)
        
        # Register autocomplete for material selection
        self.enable_autocomplete(self.cb_material, values_list=[]) # Initial empty, will be updated
        self.update_material_combo()
        
        # Model List display - Moved to row 4 for narrower fit
        ttk.Label(trans_frame, text="관련 모델명:").grid(row=4, column=0, padx=5, pady=2, sticky='nw')
        self.list_models = tk.Listbox(trans_frame, height=3, width=42, font=('Malgun Gothic', 9))
        self.list_models.grid(row=4, column=1, columnspan=3, padx=5, pady=2, sticky='ew')
        
        # Add scrollbar for model list
        model_vsb = ttk.Scrollbar(trans_frame, orient="vertical", command=self.list_models.yview)
        model_vsb.grid(row=4, column=4, sticky='ns', pady=2)
        self.list_models.config(yscrollcommand=model_vsb.set)
        
        ttk.Label(trans_frame, text="구분:").grid(row=1, column=0, padx=5, pady=2, sticky='w')
        self.cb_type = ttk.Combobox(trans_frame, values=["IN", "OUT"], state="readonly", width=15)
        self.cb_type.grid(row=1, column=1, padx=5, pady=2, sticky='ew')
        self.cb_type.set("OUT")
        self.cb_type.bind('<Return>', lambda e: self.ent_qty.focus_set())
        
        ttk.Label(trans_frame, text="수량:").grid(row=1, column=2, padx=5, pady=2, sticky='w')
        self.ent_qty = ttk.Entry(trans_frame, width=30)
        self.ent_qty.grid(row=1, column=3, padx=5, pady=2, sticky='ew')
        self.ent_qty.bind('<Return>', lambda e: self.cb_trans_site.focus_set())
        
        ttk.Label(trans_frame, text="현장:").grid(row=2, column=0, padx=5, pady=2, sticky='w')
        self.cb_trans_site = ttk.Combobox(trans_frame, width=28, values=self.sites)
        self.cb_trans_site.grid(row=2, column=1, padx=5, pady=2, sticky='ew')
        self.cb_trans_site.bind('<FocusOut>', lambda e: self.auto_save_to_list(e, self.cb_trans_site, self.sites, 'sites'))
        self.cb_trans_site.bind('<Return>', lambda e: self._on_trans_site_return(e))
        
        ttk.Label(trans_frame, text="창고:").grid(row=2, column=2, padx=5, pady=2, sticky='w')
        self.cb_warehouse = ttk.Combobox(trans_frame, width=28, values=self.warehouses)
        self.cb_warehouse.grid(row=2, column=3, padx=5, pady=2, sticky='ew')
        self.cb_warehouse.bind('<FocusOut>', lambda e: self.auto_save_to_list(e, self.cb_warehouse, self.warehouses, 'warehouses'))
        self.cb_warehouse.bind('<Return>', lambda e: self._on_warehouse_return(e))
        
        ttk.Label(trans_frame, text="담당자:").grid(row=3, column=0, padx=5, pady=2, sticky='w')
        self.ent_user = ttk.Combobox(trans_frame, width=28, values=getattr(self, 'users', []))
        self.ent_user.grid(row=3, column=1, padx=5, pady=2, sticky='ew')
        self.ent_user.bind('<FocusOut>', lambda e: self.auto_save_to_list(e, self.ent_user, self.users, 'users'))
        self.ent_user.bind('<Return>', lambda e: self._on_user_return(e))
        
        ttk.Label(trans_frame, text="비고:").grid(row=3, column=2, padx=5, pady=2, sticky='w')
        self.ent_note = ttk.Entry(trans_frame, width=30)
        self.ent_note.grid(row=3, column=3, padx=5, pady=2, sticky='ew')
        self.ent_note.bind('<Return>', lambda e: btn_trans.focus_set())
        
        btn_trans = ttk.Button(trans_frame, text="💾 기록 저장", command=self.add_transaction, style="Action.TButton")
        btn_trans.grid(row=5, column=0, columnspan=4, pady=10)
        
        # Frame for displaying transaction history
        history_frame = ttk.LabelFrame(history_container, text="📜 최근 입출고 내역")
        history_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Filter/Control frame for history
        history_ctrl_frame = ttk.Frame(history_frame)
        history_ctrl_frame.pack(fill='x', padx=5, pady=2)
        
        # First row of filters
        filter_row1 = ttk.Frame(history_ctrl_frame)
        filter_row1.pack(fill='x', pady=2)
        
        btn_del_trans = ttk.Button(filter_row1, text="🗑️ 선택 삭제", command=self.delete_transaction_entry, style="Danger.TButton")
        btn_del_trans.pack(side='left', padx=3)
        
        btn_export_trans = ttk.Button(filter_row1, text="📊 엑셀 내보내기", command=self.export_transaction_history, style="Action.TButton")
        btn_export_trans.pack(side='left', padx=3)

        btn_refresh_trans = ttk.Button(filter_row1, text="🔄 내역 새로고침", command=self.refresh_inout_history, style="Primary.TButton")
        btn_refresh_trans.pack(side='left', padx=3)

        # [NEW] Cleanup button for automatic transactions (Equipment only)
        btn_cleanup_trans = ttk.Button(filter_row1, text="🧹 자동기록 정리", command=self.cleanup_auto_transactions, style="Warning.TButton")
        btn_cleanup_trans.pack(side='left', padx=6)
        
        # Second row of filters
        filter_row2 = ttk.Frame(history_ctrl_frame)
        filter_row2.pack(fill='x', pady=2)
        
        ttk.Label(filter_row2, text="자재/장비명 필터:").pack(side='left', padx=5)
        self.cb_trans_filter_mat = ttk.Combobox(filter_row2, width=35, state="readonly")
        self.cb_trans_filter_mat.pack(side='left', padx=5)
        self.cb_trans_filter_mat.bind('<<ComboboxSelected>>', lambda e: self.update_transaction_view())
        
        ttk.Label(filter_row2, text="현장 필터:").pack(side='left', padx=(20, 5))
        self.cb_trans_filter_site = ttk.Combobox(filter_row2, width=12, state="readonly")
        self.cb_trans_filter_site.pack(side='left', padx=5)
        self.cb_trans_filter_site.bind('<<ComboboxSelected>>', lambda e: self.update_transaction_view())
        
        # Third row of filters
        filter_row3 = ttk.Frame(history_ctrl_frame)
        filter_row3.pack(fill='x', pady=2)
        
        ttk.Label(filter_row3, text="차량번호 필터:").pack(side='left', padx=5)
        self.cb_trans_filter_vehicle = ttk.Combobox(filter_row3, width=12, state="readonly")
        self.cb_trans_filter_vehicle.pack(side='left', padx=5)
        self.cb_trans_filter_vehicle.bind('<<ComboboxSelected>>', lambda e: self.update_transaction_view())
        self.cb_trans_filter_vehicle.bind('<F5>', lambda e: (self.refresh_inout_history(), 'break')[1])
        
        # Treeview for history
        tree_scroll_frame = ttk.Frame(history_frame)
        tree_scroll_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        inout_vsb = ttk.Scrollbar(tree_scroll_frame, orient="vertical")
        inout_hsb = ttk.Scrollbar(tree_scroll_frame, orient="horizontal")
        
        # [RESTRUCTURED] Optimized for NDT consumables & Equipment: Focus on RT, PT, MT quantities and Equipment
        columns = ('날짜', '현장', '구분', '품목명', '장비명', '수량', '세척제', '침투제', '현상제', '백색페인트', '흑색자분', 'SN', '규격', '창고', '담당자', '비고', '검사비')
        self.inout_tree = ttk.Treeview(tree_scroll_frame, columns=columns, show='headings', height=10,
                                       yscrollcommand=inout_vsb.set, xscrollcommand=inout_hsb.set)
        
        inout_vsb.config(command=self.inout_tree.yview)
        inout_hsb.config(command=self.inout_tree.xview)
        
        # [NEW] Default Column Widths for NDT-focused View
        col_widths = {
            '날짜': 130, '현장': 100, '구분': 60, '품목명': 150, '장비명': 150,
            '수량': 70, '세척제': 60, '침투제': 60, '현상제': 60, '백색페인트': 80, '흑색자분': 80,
            'SN': 100, '규격': 100, '창고': 80, '담당자': 80, '비고': 150, '검사비': 90
        }
        for col in columns:
            self.inout_tree.heading(col, text=("자재명" if col == "품목명" else col))
            width = col_widths.get(col, 100)
            self.inout_tree.column(col, width=width, minwidth=50, stretch=False, anchor='center')
        
        # [FIX] Select core columns including both 자재명 and 장비명 for initial display
        visible_cols = ('날짜', '현장', '구분', '품목명', '장비명', '수량', '세척제', '침투제', '현상제', '백색페인트', '흑색자분', '비고')
        self.inout_tree['displaycolumns'] = visible_cols

        def _show_inout_tree_context_menu(event):
            region = self.inout_tree.identify_region(event.x, event.y)
            if region == "heading":
                self._show_generic_tree_heading_context_menu(event, self.inout_tree)
                return
            item = self.inout_tree.identify_row(event.y)
            if item:
                self.inout_tree.selection_set(item)
                menu = tk.Menu(self.root, tearoff=0)
                menu.add_command(label="🔄 내역 새로고침", command=self.refresh_inout_history)
                menu.add_command(label="📊 엑셀 내보내기", command=self.export_transaction_history)
                menu.add_separator()
                menu.add_command(label="⚙️ 컬럼 관리(숨기기/보이기)...", command=self.show_column_visibility_dialog)
                try:
                    menu.tk_popup(event.x_root, event.y_root)
                finally:
                    menu.grab_release()

        self.enable_tree_column_drag(self.inout_tree, context_menu_handler=_show_inout_tree_context_menu)
        self.inout_tree.bind('<ButtonRelease-1>', lambda e: self.save_tab_config(), add='+')
        self.inout_tree.bind('<F5>', lambda e: (self.refresh_inout_history(), 'break')[1])

        self.inout_tree.grid(row=0, column=0, sticky='nsew')
        inout_vsb.grid(row=0, column=1, sticky='ns')
        inout_hsb.grid(row=1, column=0, sticky='ew')
        
        tree_scroll_frame.grid_rowconfigure(0, weight=1)
        tree_scroll_frame.grid_columnconfigure(0, weight=1)
        
        # Initial populate
        self.update_transaction_view()
        
        # Set initial sash position for better balance
        self.inout_paned.after(200, self._ensure_inout_sash_visibility)

    def _ensure_inout_sash_visibility(self):
        """Ensure the inout sash position is properly set"""
        try:
            if hasattr(self, 'inout_paned'):
                self.inout_paned.update_idletasks()
                total_h = self.inout_paned.winfo_height()
                if total_h > 200:
                    # Set vertical sash to 40% of total height (registration area smaller)
                    new_pos = int(total_h * 0.4)
                    self.inout_paned.sashpos(0, new_pos)
                    
                    # Set horizontal sash for side-by-side
                    total_w = self.inout_top_paned.winfo_width()
                    if total_w > 200:
                        self.inout_top_paned.sashpos(0, int(total_w * 0.45))
                    print(f"Set inout sash to {new_pos} (total height: {total_h})")
        except Exception as e:
            print(f"Error ensuring inout sash visibility: {e}")


    def get_material_display_name(self, mat_id):
        """Get formatted material name as '품목명 (SN: SN번호) - 규격'"""
        # [FIX] Handle NaN/empty/zero IDs gracefully to prevent showing '0.0' or matching empty strings
        if pd.isna(mat_id) or str(mat_id).lower().strip() in ('nan', 'none', '', '0', '0.0', '0.00', '-'):
            return ""
        try:
            if float(str(mat_id).replace(',', '').strip()) == 0.0:
                return ""
        except:
            pass
            
        if self.materials_df.empty:
            return f"ID: {mat_id}"
            
        target_id = self.normalize_id(mat_id)
        if not target_id:
            return ""

        # [PERF] Fast O(1) cache lookup
        if not hasattr(self, '_mat_display_cache'):
            self._mat_display_cache = {}
        if mat_id in self._mat_display_cache:
            return self._mat_display_cache[mat_id]
        if target_id in self._mat_display_cache:
            return self._mat_display_cache[target_id]
            
        mask = (self.materials_df['MaterialID'].notna()) & (self.materials_df['MaterialID'].apply(self.normalize_id) == target_id)
        mat_row = self.materials_df[mask]
            
        if mat_row.empty:
            # [NEW] Handle non-numeric IDs (raw names for PAUT) and orphans gracefully
            self._mat_display_cache[mat_id] = str(mat_id)
            return str(mat_id)
            
        mat = mat_row.iloc[0]
        name = mat.get('품목명', '')
        
        # [SAFETY] If 품목명 is NaN/empty, fall back to str(mat_id)
        if pd.isna(name) or str(name).strip() == '':
            self._mat_display_cache[mat_id] = str(mat_id)
            return str(mat_id)
        
        sn = mat.get('SN', '')
        spec = mat.get('규격', '')
        
        # Build display string: 품목명-SN (모델명) [규격]
        display = name
        if sn and pd.notna(sn) and str(sn).strip():
            display = f"{name}-{str(sn).strip()}"
            
        model_name = mat.get('모델명', '')
        model_name_val = str(model_name).strip()
        if model_name_val and pd.notna(model_name) and model_name_val != str(name).strip():
            display = f"{display} ({model_name_val})"
        
        if spec and pd.notna(spec) and str(spec).strip():
            display = f"{display} [{str(spec).strip()}]"
            
        self._mat_display_cache[mat_id] = display
        self._mat_display_cache[target_id] = display
        return display


    def get_material_name_only(self, mat_id):
        """Get only the 품목명 for a material"""
        if pd.isna(mat_id) or str(mat_id).lower().strip() in ('nan', 'none', '', '0'):
            return ""
        if self.materials_df.empty:
            return f"(미등록) {mat_id}"
        
        target_id = self.normalize_id(mat_id)
        if not target_id:
            return ""
        # Match using normalized IDs, strictly ignoring NaN MaterialIDs
        mask = (self.materials_df['MaterialID'].notna()) & (self.materials_df['MaterialID'].apply(self.normalize_id) == target_id)
        mat_row = self.materials_df[mask]
        
        if mat_row.empty:
            return f"(미등록) {mat_id}"
        name = mat_row.iloc[0].get('품목명', '')
        return "" if pd.isna(name) else str(name).strip()

    def get_material_info(self, mat_id):
        """Get all material columns as a dictionary"""
        if pd.isna(mat_id) or str(mat_id).lower().strip() in ('nan', 'none', '', '0'):
            return {}
        if self.materials_df.empty:
            return {}
        
        target_id = self.normalize_id(mat_id)
        if not target_id:
            return {}
        mask = (self.materials_df['MaterialID'].notna()) & (self.materials_df['MaterialID'].apply(self.normalize_id) == target_id)
        mat_row = self.materials_df[mask]
        
        if mat_row.empty:
            return {}
        return mat_row.iloc[0].to_dict()

    def update_material_combo(self):
        """Update both In/Out and Daily Usage material comboboxes with unified list"""
        mat_list = []
        inactive_ids = set()
        inactive_names = set()
        if not hasattr(self, '_mat_display_cache'):
            self._mat_display_cache = {}
            
        if not self.materials_df.empty:
            # [PERF] Fast single-pass vector iteration using itertuples
            active_vals = pd.to_numeric(self.materials_df.get('Active', 1), errors='coerce').fillna(1)
            for idx, row in enumerate(self.materials_df.itertuples()):
                is_active = active_vals.iloc[idx] > 0.5
                mat_id = getattr(row, 'MaterialID', None)
                norm_id = self.normalize_id(mat_id)
                name = getattr(row, '품목명', '')
                
                if not is_active:
                    if norm_id: inactive_ids.add(norm_id)
                    if name: inactive_names.add(str(name).strip())
                    continue
                
                if not name or str(name).strip() == '':
                    continue
                
                sn = str(getattr(row, 'SN', '') or '').strip()
                spec = str(getattr(row, '규격', '') or '').strip()
                model = str(getattr(row, '모델명', '') or '').strip()
                
                display = name
                if sn: display = f"{name}-{sn}"
                if model and model != str(name).strip(): display = f"{display} ({model})"
                if spec: display = f"{display} [{spec}]"
                
                mat_list.append(display)
                if mat_id is not None:
                    self._mat_display_cache[mat_id] = display
                if norm_id:
                    self._mat_display_cache[norm_id] = display

        # [NEW] Also collect historical material names from daily usage history (Only if still Active)
        history_materials = []
        if not self.daily_usage_df.empty and 'MaterialID' in self.daily_usage_df.columns:
            unique_mat_ids = self.daily_usage_df['MaterialID'].dropna().unique()
            for mat_id in unique_mat_ids:
                str_id = self.normalize_id(mat_id)
                if str_id in inactive_ids:
                    continue
                name = self.get_material_display_name(mat_id)
                if name: history_materials.append(name)
        
        # Merge database items with centralized films list and historical materials, unique and sort
        all_vals = list(set([str(m) for m in mat_list + self.carestream_films + history_materials if pd.notna(m) and str(m).strip()]))
        all_vals.sort()
        
        # Update ComboBoxes
        self.cached_material_list = all_vals
        if hasattr(self, 'cb_material'):
            self.cb_material['values'] = all_vals
        if hasattr(self, 'cb_daily_material'):
            if isinstance(self.cb_daily_material, ttk.Combobox):
                # [NEW] Filter: Only include actual consumables (films, chemicals) in the material dropdown
                consumable_only_vals = [v for v in all_vals if self._is_consumable_material(v.split('[')[0].split('-')[0].strip(), '')]
                self.cb_daily_material['values'] = consumable_only_vals
                try:
                    max_chars = max((len(str(v)) for v in consumable_only_vals), default=12)
                    self.cb_daily_material.configure(width=min(40, max(22, max_chars + 2)))
                except Exception:
                    pass
            
        if hasattr(self, 'cb_daily_equip'):
            # [NEW] Collect history of equipment names from daily usage records (Only if Active)
            history_equip = []
            if not self.daily_usage_df.empty and '장비명' in self.daily_usage_df.columns:
                raw_history = self.daily_usage_df['장비명'].dropna().unique().tolist()
                for eq in raw_history:
                    eq_str = str(eq).strip()
                    if not eq_str or eq_str in inactive_names:
                        continue
                    history_equip.append(eq_str)
                
            # [NEW] Filter: Only include actual equipment (non-consumables) in the equipment dropdown
            equip_only_vals = []
            for v in all_vals:
                base_name = v.split('[')[0].split('-')[0].strip()
                if not self._is_consumable_material(base_name, ''):
                    equip_only_vals.append(v)
                
            # Combine custom equipments, history, and the filtered equipment list
            combined_equip = list(set([str(e).strip() for e in self.equipments + history_equip + equip_only_vals if pd.notna(e) and str(e).strip()]))
            combined_equip.sort()
            
            # [FIX] Only set values if it's a Combobox
            if isinstance(self.cb_daily_equip, ttk.Combobox):
                self.cb_daily_equip['values'] = combined_equip
            
            try:
                max_chars = max((len(str(v)) for v in combined_equip), default=12)
                self.cb_daily_equip.configure(width=min(40, max(22, max_chars + 2)))
            except Exception:
                pass
            self.equipment_suggestions = combined_equip
        
        self.materials_display_list = all_vals
        
        # [NEW] Calculate consumable-only list for Inventory Status filters (Exclude PAUT, etc.)
        self.consumable_display_list = []
        for val in all_vals:
            # Extract the base name (before spec '-' or ID '[')
            base_name = val.split('[')[0].split('-')[0].strip()
            if self._is_consumable_material(base_name, ''):
                self.consumable_display_list.append(val)
        self.consumable_display_list.sort()


        if hasattr(self, 'cb_trans_filter_mat'):
            self.cb_trans_filter_mat['values'] = ["전체"] + all_vals
            if not self.cb_trans_filter_mat.get():
                self.cb_trans_filter_mat.set("전체")
        
        if hasattr(self, 'cb_trans_filter_site'):
            self.cb_trans_filter_site['values'] = ["전체"] + sorted(self.sites)
            if not self.cb_trans_filter_site.get():
                self.cb_trans_filter_site.set("전체")
        
        if hasattr(self, 'cb_trans_filter_vehicle'):
            # Get unique vehicle numbers from transactions_df
            if '차량번호' in self.transactions_df.columns:
                vehicle_numbers = self.transactions_df['차량번호'].dropna().astype(str).str.strip().unique()
                vehicle_list = sorted([v for v in vehicle_numbers if v and v != 'nan'])
            else:
                vehicle_list = []
            
            # Add sample vehicle numbers for testing if no data exists
            if not vehicle_list:
                vehicle_list = ['12가1234', '34나5678', '89다9012']
            
            self.cb_trans_filter_vehicle['values'] = ["전체"] + vehicle_list
            if not self.cb_trans_filter_vehicle.get():
                self.cb_trans_filter_vehicle.set("전체")

    def _get_material_candidates(self, include_all=False):
        from site_apps.central.src.method_catalog import current_method, material_candidates
        if current_method(self):
            result = material_candidates(self)
            return (['전체'] + result) if include_all else result
        vals = []
        if hasattr(self, 'materials_display_list'):
            vals.extend(getattr(self, 'materials_display_list', []) or [])
        if hasattr(self, 'cb_daily_material'):
            pass # Removed reading from widget to prevent recursive filtering issues

        cleaned = sorted(set(str(v).strip() for v in vals if str(v).strip() and str(v).strip() != '전체' and not str(v).strip().startswith('ID: ')))
        return (['전체'] + cleaned) if include_all else cleaned

    def _get_history_material_candidates(self, include_all=False):
        """Get material names from both current inventory and history for selection"""
        raw_materials = set()
        
        # 1. Add all active materials from stock master (Primary Source)
        if not self.materials_df.empty:
            for _, mat in self.materials_df.iterrows():
                if mat.get('Active', 1) == 1:
                    name = self.get_material_display_name(mat['MaterialID'])
                    if name: raw_materials.add(name)
        
        # 2. Add materials exclusively from daily usage history (Secondary Source)
        if not self.daily_usage_df.empty and 'MaterialID' in self.daily_usage_df.columns:
            unique_mat_ids = self.daily_usage_df['MaterialID'].dropna().unique()
            for mat_id in unique_mat_ids:
                name = self.get_material_display_name(mat_id)
                if name: raw_materials.add(name)
        
        cleaned = sorted(list(raw_materials))
        return (['전체'] + cleaned) if include_all else cleaned

    def _get_equipment_candidates(self, include_all=False):
        """Get unified equipment suggestions from master list and history"""
        raw_equip = set()
        
        # 1. Base equipment list
        if hasattr(self, 'equipments'):
            for e in getattr(self, 'equipments', []):
                if e and str(e).strip(): raw_equip.add(str(e).strip())
                
        # 2. Dynamic suggestions
        if hasattr(self, 'equipment_suggestions'):
            for e in getattr(self, 'equipment_suggestions', []):
                if e and str(e).strip(): raw_equip.add(str(e).strip())
                
        # 3. History
        if hasattr(self, 'daily_usage_df') and not self.daily_usage_df.empty and '장비명' in self.daily_usage_df.columns:
            unique_equip = self.daily_usage_df['장비명'].dropna().astype(str).str.strip().unique()
            for e in unique_equip:
                if e: raw_equip.add(e)

        if hasattr(self, 'daily_usage_df') and '보조장비' in self.daily_usage_df.columns:
            from site_apps.central.src.aux_equipment import names
            for text in self.daily_usage_df['보조장비'].dropna():
                raw_equip.update(names(text))
        cleaned = sorted([e for e in raw_equip if e != '전체'])
        return (['전체'] + cleaned) if include_all else cleaned

    def _get_inspection_item_candidates(self):
        """Get candidates for Inspection Item (Inspection Item Name)"""
        items = set(getattr(self, 'test_items', []))
        if not self.daily_usage_df.empty and '검사품명' in self.daily_usage_df.columns:
            hist = self.daily_usage_df['검사품명'].dropna().astype(str).str.strip().unique()
            for h in hist: 
                if h: items.add(h)
        return sorted(list(items))

    def _get_applied_code_candidates(self):
        """Get candidates for Applied Code"""
        codes = set(getattr(self, 'applied_codes', []))
        if not self.daily_usage_df.empty and '적용코드' in self.daily_usage_df.columns:
            hist = self.daily_usage_df['적용코드'].dropna().astype(str).str.strip().unique()
            for h in hist: 
                if h: codes.add(h)
        return sorted(list(codes))

    def _apply_combobox_word_suggest(self, combobox, source_values, open_dropdown=False):
        if not combobox or not isinstance(combobox, ttk.Combobox):
            return

        text_raw = combobox.get()
        text = text_raw.strip()
        had_focus = False
        try:
            # More robust focus check
            curr_focus = self.root.focus_get()
            had_focus = (curr_focus == combobox or str(curr_focus).startswith(str(combobox)))
        except Exception:
            had_focus = False
        
        try:
            cursor_pos = combobox.index(tk.INSERT)
        except Exception:
            cursor_pos = len(text_raw)
            
        source = [str(v).strip() for v in (source_values or []) if str(v).strip()]

        if not text:
            filtered = source
        else:
            q = text.lower()
            filtered = [v for v in source if q in v.lower()]

        # [FIX] Only update if values changed to avoid recursion/flicker
        values_changed = list(combobox['values']) != list(filtered)
        if values_changed:
            combobox['values'] = filtered

        # Always ensure state is reconciled if it had focus
        if had_focus:
            def _restore_state():
                try:
                    if self.root.focus_get() != combobox:
                        combobox.focus_set()
                    combobox.set(text_raw)
                    combobox.icursor(cursor_pos)
                    combobox.selection_clear()
                except Exception:
                    pass

            _restore_state()
            # Double tap with after_idle to ensure it sticks after the internal Tk events
            try: combobox.after_idle(_restore_state)
            except Exception: pass
        
        # [REVISION] Use 'Alt-Down' for more stable dropdown posting
        if open_dropdown and had_focus and filtered and len(text) > 0:
            try:
                # [STABILITY] Don't re-open if just selected
                if getattr(combobox, '_just_selected', False):
                    return
                # Alt-Down is the standard shortcut to post the dropdown in many OS/Toolkits
                combobox._auto_opening = True
                combobox.event_generate('<Alt-Down>')
                combobox._auto_opening = False
                # Re-verify cursor immediately after posting
                combobox.icursor(cursor_pos)
            except:
                pass

    def _bind_combobox_word_suggest(self, combobox, source_getter):
        if not combobox:
            return

        def _on_interaction(e=None, open=True):
            # Skip for navigation keys which have native behaviors
            if e is not None and hasattr(e, 'keysym'):
                if e.keysym in {
                    'Left', 'Right', 'Tab', 'Up', 'Down', 'Return', 'Escape',
                    'Shift_L', 'Shift_R', 'Control_L', 'Control_R',
                    'Alt_L', 'Alt_R', 'Caps_Lock'
                }:
                    return
                # [NEW] Skip 'Process' keysym used by Korean IME during composition
                if e.keysym == 'Process':
                    return
            
            # [FIX] Debounce: Wait 300ms before showing suggestions to avoid interrupting continuous typing (especially for IME)
            if hasattr(combobox, '_suggest_after_id'):
                combobox.after_cancel(combobox._suggest_after_id)
            
            def _do_suggest():
                # 방금 목록에서 선택한 직후에는 값 목록/포커스를 다시 만지지 않는다(선택 지연 방지).
                if combobox.winfo_exists() and not getattr(combobox, '_just_selected', False):
                    self._apply_combobox_word_suggest(combobox, source_getter(), open_dropdown=open)
            
            combobox._suggest_after_id = self.root.after(300, _do_suggest)

        combobox.bind('<KeyRelease>', lambda e: _on_interaction(e, open=True), add='+')
        # FocusIn에서 목록·포커스를 재설정하면 드롭다운을 여는 순간 값 목록이 바뀌어 선택이 지연될 수 있어
        # 타이핑(KeyRelease) 때만 제안 목록을 갱신한다.
        
        # [NEW] Guard against "sticky" dropdown after selection
        def _on_selected(e=None):
            combobox._just_selected = True
            def _reset(): combobox._just_selected = False
            combobox.after(500, _reset)
            
        combobox.bind('<<ComboboxSelected>>', _on_selected, add='+')
        
        # [FIX] Use postcommand instead of <ButtonPress-1> to restore the full list.
        # Modifying values during <ButtonPress-1> can break Tkinter dropdown selection (관경 선택 불가능 현상 수정).
        def _on_post():
            if not getattr(combobox, '_auto_opening', False):
                try:
                    new_vals = source_getter()
                    # Only update if changed to prevent breaking selection internally
                    if list(combobox['values']) != list(new_vals):
                        combobox['values'] = new_vals
                except: pass
        if isinstance(combobox, ttk.Combobox):
            try:
                combobox.config(postcommand=_on_post)
            except Exception: pass

    def update_registration_combos(self):
        """Update registration comboboxes with unique values from database and centralized list"""
        # 1. Update registration fields from database
        fields = {}
        for key, attr in [
            ('회사코드', 'cb_co_code'),
            ('관리품번', 'cb_eq_code'),
            ('품목명', 'cb_item_name'),
            ('품목군코드', 'cb_class'),
            ('규격', 'cb_spec'),
            ('관리단위', 'cb_unit'),
            ('공급업체', 'cb_supplier'),
            ('제조사', 'cb_mfr'),
            ('제조국', 'cb_origin')
        ]:
            if hasattr(self, attr):
                fields[key] = getattr(self, attr)
        
        # Store lists for autocomplete (Main Catalog + History)
        self.co_code_list = []
        self.eq_code_list = []
        self.item_name_list = []
        self.class_list = []
        self.spec_list = []
        self.unit_list = []
        self.supplier_list = []
        self.mfr_list = []
        self.origin_list = []
        self.sn_list = []
        self.model_list = []

        attr_mapping = {
            '회사코드': 'co_code_list',
            '관리품번': 'eq_code_list',
            '품목명': 'item_name_list',
            '품목군코드': 'class_list',
            '규격': 'spec_list',
            '관리단위': 'unit_list',
            '공급업체': 'supplier_list',
            '제조사': 'mfr_list',
            '제조국': 'origin_list',
            'SN': 'sn_list',
            '모델명': 'model_list'
        }

        # [PERF] Pre-filter active materials and precompute consumable sets in a single pass
        if not self.materials_df.empty:
            active_mask = pd.to_numeric(self.materials_df.get('Active', 1), errors='coerce').fillna(1) > 0.5
            active_df = self.materials_df[active_mask]
        else:
            active_df = self.materials_df

        consumable_names = set()
        non_consumable_names = set()
        consumable_models = set()
        consumable_mfrs = set()
        if not active_df.empty and '품목명' in active_df.columns:
            p_series = active_df['품목명'].fillna('').astype(str).str.strip()
            m_series = active_df['모델명'].fillna('').astype(str).str.strip() if '모델명' in active_df.columns else pd.Series([''] * len(active_df), index=active_df.index)
            mfr_series = active_df['제조사'].fillna('').astype(str).str.strip() if '제조사' in active_df.columns else pd.Series([''] * len(active_df), index=active_df.index)
            is_cons_cache = {}
            for p_name, model_val, mfr_val in zip(p_series, m_series, mfr_series):
                if not p_name:
                    continue
                is_c = is_cons_cache.get(p_name)
                if is_c is None:
                    is_c = self._is_consumable_material(p_name, '')
                    is_cons_cache[p_name] = is_c
                if is_c:
                    consumable_names.add(p_name)
                    if model_val and model_val.lower() != 'nan':
                        consumable_models.add(model_val)
                    if mfr_val and mfr_val.lower() != 'nan':
                        consumable_mfrs.add(mfr_val)
                else:
                    non_consumable_names.add(p_name)

        for col, list_attr in attr_mapping.items():
            vals = []
            if not active_df.empty and col in active_df.columns:
                unique_vals = active_df[col].dropna().unique()
                vals = sorted([str(v).strip() for v in unique_vals if v and str(v).strip()])
            
            # Update instance attributes for autocomplete
            if col == '관리단위':
                # Merge DB units with the managed daily_units
                self.unit_list = list(dict.fromkeys(self.daily_units + vals))
                if hasattr(self, 'cb_daily_unit'):
                    self.cb_daily_unit['values'] = self.unit_list
            else:
                if col == '회사코드':
                    extra_co = set(getattr(self, 'co_code_list', []) or []) | consumable_mfrs
                    if hasattr(self, 'daily_usage_df') and not self.daily_usage_df.empty and '회사코드' in self.daily_usage_df.columns:
                        for h_co in self.daily_usage_df['회사코드'].dropna().unique():
                            h_co_s = str(h_co).strip()
                            if h_co_s and h_co_s.lower() != 'nan' and h_co_s not in ('한국지역난방공사', '한국지역난방공사 중앙지사'):
                                extra_co.add(h_co_s)
                    vals = sorted(set(vals) | extra_co)
                setattr(self, list_attr, vals)
                if col == '회사코드' and hasattr(self, 'ndt_company_entries'):
                    for entries in self.ndt_company_entries:
                        cb_ndt_co = entries.get('_company')
                        if cb_ndt_co is not None:
                            try:
                                cb_ndt_co['values'] = vals
                            except Exception:
                                pass
            
            # 1. Update registration fields (without "전체")
            if col in fields:
                fields[col]['values'] = vals

            # 2. Update Stock View filters (with "전체")
            # [NEW] For Stock View filters, also exclude non-consumable equipment to match the list view
            filter_map = {
                '회사코드': getattr(self, 'cb_filter_co', None),
                '품목군코드': getattr(self, 'cb_filter_class', None),
                '제조사': getattr(self, 'cb_filter_mfr', None),
                '품목명': getattr(self, 'cb_filter_name', None),
                'SN': getattr(self, 'cb_filter_sn', None),
                '모델명': getattr(self, 'cb_filter_model', None),
                '관리품번': getattr(self, 'cb_filter_eq', None)
            }
            if col in filter_map and filter_map[col] is not None:
                combo = filter_map[col]
                
                # Filter vals for consumables if it's the item name or model
                if col == '품목명':
                    c_vals = [v for v in vals if v in consumable_names or self._is_consumable_material(v, '')]
                elif col == '모델명':
                    c_vals = [v for v in vals if v in consumable_models]
                else:
                    c_vals = vals
                
                combo['values'] = ["전체"] + c_vals
                if not combo.get():
                    combo.set("전체")
        
        # Add Carestream film options to 품목명 (Registration tab only)
        if hasattr(self, 'cb_item_name'):
            all_mat_vals = sorted(list(set([str(mat) for mat in self.item_name_list + self.carestream_films if pd.notna(mat) and str(mat).strip()])))
            self.item_name_list[:] = all_mat_vals
            self.cb_item_name['values'] = self.item_name_list
            
        # Update cb_eq_name with equipment names
        if hasattr(self, 'cb_eq_name'):
            eq_vals = list(getattr(self, 'equipments', []) or [])
            eq_vals.extend(non_consumable_names)
            self.cb_eq_name['values'] = sorted(list(set(eq_vals)))
            
        # [FIX] Do NOT overwrite cb_filter_name values here; it was already set with consumable filtering above.
        # Just ensure the autocomplete list is refreshed when already ready (create_widgets schedules initial pass via after(200)).
        if getattr(self, 'is_ready', False) and hasattr(self, 'apply_autocomplete_to_all_comboboxes'):
            self.apply_autocomplete_to_all_comboboxes()
        
        # 3. Final view update now that filters are set to "전체"
        import threading
        self.root.after(100, self.update_stock_view)

    def register_material(self):
        co_code = self.cb_co_code.get().strip()
        item_name = self.cb_item_name.get().strip()
        eq_name = self.cb_eq_name.get().strip() if hasattr(self, 'cb_eq_name') else ''
        eq_code = self.cb_eq_code.get().strip()
        sn = self.ent_sn.get().strip()
        classification = self.cb_class.get().strip()
        spec = self.cb_spec.get().strip()
        unit = self.cb_unit.get().strip()
        supplier = self.cb_supplier.get().strip()
        manufacturer = self.cb_mfr.get().strip()
        origin = self.cb_origin.get().strip()
        
        try:
            init_stock = float(self.ent_init.get())
            reorder_point = float(self.ent_reorder.get())
            price_val = float(self.ent_price.get() if self.ent_price.get() else 0)
            cost_val = float(self.ent_cost.get() if self.ent_cost.get() else 0)
        except ValueError:
            messagebox.showwarning("입력 오류", "재고, 단가, 원가 정보는 숫자여야 합니다.")
            return
            
        if not item_name and not eq_name:
            messagebox.showwarning("입력 오류", "자재명 또는 장비명을 입력해주세요.")
            return

        final_item_name = eq_name if eq_name else item_name
        is_equipment = bool(eq_name)
        
        old_materials = self.materials_df.copy(deep=True)
        old_equipments = list(getattr(self, 'equipments', []))
        if is_equipment and hasattr(self, 'equipments') and eq_name not in self.equipments:
            self.equipments.append(eq_name)
        
        # Generate MaterialID
        if self.materials_df.empty:
            mat_id = 1
        else:
            valid_ids = pd.to_numeric(self.materials_df['MaterialID'], errors='coerce').dropna()
            mat_id = int(valid_ids.max()) + 1 if not valid_ids.empty else 1
        
        # Extract SN from Model Name (Now using newly added model name combobox)
        model_name = self.cb_model.get().strip() if hasattr(self, 'cb_model') else ''
        new_model, new_sn = self.extract_sn_from_model(model_name, sn)
        new_row = {
            'MaterialID': mat_id,
            '회사코드': co_code,
            '관리품번': eq_code,
            '품목명': final_item_name,
            'SN': new_sn,
            '창고': '',
            '모델명': new_model,
            '규격': spec,
            '품목군코드': classification,
            '공급업체': supplier,
            '제조사': manufacturer,
            '제조국': origin,
            '가격': price_val,
            '원가': cost_val,
            '관리단위': unit if unit else 'EA',
            '수량': init_stock,
            '재고하한': reorder_point,
            'Active': 1
        }
        
        try:
            self.materials_df = pd.concat([self.materials_df, pd.DataFrame([new_row])], ignore_index=True)
            if not self.save_data(require_primary=True):
                raise RuntimeError('원본 재고 파일에 저장하지 못했습니다.')
        except Exception as error:
            self.materials_df = old_materials
            if hasattr(self, 'equipments'):
                self.equipments[:] = old_equipments
            messagebox.showerror("등록 저장 실패", f"{error}\n입력 내용은 그대로 유지됩니다.", parent=self.root)
            return
        self.update_material_combo()
        self.update_registration_combos()
        self.update_stock_view()
        type_str = "장비" if is_equipment else "자재"
        messagebox.showinfo("완료", f"'{final_item_name}' {type_str}가 등록되었습니다.")
        
        # Clear entries
        self.cb_co_code.set('')
        self.cb_eq_code.set('')
        self.cb_item_name.set('')
        if hasattr(self, 'cb_eq_name'): self.cb_eq_name.set('')
        self.ent_sn.delete(0, tk.END)
        self.cb_class.set('')
        self.cb_spec.set('')
        self.cb_unit.set('')
        self.cb_supplier.set('')
        self.cb_mfr.set('')
        self.cb_origin.set('')
        self.ent_reorder.delete(0, tk.END)
        self.ent_reorder.insert(0, "0")
        self.ent_init.delete(0, tk.END)
        self.ent_init.insert(0, "0")
        # Clear additional fields
        if hasattr(self, 'cb_model'):
            self.cb_model.set('')
        if hasattr(self, 'ent_price'):
            self.ent_price.delete(0, tk.END)
        if hasattr(self, 'ent_cost'):
            self.ent_cost.delete(0, tk.END)

    def add_transaction(self):
        """Record an IN or OUT transaction"""
        try:
            mat_selection = self.cb_material.get().strip()
            t_type = self.cb_type.get()
            user = self.ent_user.get().strip()
            
            if not mat_selection:
                messagebox.showwarning("입력 오류", "자재 또는 장비를 선택해주세요.")
                return
            if not t_type:
                messagebox.showwarning("입력 오류", "구분(입고/출고)을 선택해주세요.")
                return
                
            try:
                qty_str = self.ent_qty.get().strip()
                if not qty_str:
                    messagebox.showwarning("입력 오류", "수량을 입력해주세요.")
                    return
                qty = float(qty_str)
            except ValueError:
                messagebox.showwarning("입력 오류", "수량은 숫자여야 합니다.")
                return
                
            note = self.ent_note.get().strip()
            
            # Extract pure material name from selection
            mat_name = mat_selection
            if " - " in mat_name: mat_name = mat_name.split(" - ")[0]
            if " (SN: " in mat_name: mat_name = mat_name.split(" (SN: ")[0]
            pure_mat_name = mat_name.strip()
            
            # Find MaterialID
            # Ensure MaterialID is treated consistently
            mat_rows = self.materials_df[self.materials_df['품목명'] == pure_mat_name]
            
            if mat_rows.empty:
                # If exact match fails, try a case-insensitive or stripped match
                mat_rows = self.materials_df[self.materials_df['품목명'].str.strip() == pure_mat_name]
            
            if mat_rows.empty:
                # [NEW] Ask to register new material instead of just showing error
                confirm = messagebox.askyesno("신규 자재", f"'{pure_mat_name}'은(는) 등록되지 않은 자재입니다. 신규 자재로 등록하고 진행할까요?")
                if confirm:
                    mat_id = self.register_new_material(pure_mat_name)
                else:
                    return
            else:
                mat_id = mat_rows['MaterialID'].values[0]
            
            old_materials = self.materials_df.copy(deep=True)
            old_transactions = self.transactions_df.copy(deep=True)
            # Update Warehouse in materials_df
            warehouse = str(self.cb_warehouse.get()).strip()
            if warehouse:
                # Type-safe assignment
                if '창고' in self.materials_df.columns:
                    mask = self.materials_df['MaterialID'] == mat_id
                    if mask.any():
                        self.materials_df.loc[mask, '창고'] = warehouse
            
            # Create transaction record
            new_trans = {
                'Date': datetime.datetime.now(),
                'MaterialID': mat_id,
                'Type': t_type,
                'Quantity': qty if t_type == 'IN' else -qty,
                'Note': note,
                'User': user,
                'Site': self.cb_trans_site.get() if hasattr(self, 'cb_trans_site') else '',
                '차량번호': '',
                '주행거리': '',
                '차량점검': '',
                '차량비고': ''
            }
            
            try:
                # Add to dataframe
                self.transactions_df = pd.concat([self.transactions_df, pd.DataFrame([new_trans])], ignore_index=True)
                # Save all data (force sync)
                if not self.save_data(require_primary=True):
                    raise RuntimeError('원본 재고 파일에 입출고 내역을 저장하지 못했습니다.')
            except Exception as save_err:
                self.materials_df = old_materials
                self.transactions_df = old_transactions
                messagebox.showerror("입출고 저장 실패", f"{save_err}\n원본 기록을 복원했으며 입력 내용은 유지됩니다.", parent=self.root)
                return
            
            # Check if it was really added (for debug feedback)
            last_count = len(self.transactions_df)
            
            # Refresh views
            self.update_stock_view()
            self.update_transaction_view()
            self.update_material_combo() 
            
            # Auto-save site/user to lists
            site_value = new_trans['Site']
            if site_value and site_value not in self.sites:
                self.sites.append(site_value)
                self.sites.sort()
                self.save_tab_config()
            
            if user and user not in self.users:
                self.users.append(user)
                self.users.sort()
                self.save_tab_config()
            
            # Success feedback
            messagebox.showinfo("완료", f"{pure_mat_name} {t_type} 처리되었습니다.\n(전체 기록 수: {last_count}개)")
            
            # Clear UI fields
            self.ent_qty.delete(0, tk.END)
            self.ent_note.delete(0, tk.END)
            if hasattr(self, 'ent_user') and hasattr(self.ent_user, 'set'):
                self.ent_user.set('')
            elif hasattr(self, 'ent_user'):
                self.ent_user.delete(0, tk.END)
            
            if hasattr(self, 'cb_warehouse'): self.cb_warehouse.set('')
            if hasattr(self, 'cb_trans_site'): self.cb_trans_site.set('')
            
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            self.show_error_dialog("저장 오류", f"기록 저장 중 기술적인 오류가 발생했습니다:\n{e}\n\n상세 정보:\n{error_details}")
        
        # Ensure view is updated regardless
        self.update_transaction_view()

    def update_transaction_view(self):
        """Unified transaction view that combines manual stock transactions and site usage."""
        try:
            # 0. Helper Functions (Global normalize_id is now used)
            def _f(val):
                if val is None or pd.isna(val): return 0.0
                try: 
                    s = str(val).replace(',', '').strip()
                    return float(s) if s else 0.0
                except: return 0.0
                
            def _q_str(val):
                try:
                    v = abs(float(str(val).replace(',', '').strip()))
                    return f"{v:g}" if v > 0 else ""
                except: return ""

            # 1. Clear existing view
            for item in self.inout_tree.get_children():
                self.inout_tree.delete(item)
            
            # 2. Prepare Transaction Records
            df_trans = pd.DataFrame()
            active_ids = []
            if not self.materials_df.empty and 'MaterialID' in self.materials_df.columns:
                # Get list of MaterialIDs that are NOT deleted (Active != 0)
                active_mask = (self.materials_df.get('Active', 1) != 0)
                active_ids = [self.normalize_id(x) for x in self.materials_df[active_mask]['MaterialID'].dropna().tolist()]
                
            if not self.transactions_df.empty:
                df_trans = self.transactions_df.copy()
                # Column Compatibility
                if 'Site' not in df_trans.columns and '현장' in df_trans.columns: df_trans['Site'] = df_trans['현장']
                if 'Date' not in df_trans.columns and '날짜' in df_trans.columns: df_trans['Date'] = df_trans['날짜']
                if 'Date' not in df_trans.columns and '일자' in df_trans.columns: df_trans['Date'] = df_trans['일자']
                
                df_trans['Date'] = pd.to_datetime(df_trans['Date'], errors='coerce')
                if 'EntryDate' in df_trans.columns:
                    df_trans['Date'] = df_trans['Date'].fillna(pd.to_datetime(df_trans['EntryDate'], errors='coerce'))
                
                if active_ids:
                    df_trans = df_trans[df_trans['MaterialID'].apply(self.normalize_id).isin(active_ids)]
            
            # 2b. Prepare Daily Usage Records
            df_daily = pd.DataFrame()
            if hasattr(self, 'daily_usage_df') and not self.daily_usage_df.empty:
                df_daily = self.daily_usage_df.copy()
                # [FIX] Normalize column names to ensure reliable extraction of NDT_ fields
                df_daily.columns = [str(c).strip().replace(' ', '') for c in df_daily.columns]
                
                df_daily['Date'] = pd.to_datetime(df_daily['Date'], errors='coerce')
                # If Date is missing but EntryTime exists, use EntryTime
                if 'EntryTime' in df_daily.columns:
                    df_daily['Date'] = df_daily['Date'].fillna(pd.to_datetime(df_daily['EntryTime'], errors='coerce'))
                
            print(f"DEBUG: update_transaction_view - TRANS: {len(df_trans)}, DAILY: {len(df_daily)}")
            def _f(val):
                if val is None or pd.isna(val): return 0.0
                try: 
                    s = str(val).replace(',', '').strip()
                    return float(s) if s else 0.0
                except: return 0.0
            # [FIX] Do NOT filter by active_ids here; process everything and handle missing info later

            # 3. Create Harmonized Unified Data
            unified_rows = []
                
            # Add Transaction rows (Source: TRANS)
            if not df_trans.empty:
                for idx, row in df_trans.iterrows():
                    mat_id = self.normalize_id(row.get('MaterialID'))
                    mat_name = self.get_material_name_only(mat_id)
                    mat_info = self.get_material_info(mat_id)
                    
                    is_con = self._is_consumable_material(mat_name, '')
                    mat_col = mat_name if is_con else "-"
                    eq_col = "-" if is_con else mat_name

                    # [NEW] Deduplication: Skip auto-generated transactions that are already covered by DAILY usage entries
                    # This removes the "50 vs -50" duplication while keeping the richer DAILY info.
                    note = str(row.get('Note', ''))
                    if "(자동 차감)" in note:
                        continue # Skip ALL auto-generated TRANS records as they are aggregated in the DAILY record

                    unified_rows.append({
                        'Date': row['Date'],
                        'Site': row.get('Site', ''),
                        'Type': row.get('Type', 'OUT'),
                        '품목명': mat_col,
                        '장비명': eq_col,
                        'SN': mat_info.get('SN', ''),
                        '규격': mat_info.get('규격', ''),
                        'MaterialID': mat_id,
                        'Quantity': row.get('Quantity', 0),
                        'Warehouse': row.get('Warehouse', ''),
                        'User': row.get('User', ''),
                        'Note': row.get('Note', ''),
                        '차량번호': row.get('차량번호', ''),
                        'Source': 'TRANS',
                        'Fee': 0, # Transactions usually don't have inspection fees directly
                        'OrigIdx': idx
                    })


            # Add Daily Usage rows (Source: DAILY)
            if not df_daily.empty:
                for idx, row in df_daily.iterrows():
                    # [FLEXIBLE] Extract usage with multiple fallback names
                    u_val = _f(row.get('Usage', row.get('검사량', row.get('수량', row.get('Quantity', 0)))))
                    f_val = _f(row.get('FilmCount', row.get('매수', 0)))
                    
                    # If everything is 0, we still might want to show it if there's a note or site info
                    # but for now let's just make sure we capture NDT properly
                    ndt_vals_list = [
                        _f(row.get(f'NDT_{n}', row.get(n, 0))) 
                        for n in ['세척제', '침투제', '현상제', '백색페인트', '흑색자분', '형광자분', '형광침투제', '자분페인트']
                    ]
                    
                    # If any sign of life, keep the row
                    if u_val <= 0 and f_val <= 0 and not any(v > 0 for v in ndt_vals_list) and not str(row.get('Site', '')):
                        continue

                    method_val = str(row.get('검사방법', '')).strip().upper()
                    mat_id = self.normalize_id(row.get('MaterialID'))
                    mat_name = self.get_material_name_only(mat_id) if mat_id else ""
                    mat_info = self.get_material_info(mat_id) if mat_id else {}
                    
                    # Extract Equipment from DailyUsage
                    eq_name = str(row.get('장비명', '')).strip()
                    aux_eq = str(row.get('보조장비', '')).strip()
                    if eq_name == 'nan': eq_name = ''
                    if aux_eq == 'nan': aux_eq = ''
                    full_eq = f"{eq_name} ({aux_eq})" if (eq_name and aux_eq) else (eq_name if eq_name else (aux_eq if aux_eq else ""))

                    has_chem = any(v > 0 for v in ndt_vals_list)
                    has_mat = bool(mat_id and self._is_consumable_material(mat_name, method_val))
                    has_eq = bool(full_eq)

                    # [NEW] Filter: Keep row if it has consumable material, NDT chemicals, OR equipment used
                    if not (has_mat or has_chem or has_eq):
                        continue
                    
                    # [FIX] Filter removed to show all site usage in history
                    
                    # Collect all worker names from User, User2...User10
                    workers = []
                    for i in range(1, 11):
                        u_col = 'User' if i == 1 else f'User{i}'
                        u_val_str = str(row.get(u_col, '')).strip()
                        if u_val_str and u_val_str not in ['nan', 'None', '']:
                            # Handle (주간) suffix if present
                            u_clean = re.sub(r'\(.*?\)\s*', '', u_val_str).strip()
                            if u_clean: workers.append(u_clean)
                    
                    worker_str = ", ".join(workers) if workers else str(row.get('User', ''))

                    # [FIX] Quantity for display:
                    # If RT: show film count / usage
                    # If not RT and no main consumable material: disp_qty should be 0 (chemicals have their own columns)
                    if method_val == 'RT' or (mat_id and self._is_consumable_material(mat_name, method_val)):
                        disp_qty = f_val if f_val > 0 else u_val
                    else:
                        disp_qty = 0

                    # [ROBUST] Try to parse ndt_data if it exists as a JSON string
                    ndt_json = {}
                    nj_raw = row.get('ndt_data', '')
                    if isinstance(nj_raw, str) and nj_raw.strip().startswith('{'):
                        try: ndt_json = json.loads(nj_raw)
                        except: pass
                    elif isinstance(nj_raw, dict):
                        ndt_json = nj_raw

                    unified_rows.append({
                        'Date': row['Date'],
                        'Site': row.get('Site', ''),
                        'Type': 'OUT',
                        '품목명': mat_name if has_mat else "-",
                        '장비명': full_eq if full_eq else "-",
                        'SN': mat_info.get('SN', ''),
                        '규격': mat_info.get('규격', ''),
                        'MaterialID': mat_id,
                        'Quantity': disp_qty,
                        'NDT_세척제': _f(row.get('NDT_세척제', row.get('세척제', ndt_json.get('세척제', ndt_json.get('PT Cleaner', ndt_json.get('PTCleaner', 0)))))),
                        'NDT_침투제': _f(row.get('NDT_침투제', row.get('침투제', ndt_json.get('침투제', ndt_json.get('PT Penetrant', ndt_json.get('PTPenetrant', 0)))))),
                        'NDT_현상제': _f(row.get('NDT_현상제', row.get('현상제', ndt_json.get('현상제', ndt_json.get('PT Developer', ndt_json.get('PTDeveloper', 0)))))),
                        'NDT_백색페인트': _f(row.get('NDT_백색페인트', row.get('NDT_페인트', row.get('백색페인트', ndt_json.get('백색페인트', ndt_json.get('MT WHITE', ndt_json.get('MT7C-WHITE', 0))))))),
                        'NDT_흑색자분': row.get('NDT_흑색자분', row.get('NDT_자분', ndt_json.get('흑색자분', ndt_json.get('MT 7C-BLACK', 0)))),
                        'Warehouse': '',
                        'User': worker_str,
                        'Note': row.get('Note', ''),
                        '차량번호': row.get('차량번호', ''),
                        'Source': 'DAILY',
                        'Fee': row.get('검사비', 0.0), # Bring fee from daily usage
                        'OrigIdx': idx
                    })
            
            if not unified_rows:
                return
                
            df_unified = pd.DataFrame(unified_rows)
            
            # Apply Filters
            # Material / Equipment Filter
            if hasattr(self, 'cb_trans_filter_mat'):
                selected_mat = str(self.cb_trans_filter_mat.get()).strip()
                if selected_mat and selected_mat != "전체":
                    matching_rows = []
                    for idx, row in df_unified.iterrows():
                        mid_name = self.get_material_display_name(row['MaterialID'])
                        pm_name = str(row.get('품목명', '')).strip()
                        eq_name = str(row.get('장비명', '')).strip()
                        if selected_mat in (mid_name, pm_name, eq_name):
                            matching_rows.append(idx)
                    df_unified = df_unified.loc[matching_rows]
            
            # Site Filter
            if hasattr(self, 'cb_trans_filter_site'):
                selected_site = str(self.cb_trans_filter_site.get()).strip()
                if selected_site and selected_site != "전체":
                    df_unified = df_unified[df_unified['Site'].astype(str).str.contains(selected_site, na=False, case=False, regex=False)]
            
            # Vehicle Filter
            if hasattr(self, 'cb_trans_filter_vehicle'):
                selected_vehicle = str(self.cb_trans_filter_vehicle.get()).strip()
                if selected_vehicle and selected_vehicle != "전체":
                    df_unified = df_unified[df_unified['차량번호'].astype(str).str.contains(selected_vehicle, na=False, case=False, regex=False)]

            # [NEW] Update Filter Dropdowns with latest unified data unique values
            if hasattr(self, 'cb_trans_filter_site') and not getattr(self, '_site_filter_busy', False):
                self._site_filter_busy = True
                try:
                    s_current = self.cb_trans_filter_site.get()
                    s_vals = sorted(df_unified['Site'].dropna().astype(str).unique().tolist())
                    self.cb_trans_filter_site['values'] = ["전체"] + [v for v in s_vals if v.strip()]
                    if s_current: self.cb_trans_filter_site.set(s_current)
                    else: self.cb_trans_filter_site.set("전체")
                finally: self._site_filter_busy = False

            if hasattr(self, 'cb_trans_filter_mat') and not getattr(self, '_mat_filter_busy', False):
                self._mat_filter_busy = True
                try:
                    m_current = self.cb_trans_filter_mat.get()
                    # Get display names for all MaterialIDs in the unified set
                    unique_mat_ids = df_unified['MaterialID'].dropna().unique()
                    m_vals = [self.get_material_display_name(mid) for mid in unique_mat_ids]
                    for eq in df_unified['장비명'].dropna().unique():
                        eq_s = str(eq).strip()
                        if eq_s and eq_s != '-':
                            m_vals.append(eq_s)
                    for pm in df_unified['품목명'].dropna().unique():
                        pm_s = str(pm).strip()
                        if pm_s and pm_s != '-':
                            m_vals.append(pm_s)
                    m_vals = sorted(list(set([v for v in m_vals if v.strip()])))
                    self.cb_trans_filter_mat['values'] = ["전체"] + m_vals
                    if m_current: self.cb_trans_filter_mat.set(m_current)
                    else: self.cb_trans_filter_mat.set("전체")
                finally: self._mat_filter_busy = False
                
            # [STABILITY] Ensure all standard columns are visible and reset any previous hiding
            all_inout_cols = ('날짜', '현장', '구분', '품목명', '장비명', '수량', '세척제', '침투제', '현상제', '백색페인트', '흑색자분', 'SN', '규격', '창고', '담당자', '비고', '검사비')
            
            # Use saved configuration if available
            if hasattr(self, 'tab_config') and 'inout_visible_cols' in self.tab_config:
                saved_cols = [c for c in self.tab_config['inout_visible_cols'] if c in all_inout_cols]
                if '장비명' not in saved_cols:
                    if '품목명' in saved_cols:
                        idx_pm = saved_cols.index('품목명')
                        saved_cols.insert(idx_pm + 1, '장비명')
                    else:
                        saved_cols.append('장비명')
                    self.tab_config['inout_visible_cols'] = saved_cols
                self.inout_tree['displaycolumns'] = saved_cols
            else:
                self.inout_tree['displaycolumns'] = ('날짜', '현장', '구분', '품목명', '장비명', '수량', '세척제', '침투제', '현상제', '백색페인트', '흑색자분', '비고')

            # Sort and display
            df_sorted = df_unified.sort_values(by='Date', ascending=False)
            
            for idx, row in df_sorted.iterrows():
                try:
                    # Robust date formatting with fallback for NaT
                    if pd.isna(row['Date']):
                        usage_date = "알 수 없음"
                    else:
                        usage_date = row['Date'].strftime('%Y-%m-%d %H:%M')
                    
                    # Manage Numeric Fee Formatting
                    f_val = row.get('Fee', 0.0)
                    fee_str = f"{f_val:,.0f}" if (isinstance(f_val, (int, float)) and f_val > 0) else ""

                    # Normalize Quantity for display: Use absolute value since 'Type' (IN/OUT) indicates direction
                    raw_qty = row.get('Quantity', row.get('수량', 0))
                    try:
                        clean_qty = abs(float(str(raw_qty).replace(',', '')))
                        qty_str = (f"{clean_qty:.1f}" if clean_qty % 1 != 0 else f"{int(clean_qty)}") if clean_qty > 0 else ""
                    except:
                        qty_str = self.clean_nan(raw_qty)

                    # NDT columns formatting
                    def _q_str(val):
                        try:
                            v = abs(float(str(val).replace(',', '')))
                            return f"{v:g}" if v > 0 else ""
                        except: return ""

                    self.inout_tree.insert('', tk.END, values=(
                        usage_date,
                        self.clean_nan(row.get('Site', '')),
                        self.clean_nan(row.get('Type', '-')),
                        self.clean_nan(row.get('품목명', '')),
                        self.clean_nan(row.get('장비명', '')),
                        qty_str,
                        _q_str(row.get('NDT_세척제', 0)),
                        _q_str(row.get('NDT_침투제', 0)),
                        _q_str(row.get('NDT_현상제', 0)),
                        _q_str(row.get('NDT_백색페인트', 0)),
                        _q_str(row.get('NDT_흑색자분', 0)),
                        self.clean_nan(row.get('SN', '')),
                        self.clean_nan(row.get('규격', '')),
                        self.clean_nan(row.get('Warehouse', '')),
                        self.clean_nan(row.get('User', '')),
                        self.clean_nan(row.get('Note', '')),
                        fee_str
                    ), tags=(str(row.get('OrigIdx', '')), row.get('Source', 'TRANS')))


                except Exception as e:
                    print(f"Row Display Error: {e}")
                    continue
                    
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"Unified Transaction View Error: {e}\n{error_details}")

    def refresh_inout_history(self):
        """Refresh In/Out history and related filter source lists."""
        try:
            # Ensure filter candidates reflect latest data first
            self.update_material_combo()
            # Then redraw history view (skip hidden stock view recalculation on simple tab switch)
            self.update_transaction_view()
        except Exception as e:
            messagebox.showerror("새로고침 오류", f"입출고 내역 새로고침 중 오류가 발생했습니다.\n{e}")

    def delete_transaction_entry(self):
        """Delete selected transaction from history and refresh"""
        selection = self.inout_tree.selection()
        if not selection:
            messagebox.showwarning("선택 오류", "삭제할 기록을 선택해주세요.")
            return
            
        targets = []
        try:
            for item in selection:
                tags = self.inout_tree.item(item, 'tags')
                if len(tags) < 2 or tags[1] not in ('TRANS', 'DAILY'):
                    raise ValueError('기록의 원본 정보를 확인할 수 없습니다.')
                idx, source = int(tags[0]), tags[1]
                frame = self.transactions_df if source == 'TRANS' else self.daily_usage_df
                if idx not in frame.index:
                    raise ValueError('기록이 변경되었습니다. 새로고침 후 다시 선택해주세요.')
                targets.append((source, idx))
        except (ValueError, TypeError) as error:
            messagebox.showwarning('삭제 오류', str(error), parent=self.root)
            return
        daily_count = len({idx for source, idx in targets if source == 'DAILY'})
        prompt = '선택한 기록을 삭제할까요? 삭제 시 재고 계산에 반영됩니다.'
        if daily_count:
            prompt += (f'\n\n현장 사용 {daily_count}건은 원본 일일 작업 기록 전체가 삭제됩니다.'
                       '\n검사량·작업자·작업시간·메모도 함께 삭제되어 조회 및 집계에 영향을 줍니다.')
        if not messagebox.askyesno('삭제 확인', prompt, parent=self.root):
            return
        old_transactions = self.transactions_df.copy(deep=True)
        old_daily = self.daily_usage_df.copy(deep=True)
        try:
            self.transactions_df = self.transactions_df.drop(
                list({idx for source, idx in targets if source == 'TRANS'}))
            self.daily_usage_df = self.daily_usage_df.drop(
                list({idx for source, idx in targets if source == 'DAILY'}))
            if not self.save_data(require_primary=True):
                raise RuntimeError('원본 파일에 저장하지 못했습니다.')
        except Exception as error:
            self.transactions_df = old_transactions
            self.daily_usage_df = old_daily
            messagebox.showerror('삭제 저장 실패', f'{error}\n화면의 원본 기록을 복원했습니다.', parent=self.root)
            return
        self.update_transaction_view()
        self.update_stock_view()
        
        if daily_count:
            self.update_recent_entries_view()
            self.update_daily_usage_view()
            self.update_monthly_usage_view()
        if hasattr(self, 'budget_view_tree'):
            self.update_budget_site_view()

        messagebox.showinfo("완료", "거래 기록이 삭제되었습니다.")

    def on_material_selected(self, *args, **kwargs):
        from site_apps.central.src.controllers.event_controller import on_material_selected_impl
        return on_material_selected_impl(self, *args, **kwargs)

    def _on_trans_site_return(self, *args, **kwargs):
        from site_apps.central.src.controllers.event_controller import _on_trans_site_return_impl
        return _on_trans_site_return_impl(self, *args, **kwargs)

    def _on_warehouse_return(self, *args, **kwargs):
        from site_apps.central.src.controllers.event_controller import _on_warehouse_return_impl
        return _on_warehouse_return_impl(self, *args, **kwargs)

    def _on_user_return(self, *args, **kwargs):
        from site_apps.central.src.controllers.event_controller import _on_user_return_impl
        return _on_user_return_impl(self, *args, **kwargs)


