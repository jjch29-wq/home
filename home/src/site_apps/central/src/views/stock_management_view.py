from site_apps.central.src.utils.helpers import NAN_PATTERN, DOT_ZERO_PATTERN, MARKER_PATTERN
from site_apps.central.src.views.components import *
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
    
    btn_refresh = ttk.Button(action_row, text="🔄 재고 새로고침", command=self.reload_and_update_stock_view, style="Primary.TButton")
    btn_refresh.pack(side='left', padx=3)
    
    btn_alert = ttk.Button(action_row, text="⚠️ 재주문 필요 항목", command=self.show_low_stock, style="Warning.TButton")
    btn_alert.pack(side='left', padx=3)
    
    btn_edit = ttk.Button(action_row, text="✏️ 품목 수정", command=self.open_edit_material_dialog, style="Info.TButton")
    btn_edit.pack(side='left', padx=3)

    from site_apps.central.src.views.stock_history import show_material_history
    ttk.Button(action_row, text="📜 품목 이력", command=lambda: show_material_history(self), style="Secondary.TButton").pack(side='left', padx=3)

    btn_delete = ttk.Button(action_row, text="🗑️ 품목 삭제", command=self.delete_selected_material, style="Danger.TButton")
    btn_delete.pack(side='left', padx=3)

    btn_export = ttk.Button(action_row, text="📊 엑셀 내보내기", command=self.export_stock_to_excel, style="Action.TButton")
    btn_export.pack(side='left', padx=3)
    
    btn_select_all = ttk.Button(action_row, text="전체 선택", command=self.select_all_stock, style="Secondary.TButton")
    btn_select_all.pack(side='left', padx=3)
    self.stock_allocation_notice = ttk.Label(control_frame, foreground='#A33A00')
    self.stock_allocation_notice.pack(fill='x')
    
    # [NEW] Popout Button
    btn_popout_stock = ttk.Button(action_row, text="🗗 팝업창으로 열기", command=self.open_detached_stock_view, style="Info.TButton")
    btn_popout_stock.pack(side='right', padx=5)
    
    # Row 2: Search and Filter Frame
    filter_row = ttk.Frame(control_frame)
    filter_row.pack(fill='x', side='top')
    
    filter_frame = ttk.LabelFrame(filter_row, text="🔍 검색 필터")
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
    def _debounced_stock_search(*args):
        if hasattr(self, '_stock_search_timer') and self._stock_search_timer:
            self.root.after_cancel(self._stock_search_timer)
        self._stock_search_timer = self.root.after(150, self.update_stock_view)
    self.search_var.trace_add('write', _debounced_stock_search)
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

    def _show_stock_tree_context_menu(event):
        region = self.stock_tree.identify_region(event.x, event.y)
        if region == "heading":
            self._show_generic_tree_heading_context_menu(event, self.stock_tree)
            return
        item = self.stock_tree.identify_row(event.y)
        if item:
            self.stock_tree.selection_set(item)
            menu = tk.Menu(self.root, tearoff=0)
            menu.add_command(label="✏️ 품목 수정...", command=self.open_edit_material_dialog)
            menu.add_command(label="📜 품목 이력 보기...", command=lambda: show_material_history(self))
            menu.add_command(label="🗑️ 품목 삭제", command=self.delete_selected_material)
            menu.add_separator()
            menu.add_command(label="🔄 재고 새로고침", command=self.reload_and_update_stock_view)
            menu.add_command(label="📊 엑셀 내보내기", command=self.export_stock_to_excel)
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()

    self.enable_tree_column_drag(self.stock_tree, context_menu_handler=_show_stock_tree_context_menu)
    
    # Grid layout
    self.stock_tree.grid(row=0, column=0, sticky='nsew')
    vsb.grid(row=0, column=1, sticky='ns')
    hsb.grid(row=1, column=0, sticky='ew')
    
    tree_frame.grid_rowconfigure(0, weight=1)
    tree_frame.grid_columnconfigure(0, weight=1)
    
    # Initial view update will be triggered after update_registration_combos 
    # to ensure filters are properly initialized to "전체" first.
    if getattr(self, 'is_ready', False):
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
    
    from site_apps.central.src.models.stock_calculation import build_stock_lookups, stock_breakdown
    stock_lookups = build_stock_lookups(self)
    unassigned = stock_lookups[4].unassigned
    if hasattr(self, 'stock_allocation_notice'):
        self.stock_allocation_notice.configure(text=(
            f'약품 사용 {len(unassigned)}건의 품목·창고를 특정할 수 없어 재고 차감을 보류했습니다. 입력 기록의 자재 연결·창고를 확인하세요.'
            if unassigned else ''))

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

import json
import shutil
import traceback
from tkinter import simpledialog


def delete_selected_material_impl(self):
    """선택된 자재 항목들을 재고에서 삭제"""
    selected_items = self.stock_tree.selection()

    if not selected_items:
        messagebox.showwarning("선택 오류", "삭제할 품목을 선택해주세요.")
        return

    # Confirm deletion
    confirm = messagebox.askyesno("삭제 확인", f"선택한 {len(selected_items)}개의 품목을 재고에서 영구히 삭제하시겠습니까?\n이 작업은 되돌릴 수 없습니다.")

    if not confirm:
        return

    # Get MaterialIDs to delete
    mat_ids_to_remove = []
    for item in selected_items:
        values = self.stock_tree.item(item, 'values')
        if values:
            # Ensure we match the type of MaterialID in the dataframe
            # [UX IMPROVEMENT] Get hidden MaterialID from the end of values
            mat_ids_to_remove.append(type(self.materials_df['MaterialID'].iloc[0])(values[-1]))

    old_materials = self.materials_df.copy(deep=True)
    # Soft delete: Set Active=0 instead of removing from materials_df
    if 'Active' not in self.materials_df.columns:
        self.materials_df['Active'] = 1

    # Standardize Active column to numeric to avoid pandas TypeError
    self.materials_df['Active'] = pd.to_numeric(self.materials_df['Active'], errors='coerce').fillna(1)

    initial_count = len(self.materials_df[self.materials_df['Active'] != 0])

    # Apply deletion (set to 0)
    mask = self.materials_df['MaterialID'].isin(mat_ids_to_remove)
    self.materials_df.loc[mask, 'Active'] = 0

    final_count = len(self.materials_df[self.materials_df['Active'] != 0])
    removed_count = initial_count - final_count

    if removed_count > 0:
        try:
            if not self.save_data(require_primary=True):
                raise RuntimeError('원본 재고 파일에 삭제 내역을 저장하지 못했습니다.')
        except Exception as error:
            self.materials_df = old_materials
            messagebox.showerror("삭제 저장 실패", f"{error}\n원본 품목 상태를 복원했습니다.", parent=self.root)
            return
        self.update_stock_view()
        self.update_material_combo()

        messagebox.showinfo("완료", f"{removed_count}개의 품목이 삭제되었습니다.")
    else:
        messagebox.showwarning("실패", "데이터프레임에서 항목을 삭제하지 못했습니다.")


def save_material_edits_impl(self, mat_id, new_data):
    """Save edited material data back to the database"""
    # Update DataFrame (Robust matching)
    target_id_str = self.normalize_id(mat_id)
    mask = self.materials_df['MaterialID'].apply(self.normalize_id) == target_id_str
    idx = self.materials_df.index[mask].tolist()
    if not idx:
        messagebox.showerror("오류", "자재를 찾을 수 없습니다.")
        return False

    old_materials = self.materials_df.copy(deep=True)
    for col, val in new_data.items():
        if col in self.materials_df.columns:
            # Check the actual dtype of the column
            col_dtype = self.materials_df[col].dtype

            # Handle different data types appropriately
            if col_dtype == 'float64' or col_dtype == 'int64':
                # Numeric columns - convert to float, empty becomes 0.0
                try:
                    val = float(val) if str(val).strip() else 0.0
                except (ValueError, TypeError):
                    val = 0.0
            else:
                # String/object columns - handle empty values
                if val == '' or val == 'nan' or pd.isna(val):
                    val = ''
                else:
                    val = str(val)

            self.materials_df.at[idx[0], col] = val

    # Re-check SN extraction after edit
    model = self.materials_df.at[idx[0], '모델명']
    sn = self.materials_df.at[idx[0], 'SN']
    new_model, new_sn = self.extract_sn_from_model(model, sn)
    self.materials_df.at[idx[0], '모델명'] = new_model
    self.materials_df.at[idx[0], 'SN'] = new_sn

    try:
        if not self.save_data(require_primary=True):
            raise RuntimeError('원본 재고 파일에 수정 내역을 저장하지 못했습니다.')
    except Exception as error:
        self.materials_df = old_materials
        messagebox.showerror("수정 저장 실패", f"{error}\n원본 품목 정보를 복원했습니다.", parent=self.root)
        return False

    # Refresh everything
    self.update_stock_view()
    self.update_registration_combos()
    self.update_material_combo()

    messagebox.showinfo("완료", "자재 정보가 성공적으로 수정되었습니다.")
    return True


def clean_df_export_impl(self, df):
    """Sanitize data for Excel: remove illegal characters, handle empty values, and protect formulas"""
    # 1. Pre-clean common empty representations
    df = df.replace(['nan', 'NaN', 'None'], "")

    # 2. Identify columns that have at least one meaningful value
    def is_really_empty(col):
        # A column is empty if all stripped values are "", NaN, or 0 sequences
        return df[col].astype(str).str.strip().replace(['nan', 'None', '', '0', '0.0', '0.00'], pd.NA).dropna().empty

    non_empty_cols = [col for col in df.columns if not is_really_empty(col)]

    # Always keep essential columns even if empty
    essential = ['날짜', '현장', '품목명', 'Date', 'Site']
    # Also keep standard RTK and NDT result columns
    rtk_cats = ["센터미스", "농도", "마킹미스", "필름마크", "취급부주의", "고객불만", "기타", "RTK총계"]
    ndt_materials = ["형광자분", "흑색자분", "백색페인트", "침투제", "세척제", "현상제", "형광침투제"]
    essential.extend(rtk_cats)
    essential.extend([f'RTK_{c}' for c in rtk_cats])
    essential.extend(ndt_materials)
    essential.extend([f'NDT_{"".join(m.split())}' for m in ndt_materials])

    for col in essential:
        if col in df.columns and col not in non_empty_cols:
            non_empty_cols.append(col)

    # 3. Sanitize content and protect formulas
    def sanitize_and_escape(val):
        if pd.isna(val) or val == "" or val == 0 or val == 0.0:
            return ""

        # String conversion and cleaning
        s_val = str(val).strip()
        if not s_val or s_val.lower() in ['nan', 'none']:
            return ""

        # Remove illegal XML control characters (ASCII < 32 except tab, LF, CR)
        # These characters cause Excel to report file corruption.
        s_val = "".join(ch for ch in s_val if ord(ch) >= 32 or ch in "\t\n\r")

        # Escape potential formula characters (=, +, -, @)
        # IMPORTANT: Skip escaping if the value is a valid number (e.g., -5000)
        if s_val.startswith(('=', '+', '-', '@')):
            try:
                # If it's a number, don't prefix with '
                float(s_val)
                return val
            except ValueError:
                # It's a text string starting with a formula char
                return f"'{s_val}"

        return s_val

    # Apply sanitization to all columns
    for col in df.columns:
        df[col] = df[col].apply(sanitize_and_escape)

    # Return only non-empty columns in original order
    final_cols = [c for c in df.columns if c in non_empty_cols]
    return df[final_cols]


def setup_import_tab_impl(self):
    import_frame = ttk.LabelFrame(self.tab_import, text="데이터 관리")
    import_frame.pack(pady=20, padx=20, fill='both', expand=True)

    # Import Section
    ttk.Label(import_frame, text="엑셀 파일에서 자재 데이터 가져오기", font=('Arial', 11, 'bold')).pack(pady=10)
    ttk.Label(import_frame, text="형식: MaterialID, 회사코드, 관리품번, 품목명, 창고, 모델명, 규격, 품목군코드, 제조사, 제조국, 가격, 관리단위, 수량", 
             wraplength=600).pack(pady=5)

    btn_import = ttk.Button(import_frame, text="엑셀 파일 가져오기", command=self.import_from_excel)
    btn_import.pack(pady=10)

    ttk.Separator(import_frame, orient='horizontal').pack(fill='x', pady=20)

    # Export Section
    ttk.Label(import_frame, text="현재 데이터 엑셀로 내보내기", font=('Arial', 11, 'bold')).pack(pady=10)

    btn_export_materials = ttk.Button(import_frame, text="자재 목록 내보내기", command=self.export_materials)
    btn_export_materials.pack(pady=5)

    btn_export_trans = ttk.Button(import_frame, text="거래 내역 내보내기", command=self.export_transactions)
    btn_export_trans.pack(pady=5)

    btn_export_all = ttk.Button(import_frame, text="전체 데이터 내보내기", command=self.export_all)
    btn_export_all.pack(pady=5)


def import_from_excel_impl(self):
    file_path = filedialog.askopenfilename(
        title="엑셀 파일 선택",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

    if not file_path:
        return

    try:
        # Load raw data
        raw_imported_df = pd.read_excel(file_path)

        # 1. Standardize column names (mapping) summerly
        col_map = {
            '품목명': ['품목명', '자재명', '품명', 'Item Name', 'Name'],
            'SN': ['SN', 'SN번호', 'Serial Number', 'S/N'],
            '규격': ['규격', 'Specification', 'Spec'],
            '수량': ['수량', 'Initial Stock', 'Current Stock', 'Qty', 'Amount'],
            '모델명': ['모델명', 'Model', 'Machine'],
            '창고': ['창고', 'Warehouse', 'Location'],
            '가격': ['가격', 'Price', 'Unit Price'],
            '제조사': ['제조사', 'Manufacturer', 'Brand'],
            '공급업체': ['공급업체', '공급업자', 'Supplier', 'Vendor'],
            '관리단위': ['관리단위', 'Unit', 'UOM'],
            '재고하한': ['재고하한', '재주문 수준', 'Reorder Point', 'Min Stock']
        }

        # Simple column normalization
        working_df = pd.DataFrame()
        for std_name, aliases in col_map.items():
            for alias in aliases:
                if alias in raw_imported_df.columns:
                    working_df[std_name] = raw_imported_df[alias]
                    break
            if std_name not in working_df.columns:
                if std_name in ['품목명', '수량']: # Critical columns defaults
                    if std_name == '수량': working_df[std_name] = 0
                    else: 
                        messagebox.showerror("오류", f"필수 컬럼 '{std_name}'을(를) 찾을 수 없습니다.")
                        return
                else:
                    working_df[std_name] = ""

        # Standardize values: clean strings, handle NaNs summerly
        working_df['품목명'] = working_df['품목명'].astype(str).str.strip()
        working_df['SN'] = working_df['SN'].astype(str).replace('nan', '').str.strip()
        working_df['규격'] = working_df['규격'].astype(str).replace('nan', '').str.strip()
        working_df['수량'] = pd.to_numeric(working_df['수량'], errors='coerce').fillna(0)
        working_df['가격'] = pd.to_numeric(working_df['가격'], errors='coerce').fillna(0)

        # [Added] Extract SN from Model name if necessary summerly
        def _extract_row_sn(row):
            m, s = self.extract_sn_from_model(row['모델명'], row['SN'])
            return pd.Series([m, s])

        working_df[['모델명', 'SN']] = working_df.apply(_extract_row_sn, axis=1)

        # Filter out rows without a name
        working_df = working_df[working_df['품목명'] != 'nan']
        working_df = working_df[working_df['품목명'] != '']

        # 2. Interior Deduplication: merge same items within the Excel file itself summerly
        # We group by Name, SN, and Spec.
        final_excel_df = working_df.groupby(['품목명', 'SN', '규격'], as_index=False).agg({
            '수량': 'sum',
            '가격': 'max',
            '모델명': 'first',
            '창고': 'first',
            '제조사': 'first',
            '공급업체': 'first',
            '관리단위': 'first',
            '재고하한': 'first'
        })

        # 3. Preparation: Match with existing inventory
        count_new = 0
        count_merged = 0
        count_overwritten = 0
        count_skipped = 0

        duplicate_indices = [] # (excel_row_idx, db_row_idx)
        new_items = []

        for idx, row in final_excel_df.iterrows():
            mat_name, sn, spec = row['품목명'], row['SN'], row['규격']

            existing_idx = -1
            if not self.materials_df.empty:
                mask = (self.materials_df['품목명'].astype(str) == mat_name) & \
                       (self.materials_df['SN'].astype(str).replace('nan', '') == sn) & \
                       (self.materials_df['규격'].astype(str).replace('nan', '') == spec)
                matches = self.materials_df.index[mask].tolist()
                if matches:
                    existing_idx = matches[0]

            if existing_idx != -1:
                duplicate_indices.append((idx, existing_idx))
            else:
                new_items.append(idx)

        # 4. Conflict Resolution Choice Dialog summerly
        mode = "SKIP"
        if duplicate_indices:
            choices_dlg = tk.Toplevel(self.root)
            choices_dlg.title("중복 항목 처리 선택")
            choices_dlg.geometry("450x300")
            choices_dlg.transient(self.root)
            choices_dlg.grab_set()

            selection = tk.StringVar(value="CANCEL")

            ttk.Label(choices_dlg, text=f"기존 재고 목록과 일치하는 항목 {len(duplicate_indices)}건이 발견되었습니다.", 
                      font=('Malgun Gothic', 10, 'bold')).pack(pady=15)
            ttk.Label(choices_dlg, text="어떻게 처리할까요?").pack(pady=5)

            def set_choice(c):
                selection.set(c)
                choices_dlg.destroy()

            btn_frame = ttk.Frame(choices_dlg, padding=10)
            btn_frame.pack(fill='both', expand=True)

            ttk.Button(btn_frame, text="기존 재고에 수량 합치기 (추천)", 
                       command=lambda: set_choice("MERGE"), width=40).pack(pady=5)
            ttk.Button(btn_frame, text="기존 정보 덮어쓰기 (Excel 정보로 교체)", 
                       command=lambda: set_choice("OVERWRITE"), width=40).pack(pady=5)
            ttk.Button(btn_frame, text="중복 항목 건너뛰기 (신규만 추가)", 
                       command=lambda: set_choice("SKIP"), width=40).pack(pady=5)
            ttk.Button(btn_frame, text="가져오기 취소", 
                       command=lambda: set_choice("CANCEL"), width=40).pack(pady=15)

            self.root.wait_window(choices_dlg)
            mode = selection.get()
            if mode == "CANCEL":
                return

        # 5. Process Import based on Mode summerly
        old_materials = self.materials_df.copy(deep=True)
        try:
            # Add new items first
            new_rows = []
            for idx in new_items:
                row_data = final_excel_df.iloc[idx].to_dict()
                valid_ids = pd.to_numeric(self.materials_df['MaterialID'], errors='coerce').dropna() if not self.materials_df.empty else pd.Series(dtype=float)
                base_max = int(valid_ids.max()) if not valid_ids.empty else 0
                new_mat_id = base_max + len(new_rows) + 1
                row_data['MaterialID'] = new_mat_id
                row_data['Active'] = 1
                new_rows.append(row_data)
                count_new += 1

            if new_rows:
                self.materials_df = pd.concat([self.materials_df, pd.DataFrame(new_rows)], ignore_index=True)

            # Handle duplicates
            if mode != "SKIP":
                for ex_idx, db_idx in duplicate_indices:
                    ex_row = final_excel_df.iloc[ex_idx]

                    if mode == "MERGE":
                        # Add to existing quantity
                        current_qty = pd.to_numeric(self.materials_df.at[db_idx, '수량'], errors='coerce')
                        if pd.isna(current_qty): current_qty = 0
                        self.materials_df.at[db_idx, '수량'] = current_qty + ex_row['수량']
                        # Also update Manufacturer/Supplier if empty in DB but present in Excel
                        for field in ['제조사', '공급업체', '규격', '창고']:
                            if field in self.materials_df.columns and field in ex_row:
                                if not self.materials_df.at[db_idx, field] or str(self.materials_df.at[db_idx, field]) == 'nan':
                                    self.materials_df.at[db_idx, field] = ex_row[field]
                        count_merged += 1

                    elif mode == "OVERWRITE":
                        # Full replace (except ID and Active status)
                        for col in final_excel_df.columns:
                            if col in self.materials_df.columns:
                                self.materials_df.at[db_idx, col] = ex_row[col]
                        count_overwritten += 1
            else:
                count_skipped = len(duplicate_indices)

            # 6. Finalize summerly
            if not self.save_data(require_primary=True):
                raise RuntimeError('원본 재고 파일에 가져온 데이터를 저장하지 못했습니다.')
        except Exception:
            self.materials_df = old_materials
            raise
        self.update_material_combo()
        self.update_stock_view()
        self.update_registration_combos()

        summary = f"자재 가져오기가 완료되었습니다.\n\n"
        summary += f"• 신규 등록: {count_new}건\n"
        if mode == "MERGE":
            summary += f"• 재고량 합산: {count_merged}건\n"
        elif mode == "OVERWRITE":
            summary += f"• 정보 덮어쓰기: {count_overwritten}건\n"
        elif mode == "SKIP":
            summary += f"• 중복 건너뜀: {count_skipped}건\n"

        messagebox.showinfo("가져오기 완료", summary)

    except Exception as e:
        import traceback
        print(traceback.format_exc())
        messagebox.showerror("오류", f"파일을 가져오는데 실패했습니다: {e}")

    except Exception as e:
        messagebox.showerror("오류", f"파일을 가져오는데 실패했습니다: {e}")


def view_monthly_usage_impl(self):
    """Display monthly usage by item name in the treeview"""
    # Clear current view
    for item in self.usage_tree.get_children():
        self.usage_tree.delete(item)

    year = int(self.cb_year.get())
    month = int(self.cb_month.get())

    # Filter transactions for the selected month
    month_mask = (self.transactions_df['Date'].dt.year == year) & \
                 (self.transactions_df['Date'].dt.month == month) & \
                 (self.transactions_df['Type'] == 'OUT')
    monthly_trans = self.transactions_df[month_mask]

    # Filter transactions for cumulative (from start of year to selected month)
    cumulative_mask = (self.transactions_df['Date'].dt.year == year) & \
                      (self.transactions_df['Date'].dt.month <= month) & \
                      (self.transactions_df['Type'] == 'OUT')
    cumulative_trans = self.transactions_df[cumulative_mask]

    # Build usage data for each material
    usage_data = []
    for _, mat in self.materials_df.iterrows():
        mat_id = mat['MaterialID']

        # Calculate monthly usage
        month_usage = monthly_trans[monthly_trans['MaterialID'] == mat_id]['Quantity'].sum()

        # Calculate cumulative usage (year-to-date)
        cumulative_usage = cumulative_trans[cumulative_trans['MaterialID'] == mat_id]['Quantity'].sum()

        # Only show items with usage
        if month_usage > 0 or cumulative_usage > 0:
            usage_data.append({
                '품목명': mat.get('품목명', ''),  # Using 재고현황's 품목명 field
                '관리품번': mat.get('관리품번', ''),
                '규격': mat.get('규격', ''),
                '단위': mat.get('관리단위', 'EA'),
                '월사용량': month_usage,
                '누계사용량': cumulative_usage
            })

    # Sort by item name
    usage_data.sort(key=lambda x: x['품목명'])

    # Display in treeview
    for data in usage_data:
        self.usage_tree.insert('', tk.END, values=(
            data['품목명'],
            data['관리품번'],
            data['규격'],
            data['단위'],
            f"{data['월사용량']:.1f}",
            f"{data['누계사용량']:.1f}"
        ))

    # Show message if no data
    if not usage_data:
        messagebox.showinfo("알림", f"{year}년 {month}월에 사용 내역이 없습니다.")


def generate_yearly_report_impl(self):
    year = int(self.cb_year.get())
    # Filter transactions for the year
    mask = (self.transactions_df['Date'].dt.year == year) & (self.transactions_df['Type'] == 'OUT')
    yearly_trans = self.transactions_df[mask]

    report = []
    for _, mat in self.materials_df.iterrows():
        mat_id = mat['MaterialID']

        # Calculate monthly usage for each month
        row_data = {
            '설비코드': mat.get('관리품번', ''),
            '자재명': mat.get('품목명', ''),  # Using 재고현황's 품목명 field
            '분류': mat.get('품목군코드', ''),
            '규격': mat.get('규격', ''),
            '단위': mat.get('관리단위', ''),
            '제조사': mat.get('제조사', '')
        }

        # Add monthly columns (1월 ~ 12월)
        monthly_values = []
        for month in range(1, 13):
            month_mask = (yearly_trans['MaterialID'] == mat_id) & \
                        (yearly_trans['Date'].dt.month == month)
            month_usage = yearly_trans[month_mask]['Quantity'].sum()
            row_data[f'{month}월'] = month_usage
            monthly_values.append(month_usage)

        # Calculate totals
        total = sum(monthly_values)
        row_data['합계'] = total

        # Calculate cumulative total
        cumulative = 0
        for i, val in enumerate(monthly_values, 1):
            cumulative += val
            if i == 12:  # Only show final cumulative at the end
                row_data['누계'] = cumulative

        report.append(row_data)

    report_df = pd.DataFrame(report)
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                             initialfile=f"Yearly_Usage_{year}.xlsx",
                                             title="보고서 저장")
    if save_path:
        report_df.to_excel(save_path, index=False)
        messagebox.showinfo("완료", f"{year}년 보고서가 저장되었습니다.")


def generate_monthly_report_impl(self):
    year = int(self.cb_year.get())
    month = int(self.cb_month.get())

    mask = (self.transactions_df['Date'].dt.year == year) & \
           (self.transactions_df['Date'].dt.month == month) & \
           (self.transactions_df['Type'] == 'OUT')
    monthly_trans = self.transactions_df[mask]

    report = []
    for _, mat in self.materials_df.iterrows():
        mat_id = mat['MaterialID']
        total_usage = monthly_trans[monthly_trans['MaterialID'] == mat_id]['Quantity'].sum()
        report.append({
            '설비코드': mat.get('관리품번', ''),
            '자재명': mat.get('품목명', ''),  # Using 재고현황's 품목명 field
            '분류': mat.get('품목군코드', ''),
            '규격': mat.get('규격', ''),
            '단위': mat.get('관리단위', ''),
            '제조사': mat.get('제조사', ''),
            '월간 총 사용량': total_usage
        })

    report_df = pd.DataFrame(report)
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                             initialfile=f"Monthly_Usage_{year}_{month}.xlsx",
                                             title="보고서 저장")
    if save_path:
        report_df.to_excel(save_path, index=False)
        messagebox.showinfo("완료", f"{year}년 {month}월 보고서가 저장되었습니다.")

