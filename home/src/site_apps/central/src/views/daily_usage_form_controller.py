# -*- coding: utf-8 -*-
import os
import re
import sys
import json
import math
import time
import shutil
import datetime
import traceback
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import pandas as pd
import numpy as np
from tkcalendar import DateEntry, Calendar
from site_apps.central.src.utils.helpers import NAN_PATTERN, DOT_ZERO_PATTERN, MARKER_PATTERN
from site_apps.central.src.views.components import *


def setup_daily_usage_query_tab_impl(self):
    """Setup the daily usage query tab"""
    display_frame = ttk.Frame(self.tab_daily_usage_query)
    display_frame.pack(fill="both", expand=True, padx=5, pady=5)

    # KPI Summary Frame
    self.kpi_frame = ttk.LabelFrame(display_frame, text="조회 기간 요약 (KPI)")
    self.kpi_frame.pack(fill='x', padx=5, pady=(0, 5))
    self.lbl_kpi_summary = ttk.Label(self.kpi_frame, text="데이터를 불러오는 중...", font=('Malgun Gothic', 10, 'bold'), foreground='#00529B')
    self.lbl_kpi_summary.pack(side='left', padx=10, pady=5)

    # Filter controls
    filter_frame = ttk.Frame(display_frame)
    filter_frame.pack(fill='x', padx=5, pady=0)

    # --- Row 1: Date Filters ---
    date_row = ttk.Frame(filter_frame)
    date_row.pack(fill='x', pady=2)

    ttk.Label(date_row, text="시작일:").pack(side='left', padx=5)
    self.ent_daily_start_date = DateEntry(date_row, width=12, date_pattern='yyyy-mm-dd', locale='ko_KR', state='readonly', showweeknumbers=True)
    self.ent_daily_start_date.pack(side='left', padx=5)
    # Default to showing all history (starting from 2024)
    start_date = datetime.datetime(2024, 1, 1)
    self.ent_daily_start_date.set_date(start_date)

    ttk.Label(date_row, text="종료일:").pack(side='left', padx=5)
    self.ent_daily_end_date = DateEntry(date_row, width=12, date_pattern='yyyy-mm-dd', locale='ko_KR', state='readonly', showweeknumbers=True)
    self.ent_daily_end_date.pack(side='left', padx=5)
    self.ent_daily_end_date.set_date(datetime.datetime.now())

    # --- Row 2: Search Filters ---
    filter_row = ttk.Frame(filter_frame)
    filter_row.pack(fill='x', pady=2)

    ttk.Label(filter_row, text="업체명:").pack(side='left', padx=(5, 2))
    self.cb_daily_filter_company = ttk.Combobox(filter_row, width=12)
    self.cb_daily_filter_company.pack(side='left', padx=2)
    self.cb_daily_filter_company.set('전체')

    def _on_filter_company_changed(event=None):
        comp = self.cb_daily_filter_company.get().strip()
        if not comp or comp == '전체':
            sites = ['전체'] + sorted(self.sites)
        elif hasattr(self, 'daily_usage_df') and not self.daily_usage_df.empty:
            co_col = '업체명' if '업체명' in self.daily_usage_df.columns else ('Company' if 'Company' in self.daily_usage_df.columns else None)
            site_col = '현장' if '현장' in self.daily_usage_df.columns else ('Site' if 'Site' in self.daily_usage_df.columns else None)
            if co_col and site_col:
                matched = sorted(set(self.daily_usage_df[self.daily_usage_df[co_col] == comp][site_col].dropna().astype(str).str.strip().unique()))
                matched = [s for s in matched if s and s != 'nan']
                sites = ['전체'] + (matched or sorted(self.sites))
            else:
                sites = ['전체'] + sorted(self.sites)
        else:
            sites = ['전체'] + sorted(self.sites)
        self.cb_daily_filter_site['values'] = sites
        site_get = getattr(self.cb_daily_filter_site, 'get', None)
        if site_get and site_get() not in sites:
            self.cb_daily_filter_site.set('전체')
        self.update_daily_usage_view()

    self.cb_daily_filter_company.bind("<<ComboboxSelected>>", _on_filter_company_changed)
    tk.Button(filter_row, text="⚙️", font=('Malgun Gothic', 8), bd=0, bg=self.theme_bg, fg='blue', cursor='hand2',
              command=lambda: self.open_list_management_dialog('companies', target_cb=self.cb_daily_filter_company)).pack(side='left', padx=(0, 5))

    ttk.Label(filter_row, text="현장:").pack(side='left', padx=(5, 2))
    self.cb_daily_filter_site = ttk.Combobox(filter_row, width=12)
    self.cb_daily_filter_site.pack(side='left', padx=2)
    self.cb_daily_filter_site.set('전체')
    tk.Button(filter_row, text="⚙️", font=('Malgun Gothic', 8), bd=0, bg=self.theme_bg, fg='blue', cursor='hand2',
              command=lambda: self.open_list_management_dialog('sites', target_cb=self.cb_daily_filter_site)).pack(side='left', padx=(0, 5))

    ttk.Label(filter_row, text="자재명:").pack(side='left', padx=(5, 2))
    self.cb_daily_filter_material = ttk.Combobox(filter_row, width=15)
    self.cb_daily_filter_material.pack(side='left', padx=2)
    self.cb_daily_filter_material.set('전체')
    tk.Button(filter_row, text="⚙️", font=('Malgun Gothic', 8), bd=0, bg=self.theme_bg, fg='blue', cursor='hand2',
              command=lambda: self.open_list_management_dialog('materials', target_cb=self.cb_daily_filter_material)).pack(side='left', padx=(0, 5))

    ttk.Label(filter_row, text="장비명:").pack(side='left', padx=(5, 2))
    self.cb_daily_filter_equipment = ttk.Combobox(filter_row, width=15)
    self.cb_daily_filter_equipment.pack(side='left', padx=2)
    self.cb_daily_filter_equipment.set('전체')
    tk.Button(filter_row, text="⚙️", font=('Malgun Gothic', 8), bd=0, bg=self.theme_bg, fg='blue', cursor='hand2',
              command=lambda: self.open_list_management_dialog('equipments', target_cb=self.cb_daily_filter_equipment)).pack(side='left', padx=(0, 5))

    ttk.Label(filter_row, text="작업자:").pack(side='left', padx=(5, 2))
    self.cb_daily_filter_worker = ttk.Combobox(filter_row, width=10)
    self.cb_daily_filter_worker.pack(side='left', padx=2)
    self.cb_daily_filter_worker.set('전체')

    ttk.Label(filter_row, text="분류:").pack(side='left', padx=(5, 2))
    self.cb_daily_filter_shift = ttk.Combobox(filter_row, width=8, state="readonly", values=["전체", "주간", "야간", "주야간", "휴일"])
    self.cb_daily_filter_shift.pack(side='left', padx=2)
    self.cb_daily_filter_shift.set('전체')

    ttk.Label(filter_row, text="검사방법:").pack(side='left', padx=(5, 2))
    self.cb_daily_filter_method = ttk.Combobox(filter_row, width=8, state="readonly", values=["전체", "PAUT", "RT", "MT", "PT"])
    self.cb_daily_filter_method.pack(side='left', padx=2)
    self.cb_daily_filter_method.set('전체')

    # --- Row 3: Action Buttons (1줄 통합 툴바) ---
    btn_row = ttk.Frame(filter_frame)
    btn_row.pack(fill='x', pady=(4, 2))

    # [그룹 1: 조회 제어]
    btn_filter = ttk.Button(btn_row, text="🔍 조회", style='Primary.TButton', command=self.update_daily_usage_view)
    btn_filter.pack(side='left', padx=(5, 3))

    btn_filter_reset = ttk.Button(btn_row, text="♻️ 필터 초기화", style='Warning.TButton', command=self.reset_daily_usage_filters)
    btn_filter_reset.pack(side='left', padx=3)

    btn_col_manage = ttk.Button(btn_row, text="⚙️ 컬럼 관리", style='Secondary.TButton', command=self.show_column_visibility_dialog)
    btn_col_manage.pack(side='left', padx=(3, 8))

    # [그룹 2: 레코드 편집 및 저장]
    ttk.Button(btn_row, text="📥 입력폼에 불러오기", style='Primary.TButton',
               command=lambda: self._on_daily_usage_select(None)).pack(side='left', padx=3)
    btn_edit = ttk.Button(btn_row, text="✏️ 선택 수정", style='Info.TButton', command=self.open_edit_daily_usage_dialog)
    btn_edit.pack(side='left', padx=3)

    btn_delete = ttk.Button(btn_row, text="🗑️ 선택 삭제", style='Danger.TButton', command=self.delete_daily_usage_entry)
    btn_delete.pack(side='left', padx=3)

    self.btn_daily_save_list = ttk.Button(btn_row, text="💾 변경사항 저장", command=self.save_all_daily_usage_changes, style='Action.TButton')
    self.btn_daily_save_list.pack(side='left', padx=(3, 8))

    # [그룹 3: 데이터 내보내기 및 백업]
    btn_export = ttk.Button(btn_row, text="📊 엑셀 내보내기", style='Action.TButton', command=self.export_daily_usage_history)
    btn_export.pack(side='left', padx=3)

    btn_export_all = ttk.Button(btn_row, text="📦 전체 백업", style='Secondary.TButton', command=self.export_all_daily_usage)
    btn_export_all.pack(side='left', padx=3)

    # Bindings
    filter_widgets = [
        self.cb_daily_filter_site, self.cb_daily_filter_company, self.cb_daily_filter_material, 
        self.cb_daily_filter_equipment, self.cb_daily_filter_worker, 
        self.cb_daily_filter_shift, self.cb_daily_filter_method
    ]
    for widget in filter_widgets:
        widget.bind("<Return>", lambda e: self.update_daily_usage_view())
        widget.bind("<<ComboboxSelected>>", lambda e: self.update_daily_usage_view())

    for date_widget in [self.ent_daily_start_date, self.ent_daily_end_date]:
        date_widget.bind("<<DateEntrySelected>>", lambda e: self.update_daily_usage_view())
        try:
            date_widget.bind("<Return>", lambda e: self.update_daily_usage_view())
            for child in date_widget.winfo_children():
                if isinstance(child, (tk.Entry, ttk.Entry)):
                    child.bind("<Return>", lambda e: self.update_daily_usage_view())
        except: pass

    # Treeview for daily usage records
    tree_container = ttk.Frame(display_frame)
    tree_container.pack(expand=True, fill='both', padx=5, pady=5)

    list_frame = ttk.Frame(tree_container)
    list_frame.pack(fill='both', expand=True)

    # Scrollbars
    vsb = ttk.Scrollbar(list_frame, orient="vertical")
    hsb = ttk.Scrollbar(list_frame, orient="horizontal")

    # Treeview with RTK categories and NDT materials
    # Note: Workers 1-10 columns are kept in the 'columns' tuple for data storage,
    # but we will only show a consolidated '작업자' in 'displaycolumns'.
    # Added '(Full작업자)' for Excel export backup.
    columns = ('날짜', '업체명', '적용코드', '현장', '구분', '검사품명', '성적서번호', '작업자', '작업시간', 'OT1', 'OT2', 'OT3', 'OT4', 'OT5', 'OT6', 'OT7', 'OT8', 'OT9', 'OT10', '장비명', '검사방법', '회사코드', '수량', '단위', '단가', '출장비', '일식', '검사비', 'OT시간', 'OT금액', '품목명', '센터미스', '농도', '마킹미스', '필름마크', '취급부주의', '고객불만', '기타', 'RTK총계', '형광자분', '흑색자분', '백색페인트', '침투제', '세척제', '현상제', '형광침투제', '비고', '입력시간', '차량번호', '주행거리', '차량점검', '차량비고', '(Full작업자)')
    columns += ('메모',)
    self.daily_usage_tree = ttk.Treeview(list_frame, columns=columns, show='headings',
                                          yscrollcommand=vsb.set, xscrollcommand=hsb.set)

    # Define display columns - 목적에 최적화된 추천 기본 컬럼 (원가 세부항목 대신 단가·검사금액으로 통합)
    visible_defaults = ['날짜', '메모', '구분', '검사방법', '품목명', '수량', '단위', '단가', '검사비', '작업자', '작업시간', 'OT금액', '장비명', '업체명', '비고']
    self.daily_usage_tree['displaycolumns'] = visible_defaults

    # [NEW NDT COLUMNS] Add dynamically
    new_cols = ['작업형태', '조건1', '조건2', '보정계수', '제경비', '기술료', '환산물량', '재료비', '인건비', '검사구분', '조인트수', '불량수', '관경(Inch)', '핫태핑차수', '핫태핑개소', '핫태핑관경']
    new_cols += ['보조장비', '원본매수', '재검사매수', '참여인원', '참여시간', '참여OT']
    c_list = list(columns)
    for nc in new_cols:
        if nc not in c_list:
            c_list.append(nc)
    columns = tuple(c_list)

    v_list = list(visible_defaults)

    self.daily_usage_tree.configure(columns=columns)
    self.daily_usage_tree['displaycolumns'] = v_list

    for nc in new_cols:
        self.daily_usage_tree.heading(nc, text=("자재명" if nc == "품목명" else nc), command=lambda c=nc: self.treeview_sort_column(self.daily_usage_tree, c, False))
        self.daily_usage_tree.column(nc, width=100, minwidth=20, stretch=False, anchor='center')


    vsb.config(command=self.daily_usage_tree.yview)
    hsb.config(command=self.daily_usage_tree.xview)

    # Column configuration
    col_widths = {
        '날짜': 105, '메모': 55, '구분': 85, '검사방법': 75, '수량': 70, '단위': 45,
        '단가': 85, '검사비': 105,
        '작업자': 150, '작업시간': 130, 'OT금액': 90, '장비명': 130, '업체명': 120, '비고': 180,
        '현장': 90, '검사품명': 100, '적용코드': 90, '성적서번호': 110,
        'OT시간': 75, '출장비': 85, '일식': 75,
        '차량번호': 110, '주행거리': 85, '차량점검': 150, '차량비고': 130,
        '품목명': 160, '입력시간': 150,
        'OT1': 90, 'OT2': 90, 'OT3': 90, 'OT4': 90, 'OT5': 90,
        'OT6': 90, 'OT7': 90, 'OT8': 90, 'OT9': 90, 'OT10': 90,
    }

    for col in columns:
        disp_title = "자재명" if col == "품목명" else ("검사금액" if col == "검사비" else col)
        self.daily_usage_tree.heading(col, text=disp_title, command=lambda c=col: self.treeview_sort_column(self.daily_usage_tree, c, False))
        width = col_widths.get(col, 100)
        anchor = 'e' if col in ('수량', '단가', '검사비', 'OT금액', '출장비', '일식') else 'center'
        self.daily_usage_tree.column(col, width=width, minwidth=20, stretch=False, anchor=anchor)

    # Grid layout for list_frame
    self.daily_usage_tree.grid(row=0, column=0, sticky='nsew')
    vsb.grid(row=0, column=1, sticky='ns')
    hsb.grid(row=1, column=0, sticky='ew')

    list_frame.grid_rowconfigure(0, weight=1)
    list_frame.grid_columnconfigure(0, weight=1)

    # Auto-save column widths when user resizes columns
    # Worker details remain available on a single click; double-click opens notes.
    self.daily_usage_tree.bind("<Button-1>", lambda e: self.show_worker_popup(e, self.daily_usage_tree), add="+")
    self.daily_usage_tree.bind("<Double-1>", self._show_daily_usage_note)

    # [NEW] Enable column reordering via drag & drop and context menu
    def _show_daily_usage_tree_context_menu(event):
        region = self.daily_usage_tree.identify_region(event.x, event.y)
        if region == "heading":
            self._show_heading_context_menu(event, self.daily_usage_tree)
            return
        item = self.daily_usage_tree.identify_row(event.y)
        if item:
            self.daily_usage_tree.selection_set(item)
            menu = tk.Menu(self.root, tearoff=0)
            menu.add_command(label="📥 입력 화면에 불러오기", command=lambda: self._on_daily_usage_select(None))
            menu.add_command(label="✏️ 선택 항목 수정...", command=self.open_edit_daily_usage_dialog)
            menu.add_command(label="🗑️ 선택 항목 삭제", command=self.delete_daily_usage_entry)
            menu.add_separator()
            menu.add_command(label="🔍 조회 (새로고침)", command=self.update_daily_usage_view)
            menu.add_command(label="📊 조회결과 엑셀 내보내기", command=self.export_daily_usage_history)
            menu.add_separator()
            menu.add_command(label="⚙️ 컬럼 관리(숨기기/보이기)...", command=self.show_column_visibility_dialog)
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()

    self.enable_tree_column_drag(self.daily_usage_tree, context_menu_handler=_show_daily_usage_tree_context_menu)

    def save_column_widths(event=None):
        self.save_tab_config()

    self.daily_usage_tree.bind('<ButtonRelease-1>', save_column_widths)

    # Selection and context menus must not replace unfinished entry data.
    # Loading is explicit; keep double-click available for the note viewer.
    self.daily_usage_tree.bind('<Return>', self._on_daily_usage_select)

    # Optimize treeview scroll region to prevent unnecessary scrollbars
    def optimize_treeview_scroll_region():
        try:
            if hasattr(self, 'daily_usage_tree'):
                self.daily_usage_tree.update_idletasks()

                # Get treeview content dimensions
                total_items = len(self.daily_usage_tree.get_children())
                if total_items > 0:
                    # Calculate approximate content height
                    item_height = 25  # Approximate height per row
                    content_height = total_items * item_height + 50  # Add some padding

                    # Get treeview dimensions
                    tree_width = self.daily_usage_tree.winfo_width()
                    tree_height = self.daily_usage_tree.winfo_height()

                    # Set scroll region to content size if smaller than treeview
                    if content_height < tree_height:
                        self.daily_usage_tree.configure(yscrollcommand=(0, 0, tree_width, content_height))
                        print(f"Treeview scroll region optimized: (0, 0, {tree_width}, {content_height})")
                    else:
                        # Use full treeview dimensions
                        self.daily_usage_tree.configure(yscrollcommand=(0, 0, tree_width, tree_height))
                        print(f"Treeview scroll region set to full size: (0, 0, {tree_width}, {tree_height})")
        except Exception as e:
            print(f"Error optimizing treeview scroll region: {e}")

    # Apply optimization after a short delay
    self.root.after(500, optimize_treeview_scroll_region)


def _show_daily_usage_note_impl(self, event):
    """Display the clicked record's full note without opening the editor."""
    tree = self.daily_usage_tree
    if tree.identify_region(event.x, event.y) not in ('cell', 'tree'):
        return
    item = tree.identify_row(event.y)
    if not item:
        return
    tags = tree.item(item, 'tags')
    if not tags or 'total' in tags:
        return 'break'
    try:
        entry = self.daily_usage_df.loc[int(tags[0])]
    except (ValueError, KeyError, IndexError):
        return 'break'
    from site_apps.central.src.views.memo_details import daily_memo_lookup, record_notes
    note, daily_memo = record_notes(entry, daily_memo_lookup(self.daily_usage_df))
    details = f"[기록 비고]\n{note or '등록된 비고가 없습니다.'}\n\n[날짜별 메모]\n{daily_memo or '등록된 메모가 없습니다.'}"
    popup = tk.Toplevel(self.root)
    popup.title('비고 · 날짜별 메모')
    popup.geometry('650x320')
    popup.transient(self.root)
    frame = ttk.Frame(popup, padding=10)
    frame.pack(fill='both', expand=True)
    text = tk.Text(frame, wrap='word', font=('Malgun Gothic', 11))
    scrollbar = ttk.Scrollbar(frame, orient='vertical', command=text.yview)
    text.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side='right', fill='y')
    text.pack(fill='both', expand=True)
    text.insert('1.0', details)
    text.configure(state='disabled')
    ttk.Button(popup, text='닫기', command=popup.destroy).pack(pady=(0, 10))
    popup.bind('<Escape>', lambda e: popup.destroy())
    text.focus_set()
    return 'break'


def add_ndt_company_section_impl(self):
    """회사별 NDT 입력 섹션 추가"""
    if not hasattr(self, 'ndt_company_entries'):
        self.ndt_company_entries = []

    company_idx = len(self.ndt_company_entries)
    ndt_materials = self.ndt_materials_all



    # Company frame (Standard frame to save space, no title border)
    company_frame = ttk.Frame(self.ndt_company_container)
    company_frame.pack(fill='x', padx=2, pady=2)

    # Single Grid frame for everything to save vertical space
    grid_frame = ttk.Frame(company_frame)
    for c in range(8): grid_frame.columnconfigure(c, weight=1, uniform="ndt_rtk")
    grid_frame.pack(fill='x', padx=2, pady=2)

    # Company code selector (Index 0)
    ttk.Label(grid_frame, text="회사코드:", font=('Arial', 8, 'bold')).grid(row=0, column=0, padx=1, pady=1, sticky='e')
    cb_co = ttk.Combobox(grid_frame, width=8, values=getattr(self, 'co_code_list', []))
    cb_co.grid(row=0, column=1, padx=1, pady=1, sticky='ew')
    cb_co.set('')  # Default empty

    entries = {'_company': cb_co}  # Store company combobox

    # NDT entries grid (Index 1 to 7)
    for i, mat in enumerate(ndt_materials):
        idx = i + 1
        r = idx // 4
        c = (idx % 4) * 2
        ttk.Label(grid_frame, text=f"{mat}:", font=('Arial', 8)).grid(row=r, column=c, padx=1, pady=1, sticky='e')
        e = ttk.Entry(grid_frame, width=6)
        e.grid(row=r, column=c+1, padx=1, pady=1, sticky='ew')
        entries[mat] = e

    self.ndt_company_entries.append(entries)


def remove_last_ndt_company_impl(self):
    """마지막 회사 섹션 삭제"""
    if hasattr(self, 'ndt_company_entries') and len(self.ndt_company_entries) > 1:
        # Remove last entry dict
        self.ndt_company_entries.pop()
        # Destroy last widget in container
        widgets = self.ndt_company_container.winfo_children()
        if widgets:
            widgets[-1].destroy()


def _load_ndt_product_map_impl(self):
    """config 에서 NDT 약품 → 실제 DB 품목명 매핑 불러오기"""
    try:
        if os.path.exists(self.config_path):
            with open(self.config_path, 'r', encoding='utf-8') as f:
                cfg = json.load(f)
            return cfg.get('ndt_product_map', {})
    except Exception:
        pass
    return {}


def _save_ndt_product_map_impl(self, map_data):
    """NDT 약품 -> 실제 DB 품목명 매핑 저장"""
    try:
        if os.path.exists(self.config_path):
            with open(self.config_path, 'r', encoding='utf-8') as f:
                cfg = json.load(f)
        else:
            cfg = {}
        cfg['ndt_product_map'] = map_data
        with open(self.config_path, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, indent=4, ensure_ascii=False)
        return True
    except Exception:
        return False


def open_ndt_product_map_dialog_impl(self):
    """NDT 약품 -> 실제 DB 품목명 매핑 설정 다이얼로그"""
    dlg = tk.Toplevel(self.root)
    dlg.title("NDT 약품-품목 매핑 설정")
    dlg.geometry("500x600")
    dlg.transient(self.root)
    dlg.grab_set()

    main_frame = ttk.Frame(dlg, padding=20)
    main_frame.pack(fill='both', expand=True)

    ttk.Label(main_frame, text="현장 입력 약품명", font=('Arial', 10, 'bold')).grid(row=0, column=0, pady=10)
    ttk.Label(main_frame, text="창고 재고 품목 (매핑)", font=('Arial', 10, 'bold')).grid(row=0, column=1, pady=10)

    ndt_materials = ["형광자분", "흑색자분", "백색페인트", "침투제", "세척제", "현상제", "형광침투제"]
    current_map = self._load_ndt_product_map()

    material_options = []
    for _, row in self.materials_df.iterrows():
        disp = self.get_material_display_name(row['MaterialID'])
        material_options.append(disp)
    material_options.sort()

    combos = {}
    for i, mat in enumerate(ndt_materials):
        ttk.Label(main_frame, text=mat).grid(row=i+1, column=0, padx=5, pady=5, sticky='w')
        cb = ttk.Combobox(main_frame, values=material_options, width=40)
        cb.grid(row=i+1, column=1, padx=5, pady=5, sticky='ew')
        current_id = current_map.get(mat, "")
        if current_id:
            cb.set(self.get_material_display_name(current_id))
        combos[mat] = cb

    def _save():
        new_map = {}
        for mat, cb in combos.items():
            disp = cb.get().strip()
            if disp:
                for _, row in self.materials_df.iterrows():
                    if self.get_material_display_name(row['MaterialID']) == disp:
                        new_map[mat] = row['MaterialID']
                        break
        if self._save_ndt_product_map(new_map):
            messagebox.showinfo("성공", "매핑 설정이 저장되었습니다.")
            dlg.destroy()
        else:
            messagebox.showerror("오류", "설정을 저장하지 못했습니다.")

    def _clear():
        for cb in combos.values():
            cb.set('')

    btn_frame = ttk.Frame(main_frame)
    btn_frame.grid(row=len(ndt_materials)+1, column=0, columnspan=2, pady=20)

    ttk.Button(btn_frame, text="저장", command=_save, width=10).pack(side='left', padx=8)
    ttk.Button(btn_frame, text="전체 초기화", command=_clear, width=10).pack(side='left', padx=8)
    ttk.Button(btn_frame, text="닫기", command=dlg.destroy, width=10).pack(side='left', padx=8)


def _get_merged_memo_and_note_impl(self):
    """[NEW] Collects user-entered note and all open memo contents for archiving"""
    main_note = ""
    if hasattr(self, 'ent_daily_note'):
        main_note = self.ent_daily_note.get().strip()

    memo_archive = []
    if hasattr(self, 'memos'):
        for key, m in self.memos.items():
            try:
                title = m['title_entry'].get().strip()
                content = m['text_widget'].get('1.0', 'end-1c').strip()
                if content:
                    memo_archive.append(content)
            except: continue

    if not memo_archive:
        return main_note

    archive_text = " | ".join(memo_archive)
    return f"{main_note} | {archive_text}" if main_note else archive_text


def load_daily_usage_to_form_impl(self, record):
    """[NEW] Load a database record back into the input form for editing or report generation"""
    from site_apps.central.src.inspection_batch import participation_record
    record = participation_record(record)
    try:
        def to_f_local(v):
            try: return float(str(v).replace(',', '')) if v else 0.0
            except: return 0.0

        # 1. Basic Info
        if 'Date' in record:
            try: self.ent_daily_date.set_date(pd.to_datetime(record['Date']))
            except: pass
        self._last_daily_date = self.ent_daily_date.get_date()
        if hasattr(self, 'cb_daily_client'):
            client_val = self.clean_nan(record.get('회사코드', '')).strip()
            valid_clients = ['한국지역난방공사', '한국지역난방공사 중앙지사']
            if hasattr(self.cb_daily_client, 'cget'):
                try:
                    vals = list(self.cb_daily_client.cget('values'))
                    if vals:
                        valid_clients = vals
                except: pass
            if client_val in valid_clients:
                self.cb_daily_client.set(client_val)
            elif '한국지역난방' in client_val:
                self.cb_daily_client.set(client_val)
            else:
                self.cb_daily_client.set('한국지역난방공사')
        if '업체명' in record:
            comp_val = self.clean_nan(record['업체명']).split(' [')[0]
            self.cb_daily_company.set(comp_val)
        if 'Site' in record:
            self.cb_daily_site.set(self.clean_nan(record['Site']))
        if '구분' in record:
            self.ndt_loc_type_var.set(self.clean_nan(record['구분']))
        if '적용코드' in record:
            self.ent_daily_applied_code.delete(0, tk.END)
            self.ent_daily_applied_code.insert(0, self.clean_nan(record['적용코드']))
        from site_apps.central.src.aux_equipment import restore
        restore(self, record.get('보조장비', ''))
        if '장비명' in record:
            self.cb_daily_equip.delete(0, tk.END)
            self.cb_daily_equip.insert(0, self.clean_nan(record['장비명']))
        if '성적서번호' in record:
            self.ent_daily_report_no.delete(0, tk.END)
            self.ent_daily_report_no.insert(0, self.clean_nan(record['성적서번호']))
        if '검사품명' in record:
            self.ent_daily_inspection_item.delete(0, tk.END)
            self.ent_daily_inspection_item.insert(0, self.clean_nan(record['검사품명']))
        if '검사방법' in record:
            method = self.clean_nan(record.get('검사방법', ''))
            self.cb_daily_test_method.set(method)

            # [FIX] Explicitly show/hide NDT related frames to ensure they appear
            if method in ["MT", "PT"]:
                if hasattr(self, 'ndt_frame'):
                    try: self.ndt_frame.grid()
                    except: pass
            else:
                if hasattr(self, 'ndt_frame'):
                    try: self.ndt_frame.grid_remove()
                    except: pass

            if method in ["RT", "UT", "PT", "PAUT"]:
                if hasattr(self, 'ndt_calc_frame'):
                    try:
                        self.ndt_calc_frame.grid()
                        self.ndt_calc_frame.lift()
                    except: pass
                if method == "RT":
                    if hasattr(self, 'rtk_grid'):
                        try: self.rtk_grid.grid()
                        except: pass
                else:
                    if hasattr(self, 'rtk_grid'):
                        try: self.rtk_grid.grid_remove()
                        except: pass
            else:
                if hasattr(self, 'ndt_calc_frame'):
                    try: self.ndt_calc_frame.grid_remove()
                    except: pass
                if hasattr(self, 'rtk_grid'):
                    try: self.rtk_grid.grid_remove()
                    except: pass

            self.cb_daily_test_method.event_generate('<<ComboboxSelected>>')
        if 'Unit' in record:
            self.cb_daily_unit.set(self.clean_nan(record['Unit']))
        elif '단위' in record:
            self.cb_daily_unit.set(self.clean_nan(record['단위']))

        self.hot_tap_stage_var.set(self.clean_nan(record.get('핫태핑차수', '')) or '1차')
        self.hot_tap_size_var.set(self.clean_nan(record.get('핫태핑관경', '')) or '600A')
        self.hot_tap_location_var.set(self.clean_nan(record.get('핫태핑개소', '')))
        # 2. Quantities & Costs
        def set_val(ent, key):
            val = record.get(key, 0)
            ent.delete(0, tk.END)
            if isinstance(val, (int, float)):
                if val != 0: ent.insert(0, f"{val:,.0f}" if any(x in key for x in ['비', '가', '일식']) else str(val))
            else: 
                ent.insert(0, str(val))

        set_val(self.ent_daily_test_amount, '검사량')
        set_val(self.ent_daily_unit_price, '단가')
        set_val(self.ent_daily_travel_cost, '출장비')
        set_val(self.ent_daily_test_fee, '검사비')
        set_val(self.ent_daily_meal_cost, '일식')
        if '작업형태' in record and record['작업형태']: self.ndt_work_time_var.set(record['작업형태'])
        if '조건1' in record and record['조건1']: 
            if record.get('검사방법') == 'RT': self.ndt_source_var.set(record['조건1'])
            else: self.ndt_pipe_var.set(record['조건1'])
        if '조건2' in record and record['조건2']: self.ndt_thickness_var.set(record['조건2'])
        def get_valid_rate(key, default_val):
            if key in record:
                val = record[key]
                if pd.notna(val) and str(val).strip() and str(val).lower() != 'nan':
                    try: return float(val)
                    except: pass
            return default_val

        self.ndt_overhead_var.set(get_valid_rate('제경비율', 110.0))
        self.ndt_tech_var.set(get_valid_rate('기술료율', 20.0))

        # NDT 상세 조건 필드 로딩
        if '관경(Inch)' in record and pd.notna(record['관경(Inch)']):
            self.ndt_report_pipe_var.set(str(record['관경(Inch)']).replace('.0', '') if str(record['관경(Inch)']).endswith('.0') else str(record['관경(Inch)']))
        else:
            self.ndt_report_pipe_var.set("")

        insp_type = self.clean_nan(record.get('검사구분', 'ORI')).upper()
        if not insp_type: insp_type = 'ORI'
        if insp_type == 'REP':
            self.ndt_rep_joint_var.set(self.clean_nan(record.get('조인트수', '')))
            self.ndt_rep_qty_var.set(self.clean_nan(record.get('재검사매수', record.get('Usage', ''))))
        else:
            self.ndt_ori_joint_var.set(self.clean_nan(record.get('조인트수', '')))
            self.ndt_ori_qty_var.set(self.clean_nan(record.get('원본매수', record.get('Usage', ''))))

        if '불량수' in record and pd.notna(record['불량수']):
            self.ndt_rej_joint_var.set(str(record['불량수']))
        else:
            self.ndt_rej_joint_var.set("")

        # 3. Material
        if 'MaterialID' in record:
            mat_id_val = record['MaterialID']
            disp_name = self.get_material_display_name(mat_id_val)
            # [FIX] If display name is empty but MaterialID is a non-NaN string (e.g. 'JIREH Scanner'),
            # show the raw MaterialID string as a fallback
            if not disp_name:
                try:
                    if not pd.isna(mat_id_val) and str(mat_id_val).strip().lower() not in ('nan', ''):
                        disp_name = str(mat_id_val).strip()
                except: pass
            print(f"[DEBUG] load_to_form: MaterialID={repr(mat_id_val)}, disp={repr(disp_name)}")
            if isinstance(self.cb_daily_material, ttk.Combobox):
                self.cb_daily_material.set(disp_name)
            else:
                self.cb_daily_material.delete(0, tk.END)
                self.cb_daily_material.insert(0, disp_name)
            from site_apps.central.src.material_selection import remember_selection
            remember_selection(self, mat_id_val)

        # 4. Workers
        # First, clear existing workers if needed or just overwrite
        for i in range(1, 11):
            group = getattr(self, f'worker_group{i}', None)
            if not group: continue
            u_key = 'User' if i == 1 else f'User{i}'
            wt_key = 'WorkTime' if i == 1 else f'WorkTime{i}'
            ot_key = 'OT' if i == 1 else f'OT{i}'
            m_key = 'Meal' if i == 1 else f'Meal{i}'

            group.cb_name.set(self.clean_nan(record.get(u_key, '')))
            # [FIX] Use set_time to properly parse (Shift) Time string and populate both shift and time widgets
            group.set_time(self.clean_nan(record.get(wt_key, '')))

            group.ent_ot.delete(0, tk.END)
            group.ent_ot.insert(0, self.clean_nan(record.get(ot_key, '')))

            group.ent_meal.delete(0, tk.END)
            group.ent_meal.insert(0, self.clean_nan(record.get(m_key, '')))

        # 5. RTK

        for cat, ent in self.rtk_entries.items():
            if cat == "총계": continue
            val = record.get(f'RTK_{cat}', 0)
            ent.delete(0, tk.END)
            if abs(to_f_local(val)) > 0.001: ent.insert(0, str(val))
        self.calculate_rtk_total()


        # 6. Vehicle
        v_no_raw = self.clean_nan(record.get('차량번호', ''))
        v_insp_raw = self.clean_nan(record.get('차량점검', ''))
        v_mileage_raw = self.clean_nan(record.get('주행거리', ''))
        v_remarks_raw = self.clean_nan(record.get('차량비고', ''))

        v_no_list = [x.strip() for x in str(v_no_raw).split('||')] if v_no_raw else []
        v_insp_list = [x.strip() for x in str(v_insp_raw).split('||')] if v_insp_raw else []
        v_mileage_list = [x.strip() for x in str(v_mileage_raw).split('||')] if v_mileage_raw else []
        v_remarks_list = [x.strip() for x in str(v_remarks_raw).split('||')] if v_remarks_raw else []

        max_v_count = max(len(v_no_list), len(v_insp_list), len(v_mileage_list), len(v_remarks_list))

        # 기존 플로팅 창 닫기 및 고정 패널 비우기
        if hasattr(self, 'vehicle_panel'):
            self.vehicle_panel.reset()
        if hasattr(self, 'draggable_items'):
            for key, cont in list(self.draggable_items.items()):
                w = getattr(cont, '_widget', None)
                if not w: continue
                if type(w).__name__ == 'VehicleInspectionWidget':
                    self.remove_box(key)
        if hasattr(self, 'fixed_vehicle_widget'):
            self.fixed_vehicle_widget.reset_fields()

        for i in range(max_v_count):
            cur_no = v_no_list[i] if i < len(v_no_list) else ""
            cur_insp = v_insp_list[i] if i < len(v_insp_list) else ""
            cur_mileage = v_mileage_list[i] if i < len(v_mileage_list) else ""
            cur_remarks = v_remarks_list[i] if i < len(v_remarks_list) else ""

            if not (cur_no or cur_insp or cur_mileage or cur_remarks):
                continue

            v_parsed = {'vehicle_info': cur_no, 'mileage': cur_mileage, 'remarks': cur_remarks}
            if ':' in cur_insp:
                for pair in cur_insp.split('|'):
                    if ':' in pair:
                        k, v_val = pair.split(':', 1)
                        v_parsed[k.strip()] = v_val.strip()
            elif cur_insp and cur_insp != 'nan':
                for k in cur_insp.split(','):
                    k_clean = k.strip()
                    if k_clean:
                        if 'locking' in k_clean: v_parsed[k_clean] = '잠금'
                        elif 'cleaning' in k_clean: v_parsed[k_clean] = '함'
                        else: v_parsed[k_clean] = '양호'

            if i == 0 and hasattr(self, 'fixed_vehicle_widget'):
                v_widget = self.fixed_vehicle_widget
                v_widget.set_data(v_parsed)
            else:
                if hasattr(self, 'add_vehicle_inspection_box'):
                    v_widget = self.add_vehicle_inspection_box()
                    if hasattr(v_widget, 'set_data'):
                        v_widget.set_data(v_parsed)

        # 7. NDT Chemicals (다중 회사코드 형제 행 포함 전체 복원)
        if hasattr(self, 'ndt_company_entries') and self.ndt_company_entries:
            while len(self.ndt_company_entries) > 1:
                self.remove_last_ndt_company()
            ndt_names = ["형광자분", "흑색자분", "백색페인트", "침투제", "세척제", "현상제", "형광침투제"]
            sibling_records = [record]
            rec_entry_time = record.get('EntryTime')
            if (rec_entry_time is not None and not pd.isna(rec_entry_time)
                    and str(rec_entry_time).strip() and hasattr(self, 'daily_usage_df')
                    and not self.daily_usage_df.empty and 'EntryTime' in self.daily_usage_df.columns):
                matched_siblings = self.daily_usage_df[
                    (self.daily_usage_df.get('Date') == record.get('Date')) &
                    (self.daily_usage_df['EntryTime'] == rec_entry_time) &
                    (self.daily_usage_df.get('검사구분', '') != 'REP')
                ]
                if len(matched_siblings) > 1:
                    sibling_records = [row for _, row in matched_siblings.iterrows()]

            for idx_co, co_rec in enumerate(sibling_records):
                if idx_co >= len(self.ndt_company_entries):
                    self.add_ndt_company_section()
                target_entries = self.ndt_company_entries[idx_co]
                for name in ndt_names:
                    col_name = f"NDT_{name}"
                    if name in target_entries:
                        target_entries[name].delete(0, tk.END)
                        if col_name in co_rec:
                            val = co_rec[col_name]
                            if abs(to_f_local(val)) > 0.001:
                                target_entries[name].insert(0, str(val))
                raw_co = self.clean_nan(co_rec.get('회사코드', '')).strip()
                if raw_co and raw_co not in ('한국지역난방공사', '한국지역난방공사 중앙지사'):
                    target_entries['_company'].set(raw_co)
                else:
                    target_entries['_company'].set('')

        # 8. Note
        note_val = self.clean_nan(record.get('Note', ''))
        if 'Note' in record:
            self.ent_daily_note.delete(0, tk.END)
            # [REMOVED] 사용자 요청으로 과거 비고(Note) 내용을 불러오지 않음
            # self.ent_daily_note.insert(0, note_val)

        # 9. DailyMemo
        # 사용자가 저장한 메모를 불러오도록 복구 (과거 기록 호환성 포함)
        memo_val = self.clean_nan(record.get('DailyMemo', ''))
        if not memo_val and note_val:
            if ' | ' in str(note_val):
                parts = str(note_val).split(' | ', 1)
                if len(parts) > 1:
                    memo_val = parts[1]
            else:
                memo_val = note_val

        if hasattr(self, 'main_memo_text') and self.main_memo_text:
            try:
                self.main_memo_text.delete('1.0', tk.END)
                if memo_val:
                    self.main_memo_text.insert(tk.END, memo_val)
            except: pass

    except Exception as e:
        print(f"Error loading record to form: {e}")


def clear_daily_usage_form_all_impl(self, keep_date=False, keep_workers=False):
    """현장별 일일 사용량 입력 폼의 모든 필드를 초기화합니다."""
    # 1. 콤보박스 선택 해제
    self._daily_edit_source = None
    self._daily_material_selection = None
    for cb in [self.cb_daily_company, self.cb_daily_site, 
               self.cb_daily_test_method, self.cb_daily_unit]:
        cb.set('')

    # 2. 날짜 초기화 (연속 기입 시 편의를 위해 keep_date 옵션 제공)
    if not keep_date:
        try:
            self.ent_daily_date.set_date(datetime.date.today())
        except: pass

    self._last_daily_date = self.ent_daily_date.get_date()
    from site_apps.central.src.aux_equipment import restore
    restore(self)
    # 3. 모든 Entry 필드 비우기
    for ent in [self.ent_daily_inspection_item, self.ent_daily_applied_code, 
               self.ent_daily_equip, self.cb_daily_material, # Both are now Entries
               self.ent_daily_test_amount, self.ent_daily_unit_price, 
               self.ent_daily_report_no, self.ent_daily_travel_cost, 
               self.ent_daily_meal_cost, self.ent_daily_test_fee, self.ent_daily_note]:
        ent.delete(0, tk.END)

    # 기본값 복구
    self.cb_daily_unit.set('매')
    self.ndt_work_time_var.set("일반")
    self.ndt_source_var.set("Se-75 (1.0)")
    self.ndt_thickness_var.set("조건없음 (1.0)")
    self.ndt_pipe_var.set("250mm 초과 [10인치 이상] (1.0)")
    self.ndt_overhead_var.set(110.0)
    self.ndt_tech_var.set(20.0)
    self.ndt_ori_joint_var.set("")
    self.ndt_ori_qty_var.set("")
    self.ndt_rep_joint_var.set("")
    self.ndt_rep_qty_var.set("")
    self.ndt_rej_joint_var.set("")
    self.ndt_report_pipe_var.set("")
    self.ndt_loc_type_var.set('열수송관')
    self.hot_tap_stage_var.set('1차')
    self.hot_tap_size_var.set('600A')
    self.hot_tap_location_var.set('')
    self.ndt_calc_frame.grid_remove()
    self.rtk_grid.grid_remove() # [NEW] Hide RTK on clear
    self.ndt_frame.grid_remove() # [NEW] Hide NDT frame on clear

    # 4. NDT 자재 및 RTK 필드 초기화
    if hasattr(self, 'ndt_company_entries'):
        for entries in self.ndt_company_entries:
            for mat in ["형광자분", "흑색자분", "백색페인트", "침투제", "세척제", "현상제", "형광침투제"]:
                if mat in entries: entries[mat].delete(0, tk.END)

    for ent in self.rtk_entries.values():
        if ent.winfo_exists():
            ent.config(state='normal')
            ent.delete(0, tk.END)
            if ent == self.rtk_entries.get("총계"): 
                ent.config(state='readonly')

    # 5. 차량 점검 섹션 초기화 (중복 방지 및 고정 패널 사용)
    for key in list(getattr(self, 'vehicle_inspections', {})):
        self.remove_box(key)
    if hasattr(self, 'draggable_items'):
        for key, cont in list(self.draggable_items.items()):
            w = getattr(cont, '_widget', None)
            if not w: continue
            lbl_text = ""
            if hasattr(cont, '_label_widget'):
                lbl_text = cont._label_widget.cget('text')
            is_vehicle = (type(w).__name__ == 'VehicleInspectionWidget') or ('차량' in lbl_text and '점검' in lbl_text) or ('상시' in lbl_text and '차량' in lbl_text)
            if is_vehicle:
                print(f"[DEBUG] Removing legacy floating vehicle box: {lbl_text}")
                self.remove_box(key)

    if hasattr(self, 'fixed_vehicle_widget'):
        if hasattr(self.fixed_vehicle_widget, 'reset_fields'):
            self.fixed_vehicle_widget.reset_fields()
            self.fixed_vehicle_widget.clear_photo()

    # 6. [REMOVED] 메모 패널 초기화 방지 (사용자 요청)

    # 6. 작업자 섹션 초기화
    for i in range(1, 11):
        if keep_workers:
            continue
        group = getattr(self, f'worker_group{i}', None)
        if group:
            if hasattr(group, 'participation_var'):
                group.participation_var.set(True)
            group.cb_name.set('')
            group.cb_shift.set('주간')
            group.ent_worktime.set('')
            group.ent_ot.delete(0, tk.END)


def add_daily_usage_entry_impl(self):
    if hasattr(self, 'inspection_batch') and self.inspection_batch.plans:
        self.inspection_batch.save()
        return
    if self.ndt_loc_type_var.get() == '핫태핑':
        if self.cb_daily_test_method.get().strip().upper() not in ('MT', 'PT'):
            messagebox.showwarning('핫태핑', '핫태핑 길이표는 MT/PT 검사에 적용됩니다.')
            return
        if not self.hot_tap_location_var.get().strip():
            messagebox.showwarning('핫태핑', '1·2차 기록을 연결할 개소명/번호를 입력하세요.')
            return
        try:
            counts = [float(v.get().strip().replace(',', '') or '0') for v in
                      (self.ndt_ori_joint_var, self.ndt_rep_joint_var)]
            if any(not math.isfinite(n) or n < 0 or not n.is_integer() for n in counts) or sum(counts) <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning('핫태핑', 'ORI/REP 조인트수에 검사 개소 수를 양의 정수로 입력하세요.')
            return
        self._refresh_ndt_quantity()
    """현장별 일일 사용량 단일 등록 (리팩토링 버전)"""
    try:
        # 1. 기본 정보 및 유효성 검사
        date_val = self.ent_daily_date.get_date()
        site = self.cb_daily_site.get().strip()
        if not site:
            messagebox.showwarning("입력 오류", "현장명을 입력해주세요.")
            return

        # [NEW VALIDATION START]
        # 1-1. 차량 운행 유효성 검사
        if hasattr(self, 'vehicle_widget'):
            veh_data = self.vehicle_widget.get_data()
            if veh_data.get('vehicle_info'):
                if not veh_data.get('mileage'):
                    if not messagebox.askyesno("입력 확인", "차량 정보는 선택되었으나 주행거리가 입력되지 않았습니다.\n이대로 저장하시겠습니까?"):
                        return

        # 1-2. 투입 인원 대비 검사 물량 검사
        has_workers = False
        if hasattr(self, 'worker_groups'):
            from site_apps.central.src.worker_selection import worker_is_selected
            for wg in self.worker_groups:
                if worker_is_selected(wg) and wg.get_worker() and wg.get_time():
                    has_workers = True
                    break

        has_ndt_qty = False
        if hasattr(self, 'ndt_company_entries'):
            for comp_dict in self.ndt_company_entries:
                for mat, entry_widget in comp_dict.items():
                    if mat == '_company': continue
                    try:
                        qty = float(entry_widget.get().strip() or 0)
                        if qty > 0:
                            has_ndt_qty = True
                            break
                    except: pass
                if has_ndt_qty: break

        if has_workers and not has_ndt_qty:
            if not messagebox.askyesno("입력 확인", "투입된 인원(작업시간)은 있으나 검사 물량(NDT 수량)이 모두 0입니다.\n결측치일 가능성이 있습니다. 이대로 저장하시겠습니까?"):
                return
        # [NEW VALIDATION END]

        # 2. 품목 ID 확인 (디스플레이 명칭 기반 매칭)
        mat_display = self.cb_daily_material.get().strip()
        mat_id = ""
        current_method = self.cb_daily_test_method.get().strip().upper()
        if current_method in ('PAUT', 'MT'):
            # PAUT(길이검사) 및 MT(약품별도기입)는 MaterialID에 자재ID/장비명을 기입하지 않음
            mat_id = ""
        elif mat_display:
            from site_apps.central.src.material_selection import resolve_daily_material
            selected_material_id = resolve_daily_material(self, mat_display)
            if selected_material_id is not None:
                selected_rows = self.materials_df[self.materials_df['MaterialID'].apply(self.normalize_id) == self.normalize_id(selected_material_id)]
                m_name = str(selected_rows.iloc[0].get('품목명', '')).strip()
                if self._is_consumable_material(m_name, current_method):
                    mat_id = selected_material_id

            # [FIX] Only auto-register if it's a CONSUMABLE (Drug/Film). Equipment should NOT be registered.
            if not mat_id and mat_display:
                if self._is_consumable_material(mat_display, self.cb_daily_test_method.get().strip()):
                    # [UX IMPROVEMENT] 등록되지 않은 소모품 감지 시, 사용자의 승인을 받도록 알림 팝업 추가
                    if messagebox.askyesno("신규 자재 등록", f"등록되지 않은 소모품 자재명('{mat_display}')입니다.\n재고 마스터에 신규 자재로 등록하시겠습니까?\n\n(아니오를 누르면 작업일보에는 저장되지만 재고 연동(차감) 대상에서 제외됩니다.)"):
                        mat_id = self.register_new_material(mat_display, warehouse='현장', 규격='자동등록')
                    else:
                        mat_id = mat_display
                else:
                    # For equipment or non-consumables, do not store in MaterialID
                    mat_id = ""

        # Prepare first; replacement and persistence share one rollback boundary.
        from site_apps.central.src.views.daily_record_save import save_prepared
        plan = self._add_single_usage_record_logic(mat_id, date_val, site, prepare_only=True)
        saved_count = save_prepared(self, plan)
        if saved_count > 0:
            # [SMART VISIBILITY] Adjust filters to ensure the NEW record is visible
            try:
                current_start = self.ent_daily_start_date.get_date()
                current_end = self.ent_daily_end_date.get_date()

                if date_val < current_start:
                    self.ent_daily_start_date.set_date(date_val)
                if date_val > current_end:
                    self.ent_daily_end_date.set_date(date_val)
            except: pass

            # 성공 시 필드 초기화 및 뷰 갱신
            if hasattr(self, 'ndt_company_entries'):
                for company_entries in self.ndt_company_entries:
                    for name in self.ndt_materials_all:
                        if name in company_entries:
                            company_entries[name].delete(0, tk.END)

            # RTK 필드 초기화
            for cat, ent in self.rtk_entries.items():
                if cat != "총계":
                    ent.delete(0, tk.END)
            self.rtk_entries["총계"].config(state='normal')
            self.rtk_entries["총계"].delete(0, tk.END)
            self.rtk_entries["총계"].config(state='readonly')

            # [V13_RESET_REQUESTED] Clearing fields after successful entry based on user request
            # [FIX] Keep the current date for consecutive entries on the same day
            self.clear_daily_usage_form_all(keep_date=True, keep_workers=True)

            # Reset focus to first logical field (Company)
            self.cb_daily_company.focus_set()

            # [V13.1] Keep workers too for consecutive entries
            # for i in range(1, 11):
            #     group = getattr(self, f'worker_group{i}', None)
            #     if group:
            #         group.cb_name.set('')
            #         group.cb_shift.set('')
            #         group.ent_worktime.set('')
            #         group.ent_ot.delete(0, tk.END)

            # NDT 섹션 초기화
            if hasattr(self, 'ndt_company_entries'):
                while len(self.ndt_company_entries) > 1:
                    self.remove_last_ndt_company()
                if self.ndt_company_entries:
                    first = self.ndt_company_entries[0]
                    first['_company'].set('')
                    for k in self.ndt_materials_all:
                        if k in first: first[k].delete(0, tk.END)

            # 차량 점검 필드 초기화 (safeguarded)
            # [NEW] 현장별 저장 시 차량, 목록, 메모 레이블창 자동 닫기 (User request)
            for key in list(getattr(self, 'vehicle_inspections', {}).keys()):
                self.remove_box(key)
            for key in list(getattr(self, 'checklists', {}).keys()):
                self.remove_box(key)
            for key in list(getattr(self, 'memos', {}).keys()):
                self.remove_box(key)

            self.update_daily_usage_view()
            self.update_recent_entries_view()
            self.refresh_inquiry_filters()
            self.update_stock_view()

            # [NEW] Scroll to the bottom to see new entries
            if hasattr(self, 'daily_usage_tree'):
                self.daily_usage_tree.yview_moveto(1.0)

            messagebox.showinfo("성공", f"{saved_count}건의 기록이 안전하게 저장되었습니다.")
            self._recover_combobox_interaction()
    except Exception as e:
        messagebox.showerror("오류", f"기록 저장 중 오류 발생: {e}")
        self._recover_combobox_interaction()


def _recover_combobox_interaction_impl(self):
    """저장 직후 드롭다운이 열리지 않는 현상 방지: 남은 입력 점유(grab)·디바운스 상태를 정리한다."""
    try:
        # 1. 이미 사라졌거나 모달이 아닌 위젯이 잡고 있는 grab 해제
        holder = self.root.grab_current()
        if holder is not None:
            keep = False
            try:
                keep = holder.winfo_exists() and holder.winfo_ismapped() and holder is not self.root
            except Exception:
                keep = False
            if not keep:
                try: holder.grab_release()
                except Exception: pass
        # 2. 예약된 자동 제안(디바운스)·열림 플래그 초기화
        for name in ('cb_daily_company', 'cb_daily_site', 'cb_daily_material', 'cb_daily_equip',
                     'cb_daily_test_method', 'cb_daily_unit', 'ent_daily_inspection_item',
                     'ent_daily_applied_code'):
            cb = getattr(self, name, None)
            if cb is None:
                continue
            try:
                after_id = getattr(cb, '_suggest_after_id', None)
                if after_id:
                    cb.after_cancel(after_id)
                    cb._suggest_after_id = None
                cb._auto_opening = False
                cb._just_selected = False
            except Exception:
                pass
        # 3. 포커스를 메인 창으로 복귀시켜 첫 클릭이 드롭다운으로 전달되게 함
        self.root.after_idle(lambda: self.root.focus_force())
    except Exception as e:
        print(f"[WARN] 콤보박스 상태 복구 실패: {e}")


def load_report_mapping_impl(self):
    """Load report mapping from JSON file or return defaults"""
    mapping_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'resources', 'report_mapping.json')
    if not os.path.exists(mapping_path):
        # Return defaults matching the manager's default_mapping
        return {
            'header': {'date': 'F2'},
            'general': {
                'company': 'E5', 'project_name': 'E6', 'standard': 'E7', 'equipment': 'E8',
                'report_no': 'K5', 'inspection_item': 'K6', 'inspector': 'K7', 'car_no': 'K8'
            },
            'methods': {
                'RT': {'row': '12'}, 'UT': {'row': '13'}, 'MT': {'row': '14'}, 'PT': {'row': '15'},
                'HT': {'row': '16'}, 'VT': {'row': '17'}, 'LT': {'row': '18'}, 'ET': {'row': '19'},
                'PAUT': {'row': '20'}
            },
            'rtk': {
                'center_miss': 'C32', 'density': 'E32', 'marking_miss': 'G32', 'film_mark': 'I32',
                'handling': 'K32', 'customer_complaint': 'M32', 'etc': 'O32', 'total': 'Q32'
            },
            'ot': {
                'row1_name': 'B38', 'row1_hours': 'K38',
                'row2_name': 'B39', 'row2_hours': 'K39'
            },
            'materials': {
                'RT T200': '43', 'RT AA400': '44', 'MT WHITE': '46', 'MT 7C-BLACK': '47', 'MT FLUORESCENT': '48',
                'PT Penetrant': '49', 'PT Cleaner': '50', 'PT Developer': '51'
            }
        }
    try:
        with open(mapping_path, 'r', encoding='utf-8') as f:
            import json
            mapping = json.load(f)

            # Robust cleaning of method rows (e.g. 'E13' -> 13)
            if 'methods' in mapping:
                import re
                for m_key, m_val in mapping['methods'].items():
                    if isinstance(m_val, dict) and 'row' in m_val:
                        row_str = str(m_val['row'])
                        match = re.search(r'\d+', row_str)
                        if match: m_val['row'] = int(match.group())
            return mapping
    except:
        return {}


def save_report_mapping_impl(self, mapping):
    """Save report mapping to JSON file"""
    resources_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'resources')
    if not os.path.exists(resources_dir): 
        try: os.makedirs(resources_dir)
        except: pass
    mapping_path = os.path.join(resources_dir, 'report_mapping.json')
    try:
        print(f"DEBUG: Saving report mapping to {mapping_path}")
        with open(mapping_path, 'w', encoding='utf-8') as f:
            json.dump(mapping, f, indent=4, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"DEBUG: Save failed: {e}")
        messagebox.showerror("저장 오류", f"매핑 설정을 저장하지 못했습니다: {e}")
        return False


def open_report_mapping_dialog_impl(self):
    """Open a dialog to configure Excel mapping for Daily Work Report"""
    dialog = tk.Toplevel(self.root)
    dialog.title("작업일보 엑셀 매핑 설정")
    dialog.geometry("600x800")
    dialog.transient(self.root)
    dialog.grab_set()

    mapping = self.load_report_mapping()

    main_frame = ttk.Frame(dialog, padding=10)
    main_frame.pack(fill='both', expand=True)

    notebook = ttk.Notebook(main_frame)
    notebook.pack(fill='both', expand=True)

    # Helper to create grid of entries
    def create_entry_grid(parent, fields, current_data):
        entries = {}
        for i, (label_text, key) in enumerate(fields):
            ttk.Label(parent, text=label_text).grid(row=i, column=0, padx=5, pady=5, sticky='w')
            ent = ttk.Entry(parent, width=15)
            # Ensure we handle missing keys or nested dicts in current_data
            val = current_data.get(key, '')
            if isinstance(val, dict): val = val.get('row', '')
            ent.insert(0, str(val))
            ent.grid(row=i, column=1, padx=5, pady=5, sticky='w')
            entries[key] = ent
        return entries

    # Tab 1: General & Header
    tab1 = ttk.Frame(notebook, padding=10)
    notebook.add(tab1, text="일반/상단")
    gen_fields = [
        ("날짜 (Date)", "date_header"),
        ("업체명 (Company)", "company"), ("공사명 (Project)", "project_name"),
        ("적용규격 (Standard)", "standard"), ("장비명 (Equipment)", "equipment"),
        ("성적서번호 (Report No)", "report_no"), ("검사품명 (Item)", "inspection_item"),
        ("검사자 (Inspector)", "inspector"), ("차량번호 (Car No)", "car_no")
    ]
    tab1_data = mapping.get('general', {}).copy()
    tab1_data['date_header'] = mapping.get('header', {}).get('date', 'F2')
    tab1_entries = create_entry_grid(tab1, gen_fields, tab1_data)

    # Tab 2: Methods (Row Numbers)
    tab2 = ttk.Frame(notebook, padding=10)
    notebook.add(tab2, text="검사공법(행)")
    method_keys = ["RT", "UT", "MT", "PT", "HT", "VT", "LT", "ET", "PAUT"]
    method_fields = [(f"{m} (시작 행 번호)", m) for m in method_keys]
    tab2_entries = create_entry_grid(tab2, method_fields, mapping.get('methods', {}))

    # Tab 3: RTK Results
    tab3 = ttk.Frame(notebook, padding=10)
    notebook.add(tab3, text="불량율(RTK)")
    rtk_keys = ['center_miss', 'density', 'marking_miss', 'film_mark', 'handling', 'customer_complaint', 'etc', 'total']
    rtk_fields = [
        ("센터미스", "center_miss"), ("농도", "density"), ("마킹미스", "marking_miss"),
        ("필름마크", "film_mark"), ("취급부주의", "handling"), ("고객불만", "customer_complaint"),
        ("기타", "etc"), ("총계", "total")
    ]
    tab3_entries = create_entry_grid(tab3, rtk_fields, mapping.get('rtk', {}))

    # Tab 4: Materials (Row Numbers)
    tab4 = ttk.Frame(notebook, padding=10)
    notebook.add(tab4, text="자재(행)")
    mat_keys = list(mapping.get('materials', {}).keys())
    if not mat_keys: mat_keys = ['RT T200', 'RT AA400', 'MT WHITE', 'MT 7C-BLACK', 'MT FLUORESCENT', 'PT Penetrant', 'PT Cleaner', 'PT Developer']
    mat_fields = [(f"{m} (행 번호)", m) for m in mat_keys]
    tab4_entries = create_entry_grid(tab4, mat_fields, mapping.get('materials', {}))

    def save_and_close():
        try:
            new_mapping = {
                'header': {'date': tab1_entries['date_header'].get().strip()},
                'general': {k: tab1_entries[k].get().strip() for k in ['company', 'project_name', 'standard', 'equipment', 'report_no', 'inspection_item', 'inspector', 'car_no']},
                'methods': {m: {'row': tab2_entries[m].get().strip()} for m in method_keys},
                'rtk': {k: tab3_entries[k].get().strip() for k in rtk_keys},
                'ot': mapping.get('ot', {
                    'row1_name': 'B38', 'row1_hours': 'K38',
                    'row2_name': 'B39', 'row2_hours': 'K39'
                }),
                'materials': {k: tab4_entries[k].get().strip() for k in mat_keys}
            }

            if self.save_report_mapping(new_mapping):
                messagebox.showinfo("저장 완료", "매핑 설정이 저장되었습니다.")
                dialog.destroy()
        except Exception as e:
            messagebox.showerror("오류", f"설정 구성 중 오류가 발생했습니다: {e}")

    btn_frame = ttk.Frame(main_frame)
    btn_frame.pack(fill='x', pady=10)
    ttk.Button(btn_frame, text="✅ 설정 저장", command=save_and_close).pack(side='right', padx=5)
    ttk.Button(btn_frame, text="❌ 취소", command=dialog.destroy).pack(side='right', padx=5)


def _get_inout_session_data_impl(self):
    """Helper to collect In/Out tab entry data"""
    data = {
        'reg': {
            'co_code': self.cb_co_code.get().strip(),
            'eq_code': self.cb_eq_code.get().strip(),
            'item_name': self.cb_item_name.get().strip(),
            'model': self.cb_model.get().strip(),
            'sn': self.ent_sn.get().strip(),
            'class': self.cb_class.get().strip(),
            'spec': self.cb_spec.get().strip(),
            'unit': self.cb_unit.get().strip(),
            'supplier': self.cb_supplier.get().strip(),
            'mfr': self.cb_mfr.get().strip(),
            'origin': self.cb_origin.get().strip(),
            'reorder': self.ent_reorder.get().strip(),
            'init': self.ent_init.get().strip(),
            'price': self.ent_price.get().strip(),
            'cost': self.ent_cost.get().strip()
        },
        'trans': {
            'material': self.cb_material.get().strip(),
            'type': self.cb_type.get().strip(),
            'qty': self.ent_qty.get().strip(),
            'site': self.cb_trans_site.get().strip(),
            'warehouse': self.cb_warehouse.get().strip(),
            'user': self.ent_user.get().strip(),
            'note': self.ent_note.get().strip()
        }
    }
    return data


def _set_inout_session_data_impl(self, data):
    """Helper to restore In/Out tab entry data"""
    if not data: return
    reg = data.get('reg', {})
    self.cb_co_code.set(reg.get('co_code', ''))
    self.cb_eq_code.set(reg.get('eq_code', ''))
    self.cb_item_name.set(reg.get('item_name', ''))
    self.cb_model.set(reg.get('model', ''))
    self.ent_sn.delete(0, tk.END); self.ent_sn.insert(0, reg.get('sn', ''))
    self.cb_class.set(reg.get('class', ''))
    self.cb_spec.set(reg.get('spec', ''))
    self.cb_unit.set(reg.get('unit', ''))
    self.cb_supplier.set(reg.get('supplier', ''))
    self.cb_mfr.set(reg.get('mfr', ''))
    self.cb_origin.set(reg.get('origin', ''))
    self.ent_reorder.delete(0, tk.END); self.ent_reorder.insert(0, reg.get('reorder', '0'))
    self.ent_init.delete(0, tk.END); self.ent_init.insert(0, reg.get('init', '0'))
    self.ent_price.delete(0, tk.END); self.ent_price.insert(0, reg.get('price', '0'))
    self.ent_cost.delete(0, tk.END); self.ent_cost.insert(0, reg.get('cost', '0'))

    trans = data.get('trans', {})
    self.cb_material.set(trans.get('material', ''))
    self.cb_type.set(trans.get('type', 'OUT'))
    self.ent_qty.delete(0, tk.END); self.ent_qty.insert(0, trans.get('qty', ''))
    self.cb_trans_site.set(trans.get('site', ''))
    self.cb_warehouse.set(trans.get('warehouse', ''))
    self.ent_user.set(trans.get('user', ''))
    self.ent_note.delete(0, tk.END); self.ent_note.insert(0, trans.get('note', ''))


def _get_budget_session_data_impl(self):
    """Helper to collect Budget tab entry data"""
    data = {
        'main': {k: w.get().strip() for k, w in self.budget_widgets.items()},
        'labor_detail': self.labor_detail_widget.get_data(),
        'material_detail': self.material_detail_widget.get_data(),
        'expense_detail': self.expense_detail_widget.get_data()
    }
    return data


def _set_budget_session_data_impl(self, data):
    """Helper to restore Budget tab entry data"""
    if not data: return
    main = data.get('main', {})
    for k, val in main.items():
        if k in self.budget_widgets:
            w = self.budget_widgets[k]
            if hasattr(w, 'set'): w.set(val)
            else: w.delete(0, tk.END); w.insert(0, val)
    self.labor_detail_widget.set_data(data.get('labor_detail', {}))
    self.material_detail_widget.set_data(data.get('material_detail', {}))
    self.expense_detail_widget.set_data(data.get('expense_detail', {}))
    self._update_budget_kpis()


def save_form_session_impl(self):
    """Orchestrate saving session data from all entry tabs into one JSON"""
    try:
        global_data = {
            'usage': self._get_usage_session_data(),
            'inout': self._get_inout_session_data(),
            'budget': self._get_budget_session_data(),
            'saved_at': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }
        file_path = filedialog.asksaveasfilename(
            title="전체 세션 저장 (일괄)",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")],
            initialfile=f"PMI_Global_Session_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        )
        if file_path:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(global_data, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("저장 완료", f"현장/입출고/실행예산 탭의 입력 내용이 일괄 저장되었습니다.\n{os.path.basename(file_path)}")
    except Exception as e:
        messagebox.showerror("오류", f"전체 세션 저장 중 오류 발생: {e}")


def load_previous_day_data_impl(self):
    """선택된 현장의 가장 최근 기록(전일 데이터)을 불러와 폼을 채웁니다."""
    site = self.cb_daily_site.get().strip()
    if not site:
        messagebox.showwarning("입력 오류", "먼저 현장명을 선택하거나 입력해주세요.")
        return

    if self.daily_usage_df.empty:
        messagebox.showinfo("알림", "저장된 기록이 없습니다.")
        return

    df_site = self.daily_usage_df[self.daily_usage_df['Site'] == site].copy()
    if df_site.empty:
        messagebox.showinfo("알림", f"'{site}' 현장의 이전 기록이 없습니다.")
        return

    # 가장 최근 날짜 찾기
    df_site['Date'] = pd.to_datetime(df_site['Date'])
    recent_date = df_site['Date'].max()
    df_recent = df_site[df_site['Date'] == recent_date]

    date_str = recent_date.strftime('%Y-%m-%d')
    if not messagebox.askyesno("전일 데이터 불러오기", f"'{site}' 현장의 가장 최근 기록({date_str})을 불러오시겠습니까?\n(기존 입력된 작업자/차량 정보는 덮어쓰기 됩니다.)"):
        return

    try:
        # 1. 작업자 정보 복원 (최초 레코드 기준)
        first_record = df_recent.iloc[0]
        # 0. 일반 정보 복원 (업체명, 장비명 등)
        def _set_val(widget, val):
            if hasattr(widget, 'set'): widget.set(val)
            else:
                widget.delete(0, 'end')
                widget.insert(0, val)

        _set_val(self.cb_daily_company, self.clean_nan(first_record.get('업체명', '')))
        _set_val(self.cb_daily_equip, self.clean_nan(first_record.get('장비명', '')))
        from site_apps.central.src.aux_equipment import restore
        restore(self, first_record.get('보조장비', ''))

        self.ent_daily_applied_code.delete(0, 'end')
        self.ent_daily_applied_code.insert(0, self.clean_nan(first_record.get('적용코드', '')))

        self.ent_daily_inspection_item.delete(0, 'end')
        self.ent_daily_inspection_item.insert(0, self.clean_nan(first_record.get('검사품명', '')))

        _set_val(self.cb_daily_test_method, self.clean_nan(first_record.get('검사방법', '')))
        _set_val(self.cb_daily_unit, self.clean_nan(first_record.get('Unit', first_record.get('단위', ''))))

        self.ent_daily_unit_price.delete(0, 'end')
        unit_price = self.clean_nan(first_record.get('단가', ''))
        if str(unit_price) in ["0", "0.0"]: unit_price = ""
        self.ent_daily_unit_price.insert(0, unit_price)

        self.ent_daily_report_no.delete(0, 'end')
        self.ent_daily_report_no.insert(0, self.clean_nan(first_record.get('성적서번호', '')))

        self.ent_daily_note.delete(0, 'end')
        self.ent_daily_note.insert(0, self.clean_nan(first_record.get('Note', first_record.get('비고', ''))))

        for i in range(1, 11):
            group = getattr(self, f'worker_group{i}', None)
            if group:
                u_col = 'User' if i == 1 else f'User{i}'
                t_col = 'WorkTime' if i == 1 else f'WorkTime{i}'
                o_col = 'OT' if i == 1 else f'OT{i}'
                m_col = 'Meal' if i == 1 else f'Meal{i}'

                worker = first_record.get(u_col, first_record.get('작업자' if i == 1 else f'작업자{i}', ''))
                worktime = first_record.get(t_col, first_record.get('작업시간' if i == 1 else f'작업시간{i}', ''))
                ot = first_record.get(o_col, first_record.get('OT시간' if i == 1 else f'OT시간{i}', ''))
                meal = first_record.get(m_col, first_record.get('일비' if i == 1 else f'일비{i}', ''))

                group.set_worker(self.clean_nan(worker))
                group.set_time(self.clean_nan(worktime))
                group.set_ot(self.clean_nan(ot))
                if hasattr(group, 'set_meal'):
                    group.set_meal(self.clean_nan(meal))

        # 2. 차량 정보 복원
        if hasattr(self, 'fixed_vehicle_widget'):
            v_no_str = self.clean_nan(first_record.get('VehicleNo', first_record.get('차량번호', '')))
            v_nos = [x.strip() for x in v_no_str.split("||")] if "||" in v_no_str else [v_no_str]

            if v_nos and v_nos[0]:
                self.fixed_vehicle_widget.cb_vehicle_info.set(v_nos[0])

            if len(v_nos) > 1:
                existing_boxes = getattr(self, 'vehicle_boxes', [])
                for idx, extra_v in enumerate(v_nos[1:]):
                    if not extra_v: continue
                    if idx < len(existing_boxes):
                        existing_boxes[idx].cb_vehicle_info.set(extra_v)
                    elif hasattr(self, 'add_vehicle_inspection_box'):
                        new_box = self.add_vehicle_inspection_box()
                        if new_box: new_box.cb_vehicle_info.set(extra_v)

        messagebox.showinfo("완료", f"{date_str}의 작업자 및 차량 정보가 성공적으로 불러와졌습니다.\n(검사 물량 및 자재 소모량은 오늘 기준에 맞게 새로 입력해주세요.)")
    except Exception as e:
        messagebox.showerror("오류", f"데이터를 불러오는 중 오류가 발생했습니다: {e}")


def load_form_session_impl(self):
    """Orchestrate loading session data back to all entry tabs from a JSON file"""
    try:
        file_path = filedialog.askopenfilename(
            title="전체 세션 불러오기 (일괄)",
            filetypes=[("JSON files", "*.json")]
        )
        if not file_path: return
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        # Restore all tabs
        self._set_usage_session_data(data.get('usage'))
        self._set_inout_session_data(data.get('inout'))
        self._set_budget_session_data(data.get('budget'))
        messagebox.showinfo("불러오기 완료", "현장/입출고/실행예산 탭의 데이터가 일괄 복원되었습니다.")
    except Exception as e:
        messagebox.showerror("오류", f"전체 세션 불러오기 중 오류 발생: {e}")


def _auto_reconcile_and_register_ndt_impl(self, date_val, site, ndt_data, workers, company_code=""):
    """Unified helper to auto-register missing consumables and create 'OUT' transactions"""
    ndt_product_map = self._load_ndt_product_map()
    ndt_groups = self.ndt_groups

    note_suffix = f" ({company_code})" if company_code else ""
    note = f"{site} 현장 사용 (자동 차감){note_suffix}"

    # Ensure date_val is datetime and handle date vs datetime objects
    if isinstance(date_val, str):
        date_obj = pd.to_datetime(date_val)
    else:
        date_obj = date_val

    # extract date part safely
    pure_date = date_obj.date() if hasattr(date_obj, 'date') else date_obj
    trans_date = datetime.datetime.combine(pure_date, datetime.datetime.now().time())

    for name, qty in ndt_data.items():
        if qty <= 0: continue

        # 1. Resolve Item Name (Group) and Model
        db_item_name = ndt_product_map.get(name, '')
        db_model_name = name
        if not db_item_name:
            for grp, members in ndt_groups.items():
                if name in members:
                    db_item_name = grp
                    db_model_name = name # [FIX] Remove prefix for consistency with display logic
                    break
            if not db_item_name: db_item_name = "기타소모품"

        # 2. Find or Register (Highly flexible matching for NDT)
        target_item = db_item_name.replace(' ', '')

        # [FIX] Fill NaN with empty string before comparison to ensure reliable matching
        m_df = self.materials_df.copy()
        for col in ['품목명', '모델명']:
            if col in m_df.columns:
                m_df[col] = m_df[col].fillna('').astype(str).str.replace(' ', '', regex=False)

        # Try to find existing item
        # First, check for exact (item+model) match
        target_model = db_model_name.replace(' ', '')
        exact_match = m_df[(m_df['품목명'] == target_item) & (m_df['모델명'] == target_model)]

        if not exact_match.empty:
            mat_id = exact_match.iloc[0]['MaterialID']
        else:
            # If no exact match, try matching by item name ONLY for NDT groups
            item_only_match = m_df[m_df['품목명'] == target_item]
            if not item_only_match.empty:
                # Use the first available item in this group (e.g., generic "PT약품")
                mat_id = item_only_match.iloc[0]['MaterialID']
            else:
                # Truly missing - [REFINED] Do NOT register new one automatically
                print(f"DEBUG: Skipping auto-registration for missing NDT item '{db_item_name}' (Model: {db_model_name})")
                mat_id = None

        # 3. Create Transaction (Only if mat_id exists)
        if mat_id:
            self._create_manual_stock_transaction(trans_date, mat_id, 'OUT', qty, site, workers, note)


def _create_manual_stock_transaction_impl(self, date_val, mat_id, trans_type, qty, site, user, note):
    # [V23.1_WATERTIGHT_FILTER] Default to block everything unless it matches the whitelist
    try:
        if self.materials_df.empty:
            return # Can't identify item, don't save

        # Robust ID matching
        str_target_id = str(mat_id).strip().replace('.0', '')
        # Vectorized lookup for speed
        mask = self.materials_df['MaterialID'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True) == str_target_id
        mat_info = self.materials_df[mask]

        if mat_info.empty:
            # If we can't find it in DB, we don't know what it is. To be safe, DON'T save it to history.
            return

        row = mat_info.iloc[0]
        item_name = str(row.get('품목명', '')).strip()
        model_name = str(row.get('모델명', '')).strip()
        item_name_up = item_name.upper()
        model_name_up = model_name.upper()

        # 1. Check for Carestream (Partial Match, Case Insensitive)
        is_carestream = "CARESTREAM" in item_name_up or "CARESTREAM" in model_name_up

        # 2. Check for NDT Chemicals (More inclusive partial match for PT/MT/NDT)
        ndt_keywords = ["PT", "MT", "NDT", "침투제", "세척제", "현상제", "자분", "페인트"]
        is_ndt_chem = any(kw in item_name_up or kw in model_name_up for kw in ndt_keywords)

        # [CRITICAL] If not a whitelist item, exit NOW
        if not (is_carestream or is_ndt_chem):
            return

    except Exception as e:
        print(f"Filter error in _create_manual_stock_transaction: {e}")
        return # Block on error

    # --- IF WE GOT HERE, THE ITEM IS ALLOWED ---
    try:
        # Ensure quantity is negative for OUT
        signed_qty = -abs(float(qty)) if trans_type == 'OUT' else abs(float(qty))

        new_trans = {
            'Date': pd.to_datetime(date_val),
            'MaterialID': mat_id,
            'Type': trans_type,
            'Quantity': signed_qty,
            'User': user,
            'Site': site,
            'Note': note
        }
        self.transactions_df = pd.concat([self.transactions_df, pd.DataFrame([new_trans])], ignore_index=True)
    except: pass


def _auto_adjust_tree_columns_impl(self, tree, expand_only=False):
    """Automatically adjust column widths to fit content"""
    import tkinter.font as tkfont

    # Use the actual font used in the Treeview (12pt Malgun Gothic)
    # This ensures measurement matches display perfectly.
    content_font = tkfont.Font(family="Malgun Gothic", size=12)
    heading_font = tkfont.Font(family="Malgun Gothic", size=12, weight="bold")

    # We only care about visible columns
    visible_cols = tree['displaycolumns']
    if not visible_cols or visible_cols == ('#all'):
         visible_cols = tree['columns']

    for col in visible_cols:
        # Maximum Overkill: Character Count Heuristic
        # Assume 40px per character + 200px buffer
        # This GUARANTEES visibility in any environment

        # Measure heading
        w = len(col) * 40 + 200

        # Measure content
        col_index = list(tree['columns']).index(col)
        for item in tree.get_children():
            val = tree.item(item, 'values')
            if val and col_index < len(val):
                text_len = len(str(val[col_index]))
                content_w = text_len * 40 + 200
                if content_w > w:
                    w = content_w

        # Apply specific constraints
        # Ultra-Safe: No max limit for potentially long text columns
        if col in ['품목명', '작업자', '비고']:
            w = max(w, 300) # Minimum base
            if col == '품목명': w = max(w, 700)
            elif col == '작업자': w = max(w, 300)
            elif col == '비고': w = max(w, 600)
        else:
            # Other columns get a relaxed global max
            min_w = 200
            if col == '날짜': min_w = 400
            elif col == '현장': min_w = 300
            elif col == '장비명': min_w = 250
            elif col == '검사방법': min_w = 200

            w = max(min_w, min(w, 2000)) 

        # Smart Auto-Expand: Only grow if content requires it
        if expand_only:
            current_w = int(tree.column(col, 'width'))
            if w <= current_w:
                continue  # Keep user's manual adjustment

        # All columns are now user-resizable
        tree.column(col, width=w, minwidth=20, stretch=False, anchor='center')

    # Force layout update
    tree.update_idletasks()


def export_daily_usage_history_impl(self):
    """Export current filtered daily usage history to Excel"""
    try:
        # Get selected items only, or all if none selected
        selected_items = self.daily_usage_tree.selection()
        if not selected_items:
            # If no items selected, export all visible items
            items_to_export = self.daily_usage_tree.get_children()
        else:
            # Export only selected items
            items_to_export = selected_items

        data = []
        for item in items_to_export:
            data.append(self.daily_usage_tree.item(item, 'values'))

        if not data:
            messagebox.showinfo("알림", "내보낼 데이터가 없습니다.")
            return

        # Export columns currently visible in the Treeview
        display_cols = self.daily_usage_tree['displaycolumns']
        if not display_cols or display_cols == ('#all'):
            selected_cols = list(self.daily_usage_tree['columns'])
        else:
            selected_cols = list(display_cols)

        # Map selected column names to their indices in the Treeview values
        all_cols = list(self.daily_usage_tree['columns'])
        col_indices = [all_cols.index(col) for col in selected_cols]

        # Reconstruct data
        filtered_data = []
        for item_values in data:
            row = [item_values[i] for i in col_indices if i < len(item_values)]
            filtered_data.append(row)

        filename = f"일일사용내역_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=filename, filetypes=[("Excel files", "*.xlsx")])

        if save_path:
            cols = list(self.daily_usage_tree['columns'])
            col_to_idx = {col: i for i, col in enumerate(cols)}

            # Rebuild data from 'data' (which has ALL columns including hidden ones)
            # and swap summarized '작업자' with full list from '(Full작업자)'
            final_filtered_data = []
            for item_values in data:
                row_list = list(item_values)

                # Perform the swap if both columns exist
                if '작업자' in cols and '(Full작업자)' in col_to_idx:
                    full_idx = col_to_idx['(Full작업자)']
                    if full_idx < len(row_list):
                        full_val = row_list[full_idx]
                        if full_val:
                            # Find index of '작업자' in ALL columns to swap in row_list
                            idx_worker = cols.index('작업자')
                            row_list[idx_worker] = full_val

                # Extract only the selected display columns for final export
                final_row = [row_list[i] for i in col_indices if i < len(row_list)]
                final_filtered_data.append(final_row)

            df = pd.DataFrame(final_filtered_data, columns=selected_cols)
            df = self.clean_df_export(df)
            self.save_df_to_excel_autofit(df, save_path, "일일사용내역")
            messagebox.showinfo("완료", f"데이터가 엑셀로 저장되었습니다.\n{save_path}")
    except Exception as e:
        messagebox.showerror("오류", f"엑셀 내보내기 중 오류가 발생했습니다: {e}")


def show_column_visibility_dialog_impl(self, tree=None):
    """Open dialog to manually show/hide columns in the specified tree view"""
    if tree is None:
        tree = self.daily_usage_tree

    all_cols = list(tree['columns'])

    # Mandatory columns identification
    # NDT/RTK columns are at the end, workers in middle, basic info at start
    selectable_cols = [c for c in all_cols if c != '(Full작업자)']

    # [FIX] Use stored manual preferences instead of current filtered displaycolumns
    # so that hidden-by-data columns still appear checked if user wants them.
    preferred_defaults = ['날짜', '구분', '검사방법', '품목명', '수량', '단위', '단가', '검사비', '작업자', '작업시간', 'OT금액', '장비명', '업체명', '비고']
    if tree == self.daily_usage_tree:
        active_cols = getattr(self, 'manual_visible_cols', [])
        if not active_cols:
            active_cols = list(preferred_defaults)
    else:
        active_cols = tree['displaycolumns']
        if not active_cols or active_cols == ('#all'):
             active_cols = all_cols

    dialog = ColumnSelectionDialog(self.root, selectable_cols, title="표시 컬럼 관리")
    # Overwrite vars with current visibility state
    for col, var in dialog.vars.items():
        var.set(col in active_cols)

    dialog.wait_window()

    if dialog.result is not None:
        # Result set from dialog
        final_selection = list(dialog.result)

        # Reconstruct final display based on user result + mandatory ones
        mandatory_lookup = {
            self.daily_usage_tree: ['날짜', '구분', '검사방법', '품목명', '수량', '작업자'],
            self.inout_tree: ['날짜', '구분', '품목명', '장비명']
        }

        mandatory_cols = mandatory_lookup.get(tree, ['날짜'])
        for mc in mandatory_cols:
            if mc in all_cols and mc not in final_selection:
                final_selection.append(mc)

        # Re-order according to preferred order first, then remaining
        if tree == self.daily_usage_tree:
            sorted_selection = [c for c in preferred_defaults if c in final_selection]
            for c in all_cols:
                if c in final_selection and c not in sorted_selection:
                    sorted_selection.append(c)
        else:
            sorted_selection = [c for c in all_cols if c in final_selection]

        tree['displaycolumns'] = sorted_selection

        # Save configuration to the appropriate key
        if tree == self.daily_usage_tree:
            self.manual_visible_cols = sorted_selection
            if not hasattr(self, 'tab_config'): self.tab_config = {}
            self.tab_config['daily_usage_visible_cols'] = sorted_selection
        elif tree == self.inout_tree:
            if not hasattr(self, 'tab_config'): self.tab_config = {}
            self.tab_config['inout_visible_cols'] = sorted_selection

        self.save_tab_config()

        # Immediate refresh to apply selection
        if tree == self.daily_usage_tree:
            self.update_daily_usage_view()
        elif tree == self.inout_tree:
            self.update_transaction_view()


def export_all_daily_usage_impl(self):
    """Export all daily usage records to Excel"""
    try:
        if self.daily_usage_df.empty:
            messagebox.showinfo("알림", "기록된 데이터가 없습니다.")
            return

        # Export all columns from the tree structure
        selected_cols = list(self.daily_usage_tree['columns'])

        filename = f"전체사용기록_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=filename, filetypes=[("Excel files", "*.xlsx")])

        if save_path:
            # Ensure we only try to export columns that exist in the dataframe
            valid_cols = [c for c in selected_cols if c in self.daily_usage_df.columns]
            # If dataframe has extra columns not in tree (like raw data), include them? 
            # Better to stick to what the user sees/defines in tree interface + critical data
            # Actually, export_all usually implies backing up the raw dataframe. 
            # But previous logic used tree columns. Let's stick to tree columns if they exist in DF.

            # Fallback: if tree columns are not in DF (e.g. calculated columns), we might have issues.
            # The daily_usage_df has raw data. The tree has formatted data.
            # daily_usage_df columns: Date, Site, MaterialID, Usage, Note, etc. + RTK_..., NDT_...
            # Tree columns: 날짜, 현장, 작업자... OT합계...

            # If we want "All Daily Usage", we should probably export the RAW dataframe for backup purposes, 
            # OR the formatted view for reporting.
            # The previous code seemed to try to export based on tree columns but mapped from DF?
            # "export_df = self.daily_usage_df[selected_cols].copy()"
            # This implies 'selected_cols' MUST exist in daily_usage_df.
            # Let's check daily_usage_df columns again.
            # It has 'Date', 'Site', 'MaterialID'...
            # It does NOT have '날짜', '현장' (Korean names).
            # So the previous code 'self.daily_usage_df[selected_cols]' would have FAILED if selected_cols came from tree columns!

            # WAIT. The user said "remove column selection". 
            # The previous working code (before my changes) probably did:
            # export_df = self.daily_usage_df.copy()
            # Let's revert to a safe "Backup" style export for "All Data".

            export_df = self.daily_usage_df.copy()
            export_df = self.clean_df_export(export_df)
            self.save_df_to_excel_autofit(export_df, save_path, "전체사용기록")

            export_df = self.clean_df_export(export_df)
            self.save_df_to_excel_autofit(export_df, save_path, "전체기록")
            messagebox.showinfo("완료", f"전체 기록이 엑셀로 저장되었습니다.\n{save_path}")
    except Exception as e:
        messagebox.showerror("오류", f"엑셀 내보내기 중 오류가 발생했습니다: {e}")


def export_stock_to_excel_impl(self):
    """Export current stock data to Excel"""
    try:
        # Get current filtered data from treeview
        stock_data = []

        cols = list(self.stock_tree['columns'])
        for item in self.stock_tree.get_children():
            values = self.stock_tree.item(item, 'values')
            row_dict = {}
            for idx, col in enumerate(cols):
                if idx < len(values):
                    heading_text = self.stock_tree.heading(col, 'text') or col
                    row_dict[heading_text] = values[idx]
            stock_data.append(row_dict)

        if not stock_data:
            messagebox.showinfo("알림", "내보낼 데이터가 없습니다.")
            return

        # Prepare filename with current date
        current_date = datetime.datetime.now().strftime('%Y%m%d')
        filename = f"재고현황_{current_date}.xlsx"

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=filename,
            title="재고 현황 저장",
            filetypes=[("Excel files", "*.xlsx")]
        )

        if save_path:
            stock_df = pd.DataFrame(stock_data)
            stock_df = self.clean_df_export(stock_df)
            self.save_df_to_excel_autofit(stock_df, save_path, "재고현황")
            messagebox.showinfo("완료", f"재고 현황이 저장되었습니다.\n저장 위치: {save_path}")

    except Exception as e:
        messagebox.showerror("오류", f"내보내기 실패: {e}")


def enable_tree_column_drag_impl(self, tree, context_menu_handler=None):
    """Ultra-robust drag-and-drop column reordering using root-level event tracking"""
    tree._drag_info = {"col_id": None, "col_name": None, "start_x": 0, "active": False, "ghost": None}

    def on_press(event):
        region = tree.identify_region(event.x, event.y)
        if region == "heading":
            col_id = tree.identify_column(event.x)
            col_name = self._get_column_name_from_id(tree, col_id)
            if col_name:
                tree._drag_info.update({
                    "col_id": col_id,
                    "col_name": col_name,
                    "start_x": event.x_root,
                    "active": False,
                    "motion_id": self.root.bind("<B1-Motion>", on_motion, add="+"),
                    "release_id": self.root.bind("<ButtonRelease-1>", on_release, add="+")
                })

    def on_motion(event):
        di = tree._drag_info
        if not di["col_name"]: return

        # Start drag if moved > 5 pixels
        if not di["active"] and abs(event.x_root - di["start_x"]) > 5:
            di["active"] = True
            tree.config(cursor="fleur")

            # Visual ghost label
            try:
                if di["ghost"]: di["ghost"].destroy()
                di["ghost"] = tk.Toplevel(self.root)
                di["ghost"].overrideredirect(True)
                di["ghost"].attributes("-alpha", 0.8)
                di["ghost"].attributes("-topmost", True)

                lbl = tk.Label(di["ghost"], text=f" ↔ {di['col_name']} 이동 중 ", 
                               bg="#424242", fg="white", relief="raised", borderwidth=1, 
                               padx=8, pady=4, font=('Malgun Gothic', 10, 'bold'))
                lbl.pack()
            except: pass

        if di["active"] and di["ghost"]:
            try:
                # Follow mouse with offset to stay visible
                di["ghost"].geometry(f"+{event.x_root + 20}+{event.y_root + 10}")
            except: pass

    def on_release(event):
        di = tree._drag_info
        if di["active"] and di["col_name"]:
            # Convert screen X back to tree-relative X
            tree_x = event.x_root - tree.winfo_rootx()
            target_id = tree.identify_column(tree_x)

            if target_id and target_id != di["col_id"]:
                self._reorder_tree_columns(tree, di["col_id"], target_id)

        # Cleanup all bindings and state
        tree.config(cursor="")
        if di["ghost"]:
            try: di["ghost"].destroy()
            except: pass
            di["ghost"] = None

        # Unbind specific root-level drag listeners securely
        try:
            if di.get("motion_id"):
                self.root.unbind("<B1-Motion>", di["motion_id"])
            if di.get("release_id"):
                self.root.unbind("<ButtonRelease-1>", di["release_id"])
        except: pass

        di.update({"col_id": None, "col_name": None, "active": False, "motion_id": None, "release_id": None})

    # Primary entry point: tree-level press
    tree.bind("<Button-1>", on_press, add="+")
    # Direct right-click menu as reliable alternative
    if context_menu_handler is not None:
        tree.bind("<Button-3>", context_menu_handler, add="+")
    else:
        tree.bind("<Button-3>", lambda e: self._show_heading_context_menu(e, tree), add="+")


def _show_heading_context_menu_impl(self, event, tree):
    """Show context menu on Treeview header to move columns"""
    region = tree.identify_region(event.x, event.y)
    if region == "heading":
        column_id = tree.identify_column(event.x)

        # Identify the column name from visual ID
        col_name = self._get_column_name_from_id(tree, column_id)
        if not col_name: return

        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label=f"⬅️ '{col_name}' 왼쪽으로 이동", 
                         command=lambda: self._move_column_visual(tree, column_id, -1))
        menu.add_command(label=f"➡️ '{col_name}' 오른쪽으로 이동", 
                         command=lambda: self._move_column_visual(tree, column_id, 1))
        menu.add_separator()
        menu.add_command(label="⚙️ 컬럼 관리(숨기기/보이기)...", command=self.show_column_visibility_dialog)

        menu.post(event.x_root, event.y_root)


def _show_generic_tree_heading_context_menu_impl(self, event, tree):
    """Generic header context menu for trees that support only column move."""
    try:
        if tree.identify_region(event.x, event.y) != "heading":
            return
        column_id = tree.identify_column(event.x)
        col_name = self._get_column_name_from_id(tree, column_id)
        if not col_name:
            return
    except Exception:
        return

    menu = tk.Menu(self.root, tearoff=0)
    menu.add_command(label=f"⬅️ '{col_name}' 왼쪽으로 이동",
                     command=lambda: self._move_column_visual(tree, column_id, -1))
    menu.add_command(label=f"➡️ '{col_name}' 오른쪽으로 이동",
                     command=lambda: self._move_column_visual(tree, column_id, 1))
    try:
        menu.tk_popup(event.x_root, event.y_root)
    finally:
        menu.grab_release()


def _get_column_name_from_id_impl(self, tree, column_id):
    """Helper to get actual column name from visual ID like #1"""
    try:
        if not column_id or not column_id.startswith("#"): return None
        # Extract N from #N
        vis_idx = int(column_id[1:]) - 1

        # Get displaycolumns accurately
        disp_tuple = tree.cget("displaycolumns")
        if not disp_tuple or disp_tuple == ("#all",) or disp_tuple == ["#all"] or disp_tuple == "":
            visible_names = list(tree.cget("columns"))
        else:
            visible_names = list(disp_tuple)

        if 0 <= vis_idx < len(visible_names):
            return visible_names[vis_idx]
    except: pass
    return None


def _move_column_visual_impl(self, tree, source_id, direction):
    """Move a column left (-1) or right (1) based on its current position"""
    col_name = self._get_column_name_from_id(tree, source_id)
    if not col_name: return

    try:
        disp_tuple = tree.cget("displaycolumns")
        if not disp_tuple or disp_tuple == ("#all",) or disp_tuple == ["#all"] or disp_tuple == "":
            visible_names = list(tree.cget("columns"))
        else:
            visible_names = list(disp_tuple)

        if col_name not in visible_names: return

        src_idx = visible_names.index(col_name)
        tgt_idx = src_idx + direction

        if 0 <= tgt_idx < len(visible_names):
            visible_names.pop(src_idx)
            visible_names.insert(tgt_idx, col_name)
            tree["displaycolumns"] = visible_names

            # Update persistence
            if tree == self.daily_usage_tree:
                self.manual_visible_cols = visible_names
                self.save_tab_config()
            elif tree == self.monthly_usage_tree:
                self.monthly_visible_cols = visible_names
                self.save_tab_config()
            elif hasattr(self, 'budget_view_tree') and tree == self.budget_view_tree:
                self.budget_view_visible_cols = visible_names
                self.save_tab_config()
    except: pass


def _reorder_tree_columns_impl(self, tree, source_id, target_id):
    """Internal helper to swap displaycolumns for drag-and-drop"""
    src_name = self._get_column_name_from_id(tree, source_id)
    tgt_name = self._get_column_name_from_id(tree, target_id)

    if not src_name or not tgt_name or src_name == tgt_name: return

    try:
        disp_tuple = tree.cget("displaycolumns")
        if not disp_tuple or disp_tuple == ("#all",) or disp_tuple == ["#all"] or disp_tuple == "":
            visible_names = list(tree.cget("columns"))
        else:
            visible_names = list(disp_tuple)

        if src_name in visible_names and tgt_name in visible_names:
            src_idx = visible_names.index(src_name)
            tgt_idx = visible_names.index(tgt_name)

            # Strategic insertion: move before or after target 
            # (Simple swap approach: just pop and insert)
            visible_names.pop(src_idx)
            visible_names.insert(tgt_idx, src_name)

            tree["displaycolumns"] = visible_names

            # Update persistence
            if tree == self.daily_usage_tree:
                self.manual_visible_cols = visible_names
                self.save_tab_config()
            elif tree == self.monthly_usage_tree:
                self.monthly_visible_cols = visible_names
                self.save_tab_config()
            elif hasattr(self, 'budget_view_tree') and tree == self.budget_view_tree:
                self.budget_view_visible_cols = visible_names
                self.save_tab_config()
    except: pass

