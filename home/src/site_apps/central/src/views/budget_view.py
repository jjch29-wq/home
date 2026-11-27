# -*- coding: utf-8 -*-
"""공사실행예산서 탭 UI 및 집계·실적 반영 모듈 (Budget View Module)"""
import os
import re
import datetime
import traceback
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
import pandas as pd
import numpy as np

from site_apps.central.src.views.components import (
    LaborCostDetailWidget, MaterialCostDetailWidget, ExpenseProfitDetailWidget, ColumnSelectionDialog
)
from site_apps.central.src.utils.helpers import MARKER_PATTERN


def setup_budget_tab_impl(self):
    """Setup the project execution budget management tab with detailed labor cost breakdown"""
    # KPI Summary Panel [NEW]
    kpi_frame = tk.Frame(self.tab_budget, background="#ffffff", highlightthickness=1, highlightbackground="#e5e7eb")
    kpi_frame.pack(fill='x', padx=10, pady=(5, 0))

    kpis = [
        ("계약금액", "lbl_kpi_rev", "#374151"),
        ("실행원가", "lbl_kpi_cost", "#374151"),
        ("영업이익", "lbl_kpi_profit", "#10b981"),
        ("이익률", "lbl_kpi_margin", "#10b981")
    ]

    for i, (label, attr, color) in enumerate(kpis):
        f = tk.Frame(kpi_frame, background="#ffffff")
        f.pack(side='left', expand=True, fill='both', pady=10)
        if i > 0:
            tk.Frame(kpi_frame, width=1, background="#e5e7eb").pack(side='left', fill='y', pady=10)

        tk.Label(f, text=label, font=("Malgun Gothic", 9), background="#ffffff", foreground="#6b7280").pack()
        lbl_val = tk.Label(f, text="0원", font=("Malgun Gothic", 12, "bold"), background="#ffffff", foreground=color)
        lbl_val.pack()
        setattr(self, attr, lbl_val)

    # Fixed Button Frame [NEW]
    top_btn_frame = ttk.Frame(self.tab_budget)
    top_btn_frame.pack(fill='x', padx=10, pady=5)

    ttk.Button(top_btn_frame, text="📥 실적 데이터 채우기", command=self.fill_budget_from_actuals, style="Primary.TButton").pack(side='left', padx=3)
    ttk.Button(top_btn_frame, text="💾 저장/수정", command=self.save_budget_entry, style="Action.TButton").pack(side='left', padx=3)
    ttk.Button(top_btn_frame, text="🗑️ 삭제", command=self.delete_budget_entry, style="Danger.TButton").pack(side='left', padx=3)
    ttk.Button(top_btn_frame, text="🧹 초기화", command=self.clear_budget_form, style="Warning.TButton").pack(side='left', padx=3)

    # Pop-out button
    ttk.Button(top_btn_frame, text="🗗 팝업창 분리", command=self.open_detached_budget_view, style="Info.TButton").pack(side='left', padx=(10, 4))

    # [NEW] Toggle button for bottom pane
    self.btn_toggle_bottom = ttk.Button(top_btn_frame, text="🔽 하단 일일기록 숨기기", style="Secondary.TButton")
    self.btn_toggle_bottom.pack(side='left', padx=4)

    ttk.Button(top_btn_frame, text="📊 실행예산·정산서 출력", command=self.export_budget_estimation, style="Action.TButton").pack(side='right', padx=4)
    ttk.Button(top_btn_frame, text="📑 엑셀 내보내기", command=self.export_budget_sales_status, style="Secondary.TButton").pack(side='right', padx=4)

    main_paned = ttk.PanedWindow(self.tab_budget, orient='vertical')
    main_paned.pack(fill='both', expand=True, padx=10, pady=(0, 5))

    # --- 상단: Scrollable Form area ---
    top_container = ttk.Frame(main_paned)
    main_paned.add(top_container, weight=2)

    canvas = tk.Canvas(top_container, highlightthickness=0)
    v_scroll = ttk.Scrollbar(top_container, orient="vertical", command=canvas.yview)
    form_scrollable = ttk.Frame(canvas)

    form_scrollable.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    _canvas_win = canvas.create_window((0, 0), window=form_scrollable, anchor="nw")
    canvas.configure(yscrollcommand=v_scroll.set)

    # [FIX] 우측 빈 공간 제거: Canvas 너비 변경 시 내부 프레임 너비를 동기화
    def _on_canvas_resize(event):
        canvas.itemconfigure(_canvas_win, width=event.width)
    canvas.bind("<Configure>", _on_canvas_resize)

    # Scroll binding is handled globally in MaterialManager.__init__

    canvas.pack(side="left", fill="both", expand=True)
    v_scroll.pack(side="right", fill="y")

    # 1. Main Budget Form
    form_frame = ttk.LabelFrame(form_scrollable, text="📋 실행예산 입력/수정", padding=10)
    form_frame.pack(fill='x', pady=(0, 10), padx=5)

    # Grid layout for form fields
    # Grid layout headers
    headers = ["항목", "사전예산(계획)", "사후원가(실적)", "잔여예산(차액)"]
    for col, h in enumerate(headers):
        ttk.Label(form_frame, text=h, font=('Malgun Gothic', 9, 'bold')).grid(row=0, column=col, padx=10, pady=5)

    rows = [
        ("현장명", "cb_budget_site", None, None),
        ("공사기간 (일)", "ent_budget_period", "ent_budget_actual_period", "ent_budget_diff_period"),
        ("계약금액(Revenue) (원)", "ent_budget_revenue", "ent_budget_actual_revenue", "ent_budget_diff_revenue"),
        ("매출금액(UnitPrice) (원)", "ent_budget_unit_price", "ent_budget_actual_unit_price", "ent_budget_diff_unit_price"),
        ("실행 노무비(Labor) (원)", "ent_budget_labor", "ent_budget_actual_labor", "ent_budget_diff_labor"),
        ("실행 재료비(Material) (원)", "ent_budget_material", "ent_budget_actual_material", "ent_budget_diff_material"),
        ("실행 경비(Expense) (원)", "ent_budget_expense", "ent_budget_actual_expense", "ent_budget_diff_expense"),
        ("실행 외주비(Outsource) (원)", "ent_budget_outsource", "ent_budget_actual_outsource", "ent_budget_diff_outsource"),
        ("영업이익(Profit) (원)", "ent_budget_profit", "ent_budget_actual_profit", "ent_budget_diff_profit"),
        ("이익률(%)", "ent_budget_margin", "ent_budget_actual_margin", "ent_budget_diff_margin"),
        ("비고", "ent_budget_note", "ent_budget_actual_note", None)
    ]

    self.budget_widgets = {}
    for r_idx, (label, w_plan, w_actual, w_diff) in enumerate(rows, start=1):
        ttk.Label(form_frame, text=label).grid(row=r_idx, column=0, sticky='e', padx=10, pady=2)
        if w_plan == "cb_budget_site":
            w = ttk.Combobox(form_frame, width=20)
            w['values'] = getattr(self, 'sites', [])
            w.grid(row=r_idx, column=1, sticky='ew', padx=5, pady=2, columnspan=3)
        else:
            w = ttk.Entry(form_frame, width=20)
            w.grid(row=r_idx, column=1, sticky='ew', padx=5, pady=2)
        setattr(self, w_plan, w)
        self.budget_widgets[w_plan] = w

        if w_actual:
            w_a = ttk.Entry(form_frame, width=20)
            w_a.grid(row=r_idx, column=2, sticky='ew', padx=5, pady=2)
            setattr(self, w_actual, w_a)
            self.budget_widgets[w_actual] = w_a

        if w_diff:
            w_d = ttk.Entry(form_frame, width=20, state='readonly')
            w_d.grid(row=r_idx, column=3, sticky='ew', padx=5, pady=2)
            setattr(self, w_diff, w_d)
            self.budget_widgets[w_diff] = w_d

    def _on_budget_site_selected(e):
        # [FIX] 현장명을 가장 먼저 캡처 - focus_set() 호출 전에 해야 값이 유지됨
        selected_site = self.cb_budget_site.get().strip()
        # 폼 로드 (현장명 먼저 캡처했으므로 안전)
        self._load_budget_to_form(selected_site, silent=True)
        # 포커스 이동은 폼 로드 완료 후에 (after로 지연)
        self.root.after(50, self.ent_budget_revenue.focus_set)
    self.cb_budget_site.bind('<<ComboboxSelected>>', _on_budget_site_selected)

    # 1.5. 상세 탭 컨테이너 분리 (사전예산 / 사후원가)
    self.detail_notebook = ttk.Notebook(form_scrollable)
    self.detail_notebook.pack(fill='both', expand=True, pady=(0, 10), padx=5)

    self.tab_planned = ttk.Frame(self.detail_notebook)
    self.tab_actual = ttk.Frame(self.detail_notebook)

    self.detail_notebook.add(self.tab_planned, text='사전예산 상세 (Planned)')
    self.detail_notebook.add(self.tab_actual, text='사후원가(실적) 상세 (Actuals)')

    # ===== [ 사전예산 (Planned) 상세 위젯 ] =====
    labor_detail_frame = ttk.LabelFrame(self.tab_planned, text="인건비 상세 (Labor Cost Detail)", padding=10)
    labor_detail_frame.pack(fill='x', pady=(0, 10), padx=5)

    def on_labor_change(total):
        self.ent_budget_labor.delete(0, tk.END)
        self.ent_budget_labor.insert(0, f"{total:,.0f}")
        self._update_budget_kpis()

    self.labor_detail_widget = LaborCostDetailWidget(labor_detail_frame, on_change_callback=on_labor_change)
    self.labor_detail_widget.pack(fill='x', expand=True)

    material_detail_frame = ttk.LabelFrame(self.tab_planned, text="재료비 상세 (Material Cost Detail)", padding=10)
    material_detail_frame.pack(fill='x', pady=(0, 10), padx=5)

    def on_material_change(total):
        self.ent_budget_material.delete(0, tk.END)
        self.ent_budget_material.insert(0, f"{total:,.0f}")
        self._update_budget_kpis()

    self.material_detail_widget = MaterialCostDetailWidget(material_detail_frame, on_change_callback=on_material_change)
    self.material_detail_widget.pack(fill='x', expand=True)

    expense_detail_frame = ttk.LabelFrame(self.tab_planned, text="경비 및 이익 상세 (Expense & Profit Detail)", padding=10)
    expense_detail_frame.pack(fill='x', pady=(0, 10), padx=5)

    def on_expense_change(exp_total, outsource_total, op_profit):
        self.ent_budget_expense.delete(0, tk.END)
        self.ent_budget_expense.insert(0, f"{exp_total:,.0f}")
        self.ent_budget_outsource.delete(0, tk.END)
        self.ent_budget_outsource.insert(0, f"{outsource_total:,.0f}")
        if hasattr(self, 'ent_budget_profit'):
            rev = get_rev()
            margin = (op_profit / rev * 100) if rev > 0 else 0.0
            self.ent_budget_profit.delete(0, tk.END)
            self.ent_budget_profit.insert(0, f"{op_profit:,.0f} ({margin:.1f}%)")
        self._update_budget_kpis()

    def get_lab():
        try:
            if hasattr(self, 'labor_detail_widget') and hasattr(self.labor_detail_widget, 'get_exact_total'):
                return self.labor_detail_widget.get_exact_total()
            return float(self.ent_budget_labor.get().replace(',', '') or 0)
        except: return 0.0
    def get_mat():
        try:
            if hasattr(self, 'material_detail_widget') and hasattr(self.material_detail_widget, 'get_exact_total'):
                return self.material_detail_widget.get_exact_total()
            return float(self.ent_budget_material.get().replace(',', '') or 0)
        except: return 0.0
    def get_rev():
        try: 
            rev = float(self.ent_budget_revenue.get().replace(',', '') or 0)
            if rev == 0 and hasattr(self, 'ent_budget_unit_price'):
                rev = float(self.ent_budget_unit_price.get().replace(',', '') or 0)
            return rev
        except: return 0.0

    self.expense_detail_widget = ExpenseProfitDetailWidget(
        expense_detail_frame, 
        on_change_callback=on_expense_change,
        get_labor_total_func=get_lab,
        get_material_total_func=get_mat,
        get_revenue_func=get_rev,
        master_app=self
    )
    self.expense_detail_widget.pack(fill='x', expand=True)

    # ===== [ 사후원가 (Actual) 상세 위젯 ] =====
    a_labor_detail_frame = ttk.LabelFrame(self.tab_actual, text="실적 인건비 상세 (Actual Labor Detail)", padding=10)
    a_labor_detail_frame.pack(fill='x', pady=(0, 10), padx=5)

    def on_actual_labor_change(total):
        self.ent_budget_actual_labor.delete(0, tk.END)
        self.ent_budget_actual_labor.insert(0, f"{total:,.0f}")
        self._update_budget_kpis()

    self.actual_labor_detail_widget = LaborCostDetailWidget(a_labor_detail_frame, on_change_callback=on_actual_labor_change)
    self.actual_labor_detail_widget.pack(fill='x', expand=True)

    a_material_detail_frame = ttk.LabelFrame(self.tab_actual, text="실적 재료비 상세 (Actual Material Detail)", padding=10)
    a_material_detail_frame.pack(fill='x', pady=(0, 10), padx=5)

    def on_actual_material_change(total):
        self.ent_budget_actual_material.delete(0, tk.END)
        self.ent_budget_actual_material.insert(0, f"{total:,.0f}")
        self._update_budget_kpis()

    self.actual_material_detail_widget = MaterialCostDetailWidget(a_material_detail_frame, on_change_callback=on_actual_material_change)
    self.actual_material_detail_widget.pack(fill='x', expand=True)

    a_expense_detail_frame = ttk.LabelFrame(self.tab_actual, text="실적 경비 및 이익 상세 (Actual Expense & Profit)", padding=10)
    a_expense_detail_frame.pack(fill='x', pady=(0, 10), padx=5)

    def on_actual_expense_change(exp_total, outsource_total, op_profit):
        self.ent_budget_actual_expense.delete(0, tk.END)
        self.ent_budget_actual_expense.insert(0, f"{exp_total:,.0f}")
        self.ent_budget_actual_outsource.delete(0, tk.END)
        self.ent_budget_actual_outsource.insert(0, f"{outsource_total:,.0f}")
        if hasattr(self, 'ent_budget_actual_profit'):
            rev = get_actual_rev()
            margin = (op_profit / rev * 100) if rev > 0 else 0.0
            self.ent_budget_actual_profit.delete(0, tk.END)
            self.ent_budget_actual_profit.insert(0, f"{op_profit:,.0f} ({margin:.1f}%)")
        self._update_budget_kpis()

    def get_actual_lab():
        try:
            if hasattr(self, 'actual_labor_detail_widget') and hasattr(self.actual_labor_detail_widget, 'get_exact_total'):
                return self.actual_labor_detail_widget.get_exact_total()
            return float(self.ent_budget_actual_labor.get().replace(',', '') or 0)
        except: return 0.0
    def get_actual_mat():
        try:
            if hasattr(self, 'actual_material_detail_widget') and hasattr(self.actual_material_detail_widget, 'get_exact_total'):
                return self.actual_material_detail_widget.get_exact_total()
            return float(self.ent_budget_actual_material.get().replace(',', '') or 0)
        except: return 0.0
    def get_actual_rev():
        try: 
            rev = float(self.ent_budget_actual_unit_price.get().replace(',', '') or 0)
            if rev == 0 and hasattr(self, 'ent_budget_actual_revenue'):
                rev = float(self.ent_budget_actual_revenue.get().replace(',', '') or 0)
            return rev
        except: return 0.0

    self.actual_expense_detail_widget = ExpenseProfitDetailWidget(
        a_expense_detail_frame, 
        on_change_callback=on_actual_expense_change,
        get_labor_total_func=get_actual_lab,
        get_material_total_func=get_actual_mat,
        get_revenue_func=get_actual_rev,
        budget_mode='actual'
    )
    self.actual_expense_detail_widget.pack(fill='x', expand=True)

    # [NEW] 도급액 실시간 입력에 따른 이익금액 재계산 바인딩
    if hasattr(self, 'ent_budget_revenue'):
        self.ent_budget_revenue.bind('<KeyRelease>', lambda e: self.expense_detail_widget.calculate_all())
        self.ent_budget_revenue.bind('<FocusOut>', lambda e: self.format_entry_with_commas(e, self.ent_budget_revenue), add='+')
        self.ent_budget_revenue.bind('<Return>', lambda e: self.format_entry_with_commas(e, self.ent_budget_revenue), add='+')

    if hasattr(self, 'ent_budget_unit_price'):
        self.ent_budget_unit_price.bind('<FocusOut>', lambda e: self.format_entry_with_commas(e, self.ent_budget_unit_price), add='+')
        self.ent_budget_unit_price.bind('<Return>', lambda e: self.format_entry_with_commas(e, self.ent_budget_unit_price), add='+')

    # [MOVED to top_btn_frame]

    # ===============================================================
    # --- 하단: 현장별 일일사용량 조회 프레임 ---
    # ===============================================================
    bottom_container = ttk.Frame(main_paned)
    main_paned.add(bottom_container, weight=3)
    self._budget_bottom_visible = True

    def toggle_budget_bottom():
        if getattr(self, '_budget_bottom_visible', True):
            main_paned.remove(bottom_container)
            self.btn_toggle_bottom.config(text="🔼 하단 일일기록 펼치기")
            self._budget_bottom_visible = False
        else:
            main_paned.add(bottom_container, weight=3)
            self.btn_toggle_bottom.config(text="🔽 하단 일일기록 숨기기")
            self._budget_bottom_visible = True

    self.btn_toggle_bottom.config(command=toggle_budget_bottom)

    # 필터 영역
    bottom_filter = ttk.Frame(bottom_container)
    bottom_filter.pack(fill='x', pady=(5, 2))

    ttk.Label(bottom_filter, text="📊 현장별 실적 조회", font=('Malgun Gothic', 10, 'bold')).pack(side='left', padx=(0, 10))

    ttk.Label(bottom_filter, text="🏢 현장:").pack(side='left', padx=(4, 2))
    self.cb_budget_view_site = ttk.Combobox(bottom_filter, width=18, state='readonly')
    site_values = ['전체'] + sorted([str(s).strip() for s in (self.sites or []) if str(s).strip()])
    self.cb_budget_view_site['values'] = site_values
    self.cb_budget_view_site.pack(side='left', padx=(0, 6))
    self.cb_budget_view_site.set('전체')
    self.cb_budget_view_site.bind('<<ComboboxSelected>>', lambda e: self.update_budget_site_view())

    ttk.Label(bottom_filter, text="🗓️ 시작일:").pack(side='left', padx=(6, 2))
    self.budget_view_start = DateEntry(bottom_filter, width=12, date_pattern='yyyy-mm-dd', locale='ko_KR', state='readonly', showweeknumbers=True)
    self.budget_view_start.pack(side='left', padx=(0, 4))
    today = datetime.datetime.now()
    self.budget_view_start.set_date(datetime.date(2024, 1, 1))

    ttk.Label(bottom_filter, text=" ~ 📅 종료일:").pack(side='left', padx=(4, 2))
    self.budget_view_end = DateEntry(bottom_filter, width=12, date_pattern='yyyy-mm-dd', locale='ko_KR', state='readonly', showweeknumbers=True)
    self.budget_view_end.pack(side='left', padx=(0, 6))
    self.budget_view_end.set_date(today)

    ttk.Button(bottom_filter, text="🔍 조회", command=self.update_budget_site_view, style="Primary.TButton").pack(side='left', padx=3)
    ttk.Button(bottom_filter, text="📥 예산 수정/불러오기",
               command=lambda: self._load_budget_to_form(self.cb_budget_view_site.get()), style="Info.TButton").pack(side='left', padx=3)
    ttk.Button(bottom_filter, text="➕ 컬럼 추가", command=self.add_budget_view_custom_column, style="Secondary.TButton").pack(side='left', padx=3)
    ttk.Button(bottom_filter, text="⚙️ 컬럼 관리", command=self.show_budget_view_column_dialog, style="Secondary.TButton").pack(side='left', padx=3)
    ttk.Button(bottom_filter, text="📊 엑셀 내보내기", command=self.export_budget_sales_status, style="Action.TButton").pack(side='right', padx=10)

    # KPI 요약 행 제거 (예산입력 5번 영업이익률로 대체)
    # Treeview (일일사용량 코럼)
    tree_outer = ttk.Frame(bottom_container)
    tree_outer.pack(fill='both', expand=True)

    # 전체 컬럼: 기본 표시 컬럼 + 사용자가 "추가" 할 수 있는 숨김 기본 컬럼 + 사용자정의 컬럼
    self.budget_view_builtin_cols = (
        'Date', 'Site', '장비명', '검사방법', '작업자', '품목명',
        '총기성액', '청구인건비', '청구재료비', '제경비', '기술료',
        '총지출액', '자재원가', 'OT합계', '출장비', '일식', 
        '예상이윤', '목표이익률', '현재이익률', '검사량', '단가', '검사단가', '합계', 
        '침투제', '세척제', '현상제', '자재단가', '비고',
        '입력시간', '차량번호', '주행거리', '차량점검', '차량비고', 'MaterialID'
    )
    self.budget_view_builtin_width_map = {
        'Date': 90, 'Site': 120, '장비명': 100, '검사방법': 80, '검사량': 70, '단가': 80,
        '검사단가': 90, '출장비': 80, '일식': 70, 'OT합계': 80, '침투제': 70, '세척제': 70, '현상제': 70, '자재단가': 90,
        '합계': 100, '비고': 150, '작업자': 170, '품목명': 180, '입력시간': 140,
        '차량번호': 110, '주행거리': 100, '차량점검': 120, '차량비고': 140, 'MaterialID': 90,
        '총기성액': 100, '청구인건비': 90, '청구재료비': 90, '제경비': 90, '기술료': 90,
        '총지출액': 100, '자재원가': 90, '예상이윤': 100, '목표이익률': 80, '현재이익률': 80
    }
    self.budget_view_builtin_head_map = {
        'Date': '날짜', 'Site': '현장', '장비명': '장비', '검사방법': '검사방법', '검사량': '수량', '단가': '단가',
        '검사단가': '매출금액', '출장비': '출장비', '일식': '일식', 'OT합계': 'OT지급액', '침투제': '침투제', '세척제': '세척제', '현상제': '현상제',
        '자재단가': '자재단가', '합계': '합계', '비고': '비고', '작업자': '작업자', '품목명': '자재명',
        '입력시간': '입력시간', '차량번호': '차량번호', '주행거리': '주행거리', '차량점검': '차량점검',
        '차량비고': '차량비고', 'MaterialID': '자재ID',
        '총기성액': '총기성액', '청구인건비': '인건비(청구)', '청구재료비': '재료비(청구)', '제경비': '제경비', '기술료': '기술료',
        '총지출액': '총지출액', '자재원가': '자재원가', '예상이윤': '예상이윤', '목표이익률': '목표이익률(%)', '현재이익률': '현재이익률(%)'
    }
    self.budget_view_heading_aliases = getattr(self, 'budget_view_heading_aliases', {})
    self.budget_view_custom_columns = getattr(self, 'budget_view_custom_columns', [])
    self.budget_view_default_cols = (
        'Date', 'Site', '장비명', '검사방법', '작업자', '품목명',
        '총기성액',
        '총지출액', '자재원가', 'OT합계', '출장비', '일식', 
        '예상이윤', '목표이익률', '현재이익률', '비고'
    )

    vsb = ttk.Scrollbar(tree_outer, orient='vertical')
    hsb = ttk.Scrollbar(tree_outer, orient='horizontal')
    self.budget_view_tree = ttk.Treeview(tree_outer, columns=self.budget_view_builtin_cols, show='headings',
                                          yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    self._refresh_budget_view_tree_columns(reload_data=False)

    vsb.config(command=self.budget_view_tree.yview)
    hsb.config(command=self.budget_view_tree.xview)
    self.budget_view_tree.grid(row=0, column=0, sticky='nsew')
    vsb.grid(row=0, column=1, sticky='ns')
    hsb.grid(row=1, column=0, sticky='ew')
    tree_outer.grid_rowconfigure(0, weight=1)
    tree_outer.grid_columnconfigure(0, weight=1)
    self.budget_view_tree.bind('<ButtonRelease-1>', lambda e: self.save_tab_config())
    self.enable_tree_column_drag(self.budget_view_tree, context_menu_handler=self._show_budget_view_heading_context_menu)

    # ── ESC: 현장별 탭 내 모든 입력 위젯에서 포커스 해제 ──────────────
    def _esc_to_root(e=None):
        self.root.focus_set()
        return "break"

    def _bind_esc_recursive(widget):
        try:
            cls = widget.winfo_class()
            # Toplevel(달력 팝업 등) 내부는 절대 순회하지 않음
            # → 순회만으로도 winfo_children() 가 Toplevel을 화면에 표시시킬 수 있음
            if cls == 'Toplevel':
                return
            if cls in ('TEntry', 'Entry', 'TCombobox'):
                widget.bind('<Escape>', _esc_to_root, add='+')
        except Exception:
            pass
        for child in widget.winfo_children():
            _bind_esc_recursive(child)

    # after_idle: 모든 위젯 생성 완료 후 바인딩
    def _safe_bind_esc():
        try:
            if hasattr(self, 'tab_daily_usage') and self.tab_daily_usage.winfo_exists():
                _bind_esc_recursive(self.tab_daily_usage)
        except tk.TclError:
            pass

    def _safe_update_budget():
        try:
            if hasattr(self, 'root') and self.root.winfo_exists():
                self.update_budget_site_view()
        except tk.TclError:
            pass

    self.root.after_idle(_safe_bind_esc)

    # 초기 데이터 로드 (첫 현장 자동 선택 후 표시)
    self.root.after(100, _safe_update_budget)

    self._budget_loaded_site = self.cb_budget_site.get().strip()
    self._budget_baseline = self._capture_budget_input()

def update_budget_site_view_impl(self):
    """현장별 일일사용량 데이터를 하단 Treeview에 표시하고 KPI를 갱신한다."""
    try:
        if not hasattr(self, 'budget_view_tree'):
            return

        for item in self.budget_view_tree.get_children():
            self.budget_view_tree.delete(item)

        if not hasattr(self, 'cb_budget_view_site'):
            return
        site = self.cb_budget_view_site.get().strip()

        # [NEW] Track logically unique entries to avoid double-counting materials from split-row records
        processed_entry_ids = set()

        # [FIX] Do NOT return early if df is empty, we still need to run column hiding logic
        if self.daily_usage_df.empty:
            df_empty = True
        else:
            df_empty = False

        # --- 날짜 필터 ---
        start_date = self.budget_view_start.get_date()
        end_date   = self.budget_view_end.get_date()
        df = self.daily_usage_df.copy()
        # 구버전 컬럼명 호환
        if 'Site' not in df.columns and '현장' in df.columns:
            df['Site'] = df['현장']
        if 'Date' not in df.columns and '날짜' in df.columns:
            df['Date'] = df['날짜']
        if 'Site' not in df.columns:
            df_empty = True

        if not df_empty:
            df['_site_norm'] = df['Site'].astype(str).str.strip()

        # 날짜 파싱(실패 데이터는 제외 대신 전체 누락 방지를 위해 후단에서 완화)
        if 'Date' in df.columns:
            # Use normalized Timestamps for robust comparison with tk date objects
            df['_date'] = pd.to_datetime(df['Date'], errors='coerce').dt.normalize()
            start_ts = pd.to_datetime(start_date).normalize()
            end_ts = pd.to_datetime(end_date).normalize()
            date_mask = (df['_date'] >= start_ts) & (df['_date'] <= end_ts)
            # [STRICT] Exclude any rows with invalid/missing dates (NaT)
            date_mask = date_mask & df['_date'].notna()
        else:
            date_mask = pd.Series([True] * len(df), index=df.index)

        if site and site != '전체':
            site_mask = (df['_site_norm'] == site)
        else:
            site_mask = pd.Series([True] * len(df), index=df.index)

        mask = date_mask & site_mask
        df = df[mask].copy()

        # --- 자재 원가 맵 ---
        mat_id_cost_map = {}
        if 'MaterialID' in self.materials_df.columns and '원가' in self.materials_df.columns:
            mat_id_cost_map = self.materials_df.set_index('MaterialID')['원가'].fillna(0).to_dict()

        def _f(val):
            if pd.isna(val) or str(val).lower() == 'nan': return 0.0
            try: 
                s = str(val).replace(',', '').strip()
                if not s: return 0.0
                num = float(s)
                return 0.0 if np.isnan(num) else num
            except: return 0.0

        total_income = 0.0   # 검사비 + OT 합계
        total_expense = 0.0  # 출장비 + 일식
        total_net_revenue = 0.0 # 검사비 합계
        total_sum     = 0.0
        total_mat_cost = 0.0 # 자재 실적 사용료 합계
        total_ndt_penetrant = 0.0 # NDT 침투제 합계
        total_ndt_cleaner = 0.0 # NDT 세척제 합계
        total_ndt_developer = 0.0 # NDT 현상제 합계

        # [NEW] Enhanced Budget Totals
        sys_total_revenue = 0.0
        sys_total_labor = 0.0
        sys_total_mat_bill = 0.0
        sys_total_overhead = 0.0
        sys_total_tech = 0.0
        sys_total_expense_cost = 0.0
        sys_total_profit = 0.0

        import re as _re

        # [NEW] Robust is_active helper for smart column hiding
        def is_active(val):
            if val is None: return False
            s = str(val).strip().lower()
            if s in ('', '0', '0.0', '0.00', 'nan', 'none', '-', '0원', '0.0원', '0시간'):
                return False
            try:
                # Remove common units for numeric check
                v_raw = s.replace(',', '').replace('원', '').replace('시간', '').strip()
                if not v_raw: return False
                v = float(v_raw)
                return abs(v) > 0.001
            except:
                return bool(s)

        # [NEW] Smart Column Hiding tracking
        mandatory_cols = {'Date', 'Site', '합계', '비고'}
        has_data_map = {col: False for col in self.budget_view_cols if col not in mandatory_cols}

        def _clean_str(val):
            if pd.isna(val):
                return ''
            s = str(val).strip()
            return '' if s.lower() in ('nan', 'none') else s

        def _first_text(_row, keys, default=''):
            """여러 후보 컬럼 중 첫 유효 문자열 반환 (구버전 컬럼명 호환)."""
            for k in keys:
                if k in _row.index:
                    v = _clean_str(_row.get(k, ''))
                    if v:
                        return v
            return default

        def _ndt_val(_row, *keys):
            for k in keys:
                if k in _row.index:
                    return _f(_row.get(k, 0))
            return 0.0

        # 규격별 작업 행의 검사 기성액은 모두 합산하되, 동일 근무의 OT는 한 번만 차감한다.
        processed_financial_ot_signatures = set()

        for _, row in df.iterrows():
            # [NEW] Logical entry deduplication for the display list (Treeview)
            e_date = str(row.get('Date', '')).split(' ')[0]
            e_site = str(row.get('Site', '')).strip()
            raw_time = str(row.get('EntryTime', '')).strip()
            e_time = raw_time[:16] if len(raw_time) > 16 else raw_time
            e_uid = f"{e_date}|{e_site}|{e_time}"

            if e_uid in processed_entry_ids:
                continue # Skip redundant rows from being displayed in the list
            processed_entry_ids.add(e_uid)

            # [FIX] OT 컬럼에서 금액 추출 로직 강화 (NaN 방지)
            parsed_ots = []
            for i in range(1, 11):
                raw_ot = row.get('OT' if i == 1 else f'OT{i}', 0)
                if pd.isna(raw_ot) or not str(raw_ot).strip() or str(raw_ot).lower() == 'nan':
                    parsed_ots.append(0.0)
                    continue

                v_str = str(raw_ot).strip()
                if '(' in v_str and '원)' in v_str:
                    try:
                        amt_str = v_str.split('(')[1].split('원')[0].replace(',', '').strip()
                        parsed_ots.append(_f(amt_str))
                    except: parsed_ots.append(0.0)
                else:
                    parsed_ots.append(_f(v_str))

            ot_sum = sum(parsed_ots)

            worker_time_pairs = []
            for i in range(1, 11):
                user_key = 'User' if i == 1 else f'User{i}'
                time_key = 'WorkTime' if i == 1 else f'WorkTime{i}'
                worker_name = _clean_str(row.get(user_key, ''))
                if worker_name:
                    worker_time_pairs.append((worker_name, _clean_str(row.get(time_key, ''))))

            financial_ot_signature = (
                str(row.get('Date', '')).split(' ')[0],
                str(row.get('Site', '')).strip(),
                tuple(worker_time_pairs),
            )
            is_duplicate_financial_ot = financial_ot_signature in processed_financial_ot_signatures
            if worker_time_pairs:
                processed_financial_ot_signatures.add(financial_ot_signature)
            if is_duplicate_financial_ot:
                ot_sum = 0.0

            net    = _f(row.get('검사비', 0))
            travel = _f(row.get('출장비', 0))
            meal   = _f(row.get('일식', 0))
            qty    = _f(row.get('검사량', 0))
            price  = _f(row.get('단가', 0))

            # 자재원가
            test_method = str(row.get('검사방법', '')).upper()
            if 'RT' in test_method:
                usage = _f(row.get('Usage', 0))
            else:
                usage = _f(row.get('Usage', 0))

            # [NEW] Track logically unique entries to avoid double-counting materials from split-row records
            # Based on Date + Site + EntryTime (best UID) or Date + Site + Note (as fallback)
            e_time = str(row.get('EntryTime', '')).strip()
            e_site = str(row.get('Site', row.get('현장', ''))).strip()
            e_date = str(row.get('Date', row.get('날짜', ''))).strip()
            e_note = str(row.get('Note', row.get('비고', ''))).strip()
            e_uid = f"{e_date}|{e_site}|{e_time}|{e_note}"

            is_new_logical_entry = (e_uid not in processed_entry_ids)
            processed_entry_ids.add(e_uid)

            # 일일 탭 NDT 입력값(침투제/세척제/현상제)
            # [FIX] ONLY sum consumables for the first row of each logical entry to prevent inflation
            ndt_penetrant = 0.0
            ndt_cleaner = 0.0
            ndt_developer = 0.0

            if is_new_logical_entry:
                ndt_penetrant = _ndt_val(row, 'NDT_침투제', '침투제', 'NDT_침투액', '침투액')
                ndt_cleaner = _ndt_val(row, 'NDT_세척제', '세척제', 'NDT_세척액', '세척액')
                ndt_developer = _ndt_val(row, 'NDT_현상제', '현상제', 'NDT_현상액', '현상액')

                total_ndt_penetrant += ndt_penetrant
                total_ndt_cleaner += ndt_cleaner
                total_ndt_developer += ndt_developer

            unit_cost = mat_id_cost_map.get(row.get('MaterialID'), 0)
            mat_cost = usage * float(unit_cost)
            total_mat_cost += mat_cost

            # [NEW] Calculate Revenue and Profit per row
            nd_labor = _f(row.get('인건비', 0))
            nd_mat = _f(row.get('재료비', 0))
            nd_overhead = _f(row.get('제경비', 0))
            nd_tech = _f(row.get('기술료', 0))

            # If NDT fields exist, use them.
            row_revenue = nd_labor + nd_mat + nd_overhead + nd_tech

            # [FIX] Absorb 1-won truncation differences into tech fee so component sums perfectly match the exact subtotal (net)
            if row_revenue > 0 and net > 0 and abs(net - row_revenue) < 10:
                nd_tech += (net - row_revenue)
                row_revenue = net
            elif row_revenue <= 0:
                row_revenue = net

            row_expense_cost = travel + meal + ot_sum + mat_cost
            row_profit = row_revenue - row_expense_cost
            row_margin = (row_profit / row_revenue * 100) if row_revenue > 0 else 0.0

            # --- Financial Totals (Sum for ALL rows) ---
            total_net_revenue += net
            total_expense += (travel + meal)
            total_income += (net + ot_sum)
            total_sum += (net + travel + meal)

            # [NEW] Accumulate system totals
            sys_total_revenue += row_revenue
            sys_total_labor += nd_labor
            sys_total_mat_bill += nd_mat
            sys_total_overhead += nd_overhead
            sys_total_tech += nd_tech
            sys_total_expense_cost += row_expense_cost
            sys_total_profit += row_profit

            # 작업자 통합
            worker_cols = ['User'] + [f'User{i}' for i in range(2, 11)]
            raw_workers = []
            for wcol in worker_cols:
                wv = _clean_str(row.get(wcol, ''))
                if wv and wv not in raw_workers:
                    raw_workers.append(wv)
            consolidated_workers = ', '.join(raw_workers)

            # 품목명 표시명
            try:
                mat_name = self.get_material_display_name(row.get('MaterialID')) if pd.notna(row.get('MaterialID')) else ''
            except Exception:
                mat_name = ''

            row_map = {
                'Date': _clean_str(row.get('Date', '')),
                'Site': _clean_str(row.get('Site', '')),
                '장비명': _clean_str(row.get('장비명', '')),
                '검사방법': _clean_str(row.get('검사방법', '')),
                '검사량': f"{qty:g}",
                '단가': f"{price:,.0f}",
                '검사단가': f"{net:,.0f}",
                '출장비': f"{travel:,.0f}",
                '일식': f"{meal:,.0f}",
                'OT합계': f"{ot_sum:,.0f}" if ot_sum else '',
                '침투제': f"{_f(row.get('NDT_침투제', row.get('침투제', 0))):g}",
                '세척제': f"{_f(row.get('NDT_세척제', row.get('세척제', 0))):g}",
                '현상제': f"{_f(row.get('NDT_현상제', row.get('현상제', 0))):g}",
                '자재단가': f"{unit_cost:,.0f}",
                '합계': f"{(net + travel + meal):,.0f}",
                '비고': _clean_str(row.get('비고', row.get('Note', ''))),
                '작업자': consolidated_workers,
                '품목명': mat_name,
                '입력시간': _clean_str(row.get('EntryTime', '')),
                '차량번호': _first_text(row, ['차량번호', 'vehicle_info', 'VehicleInfo']),
                '주행거리': _first_text(row, ['주행거리', 'mileage', 'Mileage']),
                '차량점검': _first_text(row, ['차량점검', '차량 점검', 'vehicle_checks', 'inspection', 'Inspection']),
                '차량비고': _first_text(row, ['차량비고', '차량 비고', 'remarks', 'vehicle_remarks', 'Remark']),
                'MaterialID': _clean_str(row.get('MaterialID', '')),
                'Usage': _clean_str(row.get('Usage', '')),
                '총기성액': f"{row_revenue:,.0f}" if row_revenue else '',
                '청구인건비': f"{nd_labor:,.0f}" if nd_labor else '',
                '청구재료비': f"{nd_mat:,.0f}" if nd_mat else '',
                '제경비': f"{nd_overhead:,.0f}" if nd_overhead else '',
                '기술료': f"{nd_tech:,.0f}" if nd_tech else '',
                '총지출액': f"{row_expense_cost:,.0f}" if row_expense_cost else '',
                '자재원가': f"{mat_cost:,.0f}" if mat_cost else '',
                '예상이윤': f"{row_profit:,.0f}" if row_profit else '',
                '목표이익률': '',
                '현재이익률': f"{row_margin:.1f}%" if row_revenue > 0 else '',
            }
            for custom_col in getattr(self, 'budget_view_custom_columns', []) or []:
                key = custom_col.get('key')
                if key:
                    row_map[key] = custom_col.get('default', '')

            # [NEW] Mark columns that have non-zero/non-empty data using robust is_active
            for col, val in row_map.items():
                if col in has_data_map and not has_data_map[col]:
                    if is_active(val):
                        has_data_map[col] = True

            self.budget_view_tree.insert('', tk.END, values=tuple(row_map.get(col, '') for col in self.budget_view_cols))

            try:
                equip_name = str(row.get('장비명', ''))
                m_name_local = str(row.get('검사방법', ''))
                date_key = str(row.get('Date', ''))
                if 'RT' in m_name_local or 'RT' in equip_name or 'RT' in mat_name:
                    rt_dates.add(date_key)

                # [NEW] Generic equipment tracking by name
                if equip_name and equip_name.lower() != 'nan':
                    if equip_name not in equip_dates_map:
                        equip_dates_map[equip_name] = set()
                    equip_dates_map[equip_name].add(date_key)
            except Exception:
                pass

        # --- 합계 행 추가 ---
        if df.shape[0] > 0:
            total_row_map = {
                'Date': f"[합계]",
                'Site': f"{site}",
                '장비명': '',
                '검사방법': '',
                '검사량': '',
                '단가': '',
                '검사단가': f"{total_net_revenue:,.0f}",
                '출장비': f"{total_expense:,.0f}",
                '일식': '',
                'OT합계': '',
                '자재사용량': '',
                '침투제': f"{total_ndt_penetrant:g}",
                '세척제': f"{total_ndt_cleaner:g}",
                '현상제': f"{total_ndt_developer:g}",
                '자재단가': '',
                '합계': f"{total_sum:,.0f}",
                '비고': f"소계: 수입 {total_income:,.0f} 경비 {total_expense:,.0f} 자재원가 {total_mat_cost:,.0f}",
                '작업자': '',
                '품목명': '',
                '입력시간': '',
                '차량번호': '',
                '주행거리': '',
                '차량점검': '',
                '차량비고': '',
                'MaterialID': '',
                'Usage': '',
                '총기성액': f"{sys_total_revenue:,.0f}" if sys_total_revenue else '',
                '청구인건비': f"{sys_total_labor:,.0f}" if sys_total_labor else '',
                '청구재료비': f"{sys_total_mat_bill:,.0f}" if sys_total_mat_bill else '',
                '제경비': f"{sys_total_overhead:,.0f}" if sys_total_overhead else '',
                '기술료': f"{sys_total_tech:,.0f}" if sys_total_tech else '',
                '총지출액': f"{sys_total_expense_cost:,.0f}" if sys_total_expense_cost else '',
                '자재원가': f"{total_mat_cost:,.0f}" if total_mat_cost else '',
                '예상이윤': f"{sys_total_profit:,.0f}" if sys_total_profit else '',
                '목표이익률': '',
                '현재이익률': f"{(sys_total_profit / sys_total_revenue * 100):.1f}%" if sys_total_revenue > 0 else '',
            }
            for custom_col in getattr(self, 'budget_view_custom_columns', []) or []:
                key = custom_col.get('key')
                if key:
                    total_row_map[key] = ''
            self.budget_view_tree.insert('', tk.END, values=tuple(total_row_map.get(col, '') for col in self.budget_view_cols), tags=('total',))

        # --- [NEW] 예산 대비 실적 요약 행 추가 ---
        if df.shape[0] > 0 and site and site != '전체' and hasattr(self, 'budget_df') and not self.budget_df.empty:
            budget_match = self.budget_df[self.budget_df['Site'] == site]
            if not budget_match.empty:
                b_row = budget_match.iloc[0]

                # Fetch Budget Values
                def to_f_b(key):
                    try: 
                        v = str(b_row.get(key, 0)).replace(',', '').strip()
                        return float(v) if v else 0.0
                    except: return 0.0

                b_rev = to_f_b('Revenue')
                b_labor = to_f_b('LaborCost')
                b_mat = to_f_b('MaterialCost')
                b_exp = to_f_b('Expense')
                b_profit = to_f_b('Profit')

                # Progress / Variance Calculations
                rev_progress = (total_income / b_rev * 100) if b_rev > 0 else 0
                mat_progress = (total_mat_cost / b_mat * 100) if b_mat > 0 else 0
                exp_progress = (total_expense / b_exp * 100) if b_exp > 0 else 0

                # Final Net Profit Achievement
                actual_profit = total_income - total_expense - total_mat_cost
                b_margin = (b_profit / b_rev * 100) if b_rev > 0 else 0
                actual_margin = (sys_total_profit / sys_total_revenue * 100) if sys_total_revenue > 0 else 0

                comparison_row_map = {
                    'Date': f"[예산대비]",
                    'Site': f"달성률: {rev_progress:.1f}%",
                    '장비명': '',
                    '검사방법': '',
                    '검사량': '',
                    '단가': '',
                    '검사단가': f"예산 {b_rev:,.0f}",
                    '출장비': f"경비 {exp_progress:.1f}%",
                    '일식': '',
                    'OT합계': '',
                    '자재사용량': f"자재 {mat_progress:.1f}%",
                    '침투제': '',
                    '세척제': '',
                    '현상제': '',
                    '자재단가': '',
                    '합계': f"{actual_profit:,.0f}",
                    '비고': f"목표이익: {b_profit:,.0f} | 현재이익: {actual_profit:,.0f} ({actual_profit - b_profit:+,.0f})",
                    '목표이익률': f"{b_margin:.1f}%",
                    '현재이익률': f"{actual_margin:.1f}%",
                    '작업자': '',
                    '품목명': '',
                    '입력시간': '',
                    '차량번호': '',
                    '주행거리': '',
                    '차량점검': '',
                    '차량비고': '',
                    'MaterialID': '',
                    'Usage': '',
                    '총기성액': '', '청구인건비': '', '청구재료비': '', '제경비': '', '기술료': '',
                    '총지출액': '', '자재원가': '', '예상이윤': ''
                }
                for custom_col in getattr(self, 'budget_view_custom_columns', []) or []:
                    key = custom_col.get('key')
                    if key:
                        comparison_row_map[key] = ''

                # Insert with a distinct tag for styling
                self.budget_view_tree.tag_configure('comparison', background='#E3F2FD', font=('Malgun Gothic', 10, 'bold'))
                self.budget_view_tree.insert('', tk.END, values=tuple(comparison_row_map.get(col, '') for col in self.budget_view_cols), tags=('comparison',))

        # --- Smart Column Hiding Application ---
        # 기성 청구 세부 내역(인건비(청구), 재료비(청구), 제경비, 기술료)은 [기성 정산 (NDT)] 탭의 고유 영역이므로
        # 공사실행예산서 기본 보기에서는 제외하고 총기성액(매출) 및 손익 중심으로 표시
        hidden_billing_cols = {'청구인건비', '청구재료비', '제경비', '기술료'}

        saved_visible = getattr(self, 'budget_view_visible_cols', [])
        if saved_visible:
            # Remove billing details from saved preference to keep clean view
            cleaned_visible = [c for c in saved_visible if c not in hidden_billing_cols and c in self.budget_view_cols]
            self.budget_view_visible_cols = cleaned_visible
            active_cols = cleaned_visible if cleaned_visible else [c for c in self.budget_view_default_cols if c in self.budget_view_cols]
        else:
            active_cols = [col for col in self.budget_view_cols 
                           if col not in hidden_billing_cols and (col in mandatory_cols or has_data_map.get(col, False))]
        self.budget_view_tree['displaycolumns'] = active_cols

        # --- Treeview 태그 스타일 (합계 행 강조) ---

        if not self.budget_view_tree.tag_configure('total'):
            self.budget_view_tree.tag_configure('total', background='#ffff00', foreground='#000000')
    except Exception as e:
        print(f"ERROR in update_budget_site_view: {e}")

def fill_budget_from_actuals_impl(self):
    """하단 필터(현장, 기간) 기준으로 실제 실적 데이터를 집계하여 상단 예산 폼에 자동 입력한다."""
    site = self.cb_budget_view_site.get().strip()
    if not site:
        messagebox.showwarning("현장 미선택", "하단 '현장별 실적 조회' 영역에서 현장을 먼저 선택해주세요.")
        return

    start_ts = pd.Timestamp(self.budget_view_start.get_date())
    end_ts = pd.Timestamp(self.budget_view_end.get_date())

    if self.daily_usage_df.empty:
        messagebox.showinfo("데이터 없음", "집계할 실적 데이터가 없습니다.")
        return

    # 1. 필터링
    df = self.daily_usage_df.copy()
    df['_date_ts'] = pd.to_datetime(df['Date'], errors='coerce')
    if site == '전체':
        mask = (df['_date_ts'] >= start_ts) & (df['_date_ts'] <= end_ts)
    else:
        mask = (df['_date_ts'] >= start_ts) & (df['_date_ts'] <= end_ts) & (df['Site'] == site)
    df = df[mask].copy()

    if df.empty:
        s_date = self.budget_view_start.get_date()
        e_date = self.budget_view_end.get_date()
        messagebox.showinfo("데이터 없음", f"선택한 기간 ({s_date} ~ {e_date}) 동안 '{site}' 현장의 실적이 없습니다.")
        return

    # 2. 집계 초기화
    def _f(val):
        if pd.isna(val) or str(val).lower() == 'nan': return 0.0
        try: 
            s = str(val).replace(',', '').strip()
            if not s: return 0.0
            num = float(s)
            return 0.0 if np.isnan(num) else num
        except: return 0.0

    total_net_revenue = 0.0 # 검사비
    total_travel = 0.0
    total_meal = 0.0
    lab_total = 0.0
    outsource_total = 0.0  # 작업자 없는 행 → 외주비

    # --- Labor Aggregation by Rank ---
    ranks = ["이사", "부장", "차장", "과장", "대리", "계장", "주임", "기사"]
    # Track unique dates for each worker to calculate active days correctly
    # Structure: {rank: {worker_name: set(date_strings)}}
    rank_labor_dates = {r: {} for r in ranks}

    # --- Special Work (OT) Classification ---
    special_types = ["연장근무", "야간근무", "휴일근무"]
    ot_data = {t: {'names': set(), 'hours': 0} for t in special_types}

    import re as _re
    # Updated pattern to support [Rank], Rank Name, or Name alone with mapping
    # Supports: "[부장] 주진철", "부장 주진철", "주진철"
    rank_pattern = _re.compile(r"\[?(이사|부장|차장|과장|대리|계장|주임|기사)\]?\s*(.*)")

    # Default mapping for core workers if rank is missing
    worker_rank_map = {
        "주진철": "부장", "우명광": "대리", "김진환": "주임", "장승대": "계장",
        "김성렬": "주임", "박광복": "부장", "주영광": "과장", "이경재": "주임",
        "황조현": "주임", "김춘호": "차장", "박원준": "과장", "이봉주": "부장",
        "김성현": "대리"
    }

    # 상세 재료비 집계를 위한 맵
    material_usage_sums = [0.0] * 11
    mat_map = {
        '세척': 0, '침투': 1, '현상': 2, '자분': 3, '흑색': 4,
        '필름': 5, '글리세': 6, '현상액': 7, '정착액': 8, '수적': 9
    }

    # 자재 마스터 정보
    mat_id_name_map = {}
    mat_id_cost_map = {}
    if not self.materials_df.empty:
        mat_id_name_map = self.materials_df.set_index('MaterialID')['품목명'].fillna('').to_dict()
        mat_id_cost_map = self.materials_df.set_index('MaterialID')['원가'].fillna(0).to_dict()

    total_mat_cost = 0.0
    rt_inspection_count = 0  # Counter for RT occurrences
    total_film_count = 0.0
    # [NEW] Use unique days instead of row count to avoid inflating days for multiple records
    total_days_count = df['_date_ts'].dt.date.nunique() if not df.empty else 0

    # [NEW] Track logically unique entries to avoid double-counting materials from split-row records
    processed_entry_ids = set()

    # 같은 날짜/현장에서 규격별 작업을 따로 저장해도 동일 작업자의 동일 OT는 1회만 집계한다.
    # 작업시간이 다른 경우에는 실제 별도 근무일 수 있으므로 시간 문자열까지 식별키에 포함한다.
    processed_ot_signatures = set()

    # [NEW] Track unique dates for vehicles and specialized equipment
    starex_dates = set()
    toptruck_dates = set()
    paut_dates = set()
    paut_manual_scanner_dates = set()
    paut_cobra_scanner_dates = set()
    mt_dates = set()
    rt_dates = set()
    # [NEW] Track unique dates for ANY equipment name found in the '장비명' column
    equip_dates_map = {}

    for _, row in df.iterrows():
        # Check weekday
        date_val = _re.sub(r'\s.*', '', str(row.get('Date', '')))

        try:
            dt_obj = pd.to_datetime(date_val)
            is_weekday = dt_obj.weekday() < 5 # 0-4: Mon-Fri
        except:
            is_weekday = True

        # [NEW] Enhanced Revenue Logic
        nd_labor = _f(row.get('인건비', 0))
        nd_mat = _f(row.get('재료비', 0))
        nd_overhead = _f(row.get('제경비', 0))
        nd_tech = _f(row.get('기술료', 0))
        has_worker = any(
            str(row.get('User' if i == 1 else f'User{i}', '')).strip() not in ('', 'nan')
            for i in range(1, 11)
        )
        if has_worker:
            lab_total += nd_labor
        else:
            outsource_total += nd_labor
        row_revenue = nd_labor + nd_mat + nd_overhead + nd_tech
        net = _f(row.get('검사비', 0))
        if row_revenue > 0 and net > 0 and abs(net - row_revenue) < 10:
            nd_tech += (net - row_revenue)
            row_revenue = net
        elif row_revenue <= 0:
            row_revenue = net
        total_net_revenue += row_revenue
        total_travel += _f(row.get('출장비', 0))
        total_meal += _f(row.get('일식', 0))

        # [NEW] Identify logically unique entries (shared save event)
        # STRATEGIC FIX: Remove 'Note' from UID because split-rows often have different notes like '(차량 추가 기록)'
        # relying only on Date + Site + EntryTime (minute level) for grouping
        raw_time = str(row.get('EntryTime', '')).strip()
        e_time = raw_time[:16] if len(raw_time) > 16 else raw_time
        e_site = str(row.get('Site', '')).strip()

        # UID composite - date_val is already YYYY-MM-DD
        e_uid = f"{date_val}|{e_site}|{e_time}"

        is_new_logical_entry = (e_uid not in processed_entry_ids)
        processed_entry_ids.add(e_uid)

        # 명칭 불일치 방지 (세척제/세척액, 침투제/침투액, 현상제/현상액 혼용 대응)
        def get_first_valid(r, keys):
            for k in keys:
                v = _f(r.get(k, 0))
                if v > 0: return v
            return 0.0

        # [NOTE] Specialist NDT columns are now handled in the consolidated block below to avoid double counting

        # --- Vehicle Tracking ---
        m_id_local = row.get('MaterialID', '')
        # [FIX] If MaterialID is a manual name string (not in master map), use it directly as the name
        m_name_local = str(mat_id_name_map.get(m_id_local, m_id_local)).upper()

        # [FIX] Resolve vehicle number from master config to capture '(스타렉스)' tags
        v_no_raw = str(row.get('차량번호', '')).strip()
        vehicles_list = getattr(self, 'vehicles', [])
        vehicle_map = {v.split('(')[0].strip(): v for v in vehicles_list}
        resolved_v_no = vehicle_map.get(v_no_raw, v_no_raw)

        # [FIX] Check both Equipment Name (장비명) and Item Name (품목명) for vehicle detection
        car_info = f"{resolved_v_no} {row.get('차량비고', '')} {row.get('장비명', '')} {m_name_local}".strip()

        try:
            dt_obj = pd.to_datetime(date_val)
            date_key = dt_obj.date()
        except:
            date_key = date_val

        if is_new_logical_entry:
            if '스타렉스' in car_info:
                starex_dates.add(date_key)
            if '탑차' in car_info:
                toptruck_dates.add(date_key)

        # PAUT 및 특수장비 작업일수(중복일 제외) 집계
        try:
            method_val = str(row.get('검사방법', '')).upper()
            equip_name = str(row.get('장비명', '')).upper()

            # [FIX] Check both Equipment Name and Item Name for PAUT/MT detection
            is_paut = 'PAUT' in method_val or 'PAUT' in equip_name or 'PAUT' in m_name_local
            is_mt = 'MT' in method_val or 'MT' in equip_name or 'MT' in m_name_local or 'YOKE' in equip_name or 'YOKE' in m_name_local

            if is_paut:
                paut_dates.add(date_key)
                # Specialized Scanner Detection (Robust matching)
                eq_clean = equip_name.replace(' ', '').upper()
                aux_clean = str(row.get('보조장비', '')).replace(' ', '').upper()
                m_clean = m_name_local.replace(' ', '').upper()
                mid_clean = str(m_id_local).replace(' ', '').upper()
                note_clean = str(row.get('Note', row.get('비고', ''))).replace(' ', '').upper()
                cond_clean = str(row.get('조건1', '')).replace(' ', '').upper()
                pipe_clean = str(row.get('관경(Inch)', row.get('관경', ''))).replace(' ', '').upper()

                combined_text = f"{eq_clean}|{aux_clean}|{m_clean}|{mid_clean}|{note_clean}"

                is_cobra_scanner = ('COBRA' in combined_text or
                                    'SCANNER(COBRA)' in combined_text or
                                    ('100A이하' in cond_clean and ('PAUT' in method_val or 'PAUT' in equip_name)))

                # [FIX] If it's a scanner and not COBRA, it's MANUAL
                is_manual_scanner = ('SCANNER' in combined_text or 'JIREH' in combined_text) and not is_cobra_scanner

                if is_manual_scanner:
                    paut_manual_scanner_dates.add(date_key)
                if is_cobra_scanner:
                    paut_cobra_scanner_dates.add(date_key)

            if is_mt:
                mt_dates.add(date_key)
            is_rt = ('RT' in method_val or 'RT' in equip_name or 'RT' in m_name_local or
                     'SE-75' in equip_name or 'SE-175' in equip_name or
                     'SE-75' in aux_clean or 'SE-175' in aux_clean)
            if is_rt:
                rt_dates.add(date_key)
        except Exception:
            pass

        # --- Labor & OT Parsing ---
        shift = str(row.get('Shift', '주간')).strip()
        for i in range(1, 11):
            suffix = '' if i == 1 else str(i)
            raw_worker = str(row.get(f'User{suffix}', '')).strip()
            if not raw_worker or raw_worker.lower() == 'nan':
                continue

            # Rank Aggregation
            match = rank_pattern.search(raw_worker)
            worker_name_only = raw_worker
            rank = None

            if match:
                # Case 1: Rank identified via pattern (e.g., "[부장] 주진철" or "부장 주진철")
                if match.group(1) in ranks:
                    rank = match.group(1)
                    worker_name_only = match.group(2).strip()
                    if not worker_name_only: # Only rank was present
                         worker_name_only = rank

            # Case 2: Fallback to mapping if rank not identified from the string
            if not rank:
                # Clean potential shift prefixes if they exist (though usually cleaned in records)
                clean_name = _re.sub(r'\(.*?\)', '', raw_worker).strip()
                rank = worker_rank_map.get(clean_name)
                worker_name_only = clean_name

            if rank and rank in rank_labor_dates:
                if worker_name_only not in rank_labor_dates[rank]:
                    rank_labor_dates[rank][worker_name_only] = set()
                try:
                    # Normalize date key for uniqueness
                    dt_obj = pd.to_datetime(date_val)
                    d_key = dt_obj.date()
                except:
                    d_key = date_val
                rank_labor_dates[rank][worker_name_only].add(d_key)

            # [FIX] WorkTime 컬럼에서 분류(주간/야간/휴일)와 시간 파싱
            # 형식: "(야간) 4h" or "(주야간) 2.5h" 또는 "8h" (구형식)
            wt_col = f'WorkTime{suffix}'
            wt_val = str(row.get(wt_col, '')).strip()

            # WorkTime에서 분류 파싱
            wt_shift = shift  # 기본값: 레코드의 Shift 컬럼
            wt_hours = 0.0
            if wt_val and wt_val.lower() != 'nan':
                # Parse shift type if present, e.g. "(야간)", "(연장)"
                shift_match = _re.search(r'\(?(주간|야간|주야간|휴일|연장)\)?', wt_val)
                if shift_match:
                    wt_shift = shift_match.group(1)

                # Compute duration calculation
                marker_pattern = MARKER_PATTERN
                clean_wt = marker_pattern.sub('', wt_val).replace('(연장)', '').replace('(주야간)', '').replace('(야간)', '').replace('(휴일)', '').replace('(주간)', '').strip()
                sep_match = _re.search(r'[:\d]\s*([~-])\s*[:\d]', clean_wt)

                if sep_match:
                    try:
                        sep = sep_match.group(1)
                        st_str, en_str = clean_wt.split(sep)
                        sh, sm = map(int, st_str.split(':'))
                        eh, em = map(int, en_str.split(':'))
                        start_f = sh + sm/60.0
                        end_f = eh + em/60.0
                        if end_f < start_f: end_f += 24
                        wt_hours = end_f - start_f
                    except:
                        pass
                else:
                    # Fallback: Just a number like "8h" or "8"
                    h_match = _re.search(r'([\d.]+)', clean_wt)
                    if h_match:
                        try: wt_hours = float(h_match.group(1))
                        except: pass

            ot_signature = (
                str(date_val).strip(),
                e_site,
                worker_name_only,
                _re.sub(r'\s+', '', wt_val).lower(),
            )
            is_duplicate_ot = ot_signature in processed_ot_signatures
            if wt_hours > 0:
                processed_ot_signatures.add(ot_signature)

            # [FIX] 분류별 OT 상세 집계
            if not is_duplicate_ot:
                from site_apps.central.src.services.ot_calculator import ot_breakdown
                breakdown = ot_breakdown(wt_val, date_val, row.get(f'OT{suffix}', ''))
                for label, key in (('연장근무', 'day'), ('야간근무', 'night'), ('휴일근무', 'holiday')):
                    if breakdown[key] > 0:
                        ot_data[label]['hours'] += breakdown[key]
                        ot_data[label]['names'].add(worker_name_only)

        # --- 자재 사용량 집계 (NDT + 필름) ---
        if is_new_logical_entry:
            m_id = row.get('MaterialID', '')
            m_name = str(mat_id_name_map.get(m_id, '')).strip()
            test_method = str(row.get('검사방법', '')).upper()
            usage = _f(row.get('Usage', 0))

            if 'RT' in test_method:
                rt_inspection_count += 1
                film_qty = get_first_valid(row, ['필름매수', 'FilmCount', 'Usage', '검사량'])
                total_film_count += film_qty
                material_usage_sums[5] += film_qty  # Index 5 = 필름
                total_mat_cost += film_qty * float(mat_id_cost_map.get(m_id, 0))
            else:
                # UT/PT/MT 검사: NDT 화학약품 등의 일반 자재 사용
                # Specialist columns take priority over the generic 'Usage' field
                ndt_vals = [
                    get_first_valid(row, ['NDT_세척제', '세척제', 'NDT_세척액', '세척액']),
                    get_first_valid(row, ['NDT_침투제', '침투제', 'NDT_침투액', '침투액']),
                    get_first_valid(row, ['NDT_현상제', '현상제', 'NDT_현상액', '현상액']),
                    get_first_valid(row, ['NDT_흑색자분', 'NDT_자분', '흑색자분', 'NDT_페인트']),
                    get_first_valid(row, ['NDT_백색페인트', '백색페인트']) # Placeholder mapping adjustment
                ]

                has_specialist_data = any(v > 0 for v in ndt_vals)

                if has_specialist_data:
                    for idx, v in enumerate(ndt_vals):
                        if idx < 5: material_usage_sums[idx] += v
                    # Cost calculation still needs a MaterialID context, fallback to usage row if available
                    total_mat_cost += usage * float(mat_id_cost_map.get(m_id, 0))
                elif usage > 0:
                    # Fallback to Usage field matching if specialist columns are empty
                    # 세척제: 청소/세척 용도 → Index 0
                    if '세척제' in m_name or '세척' in m_name or '청소' in m_name:
                        material_usage_sums[0] += usage
                    # 침투제: 침투액 → Index 1
                    elif '침투제' in m_name or '침투액' in m_name:
                        material_usage_sums[1] += usage
                    # 현상제: 현상용 시약 → Index 2
                    elif '현상제' in m_name or '현상액' in m_name:
                        material_usage_sums[2] += usage
                    # 자분: 자기 자분 → Index 3
                    elif '자분' in m_name and '흑색' not in m_name:
                        material_usage_sums[3] += usage
                    # 흑색자분 → Index 4
                    elif '흑색자분' in m_name or ('흑색' in m_name and '자분' in m_name):
                        material_usage_sums[4] += usage
                    # 글리세린 → Index 6
                    elif '글리세' in m_name or '글리세린' in m_name:
                        material_usage_sums[6] += usage
                    # 정착액 → Index 8
                    elif '정착액' in m_name or '정착' in m_name:
                        material_usage_sums[8] += usage
                    # 수적 → Index 9
                    elif '수적' in m_name:
                        material_usage_sums[9] += usage
                    # 기타 PT/MT 화학약품이면 자동 매핑
                    elif any(k in m_name.upper() for k in ['PENETRANT', 'DEVELOPER', 'CLEANER', 'NABAKEM', 'PT', 'MT']):
                        matched = False
                        for keyword, idx in mat_map.items():
                            if keyword.lower() in m_name.lower():
                                material_usage_sums[idx] += usage
                                matched = True
                                break
                        if not matched:
                            material_usage_sums[0] += usage
                    total_mat_cost += usage * float(mat_id_cost_map.get(m_id, 0))

    # 3. 폼에 입력 및 [5. 영업이익] '실행(실적)' 줄에 표시 결과 연동
    # Update Special Unit Prices if LaborCostDetailWidget is available to compute correct revenue
    ot_revenue = 0.0
    if hasattr(self, 'labor_detail_widget'):
        ot_revenue = sum(ot_data[t]['hours'] * _f(self.labor_detail_widget.entries[t]['unit_price'].get()) 
                        for t in special_types if t in self.labor_detail_widget.entries)

    # 3. 폼에 입력 (사후실적 칸만 업데이트)
    self.cb_budget_site.set(site)

    # 도급액(Revenue)
    if hasattr(self, 'ent_budget_actual_revenue'):
        self.ent_budget_actual_revenue.delete(0, tk.END)
        self.ent_budget_actual_revenue.insert(0, f"{total_net_revenue:,.0f}")

    # 검사단가
    if hasattr(self, 'ent_budget_actual_unit_price'):
        self.ent_budget_actual_unit_price.delete(0, tk.END)
        self.ent_budget_actual_unit_price.insert(0, f"{total_net_revenue:,.0f}")

    # 재료비
    if hasattr(self, 'ent_budget_actual_material'):
        self.ent_budget_actual_material.delete(0, tk.END)
        self.ent_budget_actual_material.insert(0, f"{np.nan_to_num(total_mat_cost):,.0f}")

    # 경비
    if hasattr(self, 'ent_budget_actual_expense'):
        self.ent_budget_actual_expense.delete(0, tk.END)
        self.ent_budget_actual_expense.insert(0, f"{np.nan_to_num(total_travel + total_meal):,.0f}")

    # 노무비 (lab_total is already calculated from df)
    if hasattr(self, 'ent_budget_actual_labor'):
        self.ent_budget_actual_labor.delete(0, tk.END)
        self.ent_budget_actual_labor.insert(0, f"{np.nan_to_num(lab_total):,.0f}")

    # 외주비: 작업자 없는 행의 인건비 합산
    if hasattr(self, 'ent_budget_actual_outsource'):
        self.ent_budget_actual_outsource.delete(0, tk.END)
        self.ent_budget_actual_outsource.insert(0, f"{np.nan_to_num(outsource_total):,.0f}")

    # 4. 실적 인건비 상세 탭 누적 업데이트
    if hasattr(self, 'actual_labor_detail_widget'):
        labor_data = self.actual_labor_detail_widget.get_data()
        planned_labor_data = self.labor_detail_widget.get_data() if hasattr(self, 'labor_detail_widget') else {}
        for rank in ranks:
            if rank in labor_data:
                u_count = len(rank_labor_dates[rank])
                t_days = sum(len(dates) for dates in rank_labor_dates[rank].values())
                if u_count > 0:
                    avg_d = t_days / u_count
                    labor_data[rank]['personnel'] = f"{u_count:g}"
                    labor_data[rank]['period'] = f"{avg_d:g}"
                else:
                    labor_data[rank]['personnel'] = ""
                    labor_data[rank]['period'] = ""
                # [FIX] 사전예산 단가 복사
                if rank in planned_labor_data:
                    labor_data[rank]['unit_price'] = planned_labor_data[rank].get('unit_price', '')
        for stype in special_types:
            if stype in labor_data:
                u_count = len(ot_data[stype]['names'])
                t_hours = ot_data[stype]['hours']
                if u_count > 0:
                    avg_h = t_hours / u_count
                    labor_data[stype]['personnel'] = f"{u_count:g}"
                    labor_data[stype]['period'] = f"{avg_h:.15g}"
                else:
                    labor_data[stype]['personnel'] = ""
                    labor_data[stype]['period'] = ""
                # [FIX] 사전예산 단가 복사
                if stype in planned_labor_data:
                    labor_data[stype]['unit_price'] = str({'연장근무':4000, '야간근무':5000, '휴일근무':7500}[stype])
        self.actual_labor_detail_widget.set_data(labor_data)

    # 5. 실적 자재비 상세 탭 누적 업데이트
    if hasattr(self, 'actual_material_detail_widget'):
        current_mat_data = self.actual_material_detail_widget.get_data()
        planned_mat_data = self.material_detail_widget.get_data() if hasattr(self, 'material_detail_widget') else []
        film_chem_qty = (total_film_count / 250.0) if total_film_count > 0 else 0.0
        anti_drop_qty = (total_film_count / 500.0) if total_film_count > 0 else 0.0
        if len(material_usage_sums) > 7:
            material_usage_sums[7] = film_chem_qty
        if len(material_usage_sums) > 8:
            material_usage_sums[8] = film_chem_qty
        if len(material_usage_sums) > 9:
            material_usage_sums[9] = anti_drop_qty
        if len(material_usage_sums) > 10:
            # 동위원소(선원) 구매는 작업일수가 아닌 장비 투입 실적 기준 1대(1개) 구매
            material_usage_sums[10] = 1.0 if len(rt_dates) > 0 else 0.0

        for i, qty in enumerate(material_usage_sums):
            if i < len(current_mat_data):
                current_mat_data[i]['qty'] = f"{qty:g}" if qty > 0 else ""
                # [FIX] 사전예산 단가 복사
                if i < len(planned_mat_data):
                    current_mat_data[i]['price'] = planned_mat_data[i].get('price', '')
        self.actual_material_detail_widget.set_data(current_mat_data)

    # 6. 실적 경비 상세 탭 누적 업데이트
    if hasattr(self, 'actual_expense_detail_widget'):
        current_exp_data = self.actual_expense_detail_widget.get_data()
        planned_exp_data = self.expense_detail_widget.get_data() if hasattr(self, 'expense_detail_widget') else {}

        # [NEW] 차량 대수 및 누적 투입 일수 계산 (차량유지비용 및 감가상각비)
        vehicle_dates_map = {}
        vehicles_list = getattr(self, 'vehicles', [])
        vehicle_map = {v.split('(')[0].strip(): v for v in vehicles_list}

        for _, row in df.iterrows():
            raw_v_no = str(row.get('차량번호', '')).strip()
            date_val = str(row.get('Date', '')).strip()
            if raw_v_no and raw_v_no.lower() not in ['nan', 'none']:
                # Split by || or comma in case multiple vehicles are entered in one cell
                v_list = [v.strip() for v in raw_v_no.replace('||', ',').split(',') if v.strip()]

                try:
                    dt_obj = pd.to_datetime(date_val)
                    date_key = dt_obj.date()
                except:
                    date_key = date_val

                for v_no in v_list:
                    resolved_v = vehicle_map.get(v_no, v_no)
                    if resolved_v not in vehicle_dates_map:
                        vehicle_dates_map[resolved_v] = set()
                    vehicle_dates_map[resolved_v].add(date_key)

        total_vehicle_days = sum(len(dates) for dates in vehicle_dates_map.values())
        starex_days = sum(len(dates) for v, dates in vehicle_dates_map.items() if '스타렉스' in v)
        toptruck_days = sum(len(dates) for v, dates in vehicle_dates_map.items() if '탑차' in v)

        for i, row_data in enumerate(current_exp_data.get('site_expense', [])):
            # [FIX] 사전예산 단가/규격 복사 (수량은 사용자가 직접 입력)
            if 'site_expense' in planned_exp_data and i < len(planned_exp_data['site_expense']):
                row_data['price'] = planned_exp_data['site_expense'][i].get('price', '')
                row_data['unit'] = planned_exp_data['site_expense'][i].get('unit', '')

            # 차량유지비는 작업 행 수가 아니라 차량별 고유 사용일수로 반영한다.
            # 같은 날짜/현장의 700A, 200A 작업에 같은 차량이 반복 저장돼도 1일이다.
            actual_days = len(set(str(r.get('Date', '')).strip().split(' ')[0] for _, r in df.iterrows() if str(r.get('Date', '')).strip()))
            cat_name = str(row_data.get('cat', '')).strip()
            cat_upper = cat_name.upper()
            item_upper = str(row_data.get('item', '')).upper()
            if cat_name == '차량유지비':
                row_data['ppl'] = 'N/A'
                row_data['qty'] = f"{total_vehicle_days:g}" if total_vehicle_days > 0 else ""
                row_data['unit'] = '일'
                row_data['price'] = 6667
            elif cat_name == '소모품비':
                row_data['qty'] = f"{actual_days:g}" if actual_days > 0 else ""
                row_data['unit'] = '일'
                row_data['price'] = 3333
            elif cat_name == '복리후생비':
                row_data['qty'] = f"{actual_days:g}" if actual_days > 0 else ""
                row_data['unit'] = '일'
                row_data['price'] = 1667
            elif any(k in cat_upper or k in item_upper for k in ['SE-175', 'SE175', 'SE-75', 'SE75', '방사성동위원소', '동위원소']):
                row_data['ppl'] = 'N/A'
                row_data['qty'] = f"{len(rt_dates):g}" if len(rt_dates) > 0 else ""
                row_data['unit'] = '일'
                row_data['price'] = 35714

        for i, row_data in enumerate(current_exp_data.get('depreciation', [])):
            item = row_data.get('item', '')

            updated_days = False
            if '스타렉스' in item:
                row_data['days'] = f"{starex_days:g}" if starex_days > 0 else ""
                if starex_days > 0: updated_days = True
            elif '탑차' in item:
                row_data['days'] = f"{toptruck_days:g}" if toptruck_days > 0 else ""
                if toptruck_days > 0: updated_days = True
            else:
                item_upper_clean = str(item).upper().replace(' ', '')
                if 'SCANNER(MANUAL)' in item_upper_clean:
                    row_data['days'] = f"{len(paut_manual_scanner_dates):g}" if len(paut_manual_scanner_dates) > 0 else ""
                    if len(paut_manual_scanner_dates) > 0: updated_days = True
                elif 'SCANNER(COBRA)' in item_upper_clean:
                    row_data['days'] = f"{len(paut_cobra_scanner_dates):g}" if len(paut_cobra_scanner_dates) > 0 else ""
                    if len(paut_cobra_scanner_dates) > 0: updated_days = True
                elif 'PAUT' in item_upper_clean:
                    row_data['days'] = f"{len(paut_dates):g}" if len(paut_dates) > 0 else ""
                    if len(paut_dates) > 0: updated_days = True
                elif 'YOKE' in str(item).upper() or ('MT' in str(item).upper() and 'PAUT' not in str(item).upper()):
                    row_data['days'] = f"{len(mt_dates):g}" if len(mt_dates) > 0 else ""
                    if len(mt_dates) > 0: updated_days = True
                else:
                    item_upper = str(item).upper().strip()
                    matched_days = 0
                    for e_name, e_dates in equip_dates_map.items():
                        if e_name.upper().strip() in item_upper or item_upper in e_name.upper().strip():
                            matched_days = max(matched_days, len(e_dates))
                    if matched_days > 0:
                        row_data['days'] = f"{matched_days:g}"
                        updated_days = True

            # [FIX] 사전예산 단가 복사 및 수량 처리
            if 'depreciation' in planned_exp_data and i < len(planned_exp_data['depreciation']):
                row_data['rate'] = planned_exp_data['depreciation'][i].get('rate', '')

                if updated_days:
                    # 일수가 자동 합산된 항목은 중복곱셈을 막기 위해 수량을 무조건 1로 강제
                    row_data['qty'] = '1'
                else:
                    # 자동 합산되지 않은 항목은 빈칸일 때만 사전예산 수량 복사
                    current_qty = str(row_data.get('qty', '')).strip()
                    planned_qty = str(planned_exp_data['depreciation'][i].get('qty', '')).strip()
                    if not current_qty:
                        row_data['qty'] = planned_qty if planned_qty else '1'
            else:
                if updated_days or not str(row_data.get('qty', '')).strip():
                    row_data['qty'] = '1'

        self.actual_expense_detail_widget.set_data(current_exp_data)

    # KPI & Diff Update
    self._update_budget_kpis()
    self.root.update_idletasks()

    # [NEW] Calculate profit/margin for the summary message
    try:
        a_rev = float(self.ent_budget_actual_unit_price.get().replace(',', '') or 0)
        if a_rev == 0:
            a_rev = float(self.ent_budget_actual_revenue.get().replace(',', '') or 0)
        a_lab = float(self.ent_budget_actual_labor.get().replace(',', '') or 0)
        a_mat = float(self.ent_budget_actual_material.get().replace(',', '') or 0)
        a_exp = float(self.ent_budget_actual_expense.get().replace(',', '') or 0)
        a_out = float(self.ent_budget_actual_outsource.get().replace(',', '') or 0)
        a_cost = a_lab + a_mat + a_exp + a_out
        a_profit = a_rev - a_cost
        a_margin = (a_profit / a_rev * 100) if a_rev > 0 else 0.0

        if hasattr(self, 'ent_budget_actual_profit'):
            self.ent_budget_actual_profit.delete(0, tk.END)
            self.ent_budget_actual_profit.insert(0, f"{a_profit:,.0f} ({a_margin:.1f}%)")

        profit_str = f"\n\n▶ 실적 집계 결과\n- 도급액(실적): {a_rev:,.0f}원\n- 사후원가: {a_cost:,.0f}원\n- 현재이익: {a_profit:,.0f}원 ({a_margin:.1f}%)"
    except:
        profit_str = ""

    messagebox.showinfo("조회 완료", 
                        f"'{site}' 현장의 실적 데이터를 집계하여 [사후원가(실적)] 칸에 표시하였습니다.\n"
                        f"사전예산은 방어되어 변경되지 않았습니다.\n"
                        f"(기간: {start_ts.strftime('%Y-%m-%d')} ~ {end_ts.strftime('%Y-%m-%d')}){profit_str}")

import json
import shutil
import traceback
from tkinter import simpledialog


def export_budget_estimation_impl(self):
    """Export the estimated budget details (Planned & Actuals) to Excel"""
    site = self.cb_budget_site.get()
    if not site:
        messagebox.showwarning("현장 미선택", "실행예산서를 내보낼 현장을 선택하거나 입력해주세요.")
        return

    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        initialfile=f"공사실행예산정산서_{site}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )
    if not file_path:
        return

    # Collect Summary Data (Plan, Actual, Diff)
    def get_val(attr):
        return getattr(self, attr).get().strip() if hasattr(self, attr) else ""

    summary_data = {
        'site': site,
        'period': get_val('ent_budget_period'),
        'actual_period': get_val('ent_budget_actual_period'),
        'diff_period': get_val('ent_budget_diff_period'),

        'revenue': get_val('ent_budget_revenue'),
        'actual_revenue': get_val('ent_budget_actual_revenue'),
        'diff_revenue': get_val('ent_budget_diff_revenue'),

        'unitprice': get_val('ent_budget_unit_price'),
        'actual_unitprice': get_val('ent_budget_actual_unit_price'),
        'diff_unitprice': get_val('ent_budget_diff_unit_price'),

        'labor': get_val('ent_budget_labor'),
        'actual_labor': get_val('ent_budget_actual_labor'),
        'diff_labor': get_val('ent_budget_diff_labor'),

        'material': get_val('ent_budget_material'),
        'actual_material': get_val('ent_budget_actual_material'),
        'diff_material': get_val('ent_budget_diff_material'),

        'expense': get_val('ent_budget_expense'),
        'actual_expense': get_val('ent_budget_actual_expense'),
        'diff_expense': get_val('ent_budget_diff_expense'),

        'outsource': get_val('ent_budget_outsource'),
        'actual_outsource': get_val('ent_budget_actual_outsource'),
        'diff_outsource': get_val('ent_budget_diff_outsource'),

        'profit': get_val('ent_budget_profit'),
        'actual_profit': get_val('ent_budget_actual_profit'),
        'diff_profit': get_val('ent_budget_diff_profit'),

        'margin': get_val('ent_budget_margin'),
        'actual_margin': get_val('ent_budget_actual_margin'),
        'diff_margin': get_val('ent_budget_diff_margin')
    }

    # Details - Planned
    labor_data = self.labor_detail_widget.get_data() if hasattr(self, 'labor_detail_widget') else {}

    raw_mat = self.material_detail_widget.get_data() if hasattr(self, 'material_detail_widget') else []
    material_data = []
    defaults = self.material_detail_widget.default_items if hasattr(self, 'material_detail_widget') else []
    for i, row in enumerate(raw_mat):
        if i < len(defaults):
            name, spec, unit, _ = defaults[i]
            row['name'] = name
            row['spec'] = spec
            row['unit'] = unit
            material_data.append(row)

    expense_data = self.expense_detail_widget.get_data() if hasattr(self, 'expense_detail_widget') else {}

    # Details - Actuals
    actual_labor_data = self.actual_labor_detail_widget.get_data() if hasattr(self, 'actual_labor_detail_widget') else {}

    raw_act_mat = self.actual_material_detail_widget.get_data() if hasattr(self, 'actual_material_detail_widget') else []
    actual_material_data = []
    act_defaults = getattr(self.actual_material_detail_widget, 'default_items', defaults)
    for i, row in enumerate(raw_act_mat):
        if i < len(act_defaults):
            name, spec, unit, _ = act_defaults[i]
            row['name'] = name
            row['spec'] = spec
            row['unit'] = unit
            actual_material_data.append(row)

    actual_expense_data = self.actual_expense_detail_widget.get_data() if hasattr(self, 'actual_expense_detail_widget') else {}

    # Base salaries for reference
    base_salaries = self.get_base_salaries() if hasattr(self, 'get_base_salaries') else getattr(self.labor_detail_widget, 'base_salaries', {})

    try:
        from site_apps.central.src.utils.export_helper import export_budget_estimation_to_excel
        if export_budget_estimation_to_excel(
            file_path,
            summary_data,
            labor_data,
            material_data,
            expense_data,
            actual_labor_data=actual_labor_data,
            actual_material_data=actual_material_data,
            actual_expense_data=actual_expense_data,
            base_salaries=base_salaries
        ):
            messagebox.showinfo("내보내기 완료", f"공사실행예산 및 정산서가 성공적으로 저장되었습니다.\n{file_path}")
    except Exception as e:
        import traceback
        traceback.print_exc()
        messagebox.showerror("오류", f"엑셀 저장 중 오류가 발생했습니다.\n{e}")


def export_budget_sales_status_impl(self):
    """Export the currently filtered budget sales status to Excel"""
    if not hasattr(self, 'budget_view_tree') or not self.budget_view_tree.get_children():
        messagebox.showwarning("데이터 없음", "내보낼 데이터가 없습니다. 먼저 조회를 해주세요.")
        return

    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        initialfile=f"매출현황_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )

    if not file_path:
        return

    # Collect data from treeview
    data = []
    columns = [self.budget_view_tree.heading(col)['text'] for col in self.budget_view_tree['columns']]
    for item in self.budget_view_tree.get_children():
        data.append(self.budget_view_tree.item(item)['values'])

    df = pd.DataFrame(data, columns=columns)

    # Add summary row at the bottom
    summary_row = {col: '' for col in columns}
    summary_row['날짜'] = '합계'
    if hasattr(self, 'lbl_bv_actual_income'):
        summary_row['검사단가'] = self.lbl_bv_actual_income.cget('text').replace(' 원', '').replace(',', '').strip()

    # Compute OT/Total directly from columns since Budget KPIs might show differently
    ot_idx = columns.index('OT합계') if 'OT합계' in columns else -1
    total_idx = columns.index('합계') if '합계' in columns else -1

    sum_ot = 0
    sum_tot = 0
    for _d in data:
        if ot_idx >= 0:
            try: sum_ot += float(str(_d[ot_idx]).replace(',', ''))
            except: pass
        if total_idx >= 0:
            try: sum_tot += float(str(_d[total_idx]).replace(',', ''))
            except: pass

    summary_row['OT합계'] = f"{sum_ot:g}"
    summary_row['합계'] = f"{sum_tot:g}"

    # Add the summary row to the data list for the DataFrame
    data.append([summary_row.get(col, '') for col in columns])

    df = pd.DataFrame(data, columns=columns)

    try:
        from openpyxl.utils import get_column_letter
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='SalesStatus', index=False)
            # Auto-adjust columns
            worksheet = writer.sheets['SalesStatus']
            for i, col in enumerate(df.columns):
                max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                col_letter = get_column_letter(i + 1)
                worksheet.column_dimensions[col_letter].width = max_len

        messagebox.showinfo("내보내기 완료", "매출현황이 엑셀로 저장되었습니다.")
    except Exception as e:
        messagebox.showerror("오류", f"엑셀 저장 중 오류가 발생했습니다: {e}")


def _refresh_budget_view_tree_columns_impl(self, reload_data=False):
    """Rebuild budget site performance tree column structure."""
    if not hasattr(self, 'budget_view_tree'):
        return

    custom_cols = getattr(self, 'budget_view_custom_columns', []) or []
    # [FIX] built-in 컬럼과 이름/키가 중복되는 사용자 컬럼 정리
    builtin_keys = set(getattr(self, 'budget_view_builtin_cols', []) or [])
    builtin_headings = set((getattr(self, 'budget_view_builtin_head_map', {}) or {}).values())
    sanitized_custom_cols = []
    seen_custom_keys = set()
    seen_custom_names = set()
    for c in custom_cols:
        key = str(c.get('key', '')).strip()
        name = str(c.get('name', key)).strip()
        if not key:
            continue
        # built-in 과 키 충돌/이름 충돌이면 제거
        if key in builtin_keys or name in builtin_headings:
            continue
        # custom 내부 중복도 제거
        if key in seen_custom_keys or name in seen_custom_names:
            continue
        seen_custom_keys.add(key)
        seen_custom_names.add(name)
        sanitized_custom_cols.append(c)
    if len(sanitized_custom_cols) != len(custom_cols):
        self.budget_view_custom_columns = sanitized_custom_cols
        custom_cols = sanitized_custom_cols

    custom_keys = [c.get('key') for c in custom_cols if c.get('key')]
    self.budget_view_cols = tuple(list(self.budget_view_builtin_cols) + custom_keys)

    head_map = dict(getattr(self, 'budget_view_builtin_head_map', {}))
    width_map = dict(getattr(self, 'budget_view_builtin_width_map', {}))
    for col in custom_cols:
        key = col.get('key')
        if not key:
            continue
        head_map[key] = col.get('name', key)
        width_map[key] = int(col.get('width', 120) or 120)

    alias_map = getattr(self, 'budget_view_heading_aliases', {}) or {}
    self.budget_view_tree.configure(columns=self.budget_view_cols)
    for col in self.budget_view_cols:
        heading_text = alias_map.get(col, head_map.get(col, col))
        self.budget_view_tree.heading(col, text=heading_text,
                                      command=lambda c=col: self.treeview_sort_column(self.budget_view_tree, c, False))
        self.budget_view_tree.column(col, width=width_map.get(col, 100), minwidth=40, anchor='center', stretch=False)

    visible_source = getattr(self, 'budget_view_visible_cols', []) or list(self.budget_view_default_cols)
    visible = [c for c in visible_source if c in self.budget_view_cols]
    self.budget_view_tree['displaycolumns'] = visible if visible else self.budget_view_default_cols
    self.budget_view_visible_cols = list(self.budget_view_tree['displaycolumns'])

    if reload_data:
        self.update_budget_site_view()


def add_budget_view_custom_column_impl(self):
    """Add a custom column to budget site performance view."""
    name = simpledialog.askstring("컬럼 추가", "새 컬럼 이름을 입력하세요:", parent=self.root)
    if not name or not str(name).strip():
        return
    name = str(name).strip()

    # [FIX] built-in/기존 사용자 컬럼과 이름 중복 방지
    builtin_names = set((getattr(self, 'budget_view_builtin_head_map', {}) or {}).values())
    existing_custom_names = {str(c.get('name', '')).strip() for c in (getattr(self, 'budget_view_custom_columns', []) or [])}
    if name in builtin_names or name in existing_custom_names:
        messagebox.showwarning("중복 컬럼", f"'{name}' 컬럼은 이미 존재합니다.")
        return

    default_value = simpledialog.askstring("컬럼 내용", f"'{name}' 컬럼의 기본 내용을 입력하세요(선택):", parent=self.root)

    existing = set(getattr(self, 'budget_view_cols', []))
    key_base = re.sub(r'[^0-9A-Za-z가-힣_]+', '_', name).strip('_') or '사용자컬럼'
    key = key_base
    idx = 1
    while key in existing:
        idx += 1
        key = f"{key_base}_{idx}"

    custom_cols = getattr(self, 'budget_view_custom_columns', []) or []
    custom_cols.append({'key': key, 'name': name, 'default': default_value or '', 'width': 120})
    self.budget_view_custom_columns = custom_cols

    visible = list(getattr(self, 'budget_view_visible_cols', []) or list(self.budget_view_default_cols))
    if key not in visible:
        visible.append(key)
    self.budget_view_visible_cols = visible

    self._refresh_budget_view_tree_columns(reload_data=True)
    self.save_tab_config()


def rename_budget_view_column_impl(self, col_key):
    """Rename a budget site performance column heading."""
    if not col_key:
        return
    current_name = self.budget_view_tree.heading(col_key)['text'] if hasattr(self, 'budget_view_tree') else col_key
    new_name = simpledialog.askstring("컬럼 이름 변경", "새 컬럼 이름을 입력하세요:", initialvalue=current_name, parent=self.root)
    if not new_name or not str(new_name).strip():
        return
    new_name = str(new_name).strip()

    # Custom column: rename the stored name itself
    updated = False
    for col in getattr(self, 'budget_view_custom_columns', []) or []:
        if col.get('key') == col_key:
            col['name'] = new_name
            updated = True
            break
    if not updated:
        aliases = getattr(self, 'budget_view_heading_aliases', {}) or {}
        default_name = getattr(self, 'budget_view_builtin_head_map', {}).get(col_key, col_key)
        if new_name == default_name:
            aliases.pop(col_key, None)
        else:
            aliases[col_key] = new_name
        self.budget_view_heading_aliases = aliases

    self._refresh_budget_view_tree_columns(reload_data=False)
    self.save_tab_config()


def set_budget_view_custom_column_content_impl(self, col_key):
    """Change default content for a custom column."""
    custom_cols = getattr(self, 'budget_view_custom_columns', []) or []
    target = next((c for c in custom_cols if c.get('key') == col_key), None)
    if not target:
        return
    new_value = simpledialog.askstring("컬럼 내용 설정", f"'{target.get('name', col_key)}' 컬럼의 기본 내용을 입력하세요:",
                                       initialvalue=target.get('default', ''), parent=self.root)
    if new_value is None:
        return
    target['default'] = new_value
    self._refresh_budget_view_tree_columns(reload_data=True)
    self.save_tab_config()


def delete_budget_view_custom_column_impl(self, col_key):
    """Delete a custom budget column."""
    custom_cols = getattr(self, 'budget_view_custom_columns', []) or []
    target = next((c for c in custom_cols if c.get('key') == col_key), None)
    if not target:
        return
    if not messagebox.askyesno("컬럼 삭제", f"'{target.get('name', col_key)}' 컬럼을 삭제하시겠습니까?", parent=self.root):
        return

    self.budget_view_custom_columns = [c for c in custom_cols if c.get('key') != col_key]
    self.budget_view_visible_cols = [c for c in getattr(self, 'budget_view_visible_cols', []) if c != col_key]
    aliases = getattr(self, 'budget_view_heading_aliases', {}) or {}
    aliases.pop(col_key, None)
    self.budget_view_heading_aliases = aliases
    self._refresh_budget_view_tree_columns(reload_data=True)
    self.save_tab_config()


def show_budget_view_column_dialog_impl(self):
    """Open dialog to show/hide columns in budget site performance view"""
    if not hasattr(self, 'budget_view_tree'):
        return

    all_cols = list(self.budget_view_tree['columns'])
    active_cols = self.budget_view_tree['displaycolumns']
    if not active_cols or active_cols == ('#all'):
        active_cols = all_cols

    dialog_cols = [(col, self.budget_view_tree.heading(col)['text']) for col in all_cols]
    dialog = ColumnSelectionDialog(self.root, dialog_cols, title="현장별 실적 표시 컬럼 관리")
    for col, var in dialog.vars.items():
        var.set(col in active_cols)

    dialog.wait_window()

    if dialog.result is not None:
        # 필수 컬럼 보정: 날짜/현장은 항상 유지
        final_selection = list(dialog.result)
        for mandatory in ('Date', 'Site'):
            if mandatory in all_cols and mandatory not in final_selection:
                final_selection.append(mandatory)

        sorted_selection = [c for c in all_cols if c in final_selection]
        self.budget_view_tree['displaycolumns'] = sorted_selection if sorted_selection else all_cols
        self.budget_view_visible_cols = list(self.budget_view_tree['displaycolumns'])
        self.save_tab_config()


def _show_budget_view_heading_context_menu_impl(self, event):
    """공사탭 현장별 실적 Treeview 헤더 우클릭 메뉴"""
    if not hasattr(self, 'budget_view_tree'):
        return
    tree = self.budget_view_tree
    try:
        if tree.identify_region(event.x, event.y) != 'heading':
            return
    except Exception:
        return

    column_id = tree.identify_column(event.x)
    col_key = self._get_column_name_from_id(tree, column_id)
    if not col_key:
        return
    col_text = tree.heading(col_key)['text']
    custom_keys = {c.get('key') for c in (getattr(self, 'budget_view_custom_columns', []) or [])}

    menu = tk.Menu(self.root, tearoff=0)
    menu.add_command(label=f"⬅️ '{col_text}' 왼쪽으로 이동", command=lambda: self._move_column_visual(tree, column_id, -1))
    menu.add_command(label=f"➡️ '{col_text}' 오른쪽으로 이동", command=lambda: self._move_column_visual(tree, column_id, 1))
    menu.add_separator()
    menu.add_command(label=f"✏️ '{col_text}' 이름 변경...", command=lambda: self.rename_budget_view_column(col_key))
    menu.add_command(label="➕ 사용자 컬럼 추가...", command=self.add_budget_view_custom_column)
    if col_key in custom_keys:
        menu.add_command(label=f"📝 '{col_text}' 내용 설정...", command=lambda: self.set_budget_view_custom_column_content(col_key))
        menu.add_command(label=f"🗑️ '{col_text}' 사용자 컬럼 삭제", command=lambda: self.delete_budget_view_custom_column(col_key))
    menu.add_separator()
    menu.add_command(label="⚙️ 컬럼 관리(추가/삭제)...", command=self.show_budget_view_column_dialog)
    try:
        menu.tk_popup(event.x_root, event.y_root)
    finally:
        menu.grab_release()


def _update_budget_kpis_impl(self):
    """Update the top KPI summary labels based on current budget form values"""
    try:
        def _get_val(widget_name):
            try: 
                w = getattr(self, widget_name, None)
                if w:
                    v = str(w.get()).replace(',', '').split(' ')[0]
                    return float(v or 0)
                return 0.0
            except: return 0.0

        def _set_val(widget_name, val, is_percent=False):
            w = getattr(self, widget_name, None)
            if w:
                st = w.cget('state')
                if st == 'readonly': w.config(state='normal')
                w.delete(0, 'end')
                w.insert(0, f"{val:,.1f}" if is_percent else f"{val:,.0f}")
                if st == 'readonly': w.config(state='readonly')

        def _calc_profit(prefix):
            rev = _get_val(f'{prefix}revenue')
            if rev == 0: rev = _get_val(f'{prefix}unit_price')
            lab = _get_val(f'{prefix}labor')
            mat = _get_val(f'{prefix}material')
            exp = _get_val(f'{prefix}expense')
            out = _get_val(f'{prefix}outsource')

            if prefix == 'ent_budget_' and hasattr(self, 'expense_detail_widget') and hasattr(self.expense_detail_widget, 'lbl_grand_total_cost'):
                try:
                    raw_t = self.expense_detail_widget.lbl_grand_total_cost.cget('text')
                    tc = float("".join(c for c in raw_t if c.isdigit() or c == '.') or 0)
                except: tc = lab + mat + exp + out
            else:
                tc = lab + mat + exp + out

            prof = rev - tc
            mar = (prof / rev * 100) if rev > 0 else 0.0
            _set_val(f'{prefix}profit', prof)
            _set_val(f'{prefix}margin', mar, True)
            return rev, lab, mat, exp, out, prof, mar

        p_r, p_l, p_m, p_e, p_o, p_p, p_mg = _calc_profit('ent_budget_')
        a_r, a_l, a_m, a_e, a_o, a_p, a_mg = _calc_profit('ent_budget_actual_')

        _set_val('ent_budget_diff_revenue', p_r - a_r)
        _set_val('ent_budget_diff_unit_price', _get_val('ent_budget_unit_price') - _get_val('ent_budget_actual_unit_price'))
        _set_val('ent_budget_diff_labor', p_l - a_l)
        _set_val('ent_budget_diff_material', p_m - a_m)
        _set_val('ent_budget_diff_expense', p_e - a_e)
        _set_val('ent_budget_diff_outsource', p_o - a_o)
        _set_val('ent_budget_diff_profit', p_p - a_p)
        _set_val('ent_budget_diff_margin', p_mg - a_mg, True)

        if hasattr(self, 'lbl_kpi_rev'):
            self.lbl_kpi_rev.config(text=f"{p_r:,.0f}원")
            self.lbl_kpi_cost.config(text=f"{p_l+p_m+p_e+p_o:,.0f}원")
            self.lbl_kpi_profit.config(text=f"{p_p:,.0f}원", foreground="#ef4444" if p_p < 0 else "#10b981")
            self.lbl_kpi_margin.config(text=f"{p_mg:.1f}%", foreground="#ef4444" if p_mg < 0 else "#10b981")
        self.root.update_idletasks()
    except Exception as e:
        print(f"DEBUG: Budget KPI sync error: {e}")


def open_detached_budget_view_impl(self):
    """공사실행예산서 탭을 별도의 팝업창으로 엽니다. (메인 화면과 유사한 KPI/목록 구성)"""
    if 'budget' in self.detached_windows and self.detached_windows['budget']['window'].winfo_exists():
        self.detached_windows['budget']['window'].lift()
        return

    popup = tk.Toplevel(self.root)
    popup.title("📋 공사 실행예산서 모니터링 (팝업)")
    popup.geometry("1400x800")

    main_frame = ttk.Frame(popup, padding=10)
    main_frame.pack(expand=True, fill='both')

    # 1. KPI Panel (메인 UI 그대로 재현)
    kpi_frame = tk.Frame(main_frame, bg='#f8fafc', height=80)
    kpi_frame.pack(fill='x', pady=(0, 10))
    kpi_frame.pack_propagate(False)

    def create_kpi_card(parent, title, color):
        frame = tk.Frame(parent, bg='white', highlightbackground='#e2e8f0', highlightthickness=1)
        frame.pack(side='left', expand=True, fill='both', padx=5, pady=5)
        tk.Label(frame, text=title, font=('Malgun Gothic', 9), fg='#64748b', bg='white').pack(pady=(5, 2))
        lbl = tk.Label(frame, text="0", font=('Malgun Gothic', 16, 'bold'), fg=color, bg='white')
        lbl.pack(pady=(0, 5))
        return lbl

    lbl_rev = create_kpi_card(kpi_frame, "총 매출액", "#2563eb")
    lbl_cost = create_kpi_card(kpi_frame, "총 실행원가", "#dc2626")
    lbl_profit = create_kpi_card(kpi_frame, "총 기대이익", "#059669")
    lbl_margin = create_kpi_card(kpi_frame, "평균 이익률", "#7c3aed")

    # 2. 예산 목록 트리뷰
    tree_frame = ttk.Frame(main_frame)
    tree_frame.pack(expand=True, fill='both')

    vsb = ttk.Scrollbar(tree_frame, orient='vertical')
    hsb = ttk.Scrollbar(tree_frame, orient='horizontal')

    columns = ('Site', 'Revenue', 'UnitPrice', 'LaborCost', 'MaterialCost', 'Expense', 'OutsourceCost', 'Profit', 'TargetMargin', 'ActualMargin', 'Note')
    tree = ttk.Treeview(tree_frame, columns=columns, show='headings', yscrollcommand=vsb.set, xscrollcommand=hsb.set)

    col_map = {
        'Site': ('현장명', 150), 'Revenue': ('매출액(목표)', 120), 'UnitPrice': ('단가', 100),
        'LaborCost': ('노무비', 120), 'MaterialCost': ('자재비', 120), 'Expense': ('경비', 120),
        'OutsourceCost': ('외주비', 120), 'Profit': ('기대이익(목표)', 120), 'TargetMargin': ('목표이익률', 100), 'ActualMargin': ('현재이익률', 100), 'Note': ('비고', 200)
    }
    for col in columns:
        head, width = col_map[col]
        tree.heading(col, text=head, command=lambda c=col: self.treeview_sort_column(tree, c, False))
        tree.column(col, width=width, anchor='center', stretch=False)

    tree.grid(row=0, column=0, sticky='nsew')
    vsb.grid(row=0, column=1, sticky='ns')
    hsb.grid(row=1, column=0, sticky='ew')
    tree_frame.grid_rowconfigure(0, weight=1); tree_frame.grid_columnconfigure(0, weight=1)

    vsb.config(command=tree.yview); hsb.config(command=tree.xview)

    # [NEW] 인터랙티브 기능 추가
    self.enable_tree_column_drag(tree)

    self.detached_windows['budget'] = {
        'window': popup,
        'tree': tree,
        'lbl_rev': lbl_rev,
        'lbl_cost': lbl_cost,
        'lbl_profit': lbl_profit,
        'lbl_margin': lbl_margin
    }

    def on_close():
        if 'budget' in self.detached_windows:
            del self.detached_windows['budget']
        popup.destroy()
    popup.protocol("WM_DELETE_WINDOW", on_close)

    self._refresh_detached_budget()


def _refresh_detached_budget_impl(self):
    """팝업창의 예산 리스트와 KPI를 최신 데이터로 갱신"""
    if 'budget' not in self.detached_windows: return
    detached = self.detached_windows['budget']
    tree = detached['tree']

    for item in tree.get_children(): tree.delete(item)

    if self.budget_df.empty:
        for lbl in [detached['lbl_rev'], detached['lbl_cost'], detached['lbl_profit']]: lbl.config(text="0")
        detached['lbl_margin'].config(text="0.0%")
        return

    total_rev = total_cost = total_profit = 0

    for _, row in self.budget_df.iterrows():
        rev = pd.to_numeric(row.get('Revenue', 0), errors='coerce') or 0
        unit_p = pd.to_numeric(row.get('UnitPrice', 0), errors='coerce') or 0
        lab = pd.to_numeric(row.get('LaborCost', 0), errors='coerce') or 0
        mat = pd.to_numeric(row.get('MaterialCost', 0), errors='coerce') or 0
        exp = pd.to_numeric(row.get('Expense', 0), errors='coerce') or 0
        out = pd.to_numeric(row.get('OutsourceCost', 0), errors='coerce') or 0
        prof = pd.to_numeric(row.get('Profit', 0), errors='coerce') or 0

        a_rev = pd.to_numeric(row.get('Actual_Revenue', 0), errors='coerce') or 0
        a_prof = pd.to_numeric(row.get('Actual_Profit', 0), errors='coerce') or 0

        total_rev += rev
        total_cost += (lab + mat + exp + out)
        total_profit += prof

        target_margin = (prof / rev * 100) if rev > 0 else 0
        actual_margin = (a_prof / a_rev * 100) if a_rev > 0 else 0

        tree.insert('', 'end', values=(
            row.get('Site', ''),
            f"{rev:,.0f}",
            f"{unit_p:,.0f}",
            f"{lab:,.0f}",
            f"{mat:,.0f}",
            f"{exp:,.0f}",
            f"{out:,.0f}",
            f"{prof:,.0f}",
            f"{target_margin:.1f}%",
            f"{actual_margin:.1f}%",
            row.get('Note', '')
        ))

    # KPI Update
    detached['lbl_rev'].config(text=f"{total_rev:,.0f}")
    detached['lbl_cost'].config(text=f"{total_cost:,.0f}")
    detached['lbl_profit'].config(text=f"{total_profit:,.0f}")
    total_margin = (total_profit / total_rev * 100) if total_rev > 0 else 0
    detached['lbl_margin'].config(text=f"{total_margin:.1f}%")

