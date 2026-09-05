from site_apps.lotte.src.utils.helpers import NAN_PATTERN, DOT_ZERO_PATTERN, MARKER_PATTERN
from site_apps.lotte.src.views.components import *
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import math
import datetime
import json
from tkcalendar import DateEntry
import traceback

def setup_monthly_usage_tab_impl(self):
    """Setup the monthly usage aggregation tab (auto-aggregated from daily usage)"""
    # Display frame for aggregated monthly data
    display_frame = ttk.LabelFrame(self.tab_monthly_usage, text="월별 사용량 집계 (현장별 데이터 자동 집계)")
    display_frame.pack(expand=True, fill='both', padx=10, pady=10)
    
    # Filter controls
    filter_frame = ttk.Frame(display_frame)
    filter_frame.pack(fill='x', padx=5, pady=5)
    
    ttk.Label(filter_frame, text="연도:").pack(side='left', padx=5)
    current_year = datetime.datetime.now().year
    year_values = ['전체'] + [str(y) for y in range(max(2024, current_year-1), current_year + 3)]
    self.cb_filter_year = ttk.Combobox(filter_frame, values=year_values, width=10)
    self.cb_filter_year.pack(side='left', padx=5)
    self.cb_filter_year.set('전체')
    
    ttk.Label(filter_frame, text="월:").pack(side='left', padx=5)
    current_month = datetime.datetime.now().month
    self.cb_filter_month = ttk.Combobox(filter_frame, values=['전체'] + [str(m) for m in range(1, 13)], width=10)
    self.cb_filter_month.pack(side='left', padx=5)
    self.cb_filter_month.set('전체')
    
    ttk.Label(filter_frame, text="현장:").pack(side='left', padx=5)
    self.cb_filter_site_monthly = ttk.Combobox(filter_frame, width=15)
    self.cb_filter_site_monthly.pack(side='left', padx=5)
    self.cb_filter_site_monthly.set('전체')
    
    btn_site_mgr_monthly = tk.Button(filter_frame, text="⚙", font=('Arial', 7), bd=0, bg=self.theme_bg, fg='gray',
                                   command=lambda: self.open_list_management_dialog('sites', target_cb=self.cb_filter_site_monthly))
    btn_site_mgr_monthly.place(in_=self.cb_filter_site_monthly, relx=1.0, x=-18, rely=0.5, anchor='e', width=16, height=16)
    
    ttk.Label(filter_frame, text="품목명:").pack(side='left', padx=5)
    self.cb_filter_material_monthly = ttk.Combobox(filter_frame, width=25)
    self.cb_filter_material_monthly.pack(side='left', padx=5)
    self.cb_filter_material_monthly.set('전체')
    
    btn_filter = ttk.Button(filter_frame, text="조회", command=self.update_monthly_usage_view)
    btn_filter.pack(side='left', padx=10)
    
    btn_export = ttk.Button(filter_frame, text="엑셀 내보내기", command=self.export_monthly_usage_history)
    btn_export.pack(side='left', padx=5)

    btn_popout = ttk.Button(filter_frame, text="🔍 팝업창으로 열기", command=self.open_detached_monthly_usage_view)
    btn_popout.pack(side='left', padx=5)
    
    # Use a PanedWindow to allow resizing between main tree and summaries
    self.monthly_paned = ttk.PanedWindow(display_frame, orient="vertical")
    self.monthly_paned.pack(expand=True, fill='both', padx=5, pady=5)
    
    # 1. Top pane: Main Monthly Usage Tree
    tree_frame = ttk.Frame(self.monthly_paned)
    self.monthly_paned.add(tree_frame, weight=3) # Give main tree more weight
    
    # Scrollbars
    vsb = ttk.Scrollbar(tree_frame, orient="vertical")
    hsb = ttk.Scrollbar(tree_frame, orient="horizontal")
    
    # Treeview with columns including worker, work time, and OT fields
    # Added '(Full작업자)' for full list backup during Excel export
    columns = ('연도', '월', '현장', '구분', '작업자', '작업시간', 'OT시간', 'OT금액', 'OT1', 'OT2', 'OT3', 'OT4', 'OT5', 'OT6', 'OT7', 'OT8', 'OT9', 'OT10', 
               '수량', '단가', '출장비', '일식', '검사비', '제경비', '기술료', '환산물량', '재료비', '인건비', '품목명', '센터미스', '농도', '마킹미스', '필름마크', 
               '취급부주의', '고객불만', '기타', 'RTK총계', '형광자분', '흑색자분', '백색페인트', '침투제', '세척제', '현상제', '형광침투제', '비고', '(Full작업자)')
    self.monthly_usage_tree = ttk.Treeview(tree_frame, columns=columns, show='headings',
                                           yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    
    # Hide the (Full작업자) column from display
    self.monthly_usage_tree['displaycolumns'] = [c for c in columns if c != '(Full작업자)']
    
    vsb.config(command=self.monthly_usage_tree.yview)
    hsb.config(command=self.monthly_usage_tree.xview)
    
    # Column configuration with added OT columns
    col_widths = {
        '연도': 90, '월': 70, '현장': 140, '구분': 100, '작업자': 100, '작업시간': 100,
        'OT시간': 100, 'OT금액': 110,
        'OT1': 100, 'OT2': 100, 'OT3': 100, 'OT4': 100, 'OT5': 100,
        'OT6': 100, 'OT7': 100, 'OT8': 100, 'OT9': 100, 'OT10': 100,
        '수량': 110, '단가': 110, '출장비': 110, '일식': 110, '검사비': 110,
        '제경비': 100, '기술료': 100, '환산물량': 100, '재료비': 100, '인건비': 100,
        '품목명': 220, '센터미스': 80, '농도': 80, '마킹미스': 80,
        '필름마크': 80, '취급부주의': 80, '고객불만': 80, '기타': 80, 'RTK총계': 80,
        '형광자분': 90, '흑색자분': 90, '백색페인트': 90, '침투제': 90, '세척제': 90,
        '현상제': 90, '형광침투제': 90, '비고': 220
    }
    
    for col in columns:
        self.monthly_usage_tree.heading(col, text=col, command=lambda c=col: self.treeview_sort_column(self.monthly_usage_tree, c, False))
        width = col_widths.get(col, 100)
        self.monthly_usage_tree.column(col, width=width, minwidth=20, stretch=False, anchor='center')
    
    # [NEW] Enable column reordering via drag & drop
    self.enable_tree_column_drag(self.monthly_usage_tree)
    
    
    # Grid layout
    self.monthly_usage_tree.grid(row=0, column=0, sticky='nsew')
    vsb.grid(row=0, column=1, sticky='ns')
    hsb.grid(row=1, column=0, sticky='ew')
    
    tree_frame.grid_rowconfigure(0, weight=1)
    tree_frame.grid_columnconfigure(0, weight=1)
    
    # [NEW] Auto-save column widths when user resizes columns
    self.monthly_usage_tree.bind('<ButtonRelease-1>', lambda e: self.save_tab_config())
    
    # [NEW] Bind selection to update site/worker summaries
    self.monthly_usage_tree.bind('<<TreeviewSelect>>', self.on_monthly_usage_select)

    # 2. Middle pane: Site Summary
    site_frame = ttk.LabelFrame(self.monthly_paned, text="현장별 누계")
    self.monthly_paned.add(site_frame, weight=1)
    
    self.monthly_usage_tree.bind("<Button-1>", lambda e: self.show_worker_popup(e, self.monthly_usage_tree), add="+")
    
    site_cols = ('현장', '검사방법', '품목명', '수량', '검사비', '출장비', '제경비', '기술료', '환산물량', '재료비', '인건비', '형광자분', '흑색자분', '백색페인트', 
                 '침투제', '세척제', '현상제', '형광침투제', '센터미스', '농도', '마킹미스', '필름마크', '취급부주의', '고객불만', '기타', 'RTK총계')
    self.site_summary_tree = ttk.Treeview(site_frame, columns=site_cols, show='headings')
    for col in site_cols:
        self.site_summary_tree.heading(col, text=col)
        # Adjust widths based on content
        if col in ['현장', '검사방법', '품목명']: width = 120
        elif col in ['검사비', '출장비', '제경비', '기술료', '환산물량', '재료비', '인건비']: width = 100
        else: width = 80
        self.site_summary_tree.column(col, width=width, anchor='center', stretch=False)
    
    # [NEW] Auto-save site summary column widths
    self.site_summary_tree.bind('<ButtonRelease-1>', lambda e: self.save_tab_config())
    self.enable_tree_column_drag(self.site_summary_tree, context_menu_handler=lambda e: self._show_generic_tree_heading_context_menu(e, self.site_summary_tree))
    
    site_vsb = ttk.Scrollbar(site_frame, orient="vertical", command=self.site_summary_tree.yview)
    self.site_summary_tree.configure(yscrollcommand=site_vsb.set)
    
    self.site_summary_tree.pack(side='left', expand=True, fill='both')
    site_vsb.pack(side='right', fill='y')
    
    # 3. Bottom pane: Worker Summary
    worker_frame = ttk.LabelFrame(self.monthly_paned, text="작업자별 누계")
    self.monthly_paned.add(worker_frame, weight=1)
    
    worker_cols = ('작업자', '총공수', '연장(시간)', '야간(시간)', '휴일(시간)', '총OT(시간)', '연장(금액)', '야간(금액)', '휴일(금액)', '총OT(금액)')
    self.worker_summary_tree = ttk.Treeview(worker_frame, columns=worker_cols, show='headings')
    
    # Set column widths for worker summary
    worker_widths = {
        '작업자': 100, '총공수': 70, '연장(시간)': 70, '야간(시간)': 70, '휴일(시간)': 70, 
        '총OT(시간)': 80, '연장(금액)': 90, '야간(금액)': 90, '휴일(금액)': 90, '총OT(금액)': 100
    }
    for col in worker_cols:
        self.worker_summary_tree.heading(col, text=col)
        width = worker_widths.get(col, 80)
        # Disable stretching
        self.worker_summary_tree.column(col, width=width, anchor='center', stretch=False)
    
    # [NEW] Auto-save worker summary column widths
    self.worker_summary_tree.bind('<ButtonRelease-1>', lambda e: self.save_tab_config())
    self.enable_tree_column_drag(self.worker_summary_tree, context_menu_handler=lambda e: self._show_generic_tree_heading_context_menu(e, self.worker_summary_tree))
        
    worker_vsb = ttk.Scrollbar(worker_frame, orient="vertical", command=self.worker_summary_tree.yview)
    self.worker_summary_tree.configure(yscrollcommand=worker_vsb.set)
    
    self.worker_summary_tree.pack(side='left', expand=True, fill='both')
    worker_vsb.pack(side='right', fill='y')

    # [NEW] 공사탭 특별근무 자동 입력 버튼
    btn_apply_frame = ttk.Frame(worker_frame)
    btn_apply_frame.pack(fill='x', pady=3)
    ttk.Button(btn_apply_frame,
               text="📋 사후원가 특별근무에 적용",
               command=self.apply_worker_shift_hours_to_budget).pack(side='right', padx=5)

    # Initial view update
    self.update_monthly_usage_view()


def update_monthly_usage_view_impl(self):
    """Update the monthly usage treeview with aggregated data from daily usage"""
    # Clear current views
    for item in self.monthly_usage_tree.get_children():
        self.monthly_usage_tree.delete(item)
    for item in self.site_summary_tree.get_children():
        self.site_summary_tree.delete(item)
    for item in self.worker_summary_tree.get_children():
        self.worker_summary_tree.delete(item)
        
    # [NEW] Clear detached monthly views if open
    if 'monthly' in self.detached_windows:
        p_tree = self.detached_windows['monthly']['tree']
        for item in p_tree.get_children():
            p_tree.delete(item)
    
    # Get filter values
    filter_year = self.cb_filter_year.get()
    filter_month = self.cb_filter_month.get()
    filter_site = self.cb_filter_site_monthly.get() if hasattr(self, 'cb_filter_site_monthly') else '전체'
    filter_material = self.cb_filter_material_monthly.get() if hasattr(self, 'cb_filter_material_monthly') else '전체'
    
    # Return if daily usage data is empty
    if self.daily_usage_df.empty:
        return
    
    # Create a copy of daily usage data and extract year/month from Date column
    df = self.daily_usage_df.copy()
    # [CRITICAL] Normalize columns to ensure detection (matched_pairs) matches data lookups
    df.columns = [str(c).strip().replace(' ', '') for c in df.columns]
    
    # Normalize column names - remove ALL types of whitespace using regex
    import re
    df.columns = [re.sub(r'\s+', '', str(c)) for c in df.columns]
    
    df['Year'] = pd.to_datetime(df['Date'], errors='coerce').dt.year
    df['Month'] = pd.to_datetime(df['Date'], errors='coerce').dt.month
    # 날짜 파싱 실패 행 제거
    df = df.dropna(subset=['Year', 'Month'])
    df['Year'] = df['Year'].astype(int)
    df['Month'] = df['Month'].astype(int)
    
    # [ROBUST] Normalize Site names to prevent duplicate grouping/mismatches (e.g., merging "Site-A" and "Site - A")
    if 'Site' in df.columns:
        df['Site'] = df['Site'].astype(str).apply(self.normalize_site_name)
    
    # Apply filters
    if filter_year != '전체':
        df = df[df['Year'] == int(filter_year)]
    
    if filter_month != '전체':
        df = df[df['Month'] == int(filter_month)]
    
    if filter_site != '전체':
        df = df[df['Site'] == filter_site]
        
    if filter_material != '전체':
        # Get matching MaterialIDs for the selected item name
        matching_ids = self.materials_df[self.materials_df['품목명'] == filter_material]['MaterialID'].tolist()
        if matching_ids:
            df = df[df['MaterialID'].isin(matching_ids)]
        else:
            # If no material found in master list, clear and return (though this shouldn't happen with sync)
            return
    
    # [NEW] Return early if filtered df is empty to avoid ValueError during assignment
    if df.empty:
        return
    # Populate site filter options from data
    if hasattr(self, 'cb_filter_site_monthly'):
        # [ROBUST] Use same collection logic as refresh_inquiry_filters to exclude hidden sites and maintain sync
        raw_sites = set()
        if not self.daily_usage_df.empty and 'Site' in self.daily_usage_df.columns:
            raw_sites.update(self.daily_usage_df['Site'].dropna().astype(str).apply(self.normalize_site_name).tolist())
        if hasattr(self, 'sites'):
            for s in self.sites:
                norm = self.normalize_site_name(s)
                if norm: raw_sites.add(norm)
        if hasattr(self, 'budget_df') and not self.budget_df.empty and 'Site' in self.budget_df.columns:
            raw_sites.update(self.budget_df['Site'].dropna().astype(str).apply(self.normalize_site_name).tolist())
        
        unique_sites = ['전체'] + sorted([s for s in raw_sites 
                                         if s and str(s).lower() != 'nan'
                                        and s not in getattr(self, 'hidden_sites', [])])
        
        self.cb_filter_site_monthly['values'] = unique_sites
        if not self.cb_filter_site_monthly.get():
            self.cb_filter_site_monthly.set('전체')
    
    # Populate material filter options from data
    if hasattr(self, 'cb_filter_material_monthly'):
        # Get unique material names from materials_df based on MaterialIDs in daily_usage_df
        # Note: MaterialID column name itself might have spaces in self.daily_usage_df
        m_id_col = 'MaterialID' if 'MaterialID' in self.daily_usage_df.columns else 'MaterialID'
        unique_mat_ids = self.daily_usage_df[m_id_col].dropna().unique()
        material_names = []
        for mat_id in unique_mat_ids:
            # [ROBUST] Use same lookup logic for filters
            try:
                m_id_f = float(mat_id)
                matches = self.materials_df[pd.to_numeric(self.materials_df['MaterialID'], errors='coerce') == m_id_f]
                if not matches.empty:
                    material_names.append(str(matches.iloc[0]['품목명']))
                else:
                    material_names.append(f"ID: {mat_id}")
            except:
                matches = self.materials_df[self.materials_df['MaterialID'].astype(str) == str(mat_id)]
                if not matches.empty:
                    material_names.append(str(matches.iloc[0]['품목명']))
                else:
                    material_names.append(f"ID: {mat_id}")
        unique_materials = ['전체'] + sorted(set(material_names))
        self.cb_filter_material_monthly['values'] = unique_materials
        if not self.cb_filter_material_monthly.get():
            self.cb_filter_material_monthly.set('전체')

    # [NEW] Sync detached filters to main window state
    if 'monthly' in self.detached_windows:
        p_filters = self.detached_windows['monthly'].get('filters', {})
        if p_filters:
            p_filters['year'].set(self.cb_filter_year.get())
            p_filters['month'].set(self.cb_filter_month.get())
            p_filters['site'].set(self.cb_filter_site_monthly.get())
            p_filters['mat'].set(self.cb_filter_material_monthly.get())
            
            # Also sync dropdown values for Site/Material
            p_filters['site']['values'] = self.cb_filter_site_monthly['values']
            p_filters['mat']['values'] = self.cb_filter_material_monthly['values']
    
    # Prepare aggregation dictionary for all numeric fields
    agg_dict = {'Usage': 'sum'}
    
    # Helper for joining workers - clean internal spaces to avoid "11시간" vs "11 시간" mismatch
    def join_unique_non_empty(series):
        # Strip outer spaces and compress internal spaces for consistency
        vals = [" ".join(str(v).split()) for v in series if pd.notna(v) and str(v).strip()]
        return " | ".join(sorted(set(vals)))
    
    # Also define a sum helper that handles potential type issues and commas
    def safe_sum(series):
        def to_f(v):
            if pd.isna(v) or str(v).lower() in ('nan', 'none', ''): return 0.0
            s = str(v).strip().lower()
            try: 
                # Remove non-numeric markers (comma, unit, etc)
                clean_s = re.sub(r'[^0-9\.\-]', '', s)
                return float(clean_s) if clean_s else 0.0
            except: return 0.0
        return series.apply(to_f).sum()

    # [NEW] Hyper-Robust Column Detection for Workers/WorkTime/OT
    # Find all columns using exact name matching (not regex with optional suffix)
    def find_paired_cols(cols):
        pairs = []
        col_set = set(cols)
        for i in range(1, 11):
            # i=1: 'User', i=2: 'User2', etc.
            u_name = 'User' if i == 1 else f'User{i}'
            w_name = 'WorkTime' if i == 1 else f'WorkTime{i}'
            o_name = 'OT' if i == 1 else f'OT{i}'
            
            u_col = u_name if u_name in col_set else None
            w_col = w_name if w_name in col_set else None
            o_col = o_name if o_name in col_set else None
            
            if u_col: pairs.append((u_col, w_col, o_col))
        return pairs

    matched_pairs = find_paired_cols(df.columns)
    for u_c, w_c, o_c in matched_pairs:
        agg_dict[u_c] = join_unique_non_empty
        if w_c: agg_dict[w_c] = join_unique_non_empty
        if o_c: agg_dict[o_c] = join_unique_non_empty
    
    # Removed FilmCount aggregation as it is now integrated into Usage
    
    # Add RTK categories
    rtk_categories = ['RTK_센터미스', 'RTK_농도', 'RTK_마킹미스', 'RTK_필름마크', 'RTK_취급부주의', 'RTK_고객불만', 'RTK_기타']
    for cat in rtk_categories:
        if cat in df.columns:
            agg_dict[cat] = safe_sum
    
    # Add NDT materials and cost fields
    other_agg_cols = ['NDT_형광자분', 'NDT_자분', 'NDT_흑색자분', 'NDT_페인트', 'NDT_백색페인트', 'NDT_침투제', 'NDT_세척제', 'NDT_현상제', 'NDT_형광', 'NDT_형광침투제',
                      '검사량', '단가', '출장비', '일식', '검사비', '제경비', '기술료', '환산물량', '재료비', '인건비']
    for col in other_agg_cols:
        if col in df.columns:
            agg_dict[col] = safe_sum

    # [NEW] Pre-aggregation Deduping to ensure parity with Site Tab
    seen_m_times = set()
    seen_m_contents = set()
    
    # Robust mapping for worker columns in Monthly Tab
    matched_pairs = find_paired_cols(df.columns)

    def sync_dedup_and_calc_ot(row):
        # 1. Calculate Activity-based OT (Max of all workers in this row)
        h_max = 0.0
        a_sum = 0
        w_count = 0
        raw_workers = []
        
        for u_c, w_c, o_c in matched_pairs:
            u_v = self.clean_nan(row.get(u_c, ''))
            if u_v:
                raw_workers.append(u_v)
                w_count += 1
            
            if o_c:
                ots = str(row.get(o_c, '')).strip()
                if ots and ots not in ('nan', '0.0', '0'):
                    try:
                        # Use same parsing logic as Site Tab
                        if '(' in ots and '원)' in ots:
                            h_p = float(ots.split('시간')[0])
                            amt_str = _re.sub(r'[^0-9]', '', ots.split('(')[1].split('원')[0])
                            a_p = int(amt_str) if amt_str else 0
                        elif ots.replace(',', '').isdigit():
                            a_p = int(ots.replace(',', ''))
                            wt_v = str(row.get(w_c, '')).strip() if w_c else ''
                            h_p, _ = self._calculate_ot_from_worktime(wt_v, pd.to_datetime(row.get('Date', pd.Timestamp.now())))
                        else:
                            a_p = self.calculate_ot_amount(ots)
                            h_p = self._parse_ot_hours(ots)
                        
                        h_max = max(h_max, h_p)
                        a_sum += a_p
                    except: pass

        # 2. Deduping Key (Date, Site, WorkTime, Material)
        # [REFINED] Exclude workers from the key to correctly catch records split across rows.
        n_date = self._safe_format_datetime(row.get('Date', ''), '%Y-%m-%d')
        n_site = str(row.get('Site', '')).strip()
        c_worktime = str(row.get('WorkTime', '')).strip()
        n_mat = str(row.get('MaterialID', ''))
        
        content_key = (n_date, n_site, c_worktime, n_mat)
        
        e_t_raw = row.get('EntryTime', '')
        try:
            if isinstance(e_t_raw, (pd.Timestamp, datetime.datetime)):
                t_key = e_t_raw.strftime('%Y-%m-%d %H:%M:%S')
            else:
                t_key = str(e_t_raw).split('.')[0].strip() if e_t_raw else ""
        except: t_key = ""
        
        is_dup = (t_key and t_key in seen_m_times) or (content_key in seen_m_contents)
        
        if not is_dup:
            if t_key: seen_m_times.add(t_key)
            seen_m_contents.add(content_key)
            
            # Primary row: return full values
            return pd.Series([h_max, a_sum, w_count, False])
        else:
            # Duplicate row: Zero out quantitative impact for aggregation
            return pd.Series([0.0, 0, 0, True])

    # Apply calculation and marking
    calc_results = df.apply(sync_dedup_and_calc_ot, axis=1)
    df[['OT시간', 'OT금액', 'WorkerCount', '_is_m_dup']] = calc_results
    
    # Zero out other quantitative fields for duplicate rows before aggregation
    # [FIX] Do NOT zero out 'Usage' and '검사량' during view-level deduping.
    # Database-level splitting already zeros them for valid splits. 
    # View-level zeroing was causing data loss if the quantity was on a secondary row.
    q_fields = ['단가', '출장비', '일식', '검사비', 'OT시간', 'OT금액', '제경비', '기술료', '환산물량', '재료비', '인건비']
    rtk_fields = [f'RTK_{c}' for c in ['센터미스', '농도', '마킹미스', '필름마크', '취급부주의', '고객불만', '기타']]
    ndt_fields = ['NDT_형광자분', 'NDT_자분', 'NDT_흑색자분', 'NDT_페인트', 'NDT_백색페인트', 'NDT_침투제', 'NDT_세척제', 'NDT_현상제', 'NDT_형광', 'NDT_형광침투제']
    
    dup_mask = df['_is_m_dup'] == True
    for f in q_fields + rtk_fields + ndt_fields:
        if f in df.columns:
            df.loc[dup_mask, f] = 0.0

    # Group by Year, Month, Site, MaterialID and aggregate
    agg_dict['OT시간'] = 'sum'
    agg_dict['OT금액'] = 'sum'
    agg_dict['WorkerCount'] = 'sum'
    grouped = df.groupby(['Year', 'Month', 'Site', 'MaterialID'], dropna=False).agg(agg_dict).reset_index()
    
    self.monthly_usage_tree.bind("<Button-1>", lambda e: self.show_worker_popup(e, self.monthly_usage_tree), add="+")
    
    # Store df for selection handling
    self.current_monthly_df = df
    
    # Initialize totals for cumulative sum
    total_worker_count = 0.0
    total_ot_hours = 0.0
    total_ot_amount = 0.0
    total_test_amount = 0.0
    total_unit_price = 0.0
    total_travel_cost = 0.0
    total_meal_cost = 0.0
    total_test_fee = 0.0
    total_overhead_cost = 0.0
    total_tech_fee = 0.0
    total_conv_qty = 0.0
    total_mat_cost = 0.0
    total_labor_cost = 0.0
    total_rtk_center = 0.0
    total_rtk_density = 0.0
    total_rtk_marking = 0.0
    total_rtk_film = 0.0
    total_rtk_handling = 0.0
    total_rtk_customer = 0.0
    total_rtk_other = 0.0
    total_ndt_fluorescent_mag = 0.0
    total_ndt_magnet = 0.0
    total_ndt_paint = 0.0
    total_ndt_penetrant = 0.0
    total_ndt_cleaner = 0.0
    total_ndt_developer = 0.0
    total_ndt_fluorescent_pen = 0.0
    total_indiv_ot_amounts = [0] * 10
    has_note = False
    
    # Display aggregated entries
    for _, entry in grouped.iterrows():
        def clean_str(val):
            return str(val).replace('nan', '').replace('None', '').strip()
            
        mat_id = entry['MaterialID']
        
        # [ROBUST] Handle type mismatch in MaterialID lookup for Monthly Tab
        def get_mat_name(m_id):
            if pd.isna(m_id) or str(m_id).lower() == 'nan': return "N/A"
            # Try to find name from materials_df
            try:
                # Convert both to float for numeric comparison
                m_id_f = float(m_id)
                matches = self.materials_df[pd.to_numeric(self.materials_df['MaterialID'], errors='coerce') == m_id_f]
                if not matches.empty:
                    return str(matches.iloc[0]['품목명'])
            except: pass
            
            # Fallback to string comparison
            matches = self.materials_df[self.materials_df['MaterialID'].astype(str) == str(m_id)]
            if not matches.empty:
                return str(matches.iloc[0]['품목명'])
            
            return f"ID: {m_id}"

        mat_name = get_mat_name(mat_id)
        
        # Apply material filter
        if filter_material != '전체' and mat_name != filter_material:
            continue
        
        # Get aggregated values from clean columns
        work_count = entry.get('WorkerCount', 0.0)
        ot_hours = entry.get('OT시간', 0.0)
        ot_amount = entry.get('OT금액', 0.0)
        test_amount = entry.get('검사량', 0.0)
        unit_price = entry.get('단가', 0.0)
        travel_cost = entry.get('출장비', 0.0)
        meal_cost = entry.get('일식', 0.0)
        test_fee = entry.get('검사비', 0.0)
        overhead_cost = entry.get('제경비', 0.0)
        tech_fee = entry.get('기술료', 0.0)
        conv_qty = entry.get('환산물량', 0.0)
        mat_cost = entry.get('재료비', 0.0)
        labor_cost = entry.get('인건비', 0.0)
        
        # RTK values
        rtk_center = entry.get('RTK_센터미스', 0.0)
        rtk_density = entry.get('RTK_농도', 0.0)
        rtk_marking = entry.get('RTK_마킹미스', 0.0)
        rtk_film = entry.get('RTK_필름마크', 0.0)
        rtk_handling = entry.get('RTK_취급부주의', 0.0)
        rtk_customer = entry.get('RTK_고객불만', 0.0)
        rtk_other = entry.get('RTK_기타', 0.0)
        rtk_total = rtk_center + rtk_density + rtk_marking + rtk_film + rtk_handling + rtk_customer + rtk_other
        
        # NDT values
        ndt_fluorescent_mag = entry.get('NDT_형광자분', 0.0)
        ndt_magnet = entry.get('NDT_자분', 0.0) + entry.get('NDT_흑색자분', 0.0)
        ndt_paint = entry.get('NDT_페인트', 0.0) + entry.get('NDT_백색페인트', 0.0)
        ndt_penetrant = entry.get('NDT_침투제', 0.0)
        ndt_cleaner = entry.get('NDT_세척제', 0.0)
        ndt_developer = entry.get('NDT_현상제', 0.0)
        ndt_fluorescent_pen = entry.get('NDT_형광', 0.0) + entry.get('NDT_형광침투제', 0.0)
        
        # Accumulate totals
        total_worker_count += work_count
        total_ot_hours += ot_hours
        total_ot_amount += ot_amount
        total_test_amount += test_amount
        total_unit_price += unit_price
        total_travel_cost += travel_cost
        total_meal_cost += meal_cost
        total_test_fee += test_fee
        total_overhead_cost += overhead_cost
        total_tech_fee += tech_fee
        total_conv_qty += conv_qty
        total_mat_cost += mat_cost
        total_labor_cost += labor_cost
        
        total_rtk_center += rtk_center
        total_rtk_density += rtk_density
        total_rtk_marking += rtk_marking
        total_rtk_film += rtk_film
        total_rtk_handling += rtk_handling
        total_rtk_customer += rtk_customer
        total_rtk_other += rtk_other
        total_ndt_fluorescent_mag += ndt_fluorescent_mag
        total_ndt_magnet += ndt_magnet
        total_ndt_paint += ndt_paint
        total_ndt_penetrant += ndt_penetrant
        total_ndt_cleaner += ndt_cleaner
        total_ndt_developer += ndt_developer
        total_ndt_fluorescent_pen += ndt_fluorescent_pen
        
        # Extract worker names from User, User2, ..., User10
        all_workers = []
        # [ROBUST] Extract workers from matched columns
        for u_c, _, _ in matched_pairs:
            val = self.clean_nan(entry.get(u_c, ''))
            if val:
                # Split in case it was already joined in aggregation
                all_workers.extend([v.strip() for v in val.split(' | ') if v.strip()])
        worker_str = self.format_worker_summary(all_workers)
        
        # Concatenate work times
        all_worktimes = []
        worktime_cols = ['WorkTime', 'WorkTime2', 'WorkTime3', 'WorkTime4', 'WorkTime5', 'WorkTime6', 'WorkTime7', 'WorkTime8', 'WorkTime9', 'WorkTime10']
        for col in worktime_cols:
            val = str(entry.get(col, '')).strip()
            if val and val != 'nan' and val != '0.0':
                all_worktimes.extend([v.strip() for v in val.split(' | ') if v.strip()])
        worktime_str = ", ".join(sorted(set(all_worktimes)))
        
        # Extract OT amounts only (suppress time strings)
        ot_values = []
        # [ROBUST] Pair-based OT extraction (and ensure exactly 10 items for column alignment)
        for i in range(1, 11):
            if i <= len(matched_pairs):
                u_c, _, o_c = matched_pairs[i-1]
                if not o_c:
                    ot_values.append('')
                    continue
                val = str(entry.get(o_c, '')).strip()
                
                # [FIX] Eliminate ghost OT values in monthly view: Skip if no worker is assigned to this slot
                user_val = str(entry.get(u_c, '')).strip()
                has_worker = user_val and user_val != 'nan'
                
                if has_worker and val and val != 'nan' and val != '0.0':
                    # Handle multiple OTs if joined by aggregation separator ' | '
                    sub_vals = [v.strip() for v in val.split(' | ') if v.strip()]
                    parsed_ots = []
                    for v_str in sub_vals:
                        v_clean = v_str.replace(',', '')
                        if v_clean.isdigit() and int(v_clean) > 100:
                            # [PARITY] Monetary amount: Calculate hours from paired WorkTime
                            wt_val = str(entry.get(w_c, '')).strip() if w_c else ''
                            # Note: for individuals, we usually just show the amount in these columns
                            parsed_ots.append(f"{int(v_clean):,}")
                        elif '(' in v_str and '원)' in v_str:
                            try:
                                amount_str = v_str.split('(')[1].split('원')[0].replace(',', '').strip()
                                amount = int(float(amount_str))
                                parsed_ots.append(f"{amount:,}")
                            except: pass
                        else:
                            try:
                                amt = self.calculate_ot_amount(v_str)
                                if amt > 0: parsed_ots.append(f"{amt:,}")
                                elif any(x in v_str for x in [':', '시', '시간', '~', '-']): pass 
                                else: parsed_ots.append(v_str)
                            except: pass
                    
                    ot_values.append(", ".join(parsed_ots))
                    try:
                        # [FIX] Sum ALL values accurately
                        for p_val in parsed_ots:
                            try:
                                amt = int(p_val.replace(',', ''))
                                total_indiv_ot_amounts[i-1] += amt
                            except: pass
                    except: pass
                else:
                    ot_values.append('')
            else:
                ot_values.append('')
        
        if clean_str(entry.get('Note', '')): has_note = True
        
        val_tuple = (
            int(entry['Year']),
            int(entry['Month']),
            entry.get('Site', ''),
            entry.get('구분', ''),
            worker_str,
            worktime_str,
            f"{ot_hours:.1f}" if ot_hours > 0 else '',
            f"{ot_amount:,.0f}" if ot_amount > 0 else '',
            *ot_values,  # Index 7 to 16
            f"{test_amount:.1f}" if test_amount > 0 else '',
            f"{unit_price:,.0f}" if unit_price > 0 else '',
            f"{travel_cost:,.0f}" if travel_cost > 0 else '',
            f"{meal_cost:,.0f}" if meal_cost > 0 else '',
            f"{test_fee:,.0f}" if test_fee > 0 else '',
            f"{overhead_cost:,.0f}" if overhead_cost > 0 else '',
            f"{tech_fee:,.0f}" if tech_fee > 0 else '',
            f"{conv_qty:.2f}" if conv_qty > 0 else '',
            f"{mat_cost:,.0f}" if mat_cost > 0 else '',
            f"{labor_cost:,.0f}" if labor_cost > 0 else '',
            mat_name, # 품목명
            f"{rtk_center:.1f}" if rtk_center > 0 else '', # 23
            f"{rtk_density:.1f}" if rtk_density > 0 else '', # 24
            f"{rtk_marking:.1f}" if rtk_marking > 0 else '', # 25
            f"{rtk_film:.1f}" if rtk_film > 0 else '', # 26
            f"{rtk_handling:.1f}" if rtk_handling > 0 else '', # 27
            f"{rtk_customer:.1f}" if rtk_customer > 0 else '', # 28
            f"{rtk_other:.1f}" if rtk_other > 0 else '', # 29
            f"{rtk_total:.1f}" if rtk_total > 0 else '', # 30
            f"{ndt_fluorescent_mag:.1f}" if ndt_fluorescent_mag > 0 else '', # 31
            f"{ndt_magnet:.1f}" if ndt_magnet > 0 else '', # 32
            f"{ndt_paint:.1f}" if ndt_paint > 0 else '', # 33
            f"{ndt_penetrant:.1f}" if ndt_penetrant > 0 else '', # 34
            f"{ndt_cleaner:.1f}" if ndt_cleaner > 0 else '', # 35
            f"{ndt_developer:.1f}" if ndt_developer > 0 else '', # 36: 현상제
            f"{ndt_fluorescent_pen:.1f}" if ndt_fluorescent_pen > 0 else '', # 37
            '',  # Index 38: Note
            ", ".join(sorted(set(all_workers))) # Index 39: Full작업자
        )
        # [ROBUST] Final length check for 45 columns
        while len(val_tuple) < 45: val_tuple += ("",)
        
        self.monthly_usage_tree.insert('', tk.END, values=val_tuple)
        
        # [NEW] Popup sync
        if 'monthly' in self.detached_windows:
            p_tree = self.detached_windows['monthly']['tree']
            p_tree.insert('', tk.END, values=val_tuple) # Keep all 40 cols
    
    # Add total row at the bottom if there's data
    if not grouped.empty:
        total_rtk_sum = total_rtk_center + total_rtk_density + total_rtk_marking + total_rtk_film + total_rtk_handling + total_rtk_customer + total_rtk_other
        
        # Configure tag for total row
        self.monthly_usage_tree.tag_configure('total', background='#E8F4F8', font=('Arial', 12, 'bold'))
        
        total_values = (
            '',
            '',
            '--- 전체 누계 ---',
            '',  # 구분
            '',  # 작업자
            f"{total_worker_count:.1f}" if total_worker_count > 0 else '', # 작업시간 (총 공수)
            f"{total_ot_hours:.1f}" if total_ot_hours > 0 else '',
            f"{total_ot_amount:,.0f}" if total_ot_amount > 0 else '',
            *[f"{a:,.0f}" if a > 0 else '' for a in total_indiv_ot_amounts],
            f"{total_test_amount:,.0f}" if total_test_amount > 0 else '',
            '', # 단가
            f"{total_travel_cost:,.0f}" if total_travel_cost > 0 else '',
            f"{total_meal_cost:,.0f}" if total_meal_cost > 0 else '',
            f"{total_test_fee:,.0f}" if total_test_fee > 0 else '',
            f"{total_overhead_cost:,.0f}" if total_overhead_cost > 0 else '',
            f"{total_tech_fee:,.0f}" if total_tech_fee > 0 else '',
            f"{total_conv_qty:.2f}" if total_conv_qty > 0 else '',
            f"{total_mat_cost:,.0f}" if total_mat_cost > 0 else '',
            f"{total_labor_cost:,.0f}" if total_labor_cost > 0 else '',
            '', # 품목명
            f"{total_rtk_center:.1f}" if total_rtk_center > 0 else '',
            f"{total_rtk_density:.1f}" if total_rtk_density > 0 else '',
            f"{total_rtk_marking:.1f}" if total_rtk_marking > 0 else '',
            f"{total_rtk_film:.1f}" if total_rtk_film > 0 else '',
            f"{total_rtk_handling:.1f}" if total_rtk_handling > 0 else '',
            f"{total_rtk_customer:.1f}" if total_rtk_customer > 0 else '',
            f"{total_rtk_other:.1f}" if total_rtk_other > 0 else '',
            f"{total_rtk_sum:.1f}" if total_rtk_sum > 0 else '',
            f"{total_ndt_fluorescent_mag:.1f}" if total_ndt_fluorescent_mag > 0 else '',
            f"{total_ndt_magnet:.1f}" if total_ndt_magnet > 0 else '',
            f"{total_ndt_paint:.1f}" if total_ndt_paint > 0 else '',
            f"{total_ndt_penetrant:.1f}" if total_ndt_penetrant > 0 else '',
            f"{total_ndt_cleaner:.1f}" if total_ndt_cleaner > 0 else '',
            f"{total_ndt_developer:.1f}" if total_ndt_developer > 0 else '',
            f"{total_ndt_fluorescent_pen:.1f}" if total_ndt_fluorescent_pen > 0 else '',
            '', # 비고
            ''  # (Full작업자) hidden storage
        )
        self.monthly_usage_tree.insert('', tk.END, values=total_values, tags=('total',))
        
        # [NEW] Popup sync for total row
        if 'monthly' in self.detached_windows:
            p_tree = self.detached_windows['monthly']['tree']
            p_tree.insert('', tk.END, values=total_values, tags=('total',))
        
        # --- [NEW] Populate Summaries (Initial/Total) ---
        self._populate_monthly_summary_trees(df, has_note)
        
        # --- Dynamic Column Auto-Hide (same logic as Site tab) ---
        def is_active_m(val):
            if val is None: return False
            s = str(val).strip().lower()
            if s in ('', '0', '0.0', '0.00', 'nan', 'none', '-', '0.0시간', '0시간', '0원'): return False
            try:
                import re as _rem
                clean = _rem.sub(r'[^0-9\.\-]', '', s)
                return bool(clean) and abs(float(clean)) > 0.001
            except:
                return bool(s)

        all_cols = list(self.monthly_usage_tree['columns'])
        monthly_hide = set()
        
        # [REFINED] Minimal always_show list for smarter monthly column hiding
        always_show = {'연도', '월', '현장'}
        
        # Individual OT slots
        for i in range(1, 11):
            col = f'OT{i}' if i > 1 else 'OT시간'
            # Check from total amounts
            amt = total_indiv_ot_amounts[i-1] if i <= len(total_indiv_ot_amounts) else 0
            if not is_active_m(amt):
                monthly_hide.add(f'OT{i}')
                monthly_hide.add(f'작업자{i}' if i > 1 else '')
        
        # Cost columns
        if not is_active_m(total_travel_cost): monthly_hide.add('출장비')
        if not is_active_m(total_meal_cost): monthly_hide.add('일식')
        if not is_active_m(total_test_fee): monthly_hide.add('검사비')
        if not is_active_m(total_test_amount): monthly_hide.add('수량')
        if not is_active_m(total_ot_hours): monthly_hide.add('OT시간')
        if not is_active_m(total_ot_amount): monthly_hide.add('OT금액')
        
        # RTK columns
        rtk_totals = [total_rtk_center, total_rtk_density, total_rtk_marking, total_rtk_film, total_rtk_handling, total_rtk_customer, total_rtk_other]
        rtk_col_names_m = ['센터미스', '농도', '마킹미스', '필름마크', '취급부주의', '고객불만', '기타', 'RTK총계']
        for i, col in enumerate(rtk_col_names_m[:7]):
            if not is_active_m(rtk_totals[i]): monthly_hide.add(col)
        rtk_sum = sum(rtk_totals)
        if not is_active_m(rtk_sum): monthly_hide.add('RTK총계')
        
        # NDT columns
        ndt_totals = [total_ndt_fluorescent_mag, total_ndt_magnet, total_ndt_paint, total_ndt_penetrant, total_ndt_cleaner, total_ndt_developer, total_ndt_fluorescent_pen]
        ndt_col_names_m = ['형광자분', '흑색자분', '백색페인트', '침투제', '세척제', '현상제', '형광침투제']
        for i, col in enumerate(ndt_col_names_m):
            if not is_active_m(ndt_totals[i]): monthly_hide.add(col)
        
        # note column
        if not has_note: monthly_hide.add('비고')
        
        display_cols_m = [c for c in all_cols if c not in ('(Full작업자)',) and c not in monthly_hide]
        self.monthly_usage_tree['displaycolumns'] = display_cols_m
        if 'monthly' in self.detached_windows:
            self.detached_windows['monthly']['tree']['displaycolumns'] = display_cols_m
            
        # Ensure Total Row stays at bottom
        self.monthly_usage_tree.detach(self.monthly_usage_tree.get_children()[-1])
        self.monthly_usage_tree.insert('', tk.END, values=total_values, tags=('total',))
    else:
        # If empty data
        visible_cols = ['연도', '월', '현장', '작업자', '작업시간', '품목명']
        self.monthly_usage_tree['displaycolumns'] = visible_cols
        for col in visible_cols:
            self.monthly_usage_tree.column(col, stretch=False, minwidth=20)


