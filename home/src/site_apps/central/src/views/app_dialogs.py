# -*- coding: utf-8 -*-
"""
중앙지사 자재관리 주요 팝업 다이얼로그(근로자, 장비, 자재검색, 수정 등) Mixin
"""

import json
import logging
import os
import re
import sys
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd

logger = logging.getLogger(__name__)

class AppDialogsMixin:
    """중앙지사 메인 앱 주요 팝업 및 상세 관리 다이얼로그 Mixin"""
    def open_edit_material_dialog(self):
        """Open a dialog to edit the selected material"""
        selection = self.stock_tree.selection()
        if not selection:
            messagebox.showwarning("선택 오류", "수정할 자재를 선택해주세요.")
            return
            
        item = self.stock_tree.item(selection[0])
        mat_values = item['values']
        # [UX IMPROVEMENT] Get hidden MaterialID from the end of values
        mat_id = mat_values[-1]
        
        # Get full material data from DF (Robust matching)
        target_id_str = self.normalize_id(mat_id)
        mask = self.materials_df['MaterialID'].apply(self.normalize_id) == target_id_str
        matches = self.materials_df[mask]
        
        if matches.empty:
            messagebox.showerror("오류", f"자재 ID '{mat_id}'를 데이터베이스에서 찾을 수 없습니다.")
            edit_win.destroy()
            return
            
        mat_data = matches.iloc[0]
        
        # Create Edit Dialog
        edit_win = tk.Toplevel(self.root)
        edit_win.title("자재 정보 수정")
        edit_win.geometry("550x750")
        edit_win.transient(self.root)
        edit_win.grab_set()
        
        # [NEW] Scrollable area for many fields (User request: button visibility)
        outer_frame = ttk.Frame(edit_win)
        outer_frame.pack(fill='both', expand=True)
        
        canvas = tk.Canvas(outer_frame, highlightthickness=0, bg=getattr(self, 'theme_bg', '#f0f0f0'))
        scrollbar = ttk.Scrollbar(outer_frame, orient="vertical", command=canvas.yview)
        main_frame = ttk.Frame(canvas, padding=20)
        
        main_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=main_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Enable mousewheel scrolling (Safer local binding)
        def _on_mousewheel(event):
            try:
                if canvas.winfo_exists():
                    canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            except: pass
        edit_win.bind("<MouseWheel>", _on_mousewheel)
        
        fields = [
            ('회사코드', '회사코드'),
            ('관리품번', '관리품번'),
            ('자재명', '품목명'),
            ('SN', 'SN'),
            ('창고', '창고'),
            ('모델명', '모델명'),
            ('규격', '규격'),
            ('품목군코드', '품목군코드'),
            ('공급업체', '공급업체'),
            ('제조사', '제조사'),
            ('제조국', '제조국'),
            ('가격', '가격'),
            ('원가', '원가'),
            ('관리단위', '관리단위'),
            ('수량', '수량'),
            ('재고하한', '재고하한')
        ]
        
        entries = {}
        for i, (label_text, col_name) in enumerate(fields):
            ttk.Label(main_frame, text=f"{label_text}:").grid(row=i, column=0, padx=5, pady=5, sticky='w')
            
            # Using Combobox for some fields to maintain consistency
            if col_name in ['회사코드', '품목군코드', '제조사', '제조국', '관리단위']:
                ent = ttk.Combobox(main_frame, width=35)
                # [UNIFIED_UNIT] Populate values from managed list for units, others from data
                if col_name == '관리단위':
                    ent['values'] = getattr(self, 'unit_list', self.daily_units)
                elif col_name in self.materials_df.columns:
                    unique_vals = sorted(self.materials_df[col_name].dropna().unique().tolist())
                    ent['values'] = unique_vals
            else:
                ent = ttk.Entry(main_frame, width=38)
                
            ent.grid(row=i, column=1, padx=5, pady=5)
            
            # Pre-fill value
            val = mat_data.get(col_name, '')
            if pd.isna(val): val = ''
            ent.insert(0, str(val))
            entries[col_name] = ent
            
        def on_save():
            new_data = {col: ent.get() for col, ent in entries.items()}
            self.save_material_edits(mat_id, new_data)
            edit_win.destroy()
            
        btn_save = ttk.Button(main_frame, text="변경사항 저장", command=on_save)
        btn_save.grid(row=len(fields), column=0, columnspan=2, pady=20)


    def show_worker_popup(self, event, tree, is_menu=False):
        """Show a floating window (dropdown-style) with full record details for the clicked row"""
        item_id = tree.identify_row(event.y) if event and hasattr(event, 'y') else None
        if not item_id and tree.selection():
            item_id = tree.selection()[0]
        if not item_id: return

        column_id = tree.identify_column(event.x) if event and hasattr(event, 'x') else None

        # Determine target '작업자' column ID
        worker_col_id = None
        cols = list(tree['columns'])
        display_cols = list(tree['displaycolumns']) if tree['displaycolumns'] else cols
        for idx, col in enumerate(cols):
            if col == '작업자':
                try:
                    d_idx = display_cols.index(col)
                    worker_col_id = f"#{d_idx + 1}"
                except:
                    worker_col_id = f"#{idx + 1}"
                break

        if not is_menu:
            if not column_id: return
            try:
                col_name = tree.heading(column_id)['text']
            except:
                col_name = ''
            if col_name != '작업자' and column_id != worker_col_id: return
        else:
            if worker_col_id:
                column_id = worker_col_id
            elif not column_id and display_cols:
                column_id = "#1"

        tags = tree.item(item_id, 'tags')
        if tags and 'total' in tags: return
        if tree == getattr(self, 'daily_usage_tree', None):
            if not tags: return

        try:
            full_info = []
            if tree == getattr(self, 'daily_usage_tree', None):
                try:
                    df_idx = int(tags[0])
                    if df_idx not in self.daily_usage_df.index: return
                    entry = self.daily_usage_df.loc[df_idx]
                    
                    # 1. Basic Info
                    usage_date = self._safe_format_datetime(entry.get('Date', ''), '%Y-%m-%d')
                    full_info.append(f"[기본 기록 정보]")
                    full_info.append(f"• 날짜: {usage_date}")
                    full_info.append(f"• 현장: {self.clean_nan(entry.get('Site', ''))}")
                    
                    equip = self.clean_nan(entry.get('장비명', ''))
                    if equip: full_info.append(f"• 장비명: {equip}")
                    aux = self.clean_nan(entry.get("보조장비", ""))
                    if aux: full_info.append(f"• 보조장비: {aux}")
                    
                    method = self.clean_nan(entry.get('검사방법', ''))
                    if method: full_info.append(f"• 검사방법: {method}")
                    
                    amount = entry.get('검사량', 0)
                    if amount and float(amount) > 0: full_info.append(f"• 검사량: {amount}")
                    
                    
                    full_info.append(f"• 자재명: {self.get_material_display_name(entry.get('MaterialID'))}")
                    
                    # 2. Worker Info (All 10 slots)
                    full_info.append(f"\n[작업자 및 시간 정보]")
                    for i in range(1, 11):
                        u = self.clean_nan(entry.get('User' if i==1 else f'User{i}', ''))
                        if u:
                            t = self.clean_nan(entry.get('WorkTime' if i==1 else f'WorkTime{i}', ''))
                            o = self.clean_nan(entry.get('OT' if i==1 else f'OT{i}', ''))
                            full_info.append(f"• 작업자{i}: {u} | {t} | OT: {o}")
                    
                    # 3. Cost Info (Hide zeros)
                    cost_info = []
                    for k, label in [('단가', '단가'), ('출장비', '출장비'), ('일식', '일식'), ('검사비', '검사비')]:
                        raw_val = entry.get(k, 0)
                        try:
                            if pd.isna(raw_val) or str(raw_val).lower() == 'nan': continue
                            val_float = float(str(raw_val).replace(',', ''))
                            if val_float > 0:
                                cost_info.append(f"• {label}: {int(val_float):,}")
                        except: pass
                    if cost_info:
                        full_info.append(f"\n[비용 상세]")
                        full_info.extend(cost_info)
                    
                    # 4. RTK/NDT (Hide zeros)
                    rtk_found = []
                    rtk_cats = ["센터미스", "농도", "마킹미스", "필름마크", "취급부주의", "고객불만", "기타"]
                    for cat in rtk_cats:
                        val = entry.get(f'RTK_{cat}', 0)
                        if val and float(val) > 0:
                            rtk_found.append(f"  - {cat}: {val}")
                    if rtk_found:
                        full_info.append(f"\n[RT 결함 상세]")
                        full_info.extend(rtk_found)
                        
                    ndt_found = []
                    ndt_mats = ["형광자분", "흑색자분", "백색페인트", "침투제", "세척제", "현상제", "형광침투제"]
                    for mat in ndt_mats:
                        val = entry.get(f'NDT_{mat}', 0)
                        if pd.isna(val): # Fallback
                            if mat == "흑색자분": val = entry.get('NDT_자분', 0)
                            elif mat == "백색페인트": val = entry.get('NDT_페인트', 0)
                        if val and float(val) > 0:
                            ndt_found.append(f"  - {mat}: {val}")
                    if ndt_found:
                        full_info.append(f"\n[NDT 자재 사용량]")
                        full_info.extend(ndt_found)
                    
                    note = entry.get('Note', '')
                    if note and str(note).strip(): full_info.append(f"\n• 비고: {note}")
                    
                    full_info.append(f"• 입력시간: {entry.get('EntryTime', '')}")
                    
                except Exception as e: 
                    print(f"Popup error: {e}")
                    return
            else:
                # 월별 집계 트리뷰 (메인 및 분리 팝업창 공통)
                vals = tree.item(item_id, 'values')
                row_dict = dict(zip(cols, vals))

                year_val = row_dict.get('연도', '')
                month_val = row_dict.get('월', '')
                site = str(row_dict.get('현장', '')).strip()
                method = str(row_dict.get('검사방법', row_dict.get('구분', ''))).strip()
                mat_name = str(row_dict.get('품목명', '')).strip()
                qty = str(row_dict.get('수량', '')).strip()
                unit_price = str(row_dict.get('단가', '')).strip()
                fee = str(row_dict.get('금액', row_dict.get('검사비', ''))).strip()
                ot_hours = str(row_dict.get('OT시간', '')).strip()
                ot_amt = str(row_dict.get('OT금액', '')).strip()

                full_info.append("[월별 집계 기본 정보]")
                full_info.append(f"• 연월: {year_val}년 {month_val}월")
                full_info.append(f"• 현장: {site}")
                full_info.append(f"• 검사방법: {method} | 품목명: {mat_name}")
                if qty:
                    cost_line = f"• 수량: {qty}"
                    if unit_price: cost_line += f" | 단가: {unit_price}원"
                    if fee: cost_line += f" | 금액: {fee}원"
                    full_info.append(cost_line)
                if ot_hours or ot_amt:
                    full_info.append(f"• 총 OT시간: {ot_hours}시간 | 총 OT금액: {ot_amt}원")

                # Filter subset from current_monthly_df or daily_usage_df
                df_source = getattr(self, 'current_monthly_df', None)
                if df_source is None or df_source.empty:
                    df_source = getattr(self, 'daily_usage_df', None)

                subset_df = pd.DataFrame()
                if df_source is not None and not df_source.empty:
                    try:
                        y_int = int(year_val)
                        m_int = int(month_val)
                        s_df = df_source.copy()
                        if 'Year' not in s_df.columns or 'Month' not in s_df.columns:
                            s_df['Year'] = pd.to_datetime(s_df['Date'], errors='coerce').dt.year
                            s_df['Month'] = pd.to_datetime(s_df['Date'], errors='coerce').dt.month

                        mask = (s_df['Year'] == y_int) & (s_df['Month'] == m_int)
                        if site:
                            mask = mask & (s_df['Site'].astype(str).str.strip() == site)
                        if method and method not in ('미지정', '-'):
                            mask = mask & (s_df['검사방법'].astype(str).str.strip().str.upper() == method.upper())

                        # Material matching
                        matching_ids = []
                        if hasattr(self, 'materials_df') and not self.materials_df.empty:
                            matches = self.materials_df[self.materials_df['품목명'].astype(str).str.contains(mat_name, case=False, na=False)]
                            if not matches.empty:
                                matching_ids = matches['MaterialID'].tolist()

                        if mat_name and mat_name not in ('-', 'MT약품', 'PT약품', '품목 미지정'):
                            if not matching_ids:
                                if mat_name in s_df['MaterialID'].astype(str).values:
                                    matching_ids = [mat_name]
                                else:
                                    m_norm = mat_name.replace(' ', '').upper()
                                    for p_id in s_df['MaterialID'].dropna().unique():
                                        if str(p_id).replace(' ', '').upper() == m_norm:
                                            matching_ids.append(p_id)
                            if matching_ids:
                                mask = mask & (s_df['MaterialID'].isin(matching_ids))

                        subset_df = s_df[mask]
                    except Exception as ex:
                        print(f"Error filtering monthly subset for popup: {ex}")

                if not subset_df.empty:
                    def find_paired_cols_local(columns_list):
                        pairs = []
                        col_set = set(columns_list)
                        for i in range(1, 11):
                            u_n = 'User' if i == 1 else f'User{i}'
                            w_n = 'WorkTime' if i == 1 else f'WorkTime{i}'
                            o_n = 'OT' if i == 1 else f'OT{i}'
                            if u_n in col_set:
                                pairs.append((u_n, w_n if w_n in col_set else None, o_n if o_n in col_set else None))
                        return pairs

                    s_cols = [str(c).strip().replace(' ', '') for c in subset_df.columns]
                    subset_norm = subset_df.copy()
                    subset_norm.columns = s_cols
                    pairs = find_paired_cols_local(subset_norm.columns)

                    worker_rows = []
                    from site_apps.central.src.services.ot_calculator import ot_breakdown

                    for _, r in subset_norm.iterrows():
                        d_str = self._safe_format_datetime(r.get('Date', ''), '%Y-%m-%d')
                        r_note = self.clean_nan(r.get('Note', ''))
                        for uc, wc, oc in pairs:
                            u_name = self.clean_nan(r.get(uc, ''))
                            if u_name:
                                w_time = self.clean_nan(r.get(wc, '')) if wc else ''
                                o_val = self.clean_nan(r.get(oc, '')) if oc else ''
                                ot_res = ot_breakdown(w_time, r.get('Date'), o_val)
                                worker_rows.append({
                                    'Date': d_str,
                                    'Worker': u_name,
                                    'WorkTime': w_time,
                                    'OTValue': o_val,
                                    'OT_H': ot_res['hours'],
                                    'OT_A': ot_res['amount'],
                                    'Day_H': ot_res['day'],
                                    'Night_H': ot_res['night'],
                                    'Holiday_H': ot_res['holiday'],
                                    'Day_A': ot_res['day_amount'],
                                    'Night_A': ot_res['night_amount'],
                                    'Holiday_A': ot_res['holiday_amount'],
                                    'Note': r_note
                                })

                    if worker_rows:
                        w_df = pd.DataFrame(worker_rows)
                        w_dedup = w_df.drop_duplicates(subset=['Date', 'Worker', 'WorkTime', 'OTValue'])

                        # 공수 파싱 함수
                        import re as _re
                        def parse_work_count(val):
                            s = str(val).strip()
                            if not s: return 1.0
                            try:
                                m = _re.search(r'([0-9\.]+)', s)
                                if m: return float(m.group(1))
                            except: pass
                            return 1.0

                        w_dedup['Count'] = w_dedup['WorkTime'].apply(parse_work_count)

                        full_info.append("\n[참여 작업자 종합 요약]")
                        w_summary = w_dedup.groupby('Worker').agg({
                            'Count': 'sum',
                            'OT_H': 'sum',
                            'OT_A': 'sum',
                            'Day_H': 'sum',
                            'Night_H': 'sum',
                            'Holiday_H': 'sum',
                        }).reset_index()

                        for _, w_row in w_summary.iterrows():
                            w_name = w_row['Worker']
                            cnt = w_row['Count']
                            oth = w_row['OT_H']
                            ota = w_row['OT_A']
                            summary_str = f"• {w_name}: 총 {cnt:.1f}공수"
                            if oth > 0:
                                ot_parts = []
                                if w_row['Day_H'] > 0: ot_parts.append(f"연장 {w_row['Day_H']:.1f}h")
                                if w_row['Night_H'] > 0: ot_parts.append(f"야간 {w_row['Night_H']:.1f}h")
                                if w_row['Holiday_H'] > 0: ot_parts.append(f"휴일 {w_row['Holiday_H']:.1f}h")
                                summary_str += f" | 총 OT {oth:.1f}시간 ({int(ota):,}원)"
                                if ot_parts:
                                    summary_str += f" [{', '.join(ot_parts)}]"
                            full_info.append(summary_str)

                        full_info.append("\n[일자별 상세 투입 내역]")
                        dates_sorted = sorted(w_dedup['Date'].unique())
                        for d in dates_sorted:
                            d_records = w_dedup[w_dedup['Date'] == d]
                            full_info.append(f"📅 {d}")
                            for _, dr in d_records.iterrows():
                                line = f"   - {dr['Worker']}"
                                if dr['WorkTime']:
                                    line += f" ({dr['WorkTime']})"
                                if dr['OT_H'] > 0 or dr['OTValue']:
                                    line += f" | OT: {dr['OTValue'] or dr['OT_H']}h ({int(dr['OT_A']):,}원)"
                                if dr['Note']:
                                    line += f" ※ {dr['Note']}"
                                full_info.append(line)
                else:
                    full_w = row_dict.get('(Full작업자)', '')
                    if full_w:
                        full_info.append(f"\n[참여 작업자 목록]")
                        full_info.append(f"• {full_w}")

            if not full_info: return
            
            # Get cell bounding box for precise positioning
            bbox = tree.bbox(item_id, column_id) if column_id else None
            if not bbox and worker_col_id:
                bbox = tree.bbox(item_id, worker_col_id)
            if not bbox:
                for c in (display_cols if display_cols else cols):
                    b = tree.bbox(item_id, c)
                    if b:
                        bbox = b
                        break
            
            if bbox:
                x, y, w, h = bbox
                root_x = tree.winfo_rootx() + x
                root_y = tree.winfo_rooty() + y
            elif event and hasattr(event, 'x_root') and hasattr(event, 'y_root'):
                root_x = event.x_root
                root_y = event.y_root
                w, h = 120, 25
            else:
                root_x = tree.winfo_rootx() + 50
                root_y = tree.winfo_rooty() + 50
                w, h = 120, 25
            
            previous_popup = getattr(self, '_worker_detail_popup', None)
            if previous_popup is not None and previous_popup.winfo_exists():
                try: previous_popup.destroy()
                except: pass
            popup = tk.Toplevel(self.root)
            self._worker_detail_popup = popup
            popup.overrideredirect(True)
            popup.attributes("-topmost", True)
            popup.configure(bg='white', highlightbackground="#0078d7", highlightthickness=2)
            
            popup_width = max(w, 560)
            popup_height = 460
            
            # 1. Horizontal positioning (Center relative to cell)
            final_x = root_x - (popup_width - w) // 2
            
            # 2. Vertical positioning (Show ABOVE the cell by default)
            final_y = root_y - popup_height - 5
            
            # [SMART CHECK] If the popup would go off-screen at the top, show it BELOW instead
            if final_y < 0:
                final_y = root_y + h + 5

            # Screen bounds constraint
            try:
                screen_w = self.root.winfo_screenwidth()
                screen_h = self.root.winfo_screenheight()
                if final_x + popup_width > screen_w:
                    final_x = screen_w - popup_width - 10
                if final_x < 10:
                    final_x = 10
                if final_y + popup_height > screen_h:
                    final_y = screen_h - popup_height - 30
                if final_y < 10:
                    final_y = 10
            except: pass
            
            # Initial relative offsets from the cell root
            popup._rel_offset_x = final_x - root_x
            popup._rel_offset_y = final_y - root_y

            popup.geometry(f"{popup_width}x{popup_height}+{int(final_x)}+{int(final_y)}")
            
            # --- Custom Dragging Support ---
            def start_drag(event):
                popup._drag_data = {"x": event.x, "y": event.y}
            
            def do_drag(event):
                if hasattr(popup, '_drag_data'):
                    dx = event.x - popup._drag_data["x"]
                    dy = event.y - popup._drag_data["y"]
                    new_x = popup.winfo_x() + dx
                    new_y = popup.winfo_y() + dy
                    popup.geometry(f"+{new_x}+{new_y}")
                    
                    # Update relative offsets after manual drag so sync continues from new spot
                    try:
                        curr_bbox = tree.bbox(item_id, column_id)
                        if curr_bbox:
                            cx, cy, cw, ch = curr_bbox
                            # Important: Update relative to the CURRENT root position
                            popup._rel_offset_x = new_x - (tree.winfo_rootx() + cx)
                            popup._rel_offset_y = new_y - (tree.winfo_rooty() + cy)
                    except: pass
            
            popup.bind("<Button-1>", start_drag)
            popup.bind("<B1-Motion>", do_drag)
            
            # --- Header Frame with Title & Close Button ---
            header_frame = tk.Frame(popup, bg="#0078d7")
            header_frame.pack(fill='x', side='top')

            title_text = "👤 작업자 상세 투입 내역"
            if tree != getattr(self, 'daily_usage_tree', None) and year_val and month_val:
                title_parts = [f"[{year_val}년 {month_val}월]"]
                if site: title_parts.append(site)
                if method: title_parts.append(method)
                title_parts.append("작업자 투입 내역")
                title_text = "👤 " + " ".join(title_parts)

            title_lbl = tk.Label(header_frame, text=title_text, bg="#0078d7", fg="white", 
                                 font=('Malgun Gothic', 9, 'bold'))
            title_lbl.pack(side='left', padx=8, pady=5)

            def manual_close():
                if hasattr(popup, '_dismissal'):
                    popup._dismissal.close()
                elif popup.winfo_exists():
                    popup.destroy()

            close_btn = tk.Button(header_frame, text="✕", bg="#0078d7", fg="white", bd=0, 
                                  font=('Arial', 9, 'bold'), activebackground="#c42b1c", 
                                  activeforeground="white", cursor="hand2", command=manual_close)
            close_btn.pack(side='right', padx=6, pady=2)

            header_frame.bind("<Button-1>", start_drag)
            header_frame.bind("<B1-Motion>", do_drag)
            title_lbl.bind("<Button-1>", start_drag)
            title_lbl.bind("<B1-Motion>", do_drag)

            content_frame = ttk.Frame(popup, padding=10)
            content_frame.pack(fill='both', expand=True)
            
            lb = tk.Listbox(content_frame, font=('Malgun Gothic', 10), 
                            bg='white', fg='#333333', relief='flat', 
                            borderwidth=0, highlightthickness=0)
            for info in full_info:
                lb.insert(tk.END, info)
            
            lb.pack(side='left', fill='both', expand=True)
            
            sb = ttk.Scrollbar(content_frame, orient="vertical", command=lb.yview)
            lb.configure(yscrollcommand=sb.set)
            sb.pack(side='right', fill='y')
            
            # --- Window Move Synchronization ---
            def reposition_popup(event=None):
                if not popup.winfo_exists(): return
                # Re-calculate position based on CURRENT root position + saved relative offset
                try:
                    curr_bbox = tree.bbox(item_id, column_id)
                    if not curr_bbox: return
                    cx, cy, cw, ch = curr_bbox
                    croot_x = tree.winfo_rootx() + cx
                    croot_y = tree.winfo_rooty() + cy
                    
                    # Apply relative offset to keep the sync accurate
                    new_x = croot_x + popup._rel_offset_x
                    new_y = croot_y + popup._rel_offset_y
                    
                    popup.geometry(f"+{int(new_x)}+{int(new_y)}")
                except: pass
                
            # Bind to root move/resize
            bind_id = self.root.bind("<Configure>", reposition_popup, add="+")
            
            def cleanup_popup():
                try: self.root.unbind("<Configure>", bind_id)
                except tk.TclError: pass
                if getattr(self, '_worker_detail_popup', None) is popup:
                    self._worker_detail_popup = None

            from site_apps.central.src.views.detail_popup import DetailPopupDismissal
            popup._dismissal = DetailPopupDismissal(self.root, popup, lb, tree, cleanup_popup)
            
            # [FIX] Removed global self.root.bind("<Button-1>") which leaked and interfered with DateEntry.
            # FocusOut is sufficient for closing the popup when clicking elsewhere.
        except Exception as e:
            print(f"Final popup error: {e}")

    def open_list_management_dialog(self, title_or_key, data_list=None, config_key=None, target_cb=None):
        """Open a generic dialog to manage (edit/delete) items in a data list"""
        if self.layout_locked: return

        # If data_list is None, we assume title_or_key is the config_key
        if data_list is None:
            config_key = title_or_key
            data_map = {
                'sites': ('현장 목록 관리', self.sites),
                'users': ('담당자 목록 관리', getattr(self, 'users', [])),
                'equipments': ('장비 목록 관리', getattr(self, 'equipments', [])),
                'aux_equipments': ('보조장비 목록 관리', getattr(self, 'aux_equipments', [])),
                'vehicles': ('차량 목록 관리', getattr(self, 'vehicles', [])),
                'companies': ('업체 목록 관리', getattr(self, 'companies', [])),
                'materials': ('품목 목록 관리 (기본)', getattr(self, 'carestream_films', [])),
                'daily_units': ('단위 목록 관리', self.daily_units),
                'test_items': ('검사품명 목록 관리', getattr(self, 'test_items', [])),
                'applied_codes': ('적용코드 목록 관리', getattr(self, 'applied_codes', []))
            }
            if config_key not in data_map: return
            title, data_list = data_map[config_key]
        else:
            # Traditional 3-argument call: title, data_list, config_key
            title = title_or_key
            if config_key is None: config_key = title # Fallback

        # ── 이미 열린 창이 있으면 앞으로 가져오고 종료 ──────────────
        if not hasattr(self, '_list_mgmt_dialogs'):
            self._list_mgmt_dialogs = {}
        existing = self._list_mgmt_dialogs.get(config_key)
        if existing and existing.winfo_exists():
            existing.lift()
            existing.focus_set()
            return

        dialog = tk.Toplevel(self.root)
        self._list_mgmt_dialogs[config_key] = dialog  # 창 추적 등록
        dialog.title(title)
        dialog.geometry("600x450")
        dialog.transient(self.root)
        dialog.grab_set()
        
        frame = ttk.Frame(dialog, padding=10)
        frame.pack(fill='both', expand=True)
        
        # [NEW] Search box for the list
        search_f = ttk.Frame(frame)
        search_f.pack(fill='x', pady=(0, 5))
        ttk.Label(search_f, text="🔍:").pack(side='left', padx=2)
        search_var = tk.StringVar()
        search_ent = ttk.Entry(search_f, textvariable=search_var)
        search_ent.pack(side='left', fill='x', expand=True, padx=2)
        search_ent.focus_set()

        # Add scrollbar
        scrollbar = ttk.Scrollbar(frame)
        scrollbar.pack(side='right', fill='y')
        
        listbox = tk.Listbox(frame, font=('Arial', 10))
        listbox.pack(fill='both', expand=True, side='left')
        listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=listbox.yview)
        
        def refresh_list(*args):
            query = search_var.get().strip().lower()
            listbox.delete(0, 'end')
            # Use original data_list for filtering
            for item in sorted(data_list):
                if not query or query in item.lower():
                    listbox.insert('end', item)
            # Trigger app-wide update of related comboboxes
            self.refresh_ui_for_list_change(config_key)

        search_var.trace_add("write", refresh_list)
        refresh_list()
            
        def select_and_close():
            sel = listbox.curselection()
            if not sel: return
            val = listbox.get(sel[0])
            if target_cb:
                # [FIX] Support both Entry and Combobox widgets
                if hasattr(target_cb, 'set'):
                    target_cb.set(val)
                else:
                    target_cb.delete(0, tk.END)
                    target_cb.insert(0, val)
                # Manually trigger any search/filter logic if needed
                if hasattr(self, 'update_daily_usage_view'):
                    self.update_daily_usage_view()
            dialog.destroy()

        listbox.bind("<Double-1>", lambda e: select_and_close())
            
        btn_frame = ttk.Frame(dialog, padding=5)
        btn_frame.pack(fill='x')
        
        # [NEW] Select button (Shows only if target_cb is present, or always as an alternative to double-click)
        ttk.Button(btn_frame, text="✅ 선택", command=select_and_close).pack(side='left', padx=2)

        def add_item():
            new_val = simpledialog.askstring("추가", "새 이름을 입력하세요:")
            if new_val and new_val.strip():
                val = new_val.strip()
                if val not in data_list:
                    data_list.append(val)
                    data_list.sort()
                    self.save_tab_config()
                    refresh_list()
                else:
                    messagebox.showinfo("정보", "이미 목록에 있는 이름입니다.")

        def edit_item():
            sel = listbox.curselection()
            if not sel: return
            idx = sel[0]
            old_val = listbox.get(idx)
            new_val = simpledialog.askstring("수정", f"[{old_val}] -> 새 이름을 입력하세요:", initialvalue=old_val)
            if new_val and new_val.strip() and new_val != old_val:
                if old_val in data_list: data_list.remove(old_val)
                data_list.append(new_val.strip())
                data_list.sort()
                self.save_tab_config()
                refresh_list()

        def delete_item():
            sel = listbox.curselection()
            if not sel: return
            idx = sel[0]
            val = listbox.get(idx)
            if messagebox.askyesno("삭제 확인", f"'{val}'을 목록에서 삭제하시겠습니까?"):
                if val in data_list: data_list.remove(val)
                # [NEW] If the deleted value is currently selected in the target combobox, clear it
                if target_cb:
                    current_val = target_cb.get().strip()
                    if current_val == val.strip():
                        if hasattr(target_cb, 'set'):
                            target_cb.set('')
                        else:
                            target_cb.delete(0, tk.END)
                self.save_tab_config()
                refresh_list()

        def on_close():
            self._list_mgmt_dialogs.pop(config_key, None)
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", on_close)

        def bulk_register_to_stock():
            if messagebox.askyesno("일괄 등록 확인", "현재 목록에 있는 모든 품목을 '재고 현황' 데이터베이스에 새로 등록하시겠습니까?\n(이미 등록된 품목은 중복 등록되지 않습니다.)"):
                self.bulk_register_materials_to_stock(data_list)

        ttk.Button(btn_frame, text="추가", command=add_item).pack(side='left', padx=5, expand=True)
        ttk.Button(btn_frame, text="수정", command=edit_item).pack(side='left', padx=5, expand=True)
        ttk.Button(btn_frame, text="삭제", command=delete_item).pack(side='left', padx=5, expand=True)
        
        if config_key == 'materials':
            ttk.Button(btn_frame, text="📦 재고 일괄 등록", command=bulk_register_to_stock).pack(side='left', padx=5, expand=True)
            
        ttk.Button(btn_frame, text="닫기", command=on_close).pack(side='left', padx=5, expand=True)


    def open_equipment_management_dialog(self):
        """Dialog to add, edit or delete equipment names in the list"""
        dlg = tk.Toplevel(self.root)
        dlg.title("장비명 목록 관리")
        dlg.geometry("900x700") # Increased size
        dlg.transient(self.root)
        dlg.grab_set()

        main_frame = ttk.Frame(dlg, padding=20)
        main_frame.pack(fill='both', expand=True)

        ttk.Label(main_frame, text="사용자 정의 장비 목록", font=('Malgun Gothic', 12, 'bold')).pack(pady=(0, 5))
        ttk.Label(main_frame, text="(드롭다운 목록에 표시될 기본 장비 명단입니다)", font=('Malgun Gothic', 9)).pack(pady=(0, 10))
        
        # Treeview to show equipment list
        list_container = ttk.Frame(main_frame)
        list_container.pack(fill='both', expand=True)

        # Left side: Current Equipment List
        left_frame = ttk.LabelFrame(list_container, text="현재 장비 목록", padding=5)
        left_frame.pack(side='left', fill='both', expand=True, padx=(0, 5))
        
        columns = ('name',)
        self.equip_tree = ttk.Treeview(left_frame, columns=columns, show='headings', height=15)
        self.equip_tree.heading('name', text='장비명')
        self.equip_tree.column('name', width=350) # Increased width
        
        scrollbar_l_y = ttk.Scrollbar(left_frame, orient="vertical", command=self.equip_tree.yview)
        scrollbar_l_x = ttk.Scrollbar(left_frame, orient="horizontal", command=self.equip_tree.xview)
        self.equip_tree.configure(yscrollcommand=scrollbar_l_y.set, xscrollcommand=scrollbar_l_x.set)
        
        self.equip_tree.pack(side='top', fill='both', expand=True)
        scrollbar_l_y.pack(side='right', fill='y', in_=left_frame)
        scrollbar_l_x.pack(side='bottom', fill='x')

        # Right side: Materials from Stock Status
        right_frame = ttk.LabelFrame(list_container, text="재고현황 품목 (가져오기)", padding=5)
        right_frame.pack(side='right', fill='both', expand=True, padx=(5, 0))

        # Search Bar for Right List
        search_frame = ttk.Frame(right_frame)
        search_frame.pack(fill='x', pady=(0, 5))
        ttk.Label(search_frame, text="🔍 검색:").pack(side='left', padx=2)
        search_var = tk.StringVar()
        search_ent = ttk.Entry(search_frame, textvariable=search_var)
        search_ent.pack(side='left', fill='x', expand=True, padx=2)

        columns_r = ('name', 'sn', 'model')
        self.mat_source_tree = ttk.Treeview(right_frame, columns=columns_r, show='headings', height=15)
        self.mat_source_tree.heading('name', text='자재명')
        self.mat_source_tree.heading('sn', text='SN')
        self.mat_source_tree.heading('model', text='모델명')
        self.mat_source_tree.column('name', width=150)
        self.mat_source_tree.column('sn', width=120)
        self.mat_source_tree.column('model', width=150)

        scrollbar_r_y = ttk.Scrollbar(right_frame, orient="vertical", command=self.mat_source_tree.yview)
        scrollbar_r_x = ttk.Scrollbar(right_frame, orient="horizontal", command=self.mat_source_tree.xview)
        self.mat_source_tree.configure(yscrollcommand=scrollbar_r_y.set, xscrollcommand=scrollbar_r_x.set)
        
        self.mat_source_tree.pack(side='top', fill='both', expand=True)
        scrollbar_r_y.pack(side='right', fill='y', in_=right_frame)
        scrollbar_r_x.pack(side='bottom', fill='x')

        def refresh_tree():
            for item in self.equip_tree.get_children():
                self.equip_tree.delete(item)
            for eq in sorted(self.equipments):
                self.equip_tree.insert('', tk.END, values=(eq,))

        def refresh_mat_tree(*args):
            query = search_var.get().strip().lower()
            for item in self.mat_source_tree.get_children():
                self.mat_source_tree.delete(item)
            
            if not self.materials_df.empty:
                # Group by unique combinations
                unique_items = self.materials_df[['품목명', 'SN', '모델명']].drop_duplicates()
                for _, row in unique_items.iterrows():
                    name = str(row.get('품목명', '')).strip()
                    sn = str(row.get('SN', '')).strip()
                    model = str(row.get('모델명', '')).strip()
                    
                    if name == 'nan': name = ''
                    if sn == 'nan': sn = ''
                    if model == 'nan': model = ''
                    
                    if name or sn or model:
                        # Apply filter
                        if not query or query in name.lower() or query in sn.lower() or query in model.lower():
                            self.mat_source_tree.insert('', tk.END, values=(name, sn, model))

        search_var.trace_add("write", refresh_mat_tree)
        refresh_tree()
        refresh_mat_tree()

        # Action Buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill='x', pady=10)

        def add_from_stock():
            selection = self.mat_source_tree.selection()
            if not selection:
                messagebox.showwarning("선택 오류", "가져올 품목을 오른쪽 목록에서 선택해주세요.")
                return
            
            added_count = 0
            for sel in selection:
                vals = self.mat_source_tree.item(sel)['values']
                name = str(vals[0]).strip()
                sn = str(vals[1]).strip()
                model = str(vals[2]).strip()
                
                # Cleanup 'nan' strings
                if name.lower() == 'nan': name = ''
                if sn.lower() == 'nan': sn = ''
                if model.lower() == 'nan': model = ''

                # Create a display string for the equipment list
                # Format: [Model] Name (SN)
                display_parts = []
                
                # Logic: If model exists, add it. If name exists and is different from model, add it.
                if model:
                    display_parts.append(model)
                if name and name != model:
                    display_parts.append(name)
                if sn:
                    display_parts.append(f"({sn})")
                
                display_name = " ".join(display_parts).strip()
                
                if display_name and display_name not in self.equipments:
                    self.equipments.append(display_name)
                    added_count += 1
            
            if added_count > 0:
                refresh_tree()
                self.save_tab_config()
                self.update_material_combo()
                messagebox.showinfo("성공", f"{added_count}개의 품목을 장비 목록에 추가했습니다.")

        def add_equip():
            def save_new():
                name = entry.get().strip()
                if name:
                    if name not in self.equipments:
                        self.equipments.append(name)
                        refresh_tree()
                        refresh_mat_tree()
                        self.save_tab_config()
                        self.update_material_combo()
                        add_win.destroy()
                    else:
                        messagebox.showinfo("중복", "이미 존재하는 장비명입니다.")
                else:
                    messagebox.showwarning("입력 오류", "장비명을 입력해주세요.")

            add_win = tk.Toplevel(dlg)
            add_win.title("장비 추가")
            add_win.geometry("600x150")
            add_win.transient(dlg)
            add_win.grab_set()
            
            ttk.Label(add_win, text="추가할 장비명을 입력하세요:", font=('Malgun Gothic', 10)).pack(pady=10)
            entry = ttk.Entry(add_win, width=70)
            entry.pack(padx=20, pady=5)
            entry.focus_set()
            
            btn_f = ttk.Frame(add_win)
            btn_f.pack(pady=10)
            ttk.Button(btn_f, text="확인", command=save_new, width=10).pack(side='left', padx=5)
            ttk.Button(btn_f, text="취소", command=add_win.destroy, width=10).pack(side='left', padx=5)
            
            add_win.bind('<Return>', lambda e: save_new())

        def delete_equip():
            selection = self.equip_tree.selection()
            if not selection:
                messagebox.showwarning("선택 오류", "삭제할 항목을 선택해주세요.")
                return
            
            count = len(selection)
            if messagebox.askyesno("삭제 확인", f"선택한 {count}개의 장비를 목록에서 삭제하시겠습니까?"):
                for sel in selection:
                    name = self.equip_tree.item(sel)['values'][0]
                    if name in self.equipments:
                        self.equipments.remove(name)
                
                refresh_tree()
                refresh_mat_tree()
                self.save_tab_config()
                self.update_material_combo()

        def edit_equip():
            selection = self.equip_tree.selection()
            if not selection:
                messagebox.showwarning("선택 오류", "수정할 항목을 선택해주세요.")
                return
            
            old_name = self.equip_tree.item(selection[0])['values'][0]
            
            def save_edit():
                new_name = entry.get().strip()
                if new_name and new_name != old_name:
                    if old_name in self.equipments:
                        idx = self.equipments.index(old_name)
                        self.equipments[idx] = new_name
                        refresh_tree()
                        refresh_mat_tree()
                        self.save_tab_config()
                        self.update_material_combo()
                        edit_win.destroy()
                elif new_name == old_name:
                    edit_win.destroy()
                else:
                    messagebox.showwarning("입력 오류", "장비명을 입력해주세요.")

            edit_win = tk.Toplevel(dlg)
            edit_win.title("장비명 수정")
            edit_win.geometry("600x150")
            edit_win.transient(dlg)
            edit_win.grab_set()
            
            ttk.Label(edit_win, text="장비명을 수정하세요:", font=('Malgun Gothic', 10)).pack(pady=10)
            entry = ttk.Entry(edit_win, width=70)
            entry.pack(padx=20, pady=5)
            entry.insert(0, old_name)
            entry.selection_range(0, tk.END)
            entry.focus_set()
            
            btn_f = ttk.Frame(edit_win)
            btn_f.pack(pady=10)
            ttk.Button(btn_f, text="저장", command=save_edit, width=10).pack(side='left', padx=5)
            ttk.Button(btn_f, text="취소", command=edit_win.destroy, width=10).pack(side='left', padx=5)
            
            edit_win.bind('<Return>', lambda e: save_edit())

        ttk.Button(btn_frame, text="◀ 가져오기", command=add_from_stock).pack(side='left', padx=5, expand=True, fill='x')
        ttk.Button(btn_frame, text="➕ 직접추가", command=add_equip).pack(side='left', padx=5, expand=True, fill='x')
        ttk.Button(btn_frame, text="✏️ 수정", command=edit_equip).pack(side='left', padx=5, expand=True, fill='x')
        ttk.Button(btn_frame, text="❌ 삭제", command=delete_equip).pack(side='left', padx=5, expand=True, fill='x')

        ttk.Button(main_frame, text="닫기", command=dlg.destroy).pack(pady=10)


    def open_material_search_dialog(self, target_form='transaction'):
        """Unified dialog to search materials, edit them, or select for forms"""
        dlg = tk.Toplevel(self.root)
        dlg.title("자재 마스터 검색 및 관리")
        dlg.geometry("900x700")
        dlg.transient(self.root)
        dlg.grab_set()

        main_frame = ttk.Frame(dlg, padding=20)
        main_frame.pack(fill='both', expand=True)

        header_text = "전체 재고 자재 검색 및 수정" if target_form == 'registration' else "자재 검색 및 선택"
        ttk.Label(main_frame, text=header_text, font=('Malgun Gothic', 12, 'bold')).pack(pady=(0, 10))
        include_unknown = tk.BooleanVar(master=dlg, value=False)
        if target_form == 'daily_usage':
            ttk.Label(main_frame, text='선택한 검사방법의 자재만 표시 · 기존 입력값은 유지됩니다.').pack(anchor='w')
            ttk.Checkbutton(main_frame, text='미분류 포함', variable=include_unknown,
                            command=lambda: refresh_list()).pack(anchor='w')
        
        search_frame = ttk.Frame(main_frame)
        search_frame.pack(fill='x', pady=5)
        ttk.Label(search_frame, text="🔍 검색어:").pack(side='left', padx=5)
        search_var = tk.StringVar()
        search_ent = ttk.Entry(search_frame, textvariable=search_var)
        search_ent.pack(side='left', fill='x', expand=True, padx=5)
        search_ent.focus_set()

        # Treeview for results
        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(fill='both', expand=True)
        
        columns = ('id', 'name', 'model', 'sn', 'spec', 'unit', 'stock')
        tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=15)
        tree.heading('id', text='ID')
        tree.heading('name', text='자재명')
        tree.heading('model', text='모델명')
        tree.heading('sn', text='SN')
        tree.heading('spec', text='규격')
        tree.heading('unit', text='단위')
        tree.heading('stock', text='재고')
        
        tree.column('id', width=50)
        tree.column('name', width=180)
        tree.column('model', width=150)
        tree.column('sn', width=100)
        tree.column('spec', width=100)
        tree.column('unit', width=60)
        tree.column('stock', width=60)
        
        sb_y = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb_y.set)
        
        tree.pack(side='left', fill='both', expand=True)
        sb_y.pack(side='right', fill='y')

        def refresh_list(*args):
            query = search_var.get().strip().lower()
            for item in tree.get_children():
                tree.delete(item)
            
            if not self.materials_df.empty:
                for idx, row in self.materials_df.iterrows():
                    # [STRICT] Filter out inactive items (Active=0) from search dialog
                    # Use a robust check that handles both Series and scalars
                    active_val = row.get('Active', 1)
                    num_active = pd.to_numeric(active_val, errors='coerce')
                    is_active = (num_active > 0.5) if pd.notna(num_active) else True
                    
                    if not is_active:
                        continue
                    if target_form == 'daily_usage':
                        from site_apps.central.src.method_catalog import row_allowed, current_method
                        if current_method(self) and not row_allowed(self, row, include_unknown=include_unknown.get()):
                            continue
                    
                    name = str(row.get('품목명', '')).strip()
                    sn = str(row.get('SN', '')).strip()
                    model = str(row.get('모델명', '')).strip()
                    spec = str(row.get('규격', '')).strip()
                    unit = str(row.get('관리단위', '')).strip()
                    stock = str(row.get('수량', '0'))
                    # [NEW] Format MaterialID as integer for display
                    mid_raw = row.get('MaterialID', '')
                    mid = ""
                    try:
                        if pd.notna(mid_raw) and str(mid_raw).strip():
                            mid = str(int(float(str(mid_raw).strip())))
                    except:
                        mid = str(mid_raw)
                    
                    
                    # Clean nans
                    name = '' if name.lower() == 'nan' else name
                    sn = '' if sn.lower() == 'nan' else sn
                    model = '' if model.lower() == 'nan' else model
                    spec = '' if spec.lower() == 'nan' else spec
                    
                    if name or sn or model:
                        display_name = name if is_active else f"[비활성] {name}"
                        row_str = f"{name} {sn} {model} {spec} {unit}".lower()
                        if not query or query in row_str:
                            tag = 'active' if is_active else 'inactive'
                            tree.insert('', tk.END, values=(mid, display_name, model, sn, spec, unit, stock), tags=(tag, idx))
                
                tree.tag_configure('inactive', foreground='gray')
                tree.tag_configure('active', foreground='black')

        def on_select(e=None):
            selection = tree.selection()
            if not selection: return
            
            vals = tree.item(selection[0])['values']
            name = str(vals[1])
            model = str(vals[2])
            sn = str(vals[3])
            
            if target_form == 'transaction':
                self.cb_material.set(name)
                self.on_material_selected()
                dlg.destroy()
            elif target_form == 'daily_usage':
                from site_apps.central.src.material_selection import choose_daily_material
                choose_daily_material(self, vals[0])
                dlg.destroy()
            else:
                # Registration form
                if self._is_consumable_material(name, ''):
                    self.cb_item_name.set(name)
                    if hasattr(self, 'cb_eq_name'): self.cb_eq_name.set('')
                else:
                    if hasattr(self, 'cb_eq_name'): self.cb_eq_name.set(name)
                    self.cb_item_name.set('')
                if hasattr(self, 'cb_model'): self.cb_model.set(model)
                if hasattr(self, 'ent_sn'): 
                    self.ent_sn.delete(0, tk.END)
                    self.ent_sn.insert(0, sn)
                dlg.destroy()

        def edit_material():
            selection = tree.selection()
            if not selection:
                messagebox.showwarning("선택 오류", "수정할 항목을 선택해주세요.")
                return
            
            item_id = tree.item(selection[0])['values'][0]
            # Find in DF
            mask = self.materials_df['MaterialID'].astype(str) == str(item_id)
            if not mask.any(): return
            
            row_idx = self.materials_df.index[mask][0]
            mat_data = self.materials_df.loc[row_idx]

            edit_win = tk.Toplevel(dlg)
            edit_win.title(f"자재 정보 수정 - ID: {item_id}")
            edit_win.geometry("400x500")
            edit_win.transient(dlg)
            edit_win.grab_set()

            f = ttk.Frame(edit_win, padding=20)
            f.pack(fill='both', expand=True)

            fields = [
                ("자재명", "품목명"), ("모델명", "모델명"), ("SN", "SN"),
                ("규격", "규격"), ("단위", "관리단위"), ("수량", "수량"), ("재고하한", "재고하한")
            ]
            entries = {}
            for i, (label, col) in enumerate(fields):
                ttk.Label(f, text=f"{label}:").grid(row=i, column=0, padx=5, pady=5, sticky='w')
                ent = ttk.Entry(f, width=30)
                ent.grid(row=i, column=1, padx=5, pady=5, sticky='ew')
                val = str(mat_data.get(col, ''))
                if val.lower() == 'nan': val = ''
                ent.insert(0, val)
                entries[col] = ent

            def save_edit():
                for col, ent in entries.items():
                    val = ent.get().strip()
                    if col in ['수량', '재고하한']:
                        try: val = float(val) if val else 0.0
                        except: val = 0.0
                    self.materials_df.at[row_idx, col] = val
                
                self.save_data()
                self.update_stock_view()
                self.update_material_combo()
                refresh_list()
                messagebox.showinfo("수정 완료", "자재 정보가 업데이트되었습니다.")
                edit_win.destroy()

            ttk.Button(edit_win, text="변경사항 저장", command=save_edit).pack(pady=10)

        def add_new_material():
            """Small dialog to add a missing material quickly"""
            add_win = tk.Toplevel(dlg)
            add_win.title("자재 신규 등록")
            add_win.geometry("450x350")
            add_win.transient(dlg)
            add_win.grab_set()

            form_f = ttk.Frame(add_win, padding=20)
            form_f.pack(fill='both', expand=True)

            entries = {}
            labels = [("자재명", "name"), ("모델명", "model"), ("SN", "sn"), ("규격", "spec"), ("단위", "unit")]
            
            for i, (label, key) in enumerate(labels):
                ttk.Label(form_f, text=f"{label}:").grid(row=i, column=0, padx=5, pady=5, sticky='w')
                ent = ttk.Entry(form_f, width=40)
                ent.grid(row=i, column=1, padx=5, pady=5, sticky='ew')
                entries[key] = ent
                if key == "name":
                    ent.insert(0, search_var.get().strip())
                    ent.focus_set()
                if key == "unit":
                    ent.insert(0, "매")

            def save_new():
                data = {k: e.get().strip() for k, e in entries.items()}
                if not data['name']:
                    messagebox.showwarning("오류", "자재명은 필수입니다.")
                    return
                
                new_id = 10001
                if not self.materials_df.empty:
                    valid_ids = pd.to_numeric(self.materials_df['MaterialID'], errors='coerce').dropna()
                    if not valid_ids.empty:
                        new_id = int(valid_ids.max()) + 1

                new_row = {
                    'MaterialID': new_id,
                    '품목명': data['name'],
                    '모델명': data['model'],
                    'SN': data['sn'],
                    '규격': data['spec'],
                    '관리단위': data['unit'],
                    'Active': 1,
                    '수량': 0,
                    '재고하한': 0
                }
                for col in self.materials_df.columns:
                    if col not in new_row: new_row[col] = ''
                
                self.materials_df = pd.concat([self.materials_df, pd.DataFrame([new_row])], ignore_index=True)
                self.save_data()
                self.update_material_combo()
                self.update_stock_view()
                refresh_list()
                messagebox.showinfo("완료", f"'{data['name']}' 자재가 성공적으로 등록되었습니다.")
                add_win.destroy()

            ttk.Button(add_win, text="등록 완료", command=save_new).pack(pady=10)

        search_var.trace_add("write", refresh_list)
        tree.bind('<Double-1>', on_select)
        refresh_list()

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="✅ 선택 완료", command=on_select, width=15).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="✏️ 정보 수정", command=edit_material, width=15).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="➕ 직접 등록", command=add_new_material, width=15).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="취소", command=dlg.destroy, width=10).pack(side='left', padx=5)


    def open_edit_daily_usage_dialog(self):
        """Open a dialog to edit the selected daily usage record"""
        selection = self.daily_usage_tree.selection()
        if not selection:
            messagebox.showwarning("선택 오류", "수정할 항목을 선택해주세요.")
            return
            
        item = self.daily_usage_tree.item(selection[0])
        tags = item.get('tags', [])
        if not tags:
            messagebox.showwarning("데이터 오류", "선택한 항목의 데이터 정보를 찾을 수 없습니다.")
            return
            
        try:
            df_idx = int(tags[0])
            entry_data = self.daily_usage_df.loc[df_idx]
        except (ValueError, KeyError, IndexError):
            messagebox.showwarning("데이터 오류", "데이터를 불러오는 중 오류가 발생했습니다.")
            return

        # Create Edit Dialog
        edit_win = tk.Toplevel(self.root)
        edit_win.title("일일 사용 기록 수정")
        edit_win.geometry("800x900")
        edit_win.transient(self.root)
        edit_win.grab_set()

        # Main scrollable frame
        container = ttk.Frame(edit_win)
        container.pack(expand=True, fill='both')
        
        canvas = tk.Canvas(container)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, padding=20)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Scroll binding is handled globally in MaterialManager.__init__

        fields = {}
        
        # 1. Basic Info
        basic_frame = ttk.LabelFrame(scrollable_frame, text="기본 정보", padding=10)
        basic_frame.pack(fill='x', pady=5)
        
        # Date
        ttk.Label(basic_frame, text="날짜:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        ent_date = DateEntry(basic_frame, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd', locale='ko_KR', state='readonly', showweeknumbers=True)
        curr_date = entry_data.get('Date', datetime.datetime.now().strftime('%Y-%m-%d'))
        try: ent_date.set_date(pd.to_datetime(curr_date))
        except: pass
        ent_date.grid(row=0, column=1, padx=5, pady=5, sticky='w')
        fields['Date'] = ent_date

        # Site
        ttk.Label(basic_frame, text="현장:").grid(row=0, column=2, padx=5, pady=5, sticky='w')
        cb_site = ttk.Combobox(basic_frame, width=20, values=self.sites)
        cb_site.set(self.clean_nan(entry_data.get('Site', '')))
        cb_site.grid(row=0, column=3, padx=5, pady=5, sticky='w')
        fields['Site'] = cb_site
        
        # Company & Applied Code
        ttk.Label(basic_frame, text="업체명:").grid(row=1, column=0, padx=5, pady=5, sticky='w')
        cb_company = ttk.Combobox(basic_frame, width=20, values=self.companies)
        cb_company.set(self.clean_nan(entry_data.get('업체명', '')))
        cb_company.grid(row=1, column=1, padx=5, pady=5, sticky='w')
        fields['업체명'] = cb_company
        
        ttk.Label(basic_frame, text="적용코드:").grid(row=1, column=2, padx=5, pady=5, sticky='w')
        ent_code = ttk.Entry(basic_frame, width=12)
        ent_code.insert(0, self.clean_nan(entry_data.get('적용코드', '')))
        ent_code.grid(row=1, column=3, padx=5, pady=5, sticky='w')
        fields['적용코드'] = ent_code

        # Material
        ttk.Label(basic_frame, text="자재명:").grid(row=2, column=0, padx=5, pady=5, sticky='w')
        cb_mat = ttk.Combobox(basic_frame, width=40)
        mat_options = []
        for _, m_row in self.materials_df.iterrows():
            if m_row.get('Active', 1) == 0: continue
            mat_options.append(self.get_material_display_name(m_row['MaterialID']))
        
        # Deduplicate and sort for a clean list
        mat_options = sorted(list(set([m for m in mat_options if m])))
        cb_mat['values'] = mat_options
        
        # Determine current display name
        curr_mat_id = entry_data.get('MaterialID')
        curr_mat_display = self.get_material_display_name(curr_mat_id)
        
        # [FIX] Use raw MaterialID string if not found in master data (e.g. JIREH Scanner)
        if not curr_mat_display and curr_mat_id and not pd.isna(curr_mat_id):
            if str(curr_mat_id).strip().lower() not in ('nan', ''):
                curr_mat_display = str(curr_mat_id).strip()
        
        cb_mat.set(curr_mat_display)
        cb_mat.grid(row=2, column=1, columnspan=3, padx=5, pady=5, sticky='w')
        fields['Material'] = cb_mat

        # Equipment & Method
        ttk.Label(basic_frame, text="장비명:").grid(row=3, column=0, padx=5, pady=5, sticky='w')
        cb_equip = ttk.Combobox(basic_frame, width=20, values=self.equipments)
        cb_equip.set(self.clean_nan(entry_data.get('장비명', '')))
        cb_equip.grid(row=3, column=1, padx=5, pady=5, sticky='w')
        fields['장비명'] = cb_equip

        ttk.Label(basic_frame, text="검사방법:").grid(row=3, column=2, padx=5, pady=5, sticky='w')
        cb_method = ttk.Combobox(basic_frame, width=10, values=['RT', 'PAUT', 'UT', 'MT', 'PT', 'ETC'])
        cb_method.set(self.clean_nan(entry_data.get('검사방법', '')))
        cb_method.grid(row=3, column=3, padx=5, pady=5, sticky='w')
        fields['검사방법'] = cb_method
        
        # [V13_FIX] Add Unit to basic info - Now using dynamic list from unit management
        ttk.Label(basic_frame, text="단위:").grid(row=3, column=4, padx=5, pady=5, sticky='w')
        cb_unit = ttk.Combobox(basic_frame, width=10, values=self.daily_units)
        cb_unit.set(self.clean_nan(entry_data.get('Unit', entry_data.get('단위', '매'))))
        cb_unit.grid(row=3, column=5, padx=5, pady=5, sticky='w')
        fields['Unit'] = cb_unit
        
        from site_apps.central.src.aux_equipment import add_selector, names
        aux_var = tk.StringVar(value='; '.join(names(entry_data.get('보조장비', ''))))
        ttk.Label(basic_frame, text='보조장비:').grid(row=5, column=0, padx=5, pady=5, sticky='w')
        add_selector(basic_frame, self, aux_var).grid(row=5, column=1, columnspan=5, padx=5, pady=5, sticky='w')

        # Inspection Item
        ttk.Label(basic_frame, text="검사품명:").grid(row=4, column=0, padx=5, pady=5, sticky='w')
        ent_item = ttk.Entry(basic_frame, width=40)
        ent_item.insert(0, self.clean_nan(entry_data.get('검사품명', '')))
        ent_item.grid(row=4, column=1, columnspan=3, padx=5, pady=5, sticky='w')
        fields['검사품명'] = ent_item

        # 2. Quantities & Costs
        cost_frame = ttk.LabelFrame(scrollable_frame, text="수량 및 비용", padding=10)
        cost_frame.pack(fill='x', pady=5)
        
        qty_configs = [
            ('검사량', '검사량'), ('단가', '단가'), ('출장비', '출장비'),
            ('일식', '일식'), ('검사비', '검사비'), ('필름매수', 'FilmCount')
        ]
        for i, (lbl, key) in enumerate(qty_configs):
            row = i // 3
            col = (i % 3) * 2
            ttk.Label(cost_frame, text=f"{lbl}:").grid(row=row, column=col, padx=5, pady=5, sticky='w')
            ent = ttk.Entry(cost_frame, width=12)
            val = entry_data.get(key, 0)
            if pd.isna(val): val = 0
            ent.insert(0, str(val))
            ent.grid(row=row, column=col+1, padx=5, pady=5, sticky='w')
            fields[key] = ent

        # 3. RTK & NDT
        material_frame = ttk.Frame(scrollable_frame)
        material_frame.pack(fill='x', pady=5)
        
        rtk_frame = ttk.LabelFrame(material_frame, text="RTK 사용량 (재촬영)", padding=10)
        rtk_frame.pack(side='left', fill='both', expand=True, padx=(0, 5))
        
        rtk_cats = ["센터미스", "농도", "마킹미스", "필름마크", "취급부주의", "고객불만", "기타"]
        for i, cat in enumerate(rtk_cats):
            ttk.Label(rtk_frame, text=f"{cat}:").grid(row=i, column=0, padx=5, pady=2, sticky='w')
            ent = ttk.Entry(rtk_frame, width=10)
            val = entry_data.get(f'RTK_{cat}', 0)
            if pd.isna(val): val = 0
            ent.insert(0, str(val))
            ent.grid(row=i, column=1, padx=5, pady=2, sticky='w')
            fields[f'RTK_{cat}'] = ent

        ndt_frame = ttk.LabelFrame(material_frame, text="NDT 약품 사용량", padding=10)
        ndt_frame.pack(side='left', fill='both', expand=True, padx=(5, 0))
        
        ndt_cats = ["형광자분", "흑색자분", "백색페인트", "침투제", "세척제", "현상제", "형광침투제"]
        for i, cat in enumerate(ndt_cats):
            ttk.Label(ndt_frame, text=f"{cat}:").grid(row=i, column=0, padx=5, pady=2, sticky='w')
            ent = ttk.Entry(ndt_frame, width=10)
            m_key = "".join(cat.split())
            val = entry_data.get(f'NDT_{m_key}', 0)
            if pd.isna(val): val = 0
            ent.insert(0, str(val))
            ent.grid(row=i, column=1, padx=5, pady=2, sticky='w')
            fields[f'NDT_{m_key}'] = ent

            fields[f'NDT_{m_key}'] = ent
            
        # 4. NDT 상세 조건
        ndt_cond_frame = ttk.LabelFrame(scrollable_frame, text="NDT 상세 조건", padding=10)
        ndt_cond_frame.pack(fill='x', pady=5)
        
        cond_configs = [
            ('작업형태', '작업형태'), ('조건1', '조건1'), ('조건2', '조건2'),
            ('제경비율(%)', '제경비율'), ('기술료율(%)', '기술료율'), ('보정계수', '보정계수'),
            ('환산물량', '환산물량'), ('제경비', '제경비'), ('기술료', '기술료')
        ]
        for c_i, (c_lbl, c_key) in enumerate(cond_configs):
            c_row = c_i // 3
            c_col = (c_i % 3) * 2
            ttk.Label(ndt_cond_frame, text=f"{c_lbl}:").grid(row=c_row, column=c_col, padx=5, pady=2, sticky='w')
            c_ent = ttk.Entry(ndt_cond_frame, width=15)
            c_val = entry_data.get(c_key, '')
            if pd.isna(c_val): c_val = ''
            c_ent.insert(0, self.clean_nan(c_val))
            c_ent.grid(row=c_row, column=c_col+1, padx=5, pady=2, sticky='w')
            fields[c_key] = c_ent

        # 5. Vehicle Info
        vehicle_frame = ttk.LabelFrame(scrollable_frame, text="차량 점검 정보", padding=10)
        vehicle_frame.pack(fill='x', pady=5)
        
        v_configs = [
            ('차량번호', '차량번호'), ('주행거리', '주행거리'),
            ('점검내용', '차량점검'), ('차량비고', '차량비고')
        ]
        for v_i, (v_lbl, v_key) in enumerate(v_configs):
            v_row = v_i // 2
            v_col = (v_i % 2) * 2
            ttk.Label(vehicle_frame, text=f"{v_lbl}:").grid(row=v_row, column=v_col, padx=5, pady=5, sticky='w')
            v_ent = ttk.Entry(vehicle_frame, width=25)
            v_val = entry_data.get(v_key, '')
            v_ent.insert(0, self.clean_nan(v_val))
            v_ent.grid(row=v_row, column=v_col+1, padx=5, pady=5, sticky='w')
            fields[v_key] = v_ent

        # [NEW] 5.5 Remark (비고/메모)
        remark_frame = ttk.Frame(scrollable_frame, padding=10)
        remark_frame.pack(fill='x', pady=5)
        ttk.Label(remark_frame, text="비고(메모):").pack(side='left', padx=5)
        ent_remark = ttk.Entry(remark_frame, width=80)
        curr_rem = entry_data.get('Note', entry_data.get('비고', ''))
        if pd.isna(curr_rem): curr_rem = ''
        ent_remark.insert(0, self.clean_nan(curr_rem))
        ent_remark.pack(side='left', fill='x', expand=True, padx=5)
        fields['Note'] = ent_remark

        # 6. Workers
        worker_main_frame = ttk.LabelFrame(scrollable_frame, text="작업자 정보 (1~10)", padding=10)
        worker_main_frame.pack(fill='x', pady=5)
        


        # [NEW] Inner-dialog helper for Fee calculation
        def update_edit_fee_calc(event=None):
            try:
                def get_f(val):
                    try:
                        v = str(val).strip().replace(',', '')
                        return float(v) if v else 0.0
                    except: return 0.0

                qty = get_f(fields['검사량'].get())
                price = get_f(fields['단가'].get())
                travel = get_f(fields['출장비'].get())
                meal = get_f(fields['일식'].get())
                
                calc_fee = (qty * price) + travel + meal
                fields['검사비'].delete(0, tk.END)
                fields['검사비'].insert(0, f"{calc_fee:,.0f}")
            except: pass

        # Bindings for auto-calc
        fields['검사량'].bind('<KeyRelease>', lambda e: update_edit_fee_calc())
        
        # [NEW] Add comma auto-formatting for edit fields
        for k in ['단가', '출장비', '일식', '검사비']:
            if k in fields:
                fields[k].bind('<KeyRelease>', update_edit_fee_calc, add='+')
                fields[k].bind('<FocusOut>', lambda e, widget=fields[k]: self.format_entry_with_commas(e, widget), add='+')
                fields[k].bind('<Return>', lambda e, widget=fields[k]: self.format_entry_with_commas(e, widget), add='+')
        
        # Ensure FilmCount field also formats nicely (integer)
        if 'FilmCount' in fields:
            fields['FilmCount'].bind('<FocusOut>', lambda e, widget=fields['FilmCount']: self.format_entry_with_commas(e, widget), add='+')

        # Grid for workers: 5 rows x 2 cols
        worker_fields = {}
        for i in range(1, 11):
            w_row = (i-1) // 2
            w_col = (i-1) % 2
            
            w_frame = ttk.Frame(worker_main_frame, padding=2)
            w_frame.grid(row=w_row, column=w_col, padx=10, pady=5, sticky='nw')
            
            ttk.Label(w_frame, text=f"{i}:", width=2).pack(side='left')
            
            # Name
            cb_name = ttk.Combobox(w_frame, width=8, values=[''] + self.users)
            name_key = 'User' if i == 1 else f'User{i}'
            cb_name.set(self.clean_nan(entry_data.get(name_key, '')))
            cb_name.pack(side='left', padx=2)
            worker_fields[f'name{i}'] = cb_name
            
            # Time
            cb_time = ttk.Combobox(w_frame, width=16, values=[''] + self.worktimes)
            time_key = 'WorkTime' if i == 1 else f'WorkTime{i}'
            cb_time.set(self.clean_nan(entry_data.get(time_key, '')))
            cb_time.pack(side='left', padx=2)
            worker_fields[f'time{i}'] = cb_time
            
            # OT
            ent_ot = ttk.Entry(w_frame, width=18)
            ot_key = 'OT' if i == 1 else f'OT{i}'
            ent_ot.insert(0, self.clean_nan(entry_data.get(ot_key, '')))
            ent_ot.pack(side='left', padx=2)
            worker_fields[f'ot{i}'] = ent_ot
            
            # Auto-calculate OT for each worker when time changes
            def on_time_changed(event, idx=i, w_ent=cb_time, ot_ent=ent_ot, d_ent=ent_date):
                self.calculate_and_update_ot_manual(w_ent.get(), ot_ent, d_ent.get_date())
            
            cb_time.bind('<Return>', on_time_changed)
        
        # [NEW] Apply autocomplete to main edit fields
        self.enable_autocomplete(cb_site, values_list_attr='sites')
        self.enable_autocomplete(cb_mat, values_list=mat_options)
        self.enable_autocomplete(cb_equip, values_list_attr='equipment_suggestions')

        # 5. Buttons
        btn_frame = ttk.Frame(scrollable_frame, padding=20)
        btn_frame.pack(fill='x')
        
        def save_edits():
            # Collect data
            new_data = {}
            new_data['Date'] = ent_date.get_date().strftime('%Y-%m-%d')
            new_data['Site'] = cb_site.get().strip()
            new_data['업체명'] = cb_company.get().strip()
            new_data['적용코드'] = ent_code.get().strip()
            new_data['검사품명'] = ent_item.get().strip()
            
            full_mat_name = cb_mat.get().strip()
            # [FIX] Robust MaterialID mapping: Try to find matching ID using standard display name generator
            found_id = None
            if full_mat_name:
                for _, m_row in self.materials_df.iterrows():
                    m_id_check = m_row['MaterialID']
                    if self.get_material_display_name(m_id_check) == full_mat_name:
                        found_id = m_id_check
                        break
                
                # [FIX] Only auto-register if it's a CONSUMABLE. Equipment should NOT be registered.
                if not found_id:
                    if self._is_consumable_material(full_mat_name, cb_method.get().strip()):
                        found_id = full_mat_name # Resolve/register inside the guarded save transaction.
                    else:
                        found_id = full_mat_name
            
            new_data['Material'] = full_mat_name
            new_data['MaterialID'] = found_id if found_id is not None else entry_data.get('MaterialID')
            
            new_data['보조장비'] = '; '.join(names(aux_var.get()))
            new_data['장비명'] = cb_equip.get().strip()
            new_data['검사방법'] = cb_method.get().strip()
            new_data['Unit'] = cb_unit.get().strip() # [V13_FIX] Save Unit field
            
            # Numeric fields
            for key in ['검사량', '단가', '출장비', '일식', '검사비', 'FilmCount']:
                try: 
                    v_str = fields[key].get().strip().replace(',', '')
                    new_data[key] = float(v_str) if v_str else 0.0
                except: new_data[key] = 0.0
                
            for cat in rtk_cats:
                try: new_data[f'RTK_{cat}'] = float(fields[f'RTK_{cat}'].get())
                except: new_data[f'RTK_{cat}'] = 0.0
                
            for cat in ndt_cats:
                m_key = "".join(cat.split())
                try: new_data[f'NDT_{m_key}'] = float(fields[f'NDT_{m_key}'].get())
                except: new_data[f'NDT_{m_key}'] = 0.0
                
            # String fields
            for key in ['차량번호', '주행거리', '차량점검', '차량비고', '작업형태', '조건1', '조건2', 'Note']:
                if key in fields:
                    new_data[key] = fields[key].get().strip()
            
            # Additional numeric fields
            for key in ['제경비율', '기술료율', '보정계수', '환산물량', '제경비', '기술료']:
                if key in fields:
                    try:
                        v_str = fields[key].get().strip().replace(',', '')
                        new_data[key] = float(v_str) if v_str else 0.0
                    except: new_data[key] = 0.0
                
            for i in range(1, 11):
                name_key = 'User' if i == 1 else f'User{i}'
                time_key = 'WorkTime' if i == 1 else f'WorkTime{i}'
                ot_key = 'OT' if i == 1 else f'OT{i}'
                
                new_data[name_key] = worker_fields[f'name{i}'].get().strip()
                new_data[time_key] = worker_fields[f'time{i}'].get().strip()
                new_data[ot_key] = worker_fields[f'ot{i}'].get().strip()

            # Provide immediate feedback
            btn_save_edit.config(state='disabled', text="저장 중...")
            edit_win.update_idletasks()
            
            try:
                if self.save_daily_usage_edits(df_idx, new_data):
                    edit_win.destroy()
                    messagebox.showinfo("수정 완료", "기록이 성공적으로 수정되었으며 재고 내역이 업데이트되었습니다.")
                else:
                    btn_save_edit.config(state='normal', text="수정사항 저장")
            except Exception as e:
                btn_save_edit.config(state='normal', text="수정사항 저장")
                messagebox.showerror("오류", f"수정 중 오류가 발생했습니다: {e}")

        btn_save_edit = ttk.Button(btn_frame, text="수정사항 저장", style='Big.TButton', command=save_edits)
        btn_save_edit.pack(side='right', padx=10)
        ttk.Button(btn_frame, text="취소", command=edit_win.destroy).pack(side='right', padx=10)

