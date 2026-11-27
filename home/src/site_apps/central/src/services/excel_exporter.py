from site_apps.central.src.tagged_ndt_writer import write_all_tagged_sections
from openpyxl import Workbook
from tkcalendar import DateEntry
import pandas as pd
import os
import traceback
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from site_apps.central.src.data_sync import get_history_path
import openpyxl
from openpyxl.styles import Alignment, Border, Side, PatternFill, Font
from openpyxl.utils import get_column_letter
from site_apps.central.src.utils.helpers import normalize_id
from site_apps.central.src.worker_selection import worker_is_selected
from site_apps.central.src.vehicle_records import format_vehicle_numbers
import json
import re
import sys
import subprocess
import re
from site_apps.central.src.daily_work_report_manager import DailyWorkReportManager

def export_daily_work_report_impl(self):
    """작업일보를 엑셀 템플릿에 출력합니다."""
    try:
        template_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'resources', 'Template_DailyWorkReport.xlsx')
        if not os.path.exists(template_path):
            template_path = r'c:\Users\jjch2\Desktop\보고서Project PROVIDENCE\Request\PMI\Na-aba\home\resources\Template_DailyWorkReport.xlsx'
        
        if not os.path.exists(template_path):
            messagebox.showerror("오류", "작업일보 템플릿(Template_DailyWorkReport.xlsx)을 찾을 수 없습니다.")
            return

        def _clean_str(v, default=''):
            if v is None: return default
            s = str(v).strip()
            if s.lower() in ('nan', 'none', ''): return default
            return s

        date_val = self.ent_daily_date.get_date()
        site = self.cb_daily_site.get().strip()
        method = self.cb_daily_test_method.get().strip() # Default method from UI
        
        data = {
            'date': date_val,
            'company': _clean_str(self.cb_daily_company.get(), '원자력건설'),
            'project_name': site,
            'standard': _clean_str(self.ent_daily_applied_code.get(), 'KS'),
            'equipment': _clean_str(self.cb_daily_equip.get()),
            'report_no': _clean_str(self.ent_daily_report_no.get()), 
            'inspection_item': _clean_str(self.ent_daily_inspection_item.get()), 
            'inspector': '', 
            'car_no': '', 
            'methods': {},
            'inspection_details': [],
            'rtk': {},
            'ot_status': [],
            'materials': {}
        }

        # data['company']는 시공사(cb_daily_company)를 유지하며 NDT 자재 제조사(_company)로 덮어쓰지 않음

        # DB 데이터 집합 (Option 1: 전체 합산 버전)
        site_records = pd.DataFrame()
        if not self.daily_usage_df.empty:
            try:
                df_copy = self.daily_usage_df.copy()
                
                # [SMART SELECTION] 리스트에서 선택했는지 확인
                selection = self.daily_usage_tree.selection()
                if len(selection) >= 1:
                    # 선택된 항목들 모아서 출력
                    selected_indices = []
                    for item in selection:
                        tags = self.daily_usage_tree.item(item, 'tags')
                        if tags and tags[0].isdigit():
                            idx = int(tags[0])
                            if idx in df_copy.index:
                                selected_indices.append(idx)
                    if selected_indices:
                        site_records = df_copy.loc[selected_indices]
                        # 현장명/날짜는 첫 번째 선택된 항목 기준으로 업데이트 (데이터 정합성 보장)
                        if 'Site' in site_records.columns: site = str(site_records.iloc[0]['Site']).strip()
                        if 'Date' in site_records.columns: date_val = pd.to_datetime(site_records.iloc[0]['Date']).date()
                    
                if site_records.empty:
                    # 선택이 없는 경우: 현재 폼의 날짜/현장 기준으로 전체 합산
                    # [ROBUST MATCHING] 날짜와 현장명 비교 시 공백 등 무시하여 정확도 향상
                    check_date = pd.to_datetime(date_val).date()
                    site_col = 'Site' if 'Site' in df_copy.columns else '현장' if '현장' in df_copy.columns else ''
                    date_col = 'Date' if 'Date' in df_copy.columns else '날짜' if '날짜' in df_copy.columns else ''
                    
                    if site_col and date_col:
                        df_copy['Date_norm'] = pd.to_datetime(df_copy[date_col], errors='coerce').dt.date
                        site_records = df_copy[
                            (df_copy['Date_norm'] == check_date) & 
                            (df_copy[site_col].astype(str).str.strip().str.upper() == str(site).strip().upper())
                        ]
                
                # [NEW] site_records에서 장비명, 업체명 및 검사품명 집계 업데이트
                if not site_records.empty:
                    if '장비명' in site_records.columns:
                        from site_apps.central.src.aux_equipment import report_names
                        equips = report_names(site_records)
                        if equips: data['equipment'] = ", ".join(equips)
                    
                    comp_col = '업체명' if '업체명' in site_records.columns else ''
                    if comp_col:
                        comps = [_clean_str(x) for x in site_records[comp_col].dropna().unique() if _clean_str(x)]
                        if comps: data['company'] = comps[0] # 첫 번째 유효한 업체명 사용
                    
                    if '검사품명' in site_records.columns:
                        items = [_clean_str(x) for x in site_records['검사품명'].dropna().unique() if _clean_str(x)]
                        if items and not data['inspection_item']:
                            data['inspection_item'] = items[0]
            except Exception as e:
                print(f"[DEBUG] Site Records Collection Error: {e}")
                pass

        all_vehicles = []
        if hasattr(self, 'vehicle_boxes'):
            for box in self.vehicle_boxes:
                v = box.cb_vehicle_info.get().strip()
                if v and v not in all_vehicles: all_vehicles.append(v)
        if not site_records.empty and '차량번호' in site_records.columns:
            for v in site_records['차량번호'].dropna().unique():
                v_str = str(v).strip()
                if v_str and v_str not in all_vehicles: all_vehicles.append(v_str)
        data['car_no'] = format_vehicle_numbers(all_vehicles)

        note_texts = []
        if hasattr(self, 'main_memo_text'):
            try:
                ui_memo = self.main_memo_text.get('1.0', tk.END).strip()
                if ui_memo:
                    note_texts.append(ui_memo)
            except: pass

        # Also collect ent_daily_note if available and not redundant
        if hasattr(self, 'ent_daily_note'):
            try:
                ui_note = self.ent_daily_note.get().strip()
                if ui_note and ui_note not in note_texts:
                    note_texts.append(ui_note)
            except: pass

        if not note_texts and not site_records.empty:
            for n_col in ('Note', '비고'):
                if n_col in site_records.columns:
                    for val in site_records[n_col].dropna().unique():
                        v_clean = str(val).strip()
                        if v_clean and v_clean.lower() != 'nan' and v_clean not in note_texts:
                            note_texts.append(v_clean)

        data['note'] = "\n".join(note_texts)

        # [NEW] 공사 수행현황 (Section 1) - DB site_records에 있는 모든 방식 집계
        method_col = '검사방법' if '검사방법' in site_records.columns else 'TestMethod' if 'TestMethod' in site_records.columns else ''
        
        if not site_records.empty and method_col:
            from site_apps.central.src.services.report_inspection import collect_method_totals
            data['methods'] = collect_method_totals(site_records)
        else:
            # Fallback to UI values if DB is empty or column missing
            method = self.cb_daily_test_method.get().strip()
            unit_val = self.cb_daily_unit.get().strip() 
            qty_val = self.ent_daily_test_amount.get().strip()
            price_val = self.ent_daily_unit_price.get().strip()
            travel_val = self.ent_daily_travel_cost.get().strip()
            total_val = self.ent_daily_test_fee.get().strip()
            
            if method:
                from site_apps.central.src.services.report_inspection import inspection_method_shift
                method, shift = inspection_method_shift({'검사방법': method,
                    '작업형태': self.ndt_work_time_var.get() if hasattr(self, 'ndt_work_time_var') else ''})
                data['methods'][f'{method}({shift})'] = {
                    'unit': unit_val,
                    'qty': float(qty_val.replace(',', '')) if qty_val else 0,
                    'price': float(price_val.replace(',', '')) if price_val else 0,
                    'travel': float(travel_val.replace(',', '')) if travel_val else 0,
                    'total': float(total_val.replace(',', '')) if total_val else 0
                }

        # Original RT points and reject films; REP is a separate record.
        from site_apps.central.src.services.report_inspection import collect_inspection_details, collect_pad_details, inspection_method_shift
        data['inspection_details'] = collect_inspection_details(site_records, getattr(self, 'SIZE_LENGTH', {}))

        # [NEW] Collect PAD breakdown from work log or last sync
        pad_details = collect_pad_details(date_val, method)
        if not pad_details and hasattr(self, '_last_pad_synced_breakdown') and self._last_pad_synced_breakdown:
            pad_details = self._last_pad_synced_breakdown
        if pad_details:
            data['pad_details'] = pad_details

        # [FALLBACK] site_records가 없거나 inspection_details가 비어있는 경우 화면 입력값 기반 복원
        if not data.get('inspection_details') and hasattr(self, 'cb_daily_test_method'):
            u_method = self.cb_daily_test_method.get().strip().upper()
            u_pipe = getattr(self, 'ndt_report_pipe_var', tk.StringVar(value="")).get().strip().upper()
            if u_pipe:
                u_pts = getattr(self, 'ndt_ori_joint_var', tk.StringVar(value="")).get().strip()
                try: pts_val = float(u_pts.replace(',', '')) if u_pts else 0.0
                except: pts_val = 0.0
                u_qty = getattr(self, 'ent_daily_test_amount', None)
                try: qty_val = float(u_qty.get().replace(',', '')) if u_qty and u_qty.get() else 0.0
                except: qty_val = 0.0
                if pts_val == 0 and qty_val > 0 and u_method in ['MT', 'PT', 'PAUT']:
                    size_lookup = getattr(self, 'SIZE_LENGTH', {})
                    p_key = u_pipe if u_pipe.endswith('A') else f"{u_pipe}A"
                    unit_len = size_lookup.get(p_key, 0)
                    if unit_len > 0:
                        calc_pts = qty_val / unit_len
                        if abs(calc_pts - round(calc_pts)) < 0.05:
                            pts_val = round(calc_pts)
                data['inspection_details'].append({
                    'method': u_method,
                    'shift': inspection_method_shift({'검사방법': u_method,
                        '작업형태': self.ndt_work_time_var.get() if hasattr(self, 'ndt_work_time_var') else ''})[1],
                    'pipe_size': u_pipe,
                    'points': int(pts_val) if pts_val.is_integer() else pts_val,
                    'quantity': qty_val,
                    'reject_films': float(self.ndt_rej_joint_var.get().replace(',', '') or 0) if u_method == 'RT' else 0,
                })

        # 작업자 및 O/T (간소화 버전)
        # [NEW] 작업자 / OT 정보 DB에서 집계 (현장 탭 기록 기준)
        def _clean_name(n):
            if not n or str(n).lower() in ('nan', 'none', ''): return ''
            text = str(n).strip()
            titles = ['부장', '차장', '과장', '대리', '주임', '기사', '선임', '수석', '책임',
                      '팀장', '이사', '본부장', '실장', '소장', '직장', '반장', '팀원', '계장']
            for t in titles:
                import re as _re
                text = _re.sub(r'[\s/(\[]*' + t + r'[\s)\]]*$', '', text)
                text = _re.sub(r'^[\s/(\[]*' + t + r'[\s)\]]*', '', text)
            return text.strip()

        inspectors = []
        ot_groups = {}  # key: (work_time, ot_amount) -> {names:[], company:''}

        if not site_records.empty:
            company_val = data.get('company', '')
            for _, row in site_records.iterrows():
                for i in range(1, 11):
                    u_key   = 'User'     if i == 1 else f'User{i}'
                    wt_key  = 'WorkTime' if i == 1 else f'WorkTime{i}'
                    ot_key  = 'OT'       if i == 1 else f'OT{i}'

                    name = str(row.get(u_key, '')).strip()
                    if not name or name == 'nan':
                        if i == 1:
                            for alt_k in ['검사자', '작업자', '(Full작업자)']:
                                alt_v = str(row.get(alt_k, '')).strip()
                                if alt_v and alt_v != 'nan':
                                    name = alt_v
                                    break
                    if not name or name == 'nan': continue

                    if name not in inspectors:
                        inspectors.append(name)

                    wt  = str(row.get(wt_key, '')).strip()
                    ot_raw = str(row.get(ot_key,  '')).strip()
                    
                    # [FIX] Handle decimals correctly to prevent 26000.0 becoming 260000
                    try:
                        # Try parsing as float first to handle .0
                        oa_num = float(ot_raw.replace(',', '')) if ot_raw and ot_raw != 'nan' else 0
                        oa = str(int(oa_num)) if oa_num > 0 else ''
                    except:
                        # Fallback to digit-only if float parsing fails
                        oa = ''.join(c for c in ot_raw if c.isdigit())

                    if wt == 'nan': wt = ''
                    if not wt and not oa: continue

                    key = (wt, oa)
                    if key not in ot_groups:
                        ot_groups[key] = {'names': [], 'company': company_val}
                    if name not in ot_groups[key]['names']:
                        ot_groups[key]['names'].append(name)

        # [FALLBACK] site_records에서 작업자를 찾지 못한 경우 UI의 worker_group에서 수집
        if not inspectors:
            for i in range(1, 11):
                group = getattr(self, f'worker_group{i}', None)
                if group:
                    if not worker_is_selected(group): continue
                    w_name = group.get_worker().strip() if hasattr(group, 'get_worker') else group.cb_name.get().strip()
                    if w_name and w_name not in inspectors:
                        inspectors.append(w_name)

        # Inspector display (titles always stripped for header)
        disp_insp = [_clean_name(n) for n in inspectors if _clean_name(n)]
        if len(disp_insp) > 3:
            data['inspector'] = ', '.join(disp_insp[:3]) + f" 외 {len(disp_insp) - 3}명"
        elif disp_insp:
            data['inspector'] = ', '.join(disp_insp)
        elif hasattr(self, 'users') and self.users:
            data['inspector'] = _clean_name(self.users[0])

        # Build ot_status list
        data['ot_status'] = []
        if ot_groups:
            for (wt, oa), grp in ot_groups.items():
                names = grp['names']
                name_disp = ', '.join([_clean_name(n) for n in names])
                wt_disp = wt
                
                # Match method for this group if possible
                curr_method = method
                if not site_records.empty and method_col:
                    for _, r_match in site_records.iterrows():
                        if any(str(r_match.get(f'User{k}' if k>1 else 'User', '')).strip() in names for k in range(1, 11)):
                            m_found = str(r_match.get(method_col, '')).strip()
                            if m_found and m_found.lower() != 'nan':
                                curr_method = m_found; break

                data['ot_status'].append({
                    'names':      name_disp,
                    'ot_hours':   wt_disp,
                    'ot_amount':  f'{int(oa):,}' if oa else '',
                    'company':    grp['company'],
                    'method':     curr_method,
                })
        else:
            # Fallback to UI worker groups if DB is empty
            for i in range(1, 11):
                group = getattr(self, f'worker_group{i}', None)
                if group:
                    if not worker_is_selected(group): continue
                    w_name = group.get_worker().strip() if hasattr(group, 'get_worker') else group.cb_name.get().strip()
                    w_wt = group.get_time().strip() if hasattr(group, 'get_time') else group.ent_worktime.get().strip()
                    w_ot = group.get_ot().strip() if hasattr(group, 'get_ot') else group.ent_ot.get().strip()
                    if w_name:
                        data['ot_status'].append({
                            'names':      w_name,
                            'ot_hours':   w_wt,
                            'ot_amount':  w_ot,
                            'company':    data.get('company', ''),
                            'method':     method,
                        })
                        if not data['inspector']:
                            data['inspector'] = _clean_name(w_name)

        

        # 2.5 차량 및 안전 점검 수집 (섹션 3)

        data['vehicles'] = []
        seen_vehicles = set()
        if not site_records.empty and '차량번호' in site_records.columns:
            for _, row in site_records.iterrows():
                v_no = str(row.get('차량번호', '')).strip()
                if not v_no or v_no == 'nan': continue
                
                if v_no in seen_vehicles: continue
                seen_vehicles.add(v_no)
                
                v_insp_raw = str(row.get('차량점검', '')).strip()
                v_mileage = str(row.get('주행거리', '')).strip()
                v_remarks = str(row.get('차량비고', '')).strip()
                
                v_parsed = {
                    'vehicle_info': v_no, 
                    'mileage': v_mileage if v_mileage != 'nan' else '', 
                    'remarks': v_remarks if v_remarks != 'nan' else ''
                }
                
                if ':' in v_insp_raw:
                    for pair in v_insp_raw.split('|'):
                        if ':' in pair:
                            k, v = pair.split(':', 1)
                            v_parsed[k] = v
                elif ',' in v_insp_raw or v_insp_raw:
                    for k in v_insp_raw.split(','):
                        k_clean = k.strip()
                        if k_clean:
                            if 'locking' in k_clean: v_parsed[k_clean] = '잠금'
                            elif 'cleaning' in k_clean: v_parsed[k_clean] = '함'
                            else: v_parsed[k_clean] = '양호'
                data['vehicles'].append(v_parsed)
                
        # 폼에 아직 저장되지 않은 내용이 있을 경우 대비
        if not data['vehicles']:
            if hasattr(self, 'vehicle_boxes') and self.vehicle_boxes:
                for box in self.vehicle_boxes:
                    data['vehicles'].append(box.get_data())
            elif hasattr(self, 'vehicle_widget'):
                data['vehicles'].append(self.vehicle_widget.get_data())
                
        # [NEW] N9 셀(car_no)에 들어갈 차량번호만 추출
        car_no_list = []
        for v in data['vehicles']:
            v_no = v.get('vehicle_info', '').strip()
            if v_no:
                car_no_list.append(v_no)
        if car_no_list:
            data['car_no'] = format_vehicle_numbers(car_no_list)

        # 자재 정보 수집 (NDT 섹션)
        data['selected_material'] = self.cb_daily_material.get().strip()

        # [NEW] RT 품목별 수량 - DB site_records의 MaterialID로 그룹화 + 품목명/규격 조회
        # D열: 품목명(Name), F열: 규격(Spec), M열: 사용수량(Usage)
        if not site_records.empty and 'MaterialID' in site_records.columns:
            # MaterialID 기준으로 Usage 합산 및 Name/Spec 조회
            mat_groups = {}  # key: MaterialID_upper -> {'qty': float, 'name': str, 'spec': str, 'original_id': str}
            for _, row in site_records.iterrows():
                mat_id = str(row.get('MaterialID', '')).strip()
                if not mat_id or mat_id == 'nan':
                    continue
                qty = 0
                try: qty = float(row.get('Usage', 0) or 0)
                except: pass

                mat_id_upper = mat_id.upper()
                if mat_id_upper not in mat_groups:
                    # MaterialID를 사용하여 materials_df에서 실제 품목명과 규격 조회
                    mat_name_val = mat_id
                    mat_spec_val = ''
                    disp_name = ''
                    mat_cat_val = ''
                    
                    if hasattr(self, 'materials_df') and not self.materials_df.empty:
                        try:
                            # 숫자형 ID 대응 (ID가 405.0 등일 수 있음)
                            clean_id = str(mat_id).strip().replace('.0', '')
                            match = self.materials_df[
                                self.materials_df['MaterialID'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True) == clean_id
                            ]
                            if not match.empty:
                                # DB에서 값을 가져오되, 품목명 안에 규격이 합쳐져 있을 경우를 대비해 다시 분리
                                raw_name = str(match.iloc[0].get('품목명', mat_name_val)).strip()
                                raw_spec = str(match.iloc[0].get('규격', '')).strip()
                                
                                # [USER REQUEST] Keep 'Carestream' prefix
                                # (Removal logic removed)
                                
                                if '-' in raw_name and (not raw_spec or raw_spec == 'nan' or raw_spec == '' or raw_spec == '자동등록'):
                                    # 품목명에 하이픈이 있고 규격 칸이 비어있거나 무의미한 경우 분리 실행
                                    dash_idx = raw_name.index('-')
                                    mat_name_val = raw_name[:dash_idx].strip()
                                    mat_spec_val = raw_name[dash_idx+1:].strip()
                                else:
                                    mat_name_val = raw_name
                                    mat_spec_val = raw_spec if (raw_spec and str(raw_spec).lower() != 'nan' and raw_spec != '자동등록') else ''
                                
                                # [NEW] Get Category (Classification)
                                mat_cat_val = str(match.iloc[0].get('품목군코드', '')).strip()
                                if not mat_cat_val: mat_cat_val = str(match.iloc[0].iloc[8]) if len(match.columns) > 8 else ''
                        except: pass
                    
                    # [FIX] If we still only have an ID (like '410') and it contains a dash, try splitting it
                    if mat_name_val == mat_id and '-' in mat_id:
                        dash_idx = mat_id.index('-')
                        mat_name_val = mat_id[:dash_idx].strip()
                        mat_spec_val = mat_id[dash_idx+1:].strip()

                    disp_name = f"{mat_name_val}-{mat_spec_val}".upper() if mat_spec_val else mat_name_val.upper()
                    
                    mat_groups[mat_id_upper] = {
                        'qty': 0, 'name': mat_name_val, 'spec': mat_spec_val, 'original_id': mat_id, 'disp': disp_name,
                        'category': mat_cat_val
                    }
                mat_groups[mat_id_upper]['qty'] += qty

            # [REFINEMENT] Secondary Grouping by Display Name to merge identical items
            merged_mats = {} # key: disp_name -> grp_data
            for _, grp in mat_groups.items():
                d_name = grp.get('disp', 'Unknown')
                if d_name not in merged_mats:
                    merged_mats[d_name] = grp.copy()
                else:
                    merged_mats[d_name]['qty'] += grp['qty']

            # Build data['materials'] using merged results
            rt_keys = ['RT T200', 'RT AA400', 'RT Other']
            rt_counter = 0
            
            # [DEBUG] Log all merged materials before filtering
            print(f"DEBUG: All merged materials before RT filtering: {list(merged_mats.keys())}")
            
            # Separate RT items for sequential keying
            # [FIX] Include MT/PT keywords to prevent chemicals from being categorized as RT
            chem_names = [
                'WHITE', 'BLACK', '7C', 'PENETRANT', 'CLEANER', 'DEVELOPER',
                '백색', '흑색', '자분', '침투', '세척', '현상', '페인트',
                'MT', 'PT'
            ]
            rt_merged = {}
            for d_name, grp in merged_mats.items():
                is_chem = any(c in d_name.upper() for c in chem_names)
                # [IMPROVED] Expanded RT detection to catch films by brand and model
                rt_keywords = ['RT ', 'RT-', 'FILM', 'CARESTREAM', 'AGFA', 'FUJI', 'KODAK', 'T200', 'AA400', 'MX125', 'M100']
                cat_upper = str(grp.get('category', '')).upper()
                is_rt_cat = any(k in cat_upper for k in ['FILM', 'RT'])
                is_rt_name = any(k in d_name.upper() for k in rt_keywords)
                
                if (is_rt_cat or is_rt_name) and not is_chem:
                    rt_merged[d_name] = grp

            for d_name, grp in rt_merged.items():
                print(f"DEBUG: Processing RT Item {rt_counter+1}: {d_name}")
                if rt_counter < 3:
                    mat_key = rt_keys[rt_counter]
                else:
                    mat_key = f"RT_ROW_{rt_counter + 1}"
                
                rt_counter += 1
                data['materials'][mat_key] = {
                    'used': float(grp['qty']), 
                    'name': grp['name'], 
                    'spec': grp['spec'], 
                    'is_rt': True, 
                    'category': 'RT'
                }
            
            # [NOTE] Non-RT (Chemicals) are handled separately below

        else:
            # DB 데이터 없으면 UI에서 읽기 (fallback)
            if hasattr(self, 'ndt_company_entries') and self.ndt_company_entries:
                mats = self.ndt_company_entries[0]
                for m_key, mat_name in [('RT T200', 'T200'), ('RT AA400', 'AA400')]:
                    if m_key in mats:
                        val = mats[m_key].get().strip()
                        try: used = int(val) if val else 0
                        except: used = 0
                        data['materials'][m_key] = {'used': used}

        from site_apps.central.src.services.report_materials import collect_chemical_usage
        data['materials'].update(collect_chemical_usage(
            site_records, getattr(self, 'ndt_company_entries', [])))

        # 2.5 OT Status Gathering
        data['ot_status'] = []
        
        def clean_val(v):
            if pd.isna(v) or str(v).lower() == 'nan': return ""
            val = str(v).strip()
            # Strip common titles from both END and START
            titles = ["부장", "차장", "과장", "대리", "주임", "계장", "사원", "반장", "기사"]
            for title in titles:
                if val.endswith(title): val = val[:-len(title)].strip()
                if val.startswith(title): val = val[len(title):].strip()
            return val

        # 1) Priority: Check DB (Site Records) first for "Saved" data
        if not site_records.empty:
            print(f"DEBUG: --- ALL DB COLUMNS: {site_records.columns.tolist()} ---")
            for idx, row in site_records.iterrows():
                # [FINAL FIX] Scan all 10 user columns to ensure no one is missed (e.g. User2, User3...)
                for i in range(1, 11):
                    u_key = 'User' if i == 1 else f'User{i}'
                    wt_key = 'WorkTime' if i == 1 else f'WorkTime{i}'
                    oa_key = 'OT' if i == 1 else f'OT{i}'
                    
                    worker_name = clean_val(row.get(u_key, row.get('작업자' if i==1 else f'작업자{i}', '')))
                    if not worker_name or worker_name == "": continue
                    has_job_labor = bool(clean_val(row.get('작업별인원', '')))
                    
                    ot_val = clean_val(row.get(wt_key, ''))
                    if not ot_val and not has_job_labor:
                        ot_val = clean_val(row.get('작업시간' if i==1 else f'작업시간{i}', ''))
                    
                    # Fallback for WorkTime: If specific UserX's WorkTime is missing, try general WorkTime column
                    if not ot_val and not has_job_labor:
                        ot_val = clean_val(row.get('WorkTime', row.get('작업시간', '')))
                    
                    if not ot_val and not has_job_labor:
                        # Scan all columns for time-like strings if explicit names fail
                        for col_name in row.index:
                            potential_val = str(row[col_name])
                            if any(k in potential_val for k in ["09:00", "24:00", "(주야간)", "~"]):
                                ot_val = potential_val.strip()
                                break
                    
                    if not ot_val and not has_job_labor:
                        ot_val = clean_val(row.get('OT시간' if i==1 else f'OT시간{i}', ''))
                        
                    # For amount, 'OT' seems to have it in this DB
                    ot_amount = clean_val(row.get(oa_key, row.get('OT금액' if i==1 else f'OT금액{i}', '')))
                    if not ot_amount and i > 1 and not has_job_labor:
                        # Fallback to general OT column if specific one is empty
                        ot_amount = clean_val(row.get('OT', row.get('OT금액', '')))
                    
                    data['ot_status'].append({
                        'names': worker_name,
                        'company': clean_val(row.get('업체명', '')),
                        'method': clean_val(row.get('검사방법', '')),
                        'ot_hours': ot_val,
                        'ot_amount': ot_amount
                    })
            print(f"DEBUG: Gathered {len(data['ot_status'])} OT items from DB")

        # 2) Fallback: If DB is empty, check UI entries
        if not data['ot_status']:
            for i in range(1, 11):
                group = getattr(self, f'worker_group{i}', None)
                if group:
                    try:
                        if not worker_is_selected(group): continue
                        name = group.get_worker().strip()
                        if name and name.lower() != 'nan':
                            wt_val = group.get_time().strip() if hasattr(group, 'get_time') else (
                                group.ent_worktime.get().strip() if hasattr(group, 'ent_worktime') else ''
                            )
                            ot_amt_val = group.get_ot().strip() if hasattr(group, 'get_ot') else (
                                group.ent_ot.get().strip() if hasattr(group, 'ent_ot') else ''
                            )
                            data['ot_status'].append({
                                'names': name,
                                'company': self.cb_daily_company.get().strip(),
                                'method': self.cb_daily_test_method.get().strip(),
                                'ot_hours': wt_val,
                                'ot_amount': ot_amt_val
                            })
                    except: pass
            if data['ot_status']:
                print(f"DEBUG: Gathered {len(data['ot_status'])} OT items from UI (Fallback)")
        
        print(f"DEBUG: FINAL OT Status for report: {len(data['ot_status'])} items")

        data['rtk'] = {}
        rtk_cats = {
            '센터미스': 'center_miss', '농도': 'density', '마킹미스': 'marking_miss',
            '필름마크': 'film_mark', '취급부주의': 'handling', '고객불만': 'customer_complaint', '기타': 'etc'
        }
        rtk_total = 0
        if not site_records.empty:
            # Get clean column names for matching
            import re as _re
            def clean_name(s): return _re.sub(r'[^A-Z가-힣0-9]', '', str(s).upper())
            
            actual_cols = site_records.columns.tolist()
            clean_cols = [clean_name(c) for c in actual_cols]
            
            for kor_key, eng_key in rtk_cats.items():
                db_val = 0
                target_clean = clean_name(kor_key)
                rtk_target_clean = clean_name(f"RTK_{kor_key}")
                
                found_col = None
                for i, c_clean in enumerate(clean_cols):
                    if c_clean == target_clean or c_clean == rtk_target_clean:
                        found_col = actual_cols[i]
                        break
                
                if found_col:
                    try:
                        # Sum up and handle potential string/NaN values
                        series = pd.to_numeric(site_records[found_col], errors='coerce').fillna(0)
                        db_val = int(series.sum())
                        print(f"DEBUG: Found RTK Col '{found_col}' for '{kor_key}', sum={db_val}")
                    except Exception as e:
                        print(f"DEBUG: Error summing RTK {kor_key}: {e}")
                
                data['rtk'][kor_key] = db_val
                rtk_total += db_val
        
        # [CRITICAL FIX] If DB result is 0 or records empty, fallback to UI entries
        if rtk_total == 0:
            for kor_key, eng_key in rtk_cats.items():
                val = 0
                # Try direct widget attribute first
                widget = getattr(self, f"ent_rtk_{eng_key}", None)
                if widget:
                    try: val = int(widget.get().strip() or 0)
                    except: pass
                # Then try the rtk_entries dictionary as backup
                elif hasattr(self, 'rtk_entries') and kor_key in self.rtk_entries:
                    try: val = int(self.rtk_entries[kor_key].get().strip() or 0)
                    except: pass
                
                data['rtk'][kor_key] = val
                rtk_total += val
                
        data['rtk_total'] = rtk_total

        default_filename = f"한국지역난방 중앙지사_{date_val.strftime('%Y%m%d')}.xlsx"
        save_path = filedialog.asksaveasfilename(
            title="작업일보 저장",
            initialfile=default_filename,
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )

        if not save_path:
            return

        # 4. 리포트 생성
        mapping = self.load_report_mapping()
        manager = DailyWorkReportManager(template_path)
        manager.generate_report(data, save_path, custom_mapping=mapping)

        messagebox.showinfo("성공", f"작업일보가 생성되었습니다:\n{save_path}")
        
        # 생성된 파일 열기
        if messagebox.askyesno("확인", "생성된 작업일보를 지금 여시겠습니까?"):
            os.startfile(save_path)

    except PermissionError:
        messagebox.showerror("오류", f"파일 접근 오류: '{os.path.basename(save_path)}' 파일이 다른 프로그램(엑셀 등)에서 열려 있어 저장할 수 없습니다.\n파일을 닫고 다시 시도해 주세요.")
    except Exception as e:
        traceback.print_exc()
        messagebox.showerror("오류", f"작업일보 생성 중 오류 발생: {e}")


def export_materials_impl(self):
    save_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        initialfile="Materials_Export.xlsx",
        title="자재 목록 저장",
        filetypes=[("Excel files", "*.xlsx")]
    )
    
    if save_path:
        try:
            self.materials_df.to_excel(save_path, index=False)
            messagebox.showinfo("완료", "자재 목록이 저장되었습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"저장 실패: {e}")


def save_df_to_excel_autofit_impl(self, df, save_path, sheet_name='Sheet1'):
    """Save a DataFrame to Excel with automatic column width adjustment (AutoFit)"""
    with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        worksheet = writer.sheets[sheet_name]
        
        for idx, col in enumerate(df.columns):
            # Calculate max length of values in column + header
            # We handle Korean characters by assuming they take ~2 units of width
            def get_display_width(s):
                width = 0
                for char in str(s):
                    if ord(char) > 127: # Non-ASCII (Korean, etc.)
                        width += 2
                    else:
                        width += 1
                return width

            series = df[col].astype(str)
            # Filter out empty strings/NAs for max calc
            lengths = series.apply(get_display_width)
            max_val_len = lengths.max() if not lengths.empty else 0
            header_len = get_display_width(col)
            
            # Final width with padding - Cap at a reasonable width to force wrapping
            final_width = min(max(max_val_len, header_len) + 2, 15)
            
            # Map index to column letter
            # column_letter property is available in openpyxl cells
            col_letter = worksheet.cell(row=1, column=idx+1).column_letter
            worksheet.column_dimensions[col_letter].width = final_width

        # --- Page Setup for Printing ---
        # Set to Portrait orientation (User request)
        worksheet.page_setup.orientation = worksheet.ORIENTATION_PORTRAIT
        
        # Enable Fit to Page width (Scale everything to fit horizontally)
        worksheet.sheet_properties.pageSetUpPr.fitToPage = True
        worksheet.page_setup.fitToWidth = 1
        worksheet.page_setup.fitToHeight = 0 # Automatic heights (multiple pages if long)
        
        # Enable Wrap Text for all cells to use vertical space instead of horizontal width
        from openpyxl.styles import Alignment
        wrap_alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
        for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
            for cell in row:
                cell.alignment = wrap_alignment

        # Optional: Set small margins to maximize space (units are inches)
        worksheet.page_margins.left = 0.25
        worksheet.page_margins.right = 0.25
        worksheet.page_margins.top = 0.5
        worksheet.page_margins.bottom = 0.5





def export_central_daily_work_report_impl(self):
    """작업일보를 중앙지사 전용 v20 양식에 맞추어 엑셀로 출력합니다."""
    try:
        import pandas as pd
        import openpyxl
        from datetime import datetime
        import tkinter as tk
        from tkinter import messagebox, filedialog

        template_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'resources', 'Template_Central_DailyWorkReport.xlsx')
        if not os.path.exists(template_path):
            messagebox.showerror("오류", "중앙지사 전용 작업일보 템플릿(Template_Central_DailyWorkReport.xlsx)을 찾을 수 없습니다.")
            return

        def _clean_str(v, default=''):
            if v is None: return default
            s = str(v).strip()
            if s.lower() in ('nan', 'none', ''): return default
            return s
            
        def _clean_name(n):
            if not n or str(n).lower() in ('nan', 'none', ''): return ''
            text = str(n).strip()
            titles = ['부장', '차장', '과장', '대리', '주임', '기사', '선임', '수석', '책임',
                      '팀장', '이사', '본부장', '실장', '소장', '직장', '반장', '팀원', '계장']
            for t in titles:
                import re as _re
                text = _re.sub(r'[\s/(\[]*' + t + r'[\s)\]]*$', '', text)
                text = _re.sub(r'^[\s/(\[]*' + t + r'[\s)\]]*', '', text)
            return text.strip()

        date_val = self.ent_daily_date.get_date()
        site = self.cb_daily_site.get().strip()

        # DB 데이터 집합 
        df_copy = pd.DataFrame()
        if not self.daily_usage_df.empty:
            df_copy = self.daily_usage_df.copy()
            site_col = 'Site' if 'Site' in df_copy.columns else '현장' if '현장' in df_copy.columns else ''
            date_col = 'Date' if 'Date' in df_copy.columns else '날짜' if '날짜' in df_copy.columns else ''

            # 선택 행은 출력 기준(날짜·현장)만 정하고, 실제 출력은 같은 날·같은
            # 현장의 모든 기록을 합산한다. 한 행만 선택해도 복수 관경 작업이 누락되지 않는다.
            selection = self.daily_usage_tree.selection()
            if selection:
                tags = self.daily_usage_tree.item(selection[0], 'tags')
                if tags:
                    try:
                        selected_idx = int(tags[0])
                        if selected_idx in df_copy.index:
                            selected_row = df_copy.loc[selected_idx]
                            if site_col:
                                selected_site = str(selected_row.get(site_col, '')).strip()
                                if selected_site:
                                    site = selected_site
                            if date_col:
                                selected_date = pd.to_datetime(selected_row.get(date_col), errors='coerce')
                                if pd.notna(selected_date):
                                    date_val = selected_date.date()
                    except (TypeError, ValueError):
                        pass
            
            if site_col and date_col:
                df_copy['Date_norm'] = pd.to_datetime(df_copy[date_col], errors='coerce').dt.date
                check_date = pd.to_datetime(date_val).date()
                
                # 금일 기록
                today_records = df_copy[
                    (df_copy['Date_norm'] == check_date) & 
                    (df_copy[site_col].astype(str).str.strip().str.upper() == str(site).strip().upper())
                ]
                
                # 전일 누계 기록 (과거 기록 모두 합산)
                past_records = df_copy[
                    (df_copy['Date_norm'] < check_date) & 
                    (df_copy[site_col].astype(str).str.strip().str.upper() == str(site).strip().upper())
                ]
            else:
                today_records = pd.DataFrame()
                past_records = pd.DataFrame()
        else:
            today_records = pd.DataFrame()
            past_records = pd.DataFrame()

        # 물량 집계 로직
        def calc_method_qty(records, method_key, base_method):
            if records.empty: return 0
            method_col = '검사방법' if '검사방법' in records.columns else 'TestMethod' if 'TestMethod' in records.columns else ''
            if not method_col: return 0
            
            qty = 0
            for _, row in records.iterrows():
                m = str(row.get(method_col, '')).upper().strip()
                if base_method in m:
                    size_info = ' '.join(str(row.get(col, '')).upper() for col in (
                        '관경(Inch)', '관경', 'MaterialID', '품목명', '검사품명'
                    ))
                    # 야간 확인
                    work_type = ' '.join(str(row.get(col, '')).upper() for col in (
                        '작업형태', '근무구분', '구분', 'WorkType'
                    ))
                    is_night = '야간' in m or 'NIGHT' in m or '야간' in work_type or 'NIGHT' in work_type
                    target_is_night = '_N' in method_key
                    
                    # 관경 확인 (PAUT, RT의 경우)
                    size_match = True
                    size_numbers = [int(x) for x in re.findall(r'\d+', size_info)]
                    if '300A' in method_key and not any(x >= 300 for x in size_numbers): size_match = False
                    if '250A' in method_key and 250 not in size_numbers: size_match = False
                    if '200A' in method_key and 200 not in size_numbers: size_match = False
                    if '150A' in method_key and not any(x in (100, 125, 150) for x in size_numbers): size_match = False
                    if '80A' in method_key and not any(x <= 80 for x in size_numbers): size_match = False
                    
                    if size_match and (is_night == target_is_night):
                        qty += pd.to_numeric(row.get('Usage', 0), errors='coerce')
            return float(qty) if pd.notna(qty) else 0

        # 장비 수집
        equips = []
        if not today_records.empty and '장비명' in today_records.columns:
            from site_apps.central.src.aux_equipment import report_names
            equips = report_names(today_records)
        if not equips:
            equips = [_clean_str(self.cb_daily_equip.get())]
            
        equip_str = ", ".join(filter(None, equips))

        # 인원 수집
        inspectors = []
        managers = []
        if not today_records.empty:
            for _, row in today_records.iterrows():
                for i in range(1, 11):
                    u_key = 'User' if i == 1 else f'User{i}'
                    name = str(row.get(u_key, '')).strip()
                    if name and name != 'nan':
                        clean_n = _clean_name(name)
                        if '대리인' in name or '안전' in name or '관리' in name:
                            if clean_n not in managers: managers.append(clean_n)
                        else:
                            if clean_n not in inspectors: inspectors.append(clean_n)
                            
        if not inspectors and not managers:
            managers_raw = ""
            managers = [_clean_name(x.strip()) for x in managers_raw.split(',') if x.strip()]
            for i in range(1, 11):
                group = getattr(self, f'worker_group{i}', None)
                if group:
                    if not worker_is_selected(group): continue
                    w_name = _clean_name(group.get_worker().strip())
                    if w_name and w_name not in inspectors: inspectors.append(w_name)

        # 엑셀 오픈
        wb = openpyxl.load_workbook(template_path)
        ws = wb.active

        # 상단 데이터 맵핑
        ws['A4'] = f"검사일자 : {date_val.strftime('%Y년 %m월 %d일')}           날 씨 : "
        
        # 1. 작업 물량 및 누계 현황 맵핑 (PAUT: 행 7~11, RT: 13~16, MT: 18~19, PT: 20~21)
        # [행번호, method_key, base_method]
        mapping_rows = [
            [7, 'PAUT_300A_D', 'PAUT'],
            [8, 'PAUT_300A_N', 'PAUT'],
            [9, 'PAUT_250A', 'PAUT'],
            [10, 'PAUT_200A_D', 'PAUT'],
            [11, 'PAUT_200A_N', 'PAUT'],
            # 12 is PAUT 소계
            [13, 'RT_150A_D', 'RT'],
            [14, 'RT_150A_N', 'RT'],
            [15, 'RT_80A_D', 'RT'],
            [16, 'RT_80A_N', 'RT'],
            # 17 is RT 소계
            [18, 'MT_D', 'MT'],
            [19, 'MT_N', 'MT'],
            [20, 'PT_D', 'PT'],
            [21, 'PT_N', 'PT']
        ]

        for r, m_key, base_m in mapping_rows:
            today_qty = calc_method_qty(today_records, m_key, base_m)
            past_qty = calc_method_qty(past_records, m_key, base_m)
            total_qty = today_qty + past_qty
            
            # 예상량(C열) 가져오기
            exp_val = ws.cell(row=r, column=3).value
            try: exp_val = float(exp_val) if exp_val else 0
            except: exp_val = 0
            
            # 전일 누계 (D열)
            ws.cell(row=r, column=4).value = past_qty if past_qty > 0 else ''
            # 금일 작업 (E열)
            ws.cell(row=r, column=5).value = today_qty if today_qty > 0 else ''
            # 총 누계 (F열)
            ws.cell(row=r, column=6).value = total_qty if total_qty > 0 else ''
            # 공정률 (G열)
            if exp_val > 0 and total_qty > 0:
                ws.cell(row=r, column=7).value = round((total_qty / exp_val) * 100, 1)
            else:
                ws.cell(row=r, column=7).value = ''

        # 소계 엑셀 수식은 엑셀 자체에 없으므로 파이썬에서 계산하여 넣음
        paut_past = sum([calc_method_qty(past_records, m, 'PAUT') for m in ['PAUT_300A_D', 'PAUT_300A_N', 'PAUT_250A', 'PAUT_200A_D', 'PAUT_200A_N']])
        paut_today = sum([calc_method_qty(today_records, m, 'PAUT') for m in ['PAUT_300A_D', 'PAUT_300A_N', 'PAUT_250A', 'PAUT_200A_D', 'PAUT_200A_N']])
        ws.cell(row=12, column=4).value = paut_past if paut_past > 0 else ''
        ws.cell(row=12, column=5).value = paut_today if paut_today > 0 else ''
        ws.cell(row=12, column=6).value = (paut_past + paut_today) if (paut_past + paut_today) > 0 else ''
        
        rt_past = sum([calc_method_qty(past_records, m, 'RT') for m in ['RT_150A_D', 'RT_150A_N', 'RT_80A_D', 'RT_80A_N']])
        rt_today = sum([calc_method_qty(today_records, m, 'RT') for m in ['RT_150A_D', 'RT_150A_N', 'RT_80A_D', 'RT_80A_N']])
        ws.cell(row=17, column=4).value = rt_past if rt_past > 0 else ''
        ws.cell(row=17, column=5).value = rt_today if rt_today > 0 else ''
        ws.cell(row=17, column=6).value = (rt_past + rt_today) if (rt_past + rt_today) > 0 else ''
        
        # 3. 장비 및 인원 맵핑 (v20 양식 구조 맞춤)
        ws.cell(row=7, column=12).value = equip_str # L7 (L7:M7 병합)

        ws['Q7'] = len(managers) if managers else ''
        ws['Q8'] = ", ".join(managers) if managers else ''
        
        ws['R7'] = len(inspectors) if inspectors else ''
        ws['R8'] = ", ".join(inspectors) if inspectors else ''

        # 결과 저장
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        default_name = f"{site}_검사일보_{date_val.strftime('%y%m%d')}.xlsx"
        
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx", 
            initialfile=default_name, 
            filetypes=[("Excel File", "*.xlsx")], 
            title="중앙지사 검사일보 저장"
        )
        
        if filepath:
            wb.save(filepath)
            wb.close()
            messagebox.showinfo("완료", f"중앙지사 엑셀 작업일보 생성이 완료되었습니다.\n(2번 비파괴검사결과서는 엑셀에서 직접 기입해 주세요)\n\n{filepath}")
            os.startfile(filepath)
            
    except Exception as e:
        import traceback
        traceback.print_exc()
        messagebox.showerror("오류", f"엑셀 생성 중 오류가 발생했습니다: {str(e)}")


def export_weekly_report_impl(self):
    """주간 업무보고 엑셀(알아서 생성) 출력"""
    import pandas as pd
    import datetime
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    import os

    top = tk.Toplevel(self.root)
    top.title("주간 업무보고서 출력")
    top.geometry("460x390")
    top.transient(self.root)
    top.grab_set()

    ttk.Label(top, text="🗓️ 주간 업무보고서 자동 생성", font=('Arial', 12, 'bold')).pack(pady=10)

    now = datetime.datetime.now()
    monday = now - datetime.timedelta(days=now.weekday())
    sunday = monday + datetime.timedelta(days=6)

    frame1 = ttk.Frame(top)
    frame1.pack(pady=10)

    from tkcalendar import DateEntry

    ttk.Label(frame1, text="공사명:").grid(row=0, column=0, padx=5, pady=5, sticky='e')
    project_name_entry = ttk.Entry(frame1, width=35)
    project_name_entry.grid(row=0, column=1, padx=5, pady=5)

    ttk.Label(frame1, text="시작일 (월):").grid(row=1, column=0, padx=5, pady=5)
    start_cal = DateEntry(frame1, width=12, background='darkblue',
                          foreground='white', borderwidth=2,
                          date_pattern='yyyy-mm-dd', year=monday.year, month=monday.month, day=monday.day)
    start_cal.grid(row=1, column=1, padx=5, sticky='w')

    ttk.Label(frame1, text="종료일 (일):").grid(row=2, column=0, padx=5, pady=5)
    end_cal = DateEntry(frame1, width=12, background='darkblue',
                        foreground='white', borderwidth=2,
                        date_pattern='yyyy-mm-dd', year=sunday.year, month=sunday.month, day=sunday.day)
    end_cal.grid(row=2, column=1, padx=5, sticky='w')

    frame2 = ttk.Frame(top)
    frame2.pack(pady=5, fill='both', expand=True, padx=10)
    ttk.Label(frame2, text="다음 주 작업 예정 및 의견:").pack(anchor='w')
    next_week_txt = tk.Text(frame2, height=4, width=40)
    next_week_txt.pack(fill='both', expand=True, pady=5)

    def do_export():
        project_name = project_name_entry.get().strip()
        s_date = start_cal.get().strip()
        e_date = end_cal.get().strip()
        next_plan = next_week_txt.get("1.0", "end-1c").strip()
        if not project_name:
            messagebox.showwarning("입력 필요", "공사명을 입력해주세요.")
            project_name_entry.focus_set()
            return
        try:
            s_dt = pd.to_datetime(s_date).date()
            e_dt = pd.to_datetime(e_date).date()
        except Exception:
            messagebox.showerror("오류", "날짜 형식이 올바르지 않습니다 (YYYY-MM-DD)")
            return

        history_path = get_history_path()
        if not os.path.exists(history_path):
            messagebox.showinfo("알림", "저장된 작업일보 데이터가 없습니다.")
            return

        import json
        with open(history_path, 'r', encoding='utf-8') as f:
            history = json.load(f)

        week_data = {}
        cumulative_data = {}
        for d_str, data in history.items():
            try:
                d_obj = pd.to_datetime(d_str).date()
                if s_dt <= d_obj <= e_dt:
                    week_data[d_str] = data
                if d_obj <= e_dt:
                    cumulative_data[d_str] = data
            except: pass

        if not cumulative_data:
            messagebox.showinfo("알림", f"{e_date}까지 저장된 작업일보 데이터가 없습니다.")
            return

        def get_personnel_count(data, use_report_default=False):
            """작업일보의 검사원 인원을 안전하게 숫자로 변환한다."""
            p_data = data.get('personnel_data', {})
            count = 0
            try:
                inspector = str(p_data.get('검사원_인원', '0')).strip()
                manager = str(p_data.get('검사원_현장대리인', '0')).strip()
                count = (int(inspector) if inspector.isdigit() else 0) + (int(manager) if manager.isdigit() else 0)
            except Exception:
                count = 0
            return count if count > 0 else (1 if use_report_default else 0)

        cumulative_time = 0
        cumulative_ndt_methods = {}
        for data in cumulative_data.values():
            cumulative_time += 8
            for result in data.get('ndt_results', []):
                method = str(result.get("검사방법", "")).strip().upper()
                if method:
                    cumulative_ndt_methods[method] = cumulative_ndt_methods.get(method, 0) + 1

        wb = Workbook()
        ws = wb.active
        ws.title = "주간업무보고"

        title_font = Font(name='맑은 고딕', size=16, bold=True)
        head_font = Font(name='맑은 고딕', size=11, bold=True, color='FFFFFF')
        norm_font = Font(name='맑은 고딕', size=10)
        head_fill = PatternFill(start_color='4F81BD', end_color='4F81BD', fill_type='solid')
        center_align = Alignment(horizontal='center', vertical='center')
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        ws.merge_cells('A1:G1')
        ws['A1'] = f"주간 업무보고서 ({s_date} ~ {e_date})"
        ws['A1'].font = title_font
        ws['A1'].alignment = center_align

        ws.merge_cells('A2:G2')
        ws['A2'] = f"공사명: {project_name}"
        ws['A2'].font = Font(name='맑은 고딕', size=11, bold=True)
        ws['A2'].alignment = center_align

        headers = ['일자', '현장명', '작업내용', '투입인원', '작업시간', '검사실적', '비고']
        for col_num, head in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col_num)
            cell.value = head
            cell.font = head_font
            cell.fill = head_fill
            cell.alignment = center_align
            cell.border = thin_border

        row_idx = 4
        total_time = 0
        total_ndt_methods = {}
        center_align_wrap = Alignment(horizontal='center', vertical='center', wrap_text=True)

        # Sort dates
        sorted_dates = sorted(week_data.keys())

        if not sorted_dates:
            ws.cell(row=row_idx, column=1, value=f"{s_date} ~ {e_date}").alignment = center_align
            ws.cell(row=row_idx, column=2, value="중앙지사 관내").alignment = center_align
            ws.cell(row=row_idx, column=3, value="금주 신규 작업 없음").alignment = center_align
            ws.cell(row=row_idx, column=4, value="-").alignment = center_align
            ws.cell(row=row_idx, column=5, value="0 시간").alignment = center_align
            ws.cell(row=row_idx, column=6, value="0 POINT").alignment = center_align
            ws.cell(row=row_idx, column=7, value="누계 및 차주 계획 참조").alignment = center_align_wrap
            for col_num in range(1, 8):
                ws.cell(row=row_idx, column=col_num).border = thin_border
                ws.cell(row=row_idx, column=col_num).font = norm_font
            row_idx += 1

        for date_val in sorted_dates:
            data = week_data[date_val]

            # NDT results
            ndt_results = data.get('ndt_results', [])
            daily_ndt_methods = {}
            for r in ndt_results:
                method = str(r.get("검사방법", "")).strip().upper()
                if method:
                    daily_ndt_methods[method] = daily_ndt_methods.get(method, 0) + 1
                    total_ndt_methods[method] = total_ndt_methods.get(method, 0) + 1

            daily_ndt_texts = [f"{m} {c} POINT" for m, c in daily_ndt_methods.items()]
            ndt_text = "\n".join(daily_ndt_texts) if daily_ndt_texts else "0 POINT"

            # 같은 날짜·현장에서 검사한 관경을 한 작업내용으로 묶어 표시한다.
            pipe_sizes = {
                str(r.get("관경", "")).strip()
                for r in ndt_results
                if str(r.get("관경", "")).strip()
            }
            sorted_pipe_sizes = sorted(
                pipe_sizes,
                key=lambda value: int(re.search(r'\d+', value).group())
                if re.search(r'\d+', value) else 0,
                reverse=True,
            )
            work_description = (
                f"{', '.join(sorted_pipe_sizes)} 검사 진행"
                if sorted_pipe_sizes else "검사 작업 진행"
            )

            # Personnel data
            w_count = get_personnel_count(data, use_report_default=True)

            total_time += 8

            ws.cell(row=row_idx, column=1, value=str(date_val)).alignment = center_align
            ws.cell(row=row_idx, column=2, value="중앙지사 관내").alignment = center_align
            ws.cell(row=row_idx, column=3, value=work_description).alignment = center_align_wrap
            ws.cell(row=row_idx, column=4, value=f"{w_count} 명").alignment = center_align
            ws.cell(row=row_idx, column=5, value="8 시간").alignment = center_align
            ws.cell(row=row_idx, column=6, value=ndt_text).alignment = center_align_wrap
            ws.cell(row=row_idx, column=7, value="").alignment = center_align

            for col_num in range(1, 8):
                ws.cell(row=row_idx, column=col_num).border = thin_border
                ws.cell(row=row_idx, column=col_num).font = norm_font

            row_idx += 1

        ws.column_dimensions['A'].width = 15
        ws.column_dimensions['B'].width = 25
        ws.column_dimensions['C'].width = 25
        ws.column_dimensions['D'].width = 12
        ws.column_dimensions['E'].width = 12
        ws.column_dimensions['F'].width = 22
        ws.column_dimensions['G'].width = 20

        ws.merge_cells(f'A{row_idx}:C{row_idx}')
        ws.cell(row=row_idx, column=1, value="주간 합계").alignment = center_align
        ws.cell(row=row_idx, column=1).font = Font(name='맑은 고딕', size=11, bold=True)
        for col_num in range(1, 4):
            ws.cell(row=row_idx, column=col_num).border = thin_border

        ws.cell(row=row_idx, column=4, value=f"-").alignment = center_align
        ws.cell(row=row_idx, column=5, value=f"{total_time} 시간").alignment = center_align

        total_ndt_texts = [f"{m} {c} POINT" for m, c in total_ndt_methods.items()]
        total_ndt_text = "\n".join(total_ndt_texts) if total_ndt_texts else "0 POINT"
        ws.cell(row=row_idx, column=6, value=total_ndt_text).alignment = center_align_wrap
        ws.cell(row=row_idx, column=7, value="").alignment = center_align

        for col_num in range(4, 8):
            ws.cell(row=row_idx, column=col_num).border = thin_border
            ws.cell(row=row_idx, column=col_num).font = Font(name='맑은 고딕', size=11, bold=True)

        if not week_data:
            row_idx += 2
            cumulative_ndt_texts = [f"{m} {c} POINT" for m, c in sorted(cumulative_ndt_methods.items())]
            cumulative_ndt_text = ", ".join(cumulative_ndt_texts) if cumulative_ndt_texts else "검사실적 없음"
            ws.merge_cells(f'A{row_idx}:G{row_idx+2}')
            cell = ws.cell(row=row_idx, column=1)
            cell.value = (f"■ 작업 누계 (최초 작업일 ~ {e_date})\n"
                          f"작업일보 {len(cumulative_data)}일 / 투입 {cumulative_time} 시간 / {cumulative_ndt_text}")
            cell.font = Font(name='맑은 고딕', size=11)
            cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
            for r in range(row_idx, row_idx+3):
                for c in range(1, 8):
                    ws.cell(row=r, column=c).border = thin_border

        row_idx += 4 if not week_data else 2
        ws.merge_cells(f'A{row_idx}:G{row_idx+2}')
        cell = ws.cell(row=row_idx, column=1)
        cell.value = f"■ 다음 주 작업 예정 및 의견\n{next_plan or '미입력'}"
        cell.font = Font(name='맑은 고딕', size=11)
        cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)

        # Draw borders for the merged cell block
        for r in range(row_idx, row_idx+3):
            for c in range(1, 8):
                ws.cell(row=r, column=c).border = thin_border

        out_name = f"주간업무보고_{s_date.replace('-','')}_{e_date.replace('-','')}.xlsx"
        out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), out_name)

        try:
            wb.save(out_path)

            msg_extra = ""
            try:
                import win32com.client as win32
                import time
                excel = win32.gencache.EnsureDispatch('Excel.Application')
                excel.Visible = False
                excel.DisplayAlerts = False

                wb_com = excel.Workbooks.Open(out_path)
                ws_com = wb_com.Sheets(1)
                ws_com.UsedRange.CopyPicture(Format=2)
                time.sleep(0.5)
                wb_com.Close(False)
                excel.Quit()

                msg_extra = "\n\n💡 꿀팁: 카카오톡 전송을 위해 표가 '이미지'로 복사되었습니다!\n(카톡 창에서 Ctrl+V 를 누르시면 표가 그대로 붙여넣기 됩니다)"
            except Exception as ex:
                msg_extra = f"\n(클립보드 이미지 복사 실패: {ex})"

            messagebox.showinfo("완료", f"주간보고서가 생성되었습니다.\n{out_path}{msg_extra}")
            os.startfile(out_path)
            top.destroy()
        except Exception as e:
            messagebox.showerror("오류", f"파일 저장 실패: {e}")

    ttk.Button(top, text="보고서 생성 및 열기", command=do_export, style='Accent.TButton' if 'Accent.TButton' in self.style.theme_names() else 'TButton').pack(pady=20)

def export_monthly_ndt_report_impl(self):
    """일일 사용 데이터를 월용역진도보고서 '3. 비파괴검사 현황' 시트에 자동 채움"""
    import copy

    if self.daily_usage_df.empty:
        messagebox.showinfo("알림", "저장된 작업일보 데이터가 없습니다.")
        return

    # --- 0. 설정 다이얼로그 ---
    top = tk.Toplevel(self.root)
    top.title("월간 진도보고서 - 비파괴검사 현황 내보내기")
    top.geometry("550x500")
    top.transient(self.root)
    top.grab_set()

    ttk.Label(top, text="📊 월간 진도보고서 비파괴검사 현황 자동 입력", 
              font=('Arial', 12, 'bold')).pack(pady=10)

    # 대상 기간 선택
    period_frame = ttk.LabelFrame(top, text="1. 대상 기간 (당월)")
    period_frame.pack(fill='x', padx=15, pady=5)

    now = datetime.datetime.now()
    ttk.Label(period_frame, text="연도:").pack(side='left', padx=5)
    year_var = tk.IntVar(value=now.year)
    ttk.Spinbox(period_frame, from_=2024, to=2030, textvariable=year_var, width=6).pack(side='left')
    ttk.Label(period_frame, text="  월:").pack(side='left', padx=5)
    month_var = tk.IntVar(value=now.month)
    ttk.Spinbox(period_frame, from_=1, to=12, textvariable=month_var, width=4).pack(side='left')

    ttk.Label(period_frame, text="  문서번호:").pack(side='left', padx=5)
    doc_num_var = tk.StringVar(value="01")
    ttk.Entry(period_frame, textvariable=doc_num_var, width=5).pack(side='left')

    # 현장 → 열배관/관리소 매핑
    map_frame = ttk.LabelFrame(top, text="2. 현장명 → 시트 매핑 (열배관 / 관리소 구분)")
    map_frame.pack(fill='x', padx=15, pady=5)

    # 현재 데이터에 존재하는 현장명(Site) 목록 추출
    site_list = sorted(self.daily_usage_df['Site'].dropna().unique().tolist()) if 'Site' in self.daily_usage_df.columns else []

    ttk.Label(map_frame, text="열배관 현장명 (쉼표 구분):").pack(anchor='w', padx=5, pady=2)
    main_var = tk.StringVar(value="")
    ttk.Entry(map_frame, textvariable=main_var, width=60).pack(padx=5, pady=2)

    ttk.Label(map_frame, text="관리소 현장명 (쉼표 구분):").pack(anchor='w', padx=5, pady=2)
    mgmt_var = tk.StringVar(value="")
    ttk.Entry(map_frame, textvariable=mgmt_var, width=60).pack(padx=5, pady=2)

    ttk.Label(map_frame, text=f"💡 현재 등록된 현장: {', '.join(site_list) if site_list else '(없음)'}", 
              foreground='blue', wraplength=500).pack(padx=5, pady=5)

    # 파일 선택
    file_frame = ttk.LabelFrame(top, text="3. 대상 엑셀 파일")
    file_frame.pack(fill='x', padx=15, pady=5)

    file_var = tk.StringVar(value=r"C:\Users\jjch2\Desktop\월용역진도보고서.xlsx")
    ttk.Entry(file_frame, textvariable=file_var, width=55).pack(side='left', padx=5, pady=5)
    def browse_file():
        fp = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")], title="월용역진도보고서 선택")
        if fp: file_var.set(fp)
    ttk.Button(file_frame, text="찾기", command=browse_file).pack(side='left', padx=5)

    # 진행
    log_text = tk.Text(top, height=8, state='disabled', wrap='word')
    log_text.pack(fill='both', expand=True, padx=15, pady=5)

    def log(msg):
        log_text.config(state='normal')
        log_text.insert(tk.END, msg + "\n")
        log_text.see(tk.END)
        log_text.config(state='disabled')
        top.update()

    def do_export():
        try:
            year = year_var.get()
            month = month_var.get()
            doc_num = doc_num_var.get().strip() or "01"
            filepath = file_var.get().strip()

            main_sites = [s.strip() for s in main_var.get().split(',') if s.strip()]
            mgmt_sites = [s.strip() for s in mgmt_var.get().split(',') if s.strip()]

            if not main_sites and not mgmt_sites:
                messagebox.showwarning("입력 오류", "열배관 또는 관리소에 해당하는 현장명을 최소 1개 입력해주세요.")
                return
            if not filepath:
                messagebox.showwarning("입력 오류", "대상 엑셀 파일을 선택해주세요.")
                return

            log(f"▶ 기간: {year}년 {month}월")
            log(f"▶ 열배관 현장: {main_sites}")
            log(f"▶ 관리소 현장: {mgmt_sites}")

            # --- 1. 데이터 필터링 ---
            df = self.daily_usage_df.copy()
            df['_date'] = pd.to_datetime(df['Date'], errors='coerce')
            df = df[(df['_date'].dt.year == year) & (df['_date'].dt.month == month)]
            log(f"▶ {year}년 {month}월 데이터: {len(df)}건")

            if df.empty:
                messagebox.showinfo("결과", f"{year}년 {month}월 데이터가 없습니다.")
                return

            # --- 2. 관경 → 필름규격 매핑 ---
            def get_film_type(inch_val):
                """관경(Inch)에서 RT 필름 규격 결정"""
                try:
                    v = str(inch_val).replace('"', '').replace("'", '').strip()
                    if '/' in v and len(v) <= 5:  # 3/4 등 분수
                        parts = v.split('/')
                        num = float(parts[0]) / float(parts[1])
                    else:
                        num = float(v)

                    if num >= 20:
                        return 'B'   # B-TYPE (3⅓"×17")
                    elif num >= 6:
                        return 'A'   # A-TYPE (3⅓"×12")
                    else:
                        return 'A/2' # A/2-TYPE (3⅓"×6")
                except:
                    return 'A'  # 기본값

            # --- 3. 집계 함수 ---
            def aggregate_site(site_df):
                """한 현장(열배관 or 관리소)의 데이터를 집계"""
                result = {
                    'RT': {'B': {}, 'A': {}, 'A/2': {}},
                    'UT': {'data': {}},
                    'PT': {'data': {}}
                }
                # 각 셀의 키: (shift, insp_type)
                # shift: '주간' or '야간/휴일'
                # insp_type: 'ORI' or 'REP'

                for _, row in site_df.iterrows():
                    method = str(row.get('검사방법', '')).strip().upper()
                    if method not in ['RT', 'UT', 'PT', 'PAUT', 'MT']:
                        continue

                    work_type = str(row.get('작업형태', '주간')).strip()
                    if work_type in ['야간', '휴일', '야간/휴일']:
                        shift = '야간/휴일'
                    else:
                        shift = '주간'

                    insp_type = str(row.get('검사구분', 'ORI')).strip().upper()
                    if insp_type not in ['ORI', 'REP']:
                        insp_type = 'ORI'

                    try:
                        qty = float(str(row.get('검사량', row.get('Usage', 0))).replace(',', '') or 0)
                    except:
                        qty = 0.0

                    try:
                        joints = float(str(row.get('조인트수', 0)).replace(',', '') or 0)
                    except:
                        joints = 0.0

                    try:
                        corr_factor = float(str(row.get('보정계수', 1)).replace(',', '') or 1)
                    except:
                        corr_factor = 1.0

                    try:
                        adj_qty = float(str(row.get('환산물량', 0)).replace(',', '') or 0)
                    except:
                        adj_qty = qty * corr_factor

                    key = (shift, insp_type)

                    if method == 'RT':
                        film_type = get_film_type(row.get('관경(Inch)', ''))
                        bucket = result['RT'][film_type]
                        if key not in bucket:
                            bucket[key] = {'qty': 0, 'joints': 0}
                        bucket[key]['qty'] += qty
                        bucket[key]['joints'] += joints

                    elif method == 'UT':
                        bucket = result['UT']['data']
                        if key not in bucket:
                            bucket[key] = {'qty': 0, 'joints': 0, 'adj_qty': 0}
                        bucket[key]['qty'] += qty
                        bucket[key]['joints'] += joints
                        bucket[key]['adj_qty'] += adj_qty

                    elif method == 'PT':
                        bucket = result['PT']['data']
                        if key not in bucket:
                            bucket[key] = {'qty': 0, 'joints': 0, 'adj_qty': 0}
                        bucket[key]['qty'] += qty
                        bucket[key]['joints'] += joints
                        bucket[key]['adj_qty'] += adj_qty

                return result

            # --- 4. 열배관 / 관리소 각각 집계 ---
            main_df = df[df['Site'].isin(main_sites)] if main_sites else pd.DataFrame()
            mgmt_df = df[df['Site'].isin(mgmt_sites)] if mgmt_sites else pd.DataFrame()

            log(f"▶ 열배관 데이터: {len(main_df)}건, 관리소 데이터: {len(mgmt_df)}건")

            main_agg = aggregate_site(main_df) if not main_df.empty else None
            mgmt_agg = aggregate_site(mgmt_df) if not mgmt_df.empty else None

                            # --- 5. 엑셀 기입 (openpyxl로 변경) ---
            import openpyxl
            from openpyxl.utils import get_column_letter
            import os

            save_path = filepath
            wb = openpyxl.load_workbook(save_path)

            def write_ndt_sheet_openpyxl(ws, agg):
                if agg is None: return

                def get_val(bucket, shift, insp_type, field='qty'):
                    key = (shift, insp_type)
                    return bucket.get(key, {}).get(field, 0)

                def safe_set(row, col, val):
                    if val and val != 0:
                        ws.cell(row=row, column=col).value = round(val, 2) if isinstance(val, float) and val != int(val) else int(val) if isinstance(val, float) and val == int(val) else val

                rt_types = {'B': 8, 'A': 9, 'A/2': 10}
                rt_total = {'주간_joints': 0, '주간_ORI': 0, '주간_REP': 0,
                            '야간_joints': 0, '야간_ORI': 0, '야간_REP': 0}

                for film_key, row_num in rt_types.items():
                    bucket = agg['RT'][film_key]
                    j_day = get_val(bucket, '주간', 'ORI', 'joints') + get_val(bucket, '주간', 'REP', 'joints')
                    ori_day = get_val(bucket, '주간', 'ORI', 'qty')
                    rep_day = get_val(bucket, '주간', 'REP', 'qty')
                    sum_day = ori_day + rep_day
                    j_night = get_val(bucket, '야간/휴일', 'ORI', 'joints') + get_val(bucket, '야간/휴일', 'REP', 'joints')
                    ori_night = get_val(bucket, '야간/휴일', 'ORI', 'qty')
                    rep_night = get_val(bucket, '야간/휴일', 'REP', 'qty')
                    sum_night = ori_night + rep_night
                    j_total = j_day + j_night
                    ori_total = ori_day + ori_night
                    rep_total = rep_day + rep_night
                    sum_total = ori_total + rep_total

                    safe_set(row_num, 3, j_day)
                    safe_set(row_num, 4, ori_day)
                    safe_set(row_num, 5, rep_day)
                    safe_set(row_num, 6, sum_day)
                    safe_set(row_num, 7, j_night)
                    safe_set(row_num, 8, ori_night)
                    safe_set(row_num, 9, rep_night)
                    safe_set(row_num, 10, sum_night)
                    safe_set(row_num, 11, j_total)
                    safe_set(row_num, 12, ori_total)
                    safe_set(row_num, 13, rep_total)
                    safe_set(row_num, 14, sum_total)

                    rt_total['주간_joints'] += j_day
                    rt_total['주간_ORI'] += ori_day
                    rt_total['주간_REP'] += rep_day
                    rt_total['야간_joints'] += j_night
                    rt_total['야간_ORI'] += ori_night
                    rt_total['야간_REP'] += rep_night

                safe_set(11, 3, rt_total['주간_joints'])
                safe_set(11, 4, rt_total['주간_ORI'])
                safe_set(11, 5, rt_total['주간_REP'])
                safe_set(11, 6, rt_total['주간_ORI'] + rt_total['주간_REP'])
                safe_set(11, 7, rt_total['야간_joints'])
                safe_set(11, 8, rt_total['야간_ORI'])
                safe_set(11, 9, rt_total['야간_REP'])
                safe_set(11, 10, rt_total['야간_ORI'] + rt_total['야간_REP'])
                safe_set(11, 11, rt_total['주간_joints'] + rt_total['야간_joints'])
                safe_set(11, 12, rt_total['주간_ORI'] + rt_total['야간_ORI'])
                safe_set(11, 13, rt_total['주간_REP'] + rt_total['야간_REP'])
                safe_set(11, 14, rt_total['주간_ORI'] + rt_total['주간_REP'] + rt_total['야간_ORI'] + rt_total['야간_REP'])

                ut = agg['UT']['data']
                ut_day_j = get_val(ut, '주간', 'ORI', 'joints') + get_val(ut, '주간', 'REP', 'joints')
                ut_day_ori = get_val(ut, '주간', 'ORI', 'qty')
                ut_day_rep = get_val(ut, '주간', 'REP', 'qty')
                ut_night_j = get_val(ut, '야간/휴일', 'ORI', 'joints') + get_val(ut, '야간/휴일', 'REP', 'joints')
                ut_night_ori = get_val(ut, '야간/휴일', 'ORI', 'qty')
                ut_night_rep = get_val(ut, '야간/휴일', 'REP', 'qty')

                safe_set(12, 3, ut_day_j)
                safe_set(12, 4, ut_day_ori + ut_day_rep)
                safe_set(12, 6, ut_day_ori + ut_day_rep)
                safe_set(12, 7, ut_night_j)
                safe_set(12, 8, ut_night_ori + ut_night_rep)
                safe_set(12, 10, ut_night_ori + ut_night_rep)
                safe_set(12, 11, ut_day_j + ut_night_j)
                safe_set(12, 12, ut_day_ori + ut_day_rep + ut_night_ori + ut_night_rep)
                safe_set(12, 14, ut_day_ori + ut_day_rep + ut_night_ori + ut_night_rep)

                ut_day_adj = get_val(ut, '주간', 'ORI', 'adj_qty') + get_val(ut, '주간', 'REP', 'adj_qty')
                ut_night_adj = get_val(ut, '야간/휴일', 'ORI', 'adj_qty') + get_val(ut, '야간/휴일', 'REP', 'adj_qty')
                safe_set(13, 4, ut_day_adj)
                safe_set(13, 6, ut_day_adj)
                safe_set(13, 8, ut_night_adj)
                safe_set(13, 10, ut_night_adj)
                safe_set(13, 12, ut_day_adj + ut_night_adj)
                safe_set(13, 14, ut_day_adj + ut_night_adj)

                pt = agg['PT']['data']
                pt_day_j = get_val(pt, '주간', 'ORI', 'joints') + get_val(pt, '주간', 'REP', 'joints')
                pt_day_ori = get_val(pt, '주간', 'ORI', 'qty')
                pt_day_rep = get_val(pt, '주간', 'REP', 'qty')
                pt_night_j = get_val(pt, '야간/휴일', 'ORI', 'joints') + get_val(pt, '야간/휴일', 'REP', 'joints')
                pt_night_ori = get_val(pt, '야간/휴일', 'ORI', 'qty')
                pt_night_rep = get_val(pt, '야간/휴일', 'REP', 'qty')

                safe_set(14, 3, pt_day_j)
                safe_set(14, 4, pt_day_ori + pt_day_rep)
                safe_set(14, 6, pt_day_ori + pt_day_rep)
                safe_set(14, 7, pt_night_j)
                safe_set(14, 8, pt_night_ori + pt_night_rep)
                safe_set(14, 10, pt_night_ori + pt_night_rep)
                safe_set(14, 11, pt_day_j + pt_night_j)
                safe_set(14, 12, pt_day_ori + pt_day_rep + pt_night_ori + pt_night_rep)
                safe_set(14, 14, pt_day_ori + pt_day_rep + pt_night_ori + pt_night_rep)

                pt_day_adj = get_val(pt, '주간', 'ORI', 'adj_qty') + get_val(pt, '주간', 'REP', 'adj_qty')
                pt_night_adj = get_val(pt, '야간/휴일', 'ORI', 'adj_qty') + get_val(pt, '야간/휴일', 'REP', 'adj_qty')
                safe_set(15, 4, pt_day_adj)
                safe_set(15, 6, pt_day_adj)
                safe_set(15, 8, pt_night_adj)
                safe_set(15, 10, pt_night_adj)
                safe_set(15, 12, pt_day_adj + pt_night_adj)
                safe_set(15, 14, pt_day_adj + pt_night_adj)

                total_film = rt_total['주간_ORI'] + rt_total['주간_REP'] + rt_total['야간_ORI'] + rt_total['야간_REP']
                safe_set(16, 11, 0)
                safe_set(16, 12, total_film)

            sheet_names = wb.sheetnames
            if main_agg:
                sheet_name = '3. 비파괴검사 현황 (열배관)'
                if sheet_name in sheet_names:
                    write_ndt_sheet_openpyxl(wb[sheet_name], main_agg)
                    log(f"✅ '{sheet_name}' 시트 기입 완료")
                else:
                    log(f"⚠️ '{sheet_name}' 시트를 찾을 수 없습니다.")

            if mgmt_agg:
                sheet_name = '3. 비파괴검사 현황 (관리소)'
                if sheet_name in sheet_names:
                    write_ndt_sheet_openpyxl(wb[sheet_name], mgmt_agg)
                    log(f"✅ '{sheet_name}' 시트 기입 완료")
                else:
                    log(f"⚠️ '{sheet_name}' 시트를 찾을 수 없습니다.")

            history = {}
            import json, os
            history_path = get_history_path()
            if os.path.exists(history_path):
                with open(history_path, 'r', encoding='utf-8') as f:
                    history = json.load(f)

            # (Old PAUT block removed in favor of tagged writer below)
            wb.save(save_path)
            wb.close()
            log(f"\n🎉 저장 완료: {save_path}")
            messagebox.showinfo("완료", f"월간 진도보고서 비파괴검사 현황이 업데이트되었습니다.\n{save_path}")
            import os
            os.startfile(os.path.dirname(save_path))


            # --- 4.5 NDT 결과서 섹션 태그 기반 자동 기입 ---
            try:
                import sys as _sys
                import os as _os
                _src = _os.path.dirname(_os.path.abspath(__file__))
                if _src not in _sys.path:
                    _sys.path.insert(0, _src)

                from site_apps.central.src.tagged_ndt_writer import write_all_tagged_sections
                target_month_str = f"{year}-{month:02d}"

                # 저장된 파일을 다시 열어 NDT 기입
                import openpyxl as _opx
                wb2 = _opx.load_workbook(save_path)
                ws2 = wb2.worksheets[0]

                # 태그 기반 NDT 섹션 기입
                write_all_tagged_sections(ws2, history, target_month_str, log_func=log)

                wb2.save(save_path)
                wb2.close()
                log("✅ 태그 기반 NDT 결과서 전체 기입 완료")

                log("✅ NDT 결과서 전체 기입 완료")

                log("✅ NDT 결과서 전체 기입 완료")

            except Exception as ex:
                log(f"⚠️ NDT 결과서 기입 오류 (무시됨): {ex}")
                import traceback
                log(traceback.format_exc())
                log("✅ 태그 변환 완료")
            except Exception as ex:
                log(f"⚠️ 태그 변환 중 오류 (무시됨): {ex}")

        except Exception as e:

            log(f"❌ 오류 발생: {e}")
            import traceback
            log(traceback.format_exc())
            messagebox.showerror("오류", f"내보내기 중 오류: {e}")

    btn_frame = ttk.Frame(top)
    btn_frame.pack(pady=10)
    ttk.Button(btn_frame, text="📊 내보내기 실행", command=do_export, style='Accent.TButton' if 'Accent.TButton' in self.style.theme_names() else 'TButton').pack(side='left', padx=10)
    ttk.Button(btn_frame, text="취소", command=top.destroy).pack(side='left', padx=10)

def export_invoice_excel_impl(self):
    """Export current filtered daily usage to official Invoice Excel (가산가평기성양식)"""
    try:
        import win32com.client as win32
    except ImportError:
        messagebox.showerror("오류", "pywin32 라이브러리가 설치되어 있지 않습니다.\n(pip install pywin32)")
        return

    selected_items = self.daily_usage_tree.selection()
    if not selected_items:
        items_to_export = self.daily_usage_tree.get_children()
    else:
        items_to_export = selected_items

    if not items_to_export:
        messagebox.showinfo("알림", "내보낼 데이터가 없습니다.")
        return

    records = []
    cols = list(self.daily_usage_tree['columns'])

    def safe_float(val):
        try: return float(str(val).replace(',', '').strip())
        except: return 0.0

    def safe_int(val):
        try: return int(float(str(val).replace(',', '').strip()))
        except: return 0

    for item in items_to_export:
        tags = self.daily_usage_tree.item(item, 'tags')
        if tags and 'total' in tags: continue

        vals = self.daily_usage_tree.item(item, 'values')
        if not vals: continue

        row = dict(zip(cols, vals))

        records.append({
            "date": row.get("날짜", ""),
            "loc": f"{row.get('현장', '')} / {row.get('검사품명', '')}".strip(" /"),
            "ndt_type": row.get("검사방법", ""),
            "work_time": row.get("작업형태", "일반") or "일반",
            "material_type": row.get("품목명", ""),
            "qty": safe_float(row.get("수량", 0)),
            "unit": row.get("단위", ""),
            "corr": safe_float(row.get("보정계수", 1)),
            "adjusted_qty": safe_float(row.get("환산물량", 0)),
            "mat_cost": safe_int(row.get("재료비", 0)),
            "lab_cost": safe_int(row.get("인건비", 0)),
            "overhead": safe_int(row.get("제경비", 0)),
            "tech": safe_int(row.get("기술료", 0)),
            "subtotal": safe_int(row.get("검사비", 0))
        })

    if not records:
        messagebox.showwarning("기록 없음", "내보낼 작업 기록이 없습니다.")
        return

    top = tk.Toplevel(self.root)
    top.title("실비 정산 및 기성청구서 출력")
    top.geometry("350x400")
    top.transient(self.root)
    top.grab_set()

    ttk.Label(top, text="엑셀 청구서 하단에 합산될\n추가 실비정산 금액을 입력하세요.\n(세액 미포함 금액)", justify=tk.CENTER).pack(pady=15)

    exp_vars = {}
    for label in ["장비손료", "안전관리비", "주재비 및 출장여비", "도서인쇄비"]:
        f = ttk.Frame(top)
        f.pack(fill=tk.X, padx=20, pady=5)
        ttk.Label(f, text=label, width=15).pack(side=tk.LEFT)
        var = tk.IntVar(value=0)
        ttk.Entry(f, textvariable=var, justify='right').pack(side=tk.RIGHT, expand=True, fill=tk.X)
        exp_vars[label] = var

    def do_export():
        extra_expenses = {k: v.get() for k, v in exp_vars.items()}
        top.destroy()

        default_name = f"기성청구내역서_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=default_name, filetypes=[("Excel File", "*.xlsx")], title="정식 기성청구 엑셀 양식으로 저장")
        if not filepath: return

        excel = None
        try:
            excel = win32.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            try:
                excel.ScreenUpdating = False
                excel.EnableEvents = False
            except Exception:
                pass
            wb = excel.Workbooks.Add()
            try:
                excel.Calculation = -4135  # xlCalculationManual
            except Exception:
                pass
            ws = wb.ActiveSheet
            ws.Name = "기성청구내역서"

            ws.Range("A1:O2").Merge()
            ws.Range("A1").Value = "비파괴검사기술용역 기성청구 내역서"
            ws.Range("A1").Font.Size = 20
            ws.Range("A1").Font.Bold = True
            ws.Range("A1").HorizontalAlignment = -4108
            ws.Range("A1").VerticalAlignment = -4108

            ws.Range("A4:B4").Merge()
            ws.Range("A4").Value = "공 사 명 :"
            ws.Range("A4").Font.Bold = True
            ws.Range("C4:J4").Merge()
            project_name = records[0]["loc"].split('/')[0].strip() if records else ""
            if not project_name: project_name = "가산~가평 천연가스 공급시설 건설공사 비파괴검사 기술용역"
            ws.Range("C4").Value = project_name

            ws.Range("L4:M4").Merge()
            ws.Range("L4").Value = "청구일자 :"
            ws.Range("L4").Font.Bold = True
            ws.Range("L4").HorizontalAlignment = -4152
            ws.Range("N4:O4").Merge()
            ws.Range("N4").Value = datetime.datetime.now().strftime('%Y년 %m월 %d일')

            headers = ["No.", "검사일자", "작업구간", "검사종류", "규격/자재", "근무형태", "실물량", "단위", "보정계수", "환산물량", 
                       "재료비", "직접인건비", "제경비", "기술료", "공급가액소계"]

            start_row = 6
            for col, h in enumerate(headers, start=1):
                cell = ws.Cells(start_row, col)
                cell.Value = h
                cell.Font.Bold = True
                cell.Interior.Color = 14277081
                cell.HorizontalAlignment = -4108
                cell.Borders.LineStyle = 1

            widths = [4, 11, 20, 9, 22, 9, 8, 5, 9, 9, 11, 12, 11, 11, 13]
            for idx, w in enumerate(widths, start=1):
                ws.Columns(idx).ColumnWidth = w

            current_row = start_row + 1
            total_mat = total_lab = total_ovr = total_tech = total_sub = 0
            item_idx = 1

            for g_type in ["RT", "UT", "PT", "MT", "ETC"]:
                if g_type == "ETC":
                    group_records = [r for r in records if not any(r["ndt_type"].upper().startswith(t) for t in ["RT", "UT", "PT", "MT"])]
                else:
                    group_records = [r for r in records if r["ndt_type"].upper().startswith(g_type)]

                if not group_records: continue

                sub_mat = sub_lab = sub_ovr = sub_tech = sub_sub = 0
                grp_start_row = current_row
                grp_rows_batch = []
                for r in group_records:
                    loc_disp = r["loc"].split()[0] if isinstance(r["loc"], str) and r["loc"] else r["loc"]
                    grp_rows_batch.append((
                        item_idx, r["date"], loc_disp, r["ndt_type"], r["material_type"],
                        r["work_time"], r["qty"], r["unit"], r["corr"], r["adjusted_qty"],
                        r["mat_cost"], r["lab_cost"], r["overhead"], r["tech"], r["subtotal"]
                    ))
                    sub_mat += r["mat_cost"]; sub_lab += r["lab_cost"]; sub_ovr += r["overhead"]
                    sub_tech += r["tech"]; sub_sub += r["subtotal"]
                    item_idx += 1
                grp_end_row = grp_start_row + len(grp_rows_batch) - 1
                grp_rng = ws.Range(ws.Cells(grp_start_row, 1), ws.Cells(grp_end_row, 15))
                grp_rng.Value = grp_rows_batch
                grp_rng.Borders.LineStyle = 1
                ws.Range(ws.Cells(grp_start_row, 1), ws.Cells(grp_end_row, 2)).HorizontalAlignment = -4108
                ws.Range(ws.Cells(grp_start_row, 3), ws.Cells(grp_end_row, 3)).HorizontalAlignment = -4131
                ws.Range(ws.Cells(grp_start_row, 4), ws.Cells(grp_end_row, 4)).HorizontalAlignment = -4108
                ws.Range(ws.Cells(grp_start_row, 5), ws.Cells(grp_end_row, 5)).HorizontalAlignment = -4131
                ws.Range(ws.Cells(grp_start_row, 6), ws.Cells(grp_end_row, 6)).HorizontalAlignment = -4108
                ws.Range(ws.Cells(grp_start_row, 7), ws.Cells(grp_end_row, 7)).NumberFormat = "0.00"
                ws.Range(ws.Cells(grp_start_row, 8), ws.Cells(grp_end_row, 8)).HorizontalAlignment = -4108
                ws.Range(ws.Cells(grp_start_row, 9), ws.Cells(grp_end_row, 10)).NumberFormat = "0.00"
                ws.Range(ws.Cells(grp_start_row, 11), ws.Cells(grp_end_row, 15)).NumberFormat = "#,##0"
                current_row = grp_end_row + 1

                ws.Range(ws.Cells(current_row, 1), ws.Cells(current_row, 10)).Merge()
                ws.Cells(current_row, 1).Value = f"[{g_type}] 검사 소계"
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

            current_row += 1
            total_extra = 0
            for name, val in extra_expenses.items():
                if val > 0:
                    ws.Range(ws.Cells(current_row, 1), ws.Cells(current_row, 14)).Merge()
                    ws.Cells(current_row, 1).Value = f"+ {name}"
                    ws.Cells(current_row, 1).HorizontalAlignment = -4152
                    ws.Cells(current_row, 15).Value = val
                    ws.Cells(current_row, 15).NumberFormat = "#,##0"
                    for c in range(1, 16): ws.Cells(current_row, c).Borders.LineStyle = 1
                    total_extra += val
                    current_row += 1

            grand_subtotal = total_sub + total_extra

            ws.Range(ws.Cells(current_row, 1), ws.Cells(current_row, 14)).Merge()
            ws.Cells(current_row, 1).Value = "공급가액 총액"
            ws.Cells(current_row, 1).HorizontalAlignment = -4152
            ws.Cells(current_row, 1).Font.Bold = True
            ws.Cells(current_row, 15).Value = grand_subtotal
            ws.Cells(current_row, 15).NumberFormat = "#,##0"
            ws.Cells(current_row, 15).Font.Bold = True
            for c in range(1, 16): ws.Cells(current_row, c).Borders.LineStyle = 1

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

            try:
                excel.Calculation = -4105  # xlCalculationAutomatic
                excel.Calculate()
                excel.ScreenUpdating = True
                excel.EnableEvents = True
            except Exception:
                pass
            filepath = filepath.replace("/", "\\")
            wb.SaveAs(filepath)
            wb.Close()
            excel.Quit()
            excel = None

            messagebox.showinfo("저장 완료", f"엑셀 기성청구 내역서가 성공적으로 생성되었습니다.\n{filepath}")
            os.startfile(filepath)

        except Exception as e:
            messagebox.showerror("저장 오류", f"엑셀 파일 생성 중 오류가 발생했습니다.\n{str(e)}")
            try: excel.Quit()
            except: pass

    ttk.Button(top, text="엑셀 출력 실행", command=do_export).pack(pady=20, ipady=5, fill=tk.X, padx=20)
