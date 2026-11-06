import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import re
import json
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    from site_apps.central.src.data_sync import get_history_path
except ImportError:
    def get_history_path():
        p = r"F:\내 드라이브\07_Antigravity\PMI_Data\central\daily_work_history.json"
        if os.path.exists(p): return p
        return os.path.abspath(os.path.join(os.path.dirname(__file__), "daily_work_history.json"))

try:
    from site_apps.central.src.services.ndt_calculator import calculate_billing
except ImportError:
    def calculate_billing(*, quantity, adjusted_quantity, material_key, ndt_type, location, work_time, material_costs, labor_costs, overhead_rate, technical_fee_rate, vat_rate=0.1):
        mat_unit = material_costs.get(material_key, material_costs.get(ndt_type, 0))
        time_costs = labor_costs.get(location, labor_costs).get(work_time, {})
        lab_unit = time_costs.get(material_key, time_costs.get(ndt_type, 0))
        u_mat = int(float(mat_unit or 0))
        u_lab = int(float(lab_unit or 0))
        u_overhead = int(u_lab * overhead_rate)
        u_tech = int((u_lab + u_overhead) * technical_fee_rate)
        unit_price = u_mat + u_lab + u_overhead + u_tech
        subtotal = int(round(quantity * unit_price))
        vat = int(round(subtotal * vat_rate))
        return {"unit_price": unit_price, "subtotal": subtotal, "vat": vat, "total_amount": subtotal + vat}


class SectionLineBillingDialog(tk.Toplevel):
    def __init__(self, parent=None, history=None, config=None):
        super().__init__(parent)
        self.title("구간 / 라인별 기성청구 내역서 및 미리보기")
        self.geometry("1160x720")
        self.minsize(960, 560)
        
        self.history = history
        self.config = config
        self._load_data()
        
        self.exclude_reinspect_var = tk.BooleanVar(value=True)
        self.round_var = tk.StringVar(value="1")
        
        self.sections_map = {} # sec -> set(lines)
        self.section_vars = {} # sec -> BooleanVar
        self.line_vars = {}    # line -> BooleanVar
        self.line_sec_map = {} # line -> sec
        
        self._setup_ui()
        self._populate_sections()
        
    def _load_data(self):
        if self.history is None:
            try:
                hist_path = get_history_path()
                if os.path.exists(hist_path):
                    with open(hist_path, 'r', encoding='utf-8') as f:
                        self.history = json.load(f)
                else:
                    self.history = {}
            except Exception as e:
                print("Error loading history:", e)
                self.history = {}
                
        if self.config is None:
            cfg_path = os.path.join(os.path.dirname(__file__), "config.json")
            if os.path.exists(cfg_path):
                try:
                    with open(cfg_path, 'r', encoding='utf-8') as f:
                        self.config = json.load(f)
                except Exception:
                    self.config = {}
            else:
                self.config = {}

    def _setup_ui(self):
        main_frame = ttk.Frame(self, padding=8)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # -------------------------------------------------------------
        # 1. Top Filter Frame (고정 상단)
        # -------------------------------------------------------------
        filter_frame = ttk.LabelFrame(main_frame, text="■ 구간 및 정산 옵션 설정 (다중 구간/라인 선택)", padding=6)
        filter_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 6))
        
        # Row 1: 기성 회차, 제외 옵션, 새로고침 버튼
        row1 = ttk.Frame(filter_frame)
        row1.pack(fill=tk.X, pady=(0, 4))
        
        ttk.Label(row1, text="기성 회차 :", font=("맑은 고딕", 9, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Entry(row1, textvariable=self.round_var, width=5, justify="center").pack(side=tk.LEFT, padx=(0, 2))
        ttk.Label(row1, text="회").pack(side=tk.LEFT, padx=(0, 20))
        
        self.chk_exclude_re = ttk.Checkbutton(
            row1, 
            text="시공사 귀책 수정재검 물량 제외 (권장)", 
            variable=self.exclude_reinspect_var,
            command=self.update_preview
        )
        self.chk_exclude_re.pack(side=tk.LEFT, padx=(0, 15))
        
        btn_refresh = ttk.Button(row1, text="🔄 데이터 새로고침", command=self.refresh_data, width=15)
        btn_refresh.pack(side=tk.RIGHT)
        
        # Row 2: 대상 구간(섹션) 다중 선택 헤더
        row_sec_hdr = ttk.Frame(filter_frame)
        row_sec_hdr.pack(fill=tk.X, pady=(2, 1))
        
        ttk.Label(row_sec_hdr, text="대상 구간(섹션) 선택 :", font=("맑은 고딕", 9, "bold")).pack(side=tk.LEFT, padx=(0, 6))
        self.chk_all_secs_var = tk.BooleanVar(value=False)
        self.chk_all_secs = ttk.Checkbutton(row_sec_hdr, text="[ 구간 전체 선택 ]", variable=self.chk_all_secs_var, command=self._toggle_all_sections)
        self.chk_all_secs.pack(side=tk.LEFT, padx=(0, 12))
        
        self.sec_summary_lbl = ttk.Label(row_sec_hdr, text="", font=("맑은 고딕", 9), foreground="#2B579A")
        self.sec_summary_lbl.pack(side=tk.LEFT)
        
        # Sections Container Frame
        self.sections_container = ttk.Frame(filter_frame)
        self.sections_container.pack(fill=tk.X, pady=(1, 4))
        
        # Row 3: 대상 라인 선택 헤더
        row_line_hdr = ttk.Frame(filter_frame)
        row_line_hdr.pack(fill=tk.X, pady=(2, 1))
        
        ttk.Label(row_line_hdr, text="대상 라인 선택 :", font=("맑은 고딕", 9, "bold")).pack(side=tk.LEFT, padx=(0, 6))
        self.chk_all_lines_var = tk.BooleanVar(value=True)
        self.chk_all_lines = ttk.Checkbutton(row_line_hdr, text="[ 라인 전체 선택 ]", variable=self.chk_all_lines_var, command=self._toggle_all_lines)
        self.chk_all_lines.pack(side=tk.LEFT, padx=(0, 10))
        
        self.line_count_lbl = ttk.Label(row_line_hdr, text="", font=("맑은 고딕", 9), foreground="#555555")
        self.line_count_lbl.pack(side=tk.LEFT)
        
        # Lines Scrollable Canvas Frame (라인이 많아져도 상단 높이를 일정하게 유지)
        self.lines_scroll_frame = ttk.Frame(filter_frame)
        self.lines_scroll_frame.pack(fill=tk.X, pady=(1, 1))
        
        self.lines_canvas = tk.Canvas(self.lines_scroll_frame, height=68, highlightthickness=1, highlightbackground="#D3D3D3", bg="#FAFAFA")
        self.lines_scrollbar = ttk.Scrollbar(self.lines_scroll_frame, orient="vertical", command=self.lines_canvas.yview)
        self.lines_container = ttk.Frame(self.lines_canvas)
        
        def _on_inner_config(e):
            self.lines_canvas.configure(scrollregion=self.lines_canvas.bbox("all"))
        self.lines_container.bind("<Configure>", _on_inner_config)
        self.lines_canvas.create_window((0, 0), window=self.lines_container, anchor="nw")
        self.lines_canvas.configure(yscrollcommand=self.lines_scrollbar.set)
        
        def _on_wheel(e):
            self.lines_canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
        self.lines_canvas.bind("<MouseWheel>", _on_wheel)
        self.lines_container.bind("<MouseWheel>", _on_wheel)
        
        self.lines_canvas.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.lines_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # -------------------------------------------------------------
        # 2. Bottom Fixed Area (화면 맨 하단에 무조건 고정 표시)
        # -------------------------------------------------------------
        bottom_area = ttk.Frame(main_frame)
        bottom_area.pack(side=tk.BOTTOM, fill=tk.X, pady=(6, 0))
        
        # Action Buttons
        btn_frame = ttk.Frame(bottom_area)
        btn_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(4, 0))
        
        btn_close = ttk.Button(btn_frame, text="닫  기", command=self.destroy, width=12)
        btn_close.pack(side=tk.LEFT)
        
        btn_pdf = ttk.Button(btn_frame, text="📄 PDF 바로 발행", command=self.export_pdf, width=18)
        btn_pdf.pack(side=tk.RIGHT, padx=(6, 0))
        
        btn_excel = ttk.Button(btn_frame, text="📋 정식 기성청구 엑셀 양식 저장", command=self.export_excel, width=30)
        btn_excel.pack(side=tk.RIGHT, padx=(6, 0))
        
        # Summary Info Bar
        summary_frame = ttk.Frame(bottom_area)
        summary_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(2, 2))
        
        self.status_lbl_var = tk.StringVar(value="조회 준비 중...")
        lbl_status = ttk.Label(summary_frame, textvariable=self.status_lbl_var, font=("맑은 고딕", 9), foreground="#555555")
        lbl_status.pack(side=tk.LEFT)
        
        self.totals_lbl_var = tk.StringVar(value="금회 공급가액: 0 원  |  부가세: 0 원  |  합계금액(VAT포함): 0 원")
        lbl_totals = ttk.Label(summary_frame, textvariable=self.totals_lbl_var, font=("맑은 고딕", 10, "bold"), foreground="#1F4E78")
        lbl_totals.pack(side=tk.RIGHT)
        
        # -------------------------------------------------------------
        # 3. Preview Grid Frame (중간 남는 공간 100% 차지)
        # -------------------------------------------------------------
        preview_frame = ttk.LabelFrame(main_frame, text="■ 정식 기성청구 내역서 미리보기 (Official Billing Claim Preview)", padding=6)
        preview_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        
        cols = (
            "cat", "spec", "unit", "price", 
            "c_qty", "c_amt", 
            "p_qty", "p_amt", 
            "cur_qty", "cur_amt", 
            "tot_qty", "tot_amt", 
            "rem_qty", "rem_amt"
        )
        
        self.tree = ttk.Treeview(preview_frame, columns=cols, show="headings")
        
        headers = [
            ("cat", "공  종", 130, "w"),
            ("spec", "규  격", 160, "w"),
            ("unit", "단위", 50, "center"),
            ("price", "단가(원)", 85, "e"),
            ("c_qty", "계약수량", 75, "e"),
            ("c_amt", "계약금액(원)", 95, "e"),
            ("p_qty", "전회수량", 70, "e"),
            ("p_amt", "전회금액(원)", 90, "e"),
            ("cur_qty", "금회수량", 80, "e"),
            ("cur_amt", "금회기성액(원)", 105, "e"),
            ("tot_qty", "누계수량", 75, "e"),
            ("tot_amt", "누계금액(원)", 100, "e"),
            ("rem_qty", "잔여수량", 75, "e"),
            ("rem_amt", "잔여금액(원)", 100, "e"),
        ]
        
        for col_id, col_name, w, anch in headers:
            self.tree.heading(col_id, text=col_name, anchor="center")
            self.tree.column(col_id, width=w, anchor=anch)
            
        tree_scroll_y = ttk.Scrollbar(preview_frame, orient="vertical", command=self.tree.yview)
        tree_scroll_x = ttk.Scrollbar(preview_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)
        
        tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.tree.tag_configure("total_row", background="#FFF2CC", font=("맑은 고딕", 10, "bold"))
        self.tree.tag_configure("subtotal_row", background="#EBF1F5", font=("맑은 고딕", 10, "bold"))
        self.tree.tag_configure("cur_highlight", background="#F2F5FB")

    def _populate_sections(self):
        sections_map = {}
        self.line_sec_map = {}
        for dt, day_data in self.history.items():
            for r in day_data.get('ndt_results', []):
                sec = str(r.get('구간', '')).strip()
                loc = str(r.get('위치', '')).strip()
                line = str(r.get('라인번호', '')).strip()
                if not sec and not line: continue
                
                base_sec = sec
                if "15" in sec or "015" in line:
                    base_sec = "SEC#015 (한국노총)"
                elif "Sec." in sec:
                    base_sec = sec.split('-')[0]
                    
                if base_sec not in sections_map:
                    sections_map[base_sec] = set()
                if line:
                    sections_map[base_sec].add(line)
                    self.line_sec_map[line] = base_sec
                    
        sorted_secs = sorted(sections_map.keys())
        self.sections_map = sections_map
        
        for widget in self.sections_container.winfo_children():
            widget.destroy()
            
        self.section_vars = {}
        if not sorted_secs:
            ttk.Label(self.sections_container, text="(등록된 구간 데이터가 없습니다)", foreground="#888888").pack(side=tk.LEFT)
            self._on_sections_changed()
            return
            
        row_frame = ttk.Frame(self.sections_container)
        row_frame.pack(fill=tk.X, pady=1)
        
        # 기본 선택: SEC#015 (한국노총) 또는 첫 번째 구간
        default_sec = "SEC#015 (한국노총)" if "SEC#015 (한국노총)" in sorted_secs else sorted_secs[0]
        
        for idx, sec in enumerate(sorted_secs):
            is_checked = (sec == default_sec)
            var = tk.BooleanVar(value=is_checked)
            self.section_vars[sec] = var
            chk = ttk.Checkbutton(row_frame, text=sec, variable=var, command=self._on_section_toggled)
            chk.pack(side=tk.LEFT, padx=(0, 14))
            if (idx + 1) % 6 == 0 and (idx + 1) < len(sorted_secs):
                row_frame = ttk.Frame(self.sections_container)
                row_frame.pack(fill=tk.X, pady=1)
                
        self.chk_all_secs_var.set(len(sorted_secs) == 1)
        self._on_sections_changed()

    def _toggle_all_sections(self):
        val = self.chk_all_secs_var.get()
        for v in self.section_vars.values():
            v.set(val)
        self._on_sections_changed()

    def _on_section_toggled(self):
        all_checked = all(v.get() for v in self.section_vars.values()) if self.section_vars else False
        self.chk_all_secs_var.set(all_checked)
        self._on_sections_changed()

    def get_selected_sections(self):
        return [s for s, v in self.section_vars.items() if v.get()]

    def get_section_display_name(self, for_filename=False):
        selected = self.get_selected_sections()
        if not selected:
            return "미선택"
        all_secs = list(self.sections_map.keys())
        if len(selected) == len(all_secs) and len(all_secs) > 1:
            return "전체구간"
        if len(selected) == 1:
            name = selected[0]
            if for_filename:
                return name.replace("#", "").replace(" ", "_")
            return name
        elif len(selected) == 2:
            if for_filename:
                s0 = selected[0].replace("#", "").replace(" ", "_")
                s1 = selected[1].replace("#", "").replace(" ", "_")
                return f"{s0}_{s1}"
            return f"{selected[0]}, {selected[1]}"
        else:
            s0 = selected[0].replace("#", "").replace(" ", "_") if for_filename else selected[0]
            if for_filename:
                return f"{s0}_외{len(selected)-1}개구간"
            return f"{s0} 외 {len(selected)-1}개 구간"

    def refresh_data(self):
        """최신 daily_work_history.json 데이터를 다시 읽어서 구간 및 라인 목록 갱신"""
        self.history = None
        self._load_data()
        self._populate_sections()

    def _on_sections_changed(self):
        selected_secs = self.get_selected_sections()
        
        # Update summary label
        disp_name = self.get_section_display_name()
        self.sec_summary_lbl.config(text=f"(선택: {disp_name})")
        
        # Collect lines for all selected sections
        active_lines = []
        for sec in selected_secs:
            lines = sorted(list(self.sections_map.get(sec, [])))
            for line in lines:
                active_lines.append((line, sec))
                
        # Preserve previous selections
        prev_selections = {l: v.get() for l, v in self.line_vars.items()}
        
        for widget in self.lines_container.winfo_children():
            widget.destroy()
            
        self.line_vars = {}
        if not active_lines:
            lbl = ttk.Label(self.lines_container, text="(선택된 구간이 없거나 등록된 라인이 없습니다)", foreground="#888888")
            lbl.pack(side=tk.LEFT, pady=2)
            self.chk_all_lines_var.set(False)
            self.line_count_lbl.config(text="(0개)")
            self.update_preview()
            return
            
        self.line_count_lbl.config(text=f"(총 {len(active_lines)}개 라인, 상하 스크롤 가능)")
        
        row_frame = ttk.Frame(self.lines_container)
        row_frame.pack(fill=tk.X, pady=1)
        
        show_sec_tag = len(selected_secs) > 1
        for idx, (line, sec) in enumerate(active_lines):
            val = prev_selections.get(line, True)
            var = tk.BooleanVar(value=val)
            self.line_vars[line] = var
            disp_text = f"{line} ({sec})" if show_sec_tag else line
            chk = ttk.Checkbutton(row_frame, text=disp_text, variable=var, command=self._on_line_toggled)
            chk.pack(side=tk.LEFT, padx=(0, 12))
            if (idx + 1) % 4 == 0 and (idx + 1) < len(active_lines):
                row_frame = ttk.Frame(self.lines_container)
                row_frame.pack(fill=tk.X, pady=1)
                
        all_lines_checked = all(v.get() for v in self.line_vars.values()) if self.line_vars else True
        self.chk_all_lines_var.set(all_lines_checked)
        
        # 스크롤 위치 초기화
        try:
            self.lines_canvas.yview_moveto(0)
        except Exception:
            pass
            
        self.update_preview()

    def _toggle_all_lines(self):
        val = self.chk_all_lines_var.get()
        for v in self.line_vars.values():
            v.set(val)
        self.update_preview()

    def _on_line_toggled(self):
        all_checked = all(v.get() for v in self.line_vars.values()) if self.line_vars else False
        self.chk_all_lines_var.set(all_checked)
        self.update_preview()

    def get_selected_lines(self):
        return [l for l, v in self.line_vars.items() if v.get()]

    def calculate_billing_data(self):
        selected_lines = set(self.get_selected_lines())
        exclude_re = self.exclude_reinspect_var.get()
        
        mat_costs = self.config.get("MATERIAL_COST", {
            "PAUT_300A 이상": 37559, "PAUT_250A": 37559, "PAUT_200A": 37559,
            "PAUT_150A-125A": 38229, "PAUT_100A 이하": 38229,
            "RT_3 1/3 x 12\"": 9274, "RT_3 1/3 x 6\"": 8044, "MT_MT": 411, "PT_PT": 1177
        })
        lab_costs = self.config.get("LABOR_COST", {
            "열배관": {
                "일반": {
                    "PAUT_300A 이상": 65200, "PAUT_250A": 76163, "PAUT_200A": 87348, "PAUT_150A-125A": 98560, "PAUT_100A 이하": 92362,
                    "RT_3 1/3 x 12\"": 51885, "RT_3 1/3 x 6\"": 51885, "MT_MT": 29771, "PT_PT": 30847
                },
                "야간": {
                    "PAUT_300A 이상": 97800, "PAUT_250A": 114243, "PAUT_200A": 131021, "PAUT_150A-125A": 147841, "PAUT_100A 이하": 138544,
                    "RT_3 1/3 x 12\"": 77833, "RT_3 1/3 x 6\"": 77833, "MT_MT": 44659, "PT_PT": 46272
                }
            }
        })
        contract_qtys = self.config.get("CONTRACT_QTY", {})
        
        # 1. First scan all records of selected lines
        total_joints = 0
        excluded_re_count = 0
        line_records = []
        
        for dt in sorted(self.history.keys()):
            for r in self.history[dt].get('ndt_results', []):
                line = str(r.get('라인번호', '')).strip()
                if line not in selected_lines:
                    continue
                    
                total_joints += 1
                j_no = str(r.get('Joint No.', '')).strip()
                is_re = bool(r.get('원본JointNo') or (len(j_no) > 2 and j_no.endswith('R')))
                
                method = str(r.get('검사방법', '')).strip().upper()
                shift = str(r.get('근무구분', '')).strip()
                work_time = '야간' if '야간' in shift else '일반'
                size_str = str(r.get('관경', '')).strip()
                m_match = re.search(r'\d+', size_str)
                size_val = int(m_match.group()) if m_match else 0
                sec_val = str(r.get('구간', '')).strip()
                res_val = str(r.get('결과', '')).strip()
                
                # Determine qty & unit
                if method in ('PAUT', 'UT'):
                    unit = 'M'
                    raw_q = str(r.get('PAUT', '')).strip().replace(',', '')
                    qty = float(raw_q) if raw_q else (0.8401 if size_val >= 250 else 0.3591)
                    size_group = ('300A 이상' if size_val >= 300 else '250A' if size_val == 250 else
                                  '200A' if size_val == 200 else '150A-125A' if 125 <= size_val <= 150 else '100A 이하')
                    mat_key = 'PAUT_' + size_group
                    disp_spec = f"{size_str}[PAUT,{work_time}]" if size_str else f"PAUT {size_group} ({work_time})"
                elif method == 'RT':
                    unit = '매'
                    raw_q = str(r.get('RT_OR', '')).strip()
                    qty = float(raw_q) if raw_q else 3.0
                    spec_str = str(r.get('규격', '')).strip()
                    size_group = '3 1/3 x 6"' if '6' in spec_str else '3 1/3 x 12"'
                    mat_key = 'RT_' + size_group
                    disp_spec = f"{size_str}[RT,{work_time}]" if size_str else f"RT ({work_time})"
                elif method in ('MT', 'PT'):
                    unit = 'M'
                    raw_q = str(r.get(method, '')).strip().replace(',', '')
                    qty = float(raw_q) if raw_q else 1.0
                    mat_key = f'{method}_{method}'
                    disp_spec = f"{method}[{work_time}]"
                else:
                    continue
                    
                if is_re and exclude_re:
                    excluded_re_count += int(qty)
                    line_records.append({
                        "date": dt, "sec": sec_val, "line": line, "method": method, "size": size_str,
                        "shift": work_time, "joint": j_no, "qty": 0.0, "qty_raw": qty, "unit": unit,
                        "res": res_val, "is_re_excluded": True, "note": "수정재검 (기성 제외)"
                    })
                    continue
                    
                line_records.append({
                    "date": dt, "sec": sec_val, "line": line, "method": method, "size": size_str,
                    "shift": work_time, "joint": j_no, "qty": qty, "qty_raw": qty, "unit": unit,
                    "res": res_val, "is_re_excluded": False, "note": "정상 검사"
                })

        # 2. Dynamically aggregate into Billing Items
        aggregated = {}
        for r in line_records:
            if r["is_re_excluded"]:
                continue
            method = r["method"]
            work_time = r["shift"]
            size_str = r["size"]
            m_match = re.search(r'\d+', size_str)
            size_val = int(m_match.group()) if m_match else 0
            
            if method in ('PAUT', 'UT'):
                size_group = ('300A 이상' if size_val >= 300 else '250A' if size_val == 250 else
                              '200A' if size_val == 200 else '150A-125A' if 125 <= size_val <= 150 else '100A 이하')
                mat_key = 'PAUT_' + size_group
                disp_spec = f"{size_str}[PAUT,{work_time}]" if size_str else f"PAUT {size_group} ({work_time})"
            elif method == 'RT':
                spec_str = r.get('spec', '')
                size_group = '3 1/3 x 6"' if '6' in spec_str else '3 1/3 x 12"'
                mat_key = 'RT_' + size_group
                disp_spec = f"{size_str}[RT,{work_time}]" if size_str else f"RT ({work_time})"
            else:
                mat_key = f'{method}_{method}'
                disp_spec = f"{method}[{work_time}]"
                
            item_key = (method, mat_key, work_time, disp_spec, r["unit"])
            if item_key not in aggregated:
                aggregated[item_key] = 0.0
            aggregated[item_key] += r["qty"]
            
        bill_items = []
        for (method, mat_key, work_time, disp_spec, unit), cur_qty in sorted(aggregated.items(), key=lambda x: (x[0][0], x[0][2])):
            calc = calculate_billing(
                quantity=1.0, adjusted_quantity=1.0, material_key=mat_key, ndt_type=method,
                location='열배관', work_time=work_time, material_costs=mat_costs, labor_costs=lab_costs,
                overhead_rate=1.1, technical_fee_rate=0.2
            )
            price = calc['unit_price']
            c_k = f'{mat_key}_야간' if work_time == '야간' else mat_key
            c_qty = float(contract_qtys.get(c_k, 0))
            c_amt = int(round(c_qty * price))
            p_qty = 0.0
            p_amt = 0
            cur_amt = int(round(cur_qty * price))
            tot_qty = p_qty + cur_qty
            tot_amt = p_amt + cur_amt
            rem_qty = c_qty - tot_qty
            rem_amt = c_amt - tot_amt
            
            bill_items.append({
                "cat": "비파괴검사용역비",
                "spec": disp_spec,
                "unit": unit,
                "price": price,
                "c_qty": c_qty,
                "c_amt": c_amt,
                "p_qty": p_qty,
                "p_amt": p_amt,
                "cur_qty": cur_qty,
                "cur_amt": cur_amt,
                "tot_qty": tot_qty,
                "tot_amt": tot_amt,
                "rem_qty": rem_qty,
                "rem_amt": rem_amt,
                "method": method,
                "mat_key": mat_key,
                "work_time": work_time,
            })
            
        cur_supply = sum(item["cur_amt"] for item in bill_items)
        cur_vat = int(round(cur_supply * 0.1))
        cur_total = cur_supply + cur_vat
        
        return {
            "bill_items": bill_items,
            "total_joints": total_joints,
            "excluded_re_count": excluded_re_count,
            "cur_supply": cur_supply,
            "cur_vat": cur_vat,
            "cur_total": cur_total,
            "line_records": line_records,
            "selected_lines": list(selected_lines)
        }

    def update_preview(self):
        data = self.calculate_billing_data()
        
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        for item in data["bill_items"]:
            cur_qty_str = f"{item['cur_qty']:.4f}" if item["unit"] == "M" else f"{item['cur_qty']:.1f}"
            c_qty_str = f"{item['c_qty']:.4f}" if item["unit"] == "M" else f"{item['c_qty']:.1f}"
            p_qty_str = f"{item['p_qty']:.4f}" if item["unit"] == "M" else f"{item['p_qty']:.1f}"
            tot_qty_str = f"{item['tot_qty']:.4f}" if item["unit"] == "M" else f"{item['tot_qty']:.1f}"
            rem_qty_str = f"{item['rem_qty']:.4f}" if item["unit"] == "M" else f"{item['rem_qty']:.1f}"
            
            row_vals = (
                item["cat"],
                item["spec"],
                item["unit"],
                f"{item['price']:,}",
                c_qty_str,
                f"{item['c_amt']:,}",
                p_qty_str,
                f"{item['p_amt']:,}",
                cur_qty_str,
                f"{item['cur_amt']:,}",
                tot_qty_str,
                f"{item['tot_amt']:,}",
                rem_qty_str,
                f"{item['rem_amt']:,}"
            )
            self.tree.insert("", "end", values=row_vals, tags=("cur_highlight",))
            
        # Summary Row: 공급가액
        c_amt_tot = sum(i["c_amt"] for i in data["bill_items"])
        p_amt_tot = sum(i["p_amt"] for i in data["bill_items"])
        cur_amt_tot = data["cur_supply"]
        tot_amt_tot = sum(i["tot_amt"] for i in data["bill_items"])
        rem_amt_tot = sum(i["rem_amt"] for i in data["bill_items"])
        
        row_supply = (
            "공  급  가  액", "", "", "", "",
            f"{c_amt_tot:,}", "", f"{p_amt_tot:,}", "", f"{cur_amt_tot:,}", "", f"{tot_amt_tot:,}", "", f"{rem_amt_tot:,}"
        )
        self.tree.insert("", "end", values=row_supply, tags=("subtotal_row",))
        
        # Summary Row: 부가가치세 (10%)
        vat_c = int(round(c_amt_tot * 0.1))
        vat_p = int(round(p_amt_tot * 0.1))
        vat_cur = data["cur_vat"]
        vat_tot = int(round(tot_amt_tot * 0.1))
        vat_rem = int(round(rem_amt_tot * 0.1))
        row_vat = (
            "부 가 가 치 세 (10%)", "", "", "", "",
            f"{vat_c:,}", "", f"{vat_p:,}", "", f"{vat_cur:,}", "", f"{vat_tot:,}", "", f"{vat_rem:,}"
        )
        self.tree.insert("", "end", values=row_vat, tags=("subtotal_row",))
        
        # Grand Total: 합계
        row_grand = (
            "합              계", "", "", "", "",
            f"{(c_amt_tot+vat_c):,}", "", f"{(p_amt_tot+vat_p):,}", "", f"{data['cur_total']:,}", "", f"{(tot_amt_tot+vat_tot):,}", "", f"{(rem_amt_tot+vat_rem):,}"
        )
        self.tree.insert("", "end", values=row_grand, tags=("total_row",))
        
        # Update Info Labels
        lines_cnt = len(data["selected_lines"])
        self.status_lbl_var.set(
            f"선택 라인: {lines_cnt}개 | 총 검사 조인트: {data['total_joints']}개 | 수정재검 제외: {data['excluded_re_count']}매 제외됨"
        )
        self.totals_lbl_var.set(
            f"금회 공급가액: {data['cur_supply']:,} 원  |  부가가치세: {data['cur_vat']:,} 원  |  합계금액(VAT포함): {data['cur_total']:,} 원"
        )

    def generate_excel_file(self, filepath):
        data = self.calculate_billing_data()
        rnd = self.round_var.get()
        sec = self.get_section_display_name()
        
        wb = openpyxl.Workbook()
        
        font_title = Font(name='맑은 고딕', size=16, bold=True, color='1F4E78')
        font_sub = Font(name='맑은 고딕', size=11, bold=True)
        font_hdr = Font(name='맑은 고딕', size=10, bold=True, color='2C3E50')
        font_cell = Font(name='맑은 고딕', size=9.5)
        font_cell_b = Font(name='맑은 고딕', size=9.5, bold=True)
        font_tot = Font(name='맑은 고딕', size=10.5, bold=True, color='1F4E78')
        
        fill_hdr = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
        fill_subtot = PatternFill(start_color='F2F4F7', end_color='F2F4F7', fill_type='solid')
        fill_tot = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
        
        thin = Side(style='thin', color='B0C4DE')
        thin_d = Side(style='thin', color='708090')
        double = Side(style='double', color='1F4E78')
        
        b_all = Border(left=thin, right=thin, top=thin, bottom=thin)
        b_hdr = Border(left=thin_d, right=thin_d, top=thin_d, bottom=thin_d)
        b_tot = Border(left=thin_d, right=thin_d, top=thin_d, bottom=double)
        
        # ---------------------------------------------------------
        # SHEET 1: 기성청구내역서
        # ---------------------------------------------------------
        ws1 = wb.active
        ws1.title = "기성청구내역서"
        ws1.page_setup.orientation = ws1.ORIENTATION_LANDSCAPE
        ws1.page_setup.paperSize = ws1.PAPERSIZE_A4
        ws1.page_setup.fitToWidth = 1
        ws1.page_setup.fitToHeight = 0
        ws1.sheet_properties.pageSetUpPr.fitToPage = True
        
        ws1.merge_cells('A1:N2')
        ws1['A1'] = f"제 {rnd} 회 비파괴검사기술용역 기성청구 내역서"
        ws1['A1'].font = font_title
        ws1['A1'].alignment = Alignment(horizontal='center', vertical='center')
        
        ws1['A3'] = f"공 사 명 : 2026년 중앙지사 열수송관 비파괴검사용역 단가계약 ({sec})"
        ws1['A3'].font = font_sub
        
        ws1.merge_cells('A5:A6'); ws1['A5'] = "공종"; ws1['A5'].alignment = Alignment(horizontal='center', vertical='center')
        ws1.merge_cells('B5:B6'); ws1['B5'] = "규격"; ws1['B5'].alignment = Alignment(horizontal='center', vertical='center')
        ws1.merge_cells('C5:C6'); ws1['C5'] = "단위"; ws1['C5'].alignment = Alignment(horizontal='center', vertical='center')
        ws1.merge_cells('D5:D6'); ws1['D5'] = "단가(원)"; ws1['D5'].alignment = Alignment(horizontal='center', vertical='center')
        
        groups = [("E5:F5", "계약", "E6", "수량", "F6", "금액"),
                  ("G5:H5", "전회까지기성", "G6", "수량", "H6", "금액"),
                  ("I5:J5", "금회기성", "I6", "수량", "J6", "금액"),
                  ("K5:L5", "누계기성", "K6", "수량", "L6", "금액"),
                  ("M5:N5", "잔여액", "M6", "수량", "N6", "금액")]
                  
        for rng, g_name, c_q, t_q, c_a, t_a in groups:
            ws1.merge_cells(rng)
            start_cell = rng.split(':')[0]
            ws1[start_cell] = g_name
            ws1[start_cell].alignment = Alignment(horizontal='center', vertical='center')
            ws1[c_q] = t_q; ws1[c_q].alignment = Alignment(horizontal='center', vertical='center')
            ws1[c_a] = t_a; ws1[c_a].alignment = Alignment(horizontal='center', vertical='center')
            
        for r in range(5, 7):
            for c in range(1, 15):
                cell = ws1.cell(r, c)
                cell.font = font_hdr
                cell.fill = fill_hdr
                cell.border = b_hdr
                
        r_idx = 7
        for item in data["bill_items"]:
            ws1.cell(r_idx, 1, item["cat"]).alignment = Alignment(horizontal='left', vertical='center')
            ws1.cell(r_idx, 2, item["spec"]).alignment = Alignment(horizontal='left', vertical='center')
            ws1.cell(r_idx, 3, item["unit"]).alignment = Alignment(horizontal='center', vertical='center')
            
            c_p = ws1.cell(r_idx, 4, item["price"])
            c_p.alignment = Alignment(horizontal='right', vertical='center')
            c_p.number_format = '#,##0'
            
            num_data = [
                (5, item["c_qty"], False), (6, item["c_amt"], True),
                (7, item["p_qty"], False), (8, item["p_amt"], True),
                (9, item["cur_qty"], False), (10, item["cur_amt"], True),
                (11, item["tot_qty"], False), (12, item["tot_amt"], True),
                (13, item["rem_qty"], False), (14, item["rem_amt"], True),
            ]
            for col_no, val, is_curr in num_data:
                cell = ws1.cell(r_idx, col_no, val)
                cell.alignment = Alignment(horizontal='right', vertical='center')
                if is_curr:
                    cell.number_format = '#,##0'
                else:
                    cell.number_format = '0.0000' if item["unit"] == "M" else '0.0'
                    
            for c in range(1, 15):
                ws1.cell(r_idx, c).font = font_cell
                ws1.cell(r_idx, c).border = b_all
            r_idx += 1
            
        # 공급가액
        ws1.merge_cells(f'A{r_idx}:D{r_idx}')
        ws1[f'A{r_idx}'] = "공  급  가  액"
        ws1[f'A{r_idx}'].alignment = Alignment(horizontal='center', vertical='center')
        
        for col_idx, col_let in [(6, 'F'), (8, 'H'), (10, 'J'), (12, 'L'), (14, 'N')]:
            cell = ws1.cell(r_idx, col_idx)
            cell.value = f"=SUM({col_let}7:{col_let}{r_idx-1})"
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.number_format = '#,##0'
            
        for c in range(1, 15):
            ws1.cell(r_idx, c).font = font_cell_b
            ws1.cell(r_idx, c).fill = fill_subtot
            ws1.cell(r_idx, c).border = b_hdr
        r_idx += 1
        
        # 부가가치세 (10%)
        ws1.merge_cells(f'A{r_idx}:D{r_idx}')
        ws1[f'A{r_idx}'] = "부 가 가 치 세 (10%)"
        ws1[f'A{r_idx}'].alignment = Alignment(horizontal='center', vertical='center')
        
        for col_idx, col_let in [(6, 'F'), (8, 'H'), (10, 'J'), (12, 'L'), (14, 'N')]:
            cell = ws1.cell(r_idx, col_idx)
            cell.value = f"=ROUND({col_let}{r_idx-1}*0.1, 0)"
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.number_format = '#,##0'
            
        for c in range(1, 15):
            ws1.cell(r_idx, c).font = font_cell_b
            ws1.cell(r_idx, c).fill = fill_subtot
            ws1.cell(r_idx, c).border = b_hdr
        r_idx += 1
        
        # 합계
        ws1.merge_cells(f'A{r_idx}:D{r_idx}')
        ws1[f'A{r_idx}'] = "합              계"
        ws1[f'A{r_idx}'].alignment = Alignment(horizontal='center', vertical='center')
        
        for col_idx, col_let in [(6, 'F'), (8, 'H'), (10, 'J'), (12, 'L'), (14, 'N')]:
            cell = ws1.cell(r_idx, col_idx)
            cell.value = f"={col_let}{r_idx-2}+{col_let}{r_idx-1}"
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.number_format = '#,##0'
            
        for c in range(1, 15):
            ws1.cell(r_idx, c).font = font_tot
            ws1.cell(r_idx, c).fill = fill_tot
            ws1.cell(r_idx, c).border = b_tot

        col_w = {'A': 16, 'B': 22, 'C': 7, 'D': 12, 'E': 10, 'F': 14, 'G': 10, 'H': 14, 'I': 10, 'J': 15, 'K': 10, 'L': 15, 'M': 10, 'N': 15}
        for col, w in col_w.items():
            ws1.column_dimensions[col].width = w

        # ---------------------------------------------------------
        # SHEET 2: 라인별_총괄비교표 (Dynamically built per line)
        # ---------------------------------------------------------
        ws2 = wb.create_sheet(title="라인별_총괄비교표")
        ws2.page_setup.orientation = ws2.ORIENTATION_LANDSCAPE
        ws2.page_setup.paperSize = ws2.PAPERSIZE_A4
        ws2.page_setup.fitToWidth = 1
        ws2.page_setup.fitToHeight = 0
        ws2.sheet_properties.pageSetUpPr.fitToPage = True
        
        ws2['A2'] = f"{sec} 라인번호별 비파괴검사 실적 비교표"
        ws2['A2'].font = font_title
        
        hdrs_s2 = [
            ("A4:A5", "순번"), ("B4:B5", "구간/섹션"), ("C4:C5", "라인명"),
            ("D4:D5", "관경"), ("E4:E5", "검사방법"),
            ("F4:G4", "주간(일반)"), ("H4:I4", "야간"),
            ("J4:J5", "총 수량"), ("K4:K5", "공급가액(원)"), ("L4:L5", "부가세(원)"),
            ("M4:M5", "합계금액(원)\n(VAT포함)"), ("N4:N5", "비고")
        ]
        for rng, txt in hdrs_s2:
            if ':' in rng:
                ws2.merge_cells(rng); start = rng.split(':')[0]
            else:
                start = rng
            ws2[start] = txt
            ws2[start].font = font_hdr; ws2[start].fill = fill_hdr; ws2[start].alignment = Alignment(horizontal='center', vertical='center')
            
        ws2['F5'] = "수량"; ws2['F5'].alignment = Alignment(horizontal='center', vertical='center')
        ws2['G5'] = "공급가액(원)"; ws2['G5'].alignment = Alignment(horizontal='center', vertical='center')
        ws2['H5'] = "수량"; ws2['H5'].alignment = Alignment(horizontal='center', vertical='center')
        ws2['I5'] = "공급가액(원)"; ws2['I5'].alignment = Alignment(horizontal='center', vertical='center')
        for col in ['F', 'G', 'H', 'I']:
            ws2[f'{col}5'].font = font_hdr; ws2[f'{col}5'].fill = fill_hdr
            
        for r in range(4, 6):
            for c in range(1, 15): ws2.cell(r, c).border = b_hdr
            
        # Group line records by line
        line_summary_dict = {}
        for r in data["line_records"]:
            line = r["line"]
            if line not in line_summary_dict:
                line_summary_dict[line] = {
                    "sec": r["sec"], "size": r["size"], "method": r["method"],
                    "day_qty": 0.0, "day_amt": 0, "night_qty": 0.0, "night_amt": 0,
                    "unit": r["unit"], "joints": 0, "excluded_re": 0
                }
            line_summary_dict[line]["joints"] += 1
            if r["is_re_excluded"]:
                line_summary_dict[line]["excluded_re"] += int(r["qty_raw"])
            else:
                price = 0
                for item in data["bill_items"]:
                    if item["method"] == r["method"] and item["work_time"] == r["shift"]:
                        price = item["price"]
                        break
                amt = int(round(r["qty"] * price))
                if r["shift"] == "야간":
                    line_summary_dict[line]["night_qty"] += r["qty"]
                    line_summary_dict[line]["night_amt"] += amt
                else:
                    line_summary_dict[line]["day_qty"] += r["qty"]
                    line_summary_dict[line]["day_amt"] += amt
                    
        row_cur = 6
        for idx, (line, l_info) in enumerate(sorted(line_summary_dict.items()), 1):
            tot_q = l_info["day_qty"] + l_info["night_qty"]
            sub = l_info["day_amt"] + l_info["night_amt"]
            vat = int(round(sub * 0.1))
            tot = sub + vat
            
            dq_str = f"{l_info['day_qty']:.4f} M" if l_info['unit'] == 'M' and l_info['day_qty'] > 0 else (f"{l_info['day_qty']:.1f} 매" if l_info['day_qty'] > 0 else "-")
            nq_str = f"{l_info['night_qty']:.4f} M" if l_info['unit'] == 'M' and l_info['night_qty'] > 0 else (f"{l_info['night_qty']:.1f} 매" if l_info['night_qty'] > 0 else "-")
            tq_str = f"{tot_q:.4f} M" if l_info['unit'] == 'M' else f"{tot_q:.1f} 매"
            
            rem_str = f"총 {l_info['joints']}조인트"
            if l_info['excluded_re'] > 0:
                rem_str += f" (수정재검 {l_info['excluded_re']}매 제외)"
                
            ws2.cell(row_cur, 1, idx).alignment = Alignment(horizontal='center', vertical='center')
            ws2.cell(row_cur, 2, l_info['sec'] or sec).alignment = Alignment(horizontal='center', vertical='center')
            ws2.cell(row_cur, 3, line).alignment = Alignment(horizontal='left', vertical='center')
            ws2.cell(row_cur, 4, l_info['size']).alignment = Alignment(horizontal='center', vertical='center')
            ws2.cell(row_cur, 5, l_info['method']).alignment = Alignment(horizontal='center', vertical='center')
            
            ws2.cell(row_cur, 6, dq_str).alignment = Alignment(horizontal='right', vertical='center')
            c7 = ws2.cell(row_cur, 7, l_info['day_amt'] if l_info['day_amt'] > 0 else "-")
            c7.alignment = Alignment(horizontal='right', vertical='center')
            if l_info['day_amt'] > 0: c7.number_format = '#,##0'
            
            ws2.cell(row_cur, 8, nq_str).alignment = Alignment(horizontal='right', vertical='center')
            c9 = ws2.cell(row_cur, 9, l_info['night_amt'] if l_info['night_amt'] > 0 else "-")
            c9.alignment = Alignment(horizontal='right', vertical='center')
            if l_info['night_amt'] > 0: c9.number_format = '#,##0'
            
            ws2.cell(row_cur, 10, tq_str).alignment = Alignment(horizontal='right', vertical='center')
            
            for col_no, v in [(11, sub), (12, vat), (13, tot)]:
                cell = ws2.cell(row_cur, col_no, v)
                cell.alignment = Alignment(horizontal='right', vertical='center')
                cell.number_format = '#,##0'
                
            ws2.cell(row_cur, 14, rem_str).alignment = Alignment(horizontal='left', vertical='center')
            
            for c in range(1, 15):
                ws2.cell(row_cur, c).font = font_cell
                ws2.cell(row_cur, c).border = b_all
            row_cur += 1
            
        # Total row in sheet 2
        ws2.merge_cells(f'A{row_cur}:E{row_cur}')
        ws2[f'A{row_cur}'] = "합    계"
        ws2[f'A{row_cur}'].font = font_tot
        ws2[f'A{row_cur}'].alignment = Alignment(horizontal='center', vertical='center')
        
        ws2[f'F{row_cur}'] = "주간 합계"; ws2[f'F{row_cur}'].font = font_tot; ws2[f'F{row_cur}'].alignment = Alignment(horizontal='center', vertical='center')
        ws2[f'G{row_cur}'] = f"=SUM(G6:G{row_cur-1})"; ws2[f'G{row_cur}'].font = font_tot; ws2[f'G{row_cur}'].alignment = Alignment(horizontal='right', vertical='center'); ws2[f'G{row_cur}'].number_format = '#,##0'
        
        ws2[f'H{row_cur}'] = "야간 합계"; ws2[f'H{row_cur}'].font = font_tot; ws2[f'H{row_cur}'].alignment = Alignment(horizontal='center', vertical='center')
        ws2[f'I{row_cur}'] = f"=SUM(I6:I{row_cur-1})"; ws2[f'I{row_cur}'].font = font_tot; ws2[f'I{row_cur}'].alignment = Alignment(horizontal='right', vertical='center'); ws2[f'I{row_cur}'].number_format = '#,##0'
        
        ws2[f'J{row_cur}'] = "-"; ws2[f'J{row_cur}'].font = font_tot; ws2[f'J{row_cur}'].alignment = Alignment(horizontal='center', vertical='center')
        
        for col_no, col_let in [(11, 'K'), (12, 'L'), (13, 'M')]:
            cell = ws2.cell(row_cur, col_no)
            cell.value = f"=SUM({col_let}6:{col_let}{row_cur-1})"
            cell.font = font_tot
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.number_format = '#,##0'
            
        ws2[f'N{row_cur}'] = f"총 {len(line_summary_dict)}개 라인 정산 완료"
        ws2[f'N{row_cur}'].font = font_cell_b
        
        for c in range(1, 15):
            ws2.cell(row_cur, c).fill = fill_tot
            ws2.cell(row_cur, c).border = b_tot
            
        col_w2 = {'A': 6, 'B': 18, 'C': 27, 'D': 8, 'E': 10, 'F': 10, 'G': 14, 'H': 11, 'I': 15, 'J': 12, 'K': 15, 'L': 13, 'M': 16, 'N': 35}
        for col, w in col_w2.items():
            ws2.column_dimensions[col].width = w

        # ---------------------------------------------------------
        # SHEET 3: 공종별_단가산출내역
        # ---------------------------------------------------------
        ws3 = wb.create_sheet(title="공종별_단가산출내역")
        ws3.page_setup.orientation = ws3.ORIENTATION_LANDSCAPE
        ws3.page_setup.paperSize = ws3.PAPERSIZE_A4
        ws3.sheet_properties.pageSetUpPr.fitToPage = True
        
        ws3['A2'] = "2026년 중앙지사 열수송관 비파괴검사 공종별 단가산출 내역서"
        ws3['A2'].font = font_title
        ws3['A3'] = "산출기준: 재료비 + 인건비(노무비) + 제경비(인건비 110%) + 기술료((인건비+제경비) 20%)"
        ws3['A3'].font = font_cell
        
        p_hdrs = ["순번", "공종명", "검사방법", "관경/규격", "근무구분", "단위", "재료비(원)", "인건비(원)", "제경비(110%)", "기술료(20%)", "단가 (합계원)"]
        for c_idx, h in enumerate(p_hdrs, 1):
            cell = ws3.cell(5, c_idx, h)
            cell.font = font_hdr; cell.fill = fill_hdr; cell.alignment = Alignment(horizontal='center', vertical='center'); cell.border = b_hdr
            
        u_data = [
            (1, "PAUT 300A 이상[주간]", "PAUT", "300A 이상", "일반(주간)", "M", 37559, 65200, 71720, 27384, 201863),
            (2, "PAUT 300A 이상[야간]", "PAUT", "300A 이상", "야간", "M", 37559, 97800, 107580, 41076, 284015),
            (3, "PAUT 250A[주간]", "PAUT", "250A", "일반(주간)", "M", 37559, 76163, 83779, 31988, 229489),
            (4, "PAUT 250A[야간]", "PAUT", "250A", "야간", "M", 37559, 114243, 125667, 47982, 325451),
            (5, "PAUT 200A[주간]", "PAUT", "200A", "일반(주간)", "M", 37559, 87348, 96082, 36686, 257675),
            (6, "PAUT 200A[야간]", "PAUT", "200A", "야간", "M", 37559, 131021, 144123, 55028, 367731),
            (7, "PAUT 150A-125A[주간]", "PAUT", "150A-125A", "일반(주간)", "M", 38229, 98560, 108416, 41395, 286600),
            (8, "PAUT 150A-125A[야간]", "PAUT", "150A-125A", "야간", "M", 38229, 147841, 162625, 62093, 410788),
            (9, "PAUT 100A 이하[주간]", "PAUT", "100A 이하", "일반(주간)", "M", 38229, 92362, 101598, 38792, 270981),
            (10, "PAUT 100A 이하[야간]", "PAUT", "100A 이하", "야간", "M", 38229, 138544, 152398, 58188, 387359),
            (11, "RT 3 1/3 x 12\"[주간]", "RT", "3 1/3 x 12\"", "일반(주간)", "매", 9274, 51885, 57073, 21791, 140023),
            (12, "RT 3 1/3 x 12\"[야간]", "RT", "3 1/3 x 12\"", "야간", "매", 9274, 77833, 85616, 32689, 205412),
            (13, "RT 3 1/3 x 6\"[주간]", "RT", "3 1/3 x 6\"", "일반(주간)", "매", 8044, 51885, 57073, 21791, 138793),
            (14, "RT 3 1/3 x 6\"[야간]", "RT", "3 1/3 x 6\"", "야간", "매", 8044, 77833, 85616, 32689, 204182),
            (15, "MT[주간]", "MT", "MT", "일반(주간)", "M", 411, 29771, 32748, 12503, 75433),
            (16, "PT[주간]", "PT", "PT", "일반(주간)", "M", 1177, 30847, 33931, 12955, 78910),
        ]
        for idx, pdata in enumerate(u_data):
            r = 6 + idx
            for c_idx, val in enumerate(pdata, 1):
                cell = ws3.cell(r, c_idx, val)
                cell.font = font_cell; cell.border = b_all
                if c_idx in [1, 3, 5]: cell.alignment = Alignment(horizontal='center', vertical='center')
                elif c_idx in [2, 4]: cell.alignment = Alignment(horizontal='left', vertical='center')
                else: cell.alignment = Alignment(horizontal='right', vertical='center'); cell.number_format = '#,##0'
                
        col_w3 = {'A': 6, 'B': 22, 'C': 10, 'D': 14, 'E': 12, 'F': 8, 'G': 12, 'H': 12, 'I': 13, 'J': 13, 'K': 15}
        for col, w in col_w3.items(): ws3.column_dimensions[col].width = w

        # ---------------------------------------------------------
        # SHEET 4: 기성_원천상세내역 (Actual Line Records)
        # ---------------------------------------------------------
        ws4 = wb.create_sheet(title="기성_원천상세내역")
        ws4.page_setup.orientation = ws4.ORIENTATION_LANDSCAPE
        ws4.page_setup.paperSize = ws4.PAPERSIZE_A4
        ws4.sheet_properties.pageSetUpPr.fitToPage = True
        
        ws4['A2'] = f"{sec} 기성 승인 내역 및 일자별 검사 실적 상세"
        ws4['A2'].font = font_title
        
        raw_hdrs = ["순번", "검사일자", "라인번호", "위치/구간", "검사방법", "관경", "근무구분", "기성수량", "단위", "단가(원)", "공급가액(원)", "부가세(원)", "합계금액(원)", "비고"]
        for c_idx, h in enumerate(raw_hdrs, 1):
            cell = ws4.cell(4, c_idx, h)
            cell.font = font_hdr; cell.fill = fill_hdr; cell.alignment = Alignment(horizontal='center', vertical='center'); cell.border = b_hdr
            
        r_rec_idx = 5
        for idx, r in enumerate(data["line_records"], 1):
            price = 0
            for item in data["bill_items"]:
                if item["method"] == r["method"] and item["work_time"] == r["shift"]:
                    price = item["price"]
                    break
            sub = int(round(r["qty"] * price))
            vat = int(round(sub * 0.1))
            tot = sub + vat
            
            ws4.cell(r_rec_idx, 1, idx).alignment = Alignment(horizontal='center', vertical='center')
            ws4.cell(r_rec_idx, 2, r["date"]).alignment = Alignment(horizontal='center', vertical='center')
            ws4.cell(r_rec_idx, 3, r["line"]).alignment = Alignment(horizontal='left', vertical='center')
            ws4.cell(r_rec_idx, 4, r["sec"] or sec).alignment = Alignment(horizontal='center', vertical='center')
            ws4.cell(r_rec_idx, 5, r["method"]).alignment = Alignment(horizontal='center', vertical='center')
            ws4.cell(r_rec_idx, 6, r["size"]).alignment = Alignment(horizontal='center', vertical='center')
            ws4.cell(r_rec_idx, 7, r["shift"]).alignment = Alignment(horizontal='center', vertical='center')
            
            c_q = ws4.cell(r_rec_idx, 8, r["qty"])
            c_q.alignment = Alignment(horizontal='right', vertical='center')
            c_q.number_format = '0.0000' if r["unit"] == "M" else '0.0'
            
            ws4.cell(r_rec_idx, 9, r["unit"]).alignment = Alignment(horizontal='center', vertical='center')
            
            c_p = ws4.cell(r_rec_idx, 10, price if r["qty"] > 0 else 0)
            c_p.alignment = Alignment(horizontal='right', vertical='center')
            c_p.number_format = '#,##0'
            
            for col_no, v in [(11, sub), (12, vat), (13, tot)]:
                c_a = ws4.cell(r_rec_idx, col_no, v)
                c_a.alignment = Alignment(horizontal='right', vertical='center')
                c_a.number_format = '#,##0'
                
            note_str = f"Joint: {r['joint']} ({r['res']})"
            if r["is_re_excluded"]: note_str += " [수정재검 제외]"
            ws4.cell(r_rec_idx, 14, note_str).alignment = Alignment(horizontal='left', vertical='center')
            
            for c in range(1, 15):
                ws4.cell(r_rec_idx, c).font = font_cell
                ws4.cell(r_rec_idx, c).border = b_all
            r_rec_idx += 1
            
        r_tot4 = r_rec_idx
        ws4.merge_cells(f'A{r_tot4}:G{r_tot4}')
        ws4[f'A{r_tot4}'] = f"총 실적 합계 ({len(data['line_records'])}건)"
        ws4[f'A{r_tot4}'].font = font_tot; ws4[f'A{r_tot4}'].alignment = Alignment(horizontal='center', vertical='center')
        
        for col_idx, col_let in [(11, 'K'), (12, 'L'), (13, 'M')]:
            cell = ws4.cell(r_tot4, col_idx)
            cell.value = f"=SUM({col_let}5:{col_let}{r_tot4-1})"
            cell.font = font_tot
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.number_format = '#,##0'
            
        ws4[f'N{r_tot4}'] = "정산 완료"
        ws4[f'N{r_tot4}'].font = font_cell_b
        
        for c in range(1, 15):
            ws4.cell(r_tot4, c).fill = fill_tot
            ws4.cell(r_tot4, c).border = b_tot
            
        col_w4 = {'A': 6, 'B': 12, 'C': 27, 'D': 16, 'E': 9, 'F': 8, 'G': 10, 'H': 10, 'I': 6, 'J': 13, 'K': 15, 'L': 13, 'M': 16, 'N': 35}
        for col, w in col_w4.items(): ws4.column_dimensions[col].width = w

        wb.save(filepath)
        return filepath

    def export_excel(self):
        if not self.get_selected_lines():
            messagebox.showwarning("선택 확인", "기성 청구할 라인을 최소 하나 이상 선택해 주세요.", parent=self)
            return
            
        rnd = self.round_var.get()
        sec = self.get_section_display_name(for_filename=True)
        default_name = f"제{rnd}회_{sec}_기성청구내역서_{datetime.now().strftime('%Y%m%d')}.xlsx"
        
        filepath = filedialog.asksaveasfilename(
            parent=self,
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel Files", "*.xlsx")],
            title="정식 기성청구 엑셀 양식으로 저장"
        )
        if not filepath:
            return
            
        try:
            self.generate_excel_file(filepath)
            messagebox.showinfo("저장 완료", f"기성청구 엑셀 파일이 성공적으로 생성되었습니다.\n\n{filepath}", parent=self)
        except Exception as e:
            messagebox.showerror("저장 오류", f"엑셀 파일 저장 중 오류가 발생했습니다:\n{e}", parent=self)

    def export_pdf(self):
        if not self.get_selected_lines():
            messagebox.showwarning("선택 확인", "기성 청구할 라인을 최소 하나 이상 선택해 주세요.", parent=self)
            return
            
        rnd = self.round_var.get()
        sec = self.get_section_display_name(for_filename=True)
        default_name = f"제{rnd}회_{sec}_기성청구내역서_{datetime.now().strftime('%Y%m%d')}.pdf"
        
        pdf_path = filedialog.asksaveasfilename(
            parent=self,
            defaultextension=".pdf",
            initialfile=default_name,
            filetypes=[("PDF Files", "*.pdf")],
            title="기성청구 내역서 PDF로 발행"
        )
        if not pdf_path:
            return
            
        temp_xlsx = pdf_path.replace(".pdf", "_temp.xlsx")
        try:
            self.generate_excel_file(temp_xlsx)
            
            converted = False
            try:
                import win32com.client as win32
                excel = win32.Dispatch("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
                try:
                    wb = excel.Workbooks.Open(os.path.abspath(temp_xlsx))
                    wb.ExportAsFixedFormat(0, os.path.abspath(pdf_path))
                    wb.Close(False)
                    converted = True
                finally:
                    excel.Quit()
            except Exception:
                converted = False
                
            if not converted:
                import subprocess
                helper_ps1 = os.path.join(os.path.dirname(os.path.abspath(__file__)), "export_pdf_helper.ps1")
                cmd = [
                    "powershell", "-ExecutionPolicy", "Bypass", "-File", helper_ps1,
                    "-xlsxPath", os.path.abspath(temp_xlsx),
                    "-pdfPath", os.path.abspath(pdf_path)
                ]
                res = subprocess.run(cmd, capture_output=True, text=True)
                if res.returncode != 0:
                    raise RuntimeError(f"PowerShell PDF 변환 실패: {res.stderr or res.stdout}")
                
            if os.path.exists(temp_xlsx):
                try: os.remove(temp_xlsx)
                except: pass
                
            messagebox.showinfo("발행 완료", f"기성청구 PDF 파일이 성공적으로 발행되었습니다.\n\n{pdf_path}", parent=self)
        except Exception as e:
            messagebox.showerror("발행 오류", f"PDF 발행 중 오류가 발생했습니다:\n{e}", parent=self)


if __name__ == '__main__':
    root = tk.Tk()
    root.title("Test App")
    root.geometry("400x200")
    
    def open_dlg():
        dlg = SectionLineBillingDialog(root)
        dlg.grab_set()
        
    ttk.Button(root, text="구간/라인별 기성청구 열기", command=open_dlg).pack(expand=True, pady=30)
    root.mainloop()
