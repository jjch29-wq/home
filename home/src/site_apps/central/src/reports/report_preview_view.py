# -*- coding: utf-8 -*-
"""
중앙지사 비파괴검사 보고서 프리뷰 UI 및 사이드바/툴바 관리 Mixin
"""

import json
import logging
import os
import re
import sys
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

logger = logging.getLogger(__name__)

class ReportPreviewViewMixin:
    """비파괴검사 보고서 미리보기 창, 공법별 툴바, 사이드바, 갑지 메타 UI 관리 Mixin"""
    def _create_pt_preview_ui(self, parent):
        container = tk.Frame(parent, background="#f9fafb")
        container.pack(fill='both', expand=True)

        # [NEW] File Info Header
        header_info = tk.Frame(container, background="#ffffff", highlightthickness=1, highlightbackground="#e5e7eb")
        header_info.pack(fill='x', pady=(0, 5))
        tk.Label(header_info, textvariable=self.file_info_vars['PT'], background="#ffffff", 
                 foreground="#4b5563", font=("Malgun Gothic", 8, "bold"), padx=10, pady=2).pack(side='left')

        self.pt_display_cols = ["V", "No", "Date", "Report No", "ISO/Dwg", "Joint No.", "Size", "Thk.", "Material", "Test Item", "Result", "Welder No", "Remarks"]
        saved_widths = self.config.get("PT_COL_WIDTHS", {})
        default_widths = {"V": 40, "No": 50, "Date": 90, "ISO/Dwg": 300, "Joint No.": 120, "Size": 60, "Thk.": 70, "Material": 100, "Test Item": 100, "Result": 80, "Welder No": 100}

        tree_frame = tk.Frame(container, background="#f9fafb")
        self.pt_preview_tree = ttk.Treeview(tree_frame, columns=self.pt_display_cols, show='headings', height=10, selectmode='extended')
        for col in self.pt_preview_tree["columns"]:
            name_key = f"PT_NAME_{col.split('(')[0].replace(' ', '').replace('.', '').upper()}"
            if col == "Date": name_key = "PT_NAME_DATE"
            elif col == "Report No": name_key = "PT_NAME_REPORTNO"
            elif col == "ISO/Dwg": name_key = "PT_NAME_DWG"
            elif col == "Joint No.": name_key = "PT_NAME_JOINT"
            elif col == "Material": name_key = "PT_NAME_MAT"
            elif col == "Size": name_key = "PT_NAME_SIZE"
            elif col == "Thk.": name_key = "PT_NAME_THK"
            elif col == "Test Item": name_key = "PT_NAME_ITEM"
            elif col == "Result": name_key = "PT_NAME_RES"
            elif col == "Welder No": name_key = "PT_NAME_WELDER"
            else: name_key = None
            
            display_text = self.config.get(name_key, col) if name_key else col
            self.pt_preview_tree.heading(col, text=display_text, anchor='center', command=lambda _c=col: self.sort_by_column(_c, mode="PT"))
            w = saved_widths.get(col, default_widths.get(col, 80))
            self.pt_preview_tree.column(col, width=w, anchor='center', stretch=False)
        
        scroll_y = ttk.Scrollbar(tree_frame, orient="vertical", command=self.pt_preview_tree.yview)
        scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.pt_preview_tree.xview)
        self.pt_preview_tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        
        self.pt_preview_tree.grid(row=0, column=0, sticky='nsew')
        scroll_y.grid(row=0, column=1, sticky='ns')
        scroll_x.grid(row=1, column=0, sticky='ew')
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        self._setup_preview_sidebar(self.pt_preview_tree, container, mode="PT")
        tree_frame.pack(side="left", fill="both", expand=True)



    def _create_mt_preview_ui(self, parent):
        container = tk.Frame(parent, background="#f9fafb")
        container.pack(fill='both', expand=True)

        # [NEW] File Info Header
        header_info = tk.Frame(container, background="#ffffff", highlightthickness=1, highlightbackground="#e5e7eb")
        header_info.pack(fill='x', pady=(0, 5))
        tk.Label(header_info, textvariable=self.file_info_vars['MT'], background="#ffffff",
                 foreground="#4b5563", font=("Malgun Gothic", 8, "bold"), padx=10, pady=2).pack(side='left')

        self.mt_display_cols = ["V", "No", "Date", "Report No", "ISO/Dwg", "Joint No.", "Material", "Test Item", "Result", "Welder No", "Remarks"]
        saved_widths = self.config.get("MT_COL_WIDTHS", {})
        default_widths = {"V": 40, "No": 50, "Date": 90, "Report No": 100, "ISO/Dwg": 300, "Joint No.": 120, "Material": 100, "Test Item": 100, "Result": 80, "Welder No": 100}

        tree_frame = tk.Frame(container, background="#f9fafb")
        self.mt_preview_tree = ttk.Treeview(tree_frame, columns=self.mt_display_cols, show='headings', height=10, selectmode='extended')
        for col in self.mt_preview_tree["columns"]:
            name_key = f"MT_NAME_{col.split('(')[0].replace(' ', '').replace('.', '').upper()}"
            if col == "Date": name_key = "MT_NAME_DATE"
            elif col == "Report No": name_key = "MT_NAME_REPORTNO"
            elif col == "ISO/Dwg": name_key = "MT_NAME_DWG"
            elif col == "Joint No.": name_key = "MT_NAME_JOINT"
            elif col == "Material": name_key = "MT_NAME_MAT"
            elif col == "Thk": name_key = "MT_NAME_THK"
            elif col == "Test Item": name_key = "MT_NAME_ITEM"
            elif col == "Result": name_key = "MT_NAME_RES"
            elif col == "Welder No": name_key = "MT_NAME_WELDER"
            else: name_key = None
            
            display_text = self.config.get(name_key, col) if name_key else col
            self.mt_preview_tree.heading(col, text=display_text, anchor='center', command=lambda _c=col: self.sort_by_column(_c, mode="MT"))
            w = saved_widths.get(col, default_widths.get(col, 80))
            self.mt_preview_tree.column(col, width=w, anchor='center', stretch=False)
        
        scroll_y = ttk.Scrollbar(tree_frame, orient="vertical", command=self.mt_preview_tree.yview)
        scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.mt_preview_tree.xview)
        self.mt_preview_tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        
        self.mt_preview_tree.grid(row=0, column=0, sticky='nsew')
        scroll_y.grid(row=0, column=1, sticky='ns')
        scroll_x.grid(row=1, column=0, sticky='ew')
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        self._setup_preview_sidebar(self.mt_preview_tree, container, mode="MT")
        tree_frame.pack(side="left", fill="both", expand=True)



    def _create_paut_preview_ui(self, parent):
        container = tk.Frame(parent, background="#f9fafb")
        container.pack(fill='both', expand=True)

        # [NEW] File Info Header
        header_info = tk.Frame(container, background="#ffffff", highlightthickness=1, highlightbackground="#e5e7eb")
        header_info.pack(fill='x', pady=(0, 5))
        tk.Label(header_info, textvariable=self.file_info_vars['PAUT'], background="#ffffff", 
                 foreground="#4b5563", font=("Malgun Gothic", 8, "bold"), padx=10, pady=2).pack(side='left')

        tree_frame = tk.Frame(container, background="#f9fafb")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        cols = ("V", "No", "Date", "Report No", "Line No.", "Joint No.", "Size", "Th'k(mm)", "Acc", "Rej", "Start", "End", "Length(mm)", "Upper", "Lower", "Height(mm)", "TCG MAX(%)", "Type of Flaw", "a/l", "a/t", "Welder", "Tested Length", "Evaluation", "Remarks")
        self.paut_preview_tree = ttk.Treeview(tree_frame, columns=cols, show='headings', height=10, selectmode='extended')
        
        saved_widths = self.config.get("PAUT_COL_WIDTHS", {})
        default_widths = {"V": 40, "No": 50, "Report No": 100, "Line No.": 250, "Joint No.": 120, "Size": 80, "Th'k(mm)": 60, "Acc": 50, "Rej": 50, "Start": 60, "End": 60, "Length(mm)": 80, "Upper": 60, "Lower": 60, "Height(mm)": 80, "TCG MAX(%)": 95, "Type of Flaw": 100, "a/l": 60, "a/t": 60, "Welder": 80, "Tested Length": 100, "Evaluation": 80, "Remarks": 150}
        
        for col in self.paut_preview_tree["columns"]:
            # Use dynamic names from config for headings
            name_key = f"PAUT_NAME_{col.split('(')[0].replace(' ', '').replace('.', '').replace('/', '').upper()}"
            default_name = col
            if col == "No": name_key = "PAUT_COL_NO_NAME"
            elif col == "Report No": name_key = "PAUT_NAME_REPORT_NO"
            elif col == "Line No.": name_key = "PAUT_NAME_LINE"
            elif col == "Joint No.": name_key = "PAUT_NAME_JOINT"
            elif col == "Size": name_key = "PAUT_NAME_SIZE"
            elif col == "Th'k(mm)": name_key = "PAUT_NAME_THK"
            elif col == "Acc": name_key = "PAUT_NAME_ACC"
            elif col == "Rej": name_key = "PAUT_NAME_REJ"
            elif col == "Height(mm)": name_key = "PAUT_NAME_H"
            elif col == "TCG MAX(%)": name_key = "PAUT_NAME_TCG"
            elif col == "Length(mm)": name_key = "PAUT_NAME_L"
            elif col == "Upper": name_key = "PAUT_NAME_UP"
            elif col == "Lower": name_key = "PAUT_NAME_LOW"
            elif col == "Start": name_key = "PAUT_NAME_START"
            elif col == "End": name_key = "PAUT_NAME_END"
            elif col == "a/l": name_key = "PAUT_NAME_AL"
            elif col == "a/t": name_key = "PAUT_NAME_AT"
            elif col == "Type of Flaw": name_key = "PAUT_NAME_NAT"
            elif col == "Evaluation": name_key = "PAUT_NAME_EVAL"
            elif col == "Remarks": name_key = "PAUT_NAME_REM"
            else: name_key = None
            
            display_text = self.config.get(name_key, default_name) if name_key else default_name
            self.paut_preview_tree.heading(col, text=display_text, anchor='center', command=lambda _c=col: self.sort_by_column(_c, mode="PAUT"))
            
            w = saved_widths.get(col, default_widths.get(col, 80))
            self.paut_preview_tree.column(col, width=w, anchor='center', stretch=False)
        
        self.paut_preview_tree.grid(row=0, column=0, sticky='nsew')
        
        paut_vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.paut_preview_tree.yview)
        paut_hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.paut_preview_tree.xview)
        self.paut_preview_tree.configure(yscrollcommand=paut_vsb.set, xscrollcommand=paut_hsb.set)
        
        paut_vsb.grid(row=0, column=1, sticky='ns')
        paut_hsb.grid(row=1, column=0, sticky='ew')

        self._setup_preview_sidebar(self.paut_preview_tree, container, mode="PAUT")
        tree_frame.pack(side='left', fill='both', expand=True)


        # Adaptive Resizing Bindings
        self.paut_pane_ratio = self.config.get('PAUT_SASH_RATIO', 0.5)
        self.paut_paned.bind("<Configure>", self._on_paut_paned_configure)
        self.paut_paned.bind("<ButtonRelease-1>", lambda e: self.root.after(10, self._update_paut_ratio))
        self.root.after(500, lambda: self._on_paut_paned_configure(None))


    def _create_gapji_preview_ui(self, parent, mode):
        """Create the Gapji (Cover) preview tab content."""
        container = tk.Frame(parent, background="#f3f4f6")
        container.pack(fill='both', expand=True)
        
        # Toolbar
        toolbar = tk.Frame(container, background="#ffffff", pady=5, padx=10, highlightthickness=1, highlightbackground="#e5e7eb")
        toolbar.pack(fill='x')
        ttk.Label(toolbar, text=f"📄 {mode} 갑지 레이아웃 미리보기", font=("Malgun Gothic", 9, "bold"), background="#ffffff").pack(side='left')
        
        btn_f = tk.Frame(toolbar, background="#ffffff")
        btn_f.pack(side='right')
        ttk.Button(btn_f, text=" 🔄 새로고침 ", command=lambda: self._update_gapji_preview(mode)).pack(side='right', padx=5)
        
        # Scrollable area
        outer_f = tk.Frame(container, background="#9ca3af") # Desk background
        outer_f.pack(fill='both', expand=True)
        
        # Canvas
        self.gapji_canvases = getattr(self, 'gapji_canvases', {})
        # A4 @ 72dpi is ~595x842. We'll use a slightly smaller view to fit.
        self.gapji_scale = 2.5 # 1mm = 2.5px
        canv_w, canv_h = int(210 * self.gapji_scale), int(297 * self.gapji_scale)
        
        canv = tk.Canvas(outer_f, background="#ffffff", width=canv_w, height=canv_h, 
                         highlightthickness=1, highlightbackground="#374151")
        canv.pack(pady=20)
        self.gapji_canvases[mode] = canv


    def _update_gapji_preview(self, mode):
        """Render a visual representation of the Gapji on the Canvas."""
        if not hasattr(self, 'gapji_canvases') or mode not in self.gapji_canvases: return
        canv = self.gapji_canvases[mode]
        canv.delete("all")
        
        scale = getattr(self, 'gapji_scale', 2.5)
        
        # 1. Draw Title Text (Approximation of Template)
        canv.create_text(105*scale, 40*scale, text=f"{mode} EXAMINATION REPORT", 
                         font=("Arial", 16, "bold"), fill="#1e3a8a")
        canv.create_line(30*scale, 50*scale, 180*scale, 50*scale, fill="#1e3a8a", width=2)
        
        # 2. Draw Metadata
        y_start = 120 # mm
        info = [
            ("PROJECT:", self.gapji_project.get(), y_start),
            ("CUSTOMER:", self.gapji_customer.get(), y_start + 12),
            ("AREA:", self.gapji_item.get(), y_start + 24),
            ("MATERIAL:", self.gapji_material.get(), y_start + 36),
            ("DATE:", self.gapji_exam_date.get(), y_start + 48),
            ("REPORT NO:", self.gapji_report_no.get(), y_start + 60),
        ]
        
        for label, val, y in info:
            canv.create_text(40*scale, y*scale, text=label, font=("Arial", 9, "bold"), anchor='w')
            canv.create_text(80*scale, y*scale, text=val if val else "(입력 안 됨)", 
                             font=("Arial", 9), anchor='w', fill="#374151" if val else "#9ca3af")
            canv.create_line(80*scale, (y+2)*scale, 180*scale, (y+2)*scale, fill="#e5e7eb")

        # 3. Draw Logos
        from PIL import Image, ImageTk
        logo_types = [
            ("SITCO", ["로고", "SITCO"]),
            ("SEOUL", ["서울검사", "SEOUL"]),
            ("FOOTER", ["바닥글", "FOOTER"]),
            ("FOOTER_PT", ["바닥글", "PT", "LOGO"])
        ]
        
        if not hasattr(self, '_preview_img_refs'): self._preview_img_refs = {}
        
        for lt, keywords in logo_types:
            prefix = f"{lt}_{mode}_COVER"
            if prefix not in self.config and f"{lt}_COVER" in self.config: prefix = f"{lt}_COVER"
            
            path = self.config.get(f"{prefix}_PATH", "")
            if not path:
                # Fallback to smart search if empty
                found = self.find_image_smart(keywords[0])
                if found: path = found
            
            if path and os.path.exists(path):
                try:
                    img = Image.open(path)
                    w_mm = float(self.config.get(f"{prefix}_W", 40))
                    h_mm = float(self.config.get(f"{prefix}_H", 15))
                    x_off = float(self.config.get(f"{prefix}_X", 0))
                    y_off = float(self.config.get(f"{prefix}_Y", 0))
                    anchor = self.config.get(f"{prefix}_ANCHOR", "A1")
                    
                    # Convert to pixels
                    w_px, h_px = int(w_mm * scale), int(h_mm * scale)
                    ax, ay = self._excel_anchor_to_px(anchor, scale)
                    
                    final_x = ax + (x_off * scale)
                    final_y = ay + (y_off * scale)
                    
                    # Scale and keep reference
                    img = img.resize((max(1, w_px), max(1, h_px)), Image.Resampling.LANCZOS)
                    photo = ImageTk.PhotoImage(img)
                    self._preview_img_refs[f"{mode}_{lt}"] = photo
                    
                    canv.create_image(final_x, final_y, image=photo, anchor='nw')
                except Exception as e:
                    print(f"Preview error for {lt}: {e}")


    def _create_gapji_meta_ui(self, parent, use_pack=True, mode="PMI"):
        block = tk.LabelFrame(parent, text=" 리포트 기본 정보 (Gapji Info) ", padx=10, pady=5, background="#ffffff", font=("Malgun Gothic", 9, "bold"))
        if use_pack:
            block.pack(fill='x', pady=5, padx=2)
        else:
            block.grid(row=0, column=0, sticky='ew', pady=5, padx=2)
        
        # Grid for info entries
        for col in [1, 3]: block.columnconfigure(col, weight=1)
        
        customer_label = "주문주:" if mode == "PMI" else "발주처:"
        date_label = "작성일자:" if mode == "PMI" else "검사일자:"
        report_label = "보고서 번호:" if mode == "PMI" else "리포트번호:"
        fields = [
            ("공사명:", self.gapji_project, 0, 0),
            (customer_label, self.gapji_customer, 0, 2),
            ("품명:", self.gapji_item, 1, 0),
            ("재질:", self.gapji_material, 1, 2),
            (report_label, self.gapji_report_no, 2, 0),
            (date_label, self.gapji_exam_date, 2, 2)
        ]
        if mode == 'PAUT':
            fields.append(('Sec.:', self.paut_sec, 3, 0))
            fields.append(('시공사:', self.paut_contractor, 3, 2))
        if mode == 'MT':
            fields.append(('표면온도(℃):', self.mt_surface_temp, 3, 0))
        
        for lbl, var, r, c in fields:
            tk.Label(block, text=lbl, background="#ffffff", font=("Malgun Gothic", 8)).grid(row=r, column=c, sticky='e', padx=2, pady=2)
            # PT cover metadata varies by report, so project, customer and
            # item-name fields must all allow direct editing.
            fixed_field = (mode not in ("PT", "PMI") and lbl == "품명:") or (
                mode not in ("PMI", "PT") and lbl == "발주처:"
            )
            if mode == 'MT' and lbl == '품명:':
                ent = ttk.Combobox(block, textvariable=var, width=15,
                                   values=('PIPE', 'HOT TAPPING'), state='readonly')
                ent.bind('<<ComboboxSelected>>', lambda _event: self._apply_mt_item_selection())
            elif mode == 'PMI' and lbl == '품명:':
                ent = ttk.Combobox(block, textvariable=var, width=15,
                                   values=('PIPE', 'ELBOW', 'TEE', 'REDUCER', 'FLANGE', 'VALVE', 'FITTING', 'WELD', 'BEAM/COLUMN', 'HOT TAPPING'))
                ent.bind('<<ComboboxSelected>>', lambda _event: self._on_pmi_item_changed())
                ent.bind('<FocusOut>', lambda _event: [self._on_pmi_item_changed(), self.save_settings()])
                self.pmi_item_combobox = ent
            elif mode == 'PMI' and lbl == '재질:':
                ent = ttk.Combobox(block, textvariable=var, width=15,
                                   values=list(self.PMI_MATERIAL_SPECS.keys()))
                ent.bind('<<ComboboxSelected>>', lambda _event: self._on_pmi_material_changed())
                ent.bind('<FocusOut>', lambda _event: [self._on_pmi_material_changed(), self.save_settings()])
                self.pmi_material_combobox = ent
            elif lbl == '재질:':
                from site_apps.central.src.reports.pipe_material_spec import KDHC_PIPE_MATERIALS
                ent = ttk.Combobox(block, textvariable=var, width=15,
                                   values=KDHC_PIPE_MATERIALS)
                ent.bind('<<ComboboxSelected>>', lambda _event, m=mode: self._on_gapji_material_selected(m))
                ent.bind('<FocusOut>', lambda _event, m=mode: self._on_gapji_material_selected(m, from_focus_out=True))
            else:
                ent = ttk.Entry(
                    block, textvariable=var, width=15,
                    state="readonly" if fixed_field else "normal"
                )
                ent.bind("<FocusOut>", lambda e: self.save_settings())
            ent.grid(row=r, column=c+1, sticky='ew', padx=2, pady=2)
            ent.bind("<Return>", lambda e: [self.root.focus_set(), self.save_settings()])
            # Add to setting_vars for auto-save
            cfg_key = f"GAPJI_{lbl.replace(':', '').upper()}"
            if "공사명" in lbl: cfg_key = "GAPJI_PROJECT"
            elif "발주처" in lbl or "주문주" in lbl: cfg_key = "GAPJI_CUSTOMER"
            elif "품명" in lbl: cfg_key = "GAPJI_ITEM"
            elif "재질" in lbl: cfg_key = "GAPJI_MATERIAL"
            elif "리포트번호" in lbl or "보고서 번호" in lbl: cfg_key = "GAPJI_REPORT_NO"
            elif "검사일자" in lbl or "작성일자" in lbl: cfg_key = "GAPJI_EXAM_DATE"
            elif lbl == 'Sec.:': cfg_key = 'PAUT_SEC'
            elif lbl == '시공사:': cfg_key = 'PAUT_CONTRACTOR'
            elif lbl == '표면온도(℃):': cfg_key = 'MT_SURFACE_TEMP'
            self.setting_vars[cfg_key] = var

        if mode == "PMI":
            chem_frame = tk.LabelFrame(
                block, text=" Chemical Composition ", padx=6, pady=4,
                background="#ffffff", font=("Malgun Gothic", 8, "bold")
            )
            chem_frame.grid(row=3, column=0, columnspan=4, sticky='ew', padx=2, pady=(5, 2))
            
            elem_hdr = tk.Frame(chem_frame, background="#ffffff")
            elem_hdr.grid(row=0, column=0, columnspan=6, sticky='ew', pady=(0, 2))
            tk.Label(elem_hdr, text="원소명:", background="#ffffff", font=("Malgun Gothic", 8, "bold")).pack(side='left', padx=2)
            ttk.Button(elem_hdr, text="🔍 원소 선택", command=self.show_pmi_element_selector_dialog).pack(side='right', padx=2)
            ttk.Button(elem_hdr, text="➕ 다음 기준에 추가", command=self.add_current_to_next_spec).pack(side='right', padx=(0, 4))

            for idx, var in enumerate(self.pmi_chem_elements, start=1):
                ttk.Entry(chem_frame, textvariable=var, width=8).grid(
                    row=1, column=idx-1, sticky='ew', padx=2, pady=1
                )
                chem_frame.columnconfigure(idx-1, weight=1)
                self.setting_vars[f'PMI_CHEM_ELEMENT_{idx}'] = var

            # 기준 선택기 및 규격 입력창 컨테이너
            specs_container = tk.Frame(chem_frame, background="#ffffff")
            specs_container.grid(row=2, column=0, columnspan=6, sticky='ew', pady=(4, 2))
            self.pmi_specs_container = specs_container
            self.pmi_spec_slots_ui = {}

            # 기본 슬롯 (1, 2, 3) 및 실제 유효한 내용이 저장된 확장 슬롯(4~6)만 복원
            init_slots = 3
            for i in (4, 5, 6):
                s_val = (self.config.get(f'PMI_CHEM_SPEC_{i}') or '').strip()
                m_val = (self.config.get(f'PMI_SPEC{i}_MAT') or '').strip()
                has_valid_s = s_val and s_val != "-" and not any(k in s_val for k in ("공란", "비움"))
                has_valid_m = m_val and m_val != "-" and not any(k in m_val for k in ("공란", "비움"))
                if has_valid_s or has_valid_m:
                    init_slots = max(init_slots, i)
            for slot_i in range(1, init_slots + 1):
                self._create_pmi_spec_slot_ui(slot_i, specs_container)
            
            # 초기 라벨 텍스트 동기화 및 품명 전용 재질 목록 필터링
            self._update_pmi_spec_labels()
            if hasattr(self, 'gapji_item') and self.gapji_item.get():
                self._on_pmi_item_changed()
            
            # [NEW] ⚖️ 규격 합/부 자동 판정, 🧹 기준 전체 비우기, ➕ 슬롯 추가 버튼
            action_row = tk.Frame(chem_frame, background="#ffffff")
            action_row.grid(row=3, column=0, columnspan=6, sticky='ew', pady=(4, 2))
            ttk.Button(action_row, text="⚖️ 규격 합/부 자동 판정", command=self.evaluate_pmi_data_compliance).pack(side='right', padx=2)
            ttk.Button(action_row, text="🧹 기준 전체 비우기", command=self.clear_all_pmi_specs).pack(side='right', padx=2)
            ttk.Button(action_row, text="➕ 슬롯 추가", command=self.add_pmi_spec_slot).pack(side='left', padx=2)
            ttk.Button(action_row, text="↩️ 기본 3개로 원위치", command=self.reset_pmi_spec_slots_to_default).pack(side='left', padx=2)
            
        # RT radioactive-source settings (I10:M11 on the cover sheet).
        isotope_frame = tk.Frame(block, background="#ffffff")
        if mode == "RT":
            isotope_frame.grid(row=3, column=0, columnspan=4, sticky='ew', padx=2, pady=2)
        tk.Label(
            isotope_frame, text="방사성동위원소:", background="#ffffff",
            font=("Malgun Gothic", 8)
        ).pack(side='left', padx=(0, 3))
        isotope_combo = ttk.Combobox(
            isotope_frame, textvariable=self.rt_isotope,
            values=("Ir-192", "Se-75"), state="readonly", width=8
        )
        isotope_combo.pack(side='left', padx=(0, 8))
        tk.Label(
            isotope_frame, text="방사능:", background="#ffffff",
            font=("Malgun Gothic", 8)
        ).pack(side='left', padx=(0, 3))
        activity_entry = ttk.Entry(isotope_frame, textvariable=self.rt_activity, width=7)
        activity_entry.pack(side='left', padx=(0, 3))
        unit_combo = ttk.Combobox(
            isotope_frame, textvariable=self.rt_activity_unit,
            values=("Ci", "Bq", "GBq", "TBq"), state="readonly", width=5
        )
        unit_combo.pack(side='left')
        for widget in (isotope_combo, activity_entry, unit_combo):
            widget.bind("<FocusOut>", lambda e: self.save_settings())
            widget.bind("<Return>", lambda e: [self.root.focus_set(), self.save_settings()])
        self.setting_vars['RT_ISOTOPE'] = self.rt_isotope
        self.setting_vars['RT_ACTIVITY'] = self.rt_activity
        self.setting_vars['RT_ACTIVITY_UNIT'] = self.rt_activity_unit

        film_frame = tk.Frame(block, background="#ffffff")
        if mode == "RT":
            film_frame.grid(row=4, column=0, columnspan=4, sticky='ew', padx=2, pady=2)
        tk.Label(
            film_frame, text="필름 상표:", background="#ffffff",
            font=("Malgun Gothic", 8)
        ).pack(side='left', padx=(0, 3))
        brand_combo = ttk.Combobox(
            film_frame, textvariable=self.rt_film_brand,
            values=("FUJI#80", "CARESTREAM"), width=12
        )
        brand_combo.pack(side='left', padx=(0, 8))
        tk.Label(
            film_frame, text="타입:", background="#ffffff",
            font=("Malgun Gothic", 8)
        ).pack(side='left', padx=(0, 3))
        type_combo = ttk.Combobox(
            film_frame, textvariable=self.rt_film_type,
            values=("I", "T200(I)", "MX125(I)", "AA400"), width=10
        )
        type_combo.pack(side='left', padx=(0, 8))
        tk.Label(
            film_frame, text="크기:", background="#ffffff",
            font=("Malgun Gothic", 8)
        ).pack(side='left', padx=(0, 3))
        size_combo = ttk.Combobox(
            film_frame, textvariable=self.rt_film_size,
            values=("3⅓ X 6", "3⅓ X 12", "3⅓ X 17"), width=12
        )
        size_combo.pack(side='left')
        for widget in (brand_combo, type_combo, size_combo):
            widget.bind("<FocusOut>", lambda e: self.save_settings())
            widget.bind("<Return>", lambda e: [self.root.focus_set(), self.save_settings()])
            widget.bind("<<ComboboxSelected>>", lambda e: self.save_settings())
        self.setting_vars['RT_FILM_BRAND'] = self.rt_film_brand
        self.setting_vars['RT_FILM_TYPE'] = self.rt_film_type
        self.setting_vars['RT_FILM_SIZE'] = self.rt_film_size

        if mode == "RT":
            exposure_frame = ttk.Frame(block)
            exposure_frame.grid(row=5, column=0, columnspan=4, sticky='ew', padx=2, pady=2)
            ttk.Label(exposure_frame, text='노출 시간:').pack(side='left', padx=(0, 5))
            exposure_entry = ttk.Entry(exposure_frame, textvariable=self.rt_exposure_time, width=22)
            exposure_entry.pack(side='left')
            ttk.Label(exposure_frame, text='예: 2 MIN 40 SEC').pack(side='left', padx=8)
            exposure_entry.bind('<FocusOut>', lambda e: self.save_settings())
            exposure_entry.bind('<Return>', lambda e: [self.root.focus_set(), self.save_settings()])
            self.setting_vars['RT_EXPOSURE_TIME'] = self.rt_exposure_time

        # Add a refresh button for preview (PMI는 데이터 단일 뷰이므로 미노출)
        if mode != "PMI":
            btn_f = tk.Frame(block, background="#ffffff")
            btn_f.grid(row=6 if mode == "RT" else 5, column=0, columnspan=4, pady=5, sticky='ew')
            ttk.Button(btn_f, text="✨ 갑지 미리보기 업데이트", command=self._update_gapji_preview_current).pack(side='top', fill='x', padx=5)


    def _update_gapji_preview_current(self, event=None):
        """통합된 RT 탭 및 서브 미리보기 구조에 맞춰 갑지 미리보기 갱신"""
        try:
            tab_text = self.mode_notebook.tab(self.mode_notebook.select(), "text")
            mode = "PMI"
            nb_inner = None
            
            if "RT" in tab_text:
                mode = self.rt_sub_mode.get()
                nb_inner = getattr(self, 'rt_preview_nb', None)
            elif "PT" in tab_text:
                mode = "PT"
                # PT는 기존 preview_notebooks 구조 사용 가능 여부 확인
                if hasattr(self, 'preview_notebooks'): nb_inner = self.preview_notebooks.get("PT")
            elif 'MT' in tab_text:
                mode = 'MT'
                if hasattr(self, 'preview_notebooks'): nb_inner = self.preview_notebooks.get('MT')
            elif "PAUT" in tab_text:
                mode = "PAUT"
                if hasattr(self, 'preview_notebooks'): nb_inner = self.preview_notebooks.get("PAUT")
            elif "PMI" in tab_text or "OES" in tab_text:
                mode = "PMI"
                if hasattr(self, 'preview_notebooks'): nb_inner = self.preview_notebooks.get("PMI")

            # 2. Logic for tab change event
            if event and event.widget:
                try:
                    nb = event.widget
                    inner_tab_text = nb.tab(nb.select(), "text")
                    if "갑지" in inner_tab_text:
                        if nb_inner: nb_inner.select(2) # [Standard, KOGAS, Gapji] -> Gapji is 2
                        self._update_gapji_preview(mode)
                    elif "미리보기" in inner_tab_text:
                        if nb_inner:
                            # 만약 서브 미리보기가 가스공사라면 1번, 아니면 0번
                            target_idx = 1 if "가스공사" in inner_tab_text else 0
                            nb_inner.select(target_idx)
                except: pass
            else:
                # Manual trigger (Button)
                if nb_inner:
                    for t_idx, tab_id in enumerate(nb_inner.tabs()):
                        if "갑지" in nb_inner.tab(tab_id, "text"):
                            try: nb_inner.select(t_idx)
                            except: pass
                            break
                self._update_gapji_preview(mode)
        except Exception as e:
            self.log(f"미리보기 업데이트 실패: {e}")


    def _create_kogas_separated_column_mapping_ui(self, parent):
        """[NEW] 가스공사(KOGAS) 모드 전용 의뢰서(읽기) 및 보고서(쓰기) 분리 컬럼 설정 UI"""
        # Outer Frame for Scrollbar
        outer_frame = tk.Frame(parent, background="#f9fafb")
        outer_frame.pack(fill='both', expand=True)
        
        canvas = tk.Canvas(outer_frame, background="#f9fafb", highlightthickness=0)
        scrollbar = ttk.Scrollbar(outer_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, background="#f9fafb")

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        # Mouse wheel support
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        # Main horizontal container to hold Read and Write side-by-side
        main_container = tk.Frame(scrollable_frame, background="#f9fafb")
        main_container.pack(fill='both', expand=True, pady=5, padx=5)

        # 1. Left LabelFrame: [의뢰서 읽기] 설정
        read_frame = ttk.LabelFrame(main_container, text=" 📄 [의뢰서 읽기] 컬럼 설정 (Input) ", padding=10)
        read_frame.pack(side='left', fill='both', expand=True, padx=5, pady=5)

        # 2. Right LabelFrame: [보고서 쓰기] 설정
        write_frame = ttk.LabelFrame(main_container, text=" 📝 [보고서 쓰기] 컬럼 설정 (Output) ", padding=10)
        write_frame.pack(side='right', fill='both', expand=True, padx=5, pady=5)

        # Helper to render mappings inside a frame
        def render_fields(frame, items, is_read):
            frame.columnconfigure(2, weight=1)
            # Header
            ttk.Label(frame, text="항목", font=("Malgun Gothic", 8, "bold")).grid(row=0, column=0, sticky='w', pady=(0,5))
            ttk.Label(frame, text="열", font=("Malgun Gothic", 8, "bold")).grid(row=0, column=1, sticky='w', pady=(0,5))
            ttk.Label(frame, text="이름", font=("Malgun Gothic", 8, "bold")).grid(row=0, column=2, sticky='w', pady=(0,5))

            for idx, (label, key_idx, def_idx, key_name, def_name, internal_id) in enumerate(items):
                row = idx + 1
                ttk.Label(frame, text=label, font=("Malgun Gothic", 8)).grid(row=row, column=0, sticky='w', padx=2, pady=2)
                
                v_idx = tk.StringVar(value=str(self.config.get(key_idx, def_idx)))
                ent_idx = ttk.Entry(frame, textvariable=v_idx, width=4)
                ent_idx.grid(row=row, column=1, sticky='w', padx=2, pady=2)
                ent_idx.bind("<Return>", lambda e: self.root.focus_set())
                self.setting_vars[key_idx] = v_idx

                v_name = tk.StringVar(value=str(self.config.get(key_name, def_name)))
                ent_name = ttk.Entry(frame, textvariable=v_name, width=12)
                ent_name.grid(row=row, column=2, sticky='ew', padx=2, pady=2)
                ent_name.bind("<Return>", lambda e: self.root.focus_set())
                self.setting_vars[key_name] = v_name

                # Real-time heading update (only update for tree preview if it's the Read setting since tree is built on Read data!)
                if is_read:
                    v_name.trace_add("write", lambda *args, tid=internal_id, var=v_name, kn=key_name: self._update_mode_heading("KOGAS", tid, var, kn))
                else:
                    v_name.trace_add("write", lambda *args, kn=key_name, var=v_name: self.config.update({kn: var.get()}))

        # Read items
        read_items = [
            ("No:", "KOGAS_R_COL_NO", "1", "KOGAS_R_NAME_NO", "No", "No"),
            ("Date:", "KOGAS_R_COL_DATE", "2", "KOGAS_R_NAME_DATE", "Date", "Date"),
            ("Dwg No.:", "KOGAS_R_COL_DWG", "3", "KOGAS_R_NAME_DWG", "Dwg No", "Dwg"),
            ("Film Ident. No.:", "KOGAS_R_COL_JOINT", "4", "KOGAS_R_NAME_JOINT", "Joint No", "Joint"),
            ("Film Location:", "KOGAS_R_COL_LOC", "5", "KOGAS_R_NAME_LOC", "Location", "Loc"),
            ("T:", "KOGAS_R_COL_THK", "6", "KOGAS_R_NAME_THK", "T", "T"),
            ("Mat:", "KOGAS_R_COL_MAT", "7", "KOGAS_R_NAME_MAT", "Mat", "Mat"),
            ("구경(Size):", "KOGAS_R_COL_SIZE", "8", "KOGAS_R_NAME_SIZE", "Size", "Size"),
            ("판정(Result):", "KOGAS_R_COL_RES", "9", "KOGAS_R_NAME_RES", "Result", "Result"),
            ("용접사:", "KOGAS_R_COL_WELDER", "10", "KOGAS_R_NAME_WELDER", "Welder", "Welder"),
            ("비고:", "KOGAS_R_COL_REM", "11", "KOGAS_R_NAME_REM", "Remarks", "Remarks"),
            ("① Crack:", "KOGAS_R_COL_D1", "12", "KOGAS_R_NAME_D1", "① Crack", "D1"),
            ("② IP:", "KOGAS_R_COL_D2", "13", "KOGAS_R_NAME_D2", "② IP", "D2"),
            ("③ LF:", "KOGAS_R_COL_D3", "14", "KOGAS_R_NAME_D3", "③ LF", "D3"),
            ("④ Slag:", "KOGAS_R_COL_D4", "15", "KOGAS_R_NAME_D4", "④ Slag", "D4"),
            ("⑤ Por:", "KOGAS_R_COL_D5", "16", "KOGAS_R_NAME_D5", "⑤ Por", "D5"),
            ("⑥ U/C:", "KOGAS_R_COL_D6", "17", "KOGAS_R_NAME_D6", "⑥ U/C", "D6"),
            ("⑦ RUC:", "KOGAS_R_COL_D7", "18", "KOGAS_R_NAME_D7", "⑦ RUC", "D7"),
            ("⑧ BT:", "KOGAS_R_COL_D8", "19", "KOGAS_R_NAME_D8", "⑧ BT", "D8"),
            ("⑨ IC:", "KOGAS_R_COL_D9", "20", "KOGAS_R_NAME_D9", "⑨ IC", "D9"),
            ("⑩ CP:", "KOGAS_R_COL_D10", "21", "KOGAS_R_NAME_D10", "⑩ CP", "D10"),
            ("⑪ RC:", "KOGAS_R_COL_D11", "22", "KOGAS_R_NAME_D11", "⑪ RC", "D11"),
            ("⑫ Mis:", "KOGAS_R_COL_D12", "23", "KOGAS_R_NAME_D12", "⑫ Mis", "D12"),
            ("⑬ EP:", "KOGAS_R_COL_D13", "24", "KOGAS_R_NAME_D13", "⑬ EP", "D13"),
            ("⑭ SD:", "KOGAS_R_COL_D14", "25", "KOGAS_R_NAME_D14", "⑭ SD", "D14"),
            ("⑮ Oth:", "KOGAS_R_COL_D15", "26", "KOGAS_R_NAME_D15", "⑮ Oth", "D15")
        ]

        # Write items
        write_items = [
            ("No:", "KOGAS_W_COL_NO", "1", "KOGAS_W_NAME_NO", "No", "No"),
            ("Date:", "KOGAS_W_COL_DATE", "2", "KOGAS_W_NAME_DATE", "Date", "Date"),
            ("Dwg No.:", "KOGAS_W_COL_DWG", "3", "KOGAS_W_NAME_DWG", "Drawing No.", "Dwg"),
            ("Film Ident. No.:", "KOGAS_W_COL_JOINT", "4", "KOGAS_W_NAME_JOINT", "Film Ident. No.", "Joint"),
            ("Film Location:", "KOGAS_W_COL_LOC", "5", "KOGAS_W_NAME_LOC", "Film Location", "Loc"),
            ("T:", "KOGAS_W_COL_THK", "6", "KOGAS_W_NAME_THK", "T", "T"),
            ("Mat:", "KOGAS_W_COL_MAT", "7", "KOGAS_W_NAME_MAT", "Mat", "Mat"),
            ("구경(Size):", "KOGAS_W_COL_SIZE", "0", "KOGAS_W_NAME_SIZE", "Size", "Size"),
            ("합격(ACC):", "KOGAS_W_COL_ACC", "14", "KOGAS_W_NAME_ACC", "Accept", "Acc"),
            ("불합격(REJ):", "KOGAS_W_COL_REJ", "15", "KOGAS_W_NAME_REJ", "Reject", "Rej"),
            ("등급(Grade):", "KOGAS_W_COL_DEG", "16", "KOGAS_W_NAME_DEG", "Grade", "Deg"),
            ("판정(Result):", "KOGAS_W_COL_RES", "28", "KOGAS_W_NAME_RES", "Result", "Result"),
            ("용접사:", "KOGAS_W_COL_WELDER", "29", "KOGAS_W_NAME_WELDER", "Welder No", "Welder"),
            ("비고:", "KOGAS_W_COL_REM", "30", "KOGAS_W_NAME_REM", "Remarks", "Remarks"),
            ("① Crack:", "KOGAS_W_COL_D1", "17", "KOGAS_W_NAME_D1", "① Crack", "D1"),
            ("② IP:", "KOGAS_W_COL_D2", "18", "KOGAS_W_NAME_D2", "② IP", "D2"),
            ("③ LF:", "KOGAS_W_COL_D3", "19", "KOGAS_W_NAME_D3", "③ LF", "D3"),
            ("④ Slag:", "KOGAS_W_COL_D4", "20", "KOGAS_W_NAME_D4", "④ Slag", "D4"),
            ("⑤ Por:", "KOGAS_W_COL_D5", "21", "KOGAS_W_NAME_D5", "⑤ Por", "D5"),
            ("⑥ U/C:", "KOGAS_W_COL_D6", "22", "KOGAS_W_NAME_D6", "⑥ U/C", "D6"),
            ("⑦ RUC:", "KOGAS_W_COL_D7", "23", "KOGAS_W_NAME_D7", "⑦ RUC", "D7"),
            ("⑧ BT:", "KOGAS_W_COL_D8", "24", "KOGAS_W_NAME_D8", "⑧ BT", "D8"),
            ("⑨ IC:", "KOGAS_W_COL_D9", "25", "KOGAS_W_NAME_D9", "⑨ IC", "D9"),
            ("⑩ CP:", "KOGAS_W_COL_D10", "26", "KOGAS_W_NAME_D10", "⑩ CP", "D10"),
            ("⑪ RC:", "KOGAS_W_COL_D11", "27", "KOGAS_W_NAME_D11", "⑪ RC", "D11"),
            ("⑫ Mis:", "KOGAS_W_COL_D12", "28", "KOGAS_W_NAME_D12", "⑫ Mis", "D12"),
            ("⑬ EP:", "KOGAS_W_COL_D13", "29", "KOGAS_W_NAME_D13", "⑬ EP", "D13"),
            ("⑭ SD:", "KOGAS_W_COL_D14", "30", "KOGAS_W_NAME_D14", "⑭ SD", "D14"),
            ("⑮ Oth:", "KOGAS_W_COL_D15", "31", "KOGAS_W_NAME_D15", "⑮ Oth", "D15")
        ]

        render_fields(read_frame, read_items, is_read=True)
        render_fields(write_frame, write_items, is_read=False)

        ttk.Label(scrollable_frame, text="* 엑셀열: 알파벳(A, B...) 또는 숫자(1, 2...) 입력 (0=제외)", foreground="gray", font=("Malgun Gothic", 8)).pack(pady=10)


    def _create_row_settings(self, parent, mode=None):
        if mode is None or mode == "PMI":
            # PMI Rows
            pmi_frame = ttk.LabelFrame(parent, text=" PMI 행 설정 ", padding=10)
            pmi_frame.pack(fill='x', pady=5)
            pmi_rows = [
                ("을지 데이터 시작", "START_ROW", "PMI 데이터 시작"), 
                ("을지 데이터 종료", "DATA_END_ROW", "PMI 데이터 종료"), 
                ("을지 인쇄 영역 끝", "PRINT_END_ROW", "PMI 인쇄 끝"),
                ("갑지 데이터 시작", "GAPJI_START_ROW", "PMI 갑지 데이터 시작/테두리"),
                ("갑지 데이터 종료", "GAPJI_DATA_END_ROW", "PMI 갑지 데이터 종료/테두리"),
                ("갑지 인쇄 영역 끝", "GAPJI_PRINT_END_ROW", "PMI 갑지 인쇄 끝")
            ]
            for i, (label, key, tip) in enumerate(pmi_rows):
                ttk.Label(pmi_frame, text=label + ":").grid(row=i, column=0, sticky='e', padx=5, pady=2)
                # Default values for Gapji settings
                def_val = "0"
                if key == "GAPJI_PRINT_END_ROW": def_val = "51"
                elif key == "GAPJI_START_ROW": def_val = "23"
                elif key == "GAPJI_DATA_END_ROW": def_val = "38"
                else: def_val = str(self.config.get(key, 0))
                
                v = tk.StringVar(value=str(self.config.get(key, def_val)))
                ent = ttk.Entry(pmi_frame, textvariable=v, width=10)
                ent.grid(row=i, column=1, sticky='w', padx=5)
                ent.bind("<Return>", lambda e: self.root.focus_set())
                self.setting_vars[key] = v
                ttk.Label(pmi_frame, text=f"({tip})", foreground="gray", font=("Malgun Gothic", 9)).grid(row=i, column=2, sticky='w')

        if mode is None or mode == "PAUT":
            # PAUT Rows
            ut_row_frame = ttk.LabelFrame(parent, text=" PAUT 행 설정 ", padding=10)
            ut_row_frame.pack(fill='x', pady=5)
            ut_rows = [
                ("을지 데이터 시작", "PAUT_START_ROW", "PAUT 데이터 시작"), 
                ("을지 데이터 종료", "PAUT_DATA_END_ROW", "PAUT 데이터 종료"), 
                ("을지 인쇄 영역 끝", "PAUT_PRINT_END_ROW", "PAUT 인쇄 끝"),
                ("갑지 데이터 시작", "PAUT_GAPJI_START_ROW", "PAUT 갑지 데이터 시작"),
                ("갑지 데이터 종료", "PAUT_GAPJI_DATA_END_ROW", "PAUT 갑지 데이터 종료"),
                ("갑지 인쇄 영역 끝", "PAUT_GAPJI_PRINT_END_ROW", "PAUT 갑지 끝")
            ]
            for i, (label, key, tip) in enumerate(ut_rows):
                ttk.Label(ut_row_frame, text=label + ":").grid(row=i, column=0, sticky='e', padx=5, pady=2)
                def_val = "51" if "GAPJI_PRINT" in key else "0"
                v = tk.StringVar(value=str(self.config.get(key, def_val)))
                ent = ttk.Entry(ut_row_frame, textvariable=v, width=10)
                ent.grid(row=i, column=1, sticky='w', padx=5)
                ent.bind("<Return>", lambda e: self.root.focus_set())
                self.setting_vars[key] = v
                ttk.Label(ut_row_frame, text=f"({tip})", foreground="gray", font=("Malgun Gothic", 9)).grid(row=i, column=2, sticky='w')

        if mode is None or mode in ["RT", "KOGAS"]:
            # RT/KOGAS Rows
            m_label = "KOGAS" if mode == "KOGAS" else "RT"
            rt_row_frame = ttk.LabelFrame(parent, text=f" {m_label} 행 설정 ", padding=10)
            rt_row_frame.pack(fill='x', pady=5)
            
            # Use mode-specific keys
            m_key = "KOGAS" if mode == "KOGAS" else "RT"
            rt_rows = [
                ("을지 데이터 시작", f"{m_key}_START_ROW", f"{m_label} 데이터 시작"), 
                ("을지 데이터 종료", f"{m_key}_DATA_END_ROW", f"{m_label} 데이터 종료"), 
                ("을지 인쇄 영역 끝", f"{m_key}_PRINT_END_ROW", f"{m_label} 인쇄 끝"),
                ("갑지 데이터 시작", f"{m_key}_GAPJI_START_ROW", f"{m_label} 갑지 데이터 시작"),
                ("갑지 데이터 종료", f"{m_key}_GAPJI_DATA_END_ROW", f"{m_label} 갑지 데이터 종료"),
                ("갑지 인쇄 영역 끝", f"{m_key}_GAPJI_PRINT_END_ROW", f"{m_label} 갑지 끝")
            ]
            for i, (label, key, tip) in enumerate(rt_rows):
                ttk.Label(rt_row_frame, text=label + ":").grid(row=i, column=0, sticky='e', padx=5, pady=2)
                def_val = "51" if "GAPJI_PRINT" in key else "0"
                v = tk.StringVar(value=str(self.config.get(key, def_val)))
                ent = ttk.Entry(rt_row_frame, textvariable=v, width=10)
                ent.grid(row=i, column=1, sticky='w', padx=5)
                ent.bind("<Return>", lambda e: self.root.focus_set())
                self.setting_vars[key] = v
                ttk.Label(rt_row_frame, text=f"({tip})", foreground="gray", font=("Malgun Gothic", 9)).grid(row=i, column=2, sticky='w')

        if mode is None or mode == "PT":
            # PT Rows
            pt_row_frame = ttk.LabelFrame(parent, text=" PT 행 설정 ", padding=10)
            pt_row_frame.pack(fill='x', pady=5)
            pt_rows = [
                ("을지 데이터 시작", "PT_START_ROW", "PT 데이터 시작"), 
                ("을지 데이터 종료", "PT_DATA_END_ROW", "PT 데이터 종료"), 
                ("을지 인쇄 영역 끝", "PT_PRINT_END_ROW", "PT 인쇄 끝"),
                ("갑지 데이터 시작", "PT_GAPJI_START_ROW", "PT 갑지 데이터 시작"),
                ("갑지 데이터 종료", "PT_GAPJI_DATA_END_ROW", "PT 갑지 데이터 종료"),
                ("갑지 인쇄 영역 끝", "PT_GAPJI_PRINT_END_ROW", "PT 갑지 끝")
            ]
            for i, (label, key, tip) in enumerate(pt_rows):
                ttk.Label(pt_row_frame, text=label + ":").grid(row=i, column=0, sticky='e', padx=5, pady=2)
                # Default values for PT
                if key == "PT_START_ROW": def_val = "18"
                elif key == "PT_DATA_END_ROW": def_val = "32"
                elif key == "PT_PRINT_END_ROW": def_val = "35"
                elif key == "PT_GAPJI_PRINT_END_ROW": def_val = "35"
                elif "GAPJI_START" in key: def_val = "0"
                elif "GAPJI_DATA_END" in key: def_val = "35"
                else: def_val = "0"
                
                v = tk.StringVar(value=str(self.config.get(key, def_val)))
                ent = ttk.Entry(pt_row_frame, textvariable=v, width=10)
                ent.grid(row=i, column=1, sticky='w', padx=5)
                ent.bind("<Return>", lambda e: self.root.focus_set())
                self.setting_vars[key] = v
                ttk.Label(pt_row_frame, text=f"({tip})", foreground="gray", font=("Malgun Gothic", 9)).grid(row=i, column=2, sticky='w')


    def _create_preview_ui(self, parent):
        container = tk.Frame(parent, background="#f9fafb")
        container.pack(fill='both', expand=True)

        # [NEW] File Info Header (Contextual Awareness)
        header_info = tk.Frame(container, background="#ffffff", highlightthickness=1, highlightbackground="#e5e7eb")
        header_info.pack(fill='x', pady=(0, 5))
        tk.Label(header_info, textvariable=self.file_info_vars['PMI'], background="#ffffff", 
                 foreground="#4b5563", font=("Malgun Gothic", 8, "bold"), padx=10, pady=2).pack(side='left')

        # [REFINED] Compact Header inside Frame
        header_frame = tk.Frame(container, background="#f9fafb")
        header_frame.pack(fill='x', pady=(0, 5))
        
        tk.Checkbutton(header_frame, text="⚠️ 함량 미달만 보기(PMI)", variable=self.pmi_show_deficiency_only, 
                       background="#f9fafb", font=("Malgun Gothic", 9),
                       command=lambda: self.populate_preview(self.extracted_data, switch_tab=False)).pack(side='left')

        tree_frame = tk.Frame(container, background="#f9fafb")
        self.preview_tree = ttk.Treeview(tree_frame, columns=("ST", "V", "No", "Date", "Report No", "Drawing No.", "Joint No.", "Location", "Ni", "Cr", "Mo", "Result"), show='headings', height=10, selectmode='extended')
        # [NEW] Highlight tags
        self.preview_tree.tag_configure("deficient", background="#fee2e2", foreground="#991b1b") # Light red
        self.preview_tree.tag_configure("group_even", background="#ffffff")
        self.preview_tree.tag_configure("group_odd", background="#f3f4f6")
        
        saved_widths = self.config.get("PMI_COL_WIDTHS", {})
        default_widths = {"ST": 40, "V": 40, "No": 50, "Date": 90, "Report No": 100, "Drawing No.": 400, "Joint No.": 200, "Location": 300, "Ni": 60, "Cr": 60, "Mo": 60, "Result": 150}
        
        for col in self.preview_tree["columns"]:
            # Use dynamic names from config for headings
            name_key = f"PMI_NAME_{col.split('(')[0].replace(' ', '').replace('.', '').replace('-', '').upper()}"
            if col == "No": name_key = "PMI_NAME_NO"
            elif col == "Date": name_key = "PMI_NAME_DATE"
            elif col == "Drawing No.": name_key = "PMI_NAME_DWG"
            elif col == "Joint No.": name_key = "PMI_NAME_JOINT"
            elif col == "Location": name_key = "PMI_NAME_LOC"
            elif col == "Ni": name_key = "PMI_NAME_NI"
            elif col == "Cr": name_key = "PMI_NAME_CR"
            elif col == "Mo": name_key = "PMI_NAME_MO"
            elif col == "Result": name_key = "PMI_NAME_RES"
            else: name_key = None
            
            display_text = self.config.get(name_key, col) if name_key else col
            self.preview_tree.heading(col, text=display_text, anchor='center', command=lambda _c=col: self.sort_by_column(_c, mode="PMI"))
            w = saved_widths.get(col, default_widths.get(col, 80))
            self.preview_tree.column(col, width=w, anchor='center', stretch=False)
        
        # [NEW] Add Scrollbars for full coverage
        scroll_y = ttk.Scrollbar(tree_frame, orient="vertical", command=self.preview_tree.yview)
        scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.preview_tree.xview)
        self.preview_tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        
        self.preview_tree.grid(row=0, column=0, sticky='nsew')
        scroll_y.grid(row=0, column=1, sticky='ns')
        scroll_x.grid(row=1, column=0, sticky='ew')
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        self._setup_preview_sidebar(self.preview_tree, container, mode="PMI")
        tree_frame.pack(side="left", fill="both", expand=True)


    def _create_rt_preview_ui(self, parent, mode="RT"):
        container = tk.Frame(parent, background="#f9fafb")
        container.pack(fill='both', expand=True)

        # [NEW] File Info Header
        header_info = tk.Frame(container, background="#ffffff", highlightthickness=1, highlightbackground="#e5e7eb")
        header_info.pack(fill='x', pady=(0, 5))
        tk.Label(header_info, textvariable=self.file_info_vars.get(mode, tk.StringVar()), background="#ffffff", 
                 foreground="#4b5563", font=("Malgun Gothic", 8, "bold"), padx=10, pady=2).pack(side='left')

        # Inner frame for horizontal/vertical scroll
        tree_frame = tk.Frame(container, background="#f9fafb")
        tree = ttk.Treeview(tree_frame, columns=[], show='headings', height=10, selectmode='extended')
        if mode == "RT": self.rt_preview_tree = tree
        else: self.kogas_preview_tree = tree

        # [DYNAMIC] 가스공사 모드 여부에 따른 초기 컬럼 설정
        self._update_rt_preview_columns(mode=mode)
        
        scroll_y = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        
        tree.grid(row=0, column=0, sticky='nsew')
        scroll_y.grid(row=0, column=1, sticky='ns')
        scroll_x.grid(row=1, column=0, sticky='ew')
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        self._setup_preview_sidebar(tree, container, mode=mode)
        tree_frame.pack(side="left", fill="both", expand=True)


    def _setup_preview_sidebar(self, tree, container, mode):
        # 1. Event Bindings
        self.drag_start_item = None
        
        def _on_tree_press(event, t=tree, m=mode):
            t.focus_set()
            
            # Column Drag-and-Drop Init
            if t.identify_region(event.x, event.y) == "heading":
                # Ensure we are not grabbing the separator
                if t.identify_region(event.x - 3, event.y) == "separator" or t.identify_region(event.x + 3, event.y) == "separator":
                    t.drag_col_id = None
                else:
                    t.drag_col_id = t.identify_column(event.x)
                return
            else:
                t.drag_col_id = None

            item_id = t.identify_row(event.y)
            col_id = t.identify_column(event.x)
            
            # [DYNAMIC] Get consistent module info
            mode_info = self._get_mode_info(m)
            if not mode_info: return
            _, idx_map, data_list, col_keys = mode_info
            
            self.start_col = col_id
            self.end_col = col_id
            self.last_clicked_col = col_id
            
            if not item_id or not col_id: return
            
            try:
                col_idx = int(col_id.replace("#", "")) - 1
                if 0 <= col_idx < len(col_keys):
                    key = col_keys[col_idx]
                    # Checkbox Toggle (V or selected)
                    if key in ["selected", "V"] or col_id == "#1": # Fallback to first col for V
                        if key == "selected" or (m == "PMI" and col_id == "#2"):
                             view_idx = t.index(item_id)
                             if 0 <= view_idx < len(idx_map):
                                 actual_idx = idx_map[view_idx]
                                 data_list[actual_idx]['selected'] = not data_list[actual_idx].get('selected', True)
                                 self.populate_preview(data_list, switch_tab=False, mode=m)
                                 return "break"
                    
                    # RT/KOGAS/PAUT Specific specialized toggle
                    if m in ["RT", "KOGAS", "PAUT"] and ((key.startswith("D") and key[1:].isdigit()) or (key in ["Acc", "Rej"])):
                        view_idx = t.index(item_id)
                        if 0 <= view_idx < len(idx_map):
                            actual_idx = idx_map[view_idx]
                            old_v = data_list[actual_idx].get(key, "")
                            
                            new_v = "√" if old_v == "" else ""
                                
                            data_list[actual_idx][key] = new_v
                            if key in ["Acc", "Rej"] and new_v != "":
                                other = "Rej" if key == "Acc" else "Acc"
                                data_list[actual_idx][other] = ""
                            self.populate_preview(data_list, switch_tab=False, mode=m)
                            return "break"
            except: pass

            self.drag_start_item = item_id
            if not (event.state & 0x0001 or event.state & 0x0004):
                t.selection_set(item_id)

        def _on_tree_drag(event, t=tree):
            if getattr(t, 'drag_col_id', None):
                return
                
            if not self.drag_start_item: return
            curr_item = t.identify_row(event.y)
            drag_col = t.identify_column(event.x)
            if drag_col: self.end_col = drag_col
            
            if not curr_item: return
            all_items = t.get_children('')
            try:
                low = min(all_items.index(self.drag_start_item), all_items.index(curr_item))
                high = max(all_items.index(self.drag_start_item), all_items.index(curr_item))
                t.selection_set(all_items[low:high+1])
            except: pass

        def _on_tree_release(event, t=tree):
            self.drag_start_item = None
            
            if getattr(t, 'drag_col_id', None):
                target_col = t.identify_column(event.x)
                if target_col and target_col != t.drag_col_id:
                    try:
                        curr_display = list(t["displaycolumns"])
                        if curr_display == ["#all"]:
                            curr_display = list(t["columns"])
                            
                        drag_id = t.column(t.drag_col_id, "id")
                        target_id = t.column(target_col, "id")
                        
                        if drag_id in curr_display and target_id in curr_display:
                            curr_display.remove(drag_id)
                            target_idx = curr_display.index(target_id)
                            curr_display.insert(target_idx, drag_id)
                            t["displaycolumns"] = tuple(curr_display)
                    except Exception:
                        pass
                t.drag_col_id = None

        tree.bind("<Button-1>", _on_tree_press)
        tree.bind("<B1-Motion>", _on_tree_drag)
        tree.bind("<ButtonRelease-1>", _on_tree_release)
        tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        tree.bind("<Double-1>", lambda e, m=mode: self.on_tree_double_click(e, m))
        tree.bind("<Button-3>", lambda e, m=mode: self.show_context_menu(e, m))
        tree.bind("<Control-c>", lambda e, m=mode: self.copy_cell(m))
        tree.bind("<Control-C>", lambda e, m=mode: self.copy_cell(m))
        tree.bind("<Control-v>", lambda e, m=mode: self.paste_cell(m))
        tree.bind("<Control-V>", lambda e, m=mode: self.paste_cell(m))
        tree.bind("<Control-a>", lambda e, t=tree: [t.selection_set(t.get_children(''))])
        tree.bind("<Control-A>", lambda e, t=tree: [t.selection_set(t.get_children(''))])
        tree.bind("<Delete>", lambda e, m=mode: self.on_delete_key(e, m))
        tree.bind("<F2>", lambda e, m=mode: self.on_tree_double_click(e, m))

        # 2. Sidebar with SCROLL support (too many elements to fit without scrolling)
        sidebar_outer = tk.Frame(container, background="#f1f5f9", highlightthickness=1, highlightbackground="#e2e8f0", width=160)
        sidebar_outer.pack(side='left', fill='y', padx=(0, 5))
        sidebar_outer.pack_propagate(False)
        
        # --- STICKY TOP FILTERS ---
        sticky_top_frame = tk.Frame(sidebar_outer, background="#f1f5f9", padx=5, pady=3)
        sticky_top_frame.pack(side='top', fill='x')

        ttk.Label(sticky_top_frame, text="날짜 필터", font=("Malgun Gothic", 8, "bold")).pack(pady=(2, 1))
        date_frame = tk.Frame(sticky_top_frame, background="#f1f5f9")
        date_frame.pack(fill='x', pady=2)
        
        listbox = tk.Listbox(date_frame, selectmode='single', height=5, exportselection=False, font=("Malgun Gothic", 8))
        listbox.pack(side='left', fill='both', expand=True)
        sb = ttk.Scrollbar(date_frame, orient='vertical', command=listbox.yview)
        sb.pack(side='right', fill='y')
        listbox.config(yscrollcommand=sb.set)
        
        if mode == "RT": self.rt_date_listbox = listbox
        elif mode == "PT": self.pt_date_listbox = listbox
        elif mode == "MT": self.mt_date_listbox = listbox
        elif mode == "PAUT": self.paut_date_listbox = listbox
        elif mode == "KOGAS": self.kogas_date_listbox = listbox
        else: self.date_listbox = listbox
        
        def _toggle_date(event, lb=listbox):
            idx = lb.nearest(event.y)
            if idx >= 0:
                val = lb.get(idx)
                if val.startswith("[v]"):
                    lb.delete(idx); lb.insert(idx, val.replace("[v] ", "[ ] "))
                elif val.startswith("[ ]"):
                    lb.delete(idx); lb.insert(idx, val.replace("[ ] ", "[v] "))
                lb.selection_clear(0, tk.END)
        
        listbox.bind("<ButtonRelease-1>", _toggle_date)
        
        # [NEW] Batch Select/Deselect Buttons
        date_btn_frame = tk.Frame(sticky_top_frame, background="#f1f5f9")
        date_btn_frame.pack(fill='x', pady=(0, 2))
        
        def _set_all_dates(state, lb=listbox):
            for i in range(lb.size()):
                val = lb.get(i).replace("[v] ", "").replace("[ ] ", "")
                prefix = "[v] " if state else "[ ] "
                lb.delete(i)
                lb.insert(i, prefix + val)
            _apply_date_filter(m=mode, lb=listbox)

        ttk.Button(date_btn_frame, text="전체 선택 (Select All)", command=lambda: _set_all_dates(True)).pack(side='top', fill='x', pady=1)
        ttk.Button(date_btn_frame, text="전체 해제 (Clear All)", command=lambda: _set_all_dates(False)).pack(side='top', fill='x', pady=1)
        
        def _apply_date_filter(m=mode, lb=listbox):
            sel_dates = [lb.get(i).replace("[v] ", "").replace("[ ] ", "") for i in range(lb.size()) if lb.get(i).startswith("[v]")]
            if m == "RT": data = self.rt_extracted_data
            elif m == "KOGAS": data = self.kogas_extracted_data
            elif m == "PT": data = self.pt_extracted_data
            elif m == "MT": data = self.mt_extracted_data
            elif m == "PAUT": data = self.paut_extracted_data
            else: data = self.extracted_data
            
            for item in data:
                item['date_filtered'] = (item.get('Date', '') in sel_dates)
            
            if m == "PMI":
                self.update_pmi_loc_listbox()
                
            self.populate_preview(data, switch_tab=False, mode=m)
            
        ttk.Button(sticky_top_frame, text="날짜 적용", command=_apply_date_filter).pack(fill='x', pady=(0, 5))
        if mode in ('PAUT', 'MT', 'RT', 'KOGAS'):
            for widget in sticky_top_frame.winfo_children():
                widget.pack_forget()
            if mode == 'PAUT':
                tree.heading('Date', text='Date ▾', command=self._choose_paut_dates)

        # PAUT can be narrowed independently by line number and pipe size.
        # Keep these flags separate from date_filtered so changing one filter
        # never destroys the user's selections in another filter.
        if mode == "PAUT":
            paut_filter_frame = tk.Frame(sticky_top_frame, background="#f1f5f9")

            for column, title in enumerate(("Line No.", "Size")):
                column_frame = tk.Frame(paut_filter_frame, background="#f1f5f9")
                column_frame.grid(row=0, column=column, sticky='nsew', padx=(0, 2) if column == 0 else (2, 0))
                paut_filter_frame.grid_columnconfigure(column, weight=1)
                ttk.Label(column_frame, text=title, font=("Malgun Gothic", 7)).pack()
                filter_listbox = tk.Listbox(
                    column_frame, selectmode='single', height=5,
                    exportselection=False, font=("Malgun Gothic", 7)
                )
                filter_listbox.pack(fill='both', expand=True)
                if column == 0:
                    self.paut_line_listbox = filter_listbox
                else:
                    self.paut_size_listbox = filter_listbox

            def _toggle_paut_filter(event):
                lb = event.widget
                idx = lb.nearest(event.y)
                if idx < 0 or idx >= lb.size():
                    return
                value = lb.get(idx)
                prefix = "[ ] " if value.startswith("[v] ") else "[v] "
                value = value.replace("[v] ", "").replace("[ ] ", "")
                lb.delete(idx)
                lb.insert(idx, prefix + value)
                lb.selection_clear(0, tk.END)

            self.paut_line_listbox.bind("<ButtonRelease-1>", _toggle_paut_filter)
            self.paut_size_listbox.bind("<ButtonRelease-1>", _toggle_paut_filter)

            def _set_all_paut_filters(state):
                for lb in (self.paut_line_listbox, self.paut_size_listbox):
                    for idx in range(lb.size()):
                        value = lb.get(idx).replace("[v] ", "").replace("[ ] ", "")
                        lb.delete(idx)
                        lb.insert(idx, ("[v] " if state else "[ ] ") + value)
                _apply_paut_filters()

            def _apply_paut_filters():
                selected_lines = {
                    self.paut_line_listbox.get(i).replace("[v] ", "")
                    for i in range(self.paut_line_listbox.size())
                    if self.paut_line_listbox.get(i).startswith("[v] ")
                }
                selected_sizes = {
                    self.paut_size_listbox.get(i).replace("[v] ", "")
                    for i in range(self.paut_size_listbox.size())
                    if self.paut_size_listbox.get(i).startswith("[v] ")
                }
                for item in self.paut_extracted_data:
                    item['line_filtered'] = str(item.get('Line No.', '') or '').strip() in selected_lines
                    item['size_filtered'] = str(item.get('Size', '') or '').strip() in selected_sizes
                self.populate_preview(self.paut_extracted_data, switch_tab=False, mode="PAUT")
                self._sync_paut_photo_page_numbers()

            def _choose_paut_line():
                dialog = tk.Frame(self.root, background='white', borderwidth=1, relief='solid')
                self._position_paut_filter_dropdown(dialog, 'Line No.', 620, 420)
                ttk.Label(dialog, text="원하는 라인을 여러 개 체크한 후 적용하세요.").pack(anchor='w', padx=12, pady=8)
                search = tk.StringVar()
                ttk.Entry(dialog, textvariable=search).pack(fill='x', padx=12)
                choice_frame = ttk.Frame(dialog)
                choice_frame.pack(fill='both', expand=True, padx=12, pady=4)
                choices = tk.Listbox(choice_frame, font=("Malgun Gothic", 10), exportselection=False)
                choices.pack(side='left', fill='both', expand=True)
                scrollbar = ttk.Scrollbar(choice_frame, command=choices.yview)
                scrollbar.pack(side='right', fill='y')
                choices.configure(yscrollcommand=scrollbar.set)
                states = {}
                for i in range(self.paut_line_listbox.size()):
                    text = self.paut_line_listbox.get(i)
                    states[text[4:]] = text.startswith('[v] ')
                visible = []
                def refresh(*_args):
                    visible[:] = [value for value in states if search.get().strip().casefold() in value.casefold()]
                    choices.delete(0, tk.END)
                    for value in visible:
                        choices.insert(tk.END, ('[v] ' if states[value] else '[ ] ') + value)
                def toggle(event):
                    index = choices.nearest(event.y)
                    if 0 <= index < len(visible):
                        value = visible[index]
                        states[value] = not states[value]
                        refresh()
                def set_all(state):
                    for value in states:
                        states[value] = state
                    refresh()
                def choose():
                    lb = self.paut_line_listbox
                    for i in range(lb.size()):
                        value = lb.get(i)[4:]
                        lb.delete(i)
                        lb.insert(i, ("[v] " if states[value] else "[ ] ") + value)
                    _apply_paut_filters()
                    dialog.destroy()
                choices.bind('<ButtonRelease-1>', toggle)
                def toggle_selected(_event):
                    selected = choices.curselection()
                    if selected:
                        value = visible[selected[0]]
                        states[value] = not states[value]
                        refresh()
                        choices.selection_set(selected[0])
                    return 'break'
                choices.bind('<space>', toggle_selected)
                search.trace_add('write', refresh)
                refresh()
                buttons = ttk.Frame(dialog)
                buttons.pack(fill='x', padx=12, pady=8)
                for label, command in [('전체 선택', lambda: set_all(True)), ('전체 해제', lambda: set_all(False)),
                                       ('정렬', lambda: self.sort_by_column('Line No.', mode='PAUT')),
                                       ('적용', choose), ('취소', dialog.destroy)]:
                    ttk.Button(buttons, text=label, command=command).pack(side='left', padx=3)

            self._choose_paut_line = _choose_paut_line

            paut_filter_buttons = tk.Frame(sticky_top_frame, background="#f1f5f9")
            ttk.Button(
                paut_filter_buttons, text="전체",
                command=lambda: _set_all_paut_filters(True)
            ).pack(side='left', fill='x', expand=True)
            ttk.Button(
                paut_filter_buttons, text="해제",
                command=lambda: _set_all_paut_filters(False)
            ).pack(side='left', fill='x', expand=True, padx=2)
            ttk.Button(
                paut_filter_buttons, text="적용", command=_apply_paut_filters
            ).pack(side='left', fill='x', expand=True)
        
        # [NEW] Multi-Select Test Location Filter (PMI Only)
        if mode == "PMI":
            ttk.Label(sticky_top_frame, text="지점(Loc) 필터", font=("Malgun Gothic", 8, "bold"), background="#f1f5f9").pack(pady=(2, 1))
            loc_frame = tk.Frame(sticky_top_frame, background="#f1f5f9")
            loc_frame.pack(fill='x', pady=2)
            
            self.pmi_loc_listbox = tk.Listbox(loc_frame, selectmode='single', height=4, exportselection=False, font=("Malgun Gothic", 8))
            self.pmi_loc_listbox.pack(side='left', fill='both', expand=True)
            sb_loc = ttk.Scrollbar(loc_frame, orient='vertical', command=self.pmi_loc_listbox.yview)
            sb_loc.pack(side='right', fill='y')
            self.pmi_loc_listbox.config(yscrollcommand=sb_loc.set)
            
            def _toggle_loc(event):
                idx = self.pmi_loc_listbox.nearest(event.y)
                if idx >= 0:
                    val = self.pmi_loc_listbox.get(idx)
                    if val.startswith("[v]"):
                        self.pmi_loc_listbox.delete(idx); self.pmi_loc_listbox.insert(idx, val.replace("[v] ", "[ ] "))
                    elif val.startswith("[ ]"):
                        self.pmi_loc_listbox.delete(idx); self.pmi_loc_listbox.insert(idx, val.replace("[ ] ", "[v] "))
                    self.pmi_loc_listbox.selection_clear(0, tk.END)
            
            self.pmi_loc_listbox.bind("<ButtonRelease-1>", _toggle_loc)
            
            def _apply_loc_filter():
                self.populate_preview(self.extracted_data, switch_tab=False, mode="PMI")
            
            ttk.Button(sticky_top_frame, text="지점 적용", command=_apply_loc_filter).pack(fill='x', pady=(0, 5))

        # [NEW] Sticky 'Show Selected Only' Checkbox
        tk.Checkbutton(sticky_top_frame, text="선택만", variable=self.show_selected_only, 
                       command=lambda: self.populate_preview(
                           self.rt_extracted_data if mode == "RT" else (
                               self.kogas_extracted_data if mode == "KOGAS" else (
                                   self.pt_extracted_data if mode == "PT" else (
                                       self.paut_extracted_data if mode == "PAUT" else self.extracted_data
                                   )
                               )
                           ), switch_tab=False, mode=mode),
                       background="#f1f5f9", font=("Malgun Gothic", 8)).pack(pady=(0, 3), anchor='w')

        # --- SCROLLABLE BUTTONS AREA ---
        scroll_wrapper = tk.Frame(sidebar_outer, background="#f1f5f9")
        scroll_wrapper.pack(side='top', fill='both', expand=True)

        sidebar_canvas = tk.Canvas(scroll_wrapper, background="#f1f5f9", highlightthickness=0)
        sidebar_sb = ttk.Scrollbar(scroll_wrapper, orient="vertical", command=sidebar_canvas.yview)
        sidebar = tk.Frame(sidebar_canvas, background="#f1f5f9", padx=5, pady=3)
        
        sidebar.bind("<Configure>", lambda e: sidebar_canvas.configure(scrollregion=sidebar_canvas.bbox("all")))
        # Fixed width for inner sidebar relative to outer width (160)
        # Scrollbar takes about ~15px, let's make window ~140px
        sidebar_canvas.create_window((0, 0), window=sidebar, anchor="nw", width=140)
        sidebar_canvas.configure(yscrollcommand=sidebar_sb.set)
        
        sidebar_canvas.pack(side='left', fill='both', expand=True)
        sidebar_sb.pack(side='right', fill='y')
        
        # Mousewheel scroll for sidebar
        def _sidebar_scroll(event):
            # Bulletproof bounds checking to ensure no overscroll detach
            top, bottom = sidebar_canvas.yview()
            direction = int(-1*(event.delta/120))
            
            # If everything fits, lock it
            if top <= 0.0 and bottom >= 1.0:
                sidebar_canvas.yview_moveto(0)
                return "break"
                
            if direction < 0 and top <= 0.0:
                sidebar_canvas.yview_moveto(0)
                return "break"
                
            if direction > 0 and bottom >= 1.0:
                sidebar_canvas.yview_moveto(1.0)
                return "break"
                
            # Perform scroll
            sidebar_canvas.yview_scroll(direction, "units")
            
            # Post-scroll enforcement for overshoot
            new_top, new_bottom = sidebar_canvas.yview()
            if direction < 0 and new_top <= 0.0:
                sidebar_canvas.yview_moveto(0)
            elif direction > 0 and new_bottom >= 1.0:
                sidebar_canvas.yview_moveto(1.0)
                
            return "break"
            
        sidebar_canvas.bind("<MouseWheel>", _sidebar_scroll)
        sidebar.bind("<MouseWheel>", _sidebar_scroll)
        def _bind_mousewheel_recursive(widget):
            widget.bind("<MouseWheel>", _sidebar_scroll)
            for child in widget.winfo_children():
                _bind_mousewheel_recursive(child)
        sidebar.bind("<Configure>", lambda e: _bind_mousewheel_recursive(sidebar), add="+")

        ttk.Button(sidebar, text="전체 선택", width=12, command=lambda: self.select_all(mode)).pack(pady=2, fill='x')
        ttk.Button(sidebar, text="선택 해제", width=12, command=lambda: self.deselect_all(mode)).pack(pady=2, fill='x')
        
        if mode == "PMI":
            ttk.Button(sidebar, text="미달만 선택", width=12, command=lambda: self.select_deficient_items()).pack(pady=2, fill='x')
        
        ttk.Button(sidebar, text="컬럼 관리", width=12, command=lambda: self.manage_columns(mode)).pack(pady=2, fill='x')
        
        if mode == "PAUT":
            ttk.Button(sidebar, text="행 추가", width=12, command=lambda: self.add_item(mode)).pack(pady=2, fill='x')
        
        ttk.Button(sidebar, text="병합", width=12, command=lambda: self.merge_selected_iso(mode)).pack(pady=2, fill='x')
        ttk.Button(sidebar, text="일괄 변경", width=12, command=lambda: self.show_bulk_update_dialog(mode)).pack(pady=(2, 3), fill='x')
        ttk.Button(sidebar, text="📥 요청서 매핑", width=12, command=lambda: self.import_request_mapping(mode)).pack(pady=(2, 3), fill='x')

        tk.Frame(sidebar, height=1, background="#e5e7eb").pack(fill='x', pady=5)
        ttk.Button(sidebar, text="▲ 위로", width=10, command=lambda: self.move_item(-1, mode)).pack(pady=5)
        ttk.Button(sidebar, text="▼ 아래로", width=10, command=lambda: self.move_item(1, mode)).pack(pady=5)
        tk.Frame(sidebar, height=10, background="#f9fafb").pack()
        ttk.Button(sidebar, text="선택 삭제", width=12, command=lambda: self.delete_item(mode)).pack(pady=5)
        ttk.Button(sidebar, text="전체 초기화", width=12, command=lambda: self.clear_all(mode)).pack(pady=(2, 10))

        tk.Frame(sidebar, height=1, background="#e5e7eb").pack(fill='x', pady=5)
        ttk.Button(sidebar, text="💾 현재 내용 저장", width=18, command=lambda: self.save_preview_data(mode)).pack(pady=5)
        ttk.Button(sidebar, text="📂 저장된 내용 열기", width=18, command=lambda: self.load_preview_data(mode)).pack(pady=5)
        ttk.Button(sidebar, text="📊 엑셀 파일로 추출", width=18, command=lambda: self.export_to_excel(mode)).pack(pady=(10, 2))
        if mode == 'PAUT':
            sidebar_outer.pack_forget()
            self._setup_paut_preview_toolbar(container)
        elif mode == 'MT':
            sidebar_outer.pack_forget()
            self._setup_mt_preview_toolbar(container)
            self._refresh_mt_filter_headers()
        elif mode in ('RT', 'KOGAS'):
            sidebar_outer.pack_forget()
            self._setup_rt_preview_toolbar(container, mode)
            self._refresh_rt_filter_headers(mode)

        # 3. Context Menu Config
        for t in ["evenrow", "oddrow", "group_even", "group_odd", "grouped_even", "grouped_odd", "item_even", "item_odd"]:
            tree.tag_configure(t, background="", foreground="black")
        tree.tag_configure('group_even', font=("Malgun Gothic", 10, "bold"))
        tree.tag_configure('group_odd', font=("Malgun Gothic", 10, "bold"))



    def _setup_paut_preview_toolbar(self, container):
        toolbar = ttk.Frame(container, padding=(6, 4))
        toolbar.pack(side='top', fill='x')
        ttk.Checkbutton(
            toolbar, text='체크한 항목만 보기', variable=self.show_selected_only,
            command=lambda: self.populate_preview(self.paut_extracted_data, switch_tab=False, mode='PAUT')
        ).pack(side='left', padx=(0, 12))
        ttk.Button(toolbar, text='일괄 변경',
                   command=lambda: self.show_bulk_update_dialog('PAUT')).pack(side='left', padx=3)
        edit = ttk.Menubutton(toolbar, text='행 편집 ▾')
        edit_menu = tk.Menu(edit, tearoff=False)
        edit_menu.add_command(label='행 추가', command=lambda: self.add_item('PAUT'))
        edit_menu.add_command(label='선택 행 삭제', command=lambda: self.delete_item('PAUT'))
        edit.configure(menu=edit_menu)
        edit.pack(side='left', padx=3)

        ttk.Button(toolbar, text='표 설정', command=lambda: self.manage_columns('PAUT')).pack(side='right', padx=3)
        more = ttk.Menubutton(toolbar, text='더보기 ▾')
        more_menu = tk.Menu(more, tearoff=False)
        more_menu.add_command(label='전체 초기화', command=lambda: self.clear_all('PAUT'))
        more.configure(menu=more_menu)
        more.pack(side='right', padx=3)
        export = ttk.Menubutton(toolbar, text='내보내기 ▾')
        export_menu = tk.Menu(export, tearoff=False)
        export_menu.add_command(label='엑셀 파일로 추출', command=lambda: self.export_to_excel('PAUT'))
        export.configure(menu=export_menu)
        export.pack(side='right', padx=3)
        ttk.Button(toolbar, text='불러오기', command=lambda: self.load_preview_data('PAUT')).pack(side='right', padx=3)
        ttk.Button(toolbar, text='저장', command=lambda: self.save_preview_data('PAUT')).pack(side='right', padx=3)
        self.paut_check_menu = tk.Menu(self.paut_preview_tree, tearoff=False)
        self.paut_check_menu.add_command(label='현재 필터 항목 전체 체크', command=lambda: self.select_all('PAUT'))
        self.paut_check_menu.add_command(label='현재 필터 항목 체크 해제', command=lambda: self.deselect_all('PAUT'))
        self.paut_preview_tree.heading('V', text='V ▾', command=self._show_paut_check_menu)


    def _setup_mt_preview_toolbar(self, container):
        toolbar = ttk.Frame(container, padding=(6, 4))
        toolbar.pack(side='top', fill='x')
        ttk.Checkbutton(
            toolbar, text='체크한 항목만 보기', variable=self.show_selected_only,
            command=lambda: self.populate_preview(self.mt_extracted_data, switch_tab=False, mode='MT')
        ).pack(side='left', padx=(0, 12))
        ttk.Button(toolbar, text='일괄 변경',
                   command=lambda: self.show_bulk_update_dialog('MT')).pack(side='left', padx=3)
        edit = ttk.Menubutton(toolbar, text='행 편집 ▾')
        edit_menu = tk.Menu(edit, tearoff=False)
        edit_menu.add_command(label='행 추가', command=lambda: self.add_item('MT'))
        edit_menu.add_command(label='선택 행 삭제', command=lambda: self.delete_item('MT'))
        edit.configure(menu=edit_menu)
        edit.pack(side='left', padx=3)

        ttk.Button(toolbar, text='표 설정', command=lambda: self.manage_columns('MT')).pack(side='right', padx=3)
        more = ttk.Menubutton(toolbar, text='더보기 ▾')
        more_menu = tk.Menu(more, tearoff=False)
        more_menu.add_command(label='전체 초기화', command=lambda: self.clear_all('MT'))
        more.configure(menu=more_menu)
        more.pack(side='right', padx=3)
        export = ttk.Menubutton(toolbar, text='내보내기 ▾')
        export_menu = tk.Menu(export, tearoff=False)
        export_menu.add_command(label='엑셀 파일로 추출', command=lambda: self.export_to_excel('MT'))
        export.configure(menu=export_menu)
        export.pack(side='right', padx=3)
        ttk.Button(toolbar, text='불러오기', command=lambda: self.load_preview_data('MT')).pack(side='right', padx=3)
        ttk.Button(toolbar, text='저장', command=lambda: self.save_preview_data('MT')).pack(side='right', padx=3)
        self.mt_check_menu = tk.Menu(self.mt_preview_tree, tearoff=False)
        self.mt_check_menu.add_command(label='현재 필터 항목 전체 체크', command=lambda: self.select_all('MT'))
        self.mt_check_menu.add_command(label='현재 필터 항목 체크 해제', command=lambda: self.deselect_all('MT'))
        self.mt_preview_tree.heading('V', text='V ▾', command=self._show_mt_check_menu)



    def _setup_rt_preview_toolbar(self, container, mode):
        if not hasattr(self, '_rt_check_menus'):
            self._rt_check_menus = {}
        toolbar = ttk.Frame(container, padding=(6, 4))
        toolbar.pack(side='top', fill='x')
        ttk.Checkbutton(
            toolbar, text='체크한 항목만 보기', variable=self.show_selected_only,
            command=lambda: self.populate_preview(self._get_mode_info(mode)[2], switch_tab=False, mode=mode)
        ).pack(side='left', padx=(0, 12))
        ttk.Button(toolbar, text='일괄 변경',
                   command=lambda: self.show_bulk_update_dialog(mode)).pack(side='left', padx=3)
        edit = ttk.Menubutton(toolbar, text='행 편집 ▾')
        edit_menu = tk.Menu(edit, tearoff=False)
        edit_menu.add_command(label='행 추가', command=lambda: self.add_item(mode))
        edit_menu.add_command(label='선택 행 삭제', command=lambda: self.delete_item(mode))
        edit.configure(menu=edit_menu)
        edit.pack(side='left', padx=3)

        ttk.Button(toolbar, text='표 설정', command=lambda: self.manage_columns(mode)).pack(side='right', padx=3)
        more = ttk.Menubutton(toolbar, text='더보기 ▾')
        more_menu = tk.Menu(more, tearoff=False)
        more_menu.add_command(label='전체 초기화', command=lambda: self.clear_all(mode))
        more.configure(menu=more_menu)
        more.pack(side='right', padx=3)
        export = ttk.Menubutton(toolbar, text='내보내기 ▾')
        export_menu = tk.Menu(export, tearoff=False)
        export_menu.add_command(label='엑셀 파일로 추출', command=lambda: self.export_to_excel(mode))
        export.configure(menu=export_menu)
        export.pack(side='right', padx=3)
        ttk.Button(toolbar, text='불러오기', command=lambda: self.load_preview_data(mode)).pack(side='right', padx=3)
        ttk.Button(toolbar, text='저장', command=lambda: self.save_preview_data(mode)).pack(side='right', padx=3)
        self._rt_check_menus[mode] = tk.Menu(self._get_mode_info(mode)[0], tearoff=False)
        self._rt_check_menus[mode].add_command(label='현재 필터 항목 전체 체크', command=lambda: self.select_all(mode))
        self._rt_check_menus[mode].add_command(label='현재 필터 항목 체크 해제', command=lambda: self.deselect_all(mode))
        self._get_mode_info(mode)[0].heading('selected', text='V ▾', command=lambda: self._show_rt_check_menu(mode))


