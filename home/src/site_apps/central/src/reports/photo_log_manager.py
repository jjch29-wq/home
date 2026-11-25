# -*- coding: utf-8 -*-
"""사진대장 관리 UI 및 보고서 생성 믹스인 (Photo Log Management Mixin)"""
import os
import glob
import math
import json
import io
import base64
import threading
import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from PIL import Image as PILImage, ImageChops, ImageOps
import xlsxwriter

SRC_ROOT = str(Path(__file__).resolve().parents[4])
DEFAULT_SITCO_LOGO_B64 = ""


class PhotoLogMixin:
    # --- PHOTO LOG UI & LOGIC ---
    def _setup_photo_log_ui(self, parent):
        """Standardized Photo Log UI with Dual-Pane Layout."""
        # [SYNC] Match PAUT PanedWindow properties exactly
        self.photo_paned = tk.PanedWindow(parent, orient='horizontal', background="#d1d5db", 
                            sashwidth=6, sashpad=0, sashrelief='raised', borderwidth=0)
        self.photo_paned.pack(fill='both', expand=True)

        # [LEFT] Settings Sidebar
        left_container = tk.Frame(self.photo_paned, background="#f9fafb", highlightthickness=0, borderwidth=0)
        # Add with minwidth to prevent content from being completely hidden
        self.photo_paned.add(left_container, width=425, minsize=200)

        # Scrollable area (Full size)
        left_pane = self._create_scrollable_sidebar(left_container)

        # FIXED FLOATING Header
        header_frame = tk.Frame(left_container, background="#f9fafb", highlightthickness=0, borderwidth=0)
        header_frame.place(x=0, y=0, relwidth=1, height=40)

        inner_header = tk.Frame(header_frame, background="#f9fafb", padx=20)
        inner_header.pack(fill='both', expand=True, pady=(5, 0))

        tk.Label(inner_header, text="📸 사진대장 관리", font=("Malgun Gothic", 15, "bold"), 
                 background="#f9fafb", foreground="#1e3a8a").pack(side='left')

        tk.Frame(header_frame, height=1, background="#e5e7eb").pack(side='bottom', fill='x')
        header_frame.lift()

        # Spacer in scrollable area
        tk.Frame(left_pane, height=40, background="#f9fafb").pack(fill='x')

        # 1. Report Info Group
        info_frame = ttk.LabelFrame(left_pane, text=" 리포트 정보 (Report Info) ", padding=10)
        info_frame.pack(fill='x', padx=10, pady=5)

        tk.Label(info_frame, text="검사 항목:").grid(row=0, column=0, sticky='w', pady=2)
        type_combo = ttk.Combobox(info_frame, textvariable=self.photo_inspect_type, 
                                 values=list(self.photo_header_map.keys()), state="readonly")
        type_combo.grid(row=0, column=1, sticky='ew', padx=5, pady=2)
        type_combo.bind("<<ComboboxSelected>>", self._on_photo_type_change)

        tk.Label(info_frame, text="리포트 제목:").grid(row=1, column=0, sticky='w', pady=2)
        ttk.Entry(info_frame, textvariable=self.photo_report_title, width=1).grid(row=1, column=1, sticky='ew', padx=5, pady=2)

        tk.Label(info_frame, text="발주처:").grid(row=2, column=0, sticky='w', pady=2)
        ttk.Entry(info_frame, textvariable=self.photo_orderer, width=1).grid(row=2, column=1, sticky='ew', padx=5, pady=2)

        tk.Label(info_frame, text="리포트 번호:").grid(row=3, column=0, sticky='w', pady=2)
        ttk.Entry(info_frame, textvariable=self.photo_report_no, width=1).grid(row=3, column=1, sticky='ew', padx=5, pady=2)

        tk.Label(info_frame, text="검사 일자:").grid(row=4, column=0, sticky='w', pady=2)
        ttk.Entry(info_frame, textvariable=self.photo_inspect_date, width=1).grid(row=4, column=1, sticky='ew', padx=5, pady=2)

        tk.Label(info_frame, text="로고 파일:").grid(row=5, column=0, sticky='w', pady=2)
        logo_f = tk.Frame(info_frame)
        logo_f.grid(row=5, column=1, sticky='ew', padx=5, pady=2)
        ttk.Entry(logo_f, textvariable=self.photo_logo_path).pack(side='left', fill='x', expand=True)
        def browse_logo():
            f = filedialog.askopenfilename(filetypes=[("Images", "*.png;*.jpg;*.jpeg;*.bmp")])
            if f: self.photo_logo_path.set(f)
        ttk.Button(logo_f, text="...", width=3, command=browse_logo).pack(side='right')

        # [NEW] Page Numbering Inputs
        tk.Label(info_frame, text="시작 페이지:").grid(row=6, column=0, sticky='w', pady=2)
        self.photo_start_page = tk.StringVar(value=str(self.config.get("PHOTO_START_PAGE", "")))
        ttk.Entry(info_frame, textvariable=self.photo_start_page, width=1).grid(row=6, column=1, sticky='ew', padx=5, pady=2)

        tk.Label(info_frame, text="전체 페이지:").grid(row=7, column=0, sticky='w', pady=2)
        self.photo_total_pages = tk.StringVar(value=str(self.config.get("PHOTO_TOTAL_PAGES", "")))
        ttk.Entry(info_frame, textvariable=self.photo_total_pages, width=1).grid(row=7, column=1, sticky='ew', padx=5, pady=2)

        info_frame.columnconfigure(1, weight=1)

        # 2. Layout Settings Group
        layout_frame = ttk.LabelFrame(left_pane, text=" 사진 레이아웃 설정 (Layout) ", padding=10)
        layout_frame.pack(fill='x', padx=10, pady=5)

        tk.Label(layout_frame, text="한 줄당 사진:").grid(row=0, column=0, sticky='w', pady=2)
        r0_f = tk.Frame(layout_frame)
        r0_f.grid(row=0, column=1, columnspan=4, sticky='w')
        ttk.Combobox(r0_f, textvariable=self.photo_cols_per_row, values=["1", "2", "3", "4"], state="readonly", width=5).pack(side='left', padx=2)
        tk.Label(r0_f, text="페이지당 줄 수:").pack(side='left', padx=(10, 2))
        ttk.Combobox(r0_f, textvariable=self.photo_rows_per_page, values=["2", "3", "4", "5", "6", "8"], state="readonly", width=5).pack(side='left', padx=2)
        ttk.Checkbutton(r0_f, text="비율 유지", variable=self.photo_keep_aspect).pack(side='left', padx=(10, 2))

        tk.Label(layout_frame, text="칸 너비/높이:").grid(row=1, column=0, sticky='w', pady=2)
        wh_f = tk.Frame(layout_frame)
        wh_f.grid(row=1, column=1, columnspan=2, sticky='w')
        ttk.Entry(wh_f, textvariable=self.photo_cell_width_var, width=6).pack(side='left', padx=2)
        ttk.Entry(wh_f, textvariable=self.photo_cell_height_var, width=6).pack(side='left', padx=2)

        tk.Label(layout_frame, text="여백(T/B/L/R):").grid(row=2, column=0, sticky='w', pady=2)
        m_f = tk.Frame(layout_frame)
        m_f.grid(row=2, column=1, columnspan=2, sticky='w')
        ttk.Entry(m_f, textvariable=self.photo_margin_top_var, width=4).pack(side='left', padx=1)
        ttk.Entry(m_f, textvariable=self.photo_margin_bottom_var, width=4).pack(side='left', padx=1)
        ttk.Entry(m_f, textvariable=self.photo_margin_left_var, width=4).pack(side='left', padx=1)
        ttk.Entry(m_f, textvariable=self.photo_margin_right_var, width=4).pack(side='left', padx=1)

        tk.Label(layout_frame, text="설명 높이:").grid(row=3, column=0, sticky='w', pady=2)
        ttk.Entry(layout_frame, textvariable=self.photo_desc_height_var, width=10).grid(row=3, column=1, sticky='w', padx=5, pady=2)

        tk.Label(layout_frame, text="인쇄 배율:").grid(row=4, column=0, sticky='w', pady=2)
        r4_f = tk.Frame(layout_frame)
        r4_f.grid(row=4, column=1, columnspan=2, sticky='w')
        ttk.Entry(r4_f, textvariable=self.photo_print_scale_var, width=6).pack(side='left', padx=2)
        tk.Label(r4_f, text="출력 해상도(DPI):").pack(side='left', padx=(5, 2))
        ttk.Entry(r4_f, textvariable=self.photo_dpi_var, width=4).pack(side='left', padx=2)
        tk.Label(r4_f, text="(기본: 96)", font=('', 9), foreground='gray').pack(side='left')

        tk.Label(layout_frame, text="배치 설정:").grid(row=5, column=0, sticky='w', pady=2)
        b_f = tk.Frame(layout_frame)
        b_f.grid(row=5, column=1, columnspan=2, sticky='w')
        ttk.Combobox(b_f, textvariable=self.photo_align_var, values=["좌측 정렬", "중앙 정렬"], state="readonly", width=10).pack(side='left', padx=2)
        ttk.Checkbutton(b_f, text="가로 폭 맞춤 (Fit to Width)", variable=self.photo_fit_width_var).pack(side='left', padx=5)
        ttk.Checkbutton(b_f, text="A4 인쇄영역 자동맞춤", variable=self.photo_auto_fit_page_var).pack(side='left', padx=5)
        ttk.Checkbutton(b_f, text="세로 사진 자동 회전", variable=self.photo_auto_rotate_var).pack(side='left', padx=5)

        tk.Label(layout_frame, text="너비비율(%):").grid(row=6, column=0, sticky='w')
        wf_f = tk.Frame(layout_frame)
        wf_f.grid(row=6, column=1, columnspan=2, sticky='w')
        ttk.Entry(wf_f, textvariable=self.photo_width_pct_var, width=7).pack(side='left', padx=2)
        tk.Label(wf_f, text="너비여백(px):").pack(side='left', padx=(10, 0))
        ttk.Spinbox(wf_f, textvariable=self.photo_width_pixel_adj_var, from_=-500, to=500, increment=1, width=4).pack(side='left', padx=2)
        tk.Label(wf_f, text="상하여백(px):").pack(side='left', padx=(10, 0))
        ttk.Spinbox(wf_f, textvariable=self.photo_height_pixel_adj_var, from_=-500, to=500, increment=1, width=4).pack(side='left', padx=2)
        tk.Label(wf_f, text="좌우(px):").pack(side='left', padx=(10, 0))
        ttk.Spinbox(wf_f, textvariable=self.photo_shift_x_var, from_=-500, to=500, increment=1, width=4).pack(side='left', padx=2)
        ttk.Spinbox(wf_f, textvariable=self.photo_shift_y_var, from_=-500, to=500, increment=1, width=4).pack(side='left', padx=2)
        tk.Label(wf_f, text="(100% 기준 미세조정)", font=('', 9), foreground='gray').pack(side='left', padx=5)

        tk.Label(layout_frame, text="높이비율(%):").grid(row=7, column=0, sticky='w', pady=2)
        ttk.Entry(layout_frame, textvariable=self.photo_height_pct_var, width=10).grid(row=7, column=1, sticky='w', padx=2)

        # 3. Logo Options
        logo_frame = ttk.LabelFrame(left_pane, text=" 로고 및 출력 설정 ", padding=10)
        logo_frame.pack(fill='x', padx=10, pady=5)

        tk.Label(logo_frame, text="로고 너비/높이:").grid(row=0, column=0, sticky='w', pady=2)
        wh_f = tk.Frame(logo_frame)
        wh_f.grid(row=0, column=1, sticky='w')
        ttk.Entry(wh_f, textvariable=self.photo_logo_width_var, width=5).pack(side='left', padx=2)
        ttk.Entry(wh_f, textvariable=self.photo_logo_height_var, width=5).pack(side='left', padx=2)
        tk.Label(wh_f, text="(높이 공란 시 비율유지)", font=('', 8), fg='gray').pack(side='left')

        tk.Label(logo_frame, text="로고 X/Y:").grid(row=1, column=0, sticky='w', pady=2)
        xy_f = tk.Frame(logo_frame)
        xy_f.grid(row=1, column=1, sticky='w')
        ttk.Entry(xy_f, textvariable=self.photo_logo_x_var, width=5).pack(side='left', padx=2)
        ttk.Entry(xy_f, textvariable=self.photo_logo_y_var, width=5).pack(side='left', padx=2)

        # [NEW] Photo Log Layout Save Button
        btn_f = tk.Frame(logo_frame, background="#ffffff")
        btn_f.grid(row=2, column=0, columnspan=2, sticky='ew', pady=5)
        ttk.Button(btn_f, text="💾 설정 저장", command=self.save_photo_log_config).pack(side='left', padx=5)
        tk.Label(btn_f, text="* 여백, 로고위치 등 저장", font=("Malgun Gothic", 8), fg="gray", background="#ffffff").pack(side='left')

        tk.Label(logo_frame, text="출력 파일명:").grid(row=3, column=0, sticky='w', pady=2)
        ttk.Entry(logo_frame, textvariable=self.photo_output_name, width=1).grid(row=3, column=1, sticky='ew', padx=5, pady=2)
        logo_frame.columnconfigure(1, weight=1)

        # Actions moved to right Top Action Bar

        # [RIGHT] Preview & File List
        right_container = tk.Frame(self.photo_paned, background="#ffffff")
        self.photo_paned.add(right_container)

        # [TOP ACTION BAR] 사진대장 리포트 생성 및 사진 추가 툴바
        photo_action_bar = tk.Frame(right_container, background="#ffffff", padx=8, pady=6)
        photo_action_bar.pack(fill='x')

        btn_photo_gen = tk.Button(
            photo_action_bar, text="🚀 사진대장 리포트 생성", command=self.start_photo_generation,
            font=("Malgun Gothic", 10, "bold"), bg="#2563eb", fg="#ffffff",
            activebackground="#1d4ed8", activeforeground="#ffffff",
            relief="flat", cursor="hand2", padx=12, pady=5, bd=0
        )
        btn_photo_gen.pack(side='left', fill='x', expand=True, padx=(0, 6))
        btn_photo_gen.bind("<Enter>", lambda e, b=btn_photo_gen: b.configure(bg="#1d4ed8"))
        btn_photo_gen.bind("<Leave>", lambda e, b=btn_photo_gen: b.configure(bg="#2563eb"))

        btn_photo_daily = tk.Button(
            photo_action_bar, text="📋 작업일보 사진 불러오기", command=self._import_daily_work_photos,
            font=("Malgun Gothic", 9, "bold"), bg="#f8fafc", fg="#334155",
            activebackground="#e2e8f0", activeforeground="#1e293b",
            relief="solid", bd=1, cursor="hand2", padx=10, pady=5
        )
        btn_photo_daily.pack(side='left', fill='x', expand=True, padx=(0, 6))
        btn_photo_daily.bind("<Enter>", lambda e, b=btn_photo_daily: b.configure(bg="#e2e8f0"))
        btn_photo_daily.bind("<Leave>", lambda e, b=btn_photo_daily: b.configure(bg="#f8fafc"))

        btn_photo_add = tk.Button(
            photo_action_bar, text="📁 파일 개별 추가", command=self._add_photo_files,
            font=("Malgun Gothic", 9, "bold"), bg="#f8fafc", fg="#334155",
            activebackground="#e2e8f0", activeforeground="#1e293b",
            relief="solid", bd=1, cursor="hand2", padx=10, pady=5
        )
        btn_photo_add.pack(side='left', fill='x', expand=True, padx=(0, 6))
        btn_photo_add.bind("<Enter>", lambda e, b=btn_photo_add: b.configure(bg="#e2e8f0"))
        btn_photo_add.bind("<Leave>", lambda e, b=btn_photo_add: b.configure(bg="#f8fafc"))

        btn_photo_folder = tk.Button(
            photo_action_bar, text="📂 폴더 전체 추가", command=self._add_photo_folder,
            font=("Malgun Gothic", 9, "bold"), bg="#f8fafc", fg="#334155",
            activebackground="#e2e8f0", activeforeground="#1e293b",
            relief="solid", bd=1, cursor="hand2", padx=10, pady=5
        )
        btn_photo_folder.pack(side='left', fill='x', expand=True)
        btn_photo_folder.bind("<Enter>", lambda e, b=btn_photo_folder: b.configure(bg="#e2e8f0"))
        btn_photo_folder.bind("<Leave>", lambda e, b=btn_photo_folder: b.configure(bg="#f8fafc"))

        tk.Frame(right_container, height=1, background="#e2e8f0").pack(fill='x', pady=(0, 4))

        # [NEW] Live Preview Canvas
        preview_frame = tk.Frame(right_container, background="#e2e8f0", padx=10, pady=10)
        preview_frame.pack(fill='x')
        tk.Label(preview_frame, text="👁️ 실시간 레이아웃 미리보기 (선택된 사진 기준)", font=("Malgun Gothic", 10, "bold"), background="#e2e8f0", foreground="#1e293b").pack(anchor='w')
        self.photo_preview_canvas = tk.Canvas(preview_frame, bg="#ffffff", height=300, highlightthickness=1, highlightbackground="#cbd5e1")
        self.photo_preview_canvas.pack(fill='x', pady=5)

        # File List Header
        list_header = tk.Frame(right_container, background="#f8fafc", padx=15, pady=10)
        list_header.pack(fill='x')
        tk.Label(list_header, text="📁 선택된 사진 리스트 (목록 순서대로 출력)", font=("Malgun Gothic", 10, "bold"),
                 background="#f8fafc", foreground="#475569").pack(side='left')

        # Tools under header (리스트 순서 조정 및 편집 전용)
        tool_bar = tk.Frame(right_container, background="#ffffff", padx=10, pady=5)
        tool_bar.pack(fill='x')
        ttk.Button(tool_bar, text="▲ 위로", command=lambda: self._move_photo_selected(-1)).pack(side='left', padx=2)
        ttk.Button(tool_bar, text="▼ 아래로", command=lambda: self._move_photo_selected(1)).pack(side='left', padx=2)
        ttk.Button(tool_bar, text="파일명 정렬", command=self._sort_photo_list).pack(side='left', padx=2)
        ttk.Button(tool_bar, text="전체 비우기", command=self._clear_photo_all).pack(side='right', padx=2)
        ttk.Button(tool_bar, text="선택 항목 제거", command=self._remove_photo_selected).pack(side='right', padx=2)

        # Listbox area
        list_frame = tk.Frame(right_container, background="#ffffff")
        list_frame.pack(fill='both', expand=True, padx=10, pady=5)

        self.photo_listbox = tk.Listbox(list_frame, font=("Consolas", 9), selectmode="extended", borderwidth=1, relief='flat', highlightthickness=1, highlightcolor="#3b82f6", exportselection=False)
        self.photo_listbox.pack(side='left', fill='both', expand=True)

        list_vsb = ttk.Scrollbar(list_frame, orient="vertical", command=self.photo_listbox.yview)
        # [SYNC] Sash persistence bindings (Match PAUT)
        self.photo_pane_ratio = self.config.get('PHOTO_SASH_RATIO', 0.45)
        self.photo_paned.bind("<Configure>", self._on_photo_paned_configure)
        self.photo_paned.bind("<ButtonRelease-1>", lambda e: self.root.after(10, self._update_photo_ratio))

        # [RESTORED] Sync listbox if any files were loaded from settings
        if self.photo_selected_files:
            self._refresh_photo_list()

        # [NEW] Bindings for Live Preview Update
        self.photo_listbox.bind("<<ListboxSelect>>", self._update_photo_preview)
        for var in (self.photo_width_pct_var, self.photo_height_pct_var, self.photo_width_pixel_adj_var, self.photo_height_pixel_adj_var, self.photo_shift_x_var,
                    self.photo_shift_y_var, self.photo_cell_width_var, self.photo_cell_height_var,
                    self.photo_cols_per_row, self.photo_rows_per_page):
            var.trace_add("write", lambda *args: self.root.after(300, self._update_photo_preview))
        self.photo_keep_aspect.trace_add("write", lambda *args: self.root.after(300, self._update_photo_preview))
        self.photo_fit_width_var.trace_add("write", lambda *args: self.root.after(300, self._update_photo_preview))
        self.photo_auto_rotate_var.trace_add("write", lambda *args: self.root.after(300, self._update_photo_preview))

        self._apply_sash_ratio("PHOTO")
        self.root.after(500, lambda: self._on_photo_paned_configure(None))
        self._sync_paut_photo_page_numbers()

        # [NEW] 화면을 처음 켰을 때, 선택된 게 없더라도 초기 화면/안내 문구가 뜨도록 강제 1회 실행
        self.root.after(800, self._update_photo_preview)

    def _update_photo_ratio(self):
        try:
            total_w = self.photo_paned.winfo_width()
            if total_w > 100:
                current_sash = self.photo_paned.sash_coord(0)[0]
                self.photo_pane_ratio = current_sash / total_w
                self.config['PHOTO_SASH_RATIO'] = self.photo_pane_ratio
                self.save_settings()
        except: pass

    def _on_photo_paned_configure(self, event):
        try:
            total_w = self.photo_paned.winfo_width()
            if total_w > 100:
                if event and event.widget == self.photo_paned:
                    new_pos = int(total_w * self.photo_pane_ratio)
                    self.photo_paned.sash_place(0, new_pos, 0)
        except: pass

    def _on_photo_type_change(self, event=None):
        new_type = self.photo_inspect_type.get()
        if new_type in self.photo_header_map:
            self.photo_report_title.set(self.photo_header_map[new_type])
        self._sync_photo_report_info()
        self.save_settings()

    @staticmethod
    def _photo_metadata_key(path):
        return os.path.normcase(os.path.abspath(os.path.normpath(path)))

    def _central_daily_history_path(self):
        if not SRC_ROOT:
            return ""
        return os.path.join(
            SRC_ROOT, "site_apps", "central", "src", "daily_work_history.json"
        )

    def _daily_photo_description(self, img_path):
        metadata = self.photo_item_metadata.get(
            self._photo_metadata_key(img_path), {}
        )
        line_no = str(metadata.get('line_no', '')).strip()
        size = str(metadata.get('size', '')).strip()
        # 라인번호가 이미 관경으로 끝나면(예: ...-700A) 관경을 다시 붙이지 않는다.
        if line_no and size and line_no.upper().endswith(size.upper()):
            size = ""
        details = [
            line_no,
            size,
            str(metadata.get('joint_no', '')).strip(),
            str(metadata.get('welder', '')).strip(),
        ]
        details = [value for value in details if value]
        if details:
            return "_".join(details)
        saved_description = str(metadata.get('description', '')).strip()
        if saved_description:
            return saved_description
        return os.path.splitext(os.path.basename(img_path))[0].replace(
            "_capture", ""
        )

    def _is_daily_work_photo(self, img_path):
        metadata = self.photo_item_metadata.get(
            self._photo_metadata_key(img_path), {}
        )
        return metadata.get('source') == 'central_daily_work_log'

    def _should_auto_rotate_photo(self, img_path):
        """작업일보 세로 사진은 가로형 사진대장에 맞게 자동 회전한다."""
        return self.photo_auto_rotate_var.get() or self._is_daily_work_photo(
            img_path
        )

    def _schedule_photo_output(self, image_files, photos_per_page):
        """데이터 사진 뒤의 새 페이지부터 작업일보 사진 슬롯을 배정한다."""
        data_photos = [
            path for path in image_files if not self._is_daily_work_photo(path)
        ]
        daily_photos = [
            path for path in image_files if self._is_daily_work_photo(path)
        ]

        scheduled = []
        next_slot = 0
        for path in data_photos:
            scheduled.append((path, next_slot))
            next_slot += 1

        if data_photos and daily_photos:
            next_slot = (
                math.ceil(next_slot / photos_per_page) * photos_per_page
            )

        daily_start_slot = next_slot if daily_photos else None
        for path in daily_photos:
            scheduled.append((path, next_slot))
            next_slot += 1

        return scheduled, next_slot, daily_start_slot

    @staticmethod
    def _normalize_daily_photo_welder(welder):
        """교정 전 용접사 번호를 최신 발급번호로 변환한다."""
        welder_id = str(welder or '').strip()
        legacy_ids = {
            'W-2023-A-12': 'W-2026-A-08',
            'W-2023-A-13': 'W-2026-A-05',
        }
        return legacy_ids.get(welder_id, welder_id)

    def _select_daily_work_photos(self, history, inspect_type, preferred_date):
        """작업일보 사진의 날짜/관경/Joint 정보를 보여주고 복수 선택한다."""
        allowed_all = inspect_type in {"NDT", "기타 (직접 입력)".upper()}
        candidates = []
        for date_str in sorted(history):
            day_data = history[date_str]
            if not isinstance(day_data, dict):
                continue

            result_lookup = {}
            for result in day_data.get('ndt_results', []):
                if not isinstance(result, dict):
                    continue
                key = (
                    str(result.get('검사방법', '')).strip().upper(),
                    str(result.get('라인번호', '')).strip(),
                    str(result.get('Joint No.', '')).strip(),
                )
                result_lookup[key] = result

            for photo in day_data.get('process_photos', []):
                if not isinstance(photo, dict):
                    continue
                process = str(photo.get('process', '')).strip().upper()
                if not allowed_all and process != inspect_type:
                    continue
                line_no = str(photo.get('line_no', '')).strip()
                joint_no = str(photo.get('joint_no', '')).strip()
                result = result_lookup.get((process, line_no, joint_no), {})
                item = dict(photo)
                item['_source_date'] = date_str
                item['_size'] = str(result.get('관경', '')).strip()
                if not str(item.get('welder', '')).strip():
                    item['welder'] = str(result.get('용접사', '')).strip()
                candidates.append(item)

        if not candidates:
            messagebox.showinfo(
                "저장 사진 없음", f"저장된 {inspect_type} 작업일보 사진이 없습니다."
            )
            return []

        dialog = tk.Toplevel(self.root)
        dialog.title("작업일보 사진 선택")
        screen_width = dialog.winfo_screenwidth()
        screen_height = dialog.winfo_screenheight()
        dialog_width = min(940, max(760, screen_width - 80))
        dialog_height = min(600, max(440, screen_height - 120))
        dialog_x = max(20, (screen_width - dialog_width) // 2)
        dialog_y = max(20, (screen_height - dialog_height) // 2)
        dialog.geometry(
            f"{dialog_width}x{dialog_height}+{dialog_x}+{dialog_y}"
        )
        dialog.minsize(min(760, dialog_width), min(440, dialog_height))
        dialog.transient(self.root)
        dialog.grab_set()

        tk.Label(
            dialog,
            text=f"불러올 {inspect_type} 사진을 선택하세요. (Ctrl/Shift 복수 선택)",
            font=("Malgun Gothic", 10, "bold"),
            anchor='w',
        ).pack(fill='x', padx=15, pady=(15, 8))
        tk.Label(
            dialog,
            text=f"보고서 검사일자({preferred_date})는 변경되지 않습니다.",
            foreground='#555555', anchor='w',
        ).pack(fill='x', padx=15, pady=(0, 8))

        # 목록이 커져도 상태와 실행 버튼은 항상 창 하단에 남겨 둔다.
        footer_frame = ttk.Frame(dialog)
        footer_frame.pack(side='bottom', fill='x', padx=15, pady=(2, 12))

        tree_frame = ttk.Frame(dialog)
        tree_frame.pack(fill='both', expand=True, padx=15, pady=5)
        photo_tree = ttk.Treeview(
            tree_frame,
            columns=('date', 'size', 'joint', 'welder', 'section', 'description'),
            show='headings', height=13, selectmode='extended'
        )
        headings = {
            'date': ('사진 저장일', 105), 'size': ('관경', 65),
            'joint': ('Joint', 70), 'welder': ('용접사', 115),
            'section': ('구간', 80), 'description': ('설명', 360),
        }
        for column, (label, width) in headings.items():
            photo_tree.heading(column, text=label)
            photo_tree.column(
                column, width=width,
                anchor='w' if column == 'description' else 'center'
            )
        photo_tree.pack(side='left', fill='both', expand=True)
        tree_scroll = ttk.Scrollbar(
            tree_frame, orient='vertical', command=photo_tree.yview
        )
        tree_scroll.pack(side='right', fill='y')
        photo_tree.configure(yscrollcommand=tree_scroll.set)

        candidate_by_item = {}
        preferred_items = []
        latest_date = max(item['_source_date'] for item in candidates)
        for photo in candidates:
            item_id = photo_tree.insert(
                '', tk.END,
                values=(
                    photo['_source_date'], photo.get('_size', ''),
                    photo.get('joint_no', ''),
                    self._normalize_daily_photo_welder(photo.get('welder', '')),
                    photo.get('section', ''), photo.get('description', ''),
                ),
            )
            candidate_by_item[item_id] = photo
            if photo['_source_date'] == preferred_date:
                preferred_items.append(item_id)

        initial_items = preferred_items or [
            item_id for item_id, photo in candidate_by_item.items()
            if photo['_source_date'] == latest_date
        ]
        photo_tree.selection_set(*initial_items)
        photo_tree.focus(initial_items[0])
        photo_tree.see(initial_items[0])

        status_var = tk.StringVar()
        status_label = ttk.Label(
            footer_frame, textvariable=status_var, anchor='w'
        )
        status_label.pack(fill='x', pady=(2, 4))

        def update_selection_status(event=None):
            selected_count = len(photo_tree.selection())
            status_var.set(
                f"선택 {selected_count}장 / 전체 {len(candidates)}장"
            )

        update_selection_status()
        photo_tree.bind('<<TreeviewSelect>>', update_selection_status)

        result = {'photos': []}

        def confirm_selection(event=None):
            selected = set(photo_tree.selection())
            if not selected:
                messagebox.showwarning(
                    "사진 선택", "불러올 사진을 한 장 이상 선택하세요.",
                    parent=dialog,
                )
                return
            result['photos'] = [
                candidate_by_item[item_id]
                for item_id in photo_tree.get_children()
                if item_id in selected
            ]
            dialog.destroy()

        button_frame = ttk.Frame(footer_frame)
        button_frame.pack(fill='x')
        ttk.Button(
            button_frame, text="전체 선택",
            command=lambda: photo_tree.selection_set(*photo_tree.get_children())
        ).pack(side='left')
        ttk.Button(
            button_frame, text="선택 해제",
            command=lambda: photo_tree.selection_remove(*photo_tree.selection())
        ).pack(side='left', padx=5)
        ttk.Button(
            button_frame, text="불러오기", command=confirm_selection
        ).pack(side='right', padx=(5, 0))
        ttk.Button(
            button_frame, text="취소", command=dialog.destroy
        ).pack(side='right')
        dialog.protocol('WM_DELETE_WINDOW', dialog.destroy)
        dialog.wait_window()
        return result['photos']

    def _import_daily_work_photos(self):
        """검사일자/검사종류에 맞는 중앙지사 작업일보 사진을 추가한다."""
        inspect_date = self.photo_inspect_date.get().strip()
        try:
            datetime.datetime.strptime(inspect_date, "%Y-%m-%d")
        except ValueError:
            messagebox.showwarning(
                "날짜 확인", "검사 일자를 YYYY-MM-DD 형식으로 입력하세요."
            )
            return

        history_path = self._central_daily_history_path()
        if not history_path or not os.path.isfile(history_path):
            messagebox.showwarning(
                "작업일보 없음",
                "중앙지사 작업일보 저장 파일을 찾을 수 없습니다.\n"
                f"확인 경로: {history_path or '(경로 확인 실패)'}",
            )
            return

        try:
            with open(history_path, 'r', encoding='utf-8') as history_file:
                history = json.load(history_file)
        except (OSError, json.JSONDecodeError) as exc:
            messagebox.showerror(
                "불러오기 오류", f"작업일보를 읽을 수 없습니다.\n{exc}"
            )
            return

        inspect_type = self.photo_inspect_type.get().strip().upper()
        matching_photos = self._select_daily_work_photos(
            history, inspect_type, inspect_date
        )
        if not matching_photos:
            return

        history_dir = os.path.dirname(history_path)
        existing_keys = {
            self._photo_metadata_key(path) for path in self.photo_selected_files
        }
        added_count = 0
        duplicate_count = 0
        missing_files = []
        selected_dates = set()

        for photo in matching_photos:
            source_date = str(photo.get('_source_date', '')).strip()
            if source_date:
                selected_dates.add(source_date)
            stored_path = str(photo.get('file_path', '')).strip()
            if not stored_path:
                missing_files.append('(저장 경로 없음)')
                continue
            if os.path.isabs(stored_path):
                image_path = os.path.normpath(stored_path)
            else:
                image_path = os.path.normpath(os.path.join(history_dir, stored_path))
            image_path = os.path.abspath(image_path)
            metadata_key = self._photo_metadata_key(image_path)

            if not os.path.isfile(image_path):
                missing_files.append(os.path.basename(image_path) or stored_path)
                continue

            self.photo_item_metadata[metadata_key] = {
                'source': 'central_daily_work_log',
                'date': source_date,
                'process': str(photo.get('process', '')).strip(),
                'section': str(photo.get('section', '')).strip(),
                'line_no': str(photo.get('line_no', '')).strip(),
                'size': str(photo.get('_size', '')).strip(),
                'joint_no': str(photo.get('joint_no', '')).strip(),
                'welder': self._normalize_daily_photo_welder(
                    photo.get('welder', '')
                ),
                'location': str(photo.get('location', '')).strip(),
                'description': str(photo.get('description', '')).strip(),
            }
            if metadata_key in existing_keys:
                duplicate_count += 1
                continue

            self.photo_selected_files.append(image_path)
            self.photo_listbox.insert(tk.END, image_path)
            existing_keys.add(metadata_key)
            added_count += 1

        self._refresh_photo_list()
        self._update_photo_preview()
        self.save_settings()
        selected_dates_text = ", ".join(sorted(selected_dates)) or "확인 불가"
        summary = (
            f"사진 저장일: {selected_dates_text} / {inspect_type}\n"
            f"보고서 검사일자: {inspect_date}\n"
            f"추가: {added_count}장\n"
            f"중복 제외: {duplicate_count}장\n"
            f"파일 없음: {len(missing_files)}장"
        )
        if missing_files:
            summary += "\n\n찾지 못한 파일:\n" + "\n".join(missing_files[:5])
            if len(missing_files) > 5:
                summary += f"\n외 {len(missing_files) - 5}장"
        self.log(f"[PhotoLog] 작업일보 연동: {summary.replace(chr(10), ' / ')}")
        messagebox.showinfo("작업일보 사진 불러오기", summary)

    def _sort_photo_list(self):
        from site_apps.central.src.reports.ndt_history_import import sort_photo_paths
        self.photo_selected_files[:] = sort_photo_paths(self.photo_selected_files)
        self._refresh_photo_list()
        self.save_settings()

    def _refresh_photo_list(self):
        self.photo_listbox.delete(0, tk.END)
        for path in self.photo_selected_files:
            self.photo_listbox.insert(tk.END, path)

    def _move_photo_selected(self, direction):
        selected = set(self.photo_listbox.curselection())
        if not selected:
            return
        indices = sorted(selected, reverse=direction > 0)
        for index in indices:
            target = index + direction
            if 0 <= target < len(self.photo_selected_files) and target not in selected:
                self.photo_selected_files[index], self.photo_selected_files[target] = (
                    self.photo_selected_files[target], self.photo_selected_files[index])
                selected.remove(index)
                selected.add(target)
        self._refresh_photo_list()
        for index in sorted(selected):
            self.photo_listbox.selection_set(index)
        self.photo_listbox.see(min(selected))
        self._update_photo_preview()
        self.save_settings()

    def _photo_last_folder(self):
        for folder in (self.config.get('PHOTO_LAST_IMPORT_DIR', ''),
                       *(os.path.dirname(path) for path in reversed(self.photo_selected_files))):
            if folder and os.path.isdir(folder):
                return folder
        return os.getcwd()

    def _add_photo_files(self):
        files = filedialog.askopenfilenames(parent=self.root, initialdir=self._photo_last_folder(),
                                           filetypes=[("Image files", "*.png;*.jpg;*.jpeg;*.bmp")])
        if files:
            self.config['PHOTO_LAST_IMPORT_DIR'] = os.path.dirname(files[0])
            for f in files:
                f_norm = os.path.normpath(f)
                if f_norm not in self.photo_selected_files:
                    self.photo_selected_files.append(f_norm)
                    if hasattr(self, 'photo_listbox'):
                        self.photo_listbox.insert(tk.END, f_norm)
            self._refresh_photo_list()
            self.save_settings()
            self.log(f"[PhotoLog] {len(files)}개 파일 추가 시도")

    def _add_photo_folder(self):
        folder = filedialog.askdirectory(parent=self.root, title='사진 추가 폴더 선택',
                                        initialdir=self._photo_last_folder(), mustexist=True)
        if folder:
            self.config['PHOTO_LAST_IMPORT_DIR'] = folder
            added_count = 0
            for ext in ('*.png', '*.jpg', '*.jpeg', '*.bmp'):
                for f in glob.glob(os.path.join(folder, ext)):
                    f_norm = os.path.normpath(f)
                    if f_norm not in self.photo_selected_files:
                        self.photo_selected_files.append(f_norm)
                        if hasattr(self, 'photo_listbox'):
                            self.photo_listbox.insert(tk.END, f_norm)
                        added_count += 1
            self._refresh_photo_list()
            self.save_settings()
            self.log(f"[PhotoLog] 폴더에서 {added_count}개 파일 추가 완료")

    def _remove_photo_selected(self):
        idxs = list(self.photo_listbox.curselection())
        for i in reversed(idxs):
            path = self.photo_listbox.get(i)
            if path in self.photo_selected_files:
                self.photo_selected_files.remove(path)
            self.photo_item_metadata.pop(self._photo_metadata_key(path), None)
            self.photo_listbox.delete(i)

    def _clear_photo_all(self):
        self.photo_selected_files.clear()
        self.photo_item_metadata.clear()
        self.photo_listbox.delete(0, tk.END)
        self.log("[PhotoLog] 리스트 초기화 완료")
        self._update_photo_preview()

    def _photo_log_auto_row_height(self, num_rows, desc_row_height):
        """Return a photo-row height that uses the printable A4 height evenly."""
        def value_as_float(value, default):
            try:
                return float(value)
            except (TypeError, ValueError):
                return float(default)

        rows = max(1, int(num_rows))
        margin_top = value_as_float(self.photo_margin_top_var.get(), 0.5)
        margin_bottom = value_as_float(self.photo_margin_bottom_var.get(), 0.5)

        # A4 is about 842 pt high. The 40 pt allowance covers the fixed title
        # area and Excel's print rounding. Keeping this allowance small extends
        # all photo rows downward without allowing the last row onto a new page.
        available_height = 842.0 - ((margin_top + margin_bottom) * 72.0) - 40.0
        return max(
            20.0,
            (available_height - (rows * float(desc_row_height))) / float(rows),
        )

    def _update_photo_preview(self, *args):
        def safe_float(val, default=0.0):
            try: return float(val.strip())
            except: return default
        if not hasattr(self, 'photo_preview_canvas') or not self.photo_preview_canvas.winfo_exists():
            return

        self.photo_preview_canvas.delete("all")

        selected_idxs = self.photo_listbox.curselection()
        if selected_idxs:
            img_path = self.photo_listbox.get(selected_idxs[0])
        elif self.photo_listbox.size() > 0:
            img_path = self.photo_listbox.get(0)
        else:
            self.photo_preview_canvas.create_text(250, 100, text="리스트에서 사진을 선택하면 미리보기가 나타납니다.", fill="gray", font=("Malgun Gothic", 10))
            return

        if not os.path.exists(img_path):
            self.photo_preview_canvas.create_text(250, 100, text="⚠️ 선택한 사진 파일이 컴퓨터에 존재하지 않습니다.\n(파일이 삭제되었거나 이동되었을 수 있습니다.)", fill="red", font=("Malgun Gothic", 10, "bold"), justify="center")
            return

        try:

            GRID_COLS = 12
            unit_per_grid = (float(self.photo_cell_width_var.get()) * 2) / GRID_COLS
            CELL_ROW_HEIGHT = float(self.photo_cell_height_var.get())

            # A4 auto-fit uses the same row-height calculation as Excel output.
            if self.photo_auto_fit_page_var.get():
                desc_h = safe_float(self.photo_desc_height_var.get(), 15.0)
                preview_rows = max(1, int(self.photo_rows_per_page.get()))
                CELL_ROW_HEIGHT = self._photo_log_auto_row_height(preview_rows, desc_h)

            WIDTH_PCT = safe_float(self.photo_width_pct_var.get(), 100.0) / 100.0
            HEIGHT_PCT = max(0.01, safe_float(self.photo_height_pct_var.get(), 100.0) / 100.0)
            PIXEL_ADJ = safe_float(self.photo_width_pixel_adj_var.get(), 0.0)
            SHIFT_X = safe_float(self.photo_shift_x_var.get(), 0.0)
            SHIFT_Y = safe_float(self.photo_shift_y_var.get(), 0.0)
            INTERNAL_FACTOR = 8.0 # Korean Excel Font Hotfix
            num_cols = int(self.photo_cols_per_row.get())

            unit_per_px = int((unit_per_grid * INTERNAL_FACTOR) + 5)
            if self.photo_auto_fit_page_var.get():
                m_l = safe_float(self.photo_margin_left_var.get(), 0.5)
                m_r = safe_float(self.photo_margin_right_var.get(), 0.5)
                unit_per_px = int(((595.0 - (m_l + m_r) * 72.0) * 1.33333) / GRID_COLS)

            if num_cols == 1: CELL_WIDTH_PX = unit_per_px * 12
            elif num_cols == 2: CELL_WIDTH_PX = unit_per_px * 6
            elif num_cols == 3: CELL_WIDTH_PX = unit_per_px * 4
            else: CELL_WIDTH_PX = unit_per_px * 3

            SAFE_WIDTH = CELL_WIDTH_PX * WIDTH_PCT + PIXEL_ADJ
            ROW_PT_TO_PX = 1.33333
            CELL_HEIGHT_PX = (CELL_ROW_HEIGHT * 1.33333) - 2

            from PIL import Image as PILImage, ImageOps, ImageTk
            with PILImage.open(img_path) as img:
                img = ImageOps.exif_transpose(img)
                img_w, img_h = img.size

                if self._should_auto_rotate_photo(img_path) and img_h > img_w:
                    img = img.rotate(90, expand=True)
                    img_w, img_h = img.size

                HEIGHT_ADJ = safe_float(self.photo_height_pixel_adj_var.get(), 10.0)
                if self.photo_auto_fit_page_var.get():
                    scale = min(SAFE_WIDTH / img_w, CELL_HEIGHT_PX / img_h)
                    x_scale = y_scale = scale
                    final_row_h_px = CELL_ROW_HEIGHT * ROW_PT_TO_PX
                elif self.photo_fit_width_var.get():
                    x_scale = y_scale = SAFE_WIDTH / img_w
                    req_h_px = (img_h * y_scale) + HEIGHT_ADJ
                    final_row_h_px = (req_h_px / ROW_PT_TO_PX) * ROW_PT_TO_PX
                elif not self.photo_keep_aspect.get():
                    x_scale = SAFE_WIDTH / img_w
                    y_scale = CELL_HEIGHT_PX / img_h
                    final_row_h_px = (CELL_ROW_HEIGHT * ROW_PT_TO_PX) + HEIGHT_ADJ
                else:
                    scale = min(SAFE_WIDTH / img_w, CELL_HEIGHT_PX / img_h)
                    x_scale = y_scale = scale
                    final_row_h_px = (CELL_ROW_HEIGHT * ROW_PT_TO_PX) + HEIGHT_ADJ

                # Keep the calculated width fixed and adjust only image height.
                y_scale = min(y_scale * HEIGHT_PCT, CELL_HEIGHT_PX / img_h)

                x_off = round(((CELL_WIDTH_PX - (img_w * x_scale)) / 2) + SHIFT_X)
                y_off = round(((final_row_h_px - (img_h * y_scale)) / 2) + SHIFT_Y)

                preview_w = self.photo_preview_canvas.winfo_width()
                preview_h = self.photo_preview_canvas.winfo_height()
                if preview_w < 10: preview_w = 600
                if preview_h < 10: preview_h = 300

                # [NEW] 화면에서 더 크게 보일 수 있도록 1.25배 확대하되, 캔버스 높이를 넘어가면 자동 축소(줌아웃)
                VISUAL_ZOOM = 1.25
                if (final_row_h_px * VISUAL_ZOOM) > (preview_h - 20):
                    VISUAL_ZOOM = (preview_h - 20) / final_row_h_px

                center_x = preview_w / 2
                center_y = preview_h / 2

                draw_cell_w = CELL_WIDTH_PX * VISUAL_ZOOM
                draw_cell_h = final_row_h_px * VISUAL_ZOOM

                cell_left = center_x - (draw_cell_w / 2)
                cell_top = center_y - (draw_cell_h / 2)

                # [NEW] Overflow Check (1px tolerance for floating point math)
                actual_w = img_w * x_scale
                actual_h = img_h * y_scale
                is_overflow = False
                if x_off < -1 or y_off < -1 or (x_off + actual_w) > CELL_WIDTH_PX + 1 or (y_off + actual_h) > final_row_h_px + 1:
                    is_overflow = True

                box_color = "red" if is_overflow else "green"
                box_width = 3 if is_overflow else 2
                box_dash = () if is_overflow else (4, 4)

                status_msg = "⚠️ 경고: 사진이 박스를 벗어났습니다! (엑셀 레이아웃이 깨집니다)" if is_overflow else "✅ 정상: 사진이 박스 안에 안전하게 들어갑니다."
                status_color = "red" if is_overflow else "#16a34a" # Green

                # Calculate image rendering
                render_w = int(img_w * x_scale * VISUAL_ZOOM)
                render_h = int(img_h * y_scale * VISUAL_ZOOM)

                draw_x_off = x_off * VISUAL_ZOOM
                draw_y_off = y_off * VISUAL_ZOOM

                # 1. Draw Image First (so it stays in background if it overflows)
                if render_w > 0 and render_h > 0:
                    render_img = img.resize((render_w, render_h), PILImage.Resampling.LANCZOS)
                    self._preview_photo_img = ImageTk.PhotoImage(render_img)
                    self.photo_preview_canvas.create_image(
                        cell_left + draw_x_off, cell_top + draw_y_off,
                        image=self._preview_photo_img, anchor="nw"
                    )

                # 2. Draw Excel Cell Boundary On Top
                self.photo_preview_canvas.create_rectangle(
                    cell_left, cell_top, cell_left + draw_cell_w, cell_top + draw_cell_h,
                    outline=box_color, width=box_width, dash=box_dash
                )

                # 3. Draw Status Message On Top
                self.photo_preview_canvas.create_text(
                    cell_left, cell_top - 15, 
                    text=status_msg, 
                    fill=status_color, anchor="w", font=("Malgun Gothic", 10, "bold")
                )
                self.photo_preview_canvas.create_text(
                    cell_left, cell_top + draw_cell_h + 10, 
                    text=f"현재 테두리: 엑셀 셀 실제 영역 (미리보기 {VISUAL_ZOOM}배 확대 렌더링 중)", 
                    fill="gray", anchor="w", font=("Malgun Gothic", 8)
                )
        except Exception as e:
            self.photo_preview_canvas.create_text(250, 100, text=f"미리보기 렌더링 오류: {e}", fill="red")

    def start_photo_generation(self):
        if not self.photo_selected_files:
            messagebox.showwarning("경고", "리포트에 포함할 이미지를 먼저 선택해주세요.")
            return

        selected_idxs = self.photo_listbox.curselection()
        if selected_idxs:
            target_files = [self.photo_listbox.get(i) for i in selected_idxs]
        else:
            target_files = list(self.photo_selected_files)

        self.save_settings() # Auto-save before generation
        threading.Thread(target=self.generate_photo_report, args=(target_files,), daemon=True).start()

    def generate_photo_report(self, target_files=None):
        def safe_float(val, default=0.0):
            try: return float(val.strip())
            except: return default
        try:
            if target_files is None:
                if not self.photo_selected_files: return
                image_files = list(self.photo_selected_files)
            else:
                if not target_files: return
                image_files = target_files

            # [REMEMBER] Load last save dir from class variable
            initial_folder = self.last_photo_save_dir if self.last_photo_save_dir and os.path.isdir(self.last_photo_save_dir) else os.path.dirname(image_files[0])

            self._sync_photo_output_name()
            default_name = self.photo_output_name.get()
            if not default_name.endswith(".xlsx"):
                default_name += ".xlsx"

            output_folder = filedialog.askdirectory(
                parent=self.root,
                title="사진대장 저장 폴더 선택",
                initialdir=initial_folder,
                mustexist=True,
            )

            if not output_folder:
                self.log("[PhotoLog] 작업을 취소했습니다.")
                return
            from site_apps.central.src.reports.report_issue_history import report_output_path
            output_path = report_output_path(output_folder, os.path.splitext(default_name)[0], '.xlsx')

            # [REMEMBER] Update last save directory
            self.last_photo_save_dir = os.path.dirname(output_path)
            self.save_settings() # Silent save

            self.progress["value"] = 0
            self.log("[PhotoLog] 작업을 시작합니다..")

            workbook = xlsxwriter.Workbook(output_path)
            worksheet = workbook.add_worksheet('사진대장')

            # Page Setup
            worksheet.set_paper(9) # A4
            worksheet.set_portrait()
            worksheet.center_horizontally()
            worksheet.center_vertically()

            try:
                m_t = float(self.photo_margin_top_var.get())
                m_b = float(self.photo_margin_bottom_var.get())
                m_l = float(self.photo_margin_left_var.get())
                m_r = float(self.photo_margin_right_var.get())
                worksheet.set_margins(left=m_l, right=m_r, top=m_t, bottom=m_b)
            except:
                worksheet.set_margins(left=0.4, right=0.4, top=0.5, bottom=0.5)

            s_page = self.photo_start_page.get().strip()
            t_pages = self.photo_total_pages.get().strip()

            # 페이지 번호는 제목보다 한 단계 작은 10pt로 표시하고, 표 바로 위에
            # 오도록 기본 헤더 여백(0.30in)보다 약간 아래로 배치한다.
            base_header = '&"바탕"&10Page   &P   of   &N'
            header_text = '&R&10 \n\n' + base_header + ('\xa0' * 7)

            if s_page.isdigit():
                worksheet.set_start_page(int(s_page))
            if t_pages.isdigit():
                base_header = f'&"바탕"&10Page   &P   of   {t_pages}'
                header_text = '&R&10 \n\n' + base_header + ('\xa0' * 7)

            worksheet.set_header(header_text, {'margin': 0.36})
            worksheet.set_footer('')
            worksheet.repeat_rows(0, 4) 

            # Layout Calculation
            num_cols = int(self.photo_cols_per_row.get())
            num_rows = int(self.photo_rows_per_page.get())
            photos_per_page = num_cols * num_rows
            scheduled_photos, used_slot_count, _daily_start_slot = (
                self._schedule_photo_output(image_files, photos_per_page)
            )
            image_files = [path for path, _slot in scheduled_photos]
            total_pages = math.ceil(used_slot_count / photos_per_page)
            # 가로(너비)는 무조건 1페이지에 맞추고, 세로(높이)는 자동(0)으로 두어 
            # 사진이 옆 페이지로 밀려나는 현상을 완벽히 방지합니다.
            worksheet.fit_to_pages(1, 0)

            # Formats
            title_format = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'shrink': True})
            company_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_size': 9, 'text_wrap': True})
            center_border = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_size': 10})
            bold_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
            desc_format = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1, 'font_size': 10, 'shrink': True, 'text_wrap': False, 'indent': 1})

            # Fixed 6-column Grid System
            GRID_COLS = 12
            unit_per_grid = (float(self.photo_cell_width_var.get()) * 2) / GRID_COLS

            # [HOTFIX] 한국어 엑셀(맑은 고딕) 렌더링 배율 보정 (8.0)
            INTERNAL_FACTOR = 8.0

            # 엑셀의 가변적인 문자 단위 너비 대신, 파이썬이 계산한 '정확한 픽셀 너비'로 엑셀 셀을 강제 고정합니다.
            unit_per_px = int((unit_per_grid * INTERNAL_FACTOR) + 5)
            if self.photo_auto_fit_page_var.get():
                m_l = safe_float(self.photo_margin_left_var.get(), 0.5)
                m_r = safe_float(self.photo_margin_right_var.get(), 0.5)
                unit_per_px = int(((595.0 - (m_l + m_r) * 72.0) * 1.33333) / GRID_COLS)
            worksheet.set_column_pixels(0, GRID_COLS - 1, unit_per_px)

            CELL_ROW_HEIGHT = float(self.photo_cell_height_var.get())

            # [REFINED] Percentage-Based Precision Scaling

            WIDTH_PCT = safe_float(self.photo_width_pct_var.get(), 100.0) / 100.0
            HEIGHT_PCT = max(0.01, safe_float(self.photo_height_pct_var.get(), 100.0) / 100.0)
            PIXEL_ADJ = safe_float(self.photo_width_pixel_adj_var.get(), 0.0)
            SHIFT_X = safe_float(self.photo_shift_x_var.get(), 0.0)
            SHIFT_Y = safe_float(self.photo_shift_y_var.get(), 0.0)

            if num_cols == 1:
                photo_col_spans = [(0, GRID_COLS - 1)]
                CELL_WIDTH_PX = unit_per_px * 12
            elif num_cols == 2:
                photo_col_spans = [(0, 5), (6, 11)]
                CELL_WIDTH_PX = unit_per_px * 6
            elif num_cols == 3:
                photo_col_spans = [(0, 3), (4, 7), (8, 11)]
                CELL_WIDTH_PX = unit_per_px * 4
            else: # 4 Columns
                photo_col_spans = [(0, 2), (3, 5), (6, 8), (9, 11)]
                CELL_WIDTH_PX = unit_per_px * 3

            # Target width based on percentage and padding (removed hardcoded 10px margin for edge-to-edge fit)
            SAFE_WIDTH = CELL_WIDTH_PX * WIDTH_PCT + PIXEL_ADJ

            # Detailed Logging for user calibration
            self.log(f"[PhotoLog] 설정: 너비비율={WIDTH_PCT*100:.1f}%, 보정={PIXEL_ADJ}px")
            self.log(f"[PhotoLog] 이동: 가로={SHIFT_X}px, 세로={SHIFT_Y}px")
            self.log(f"[PhotoLog] 계산결과: 셀너비 {CELL_WIDTH_PX}px -> 사진너비 {SAFE_WIDTH:.1f}px")

            worksheet.set_row(0, 30)
            worksheet.merge_range(0, 0, 0, GRID_COLS-1, self.photo_report_title.get(), title_format)

            company_text = "서   울   檢   査   株   式   會   社\nSEOUL INSPECTION & TESTING Co., Ltd.\nTEL : (02) 552-1112   FAX : (02) 2058-0720"
            worksheet.merge_range(1, 0, 3, 5, company_text, company_format)

            # Logo Insertion
            logo_f = self.photo_logo_path.get().strip()

            total_header_h = 45 # 15 * 3 rows
            for r in range(1, 4): worksheet.set_row(r, 15) 

            if logo_f and os.path.exists(logo_f):
                try:
                    with PILImage.open(logo_f) as img:
                        w, h = img.size
                        max_w_logo = float(self.photo_logo_width_var.get())
                        mx = float(self.photo_logo_x_var.get())
                        my = float(self.photo_logo_y_var.get())

                        scale_x = max_w_logo / w
                        try:
                            # [FIX] Apply independent height scale if user provided it
                            max_h_logo = float(self.photo_logo_height_var.get().strip())
                            scale_y = max_h_logo / h
                        except:
                            scale_y = scale_x

                        # [FIX] Anchor at A1 with 40px base offset (row 0 height) to allow free negative movement
                        y_offset = 40 + int(my)
                        worksheet.insert_image('A1', logo_f, {'x_scale': scale_x, 'y_scale': scale_y, 'x_offset': int(mx), 'y_offset': y_offset, 'object_position': 1})
                except Exception as e:
                    self.log(f"[PhotoLog] 로고 삽입 중 오류: {e}")
            else:
                # [NEW] Use Embedded Base64 Logo if no file is provided
                try:
                    logo_data = base64.b64decode(DEFAULT_SITCO_LOGO_B64)
                    logo_buffer = io.BytesIO(logo_data)
                    with PILImage.open(logo_buffer) as img:
                        w, h = img.size
                        max_w_logo = float(self.photo_logo_width_var.get())
                        mx = float(self.photo_logo_x_var.get())
                        my = float(self.photo_logo_y_var.get())

                        scale_x = max_w_logo / w
                        try:
                            max_h_logo = float(self.photo_logo_height_var.get().strip())
                            scale_y = max_h_logo / h
                        except:
                            scale_y = scale_x

                        # [FIX] Anchor at A1 with 40px base offset (row 0 height) to allow free negative movement
                        y_offset = 40 + int(my)
                        worksheet.insert_image('A1', 'sitco_logo.png', {'image_data': logo_buffer, 'x_scale': scale_x, 'y_scale': scale_y, 'x_offset': int(mx), 'y_offset': y_offset, 'object_position': 1})
                except Exception as e:
                    self.log(f"[PhotoLog] 내장 로고 삽입 실패: {e}")

            worksheet.merge_range(1, 6, 1, 11, f"발주처: {self.photo_orderer.get()}", center_border)
            worksheet.merge_range(2, 6, 2, 11, f"REPORT NO: {self.photo_report_no.get()}", center_border)

            # PAUT 보고서와 동일한 투명 도장을 사진대장 정보 영역에도 표시한다.
            stamp_data = self._get_photo_log_stamp()
            if stamp_data:
                try:
                    stamp_buffer = io.BytesIO(stamp_data)
                    with PILImage.open(io.BytesIO(stamp_data)) as stamp_image:
                        stamp_w, stamp_h = stamp_image.size
                    stamp_size = 72.0
                    worksheet.insert_image(
                        "E2", "transparent_stamp.png",
                        {
                            "image_data": stamp_buffer,
                            "x_scale": stamp_size / stamp_w,
                            "y_scale": stamp_size / stamp_h,
                            "x_offset": 45,
                            "y_offset": 0,
                            "object_position": 1,
                        },
                    )
                except Exception as exc:
                    self.log(f"[PhotoLog] 도장 삽입 실패: {exc}")
            worksheet.merge_range(3, 6, 3, 11, f"검사일자: {self.photo_inspect_date.get()}", center_border)

            worksheet.set_row(4, 25)
            worksheet.merge_range(4, 0, 4, GRID_COLS-1, "PHOTO LOG (사진 대장)", bold_format)

            page_breaks = [
                5 + (page_index * num_rows * 2)
                for page_index in range(1, total_pages)
            ]
            num_rows = int(self.photo_rows_per_page.get())
            photos_per_page = num_cols * num_rows
            DESC_ROW_HEIGHT = float(self.photo_desc_height_var.get())

            CELL_ROW_HEIGHT = float(self.photo_cell_height_var.get())
            if self.photo_auto_fit_page_var.get():
                CELL_ROW_HEIGHT = self._photo_log_auto_row_height(
                    num_rows, DESC_ROW_HEIGHT
                )

            total = len(image_files)
            current_row_max_h_pt = CELL_ROW_HEIGHT
            CELL_HEIGHT_PX = (CELL_ROW_HEIGHT * 1.33333) - 2
            ROW_PT_TO_PX = 1.33333

            used_photo_slots = set()
            for i, (img_path, slot_index) in enumerate(scheduled_photos):
                page_index = slot_index // photos_per_page
                slot_in_page = slot_index % photos_per_page
                photo_row_in_page = slot_in_page // num_cols
                col_ptr = slot_in_page % num_cols
                is_page_start_photo = page_index > 0 and photo_row_in_page == 0
                logical_row_slot = (page_index * num_rows) + photo_row_in_page
                used_photo_slots.add((logical_row_slot, col_ptr))
                row = 5 + (logical_row_slot * 2)

                # Reset max height for a new row of photos
                if col_ptr == 0:
                    current_row_max_h_pt = 0 if self.photo_fit_width_var.get() else CELL_ROW_HEIGHT

                try:
                    with PILImage.open(img_path) as img:
                        img = ImageOps.exif_transpose(img)
                        img_w, img_h = img.size

                        # [NEW] Auto-rotate vertical images to horizontal if option enabled
                        if self._should_auto_rotate_photo(img_path) and img_h > img_w:
                            img = img.rotate(90, expand=True)
                            img_w, img_h = img.size
                            self.log(f"[PhotoLog] 자동 회전 적용: {os.path.basename(img_path)}")

                        # [FIX] Prepare image data to preserve rotation and apply User DPI for Excel
                        img_buffer = io.BytesIO()
                        try:
                            user_dpi = int(self.photo_dpi_var.get().strip())
                            if user_dpi <= 0: user_dpi = 96
                        except:
                            user_dpi = 96
                        img.save(img_buffer, format='PNG', dpi=(user_dpi, user_dpi))
                        img_buffer.seek(0)

                        c_start, c_end = photo_col_spans[col_ptr]
                        if c_start != c_end: worksheet.merge_range(row, c_start, row, c_end, "", center_border)

                        # [REFINED] Scaling Logic with 'Fit to Width' support
                        HEIGHT_ADJ = safe_float(self.photo_height_pixel_adj_var.get(), 10.0)
                        if self.photo_auto_fit_page_var.get():
                            scale = min(SAFE_WIDTH / img_w, CELL_HEIGHT_PX / img_h)
                            x_scale = y_scale = scale
                            current_row_max_h_pt = CELL_ROW_HEIGHT
                        elif self.photo_fit_width_var.get():
                            x_scale = SAFE_WIDTH / img_w
                            y_scale = x_scale
                            req_h_px = (img_h * y_scale) + HEIGHT_ADJ # Customizable Buffer
                            req_h_pt = req_h_px / ROW_PT_TO_PX
                            current_row_max_h_pt = max(current_row_max_h_pt, req_h_pt)
                        elif not self.photo_keep_aspect.get():
                            x_scale = SAFE_WIDTH / img_w
                            y_scale = CELL_HEIGHT_PX / img_h
                            current_row_max_h_pt = CELL_ROW_HEIGHT + (HEIGHT_ADJ / ROW_PT_TO_PX)
                        else:
                            scale = min(SAFE_WIDTH / img_w, CELL_HEIGHT_PX / img_h)
                            x_scale = y_scale = scale
                            current_row_max_h_pt = CELL_ROW_HEIGHT + (HEIGHT_ADJ / ROW_PT_TO_PX)

                        # Preserve horizontal size while stretching only vertically.
                        y_scale = min(y_scale * HEIGHT_PCT, CELL_HEIGHT_PX / img_h)

                        # Apply the potentially updated row height
                        worksheet.set_row(row, current_row_max_h_pt)



                        # Re-calculate Y offset based on the final row height
                        final_row_h_px = current_row_max_h_pt * ROW_PT_TO_PX

                        # [UNIFIED] Uniform Margin Distribution with Manual Shift (High Precision)
                        x_off_float = ((CELL_WIDTH_PX - (img_w * x_scale)) / 2) + SHIFT_X
                        y_off_float = ((final_row_h_px - (img_h * y_scale)) / 2) + SHIFT_Y

                        # 확대율과 수동 이동값 때문에 이미지가 행 경계를 1~2px
                        # 넘어가면 다음/이전 인쇄 페이지에 가는 선으로 나타난다.
                        # 출력 이미지가 사진 셀 안에 완전히 들어오도록 높이와
                        # 오프셋을 마지막에 다시 제한한다.
                        safe_vertical_gap = 2.0
                        # Excel이 페이지 경계에 정확히 닿은 이미지의 첫 픽셀을
                        # 이전 페이지에 그리는 현상을 피한다.
                        page_top_gap = 6.0 if is_page_start_photo else 0.0
                        max_image_height = max(
                            1.0,
                            final_row_h_px - safe_vertical_gap - page_top_gap,
                        )
                        if (img_h * y_scale) > max_image_height:
                            y_scale = max_image_height / img_h
                            y_off_float = (
                                (final_row_h_px - (img_h * y_scale)) / 2
                            ) + SHIFT_Y
                        max_y_offset = max(
                            page_top_gap,
                            final_row_h_px - (img_h * y_scale) - 1.0,
                        )
                        y_off_float = min(
                            max(y_off_float, page_top_gap), max_y_offset
                        )

                        x_off = round(x_off_float)
                        y_off = round(y_off_float)

                        # Log precision offsets for user calibration
                        self.log(f"[PhotoLog] 사진 배치: {os.path.basename(img_path)}")
                        self.log(f"           - 가로위치: {x_off_float:.2f} -> {x_off}px")
                        self.log(f"           - 세로위치: {y_off_float:.2f} -> {y_off}px")

                        worksheet.insert_image(row, c_start, img_path, {
                            'image_data': img_buffer,
                            'x_scale': x_scale, 'y_scale': y_scale, 
                            'x_offset': x_off, 'y_offset': y_off, 
                            'object_position': 2 # Move but don't size with cells
                        })
                except Exception as e:
                    self.log(f"[Error] {os.path.basename(img_path)}: {e}")

                name = self._daily_photo_description(img_path)
                worksheet.set_row(row + 1, DESC_ROW_HEIGHT)
                c_start, c_end = photo_col_spans[col_ptr]
                worksheet.merge_range(row+1, c_start, row+1, c_end, f"사진 설명: {name}", desc_format)

                self.progress["value"] = ((i + 1) / total) * 100
                self.log(f"[PhotoLog] 처리 중.. ({i+1}/{total})")

            # 모든 페이지의 빈 사진행도 같은 높이로 유지하고 테두리를 그려 표 형태를 완성한다.
            full_photo_rows = total_pages * num_rows
            for slot in range(full_photo_rows):
                padding_row = 5 + (slot * 2)
                row_has_any_photo = any((slot, c) in used_photo_slots for c in range(num_cols))
                if not row_has_any_photo:
                    worksheet.set_row(padding_row, CELL_ROW_HEIGHT)
                    worksheet.set_row(padding_row + 1, DESC_ROW_HEIGHT)

                # 빈 슬롯마다 테두리 서식을 적용하여 표 형태를 채움 (부분 빈 행 포함)
                for col_ptr in range(num_cols):
                    if (slot, col_ptr) in used_photo_slots:
                        continue
                    c_start, c_end = photo_col_spans[col_ptr]
                    if c_start != c_end:
                        worksheet.merge_range(padding_row, c_start, padding_row, c_end, "", center_border)
                        worksheet.merge_range(padding_row + 1, c_start, padding_row + 1, c_end, "사진 설명: ", desc_format)
                    else:
                        worksheet.write(padding_row, c_start, "", center_border)
                        worksheet.write(padding_row + 1, c_start, "사진 설명: ", desc_format)

            full_print_end_row = 5 + (full_photo_rows * 2) - 1
            worksheet.print_area(0, 0, full_print_end_row, GRID_COLS - 1)

            if page_breaks: worksheet.set_h_pagebreaks(page_breaks)
            workbook.close()
            self.log(f"[PhotoLog] 완료: {os.path.basename(output_path)}")
            messagebox.showinfo("성공", f"사진대장이 생성되었습니다.\n{output_path}")

            # [NEW] Open the folder
            try:
                os.startfile(os.path.dirname(output_path))
            except: pass

        except Exception as e:
            self.log(f"[Error] {e}")
            messagebox.showerror("오류", f"작업 중 오류 발생:\n{e}")

