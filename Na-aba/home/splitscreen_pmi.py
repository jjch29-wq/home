import re

file_path = r'c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\src\main.py'

with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Redefine _setup_pmi_ui for Splitscreen
# Left: Settings/Actions, Right: Preview
new_pmi_ui = r'''
    def _setup_pmi_ui(self, parent):
        container = tk.Frame(parent, background="#f9fafb")
        container.pack(fill='both', expand=True)

        # --- Dual Pane Layout (PanedWindow) ---
        pw = tk.PanedWindow(container, orient='horizontal', background="#f9fafb", sashwidth=4, sashpad=0, borderwidth=0)
        pw.pack(fill='both', expand=True)

        # [LEFT] Settings & Actions Pane
        left_pane = tk.Frame(pw, background="#f9fafb", padx=10, pady=10)
        pw.add(left_pane, width=420) # Fixed width for settings to ensure enough space for preview

        # Header
        header_frame = tk.Frame(left_pane, background="#f9fafb")
        header_frame.pack(fill='x', pady=(0, 10))
        tk.Label(header_frame, text="🔬 PMI 관리", font=("Malgun Gothic", 15, "bold"), 
                 background="#f9fafb", foreground="#1e3a8a").pack(side='left')
        tk.Label(header_frame, text=f"v{APP_VERSION}", font=("Arial", 8), 
                 background="#f1f5f9", foreground="#64748b", padx=5, pady=1).pack(side='left', padx=10)

        # 1. File Selection Group
        file_container = ttk.LabelFrame(left_pane, text=" 데이터 및 양식 (Data) ", padding=10)
        file_container.pack(fill='x', pady=(0, 10))

        def _add_very_compact_row(parent_f, label, var, row, is_dir=False, types=None):
            parent_f.columnconfigure(1, weight=1)
            ttk.Label(parent_f, text=label, font=("Malgun Gothic", 8)).grid(row=row, column=0, sticky='e', padx=2, pady=2)
            ttk.Entry(parent_f, textvariable=var, font=("Arial", 9), exportselection=False).grid(row=row, column=1, padx=2, pady=2, sticky='ew')
            cmd = (lambda : self._browse_dir(var)) if is_dir else (lambda : self._browse_file(var, types))
            ttk.Button(parent_f, text="...", width=3, command=cmd).grid(row=row, column=2, padx=2, pady=2)

        _add_very_compact_row(file_container, "로고:", self.logo_folder_path, 0, is_dir=True)
        _add_very_compact_row(file_container, "데이터:", self.target_file_path, 1, types=[("Excel Source", "*.xls;*.xlsx;*.xlsm")])
        _add_very_compact_row(file_container, "양식:", self.template_file_path, 2, types=[("Excel Template", "*.xlsx;*.xlsm")])

        # 2. Configuration Notebook (Cover, Data, Rows - NO Preview tab here)
        config_frame = ttk.LabelFrame(left_pane, text=" 설정 (Config) ", padding=2)
        config_frame.pack(fill='both', expand=True, pady=(0, 10))

        self.tab_notebook = ttk.Notebook(config_frame)
        self.tab_notebook.pack(fill='both', expand=True)

        tab_cover = ttk.Frame(self.tab_notebook, padding=5)
        tab_data = ttk.Frame(self.tab_notebook, padding=5)
        tab_rows = ttk.Frame(self.tab_notebook, padding=5)

        self.tab_notebook.add(tab_cover, text="갑지")
        self.tab_notebook.add(tab_data, text="을지")
        self.tab_notebook.add(tab_rows, text="행 설정")

        self.setting_vars = {}
        # We need to center settings inside these narrower tabs too
        inner_cover = tk.Frame(tab_cover, background="#f9fafb")
        inner_cover.pack(fill='both', expand=True)
        inner_data = tk.Frame(tab_data, background="#f9fafb")
        inner_data.pack(fill='both', expand=True)

        self._create_setting_grid(inner_cover, "COVER")
        self._create_setting_grid(inner_data, "DATA")
        self._create_margin_settings(inner_cover, "COVER")
        self._create_margin_settings(inner_data, "DATA")
        self._create_row_settings(tab_rows)
        
        # Spacer for tabs
        tk.Frame(tab_rows, background="#f9fafb").pack(fill='both', expand=True)

        # 3. Quick Options Section
        action_outer = tk.Frame(left_pane, background="#ffffff", highlightthickness=1, highlightbackground="#d1d5db", padx=10, pady=5)
        action_outer.pack(fill='x', pady=(0, 10))

        # Filter Tag Area (Self-updating)
        filter_box = tk.Frame(action_outer, background="#ffffff")
        filter_box.pack(fill='x', pady=2)
        ttk.Label(filter_box, text="🔍 성분 필터:", background="#ffffff", font=("Malgun Gothic", 9, "bold")).pack(side='left', padx=(0, 5))
        self.filter_container = tk.Frame(filter_box, background="#ffffff")
        self.filter_container.pack(side='left', fill='x', expand=True)
        
        def add_filter():
            self.element_filters.append({'key': tk.StringVar(value=''), 'min': tk.StringVar(value=''), 'max': tk.StringVar(value='')})
            self._update_pmi_filter_ui()
        ttk.Button(filter_box, text="+", width=2, command=add_filter).pack(side='right')

        # Sequence Filter & Auto Check
        tk.Checkbutton(action_outer, text="✅ 재질 자동 판정", variable=self.auto_verify, background="#ffffff", font=("Malgun Gothic", 8)).pack(anchor='w')
        seq_row = tk.Frame(action_outer, background="#ffffff")
        seq_row.pack(fill='x', pady=2)
        ttk.Label(seq_row, text="📊 특정순번:", background="#ffffff", font=("Malgun Gothic", 8)).pack(side='left')
        ttk.Entry(seq_row, textvariable=self.sequence_filter, font=("Arial", 9), exportselection=False).pack(side='left', fill='x', expand=True, padx=5)

        self._update_pmi_filter_ui()

        # 4. Action Bar
        action_bar = tk.Frame(left_pane, background="#f9fafb")
        action_bar.pack(fill='x', pady=5)
        ttk.Button(action_bar, text="추출", style="Action.TButton", command=self.extract_only).pack(side='left', fill='x', expand=True, padx=(0, 5))
        ttk.Button(action_bar, text=" ✨ 생성 시작 ", style="Action.TButton", command=self.run_process).pack(side='left', fill='x', expand=True)

        # [RIGHT] Data Preview Pane (Always visible)
        right_pane = tk.Frame(pw, background="#f9fafb", padx=10, pady=10)
        pw.add(right_pane)
        
        tk.Label(right_pane, text="📋 실시간 미리보기 (Live Preview)", font=("Malgun Gothic", 12, "bold"), 
                 background="#f9fafb", foreground="#4b5563").pack(pady=(0, 5), anchor='w')
        
        # Create Preview inside Right Pane
        self._create_preview_ui(right_pane)
'''

# 2. Update _create_preview_ui to remove redundant padding/header if necessary.
# Actually, it's fine.

pattern = re.compile(r"def _setup_pmi_ui\(self, parent\):.*?def _update_pmi_filter_ui\(self\):", re.DOTALL)
replacement = new_pmi_ui + "\n    def _update_pmi_filter_ui(self):"

new_content = pattern.sub(replacement, content)

if new_content == content:
    print("Warning: Pattern matching failed.")
else:
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(new_content)
    print("PMI Splitscreen UI applied successfully.")
