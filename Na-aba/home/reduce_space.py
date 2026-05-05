import os
import re

file_path = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\src\main.py"
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Modify the function signature and add "if mode is None or mode == '...':" checks
orig_func = """    def _create_row_settings(self, parent):
        # PMI Rows
        pmi_frame = ttk.LabelFrame(parent, text=" PMI 행 설정 ", padding=10)
        pmi_frame.pack(fill='x', pady=5)
        
        pmi_rows = [
            ("데이터 시작 행", "START_ROW", "PMI 데이터 시작"), 
            ("데이터 종료 행", "DATA_END_ROW", "PMI 데이터 종료"), 
            ("인쇄 영역 종료 행", "PRINT_END_ROW", "PMI 인쇄 끝")
        ]
        for i, (label, key, tip) in enumerate(pmi_rows):
            ttk.Label(pmi_frame, text=label + ":").grid(row=i, column=0, sticky='e', padx=5, pady=5)
            v = tk.StringVar(value=str(self.config.get(key, 0)))
            ttk.Entry(pmi_frame, textvariable=v, width=10).grid(row=i, column=1, sticky='w', padx=5)
            self.setting_vars[key] = v
            ttk.Label(pmi_frame, text=f"({tip})", foreground="gray", font=("Malgun Gothic", 9)).grid(row=i, column=2, sticky='w')

        # PAUT Rows
        ut_row_frame = ttk.LabelFrame(parent, text=" PAUT 행 설정 ", padding=10)
        ut_row_frame.pack(fill='x', pady=5)
        
        ut_rows = [
            ("PAUT 데이터 시작 행", "UT_START_ROW", "PAUT 데이터 시작"), 
            ("PAUT 데이터 종료 행", "UT_DATA_END_ROW", "PAUT 데이터 종료"), 
            ("PAUT 인쇄 영역 종료 행", "UT_PRINT_END_ROW", "PAUT 인쇄 끝")
        ]
        for i, (label, key, tip) in enumerate(ut_rows):
            ttk.Label(ut_row_frame, text=label + ":").grid(row=i, column=0, sticky='e', padx=5, pady=5)
            v = tk.StringVar(value=str(self.config.get(key, 0)))
            ttk.Entry(ut_row_frame, textvariable=v, width=10).grid(row=i, column=1, sticky='w', padx=5)
            self.setting_vars[key] = v
            ttk.Label(ut_row_frame, text=f"({tip})", foreground="gray", font=("Malgun Gothic", 9)).grid(row=i, column=2, sticky='w')

        # RT Rows
        rt_row_frame = ttk.LabelFrame(parent, text=" RT 행 설정 ", padding=10)
        rt_row_frame.pack(fill='x', pady=5)
        
        rt_rows = [
            ("RT 데이터 시작 행", "RT_START_ROW", "RT 데이터 시작"), 
            ("RT 데이터 종료 행", "RT_DATA_END_ROW", "RT 데이터 종료"), 
            ("RT 인쇄 영역 종료 행", "RT_PRINT_END_ROW", "RT 인쇄 끝")
        ]
        for i, (label, key, tip) in enumerate(rt_rows):
            ttk.Label(rt_row_frame, text=label + ":").grid(row=i, column=0, sticky='e', padx=5, pady=5)
            v = tk.StringVar(value=str(self.config.get(key, 0)))
            ttk.Entry(rt_row_frame, textvariable=v, width=10).grid(row=i, column=1, sticky='w', padx=5)
            self.setting_vars[key] = v
            ttk.Label(rt_row_frame, text=f"({tip})", foreground="gray", font=("Malgun Gothic", 9)).grid(row=i, column=2, sticky='w')
            
        # PT Rows
        pt_row_frame = ttk.LabelFrame(parent, text=" PT 행 설정 ", padding=10)
        pt_row_frame.pack(fill='x', pady=5)
        
        pt_rows = [
            ("PT 데이터 시작 행", "PT_START_ROW", "PT 데이터 시작"), 
            ("PT 데이터 종료 행", "PT_DATA_END_ROW", "PT 데이터 종료"), 
            ("PT 인쇄 영역 종료 행", "PT_PRINT_END_ROW", "PT 인쇄 끝")
        ]
        for i, (label, key, tip) in enumerate(pt_rows):
            ttk.Label(pt_row_frame, text=label + ":").grid(row=i, column=0, sticky='e', padx=5, pady=5)
            v = tk.StringVar(value=str(self.config.get(key, 18)))
            ttk.Entry(pt_row_frame, textvariable=v, width=10).grid(row=i, column=1, sticky='w', padx=5)
            self.setting_vars[key] = v
            ttk.Label(pt_row_frame, text=f"({tip})", foreground="gray", font=("Malgun Gothic", 9)).grid(row=i, column=2, sticky='w')"""

new_func = """    def _create_row_settings(self, parent, mode=None):
        if mode is None or mode == "PMI":
            # PMI Rows
            pmi_frame = ttk.LabelFrame(parent, text=" PMI 행 설정 ", padding=10)
            pmi_frame.pack(fill='x', pady=5)
            pmi_rows = [
                ("데이터 시작 행", "START_ROW", "PMI 데이터 시작"), 
                ("데이터 종료 행", "DATA_END_ROW", "PMI 데이터 종료"), 
                ("인쇄 영역 종료 행", "PRINT_END_ROW", "PMI 인쇄 끝")
            ]
            for i, (label, key, tip) in enumerate(pmi_rows):
                ttk.Label(pmi_frame, text=label + ":").grid(row=i, column=0, sticky='e', padx=5, pady=5)
                v = tk.StringVar(value=str(self.config.get(key, 0)))
                ttk.Entry(pmi_frame, textvariable=v, width=10).grid(row=i, column=1, sticky='w', padx=5)
                self.setting_vars[key] = v
                ttk.Label(pmi_frame, text=f"({tip})", foreground="gray", font=("Malgun Gothic", 9)).grid(row=i, column=2, sticky='w')

        if mode is None or mode == "PAUT":
            # PAUT Rows
            ut_row_frame = ttk.LabelFrame(parent, text=" PAUT 행 설정 ", padding=10)
            ut_row_frame.pack(fill='x', pady=5)
            ut_rows = [
                ("PAUT 데이터 시작 행", "UT_START_ROW", "PAUT 데이터 시작"), 
                ("PAUT 데이터 종료 행", "UT_DATA_END_ROW", "PAUT 데이터 종료"), 
                ("PAUT 인쇄 영역 종료 행", "UT_PRINT_END_ROW", "PAUT 인쇄 끝")
            ]
            for i, (label, key, tip) in enumerate(ut_rows):
                ttk.Label(ut_row_frame, text=label + ":").grid(row=i, column=0, sticky='e', padx=5, pady=5)
                v = tk.StringVar(value=str(self.config.get(key, 0)))
                ttk.Entry(ut_row_frame, textvariable=v, width=10).grid(row=i, column=1, sticky='w', padx=5)
                self.setting_vars[key] = v
                ttk.Label(ut_row_frame, text=f"({tip})", foreground="gray", font=("Malgun Gothic", 9)).grid(row=i, column=2, sticky='w')

        if mode is None or mode == "RT":
            # RT Rows
            rt_row_frame = ttk.LabelFrame(parent, text=" RT 행 설정 ", padding=10)
            rt_row_frame.pack(fill='x', pady=5)
            rt_rows = [
                ("RT 데이터 시작 행", "RT_START_ROW", "RT 데이터 시작"), 
                ("RT 데이터 종료 행", "RT_DATA_END_ROW", "RT 데이터 종료"), 
                ("RT 인쇄 영역 종료 행", "RT_PRINT_END_ROW", "RT 인쇄 끝")
            ]
            for i, (label, key, tip) in enumerate(rt_rows):
                ttk.Label(rt_row_frame, text=label + ":").grid(row=i, column=0, sticky='e', padx=5, pady=5)
                v = tk.StringVar(value=str(self.config.get(key, 0)))
                ttk.Entry(rt_row_frame, textvariable=v, width=10).grid(row=i, column=1, sticky='w', padx=5)
                self.setting_vars[key] = v
                ttk.Label(rt_row_frame, text=f"({tip})", foreground="gray", font=("Malgun Gothic", 9)).grid(row=i, column=2, sticky='w')
                
        if mode is None or mode == "PT":
            # PT Rows
            pt_row_frame = ttk.LabelFrame(parent, text=" PT 행 설정 ", padding=10)
            pt_row_frame.pack(fill='x', pady=5)
            pt_rows = [
                ("PT 데이터 시작 행", "PT_START_ROW", "PT 데이터 시작"), 
                ("PT 데이터 종료 행", "PT_DATA_END_ROW", "PT 데이터 종료"), 
                ("PT 인쇄 영역 종료 행", "PT_PRINT_END_ROW", "PT 인쇄 끝")
            ]
            for i, (label, key, tip) in enumerate(pt_rows):
                ttk.Label(pt_row_frame, text=label + ":").grid(row=i, column=0, sticky='e', padx=5, pady=5)
                v = tk.StringVar(value=str(self.config.get(key, 18)))
                ttk.Entry(pt_row_frame, textvariable=v, width=10).grid(row=i, column=1, sticky='w', padx=5)
                self.setting_vars[key] = v
                ttk.Label(pt_row_frame, text=f"({tip})", foreground="gray", font=("Malgun Gothic", 9)).grid(row=i, column=2, sticky='w')"""

if orig_func in content:
    content = content.replace(orig_func, new_func)
    
    # 2. Update calls
    content = content.replace('self._create_row_settings(tab_rows)', 'self._create_row_settings(tab_rows, mode="PMI")')
    content = content.replace('self._create_row_settings(rt_tab_rows)', 'self._create_row_settings(rt_tab_rows, mode="RT")')
    content = content.replace('self._create_row_settings(pt_tab_rows)', 'self._create_row_settings(pt_tab_rows, mode="PT")')
    content = content.replace('self._create_row_settings(paut_tab_rows)', 'self._create_row_settings(paut_tab_rows, mode="PAUT")')
    
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(content)
    print("Fixed _create_row_settings to collapse empty space efficiently.")
else:
    print("Failed to find orig func.")
