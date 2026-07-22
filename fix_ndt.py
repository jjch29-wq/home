import os
import re

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\자재작업일보기성서류Ver.1.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

old_layout = """        row0 = ttk.Frame(self.ndt_calc_frame)
        row0.pack(fill='x', pady=2)
        ttk.Label(row0, text="\uad6c\ubd84:").pack(side='left')
        ttk.Combobox(row0, textvariable=self.ndt_loc_type_var, values=["\uc218\uc1a1\ubc30\uad00(\uc8fc\ubc30\uad00)", "\ud50c\ub79c\ud2b8(\uad00\ub9ac\uc18c)"], width=18, state="readonly").pack(side='left', padx=2)

        row1 = ttk.Frame(self.ndt_calc_frame)
        row1.pack(fill='x', pady=2)
        ttk.Label(row1, text="\uc791\uc5c5\ud615\ud0dc:").pack(side='left')
        for t in ["\uc77c\ubc18", "\uc57c\uac04", "\ud734\uc77c"]:
            ttk.Radiobutton(row1, text=t, value=t, variable=self.ndt_work_time_var).pack(side='left', padx=2)
        
        ttk.Label(row1, text="  \uc870\uac741:").pack(side='left', padx=(5,0))
        self.cb_ndt_cond1 = ttk.Combobox(row1, textvariable=self.ndt_source_var, width=22, state='readonly')
        self.cb_ndt_cond1.pack(side='left', padx=2)

        row2 = ttk.Frame(self.ndt_calc_frame)
        row2.pack(fill='x', pady=2)
        ttk.Label(row2, text="\uc870\uac742:").pack(side='left', padx=(0,0))
        self.cb_ndt_cond2 = ttk.Combobox(row2, textvariable=self.ndt_thickness_var, width=22, state='readonly')
        self.cb_ndt_cond2.pack(side='left', padx=2)

        ttk.Label(row2, text="  \ubcf4\uace0\uc11c\uc6a9 \uad00\uacbd(Inch):").pack(side='left', padx=(5,0))
        self.cb_ndt_report_pipe = ttk.Combobox(row2, textvariable=self.ndt_report_pipe_var, width=10)
        self.cb_ndt_report_pipe.pack(side='left', padx=2)

        ttk.Label(row2, text="  \uc81c\uacbd\ube44\uc728(%):").pack(side='left', padx=(2,0))
        ttk.Entry(row2, textvariable=self.ndt_overhead_var, width=5).pack(side='left')
        ttk.Label(row2, text=" \uae30\uc220\ub8cc\uc728(%):").pack(side='left')
        ttk.Entry(row2, textvariable=self.ndt_tech_var, width=5).pack(side='left')
        
        row3 = ttk.Frame(self.ndt_calc_frame)
        row3.pack(fill='x', pady=2)
        ttk.Label(row3, text="[ORI] \uc870\uc778\ud2b8:").pack(side='left', padx=(0,0))
        ttk.Entry(row3, textvariable=self.ndt_ori_joint_var, width=5).pack(side='left', padx=2)
        ttk.Label(row3, text=" \ubb3c\ub7c9:").pack(side='left', padx=(0,0))
        ttk.Entry(row3, textvariable=self.ndt_ori_qty_var, width=5).pack(side='left', padx=2)
        
        ttk.Label(row3, text="  [REP] \uc870\uc778\ud2b8:").pack(side='left', padx=(5,0))
        ttk.Entry(row3, textvariable=self.ndt_rep_joint_var, width=5).pack(side='left', padx=2)
        ttk.Label(row3, text=" \ubb3c\ub7c9:").pack(side='left', padx=(0,0))
        ttk.Entry(row3, textvariable=self.ndt_rep_qty_var, width=5).pack(side='left', padx=2)
        
        ttk.Label(row3, text="  \ub2f9\uc77c \ubd88\ub7c9(REJ)\uc218:").pack(side='left', padx=(5,0))
        ttk.Entry(row3, textvariable=self.ndt_rej_joint_var, width=5).pack(side='left', padx=2)"""

new_layout = """        row0 = ttk.Frame(self.ndt_calc_frame)
        row0.pack(fill='x', pady=2)
        ttk.Label(row0, text="\uad6c\ubd84:").pack(side='left')
        ttk.Combobox(row0, textvariable=self.ndt_loc_type_var, values=["\uc218\uc1a1\ubc30\uad00(\uc8fc\ubc30\uad00)", "\ud50c\ub79c\ud2b8(\uad00\ub9ac\uc18c)"], width=18, state="readonly").pack(side='left', padx=2)
        ttk.Label(row0, text="  \uc791\uc5c5\ud615\ud0dc:").pack(side='left', padx=(5,0))
        for t in ["\uc77c\ubc18", "\uc57c\uac04", "\ud734\uc77c"]:
            ttk.Radiobutton(row0, text=t, value=t, variable=self.ndt_work_time_var).pack(side='left', padx=2)

        row1 = ttk.Frame(self.ndt_calc_frame)
        row1.pack(fill='x', pady=2)
        ttk.Label(row1, text="\uc870\uac741:").pack(side='left', padx=(0,0))
        self.cb_ndt_cond1 = ttk.Combobox(row1, textvariable=self.ndt_source_var, width=22, state='readonly')
        self.cb_ndt_cond1.pack(side='left', padx=2)
        ttk.Label(row1, text="  \uc870\uac742:").pack(side='left', padx=(5,0))
        self.cb_ndt_cond2 = ttk.Combobox(row1, textvariable=self.ndt_thickness_var, width=22, state='readonly')
        self.cb_ndt_cond2.pack(side='left', padx=2)

        row2 = ttk.Frame(self.ndt_calc_frame)
        row2.pack(fill='x', pady=2)
        ttk.Label(row2, text="\ubcf4\uace0\uc11c\uc6a9 \uad00\uacbd(Inch):").pack(side='left', padx=(0,0))
        self.cb_ndt_report_pipe = ttk.Combobox(row2, textvariable=self.ndt_report_pipe_var, width=10)
        self.cb_ndt_report_pipe.pack(side='left', padx=2)
        ttk.Label(row2, text="  \uc81c\uacbd\ube44\uc728(%):").pack(side='left', padx=(5,0))
        ttk.Entry(row2, textvariable=self.ndt_overhead_var, width=5).pack(side='left')
        ttk.Label(row2, text=" \uae30\uc220\ub8cc\uc728(%):").pack(side='left', padx=(5,0))
        ttk.Entry(row2, textvariable=self.ndt_tech_var, width=5).pack(side='left')
        
        row3 = ttk.Frame(self.ndt_calc_frame)
        row3.pack(fill='x', pady=2)
        ttk.Label(row3, text="[ORI] \uc870\uc778\ud2b8:").pack(side='left', padx=(0,0))
        ttk.Entry(row3, textvariable=self.ndt_ori_joint_var, width=5).pack(side='left', padx=2)
        ttk.Label(row3, text=" \ubb3c\ub7c9:").pack(side='left', padx=(0,0))
        ttk.Entry(row3, textvariable=self.ndt_ori_qty_var, width=5).pack(side='left', padx=2)
        ttk.Label(row3, text="  [REP] \uc870\uc778\ud2b8:").pack(side='left', padx=(5,0))
        ttk.Entry(row3, textvariable=self.ndt_rep_joint_var, width=5).pack(side='left', padx=2)
        ttk.Label(row3, text=" \ubb3c\ub7c9:").pack(side='left', padx=(0,0))
        ttk.Entry(row3, textvariable=self.ndt_rep_qty_var, width=5).pack(side='left', padx=2)
        
        row4 = ttk.Frame(self.ndt_calc_frame)
        row4.pack(fill='x', pady=2)
        ttk.Label(row4, text="\ub2f9\uc77c \ubd88\ub7c9(REJ)\uc218:").pack(side='left', padx=(0,0))
        ttk.Entry(row4, textvariable=self.ndt_rej_joint_var, width=5).pack(side='left', padx=2)"""

if old_layout in content:
    content = content.replace(old_layout, new_layout)
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(content)
    print("NDT Layout replaced.")
else:
    print("Could not find old layout.")
