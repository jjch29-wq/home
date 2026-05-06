
import os

filepath = r"c:\Users\-\OneDrive\바탕 화면\home\Na-aba\home\src\Material-Master-Manager-V13.py"

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

# Target start marker (line 9780 approx)
start_marker = '        # Filter controls\n'
# Target end marker (line 9897 approx)
end_marker = '        # Treeview for daily usage records\n'

start_idx = -1
end_idx = -1

for i, line in enumerate(lines):
    if i > 9000 and start_marker in line and 'filter_frame = ttk.Frame(display_frame)' in lines[i+1]:
        start_idx = i
    if start_idx != -1 and i > start_idx and end_marker in line:
        end_idx = i
        break

if start_idx != -1 and end_idx != -1:
    new_code = [
        '        # Filter controls\n',
        '        filter_frame = ttk.Frame(display_frame)\n',
        "        filter_frame.pack(fill='x', padx=5, pady=5)\n",
        '        \n',
        '        # --- Row 1: Date Filters ---\n',
        '        date_row = ttk.Frame(filter_frame)\n',
        "        date_row.pack(fill='x', pady=2)\n",
        '        \n',
        '        ttk.Label(date_row, text="시작일:").pack(side=\'left\', padx=5)\n',
        '        self.ent_daily_start_date = DateEntry(date_row, width=12, date_pattern=\'yyyy-mm-dd\', locale=\'ko_KR\', state=\'readonly\', showweeknumbers=True)\n',
        "        self.ent_daily_start_date.pack(side='left', padx=5)\n",
        '        # Default to showing all history (starting from 2024)\n',
        '        start_date = datetime.datetime(2024, 1, 1)\n',
        '        self.ent_daily_start_date.set_date(start_date)\n',
        '        \n',
        '        ttk.Label(date_row, text="종료일:").pack(side=\'left\', padx=5)\n',
        '        self.ent_daily_end_date = DateEntry(date_row, width=12, date_pattern=\'yyyy-mm-dd\', locale=\'ko_KR\', state=\'readonly\', showweeknumbers=True)\n',
        "        self.ent_daily_end_date.pack(side='left', padx=5)\n",
        '        self.ent_daily_end_date.set_date(datetime.datetime.now())\n',
        '        \n',
        '        # --- Row 2: Search Filters ---\n',
        '        filter_row = ttk.Frame(filter_frame)\n',
        "        filter_row.pack(fill='x', pady=2)\n",
        '        \n',
        '        ttk.Label(filter_row, text="업체명:").pack(side=\'left\', padx=(5, 2))\n',
        '        self.cb_daily_filter_company = ttk.Combobox(filter_row, width=12)\n',
        "        self.cb_daily_filter_company.pack(side='left', padx=2)\n",
        "        self.cb_daily_filter_company.set('전체')\n",
        '        tk.Button(filter_row, text="⚙️", font=(\'Malgun Gothic\', 8), bd=0, bg=self.theme_bg, fg=\'blue\', cursor=\'hand2\',\n',
        "                  command=lambda: self.open_list_management_dialog('companies', target_cb=self.cb_daily_filter_company)).pack(side='left', padx=(0, 5))\n",
        '        \n',
        '        ttk.Label(filter_row, text="현장:").pack(side=\'left\', padx=(5, 2))\n',
        '        self.cb_daily_filter_site = ttk.Combobox(filter_row, width=12)\n',
        "        self.cb_daily_filter_site.pack(side='left', padx=2)\n",
        "        self.cb_daily_filter_site.set('전체')\n",
        '        tk.Button(filter_row, text="⚙️", font=(\'Malgun Gothic\', 8), bd=0, bg=self.theme_bg, fg=\'blue\', cursor=\'hand2\',\n',
        "                  command=lambda: self.open_list_management_dialog('sites', target_cb=self.cb_daily_filter_site)).pack(side='left', padx=(0, 5))\n",
        '        self._bind_combobox_word_suggest(self.cb_daily_filter_site, lambda: [\'전체\'] + sorted(list(set(self.sites))))\n',
        '        \n',
        '        ttk.Label(filter_row, text="품목명:").pack(side=\'left\', padx=(5, 2))\n',
        '        self.cb_daily_filter_material = ttk.Combobox(filter_row, width=15)\n',
        "        self.cb_daily_filter_material.pack(side='left', padx=2)\n",
        "        self.cb_daily_filter_material.set('전체')\n",
        '        tk.Button(filter_row, text="⚙️", font=(\'Malgun Gothic\', 8), bd=0, bg=self.theme_bg, fg=\'blue\', cursor=\'hand2\',\n',
        "                  command=lambda: self.open_list_management_dialog('materials', target_cb=self.cb_daily_filter_material)).pack(side='left', padx=(0, 5))\n",
        '        self._bind_combobox_word_suggest(self.cb_daily_filter_material, lambda: self._get_material_candidates(include_all=True))\n',
        '        \n',
        '        ttk.Label(filter_row, text="장비명:").pack(side=\'left\', padx=(5, 2))\n',
        '        self.cb_daily_filter_equipment = ttk.Combobox(filter_row, width=15)\n',
        "        self.cb_daily_filter_equipment.pack(side='left', padx=2)\n",
        "        self.cb_daily_filter_equipment.set('전체')\n",
        '        tk.Button(filter_row, text="⚙️", font=(\'Malgun Gothic\', 8), bd=0, bg=self.theme_bg, fg=\'blue\', cursor=\'hand2\',\n',
        "                  command=lambda: self.open_list_management_dialog('equipments', target_cb=self.cb_daily_filter_equipment)).pack(side='left', padx=(0, 5))\n",
        '        self._bind_combobox_word_suggest(self.cb_daily_filter_equipment, lambda: self._get_equipment_candidates(include_all=True))\n',
        '        \n',
        '        ttk.Label(filter_row, text="작업자:").pack(side=\'left\', padx=(5, 2))\n',
        '        self.cb_daily_filter_worker = ttk.Combobox(filter_row, width=10)\n',
        "        self.cb_daily_filter_worker.pack(side='left', padx=2)\n",
        "        self.cb_daily_filter_worker.set('전체')\n",
        '        \n',
        '        ttk.Label(filter_row, text="분류:").pack(side=\'left\', padx=(5, 2))\n',
        '        self.cb_daily_filter_shift = ttk.Combobox(filter_row, width=8, state="readonly", values=["전체", "주간", "야간", "주야간", "휴일"])\n',
        "        self.cb_daily_filter_shift.pack(side='left', padx=2)\n",
        "        self.cb_daily_filter_shift.set('전체')\n",
        '        \n',
        '        # --- Row 3: Action Buttons ---\n',
        '        btn_row = ttk.Frame(filter_frame)\n',
        "        btn_row.pack(fill='x', pady=5)\n",
        '        \n',
        '        btn_filter = ttk.Button(btn_row, text="조회", style=\'Action.TButton\', command=self.update_daily_usage_view)\n',
        "        btn_filter.pack(side='left', padx=5)\n",
        '        \n',
        '        btn_filter_reset = ttk.Button(btn_row, text="♻️ 필터 초기화", command=self.reset_daily_usage_filters)\n',
        "        btn_filter_reset.pack(side='left', padx=5)\n",
        '        \n',
        '        btn_delete = ttk.Button(btn_row, text="선택 항목 삭제", command=self.delete_daily_usage_entry)\n',
        "        btn_delete.pack(side='left', padx=10)\n",
        '        \n',
        '        btn_edit = ttk.Button(btn_row, text="선택 항목 수정", command=self.open_edit_daily_usage_dialog)\n',
        "        btn_edit.pack(side='left', padx=5)\n",
        '        \n',
        '        btn_export = ttk.Button(btn_row, text="엑셀 내보내기", command=self.export_daily_usage_history)\n',
        "        btn_export.pack(side='left', padx=5)\n",
        '        \n',
        '        btn_export_all = ttk.Button(btn_row, text="전체 기록 내보내기", command=self.export_all_daily_usage)\n',
        "        btn_export_all.pack(side='left', padx=5)\n",
        '        \n',
        '        btn_col_manage = ttk.Button(btn_row, text="컬럼 관리", command=self.show_column_visibility_dialog)\n',
        "        btn_col_manage.pack(side='left', padx=10)\n",
        '\n',
        '        # Dedicated Save Button for the List View\n',
        '        self.btn_daily_save_list = ttk.Button(btn_row, text="💾 변경사항 저장", command=self.save_all_daily_usage_changes, style=\'Accent.TButton\' if \'Accent.TButton\' in self.style.theme_names() else \'TButton\')\n',
        "        self.btn_daily_save_list.pack(side='left', padx=10)\n",
        '\n',
        '        # Bindings\n',
        '        filter_widgets = [\n',
        '            self.cb_daily_filter_site, self.cb_daily_filter_company, self.cb_daily_filter_material, \n',
        '            self.cb_daily_filter_equipment, self.cb_daily_filter_worker, \n',
        '            self.cb_daily_filter_shift\n',
        '        ]\n',
        '        for widget in filter_widgets:\n',
        '            widget.bind("<Return>", lambda e: self.update_daily_usage_view())\n',
        '            widget.bind("<<ComboboxSelected>>", lambda e: self.update_daily_usage_view())\n',
        '\n',
        '        for date_widget in [self.ent_daily_start_date, self.ent_daily_end_date]:\n',
        '            date_widget.bind("<<DateEntrySelected>>", lambda e: self.update_daily_usage_view())\n',
        '            try:\n',
        '                date_widget.bind("<Return>", lambda e: self.update_daily_usage_view())\n',
        '                for child in date_widget.winfo_children():\n',
        '                    if isinstance(child, (tk.Entry, ttk.Entry)):\n',
        '                        child.bind("<Return>", lambda e: self.update_daily_usage_view())\n',
        '            except: pass\n',
        '        \n'
    ]
    
    final_lines = lines[:start_idx] + new_code + lines[end_idx:]
    
    with open(filepath, 'w', encoding='utf-8') as f:
        f.writelines(final_lines)
    print("Successfully refactored.")
else:
    print(f"Could not find markers. start_idx={start_idx}, end_idx={end_idx}")
