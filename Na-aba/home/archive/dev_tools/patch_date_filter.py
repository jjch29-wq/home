import sys, re

def patch_file(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    # 1. Insert date listbox UI
    ui_insertion = """        # [NEW] Date Filter Listbox
        ttk.Label(sidebar, text="날짜 선택 필터", font=("Malgun Gothic", 9, "bold")).pack(pady=(5, 2))
        list_frame = tk.Frame(sidebar)
        list_frame.pack(fill='x', pady=5)
        self.date_listbox = tk.Listbox(list_frame, selectmode='multiple', height=4, exportselection=False, font=("Malgun Gothic", 9))
        self.date_listbox.pack(side='left', fill='both', expand=True)
        sb = ttk.Scrollbar(list_frame, orient='vertical', command=self.date_listbox.yview)
        sb.pack(side='right', fill='y')
        self.date_listbox.config(yscrollcommand=sb.set)
        
        def apply_date_filter():
            selected_indices = self.date_listbox.curselection()
            selected_dates = [self.date_listbox.get(i) for i in selected_indices]
            for item in self.extracted_data:
                item['selected'] = (item.get('Date', '') in selected_dates)
            self.populate_preview(self.extracted_data, switch_tab=False)
            
        ttk.Button(sidebar, text="선택 날짜 적용", command=apply_date_filter).pack(fill='x', pady=(0, 10))
        
"""
    sidebar_pattern = r'(ttk\.Label\(sidebar, text="데이터 관리"[^\n]+\n)'
    if '날짜 선택 필터' not in content:
        content = re.sub(sidebar_pattern, r'\1' + ui_insertion, content)

    # 2. Insert update_date_listbox method
    method_insertion = """
    def update_date_listbox(self):
        if not hasattr(self, 'date_listbox'): return
        unique_dates = sorted(list(set(item.get('Date', '') for item in self.extracted_data if item.get('Date'))))
        self.date_listbox.delete(0, tk.END)
        for i, d in enumerate(unique_dates):
            self.date_listbox.insert(tk.END, d)
            self.date_listbox.selection_set(i) # Select all by default
"""
    if 'def update_date_listbox' not in content:
        content = content.replace('    def copy_cell(self):', method_insertion + '\n    def copy_cell(self):')
        
    # 3. Call update_date_listbox in extract_only
    if 'self.update_date_listbox()' not in content:
        content = re.sub(r'(self\.populate_preview\(self\.extracted_data.*?\n)', r'\1        self.update_date_listbox()\n', content)

    # 4. Same for clear_all
    content = re.sub(r'(self\.extracted_data = \[\]\n\s*self\.populate_preview\(self\.extracted_data, switch_tab=False\)\n)', r'\1        self.update_date_listbox()\n', content)

    # 5. Same for delete_item
    content = re.sub(r'(self\.extracted_data\.pop\(idx\)\n\s*self\.populate_preview\(self\.extracted_data, switch_tab=False\)\n)', r'\1        self.update_date_listbox()\n', content)

    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)

patch_file(r'c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\JJCHSITPMI-V2-Unified.py')
patch_file(r'c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\JJCHSITPMI-V3.py')
print("Patch applied to both files successfully.")
