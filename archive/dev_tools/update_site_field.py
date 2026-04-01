"""Script to add Site field to monthly usage tab"""

with open(r'c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\MaterialManager.py', 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Update UI - Add Site entry field (move Usage amount to row 2)
content = content.replace(
    '''        # Usage amount
        ttk.Label(entry_frame, text="사용량:").grid(row=1, column=2, padx=5, pady=5, sticky='w')
        self.ent_usage_amount = ttk.Entry(entry_frame, width=15)
        self.ent_usage_amount.grid(row=1, column=3, padx=5, pady=5)''',
    '''        # Site selection
        ttk.Label(entry_frame, text="현장:").grid(row=1, column=2, padx=5, pady=5, sticky='w')
        self.ent_usage_site = ttk.Entry(entry_frame, width=15)
        self.ent_usage_site.grid(row=1, column=3, padx=5, pady=5)
        
        # Usage amount
        ttk.Label(entry_frame, text="사용량:").grid(row=2, column=0, padx=5, pady=5, sticky='w')
        self.ent_usage_amount = ttk.Entry(entry_frame, width=15)
        self.ent_usage_amount.grid(row=2, column=1, padx=5, pady=5)'''
)

# 2. Update Note field row
content = content.replace(
    '''        # Note
        ttk.Label(entry_frame, text="비고:").grid(row=2, column=0, padx=5, pady=5, sticky='w')
        self.ent_usage_note = ttk.Entry(entry_frame, width=70)
        self.ent_usage_note.grid(row=2, column=1, columnspan=3, padx=5, pady=5)''',
    '''        # Note
        ttk.Label(entry_frame, text="비고:").grid(row=2, column=2, padx=5, pady=5, sticky='w')
        self.ent_usage_note = ttk.Entry(entry_frame, width=40)
        self.ent_usage_note.grid(row=2, column=3, columnspan=1, padx=5, pady=5)'''
)

# 3. Update treeview columns
content = content.replace(
    "columns = ('연도', '월', '품명', '사용량', '비고', '입력일')",
    "columns = ('연도', '월', '현장', '품명', '사용량', '비고', '입력일')"
)

# 4. Update column widths
content = content.replace(
    'col_widths = [80, 60, 200, 100, 300, 150]',
    'col_widths = [80, 60, 120, 200, 100, 250, 150]'
)

# 5. Update add_monthly_usage_entry to get site value
content = content.replace(
    '''    def add_monthly_usage_entry(self):
        """Add a monthly usage entry"""
        mat_name = self.cb_usage_material.get()
        year = self.cb_usage_year.get()
        month = self.cb_usage_month.get()
        usage_str = self.ent_usage_amount.get()
        note = self.ent_usage_note.get()''',
    '''    def add_monthly_usage_entry(self):
        """Add a monthly usage entry"""
        mat_name = self.cb_usage_material.get()
        year = self.cb_usage_year.get()
        month = self.cb_usage_month.get()
        site = self.ent_usage_site.get()
        usage_str = self.ent_usage_amount.get()
        note = self.ent_usage_note.get()'''
)

# 6. Update new_entry dictionary
content = content.replace(
    '''        # Create new entry
        new_entry = {
            'Material ID': mat_id,
            'Year': int(year),
            'Month': int(month),
            'Usage': usage,
            'Note': note,
            'Entry Date': datetime.datetime.now()
        }''',
    '''        # Create new entry
        new_entry = {
            'Material ID': mat_id,
            'Year': int(year),
            'Month': int(month),
            'Site': site,
            'Usage': usage,
            'Note': note,
            'Entry Date': datetime.datetime.now()
        }'''
)

# 7. Clear site field
content = content.replace(
    '''        # Clear entry fields
        self.ent_usage_amount.delete(0, tk.END)
        self.ent_usage_note.delete(0, tk.END)''',
    '''        # Clear entry fields
        self.ent_usage_site.delete(0, tk.END)
        self.ent_usage_amount.delete(0, tk.END)
        self.ent_usage_note.delete(0, tk.END)'''
)

# 8. Update display to show site
content = content.replace(
    '''            self.monthly_usage_tree.insert('', tk.END, values=(
                entry['Year'],
                entry['Month'],
                mat_name,
                f"{entry['Usage']:.1f}",
                entry.get('Note', ''),
                entry_date
            ))''',
    '''            self.monthly_usage_tree.insert('', tk.END, values=(
                entry['Year'],
                entry['Month'],
                entry.get('Site', ''),
                mat_name,
                f"{entry['Usage']:.1f}",
                entry.get('Note', ''),
                entry_date
            ))'''
)

with open(r'c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\MaterialManager.py', 'w', encoding='utf-8') as f:
    f.write(content)

print('Successfully added Site field to monthly usage tab')
