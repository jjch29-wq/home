import re

file_path = 'src/Material-Master-Manager-V13.py'
with open(file_path, 'r', encoding='utf-8') as f:
    full_content = f.read()

pattern = r'# Inner content for the basic form.*?# Category definitions'

# Complete form logic with management buttons restored
replacement = """# Inner content for the basic form
        form_content = ttk.Frame(self.master_form_panel, padding=10)
        form_content.grid(row=0, column=0, sticky='w')
        
        for c in range(4): form_content.columnconfigure(c, weight=0)

        # Row 0: 업체명, 현장명
        ttk.Label(form_content, text="업체명:").grid(row=0, column=0, padx=(5, 0), pady=1, sticky='e')
        self.cb_daily_company = ttk.Combobox(form_content, width=15, values=self.companies)
        self.cb_daily_company.grid(row=0, column=1, padx=(2, 10), pady=1, sticky='w')
        
        btn_company_mgr = tk.Button(form_content, text="⚙", font=('Arial', 7), bd=0, bg=self.theme_bg, fg='gray',
                                   command=lambda: self.open_list_management_dialog('companies'))
        btn_company_mgr.place(in_=self.cb_daily_company, relx=1.0, x=-18, rely=0.5, anchor='e', width=16, height=16)

        ttk.Label(form_content, text="현장명:").grid(row=0, column=2, padx=(5, 0), pady=1, sticky='e')
        self.cb_daily_site = ttk.Combobox(form_content, width=15, values=self.sites)
        self.cb_daily_site.grid(row=0, column=3, padx=(2, 5), pady=1, sticky='w')
        
        btn_site_mgr = tk.Button(form_content, text="⚙", font=('Arial', 7), bd=0, bg=self.theme_bg, fg='gray',
                                command=lambda: self.open_list_management_dialog('sites'))
        btn_site_mgr.place(in_=self.cb_daily_site, relx=1.0, x=-18, rely=0.5, anchor='e', width=16, height=16)

        # Row 1: 날짜, 장비명
        ttk.Label(form_content, text="날짜:").grid(row=1, column=0, padx=(5, 0), pady=1, sticky='e')
        from tkcalendar import DateEntry
        self.ent_daily_date = DateEntry(form_content, width=15, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd', locale='ko_KR', state='readonly')
        self.ent_daily_date.grid(row=1, column=1, padx=(2, 10), pady=1, sticky='w')

        ttk.Label(form_content, text="장비명:").grid(row=1, column=2, padx=(5, 0), pady=1, sticky='e')
        self.cb_daily_equip = ttk.Combobox(form_content, width=15, values=self.equipments)
        self.cb_daily_equip.grid(row=1, column=3, padx=(2, 5), pady=1, sticky='w')
        
        btn_equip_mgr = tk.Button(form_content, text="⚙", font=('Arial', 7), bd=0, bg=self.theme_bg, fg='gray',
                                 command=lambda: self.open_list_management_dialog('equipments'))
        btn_equip_mgr.place(in_=self.cb_daily_equip, relx=1.0, x=-18, rely=0.5, anchor='e', width=16, height=16)

        # Row 2: 품목명
        ttk.Label(form_content, text="품목명:").grid(row=2, column=0, padx=(5, 0), pady=1, sticky='e')
        self.cb_daily_material = ttk.Combobox(form_content, width=40)
        self.cb_daily_material.grid(row=2, column=1, columnspan=3, padx=(2, 5), pady=1, sticky='w')
        
        btn_material_mgr = tk.Button(form_content, text="⚙", font=('Arial', 7), bd=0, bg=self.theme_bg, fg='gray',
                                    command=lambda: self.open_list_management_dialog('materials'))
        btn_material_mgr.place(in_=self.cb_daily_material, relx=1.0, x=-18, rely=0.5, anchor='e', width=16, height=16)

        # Row 3: 방법, 검사품명
        ttk.Label(form_content, text="방법:").grid(row=3, column=0, padx=(5, 0), pady=1, sticky='e')
        self.cb_daily_test_method = ttk.Combobox(form_content, width=15, values=['RT', 'PAUT', 'UT', 'MT', 'PT', 'ETC'])
        self.cb_daily_test_method.grid(row=3, column=1, padx=(2, 10), pady=1, sticky='w')
        
        ttk.Label(form_content, text="검사품명:").grid(row=3, column=2, padx=(5, 0), pady=1, sticky='e')
        self.ent_daily_inspection_item = ttk.Entry(form_content, width=18)
        self.ent_daily_inspection_item.grid(row=3, column=3, padx=(2, 5), pady=1, sticky='w')

        # Row 4: 수량, 단위
        ttk.Label(form_content, text="수량:").grid(row=4, column=0, padx=(5, 0), pady=1, sticky='e')
        self.ent_daily_test_amount = ttk.Entry(form_content, width=15)
        self.ent_daily_test_amount.grid(row=4, column=1, padx=(2, 10), pady=1, sticky='w')
        
        ttk.Label(form_content, text="단위:").grid(row=4, column=2, padx=(5, 0), pady=1, sticky='e')
        self.cb_daily_unit = ttk.Combobox(form_content, width=13, values=['매', 'P,M,I/D', 'M,I/D', 'Point', 'Meter', 'Inch', 'Dia'])
        self.cb_daily_unit.grid(row=4, column=3, padx=(2, 5), pady=1, sticky='w')

        # Row 5: 단가, 출장비
        ttk.Label(form_content, text="단가:").grid(row=5, column=0, padx=(5, 0), pady=1, sticky='e')
        self.ent_daily_unit_price = ttk.Entry(form_content, width=15)
        self.ent_daily_unit_price.grid(row=5, column=1, padx=(2, 10), pady=1, sticky='w')

        ttk.Label(form_content, text="출장비:").grid(row=5, column=2, padx=(5, 0), pady=1, sticky='e')
        self.ent_daily_travel_cost = ttk.Entry(form_content, width=15)
        self.ent_daily_travel_cost.grid(row=5, column=3, padx=(2, 5), pady=1, sticky='w')

        # Row 6: 적용코드, 성적서번호
        ttk.Label(form_content, text="적용코드:").grid(row=6, column=0, padx=(5, 0), pady=1, sticky='e')
        self.ent_daily_applied_code = ttk.Entry(form_content, width=15)
        self.ent_daily_applied_code.grid(row=6, column=1, padx=(2, 10), pady=1, sticky='w')

        ttk.Label(form_content, text="성적서번호:").grid(row=6, column=2, padx=(5, 0), pady=1, sticky='e')
        self.ent_daily_report_no = ttk.Entry(form_content, width=18)
        self.ent_daily_report_no.grid(row=6, column=3, padx=(2, 5), pady=1, sticky='w')

        # Row 7: 비고, 일식
        ttk.Label(form_content, text="비고:").grid(row=7, column=0, padx=(5, 0), pady=1, sticky='e')
        self.ent_daily_note = ttk.Entry(form_content, width=15)
        self.ent_daily_note.grid(row=7, column=1, padx=(2, 10), pady=1, sticky='w')

        ttk.Label(form_content, text="일식:").grid(row=7, column=2, padx=(5, 0), pady=1, sticky='e')
        self.ent_daily_meal_cost = ttk.Entry(form_content, width=15)
        self.ent_daily_meal_cost.grid(row=7, column=3, padx=(2, 5), pady=1, sticky='w')

        # Row 8: 검사비
        ttk.Label(form_content, text="검사비:").grid(row=8, column=0, padx=(5, 0), pady=1, sticky='e')
        self.ent_daily_test_fee = ttk.Entry(form_content, width=15)
        self.ent_daily_test_fee.grid(row=8, column=1, padx=(2, 10), pady=1, sticky='w')

        # Category definitions"""

updated_content = re.sub(pattern, replacement, full_content, flags=re.DOTALL)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(updated_content)
print('Form logic restored with management buttons.')
