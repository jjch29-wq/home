import re

with open('src/Material-Master-Manager-V13.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Define the full master form content correctly
new_form_logic = """        # Inner content for the basic form
        form_content = ttk.Frame(self.master_form_panel, padding=10)
        form_content.grid(row=0, column=0, sticky='nsew')
        
        # Row 0: Company Selection & Site Selection
        ttk.Label(form_content, text="업체명:").grid(row=0, column=0, padx=(5, 0), pady=1, sticky='e')
        self.cb_daily_company = ttk.Combobox(form_content, width=12, values=self.companies)
        self.cb_daily_company.grid(row=0, column=1, padx=(0, 5), pady=1, sticky='w')
        
        btn_company_mgr = tk.Button(form_content, text="⚙", font=('Arial', 7), bd=0, bg=self.theme_bg, fg='gray',
                                   command=lambda: self.open_list_management_dialog('companies'))
        btn_company_mgr.place(in_=self.cb_daily_company, relx=1.0, x=-18, rely=0.5, anchor='e', width=16, height=16)

        def on_company_select(e):
            self.root.after(10, self.cb_daily_site.focus_set)
            return "break"
        self.cb_daily_company.bind('<<ComboboxSelected>>', on_company_select)
        self.cb_daily_company.bind('<Return>', on_company_select)

        ttk.Label(form_content, text="현장명:").grid(row=0, column=2, padx=(5, 0), pady=1, sticky='e')
        self.cb_daily_site = ttk.Combobox(form_content, width=12, values=self.sites)
        self.cb_daily_site.grid(row=0, column=3, padx=(0, 5), pady=1, sticky='w')

        btn_site_mgr = tk.Button(form_content, text="⚙", font=('Arial', 7), bd=0, bg=self.theme_bg, fg='gray',
                                command=lambda: self.open_list_management_dialog('sites'))
        btn_site_mgr.place(in_=self.cb_daily_site, relx=1.0, x=-18, rely=0.5, anchor='e', width=16, height=16)

        def on_site_select(e):
            self.root.after(10, self.ent_daily_date.focus_set)
            return "break"
        self.cb_daily_site.bind('<<ComboboxSelected>>', on_site_select)
        self.cb_daily_site.bind('<Return>', on_site_select)

        # Row 1: Date & Equipment Selection
        ttk.Label(form_content, text="날짜:").grid(row=1, column=0, padx=(5, 0), pady=1, sticky='e')
        self.ent_daily_date = DateEntry(form_content, width=12, background='darkblue',
                                         foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd',
                                         locale='ko_KR', state='readonly', showweeknumbers=True)
        self.ent_daily_date.grid(row=1, column=1, padx=(0, 5), pady=1, sticky='w')

        def on_date_select_combined(e=None):
            self.root.after(100, self.cb_daily_equip.focus_set)
            return "break"
        self.ent_daily_date.bind('<<DateEntrySelected>>', on_date_select_combined)
        self.ent_daily_date.bind('<Return>', on_date_select_combined)

        ttk.Label(form_content, text="장비명:").grid(row=1, column=2, padx=(5, 0), pady=1, sticky='e')
        self.cb_daily_equip = ttk.Combobox(form_content, width=12, values=self.equipments)
        self.cb_daily_equip.grid(row=1, column=3, padx=(0, 5), pady=1, sticky='w')
        
        btn_equip_mgr = tk.Button(form_content, text="⚙", font=('Arial', 7), bd=0, bg=self.theme_bg, fg='gray',
                                 command=lambda: self.open_list_management_dialog('equipments'))
        btn_equip_mgr.place(in_=self.cb_daily_equip, relx=1.0, x=-18, rely=0.5, anchor='e', width=16, height=16)

        self._bind_combobox_word_suggest(self.cb_daily_equip, lambda: self._get_equipment_candidates(include_all=False))

        def on_equip_select(e):
            self.root.after(10, self.cb_daily_material.focus_set)
            return "break"
        self.cb_daily_equip.bind('<<ComboboxSelected>>', on_equip_select)
        self.cb_daily_equip.bind('<Return>', on_equip_select)

        # Row 2: Material Selection (Full width)
        ttk.Label(form_content, text="품목명:").grid(row=2, column=0, padx=(5, 0), pady=1, sticky='e')
        self.cb_daily_material = ttk.Combobox(form_content, width=45)
        self.cb_daily_material.grid(row=2, column=1, columnspan=3, padx=(0, 5), pady=1, sticky='ew')
        
        btn_material_mgr = tk.Button(form_content, text="⚙", font=('Arial', 7), bd=0, bg=self.theme_bg, fg='gray',
                                    command=lambda: self.open_list_management_dialog('materials'))
        btn_material_mgr.place(in_=self.cb_daily_material, relx=1.0, x=-18, rely=0.5, anchor='e', width=16, height=16)
        
        self._bind_combobox_word_suggest(self.cb_daily_material, lambda: self._get_equipment_candidates(include_all=False))

        def on_mat_select(e):
            self.root.after(10, self.cb_daily_test_method.focus_set)
            return "break"
        self.cb_daily_material.bind('<<ComboboxSelected>>', on_mat_select)
        self.cb_daily_material.bind('<Return>', on_mat_select)

        # Row 3: Method, Item Name & Unit Price
        ttk.Label(form_content, text="방법:").grid(row=3, column=0, padx=(5, 0), pady=1, sticky='e')
        self.cb_daily_test_method = ttk.Combobox(form_content, width=12, values=['RT', 'PAUT', 'UT', 'MT', 'PT', 'ETC'])
        self.cb_daily_test_method.grid(row=3, column=1, padx=(0, 5), pady=1, sticky='w')

        def on_method_select(e):
            self.root.after(10, self.ent_daily_inspection_item.focus_set)
            return "break"
        self.cb_daily_test_method.bind('<<ComboboxSelected>>', on_method_select, add='+')
        self.cb_daily_test_method.bind('<Return>', on_method_select, add='+')
        
        def on_method_change_auto_unit(e):
            method = self.cb_daily_test_method.get().strip()
            unit_map = {'RT': '매', 'UT': 'P,M,I/D', 'MT': 'P,M,I/D', 'PT': 'P,M,I/D', 'PAUT': 'M,I/D'}
            if method in unit_map: self.cb_daily_unit.set(unit_map[method])
            return "break"
        self.cb_daily_test_method.bind('<<ComboboxSelected>>', on_method_change_auto_unit, add='+')
        
        ttk.Label(form_content, text="검사품명:").grid(row=3, column=2, padx=(5, 0), pady=1, sticky='e')
        self.ent_daily_inspection_item = ttk.Entry(form_content, width=15)
        self.ent_daily_inspection_item.insert(0, "Piping")
        self.ent_daily_inspection_item.grid(row=3, column=3, padx=(0, 5), pady=1, sticky='w')
        self.ent_daily_inspection_item.bind('<Return>', lambda e: self.ent_daily_test_amount.focus_set())

        ttk.Label(form_content, text="단가:").grid(row=3, column=4, padx=(5, 0), pady=1, sticky='e')
        self.ent_daily_unit_price = ttk.Entry(form_content, width=12)
        self.ent_daily_unit_price.grid(row=3, column=5, padx=(0, 5), pady=1, sticky='w')

        # Row 4: Quantity, Unit & Travel Cost
        ttk.Label(form_content, text="수량:").grid(row=4, column=0, padx=(5, 0), pady=1, sticky='e')
        self.ent_daily_test_amount = ttk.Entry(form_content, width=12)
        self.ent_daily_test_amount.grid(row=4, column=1, padx=(0, 5), pady=1, sticky='w')
        self.ent_daily_test_amount.bind('<Return>', lambda e: self.cb_daily_unit.focus_set())
        
        ttk.Label(form_content, text="단위:").grid(row=4, column=2, padx=(5, 0), pady=1, sticky='e')
        self.cb_daily_unit = ttk.Combobox(form_content, width=10, values=['매', 'P,M,I/D', 'M,I/D', 'Point', 'Meter', 'Inch', 'Dia'])
        self.cb_daily_unit.set('매')
        self.cb_daily_unit.grid(row=4, column=3, padx=(0, 5), pady=1, sticky='w')
        self.cb_daily_unit.bind('<Return>', lambda e: self.ent_daily_unit_price.focus_set())
        self.cb_daily_unit.bind('<<ComboboxSelected>>', lambda e: self.ent_daily_unit_price.focus_set())

        ttk.Label(form_content, text="출장비:").grid(row=4, column=4, padx=(5, 0), pady=1, sticky='e')
        self.ent_daily_travel_cost = ttk.Entry(form_content, width=12)
        self.ent_daily_travel_cost.grid(row=4, column=5, padx=(0, 5), pady=1, sticky='w')

        # Row 5: Applied Code, Report No & Note
        ttk.Label(form_content, text="적용코드:").grid(row=5, column=0, padx=(5, 0), pady=1, sticky='e')
        self.ent_daily_applied_code = ttk.Entry(form_content, width=12)
        self.ent_daily_applied_code.insert(0, "KS")
        self.ent_daily_applied_code.grid(row=5, column=1, padx=(0, 5), pady=1, sticky='w')

        ttk.Label(form_content, text="성적서번호:").grid(row=5, column=2, padx=(5, 0), pady=1, sticky='e')
        self.ent_daily_report_no = ttk.Entry(form_content, width=15)
        self.ent_daily_report_no.grid(row=5, column=3, padx=(0, 5), pady=1, sticky='w')

        ttk.Label(form_content, text="비고:").grid(row=5, column=4, padx=(5, 0), pady=1, sticky='e')
        self.ent_daily_note = ttk.Entry(form_content, width=12)
        self.ent_daily_note.grid(row=5, column=5, padx=(0, 5), pady=1, sticky='w')

        # Row 6: Meal Cost & Test Fee
        ttk.Label(form_content, text="일식:").grid(row=6, column=0, padx=(5, 0), pady=1, sticky='e')
        self.ent_daily_meal_cost = ttk.Entry(form_content, width=12)
        self.ent_daily_meal_cost.insert(0, "0")
        self.ent_daily_meal_cost.grid(row=6, column=1, padx=(0, 5), pady=1, sticky='w')

        ttk.Label(form_content, text="검사비:").grid(row=6, column=2, padx=(5, 0), pady=1, sticky='e')
        self.ent_daily_test_fee = ttk.Entry(form_content, width=15)
        self.ent_daily_test_fee.grid(row=6, column=3, padx=(0, 5), pady=1, sticky='w')
"""

# Replace the entire block from 'Inner content for the basic form' to 'Category definitions'
pattern = r'# Inner content for the basic form.*?# Category definitions'
new_text = new_form_logic + "\n\n        # Category definitions"
updated_content = re.sub(pattern, new_text, content, flags=re.DOTALL)

with open('src/Material-Master-Manager-V13.py', 'w', encoding='utf-8') as f:
    f.write(updated_content)
print('Form logic cleaned and updated with tight spacing.')
