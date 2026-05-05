    def setup_monthly_usage_tab(self):
        """Setup the monthly usage entry tab"""
        # Top frame for entry form
        entry_frame = ttk.LabelFrame(self.tab_monthly_usage, text="월별 집계")
        entry_frame.pack(fill='x', padx=10, pady=10)
        
        # Material selection
        ttk.Label(entry_frame, text="품명:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.cb_usage_material = ttk.Combobox(entry_frame, state="readonly", width=30)
        self.cb_usage_material.grid(row=0, column=1, padx=5, pady=5)
        
        # Year selection
        ttk.Label(entry_frame, text="연도:").grid(row=0, column=2, padx=5, pady=5, sticky='w')
        self.cb_usage_year = ttk.Combobox(entry_frame, values=[str(y) for y in range(2024, 2031)], width=15)
        self.cb_usage_year.grid(row=0, column=3, padx=5, pady=5)
        self.cb_usage_year.set(str(datetime.datetime.now().year))
        
        # Month selection
        ttk.Label(entry_frame, text="월:").grid(row=1, column=0, padx=5, pady=5, sticky='w')
        self.cb_usage_month = ttk.Combobox(entry_frame, values=[str(m) for m in range(1, 13)], width=30)
        self.cb_usage_month.grid(row=1, column=1, padx=5, pady=5)
        self.cb_usage_month.set(str(datetime.datetime.now().month))
        
        # Usage amount
        ttk.Label(entry_frame, text="사용량:").grid(row=1, column=2, padx=5, pady=5, sticky='w')
        self.ent_usage_amount = ttk.Entry(entry_frame, width=15)
        self.ent_usage_amount.grid(row=1, column=3, padx=5, pady=5)
        
        # Note
        ttk.Label(entry_frame, text="비고:").grid(row=2, column=0, padx=5, pady=5, sticky='w')
        self.ent_usage_note = ttk.Entry(entry_frame, width=70)
        self.ent_usage_note.grid(row=2, column=1, columnspan=3, padx=5, pady=5)
        
        # Save button
        btn_save_usage = ttk.Button(entry_frame, text="기록 저장", command=self.add_monthly_usage_entry)
        btn_save_usage.grid(row=3, column=0, columnspan=4, pady=10)
        
        # Update material combobox
        if not self.materials_df.empty:
            mats = self.materials_df['품명'].tolist()
            self.cb_usage_material['values'] = mats
        
        # Bottom frame for display
        display_frame = ttk.LabelFrame(self.tab_monthly_usage, text="월별 사용량 기록 조회")
        display_frame.pack(expand=True, fill='both', padx=10, pady=10)
        
        # Filter controls
        filter_frame = ttk.Frame(display_frame)
        filter_frame.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(filter_frame, text="연도:").pack(side='left', padx=5)
        self.cb_filter_year = ttk.Combobox(filter_frame, values=['전체'] + [str(y) for y in range(2024, 2031)], width=10)
        self.cb_filter_year.pack(side='left', padx=5)
        self.cb_filter_year.set('전체')
        
        ttk.Label(filter_frame, text="월:").pack(side='left', padx=5)
        self.cb_filter_month = ttk.Combobox(filter_frame, values=['전체'] + [str(m) for m in range(1, 13)], width=10)
        self.cb_filter_month.pack(side='left', padx=5)
        self.cb_filter_month.set('전체')
        
        btn_filter = ttk.Button(filter_frame, text="조회", command=self.update_monthly_usage_view)
        btn_filter.pack(side='left', padx=10)
        
        # Treeview for monthly usage records
        tree_frame = ttk.Frame(display_frame)
        tree_frame.pack(expand=True, fill='both', padx=5, pady=5)
        
        # Scrollbars
        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")
        
        # Treeview
        columns = ('연도', '월', '품명', '사용량', '비고', '입력일')
        self.monthly_usage_tree = ttk.Treeview(tree_frame, columns=columns, show='headings',
                                               yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        vsb.config(command=self.monthly_usage_tree.yview)
        hsb.config(command=self.monthly_usage_tree.xview)
        
        # Column configuration
        col_widths = [80, 60, 200, 100, 300, 150]
        for col, width in zip(columns, col_widths):
            self.monthly_usage_tree.heading(col, text=col)
            self.monthly_usage_tree.column(col, width=width)
        
        # Grid layout
        self.monthly_usage_tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        # Initial view update
        self.update_monthly_usage_view()
    
    def add_monthly_usage_entry(self):
        """Add a monthly usage entry"""
        mat_name = self.cb_usage_material.get()
        year = self.cb_usage_year.get()
        month = self.cb_usage_month.get()
        usage_str = self.ent_usage_amount.get()
        note = self.ent_usage_note.get()
        
        # Validation
        if not mat_name:
            messagebox.showwarning("입력 오류", "품명을 선택해주세요.")
            return
        
        if not year or not month:
            messagebox.showwarning("입력 오류", "연도와 월을 선택해주세요.")
            return
        
        try:
            usage = float(usage_str)
        except ValueError:
            messagebox.showwarning("입력 오류", "사용량은 숫자여야 합니다.")
            return
        
        # Get material ID
        mat_id = self.materials_df[self.materials_df['품명'] == mat_name]['Material ID'].values[0]
        
        # Create new entry
        new_entry = {
            'Material ID': mat_id,
            'Year': int(year),
            'Month': int(month),
            'Usage': usage,
            'Note': note,
            'Entry Date': datetime.datetime.now()
        }
        
        self.monthly_usage_df = pd.concat([self.monthly_usage_df, pd.DataFrame([new_entry])], ignore_index=True)
        self.save_data()
        self.update_monthly_usage_view()
        messagebox.showinfo("완료", f"{mat_name}의 {year}년 {month}월 사용량이 기록되었습니다.")
        
        # Clear entry fields
        self.ent_usage_amount.delete(0, tk.END)
        self.ent_usage_note.delete(0, tk.END)
    
    def update_monthly_usage_view(self):
        """Update the monthly usage treeview"""
        # Clear current view
        for item in self.monthly_usage_tree.get_children():
            self.monthly_usage_tree.delete(item)
        
        # Get filter values
        filter_year = self.cb_filter_year.get()
        filter_month = self.cb_filter_month.get()
        
        # Filter data
        filtered_df = self.monthly_usage_df.copy()
        
        if filter_year != '전체':
            filtered_df = filtered_df[filtered_df['Year'] == int(filter_year)]
        
        if filter_month != '전체':
            filtered_df = filtered_df[filtered_df['Month'] == int(filter_month)]
        
        # Display entries
        for _, entry in filtered_df.iterrows():
            mat_id = entry['Material ID']
            mat_row = self.materials_df[self.materials_df['Material ID'] == mat_id]
            
            if not mat_row.empty:
                mat_name = mat_row.iloc[0]['품명']
            else:
                mat_name = f"ID: {mat_id}"
            
            entry_date = entry.get('Entry Date',  '')
            if pd.notna(entry_date):
                entry_date = pd.to_datetime(entry_date).strftime('%Y-%m-%d %H:%M')
            
            self.monthly_usage_tree.insert('', tk.END, values=(
                entry['Year'],
                entry['Month'],
                mat_name,
                f"{entry['Usage']:.1f}",
                entry.get('Note', ''),
                entry_date
            ))

