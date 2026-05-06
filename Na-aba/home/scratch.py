with open("src/Material-Master-Manager-V13.py", "r", encoding="utf-8") as f:
    content = f.read()

# 1. Monthly site
monthly_target = """        btn_site_mgr_monthly = tk.Button(filter_frame, text="?", font=('Arial', 7), bd=0, bg=self.theme_bg, fg='gray',
                                       command=lambda: self.open_list_management_dialog('sites', target_cb=self.cb_filter_site_monthly))
        btn_site_mgr_monthly.place(in_=self.cb_filter_site_monthly, relx=1.0, x=-18, rely=0.5, anchor='e', width=16, height=16)"""
monthly_repl = """        btn_site_mgr_monthly = tk.Button(filter_frame, text="⚙️ 관리", font=('Malgun Gothic', 8), bd=0, bg=self.theme_bg, fg='blue', cursor='hand2',
                                       command=lambda: self.open_list_management_dialog('sites', target_cb=self.cb_filter_site_monthly))
        btn_site_mgr_monthly.pack(side='left', padx=(2, 10))"""
content = content.replace(monthly_target, monthly_repl)

# 2. Daily filter site
daily_filter_target = """        self.cb_daily_filter_site.pack(side='left', padx=2)
        self.cb_daily_filter_site.set('전체')"""
daily_filter_repl = """        self.cb_daily_filter_site.pack(side='left', padx=2)
        self.cb_daily_filter_site.set('전체')
        
        btn_site_mgr_daily_filter = tk.Button(filter_panel, text="⚙️ 관리", font=('Malgun Gothic', 8), bd=0, bg=self.theme_bg, fg='blue', cursor='hand2',
                                       command=lambda: self.open_list_management_dialog('sites', target_cb=self.cb_daily_filter_site))
        btn_site_mgr_daily_filter.pack(side='left', padx=(2, 10))"""
content = content.replace(daily_filter_target, daily_filter_repl)

with open("src/Material-Master-Manager-V13.py", "w", encoding="utf-8") as f:
    f.write(content)
print("Done")
