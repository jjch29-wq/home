import os

path = r'Material-Master-Manager-V13.py'
with open(path, 'r', encoding='utf-8', errors='ignore') as f:
    content = f.read()

# Find the start of the function
target_start = '    def export_daily_work_report(self):'
if target_start not in content:
    print("CRITICAL: Function start not found")
    exit(1)

# Find the end of the part we want to replace (safety anchor)
# We look for the "2.5 차량번호 및 안전 점검" comment or similar
anchor = '# 2.5 차량 및 안전 점검 수집 (섹션 3)'
if anchor not in content:
    # Try another anchor if the first one is garbled
    anchor = '# RTK 불량 정보 수집 (UI 데이터 + DB 데이터 취합)'
    if anchor not in content:
        print("CRITICAL: Anchor point not found")
        exit(1)

# Split the content
parts = content.split(target_start, 1)
prefix = parts[0]
remaining = parts[1].split(anchor, 1)
suffix = remaining[1]

# Rebuild the function with the Option 1 (Summation) logic
new_func = r'''    def export_daily_work_report(self):
        """작업일보를 엑셀 템플릿에 출력합니다."""
        try:
            template_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'resources', 'Template_DailyWorkReport.xlsx')
            if not os.path.exists(template_path):
                template_path = r'c:\Users\jjch2\Desktop\보고서Project PROVIDENCE\Request\PMI\Na-aba\home\resources\Template_DailyWorkReport.xlsx'
            
            if not os.path.exists(template_path):
                messagebox.showerror("오류", "작업일보 템플릿(Template_DailyWorkReport.xlsx)을 찾을 수 없습니다.")
                return

            date_val = self.ent_daily_date.get_date()
            site = self.cb_daily_site.get().strip()
            
            data = {
                'date': date_val,
                'company': self.cb_daily_company.get().strip() or '원자력건설',
                'project_name': site,
                'standard': self.ent_daily_applied_code.get().strip() or 'KS',
                'equipment': self.cb_daily_equip.get().strip(),
                'report_no': self.ent_daily_report_no.get().strip(), 
                'inspection_item': self.ent_daily_inspection_item.get().strip(), 
                'inspector': '', 
                'car_no': '', 
                'methods': {},
                'rtk': {},
                'ot_status': [],
                'materials': {}
            }

            if hasattr(self, 'ndt_company_entries') and self.ndt_company_entries:
                company = self.ndt_company_entries[0].get('_company', tk.Variable()).get().strip()
                if company: data['company'] = company

            # DB 데이터 집합 (Option 1: 전체 합산 버전)
            site_records = pd.DataFrame()
            if not self.daily_usage_df.empty:
                try:
                    df_copy = self.daily_usage_df.copy()
                    df_copy['Date'] = pd.to_datetime(df_copy['Date']).dt.date
                    check_date = pd.to_datetime(date_val).date()
                    site_records = df_copy[(df_copy['Date'] == check_date) & (df_copy['Site'] == site)]
                except: pass

            all_vehicles = []
            if hasattr(self, 'vehicle_boxes'):
                for box in self.vehicle_boxes:
                    v = box.cb_vehicle_info.get().strip()
                    if v and v not in all_vehicles: all_vehicles.append(v)
            if not site_records.empty and '차량번호' in site_records.columns:
                for v in site_records['차량번호'].dropna().unique():
                    v_str = str(v).strip()
                    if v_str and v_str not in all_vehicles: all_vehicles.append(v_str)
            data['car_no'] = ", ".join(all_vehicles)

            method = self.cb_daily_test_method.get().strip()
            unit_val = self.cb_daily_unit.get().strip() 
            
            if not site_records.empty:
                db_qty = pd.to_numeric(site_records['Usage'], errors='coerce').fillna(0).sum()
                db_price = pd.to_numeric(site_records['단가'], errors='coerce').fillna(0).max()
                db_travel = pd.to_numeric(site_records['출장비'], errors='coerce').fillna(0).sum()
                db_total = pd.to_numeric(site_records['검사비'], errors='coerce').fillna(0).sum()
                qty_val = str(int(db_qty))
                price_val = str(int(db_price))
                travel_val = str(int(db_travel))
                total_val = str(int(db_total))
            else:
                qty_val = self.ent_daily_test_amount.get().strip()
                price_val = self.ent_daily_unit_price.get().strip()
                travel_val = self.ent_daily_travel_cost.get().strip()
                total_val = self.ent_daily_test_fee.get().strip()

            if method:
                data['methods'][method] = {
                    'unit': unit_val,
                    'qty': float(qty_val.replace(',', '')) if qty_val else 0,
                    'price': float(price_val.replace(',', '')) if price_val else 0,
                    'travel': float(travel_val.replace(',', '')) if travel_val else 0,
                    'total': float(total_val.replace(',', '')) if total_val else 0
                }

            # 작업자 및 O/T (간소화 버전)
            inspectors = []
            if not site_records.empty:
                for _, row in site_records.iterrows():
                    for i in range(1, 11):
                        u_key = 'User' if i == 1 else f'User{i}'
                        name = str(row.get(u_key, '')).strip()
                        if name and name != 'nan' and name not in inspectors:
                            inspectors.append(name)
            data['inspector'] = ", ".join(inspectors)

            ''' + anchor

# Assemble everything
final_content = prefix + new_func + suffix

with open(path, 'w', encoding='utf-8', errors='ignore') as f:
    f.write(final_content)
print("SUCCESS: Option 1 (Summation) logic restored and simplified.")
