import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import datetime
import os

class MaterialManager:
    def __init__(self, root):
        self.root = root
        self.root.title("자재 및 소모품 관리 시스템 (Material Manager)")
        self.root.geometry("1400x800")
        
        self.db_path = r'c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Material_Inventory.xlsx'
        self.load_data()
        
        self.create_widgets()

    def load_data(self):
        try:
            if not os.path.exists(self.db_path):
                # Initialize with new schema
                self.materials_df = pd.DataFrame(columns=[
                    'Material ID', '회사코드', '관리품번', '품명', '창고',
                    '모델명', '규격', '품목군코드', '제조사', '제조국', 
                    '가격', '관리단위', '수량'
                ])
                self.transactions_df = pd.DataFrame(columns=['Date', 'Material ID', 'Type', 'Quantity', 'Note', 'User'])
                self.monthly_usage_df = pd.DataFrame(columns=['Material ID', 'Year', 'Month', 'Usage', 'Note', 'Entry Date'])
            else:
                self.materials_df = pd.read_excel(self.db_path, sheet_name='Materials')
                self.transactions_df = pd.read_excel(self.db_path, sheet_name='Transactions')
                # Ensure Date column is datetime
                if not self.transactions_df.empty:
                    self.transactions_df['Date'] = pd.to_datetime(self.transactions_df['Date'])
                
                # Load monthly usage data if exists
                try:
                    self.monthly_usage_df = pd.read_excel(self.db_path, sheet_name='MonthlyUsage')
                    if not self.monthly_usage_df.empty:
                        self.monthly_usage_df['Entry Date'] = pd.to_datetime(self.monthly_usage_df['Entry Date'])
                except:
                    self.monthly_usage_df = pd.DataFrame(columns=['Material ID', 'Year', 'Month', 'Usage', 'Note', 'Entry Date'])
                
                # Migrate old schema if needed
                if 'Equipment Code' in self.materials_df.columns and '회사코드' not in self.materials_df.columns:
                    self.migrate_old_schema()
        except Exception as e:
            messagebox.showerror("Error", f"데이터를 불러오는데 실패했습니다: {e}")
            self.materials_df = pd.DataFrame(columns=[
                'Material ID', '회사코드', '관리품번', '품명', '창고',
                '모델명', '규격', '품목군코드', '제조사', '제조국', 
                '가격', '관리단위', '수량'
            ])
            self.transactions_df = pd.DataFrame(columns=['Date', 'Material ID', 'Type', 'Quantity', 'Note', 'User'])
            self.monthly_usage_df = pd.DataFrame(columns=['Material ID', 'Year', 'Month', 'Usage', 'Note', 'Entry Date'])
    
    def migrate_old_schema(self):
        """Migrate data from old schema to new schema"""
        old_df = self.materials_df.copy()
        self.materials_df = pd.DataFrame(columns=[
            'Material ID', '회사코드', '관리품번', '품명', '창고',
            '모델명', '규격', '품목군코드', '제조사', '제조국', 
            '가격', '관리단위', '수량'
        ])
        
        for _, row in old_df.iterrows():
            new_row = {
                'Material ID': row.get('Material ID', ''),
                '회사코드': '',
                '관리품번': row.get('Equipment Code', ''),
                '품명': row.get('Item Name', row.get('Name', '')),
                '창고': '',
                '모델명': '',
                '규격': row.get('Specification', ''),
                '품목군코드': '',
                '제조사': row.get('Manufacturer', ''),
                '제조국': '',
                '가격': 0,
                '관리단위': row.get('Unit', 'EA'),
                '수량': row.get('Current Stock', row.get('Initial Stock', 0))
            }
            self.materials_df = pd.concat([self.materials_df, pd.DataFrame([new_row])], ignore_index=True)
        
        self.save_data()
        messagebox.showinfo("마이그레이션 완료", "기존 데이터가 새로운 형식으로 변환되었습니다.")

    def save_data(self):
        try:
            with pd.ExcelWriter(self.db_path, engine='openpyxl') as writer:
                self.materials_df.to_excel(writer, sheet_name='Materials', index=False)
                self.transactions_df.to_excel(writer, sheet_name='Transactions', index=False)
                self.monthly_usage_df.to_excel(writer, sheet_name='MonthlyUsage', index=False)
        except Exception as e:
            messagebox.showerror("Error", f"데이터를 저장하는데 실패했습니다: {e}")

    def create_widgets(self):
        # Notebook for Tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(expand=True, fill='both', padx=10, pady=10)
        
        # Tab 1: Current Stock
        self.tab_stock = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_stock, text='현재 재고 현황')
        self.setup_stock_tab()
        
        # Tab 2: Register/Transaction
        self.tab_inout = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_inout, text='입출고 관리')
        self.setup_inout_tab()
        
        # Tab 3: Reports
        self.tab_reports = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_reports, text='월별/년별 보고서')
        self.setup_report_tab()
        
        # Tab 4: Import/Export
        self.tab_import = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_import, text='데이터 가져오기/내보내기')
        self.setup_import_tab()
        
        # Tab 5: Monthly Usage Entry
        self.tab_monthly_usage = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_monthly_usage, text='월별 사용량 기입')
        self.setup_monthly_usage_tab()

    def setup_stock_tab(self):
        # Control Frame
        control_frame = ttk.Frame(self.tab_stock)
        control_frame.pack(fill='x', padx=5, pady=5)
        
        # Refresh Button
        btn_refresh = ttk.Button(control_frame, text="재고 새로고침", command=self.update_stock_view)
        btn_refresh.pack(side='left', padx=5)
        
        # Low Stock Alert Button
        btn_alert = ttk.Button(control_frame, text="재주문 필요 항목 보기", command=self.show_low_stock)
        btn_alert.pack(side='left', padx=5)
        
        # Search Frame
        ttk.Label(control_frame, text="검색:").pack(side='left', padx=(20, 5))
        self.search_var = tk.StringVar()
        self.search_var.trace_add('write', lambda *args: self.update_stock_view())
        search_entry = ttk.Entry(control_frame, textvariable=self.search_var, width=30)
        search_entry.pack(side='left', padx=5)
        
        # Treeview for Stock with Scrollbars
        tree_frame = ttk.Frame(self.tab_stock)
        tree_frame.pack(expand=True, fill='both', padx=5, pady=5)
        
        # Scrollbars
        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")
        
        columns = ('ID', '회사코드', '관리품번', '품명', '창고', '모델명', '규격', '품목군코드', '제조사', '제조국', '가격', '관리단위', '수량')
        self.stock_tree = ttk.Treeview(tree_frame, columns=columns, show='headings', 
                                      yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        vsb.config(command=self.stock_tree.yview)
        hsb.config(command=self.stock_tree.xview)
        
        # Column configuration
        col_widths = [50, 70, 90, 150, 80, 100, 100, 80, 100, 70, 70, 70, 70]
        for col, width in zip(columns, col_widths):
            self.stock_tree.heading(col, text=col)
            self.stock_tree.column(col, width=width)
        
        # Grid layout
        self.stock_tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        self.update_stock_view()
    
    def show_low_stock(self):
        """Show items with low stock (less than 10 units)"""
        low_stock_items = []
        for _, mat in self.materials_df.iterrows():
            current = mat.get('수량', 0)
            if pd.notna(current) and current < 10:
                low_stock_items.append((mat.get('품명', ''), current))
        
        if not low_stock_items:
            messagebox.showinfo("재고 알림", "수량이 10개 미만인 항목이 없습니다.")
        else:
            msg = "다음 항목들의 재고가 부족합니다:\n\n"
            for item, current in low_stock_items:
                msg += f"• {item}: 현재 {current}\n"
            messagebox.showwarning("재고 부족", msg)
    
    def calculate_current_stock(self, mat_id):
        """Calculate current stock for a material based on transactions"""
        # Filter transactions for this material
        mat_trans = self.transactions_df[self.transactions_df['Material ID'] == mat_id]
        in_qty = mat_trans[mat_trans['Type'] == 'IN']['Quantity'].sum()
        out_qty = mat_trans[mat_trans['Type'] == 'OUT']['Quantity'].sum()
        
        # Get the current stored quantity from materials_df
        mat = self.materials_df[self.materials_df['Material ID'] == mat_id]
        if not mat.empty:
            stored_qty = mat.iloc[0].get('수량', 0)
            # If there are transactions, calculate; otherwise return stored value
            if not mat_trans.empty:
                return stored_qty + in_qty - out_qty
            return stored_qty
        return 0

    def update_stock_view(self):
        # Clear current view
        for item in self.stock_tree.get_children():
            self.stock_tree.delete(item)
        
        search_term = self.search_var.get().lower() if hasattr(self, 'search_var') else ''
        
        # Calculate current stock
        stock_summary = []
        for _, mat in self.materials_df.iterrows():
            mat_id = mat['Material ID']
            current_stock = self.calculate_current_stock(mat_id)
            
            # Update 수량 in dataframe
            self.materials_df.loc[self.materials_df['Material ID'] == mat_id, '수량'] = current_stock
            
            stock_summary.append((
                mat_id,
                mat.get('회사코드', ''),
                mat.get('관리품번', ''),
                mat.get('품명', ''),
                mat.get('창고', ''),
                mat.get('모델명', ''),
                mat.get('규격', ''),
                mat.get('품목군코드', ''),
                mat.get('제조사', ''),
                mat.get('제조국', ''),
                mat.get('가격', 0),
                mat.get('관리단위', 'EA'),
                current_stock
            ))
        
        # Filter by search term
        for row in stock_summary:
            if search_term:
                row_str = ' '.join(str(x).lower() for x in row)
                if search_term not in row_str:
                    continue
            self.stock_tree.insert('', tk.END, values=row)

    def setup_inout_tab(self):
        # Scrollable Canvas
        canvas = tk.Canvas(self.tab_inout)
        scrollbar = ttk.Scrollbar(self.tab_inout, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Frame for Registration
        reg_frame = ttk.LabelFrame(scrollable_frame, text="자재 신규 등록")
        reg_frame.pack(fill='x', padx=10, pady=5)
        
        # Row 0
        ttk.Label(reg_frame, text="설비코드:").grid(row=0, column=0, padx=5, pady=2, sticky='w')
        self.ent_eq_code = ttk.Entry(reg_frame, width=20)
        self.ent_eq_code.grid(row=0, column=1, padx=5, pady=2)
        
        ttk.Label(reg_frame, text="품목명:").grid(row=0, column=2, padx=5, pady=2, sticky='w')
        self.ent_item_name = ttk.Entry(reg_frame, width=25)
        self.ent_item_name.grid(row=0, column=3, padx=5, pady=2)
        
        # Row 1
        ttk.Label(reg_frame, text="분류:").grid(row=1, column=0, padx=5, pady=2, sticky='w')
        self.ent_class = ttk.Entry(reg_frame, width=20)
        self.ent_class.grid(row=1, column=1, padx=5, pady=2)
        
        ttk.Label(reg_frame, text="규격:").grid(row=1, column=2, padx=5, pady=2, sticky='w')
        self.ent_spec = ttk.Entry(reg_frame, width=25)
        self.ent_spec.grid(row=1, column=3, padx=5, pady=2)
        
        # Row 2
        ttk.Label(reg_frame, text="단위:").grid(row=2, column=0, padx=5, pady=2, sticky='w')
        self.ent_unit = ttk.Entry(reg_frame, width=20)
        self.ent_unit.grid(row=2, column=1, padx=5, pady=2)
        
        ttk.Label(reg_frame, text="공급업자:").grid(row=2, column=2, padx=5, pady=2, sticky='w')
        self.ent_supplier = ttk.Entry(reg_frame, width=25)
        self.ent_supplier.grid(row=2, column=3, padx=5, pady=2)
        
        # Row 3
        ttk.Label(reg_frame, text="제조조:").grid(row=3, column=0, padx=5, pady=2, sticky='w')
        self.ent_mfr = ttk.Entry(reg_frame, width=20)
        self.ent_mfr.grid(row=3, column=1, padx=5, pady=2)
        
        ttk.Label(reg_frame, text="재주문 수준:").grid(row=3, column=2, padx=5, pady=2, sticky='w')
        self.ent_reorder = ttk.Entry(reg_frame, width=25)
        self.ent_reorder.grid(row=3, column=3, padx=5, pady=2)
        self.ent_reorder.insert(0, "0")
        
        # Row 4
        ttk.Label(reg_frame, text="초기재고:").grid(row=4, column=0, padx=5, pady=2, sticky='w')
        self.ent_init = ttk.Entry(reg_frame, width=20)
        self.ent_init.grid(row=4, column=1, padx=5, pady=2)
        self.ent_init.insert(0, "0")
        
        btn_reg = ttk.Button(reg_frame, text="자재 등록", command=self.register_material)
        btn_reg.grid(row=5, column=0, columnspan=4, pady=10)
        
        # Frame for In/Out Transaction
        trans_frame = ttk.LabelFrame(scrollable_frame, text="입출고 기록")
        trans_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(trans_frame, text="자재 선택:").grid(row=0, column=0, padx=5, pady=2, sticky='w')
        self.cb_material = ttk.Combobox(trans_frame, state="readonly", width=30)
        self.cb_material.grid(row=0, column=1, padx=5, pady=2)
        self.update_material_combo()
        
        ttk.Label(trans_frame, text="구분:").grid(row=0, column=2, padx=5, pady=2, sticky='w')
        self.cb_type = ttk.Combobox(trans_frame, values=["IN", "OUT"], state="readonly", width=15)
        self.cb_type.grid(row=0, column=3, padx=5, pady=2)
        self.cb_type.set("OUT")
        
        ttk.Label(trans_frame, text="수량:").grid(row=1, column=0, padx=5, pady=2, sticky='w')
        self.ent_qty = ttk.Entry(trans_frame, width=30)
        self.ent_qty.grid(row=1, column=1, padx=5, pady=2)
        
        ttk.Label(trans_frame, text="비고:").grid(row=1, column=2, padx=5, pady=2, sticky='w')
        self.ent_note = ttk.Entry(trans_frame, width=30)
        self.ent_note.grid(row=1, column=3, padx=5, pady=2)
        
        ttk.Label(trans_frame, text="담당자:").grid(row=2, column=0, padx=5, pady=2, sticky='w')
        self.ent_user = ttk.Entry(trans_frame, width=30)
        self.ent_user.grid(row=2, column=1, padx=5, pady=2)
        
        btn_trans = ttk.Button(trans_frame, text="기록 저장", command=self.add_transaction)
        btn_trans.grid(row=3, column=0, columnspan=4, pady=10)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def update_material_combo(self):
        if not self.materials_df.empty:
            mats = self.materials_df['품명'].tolist()
            self.cb_material['values'] = mats

    def register_material(self):
        item_name = self.ent_item_name.get()
        eq_code = self.ent_eq_code.get()
        classification = self.ent_class.get()
        spec = self.ent_spec.get()
        unit = self.ent_unit.get()
        supplier = self.ent_supplier.get()
        manufacturer = self.ent_mfr.get()
        
        try:
            init_stock = float(self.ent_init.get())
            reorder_point = float(self.ent_reorder.get())
        except ValueError:
            messagebox.showwarning("입력 오류", "초기 재고와 재주문 수준은 숫자여야 합니다.")
            return
            
        if not item_name:
            messagebox.showwarning("입력 오류", "품목명을 입력해주세요.")
            return
        
        # Generate Material ID
        if self.materials_df.empty:
            mat_id = 1
        else:
            mat_id = self.materials_df['Material ID'].max() + 1
        
        new_row = {
            'Material ID': mat_id,
            '회사코드': '',
            '관리품번': eq_code,
            '품명': item_name,
            '창고': '',
            '모델명': '',
            '규격': spec,
            '품목군코드': classification,
            '제조사': manufacturer,
            '제조국': '',
            '가격': 0,
            '관리단위': unit if unit else 'EA',
            '수량': init_stock
        }
        
        self.materials_df = pd.concat([self.materials_df, pd.DataFrame([new_row])], ignore_index=True)
        self.save_data()
        self.update_material_combo()
        self.update_stock_view()
        messagebox.showinfo("완료", f"'{item_name}' 자재가 등록되었습니다.")
        
        # Clear entries
        self.ent_eq_code.delete(0, tk.END)
        self.ent_item_name.delete(0, tk.END)
        self.ent_class.delete(0, tk.END)
        self.ent_spec.delete(0, tk.END)
        self.ent_unit.delete(0, tk.END)
        self.ent_supplier.delete(0, tk.END)
        self.ent_mfr.delete(0, tk.END)
        self.ent_reorder.delete(0, tk.END)
        self.ent_reorder.insert(0, "0")
        self.ent_init.delete(0, tk.END)
        self.ent_init.insert(0, "0")

    def add_transaction(self):
        mat_name = self.cb_material.get()
        t_type = self.cb_type.get()
        user = self.ent_user.get()
        
        try:
            qty = float(self.ent_qty.get())
        except ValueError:
            messagebox.showwarning("입력 오류", "수량은 숫자여야 합니다.")
            return
        note = self.ent_note.get()
        
        if not mat_name or not t_type:
            messagebox.showwarning("입력 오류", "자재와 구분을 선택해주세요.")
            return
        
        mat_id = self.materials_df[self.materials_df['품명'] == mat_name]['Material ID'].values[0]
        
        new_trans = {
            'Date': datetime.datetime.now(),
            'Material ID': mat_id,
            'Type': t_type,
            'Quantity': qty,
            'Note': note,
            'User': user
        }
        self.transactions_df = pd.concat([self.transactions_df, pd.DataFrame([new_trans])], ignore_index=True)
        self.save_data()
        self.update_stock_view()
        messagebox.showinfo("완료", f"{mat_name} {t_type} 처리되었습니다.")
        
        self.ent_qty.delete(0, tk.END)
        self.ent_note.delete(0, tk.END)
        self.ent_user.delete(0, tk.END)

    def setup_report_tab(self):
        # Top control frame
        report_frame = ttk.Frame(self.tab_reports)
        report_frame.pack(pady=10, fill='x', padx=10)
        
        ttk.Label(report_frame, text="연도:").grid(row=0, column=0, padx=5)
        self.cb_year = ttk.Combobox(report_frame, values=[str(y) for y in range(2024, 2031)], width=10)
        self.cb_year.grid(row=0, column=1, padx=5)
        self.cb_year.set(str(datetime.datetime.now().year))
        
        ttk.Label(report_frame, text="월:").grid(row=1, column=0, padx=5, pady=10)
        self.cb_month = ttk.Combobox(report_frame, values=[str(m) for m in range(1, 13)], width=10)
        self.cb_month.grid(row=1, column=1, padx=5, pady=10)
        self.cb_month.set(str(datetime.datetime.now().month))
        
        # Button to view monthly usage
        btn_view_usage = ttk.Button(report_frame, text="품명별 월사용량 조회", command=self.view_monthly_usage)
        btn_view_usage.grid(row=2, column=0, columnspan=2, pady=5)
        
        btn_year_report = ttk.Button(report_frame, text="년별 사용량 엑셀 추출", command=self.generate_yearly_report)
        btn_year_report.grid(row=3, column=0, columnspan=2, pady=5)
        
        btn_month_report = ttk.Button(report_frame, text="월별 사용량 엑셀 추출", command=self.generate_monthly_report)
        btn_month_report.grid(row=4, column=0, columnspan=2, pady=5)
        
        # Treeview frame for displaying monthly usage by item
        tree_frame = ttk.LabelFrame(self.tab_reports, text="품명별 월사용량")
        tree_frame.pack(expand=True, fill='both', padx=10, pady=10)
        
        # Scrollbars
        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")
        
        # Treeview for usage display
        columns = ('품명', '관리품번', '규격', '단위', '월사용량', '누계사용량')
        self.usage_tree = ttk.Treeview(tree_frame, columns=columns, show='headings',
                                       yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        vsb.config(command=self.usage_tree.yview)
        hsb.config(command=self.usage_tree.xview)
        
        # Column configuration
        col_widths = [200, 120, 150, 80, 100, 120]
        for col, width in zip(columns, col_widths):
            self.usage_tree.heading(col, text=col)
            self.usage_tree.column(col, width=width)
        
        # Grid layout
        self.usage_tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

    def setup_import_tab(self):
        import_frame = ttk.LabelFrame(self.tab_import, text="데이터 관리")
        import_frame.pack(pady=20, padx=20, fill='both', expand=True)
        
        # Import Section
        ttk.Label(import_frame, text="엑셀 파일에서 자재 데이터 가져오기", font=('Arial', 11, 'bold')).pack(pady=10)
        ttk.Label(import_frame, text="형식: Material ID, 회사코드, 관리품번, 품명, 창고, 모델명, 규격, 품목군코드, 제조사, 제조국, 가격, 관리단위, 수량", 
                 wraplength=600).pack(pady=5)
        
        btn_import = ttk.Button(import_frame, text="엑셀 파일 가져오기", command=self.import_from_excel)
        btn_import.pack(pady=10)
        
        ttk.Separator(import_frame, orient='horizontal').pack(fill='x', pady=20)
        
        # Export Section
        ttk.Label(import_frame, text="현재 데이터 엑셀로 내보내기", font=('Arial', 11, 'bold')).pack(pady=10)
        
        btn_export_materials = ttk.Button(import_frame, text="자재 목록 내보내기", command=self.export_materials)
        btn_export_materials.pack(pady=5)
        
        btn_export_trans = ttk.Button(import_frame, text="거래 내역 내보내기", command=self.export_transactions)
        btn_export_trans.pack(pady=5)
        
        btn_export_all = ttk.Button(import_frame, text="전체 데이터 내보내기", command=self.export_all)
        btn_export_all.pack(pady=5)
    
    def import_from_excel(self):
        file_path = filedialog.askopenfilename(
            title="엑셀 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        
        if not file_path:
            return
        
        try:
            imported_df = pd.read_excel(file_path)
            
            # Validate columns - accept both Korean and English column names
            required_cols = ['품명'] if '품명' in imported_df.columns else ['Item Name']
            if not any(col in imported_df.columns for col in ['품명', 'Item Name']):
                messagebox.showerror("오류", "필수 컬럼 '품명' 또는 'Item Name'이 없습니다.")
                return
            
            # Process each row
            count = 0
            for _, row in imported_df.iterrows():
                # Generate new Material ID
                if self.materials_df.empty:
                    mat_id = 1
                else:
                    mat_id = self.materials_df['Material ID'].max() + 1
                
                # Support both Korean and English column names
                new_row = {
                    'Material ID': mat_id,
                    '회사코드': row.get('회사코드', row.get('Company Code', '')),
                    '관리품번': row.get('관리품번', row.get('Equipment Code', '')),
                    '품명': row.get('품명', row.get('Item Name', '')),
                    '창고': row.get('창고', row.get('Warehouse', '')),
                    '모델명': row.get('모델명', row.get('Model', '')),
                    '규격': row.get('규격', row.get('Specification', '')),
                    '품목군코드': row.get('품목군코드', row.get('Classification', '')),
                    '제조사': row.get('제조사', row.get('Manufacturer', '')),
                    '제조국': row.get('제조국', row.get('Country', '')),
                    '가격': row.get('가격', row.get('Price', 0)),
                    '관리단위': row.get('관리단위', row.get('Unit', 'EA')),
                    '수량': row.get('수량', row.get('Initial Stock', row.get('Current Stock', 0)))
                }
                
                self.materials_df = pd.concat([self.materials_df, pd.DataFrame([new_row])], ignore_index=True)
                count += 1
            
            self.save_data()
            self.update_material_combo()
            self.update_stock_view()
            messagebox.showinfo("완료", f"{count}개의 자재가 가져와졌습니다.")
            
        except Exception as e:
            messagebox.showerror("오류", f"파일을 가져오는데 실패했습니다: {e}")
    
    def export_materials(self):
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile="Materials_Export.xlsx",
            title="자재 목록 저장",
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if save_path:
            try:
                self.materials_df.to_excel(save_path, index=False)
                messagebox.showinfo("완료", "자재 목록이 저장되었습니다.")
            except Exception as e:
                messagebox.showerror("오류", f"저장 실패: {e}")
    
    def export_transactions(self):
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile="Transactions_Export.xlsx",
            title="거래 내역 저장",
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if save_path:
            try:
                self.transactions_df.to_excel(save_path, index=False)
                messagebox.showinfo("완료", "거래 내역이 저장되었습니다.")
            except Exception as e:
                messagebox.showerror("오류", f"저장 실패: {e}")
    
    def export_all(self):
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile="Complete_Export.xlsx",
            title="전체 데이터 저장",
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if save_path:
            try:
                with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                    self.materials_df.to_excel(writer, sheet_name='Materials', index=False)
                    self.transactions_df.to_excel(writer, sheet_name='Transactions', index=False)
                messagebox.showinfo("완료", "전체 데이터가 저장되었습니다.")
            except Exception as e:
                messagebox.showerror("오류", f"저장 실패: {e}")

    def view_monthly_usage(self):
        """Display monthly usage by item name in the treeview"""
        # Clear current view
        for item in self.usage_tree.get_children():
            self.usage_tree.delete(item)
        
        year = int(self.cb_year.get())
        month = int(self.cb_month.get())
        
        # Filter transactions for the selected month
        month_mask = (self.transactions_df['Date'].dt.year == year) & \
                     (self.transactions_df['Date'].dt.month == month) & \
                     (self.transactions_df['Type'] == 'OUT')
        monthly_trans = self.transactions_df[month_mask]
        
        # Filter transactions for cumulative (from start of year to selected month)
        cumulative_mask = (self.transactions_df['Date'].dt.year == year) & \
                          (self.transactions_df['Date'].dt.month <= month) & \
                          (self.transactions_df['Type'] == 'OUT')
        cumulative_trans = self.transactions_df[cumulative_mask]
        
        # Build usage data for each material
        usage_data = []
        for _, mat in self.materials_df.iterrows():
            mat_id = mat['Material ID']
            
            # Calculate monthly usage
            month_usage = monthly_trans[monthly_trans['Material ID'] == mat_id]['Quantity'].sum()
            
            # Calculate cumulative usage (year-to-date)
            cumulative_usage = cumulative_trans[cumulative_trans['Material ID'] == mat_id]['Quantity'].sum()
            
            # Only show items with usage
            if month_usage > 0 or cumulative_usage > 0:
                usage_data.append({
                    '품명': mat.get('품명', ''),
                    '관리품번': mat.get('관리품번', ''),
                    '규격': mat.get('규격', ''),
                    '단위': mat.get('관리단위', 'EA'),
                    '월사용량': month_usage,
                    '누계사용량': cumulative_usage
                })
        
        # Sort by item name
        usage_data.sort(key=lambda x: x['품명'])
        
        # Display in treeview
        for data in usage_data:
            self.usage_tree.insert('', tk.END, values=(
                data['품명'],
                data['관리품번'],
                data['규격'],
                data['단위'],
                f"{data['월사용량']:.1f}",
                f"{data['누계사용량']:.1f}"
            ))
        
        # Show message if no data
        if not usage_data:
            messagebox.showinfo("알림", f"{year}년 {month}월에 사용 내역이 없습니다.")

    def generate_yearly_report(self):
        year = int(self.cb_year.get())
        # Filter transactions for the year
        mask = (self.transactions_df['Date'].dt.year == year) & (self.transactions_df['Type'] == 'OUT')
        yearly_trans = self.transactions_df[mask]
        
        report = []
        for _, mat in self.materials_df.iterrows():
            mat_id = mat['Material ID']
            
            # Calculate monthly usage for each month
            row_data = {
                '설비코드': mat.get('관리품번', ''),
                '자재명': mat.get('품명', ''),
                '분류': mat.get('품목군코드', ''),
                '규격': mat.get('규격', ''),
                '단위': mat.get('관리단위', ''),
                '제조사': mat.get('제조사', '')
            }
            
            # Add monthly columns (1월 ~ 12월)
            monthly_values = []
            for month in range(1, 13):
                month_mask = (yearly_trans['Material ID'] == mat_id) & \
                            (yearly_trans['Date'].dt.month == month)
                month_usage = yearly_trans[month_mask]['Quantity'].sum()
                row_data[f'{month}월'] = month_usage
                monthly_values.append(month_usage)
            
            # Calculate totals
            total = sum(monthly_values)
            row_data['합계'] = total
            
            # Calculate cumulative total
            cumulative = 0
            for i, val in enumerate(monthly_values, 1):
                cumulative += val
                if i == 12:  # Only show final cumulative at the end
                    row_data['누계'] = cumulative
            
            report.append(row_data)
            
        report_df = pd.DataFrame(report)
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                                 initialfile=f"Yearly_Usage_{year}.xlsx",
                                                 title="보고서 저장")
        if save_path:
            report_df.to_excel(save_path, index=False)
            messagebox.showinfo("완료", f"{year}년 보고서가 저장되었습니다.")

    def generate_monthly_report(self):
        year = int(self.cb_year.get())
        month = int(self.cb_month.get())
        
        mask = (self.transactions_df['Date'].dt.year == year) & \
               (self.transactions_df['Date'].dt.month == month) & \
               (self.transactions_df['Type'] == 'OUT')
        monthly_trans = self.transactions_df[mask]
        
        report = []
        for _, mat in self.materials_df.iterrows():
            mat_id = mat['Material ID']
            total_usage = monthly_trans[monthly_trans['Material ID'] == mat_id]['Quantity'].sum()
            report.append({
                '설비코드': mat.get('관리품번', ''),
                '자재명': mat.get('품명', ''),
                '분류': mat.get('품목군코드', ''),
                '규격': mat.get('규격', ''),
                '단위': mat.get('관리단위', ''),
                '제조사': mat.get('제조사', ''),
                '월간 총 사용량': total_usage
            })
            
        report_df = pd.DataFrame(report)
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                                 initialfile=f"Monthly_Usage_{year}_{month}.xlsx",
                                                 title="보고서 저장")
        if save_path:
            report_df.to_excel(save_path, index=False)
            messagebox.showinfo("완료", f"{year}년 {month}월 보고서가 저장되었습니다.")

if __name__ == "__main__":
    root = tk.Tk()
    app = MaterialManager(root)
    root.mainloop()
