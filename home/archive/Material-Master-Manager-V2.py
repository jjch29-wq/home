import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
import pandas as pd
import datetime
import os
import json
import sys
import ctypes

class MaterialManager:
    def __init__(self, root):
        self.root = root
        
        # High DPI awareness
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass
                
        self.root.title("자재 및 소모품 관리 시스템 (Material Manager)")
        self.root.geometry("1600x900")
        
        # Configure overall style
        self.style = ttk.Style()
        self.style.configure(".", font=('Malgun Gothic', 10))
        self.style.configure("Treeview.Heading", font=('Malgun Gothic', 10, 'bold'))
        self.style.configure("Treeview", font=('Malgun Gothic', 10))
        
        # Determine base directory and bundle directory for portability
        if getattr(sys, 'frozen', False):
            # If running as an executable
            self.app_dir = os.path.dirname(sys.executable)
            self.bundle_dir = getattr(sys, '_MEIPASS', self.app_dir)
        else:
            # If running as a script
            self.app_dir = os.path.dirname(os.path.abspath(__file__))
            self.bundle_dir = self.app_dir
            
        self.db_path = os.path.join(self.app_dir, 'Material_Inventory.xlsx')
        self.config_path = os.path.join(self.app_dir, 'Material_Manager_Config.json')
        
        self.sites = [] # Initialize site list
        self.users = [] # Initialize user/name list
        self.load_data()
        
        # Centralized list of Carestream films for suggestions
        self.carestream_films = [
            "Carestream AA400-3⅓*12\"",
            "Carestream AA400-3⅓*17\"", 
            "Carestream AA400-4½*12\"",
            "Carestream AA400-10*12\"",
            "Carestream M100-3⅓*12\"",
            "Carestream M100-10*12\"",
            "Carestream M100-14*17\"",
            "Carestream MX125-3⅓*6\"",
            "Carestream MX125-3⅓*12\"",
            "Carestream MX125-4½*12\"",
            "Carestream MX125-10*12\"",
            "Carestream MX125-14*17\"",
            "Carestream T200-3⅓*12\"",
            "Carestream T200-3⅓*17\"",
            "Carestream T200-4½*12\"",
            "Carestream T200-10*12\"",
            "Carestream T200-14*17\""
        ]
        
        self.create_widgets()
        self.update_registration_combos()
        
        # Enable keyboard navigation
        self.setup_keyboard_shortcuts()
        
        # Load and restore tab configuration
        self.load_tab_config()
        
        # Save tab config on window close
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def extract_sn_from_model(self, model_name, current_sn):
        """Extract SN from model name if 'S/N.' is present and update current_sn if it's empty"""
        if pd.isna(model_name):
            return model_name, current_sn
        
        model_str = str(model_name)
        if 'S/N.' in model_str:
            parts = model_str.split('S/N.', 1)
            new_model = parts[0].strip()
            # Remove trailing underscore or dot if present before S/N.
            if new_model.endswith('_') or new_model.endswith('.'):
                new_model = new_model[:-1].strip()
            
            extracted_sn = parts[1].strip()
            
            # Use extracted SN if current SN is empty or matches part of it
            if not str(current_sn).strip() or str(current_sn) == 'nan':
                return new_model, extracted_sn
            else:
                # If current SN exists, we still clean up the model name
                return new_model, current_sn
        
        return model_name, current_sn

    def load_data(self):
        try:
            # Check if database exists in app_dir. If not, try to restore from bundle_dir
            if not os.path.exists(self.db_path):
                bundled_db = os.path.join(self.bundle_dir, 'Material_Inventory.xlsx')
                if os.path.exists(bundled_db):
                    import shutil
                    try:
                        shutil.copy2(bundled_db, self.db_path)
                        # Also try to copy config if it exists in bundle but not in app_dir
                        bundled_config = os.path.join(self.bundle_dir, 'Material_Manager_Config.json')
                        if os.path.exists(bundled_config) and not os.path.exists(self.config_path):
                            shutil.copy2(bundled_config, self.config_path)
                    except Exception as e:
                        print(f"Failed to restore data from bundle: {e}")

            if not os.path.exists(self.db_path):
                # Initialize with new schema if still not found
                self.materials_df = pd.DataFrame(columns=[
                    'Material ID', '회사코드', '관리품번', '품목명', 'SN', '창고',
                    '모델명', '규격', '품목군코드', '공급업체', '제조사', '제조국', 
                    '가격', '관리단위', '수량', '재고하한'
                ])
                self.transactions_df = pd.DataFrame(columns=['Date', 'Material ID', 'Type', 'Quantity', 'Note', 'User'])
                self.monthly_usage_df = pd.DataFrame(columns=['Material ID', 'Year', 'Month', 'Site', 'Usage', 'Note', 'Entry Date'])
                self.daily_usage_df = pd.DataFrame(columns=['Date', 'Site', 'Material ID', 'Usage', 'Note', 'Entry Time', 
                                                'RTK_센터미스', 'RTK_농도', 'RTK_마킹미스', 'RTK_필름마크', 
                                                'RTK_취급부주의', 'RTK_고객불만', 'RTK_기타'])
            else:
                self.materials_df = pd.read_excel(self.db_path, sheet_name='Materials')
                # Handle column rename from '품명' to '품목명'
                if '품명' in self.materials_df.columns and '품목명' not in self.materials_df.columns:
                    self.materials_df.rename(columns={'품명': '품목명'}, inplace=True)
                
                # Add missing columns for backward compatibility
                missing_cols = {
                    'SN': '',
                    '공급업체': '',
                    '재고하한': 0
                }
                for col, default in missing_cols.items():
                    if col not in self.materials_df.columns:
                        self.materials_df[col] = default
                
                self.transactions_df = pd.read_excel(self.db_path, sheet_name='Transactions')
                # Ensure Date column is datetime
                if not self.transactions_df.empty:
                    self.transactions_df['Date'] = pd.to_datetime(self.transactions_df['Date'])
                
                # Load monthly usage data if exists
                try:
                    self.monthly_usage_df = pd.read_excel(self.db_path, sheet_name='MonthlyUsage')
                    if not self.monthly_usage_df.empty:
                        self.monthly_usage_df['Entry Date'] = pd.to_datetime(self.monthly_usage_df['Entry Date'])
                        # Add Site column if it doesn't exist (for backward compatibility)
                        if 'Site' not in self.monthly_usage_df.columns:
                            self.monthly_usage_df['Site'] = ''
                except:
                    self.monthly_usage_df = pd.DataFrame(columns=['Material ID', 'Year', 'Month', 'Site', 'Usage', 'Note', 'Entry Date'])
                
                # Load daily usage data if exists
                try:
                    self.daily_usage_df = pd.read_excel(self.db_path, sheet_name='DailyUsage')
                    if not self.daily_usage_df.empty:
                        self.daily_usage_df['Date'] = pd.to_datetime(self.daily_usage_df['Date'])
                        self.daily_usage_df['Entry Time'] = pd.to_datetime(self.daily_usage_df['Entry Time'])
                        # Add RTK columns if they don't exist (for backward compatibility)
                        rtk_columns = ['RTK_센터미스', 'RTK_농도', 'RTK_마킹미스', 'RTK_필름마크', 
                                      'RTK_취급부주의', 'RTK_고객불만', 'RTK_기타']
                        for col in rtk_columns:
                            if col not in self.daily_usage_df.columns:
                                self.daily_usage_df[col] = 0.0
                        # Remove old RTK Category column if exists
                        if 'RTK Category' in self.daily_usage_df.columns:
                            self.daily_usage_df = self.daily_usage_df.drop('RTK Category', axis=1)
                except:
                    self.daily_usage_df = pd.DataFrame(columns=['Date', 'Site', 'Material ID', 'Usage', 'Note', 'Entry Time', 
                                                        'RTK_센터미스', 'RTK_농도', 'RTK_마킹미스', 'RTK_필름마크', 
                                                        'RTK_취급부주의', 'RTK_고객불만', 'RTK_기타'])
                
                # Add SN column if it doesn't exist (for backward compatibility)
                if 'SN' not in self.materials_df.columns:
                    self.materials_df['SN'] = ''
                
                # Migrate old schema if needed
                if 'Equipment Code' in self.materials_df.columns and '회사코드' not in self.materials_df.columns:
                    self.migrate_old_schema()
                
                # Apply SN extraction from Model Name to existing data
                if not self.materials_df.empty:
                    updated = False
                    for idx, row in self.materials_df.iterrows():
                        model = row.get('모델명', '')
                        sn = row.get('SN', '')
                        new_model, new_sn = self.extract_sn_from_model(model, sn)
                        
                        if str(model) != str(new_model) or str(sn) != str(new_sn):
                            self.materials_df.at[idx, '모델명'] = new_model
                            self.materials_df.at[idx, 'SN'] = new_sn
                            updated = True
                    
                    if updated:
                        self.save_data()
        except Exception as e:
            messagebox.showerror("Error", f"데이터를 불러오는데 실패했습니다: {e}")
            self.materials_df = pd.DataFrame(columns=[
                'Material ID', '회사코드', '관리품번', '품목명', 'SN', '창고',
                '모델명', '규격', '품목군코드', '제조사', '제조국', 
                '가격', '관리단위', '수량'
            ])
            self.transactions_df = pd.DataFrame(columns=['Date', 'Material ID', 'Type', 'Quantity', 'Note', 'User'])
            self.monthly_usage_df = pd.DataFrame(columns=['Material ID', 'Year', 'Month', 'Site', 'Usage', 'Note', 'Entry Date'])
            self.daily_usage_df = pd.DataFrame(columns=['Date', 'Site', 'Material ID', 'Usage', 'Note', 'Entry Time', 
                                                'RTK_센터미스', 'RTK_농도', 'RTK_마킹미스', 'RTK_필름마크', 
                                                'RTK_취급부주의', 'RTK_고객불만', 'RTK_기타'])
    
    def migrate_old_schema(self):
        """Migrate data from old schema to new schema"""
        old_df = self.materials_df.copy()
        self.materials_df = pd.DataFrame(columns=[
            'Material ID', '회사코드', '관리품번', '품목명', 'SN', '창고',
            '모델명', '규격', '품목군코드', '제조사', '제조국', 
            '가격', '관리단위', '수량'
        ])
        
        for _, row in old_df.iterrows():
            new_row = {
                'Material ID': row.get('Material ID', ''),
                '회사코드': '',
                '관리품번': row.get('Equipment Code', ''),
                '품목명': row.get('Item Name', row.get('Name', '')),
                'SN': row.get('SN', ''),
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
                self.daily_usage_df.to_excel(writer, sheet_name='DailyUsage', index=False)
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
        self.notebook.add(self.tab_monthly_usage, text='월별 집계')
        self.setup_monthly_usage_tab()
        
        # Tab 6: Daily Usage Entry by Site
        self.tab_daily_usage = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_daily_usage, text='현장별 일일 사용량 기입')
        self.setup_daily_usage_tab()

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
        
        # Delete Button
        btn_delete = ttk.Button(control_frame, text="품목 삭제", command=self.delete_selected_material)
        btn_delete.pack(side='left', padx=5)
        
        # Edit Button
        btn_edit = ttk.Button(control_frame, text="품목 수정", command=self.open_edit_material_dialog)
        btn_edit.pack(side='left', padx=5)
        
        # Export Button
        btn_export = ttk.Button(control_frame, text="엑셀 내보내기", command=self.export_stock_to_excel)
        btn_export.pack(side='left', padx=5)
        
        # Select All Button
        btn_select_all = ttk.Button(control_frame, text="전체 선택", command=self.select_all_stock)
        btn_select_all.pack(side='left', padx=5)
        
        # Search and Filter Frame
        filter_frame = ttk.LabelFrame(control_frame, text="검색 필터")
        filter_frame.pack(side='left', padx=(20, 5), pady=2)
        
        # Row 0 of Filter Frame
        ttk.Label(filter_frame, text="회사:").grid(row=0, column=0, padx=2, pady=2, sticky='e')
        self.cb_filter_co = ttk.Combobox(filter_frame, width=15, state="readonly")
        self.cb_filter_co.grid(row=0, column=1, padx=2, pady=2)
        self.cb_filter_co.bind('<<ComboboxSelected>>', lambda e: self.update_stock_view())
        
        ttk.Label(filter_frame, text="분류:").grid(row=0, column=2, padx=2, pady=2, sticky='e')
        self.cb_filter_class = ttk.Combobox(filter_frame, width=15, state="readonly")
        self.cb_filter_class.grid(row=0, column=3, padx=2, pady=2)
        self.cb_filter_class.bind('<<ComboboxSelected>>', lambda e: self.update_stock_view())
        
        ttk.Label(filter_frame, text="제조사:").grid(row=0, column=4, padx=2, pady=2, sticky='e')
        self.cb_filter_mfr = ttk.Combobox(filter_frame, width=15, state="readonly")
        self.cb_filter_mfr.grid(row=0, column=5, padx=2, pady=2)
        self.cb_filter_mfr.bind('<<ComboboxSelected>>', lambda e: self.update_stock_view())
        
        ttk.Label(filter_frame, text="품목명:").grid(row=0, column=6, padx=2, pady=2, sticky='e')
        self.cb_filter_name = ttk.Combobox(filter_frame, width=25, state="readonly")
        self.cb_filter_name.grid(row=0, column=7, padx=2, pady=2)
        self.cb_filter_name.bind('<<ComboboxSelected>>', lambda e: self.update_stock_view())
        
        # Row 1 of Filter Frame
        ttk.Label(filter_frame, text="S/N:").grid(row=1, column=0, padx=2, pady=2, sticky='e')
        self.cb_filter_sn = ttk.Combobox(filter_frame, width=20, state="readonly")
        self.cb_filter_sn.grid(row=1, column=1, padx=2, pady=2)
        self.cb_filter_sn.bind('<<ComboboxSelected>>', lambda e: self.update_stock_view())
        
        ttk.Label(filter_frame, text="모델명:").grid(row=1, column=2, padx=2, pady=2, sticky='e')
        self.cb_filter_model = ttk.Combobox(filter_frame, width=20, state="readonly")
        self.cb_filter_model.grid(row=1, column=3, padx=2, pady=2)
        self.cb_filter_model.bind('<<ComboboxSelected>>', lambda e: self.update_stock_view())
        
        ttk.Label(filter_frame, text="관리품번:").grid(row=1, column=4, padx=2, pady=2, sticky='e')
        self.cb_filter_eq = ttk.Combobox(filter_frame, width=20, state="readonly")
        self.cb_filter_eq.grid(row=1, column=5, padx=2, pady=2)
        self.cb_filter_eq.bind('<<ComboboxSelected>>', lambda e: self.update_stock_view())
        
        ttk.Label(filter_frame, text="검색어:").grid(row=1, column=6, padx=2, pady=2, sticky='e')
        self.search_var = tk.StringVar()
        self.search_var.trace_add('write', lambda *args: self.update_stock_view())
        search_entry = ttk.Entry(filter_frame, textvariable=self.search_var, width=20)
        search_entry.grid(row=1, column=7, padx=2, pady=2)
        
        # Reset Filters Button
        btn_reset = ttk.Button(filter_frame, text="필터 초기화", command=self.reset_stock_filters)
        btn_reset.grid(row=1, column=8, padx=10, pady=2)
        
        # Treeview for Stock with Scrollbars
        tree_frame = ttk.Frame(self.tab_stock)
        tree_frame.pack(expand=True, fill='both', padx=5, pady=5)
        
        # Scrollbars
        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")
        
        columns = ('ID', '회사코드', '관리품번', '품목명', 'SN', '창고', '모델명', '규격', '품목군코드', '공급업체', '제조사', '제조국', '가격', '관리단위', '수량', '재고하한')
        self.stock_tree = ttk.Treeview(tree_frame, columns=columns, show='headings', 
                                      yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        vsb.config(command=self.stock_tree.yview)
        hsb.config(command=self.stock_tree.xview)
        
        # Column configuration
        col_widths = [50, 70, 90, 150, 80, 80, 100, 100, 80, 100, 100, 70, 70, 70, 70, 70]
        for col, width in zip(columns, col_widths):
            self.stock_tree.heading(col, text=col)
            self.stock_tree.column(col, width=width, anchor='center')
        
        # Bind double-click
        self.stock_tree.bind('<Double-1>', lambda e: self.open_edit_material_dialog())
        
        # Grid layout
        self.stock_tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        self.update_stock_view()
    
    def show_low_stock(self):
        """Show items with low stock (less than their specific reorder point)"""
        low_stock_items = []
        for _, mat in self.materials_df.iterrows():
            current = self.calculate_current_stock(mat['Material ID'])
            reorder_point = mat.get('재고하한', 10)
            if pd.isna(reorder_point) or reorder_point <= 0:
                reorder_point = 10 # Default fallback
                
            if pd.notna(current) and current < reorder_point:
                low_stock_items.append((mat.get('품목명', ''), current, reorder_point))
        
        if not low_stock_items:
            messagebox.showinfo("재고 알림", "수량이 10개 미만인 항목이 없습니다.")
        else:
            msg = "다음 항목들의 재고가 부족합니다:\n\n"
            for item, current, reorder in low_stock_items:
                msg += f"• {item}: 현재 {current:g} (필요 수준: {reorder:g})\n"
            messagebox.showwarning("재고 부족", msg)
    
    def select_all_stock(self):
        """Select all items in the stock treeview"""
        all_items = self.stock_tree.get_children()
        self.stock_tree.selection_set(all_items)

    def delete_selected_material(self):
        """선택된 자재 항목들을 재고에서 삭제"""
        selected_items = self.stock_tree.selection()
        
        if not selected_items:
            messagebox.showwarning("선택 오류", "삭제할 품목을 선택해주세요.")
            return
        
        # Confirm deletion
        confirm = messagebox.askyesno("삭제 확인", f"선택한 {len(selected_items)}개의 품목을 재고에서 영구히 삭제하시겠습니까?\n이 작업은 되돌릴 수 없습니다.")
        
        if not confirm:
            return
            
        # Get Material IDs to delete
        mat_ids_to_remove = []
        for item in selected_items:
            values = self.stock_tree.item(item, 'values')
            if values:
                # Ensure we match the type of Material ID in the dataframe
                mat_ids_to_remove.append(type(self.materials_df['Material ID'].iloc[0])(values[0]))
        
        # Remove from materials_df
        initial_count = len(self.materials_df)
        self.materials_df = self.materials_df[~self.materials_df['Material ID'].isin(mat_ids_to_remove)]
        removed_count = initial_count - len(self.materials_df)
        
        if removed_count > 0:
            # Save data and update views
            self.save_data()
            self.update_stock_view()
            self.update_material_combo()
            
            # Optional: Clear transactions related to these materials?
            # For now, let's keep them for history unless explicitly asked.
            
            messagebox.showinfo("완료", f"{removed_count}개의 품목이 삭제되었습니다.")
        else:
            messagebox.showwarning("실패", "데이터프레임에서 항목을 삭제하지 못했습니다.")

    def reset_stock_filters(self):
        """Reset all stock filters to default"""
        self.cb_filter_co.set("전체")
        self.cb_filter_class.set("전체")
        self.cb_filter_mfr.set("전체")
        self.cb_filter_name.set("전체")
        self.cb_filter_sn.set("전체")
        self.cb_filter_model.set("전체")
        self.cb_filter_eq.set("전체")
        self.search_var.set("")
        self.update_stock_view()

    def open_edit_material_dialog(self):
        """Open a dialog to edit the selected material"""
        selection = self.stock_tree.selection()
        if not selection:
            messagebox.showwarning("선택 오류", "수정할 자재를 선택해주세요.")
            return
            
        item = self.stock_tree.item(selection[0])
        mat_values = item['values']
        mat_id = mat_values[0]
        
        # Get full material data from DF
        mat_data = self.materials_df[self.materials_df['Material ID'] == mat_id].iloc[0]
        
        # Create Edit Dialog
        edit_win = tk.Toplevel(self.root)
        edit_win.title("자재 정보 수정")
        edit_win.geometry("500x600")
        edit_win.transient(self.root)
        edit_win.grab_set()
        
        main_frame = ttk.Frame(edit_win, padding=20)
        main_frame.pack(expand=True, fill='both')
        
        fields = [
            ('회사코드', '회사코드'),
            ('관리품번', '관리품번'),
            ('품목명', '품목명'),
            ('SN', 'SN'),
            ('창고', '창고'),
            ('모델명', '모델명'),
            ('규격', '규격'),
            ('품목군코드', '품목군코드'),
            ('공급업체', '공급업체'),
            ('제조사', '제조사'),
            ('제조국', '제조국'),
            ('관리단위', '관리단위'),
            ('수량', '수량'),
            ('재고하한', '재고하한')
        ]
        
        entries = {}
        for i, (label_text, col_name) in enumerate(fields):
            ttk.Label(main_frame, text=f"{label_text}:").grid(row=i, column=0, padx=5, pady=5, sticky='w')
            
            # Using Combobox for some fields to maintain consistency
            if col_name in ['회사코드', '품목군코드', '제조사', '제조국', '관리단위']:
                ent = ttk.Combobox(main_frame, width=35)
                # Populate values from existing data
                if col_name in self.materials_df.columns:
                    unique_vals = sorted(self.materials_df[col_name].dropna().unique().tolist())
                    ent['values'] = unique_vals
            else:
                ent = ttk.Entry(main_frame, width=38)
                
            ent.grid(row=i, column=1, padx=5, pady=5)
            
            # Pre-fill value
            val = mat_data.get(col_name, '')
            if pd.isna(val): val = ''
            ent.insert(0, str(val))
            entries[col_name] = ent
            
        def on_save():
            new_data = {col: ent.get() for col, ent in entries.items()}
            self.save_material_edits(mat_id, new_data)
            edit_win.destroy()
            
        btn_save = ttk.Button(main_frame, text="변경사항 저장", command=on_save)
        btn_save.grid(row=len(fields), column=0, columnspan=2, pady=20)

    def save_material_edits(self, mat_id, new_data):
        """Save edited material data back to the database"""
        # Update DataFrame
        idx = self.materials_df.index[self.materials_df['Material ID'] == mat_id].tolist()
        if not idx:
            messagebox.showerror("오류", "자재를 찾을 수 없습니다.")
            return
            
        for col, val in new_data.items():
            if col in self.materials_df.columns:
                # Ensure numeric fields are correctly typed
                if col in ['수량', '재고하한', '가격']:
                    try:
                        val = float(val) if str(val).strip() else 0.0
                    except ValueError:
                        val = 0.0
                self.materials_df.at[idx[0], col] = val
        
        # Re-check SN extraction after edit
        model = self.materials_df.at[idx[0], '모델명']
        sn = self.materials_df.at[idx[0], 'SN']
        new_model, new_sn = self.extract_sn_from_model(model, sn)
        self.materials_df.at[idx[0], '모델명'] = new_model
        self.materials_df.at[idx[0], 'SN'] = new_sn
        
        # Save to Excel
        self.save_data()
        
        # Refresh everything
        self.update_stock_view()
        self.update_registration_combos()
        self.update_material_combo()
        
        messagebox.showinfo("완료", "자재 정보가 성공적으로 수정되었습니다.")

    
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
            # Check if stored_qty is NaN and treat it as 0
            if pd.isna(stored_qty):
                stored_qty = 0
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
        
        # Helper function to safely get value and replace NaN
        def safe_get(val, default=''):
            if pd.isna(val):
                return default
            return val
        
        # Calculate current stock
        stock_summary = []
        for _, mat in self.materials_df.iterrows():
            mat_id = mat['Material ID']
            current_stock = self.calculate_current_stock(mat_id)
            
            # Note: We NO LONGER update materials_df['수량'] here to avoid infinite deduction on refresh.
            # materials_df['수량'] stays as the base/initial stock, and current_stock is calculated on-the-fly.
            
            stock_summary.append((
                mat_id,
                safe_get(mat.get('회사코드', ''), ''),
                safe_get(mat.get('관리품번', ''), ''),
                safe_get(mat.get('품목명', ''), ''),
                safe_get(mat.get('SN', ''), ''),
                safe_get(mat.get('창고', ''), ''),
                safe_get(mat.get('모델명', ''), ''),
                safe_get(mat.get('규격', ''), ''),
                safe_get(mat.get('품목군코드', ''), ''),
                safe_get(mat.get('공급업체', ''), ''),
                safe_get(mat.get('제조사', ''), ''),
                safe_get(mat.get('제조국', ''), ''),
                safe_get(mat.get('가격', 0), 0),
                safe_get(mat.get('관리단위', 'EA'), 'EA'),
                safe_get(current_stock, 0),
                safe_get(mat.get('재고하한', 0), 0)
            ))
        
        # Filter by search term and dropdowns
        filter_co = self.cb_filter_co.get()
        filter_class = self.cb_filter_class.get()
        filter_mfr = self.cb_filter_mfr.get()
        filter_name = self.cb_filter_name.get()
        filter_sn = self.cb_filter_sn.get()
        filter_model = self.cb_filter_model.get()
        filter_eq = self.cb_filter_eq.get()
        
        for row in stock_summary:
            # Dropdown Filters
            if filter_co != "전체" and str(row[1]) != filter_co: continue
            if filter_class != "전체" and str(row[8]) != filter_class: continue
            if filter_mfr != "전체" and str(row[10]) != filter_mfr: continue
            if filter_name != "전체" and str(row[3]) != filter_name: continue
            if filter_sn != "전체" and str(row[4]) != filter_sn: continue
            if filter_model != "전체" and str(row[6]) != filter_model: continue
            if filter_eq != "전체" and str(row[2]) != filter_eq: continue
            
            # General Search Term
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
        ttk.Label(reg_frame, text="회사코드:").grid(row=0, column=0, padx=5, pady=2, sticky='w')
        self.cb_co_code = ttk.Combobox(reg_frame, width=25)
        self.cb_co_code.grid(row=0, column=1, padx=5, pady=2)
        
        ttk.Label(reg_frame, text="설비코드:").grid(row=0, column=2, padx=5, pady=2, sticky='w')
        self.cb_eq_code = ttk.Combobox(reg_frame, width=30)
        self.cb_eq_code.grid(row=0, column=3, padx=5, pady=2)
        
        # Row 1
        ttk.Label(reg_frame, text="품목명:").grid(row=1, column=0, padx=5, pady=2, sticky='w')
        self.cb_item_name = ttk.Combobox(reg_frame, width=25)
        self.cb_item_name.grid(row=1, column=1, padx=5, pady=2)
        
        ttk.Label(reg_frame, text="SN번호:").grid(row=1, column=2, padx=5, pady=2, sticky='w')
        self.ent_sn = ttk.Entry(reg_frame, width=30)
        self.ent_sn.grid(row=1, column=3, padx=5, pady=2)
        
        # Row 2
        ttk.Label(reg_frame, text="분류:").grid(row=2, column=0, padx=5, pady=2, sticky='w')
        self.cb_class = ttk.Combobox(reg_frame, width=25)
        self.cb_class.grid(row=2, column=1, padx=5, pady=2)
        
        ttk.Label(reg_frame, text="규격:").grid(row=2, column=2, padx=5, pady=2, sticky='w')
        self.cb_spec = ttk.Combobox(reg_frame, width=30)
        self.cb_spec.grid(row=2, column=3, padx=5, pady=2)
        
        # Row 3
        ttk.Label(reg_frame, text="단위:").grid(row=3, column=0, padx=5, pady=2, sticky='w')
        self.cb_unit = ttk.Combobox(reg_frame, width=25)
        self.cb_unit.grid(row=3, column=1, padx=5, pady=2)
        
        ttk.Label(reg_frame, text="공급업자:").grid(row=3, column=2, padx=5, pady=2, sticky='w')
        self.cb_supplier = ttk.Combobox(reg_frame, width=30)
        self.cb_supplier.grid(row=3, column=3, padx=5, pady=2)
        
        # Row 4
        ttk.Label(reg_frame, text="제조사:").grid(row=4, column=0, padx=5, pady=2, sticky='w')
        self.cb_mfr = ttk.Combobox(reg_frame, width=25)
        self.cb_mfr.grid(row=4, column=1, padx=5, pady=2)
        
        ttk.Label(reg_frame, text="제조국:").grid(row=4, column=2, padx=5, pady=2, sticky='w')
        self.cb_origin = ttk.Combobox(reg_frame, width=30)
        self.cb_origin.grid(row=4, column=3, padx=5, pady=2)
        
        # Row 5
        ttk.Label(reg_frame, text="재주문 수준:").grid(row=5, column=0, padx=5, pady=2, sticky='w')
        self.ent_reorder = ttk.Entry(reg_frame, width=20)
        self.ent_reorder.grid(row=5, column=1, padx=5, pady=2)
        self.ent_reorder.insert(0, "0")
        
        ttk.Label(reg_frame, text="초기재고:").grid(row=5, column=2, padx=5, pady=2, sticky='w')
        self.ent_init = ttk.Entry(reg_frame, width=25)
        self.ent_init.grid(row=5, column=3, padx=5, pady=2)
        self.ent_init.insert(0, "0")
        
        btn_reg = ttk.Button(reg_frame, text="자재 등록", command=self.register_material)
        btn_reg.grid(row=6, column=0, columnspan=4, pady=10)
        
        # Frame for In/Out Transaction
        trans_frame = ttk.LabelFrame(scrollable_frame, text="입출고 기록")
        trans_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(trans_frame, text="자재 선택:").grid(row=0, column=0, padx=5, pady=2, sticky='w')
        self.cb_material = ttk.Combobox(trans_frame, state="readonly", width=65, font=('Malgun Gothic', 10))
        self.cb_material.grid(row=0, column=1, padx=5, pady=2, columnspan=3)
        self.cb_material.bind('<<ComboboxSelected>>', self.on_material_selected)
        self.update_material_combo()
        
        # Model List display
        ttk.Label(trans_frame, text="관련 모델명:").grid(row=0, column=4, padx=(20, 5), pady=2, sticky='nw')
        self.list_models = tk.Listbox(trans_frame, height=4, width=30, font=('Malgun Gothic', 10))
        self.list_models.grid(row=0, column=5, rowspan=3, padx=5, pady=2, sticky='nsew')
        
        # Add scrollbar for model list
        model_vsb = ttk.Scrollbar(trans_frame, orient="vertical", command=self.list_models.yview)
        model_vsb.grid(row=0, column=6, rowspan=3, sticky='ns', pady=2)
        self.list_models.config(yscrollcommand=model_vsb.set)
        
        ttk.Label(trans_frame, text="구분:").grid(row=1, column=0, padx=5, pady=2, sticky='w')
        self.cb_type = ttk.Combobox(trans_frame, values=["IN", "OUT"], state="readonly", width=15)
        self.cb_type.grid(row=1, column=1, padx=5, pady=2, sticky='w')
        self.cb_type.set("OUT")
        
        ttk.Label(trans_frame, text="수량:").grid(row=1, column=2, padx=5, pady=2, sticky='w')
        self.ent_qty = ttk.Entry(trans_frame, width=30)
        self.ent_qty.grid(row=1, column=3, padx=5, pady=2)
        
        ttk.Label(trans_frame, text="현장:").grid(row=2, column=0, padx=5, pady=2, sticky='w')
        self.cb_trans_site = ttk.Combobox(trans_frame, width=28, values=self.sites)
        self.cb_trans_site.grid(row=2, column=1, padx=5, pady=2, sticky='w')
        # Auto-save new site entries
        self.cb_trans_site.bind('<FocusOut>', lambda e: self.auto_save_to_list(e, self.cb_trans_site, self.sites, 'sites'))
        self.cb_trans_site.bind('<Return>', lambda e: self.auto_save_to_list(e, self.cb_trans_site, self.sites, 'sites'))
        
        ttk.Label(trans_frame, text="비고:").grid(row=2, column=2, padx=5, pady=2, sticky='w')
        self.ent_note = ttk.Entry(trans_frame, width=30)
        self.ent_note.grid(row=2, column=3, padx=5, pady=2)
        
        ttk.Label(trans_frame, text="담당자:").grid(row=3, column=0, padx=5, pady=2, sticky='w')
        self.ent_user = ttk.Combobox(trans_frame, width=28, values=getattr(self, 'users', []))
        self.ent_user.grid(row=3, column=1, padx=5, pady=2)
        # Auto-save new manager entries
        self.ent_user.bind('<FocusOut>', lambda e: self.auto_save_to_list(e, self.ent_user, self.users, 'users'))
        self.ent_user.bind('<Return>', lambda e: self.auto_save_to_list(e, self.ent_user, self.users, 'users'))
        
        btn_trans = ttk.Button(trans_frame, text="기록 저장", command=self.add_transaction)
        btn_trans.grid(row=4, column=0, columnspan=4, pady=10)
        
        # Frame for displaying transaction history
        history_frame = ttk.LabelFrame(scrollable_frame, text="최근 입출고 내역")
        history_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Filter/Control frame for history
        history_ctrl_frame = ttk.Frame(history_frame)
        history_ctrl_frame.pack(fill='x', padx=5, pady=2)
        
        btn_del_trans = ttk.Button(history_ctrl_frame, text="선택 항목 삭제", command=self.delete_transaction_entry)
        btn_del_trans.pack(side='left', padx=5)
        
        btn_export_trans = ttk.Button(history_ctrl_frame, text="엑셀 내보내기", command=self.export_transaction_history)
        btn_export_trans.pack(side='left', padx=5)
        
        # Treeview for history
        tree_scroll_frame = ttk.Frame(history_frame)
        tree_scroll_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        inout_vsb = ttk.Scrollbar(tree_scroll_frame, orient="vertical")
        inout_hsb = ttk.Scrollbar(tree_scroll_frame, orient="horizontal")
        
        columns = ('날짜', '현장', '품목명', '구분', '수량', '담당자', '비고')
        self.inout_tree = ttk.Treeview(tree_scroll_frame, columns=columns, show='headings', height=10,
                                       yscrollcommand=inout_vsb.set, xscrollcommand=inout_hsb.set)
        
        inout_vsb.config(command=self.inout_tree.yview)
        inout_hsb.config(command=self.inout_tree.xview)
        
        col_widths = [150, 120, 200, 70, 70, 100, 200]
        for col, width in zip(columns, col_widths):
            self.inout_tree.heading(col, text=col)
            self.inout_tree.column(col, width=width, anchor='center')
        
        self.inout_tree.grid(row=0, column=0, sticky='nsew')
        inout_vsb.grid(row=0, column=1, sticky='ns')
        inout_hsb.grid(row=1, column=0, sticky='ew')
        
        tree_scroll_frame.grid_rowconfigure(0, weight=1)
        tree_scroll_frame.grid_columnconfigure(0, weight=1)
        
        # Initial populate
        self.update_transaction_view()
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")


    def get_material_display_name(self, mat_id):
        """Get formatted material name as '품목명 (SN: SN번호) - 규격'"""
        if self.materials_df.empty:
            return f"ID: {mat_id}"
            
        mat_row = self.materials_df[self.materials_df['Material ID'] == mat_id]
        if mat_row.empty:
            return f"ID: {mat_id}"
            
        mat = mat_row.iloc[0]
        name = mat['품목명']
        sn = mat.get('SN', '')
        spec = mat.get('규격', '')
        
        # Build display string
        display = name
        
        # Add SN if exists
        if sn and pd.notna(sn) and str(sn).strip():
            display += f" (SN: {sn})"
        
        # Add specification if exists
        if spec and pd.notna(spec) and str(spec).strip():
            display += f" - {spec}"
            
        return display

    def update_material_combo(self):
        """Update both In/Out and Daily Usage material comboboxes with unified list"""
        mat_list = []
        if not self.materials_df.empty:
            # Create list with unified display format
            for _, mat in self.materials_df.iterrows():
                display = self.get_material_display_name(mat['Material ID'])
                mat_list.append(display)
        
        # Merge database items with centralized films list, unique and sort
        all_vals = list(set([str(m) for m in mat_list + self.carestream_films if pd.notna(m) and str(m).strip()]))
        all_vals.sort()
        
        # Update ComboBoxes
        if hasattr(self, 'cb_material'):
            self.cb_material['values'] = all_vals
        if hasattr(self, 'cb_daily_material'):
            self.cb_daily_material['values'] = all_vals

    def update_registration_combos(self):
        """Update registration comboboxes with unique values from database and centralized list"""
        # 1. Update registration fields from database
        fields = {
            '회사코드': self.cb_co_code,
            '관리품번': self.cb_eq_code,
            '품목명': self.cb_item_name,
            '품목군코드': self.cb_class,
            '규격': self.cb_spec,
            '관리단위': self.cb_unit,
            '제조사': self.cb_mfr,
            '제조국': self.cb_origin
        }
        
        for col, combo in fields.items():
            vals = []
            if not self.materials_df.empty and col in self.materials_df.columns:
                unique_vals = self.materials_df[col].dropna().unique()
                vals = sorted([str(v).strip() for v in unique_vals if v and str(v).strip()])
            combo['values'] = vals
        
        # Add Carestream film options to 품목명 combobox (registration)
        if hasattr(self, 'cb_item_name'):
            existing_vals = list(self.cb_item_name['values']) if self.cb_item_name['values'] else []
            all_vals = list(set([str(mat) for mat in existing_vals + self.carestream_films if pd.notna(mat) and str(mat).strip()]))
            all_vals.sort()
            self.cb_item_name['values'] = all_vals
        
        # 2. Update Stock View filters
        filter_fields = {
            '회사코드': self.cb_filter_co,
            '품목군코드': self.cb_filter_class,
            '제조사': self.cb_filter_mfr,
            '품목명': self.cb_filter_name,
            'SN': self.cb_filter_sn,
            '모델명': self.cb_filter_model,
            '관리품번': self.cb_filter_eq
        }
        
        for col, combo in filter_fields.items():
            vals = []
            if not self.materials_df.empty and col in self.materials_df.columns:
                unique_vals = self.materials_df[col].dropna().unique()
                vals = sorted([str(v).strip() for v in unique_vals if v and str(v).strip()])
            
            combo['values'] = ["전체"] + vals
            if not combo.get():
                combo.set("전체")
        
        # Add Carestream film options to 품목명 filter in stock view
        if hasattr(self, 'cb_filter_name'):
            existing_vals = list(self.cb_filter_name['values']) if self.cb_filter_name['values'] else []
            all_vals = list(set([str(mat) for mat in existing_vals + self.carestream_films if pd.notna(mat) and str(mat).strip()]))
            # Ensure "전체" is at the top
            if "전체" in all_vals: all_vals.remove("전체")
            all_vals = ["전체"] + sorted(all_vals)
            self.cb_filter_name['values'] = all_vals

    def register_material(self):
        co_code = self.cb_co_code.get()
        item_name = self.cb_item_name.get()
        eq_code = self.cb_eq_code.get()
        sn = self.ent_sn.get()
        classification = self.cb_class.get()
        spec = self.cb_spec.get()
        unit = self.cb_unit.get()
        supplier = self.cb_supplier.get()
        manufacturer = self.cb_mfr.get()
        origin = self.cb_origin.get()
        
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
        
        # Extract SN from Model Name if present
        model_name = '' # Default empty
        new_model, new_sn = self.extract_sn_from_model(model_name, sn)
        # Note: In register_material, model_name is currently not an input field, but SN is.
        # If the user adds model name to registration later, this will handle it.
        
        new_row = {
            'Material ID': mat_id,
            '회사코드': co_code,
            '관리품번': eq_code,
            '품목명': item_name,
            'SN': new_sn,
            '창고': '',
            '모델명': new_model,
            '규격': spec,
            '품목군코드': classification,
            '공급업체': supplier,
            '제조사': manufacturer,
            '제조국': origin,
            '가격': 0,
            '관리단위': unit if unit else 'EA',
            '수량': init_stock,
            '재고하한': reorder_point
        }
        
        self.materials_df = pd.concat([self.materials_df, pd.DataFrame([new_row])], ignore_index=True)
        self.save_data()
        self.update_material_combo()
        self.update_registration_combos()
        self.update_stock_view()
        messagebox.showinfo("완료", f"'{item_name}' 자재가 등록되었습니다.")
        
        # Clear entries
        self.cb_co_code.set('')
        self.cb_eq_code.set('')
        self.cb_item_name.set('')
        self.ent_sn.delete(0, tk.END)
        self.cb_class.set('')
        self.cb_spec.set('')
        self.cb_unit.set('')
        self.cb_supplier.set('')
        self.cb_mfr.set('')
        self.cb_origin.set('')
        self.ent_reorder.delete(0, tk.END)
        self.ent_reorder.insert(0, "0")
        self.ent_init.delete(0, tk.END)
        self.ent_init.insert(0, "0")

    def add_transaction(self):
        mat_selection = self.cb_material.get()
        t_type = self.cb_type.get()
        user = self.ent_user.get()
        
        try:
            qty = float(self.ent_qty.get())
        except ValueError:
            messagebox.showwarning("입력 오류", "수량은 숫자여야 합니다.")
            return
        note = self.ent_note.get()
        
        if not mat_selection or not t_type:
            messagebox.showwarning("입력 오류", "자재와 구분을 선택해주세요.")
            return
        
        # Extract material name from selection (handle "품명 (SN: SN번호) - 규격" format)
        mat_name = mat_selection
        
        # Remove specification if present (comes after " - ")
        if " - " in mat_name:
            mat_name = mat_name.split(" - ")[0]
        
        # Remove SN if present (comes after " (SN: ")
        if " (SN: " in mat_name:
            mat_name = mat_name.split(" (SN: ")[0]
        
        pure_mat_name = mat_name
        
        # Get material ID using 재고현황's 품목명 field
        mat_rows = self.materials_df[self.materials_df['품목명'] == pure_mat_name]
        
        if mat_rows.empty:
            # Material doesn't exist in database, create it (auto-registration)
            new_mat_id = self.materials_df['Material ID'].max() + 1 if not self.materials_df.empty else 1
            
            new_material = {
                'Material ID': new_mat_id,
                '회사코드': '',
                '관리품번': '',
                '품목명': pure_mat_name,
                'SN': '',
                '창고': '',
                '모델명': '',
                '규격': '',
                '품목군코드': 'FILM' if 'Carestream' in pure_mat_name else 'OTHER',
                '제조사': 'Carestream' if 'Carestream' in pure_mat_name else '',
                '제조국': '',
                '가격': 0,
                '관리단위': 'EA',
                '수량': 0
            }
            
            self.materials_df = pd.concat([self.materials_df, pd.DataFrame([new_material])], ignore_index=True)
            self.save_data()
            mat_id = new_mat_id
        else:
            mat_id = mat_rows['Material ID'].values[0]
        
        new_trans = {
            'Date': datetime.datetime.now(),
            'Material ID': mat_id,
            'Type': t_type,
            'Quantity': qty,
            'Note': note,
            'User': user,
            'Site': self.cb_trans_site.get() if hasattr(self, 'cb_trans_site') else ''
        }
        self.transactions_df = pd.concat([self.transactions_df, pd.DataFrame([new_trans])], ignore_index=True)
        self.save_data()
        self.update_stock_view()
        self.update_material_combo() # Refresh lists just in case
        
        # Auto-save site to sites list if it's a new site
        site_value = self.cb_trans_site.get() if hasattr(self, 'cb_trans_site') else ''
        if site_value and site_value not in self.sites:
            self.sites.append(site_value)
            self.sites.sort()
            # Update all site comboboxes
            if hasattr(self, 'cb_trans_site'):
                self.cb_trans_site['values'] = self.sites
            if hasattr(self, 'ent_daily_site'):
                self.ent_daily_site['values'] = self.sites
            self.save_tab_config()
        
        messagebox.showinfo("완료", f"{pure_mat_name} {t_type} 처리되었습니다.")
        
        self.ent_qty.delete(0, tk.END)
        self.ent_note.delete(0, tk.END)
        if isinstance(self.ent_user, tk.Entry):
            self.ent_user.delete(0, tk.END)
        else:
            self.ent_user.set('')
        if hasattr(self, 'cb_trans_site'):
            self.cb_trans_site.set('')
        self.update_transaction_view()

    def update_transaction_view(self):
        """Populate the transaction history Treeview"""
        if not hasattr(self, 'inout_tree'):
            return
            
        # Clear current view
        for item in self.inout_tree.get_children():
            self.inout_tree.delete(item)
            
        if self.transactions_df.empty:
            return
            
        # Display last 100 transactions, descending by date
        df_sorted = self.transactions_df.sort_values(by='Date', ascending=False).head(100)
        
        for idx, row in df_sorted.iterrows():
            mat_id = row['Material ID']
            mat_name = self.get_material_display_name(mat_id)
            
            date_str = row['Date'].strftime('%Y-%m-%d %H:%M:%S') if isinstance(row['Date'], datetime.datetime) else str(row['Date'])
            
            self.inout_tree.insert('', tk.END, values=(
                date_str,
                row.get('Site', ''),
                mat_name,
                row['Type'],
                row['Quantity'],
                row.get('User', ''),
                row.get('Note', '')
            ), tags=(str(idx),))

    def delete_transaction_entry(self):
        """Delete selected transaction from history and refresh"""
        selection = self.inout_tree.selection()
        if not selection:
            messagebox.showwarning("선택 오류", "삭제할 기록을 선택해주세요.")
            return
            
        if not messagebox.askyesno("삭제 확인", "선택한 기록을 삭제하시겠습니까?\n(삭제 시 재고 계산에 즉시 반영됩니다.)"):
            return
            
        for item in selection:
            tags = self.inout_tree.item(item, 'tags')
            if tags:
                idx = int(tags[0])
                if idx in self.transactions_df.index:
                    self.transactions_df = self.transactions_df.drop(idx)
        
        self.save_data()
        self.update_transaction_view()
        self.update_stock_view()
        messagebox.showinfo("완료", "거래 기록이 삭제되었습니다.")

    def on_material_selected(self, event=None):
        """Update the model listbox based on selected material"""
        selection = self.cb_material.get()
        if not selection:
            return
            
        # Clear existing models
        self.list_models.delete(0, tk.END)
        
        # Extract pure material name
        mat_name = selection
        if " - " in mat_name:
            mat_name = mat_name.split(" - ")[0]
        if " (SN: " in mat_name:
            mat_name = mat_name.split(" (SN: ")[0]
        
        pure_mat_name = mat_name
        
        # Find unique models for this material in materials_df
        if not self.materials_df.empty:
            relevant_mats = self.materials_df[self.materials_df['품목명'] == pure_mat_name]
            if not relevant_mats.empty:
                unique_models = relevant_mats['모델명'].dropna().unique()
                unique_models = sorted([str(m).strip() for m in unique_models if str(m).strip()])
                
                for model in unique_models:
                    self.list_models.insert(tk.END, model)
                
                # If no models found, add a placeholder
                if not unique_models:
                    self.list_models.insert(tk.END, "(등록된 모델명 없음)")

    def setup_report_tab(self):
        """Setup the report tab with same structure as monthly aggregation"""
        # Display frame for aggregated data
        display_frame = ttk.LabelFrame(self.tab_reports, text="월별 사용량 보고서 (현장별 데이터 자동 집계)")
        display_frame.pack(expand=True, fill='both', padx=10, pady=10)
        
        # Filter controls
        filter_frame = ttk.Frame(display_frame)
        filter_frame.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(filter_frame, text="연도:").pack(side='left', padx=5)
        self.cb_report_filter_year = ttk.Combobox(filter_frame, values=['전체'] + [str(y) for y in range(2024, 2031)], width=10)
        self.cb_report_filter_year.pack(side='left', padx=5)
        self.cb_report_filter_year.set('전체')
        
        ttk.Label(filter_frame, text="월:").pack(side='left', padx=5)
        self.cb_report_filter_month = ttk.Combobox(filter_frame, values=['전체'] + [str(m) for m in range(1, 13)], width=10)
        self.cb_report_filter_month.pack(side='left', padx=5)
        self.cb_report_filter_month.set('전체')
        
        ttk.Label(filter_frame, text="현장:").pack(side='left', padx=5)
        self.cb_report_filter_site = ttk.Combobox(filter_frame, width=15)
        self.cb_report_filter_site.pack(side='left', padx=5)
        self.cb_report_filter_site.set('전체')
        
        ttk.Label(filter_frame, text="품목명:").pack(side='left', padx=5)
        self.cb_report_filter_material = ttk.Combobox(filter_frame, width=25)
        self.cb_report_filter_material.pack(side='left', padx=5)
        self.cb_report_filter_material.set('전체')
        
        btn_filter = ttk.Button(filter_frame, text="조회", command=self.update_report_view)
        btn_filter.pack(side='left', padx=10)
        
        btn_export = ttk.Button(filter_frame, text="엑셀 내보내기", command=self.export_report_to_excel)
        btn_export.pack(side='left', padx=5)
        
        # Treeview for report data
        tree_frame = ttk.Frame(display_frame)
        tree_frame.pack(expand=True, fill='both', padx=5, pady=5)
        
        # Scrollbars
        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")
        
        # Treeview with same columns as monthly usage tab
        columns = ('연도', '월', '현장', '품목명', '필름매수', '센터미스', '농도', '마킹미스', '필름마크', '취급부주의', '고객불만', '기타', 'RT총계', '자분', '페인트', '침투제', '세척제', '현상제', '형광', '비고')
        self.report_tree = ttk.Treeview(tree_frame, columns=columns, show='headings',
                                       yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        vsb.config(command=self.report_tree.yview)
        hsb.config(command=self.report_tree.xview)
        
        # Column configuration with same widths as monthly usage tab
        col_widths = [80, 60, 120, 200, 100, 70, 70, 70, 70, 70, 70, 70, 70, 80, 80, 80, 80, 80, 80, 200]
        # Columns to center-align (numeric values)
        center_cols = ['필름매수', '센터미스', '농도', '마킹미스', '필름마크', '취급부주의', '고객불만', '기타', 'RT총계', '자분', '페인트', '침투제', '세척제', '현상제', '형광']
        
        for col, width in zip(columns, col_widths):
            self.report_tree.heading(col, text=col)
            self.report_tree.column(col, width=width, anchor='center')
        
        # Grid layout
        self.report_tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        # Initial view update
        self.update_report_view()


    def update_report_view(self):
        """Update the report treeview with aggregated data from daily usage (same logic as monthly tab)"""
        # Clear current view
        for item in self.report_tree.get_children():
            self.report_tree.delete(item)
        
        # Get filter values
        filter_year = self.cb_report_filter_year.get() if hasattr(self, 'cb_report_filter_year') else '전체'
        filter_month = self.cb_report_filter_month.get() if hasattr(self, 'cb_report_filter_month') else '전체'
        filter_site = self.cb_report_filter_site.get() if hasattr(self, 'cb_report_filter_site') else '전체'
        filter_material = self.cb_report_filter_material.get() if hasattr(self, 'cb_report_filter_material') else '전체'
        
        # Return if daily usage data is empty
        if self.daily_usage_df.empty:
            return
        
        # Create a copy of daily usage data and extract year/month from Date column
        df = self.daily_usage_df.copy()
        df['Year'] = pd.to_datetime(df['Date']).dt.year
        df['Month'] = pd.to_datetime(df['Date']).dt.month
        
        # Apply filters
        if filter_year != '전체':
            df = df[df['Year'] == int(filter_year)]
        
        if filter_month != '전체':
            df = df[df['Month'] == int(filter_month)]
        
        if filter_site != '전체':
            df = df[df['Site'] == filter_site]
        
        # Populate site filter options from data
        if hasattr(self, 'cb_report_filter_site'):
            unique_sites = ['전체'] + sorted(self.daily_usage_df['Site'].dropna().unique().tolist())
            self.cb_report_filter_site['values'] = unique_sites
            if not self.cb_report_filter_site.get():
                self.cb_report_filter_site.set('전체')
        
        # Populate material filter options from data
        if hasattr(self, 'cb_report_filter_material'):
            # Get unique material names from materials_df based on Material IDs in daily_usage_df
            unique_mat_ids = self.daily_usage_df['Material ID'].dropna().unique()
            material_names = []
            for mat_id in unique_mat_ids:
                mat_row = self.materials_df[self.materials_df['Material ID'] == mat_id]
                if not mat_row.empty:
                    material_names.append(mat_row.iloc[0]['품목명'])
            unique_materials = ['전체'] + sorted(set(material_names))
            self.cb_report_filter_material['values'] = unique_materials
            if not self.cb_report_filter_material.get():
                self.cb_report_filter_material.set('전체')
        
        # Prepare aggregation dictionary for all numeric fields
        agg_dict = {'Usage': 'sum'}
        
        # Add Film Count if exists
        if 'Film Count' in df.columns:
            agg_dict['Film Count'] = 'sum'
        
        # Add RTK categories
        rtk_categories = ['RTK_센터미스', 'RTK_농도', 'RTK_마킹미스', 'RTK_필름마크', 'RTK_취급부주의', 'RTK_고객불만', 'RTK_기타']
        for cat in rtk_categories:
            if cat in df.columns:
                agg_dict[cat] = 'sum'
        
        # Add NDT materials
        ndt_materials = ['NDT_자분', 'NDT_페인트', 'NDT_침투제', 'NDT_세척제', 'NDT_현상제', 'NDT_형광']
        for mat in ndt_materials:
            if mat in df.columns:
                agg_dict[mat] = 'sum'
        
        # Group by Year, Month, Site, Material ID and aggregate
        grouped = df.groupby(['Year', 'Month', 'Site', 'Material ID']).agg(agg_dict).reset_index()
        
        # Initialize totals for cumulative sum
        total_film_count = 0.0
        total_rtk_center = 0.0
        total_rtk_density = 0.0
        total_rtk_marking = 0.0
        total_rtk_film = 0.0
        total_rtk_handling = 0.0
        total_rtk_customer = 0.0
        total_rtk_other = 0.0
        total_ndt_magnet = 0.0
        total_ndt_paint = 0.0
        total_ndt_penetrant = 0.0
        total_ndt_cleaner = 0.0
        total_ndt_developer = 0.0
        total_ndt_fluorescent = 0.0
        
        # Display aggregated entries
        for _, entry in grouped.iterrows():
            mat_id = entry['Material ID']
            mat_row = self.materials_df[self.materials_df['Material ID'] == mat_id]
            
            if not mat_row.empty:
                mat_name = mat_row.iloc[0]['품목명']
            else:
                mat_name = f"ID: {mat_id}"
            
            # Apply material filter
            if filter_material != '전체' and mat_name != filter_material:
                continue
            
            # Get aggregated values
            film_count = entry.get('Film Count', 0.0)
            
            # RTK values
            rtk_center = entry.get('RTK_센터미스', 0.0)
            rtk_density = entry.get('RTK_농도', 0.0)
            rtk_marking = entry.get('RTK_마킹미스', 0.0)
            rtk_film = entry.get('RTK_필름마크', 0.0)
            rtk_handling = entry.get('RTK_취급부주의', 0.0)
            rtk_customer = entry.get('RTK_고객불만', 0.0)
            rtk_other = entry.get('RTK_기타', 0.0)
            rtk_total = rtk_center + rtk_density + rtk_marking + rtk_film + rtk_handling + rtk_customer + rtk_other
            
            # NDT values
            ndt_magnet = entry.get('NDT_자분', 0.0)
            ndt_paint = entry.get('NDT_페인트', 0.0)
            ndt_penetrant = entry.get('NDT_침투제', 0.0)
            ndt_cleaner = entry.get('NDT_세척제', 0.0)
            ndt_developer = entry.get('NDT_현상제', 0.0)
            ndt_fluorescent = entry.get('NDT_형광', 0.0)
            
            # Accumulate totals
            total_film_count += film_count
            total_rtk_center += rtk_center
            total_rtk_density += rtk_density
            total_rtk_marking += rtk_marking
            total_rtk_film += rtk_film
            total_rtk_handling += rtk_handling
            total_rtk_customer += rtk_customer
            total_rtk_other += rtk_other
            total_ndt_magnet += ndt_magnet
            total_ndt_paint += ndt_paint
            total_ndt_penetrant += ndt_penetrant
            total_ndt_cleaner += ndt_cleaner
            total_ndt_developer += ndt_developer
            total_ndt_fluorescent += ndt_fluorescent
            
            self.report_tree.insert('', tk.END, values=(
                int(entry['Year']),
                int(entry['Month']),
                entry.get('Site', ''),
                mat_name,
                f"{film_count:.1f}" if film_count > 0 else '',
                f"{rtk_center:.1f}" if rtk_center > 0 else '',
                f"{rtk_density:.1f}" if rtk_density > 0 else '',
                f"{rtk_marking:.1f}" if rtk_marking > 0 else '',
                f"{rtk_film:.1f}" if rtk_film > 0 else '',
                f"{rtk_handling:.1f}" if rtk_handling > 0 else '',
                f"{rtk_customer:.1f}" if rtk_customer > 0 else '',
                f"{rtk_other:.1f}" if rtk_other > 0 else '',
                f"{rtk_total:.1f}" if rtk_total > 0 else '',
                f"{ndt_magnet:.1f}" if ndt_magnet > 0 else '',
                f"{ndt_paint:.1f}" if ndt_paint > 0 else '',
                f"{ndt_penetrant:.1f}" if ndt_penetrant > 0 else '',
                f"{ndt_cleaner:.1f}" if ndt_cleaner > 0 else '',
                f"{ndt_developer:.1f}" if ndt_developer > 0 else '',
                f"{ndt_fluorescent:.1f}" if ndt_fluorescent > 0 else '',
                ''  # Empty note field
            ))
        
        # Add total row at the bottom if there's data
        if not grouped.empty:
            total_rtk_sum = total_rtk_center + total_rtk_density + total_rtk_marking + total_rtk_film + total_rtk_handling + total_rtk_customer + total_rtk_other
            
            # Configure tag for total row
            self.report_tree.tag_configure('total', background='#E8F4F8', font=('Arial', 9, 'bold'))
            
            self.report_tree.insert('', tk.END, values=(
                '',
                '',
                '',
                '=== 누계 ===',
                f"{total_film_count:.1f}" if total_film_count > 0 else '',
                f"{total_rtk_center:.1f}" if total_rtk_center > 0 else '',
                f"{total_rtk_density:.1f}" if total_rtk_density > 0 else '',
                f"{total_rtk_marking:.1f}" if total_rtk_marking > 0 else '',
                f"{total_rtk_film:.1f}" if total_rtk_film > 0 else '',
                f"{total_rtk_handling:.1f}" if total_rtk_handling > 0 else '',
                f"{total_rtk_customer:.1f}" if total_rtk_customer > 0 else '',
                f"{total_rtk_other:.1f}" if total_rtk_other > 0 else '',
                f"{total_rtk_sum:.1f}" if total_rtk_sum > 0 else '',
                f"{total_ndt_magnet:.1f}" if total_ndt_magnet > 0 else '',
                f"{total_ndt_paint:.1f}" if total_ndt_paint > 0 else '',
                f"{total_ndt_penetrant:.1f}" if total_ndt_penetrant > 0 else '',
                f"{total_ndt_cleaner:.1f}" if total_ndt_cleaner > 0 else '',
                f"{total_ndt_developer:.1f}" if total_ndt_developer > 0 else '',
                f"{total_ndt_fluorescent:.1f}" if total_ndt_fluorescent > 0 else '',
                ''
            ), tags=('total',))
    
    def export_report_to_excel(self):
        """Export current report view to Excel file"""
        # Build data from current treeview
        report_data = []
        for item in self.report_tree.get_children():
            values = self.report_tree.item(item, 'values')
            # Skip the total row
            if values[3] == '=== 누계 ===':
                continue
            report_data.append({
                '연도': values[0],
                '월': values[1],
                '현장': values[2],
                '품목명': values[3],
                '필름매수': values[4],
                '센터미스': values[5],
                '농도': values[6],
                '마킹미스': values[7],
                '필름마크': values[8],
                '취급부주의': values[9],
                '고객불만': values[10],
                '기타': values[11],
                'RT총계': values[12],
                '자분': values[13],
                '페인트': values[14],
                '침투제': values[15],
                '세척제': values[16],
                '현상제': values[17],
                '형광': values[18],
                '비고': values[19]
            })
        
        if not report_data:
            messagebox.showinfo("알림", "내보낼 데이터가 없습니다.")
            return
        
        # Prepare filename based on filters
        filter_year = self.cb_report_filter_year.get()
        filter_month = self.cb_report_filter_month.get()
        filename_parts = ["월별사용량보고서"]
        if filter_year != '전체':
            filename_parts.append(f"{filter_year}년")
        if filter_month != '전체':
            filename_parts.append(f"{filter_month}월")
        filename = "_".join(filename_parts) + ".xlsx"
        
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=filename,
            title="보고서 저장",
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if save_path:
            try:
                report_df = pd.DataFrame(report_data)
                report_df.to_excel(save_path, index=False)
                messagebox.showinfo("완료", "보고서가 저장되었습니다.")
            except Exception as e:
                messagebox.showerror("오류", f"저장 실패: {e}")
    
    def setup_import_tab(self):
        import_frame = ttk.LabelFrame(self.tab_import, text="데이터 관리")
        import_frame.pack(pady=20, padx=20, fill='both', expand=True)
        
        # Import Section
        ttk.Label(import_frame, text="엑셀 파일에서 자재 데이터 가져오기", font=('Arial', 11, 'bold')).pack(pady=10)
        ttk.Label(import_frame, text="형식: Material ID, 회사코드, 관리품번, 품목명, 창고, 모델명, 규격, 품목군코드, 제조사, 제조국, 가격, 관리단위, 수량", 
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
            required_cols = ['품목명'] if '품목명' in imported_df.columns else (['품명'] if '품명' in imported_df.columns else ['Item Name'])
            if not any(col in imported_df.columns for col in ['품목명', '품명', 'Item Name']):
                messagebox.showerror("오류", "필수 컬럼 '품목명', '품명' 또는 'Item Name'이 없습니다.")
                return
            
            # Process each row
            count_new = 0
            count_updated = 0
            
            # Ask user if they want to update existing items or just append
            update_existing = messagebox.askyesno("가져오기 방식", 
                "기존에 등록된 품목(품목명, SN, 규격 일치)이 있을 경우 정보를 업데이트하시겠습니까?\n\n"
                "'예': 기존 정보 수정\n'아니오': 무시하고 새로 추가")
            
            for _, row in imported_df.iterrows():
                # Extract values with mapping
                mat_name = str(row.get('품목명', row.get('품명', row.get('Item Name', '')))).strip()
                if not mat_name or mat_name == 'nan': continue
                
                sn = str(row.get('SN', row.get('SN번호', row.get('Serial Number', '')))).strip()
                if sn == 'nan': sn = ''
                
                spec = str(row.get('규격', row.get('Specification', ''))).strip()
                if spec == 'nan': spec = ''
                
                # Check for existing
                existing_idx = -1
                if not self.materials_df.empty:
                    # Match by Name, SN and Specification
                    mask = (self.materials_df['품목명'].astype(str) == mat_name) & \
                           (self.materials_df['SN'].astype(str).replace('nan', '') == sn) & \
                           (self.materials_df['규격'].astype(str).replace('nan', '') == spec)
                    
                    matches = self.materials_df.index[mask].tolist()
                    if matches:
                        existing_idx = matches[0]
                
                # Prepare data row
                extracted_model, extracted_sn = self.extract_sn_from_model(
                    row.get('모델명', row.get('Model', '')), 
                    sn
                )
                
                data_row = {
                    '회사코드': row.get('회사코드', row.get('Company Code', '')),
                    '관리품번': row.get('관리품번', row.get('Equipment Code', '')),
                    '품목명': mat_name,
                    'SN': extracted_sn,
                    '창고': row.get('창고', row.get('Warehouse', '')),
                    '모델명': extracted_model,
                    '규격': spec,
                    '품목군코드': row.get('품목군코드', row.get('Classification', '')),
                    '공급업체': row.get('공급업체', row.get('공급업자', row.get('Supplier', ''))),
                    '제조사': row.get('제조사', row.get('Manufacturer', '')),
                    '제조국': row.get('제조국', row.get('Country', row.get('Origin', ''))),
                    '가격': row.get('가격', row.get('Price', 0)),
                    '관리단위': row.get('관리단위', row.get('Unit', 'EA')),
                    '수량': row.get('수량', row.get('Initial Stock', row.get('Current Stock', 0))),
                    '재고하한': row.get('재고하한', row.get('재주문 수준', row.get('Reorder Point', 0)))
                }
                
                # Handle NaN values in numeric fields
                for field in ['가격', '수량', '재고하한']:
                    if pd.isna(data_row[field]): data_row[field] = 0
                
                if existing_idx != -1:
                    if update_existing:
                        # Update existing
                        for col, val in data_row.items():
                            self.materials_df.at[existing_idx, col] = val
                        count_updated += 1
                    else:
                        # Skip or append (here we append if skipping update to maintain old behavior but with new ID)
                        new_mat_id = self.materials_df['Material ID'].max() + 1 if not self.materials_df.empty else 1
                        data_row['Material ID'] = new_mat_id
                        self.materials_df = pd.concat([self.materials_df, pd.DataFrame([data_row])], ignore_index=True)
                        count_new += 1
                else:
                    # Add as new
                    new_mat_id = self.materials_df['Material ID'].max() + 1 if not self.materials_df.empty else 1
                    data_row['Material ID'] = new_mat_id
                    self.materials_df = pd.concat([self.materials_df, pd.DataFrame([data_row])], ignore_index=True)
                    count_new += 1
            
            self.save_data()
            self.update_material_combo()
            self.update_stock_view()
            self.update_registration_combos()
            
            msg = f"자재 가져오기가 완료되었습니다.\n\n"
            if count_new > 0: msg += f"• 신규 등록: {count_new}건\n"
            if count_updated > 0: msg += f"• 기존 수정: {count_updated}건"
            messagebox.showinfo("완료", msg)
            
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
    
    def export_transaction_history(self):
        """Export transaction history displayed in the inout_tree to Excel"""
        # Build data from current treeview
        history_data = []
        for item in self.inout_tree.get_children():
            values = self.inout_tree.item(item, 'values')
            history_data.append({
                '날짜': values[0],
                '현장': values[1],
                '품목명': values[2],
                '구분': values[3],
                '수량': values[4],
                '담당자': values[5],
                '비고': values[6]
            })
        
        if not history_data:
            messagebox.showinfo("알림", "내보낼 데이터가 없습니다.")
            return
        
        # Prepare filename
        today = datetime.datetime.now().strftime('%Y%m%d')
        filename = f"입출고내역_{today}.xlsx"
        
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=filename,
            title="입출고 내역 저장",
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if save_path:
            try:
                history_df = pd.DataFrame(history_data)
                history_df.to_excel(save_path, index=False)
                messagebox.showinfo("완료", "입출고 내역이 저장되었습니다.")
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
    
    def export_daily_usage_history(self):
        """일일 사용량 기록을 이젝리로 내보내기"""
        # Build data from current treeview
        daily_data = []
        for item in self.daily_usage_tree.get_children():
            values = self.daily_usage_tree.item(item, 'values')
            daily_data.append({
                '날짜': values[0],
                '현장': values[1],
                '품목명': values[2],
                '필름매수': values[3],
                '센터미스': values[4],
                '농도': values[5],
                '마킹미스': values[6],
                '필름마크': values[7],
                '취급부주의': values[8],
                '고객불만': values[9],
                '기타': values[10],
                'RT총계': values[11],
                '자분': values[12],
                '페인트': values[13],
                '침투제': values[14],
                '세척제': values[15],
                '현상제': values[16],
                '형광': values[17],
                '비고': values[18],
                '입력시간': values[19]
            })
        
        if not daily_data:
            messagebox.showinfo("알림", "내보낼 데이터가 없습니다.")
            return
        
        # Prepare filename
        today = datetime.datetime.now().strftime('%Y%m%d')
        filename = f"일일사용량기록_{today}.xlsx"
        
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=filename,
            title="일일 사용량 기록 저장",
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if save_path:
            try:
                daily_df = pd.DataFrame(daily_data)
                daily_df.to_excel(save_path, index=False)
                messagebox.showinfo("완료", "일일 사용량 기록이 저장되었습니다.")
            except Exception as e:
                messagebox.showerror("오류", f"저장 실패: {e}")
    
    def export_all_daily_usage(self):
        """데이터베이스의 모든 일일 사용량 기록을 엑셀로 내보내기 (필터 무시)"""
        if self.daily_usage_df.empty:
            messagebox.showinfo("알림", "저장된 기록이 없습니다.")
            return
        
        # Create a copy of the dataframe with material names resolved
        export_df = self.daily_usage_df.copy()
        
        # Add material names column
        material_names = []
        for idx, row in export_df.iterrows():
            mat_id = row.get('Material ID')
            if pd.notna(mat_id):
                mat_row = self.materials_df[self.materials_df['Material ID'] == mat_id]
                if not mat_row.empty:
                    material_names.append(mat_row.iloc[0]['품목명'])
                else:
                    material_names.append(f"ID: {mat_id}")
            else:
                material_names.append('')
        
        export_df.insert(2, '품목명', material_names)
        
        # Reorder and select columns for export
        columns_to_export = ['날짜', 'Date', 'Site', '품목명', 'Film Count', 
                             'RTK_센터미스', 'RTK_농도', 'RTK_마킹미스', 'RTK_필름마크',
                             'RTK_취급부주의', 'RTK_고객불만', 'RTK_기타',
                             'NDT_자분', 'NDT_페인트', 'NDT_침투제', 'NDT_세척제', 'NDT_현상제', 'NDT_형광',
                             'Note', 'User', 'Entry Time']
        
        # Filter to only include existing columns
        available_columns = [col for col in columns_to_export if col in export_df.columns]
        export_df = export_df[available_columns]
        
        # Prepare filename
        today = datetime.datetime.now().strftime('%Y%m%d')
        filename = f"일일사용량_전체기록_{today}.xlsx"
        
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=filename,
            title="전체 일일 사용량 기록 저장",
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if save_path:
            try:
                export_df.to_excel(save_path, index=False)
                record_count = len(export_df)
                messagebox.showinfo("완료", f"전체 일일 사용량 기록이 저장되었습니다.\n(총 {record_count}건)")
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
                    '품목명': mat.get('품목명', ''),  # Using 재고현황's 품목명 field
                    '관리품번': mat.get('관리품번', ''),
                    '규격': mat.get('규격', ''),
                    '단위': mat.get('관리단위', 'EA'),
                    '월사용량': month_usage,
                    '누계사용량': cumulative_usage
                })
        
        # Sort by item name
        usage_data.sort(key=lambda x: x['품목명'])
        
        # Display in treeview
        for data in usage_data:
            self.usage_tree.insert('', tk.END, values=(
                data['품목명'],
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
                '자재명': mat.get('품목명', ''),  # Using 재고현황's 품목명 field
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
                '자재명': mat.get('품목명', ''),  # Using 재고현황's 품목명 field
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

    def setup_monthly_usage_tab(self):
        """Setup the monthly usage aggregation tab (auto-aggregated from daily usage)"""
        # Display frame for aggregated monthly data
        display_frame = ttk.LabelFrame(self.tab_monthly_usage, text="월별 사용량 집계 (현장별 데이터 자동 집계)")
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
        
        ttk.Label(filter_frame, text="현장:").pack(side='left', padx=5)
        self.cb_filter_site_monthly = ttk.Combobox(filter_frame, width=15)
        self.cb_filter_site_monthly.pack(side='left', padx=5)
        self.cb_filter_site_monthly.set('전체')
        
        ttk.Label(filter_frame, text="품목명:").pack(side='left', padx=5)
        self.cb_filter_material_monthly = ttk.Combobox(filter_frame, width=25)
        self.cb_filter_material_monthly.pack(side='left', padx=5)
        self.cb_filter_material_monthly.set('전체')
        
        btn_filter = ttk.Button(filter_frame, text="조회", command=self.update_monthly_usage_view)
        btn_filter.pack(side='left', padx=10)
        
        btn_export = ttk.Button(filter_frame, text="엑셀 내보내기", command=self.export_monthly_usage_history)
        btn_export.pack(side='left', padx=5)
        
        # Treeview for monthly usage records
        tree_frame = ttk.Frame(display_frame)
        tree_frame.pack(expand=True, fill='both', padx=5, pady=5)
        
        # Scrollbars
        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")
        
        # Treeview with same columns as daily usage tab
        columns = ('연도', '월', '현장', '품목명', '필름매수', '센터미스', '농도', '마킹미스', '필름마크', '취급부주의', '고객불만', '기타', 'RT총계', '자분', '페인트', '침투제', '세척제', '현상제', '형광', '비고')
        self.monthly_usage_tree = ttk.Treeview(tree_frame, columns=columns, show='headings',
                                               yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        vsb.config(command=self.monthly_usage_tree.yview)
        hsb.config(command=self.monthly_usage_tree.xview)
        
        # Column configuration with same widths as daily usage tab
        col_widths = [80, 60, 120, 200, 100, 70, 70, 70, 70, 70, 70, 70, 70, 80, 80, 80, 80, 80, 80, 200]
        # Columns to center-align (numeric values)
        center_cols = ['필름매수', '센터미스', '농도', '마킹미스', '필름마크', '취급부주의', '고객불만', '기타', 'RT총계', '자분', '페인트', '침투제', '세척제', '현상제', '형광']
        
        for col, width in zip(columns, col_widths):
            self.monthly_usage_tree.heading(col, text=col)
            self.monthly_usage_tree.column(col, width=width, anchor='center')
        
        # Grid layout
        self.monthly_usage_tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        # Initial view update
        self.update_monthly_usage_view()
    

    def update_monthly_usage_view(self):
        """Update the monthly usage treeview with aggregated data from daily usage"""
        # Clear current view
        for item in self.monthly_usage_tree.get_children():
            self.monthly_usage_tree.delete(item)
        
        # Get filter values
        filter_year = self.cb_filter_year.get()
        filter_month = self.cb_filter_month.get()
        filter_site = self.cb_filter_site_monthly.get() if hasattr(self, 'cb_filter_site_monthly') else '전체'
        filter_material = self.cb_filter_material_monthly.get() if hasattr(self, 'cb_filter_material_monthly') else '전체'
        
        # Return if daily usage data is empty
        if self.daily_usage_df.empty:
            return
        
        # Create a copy of daily usage data and extract year/month from Date column
        df = self.daily_usage_df.copy()
        df['Year'] = pd.to_datetime(df['Date']).dt.year
        df['Month'] = pd.to_datetime(df['Date']).dt.month
        
        # Apply filters
        if filter_year != '전체':
            df = df[df['Year'] == int(filter_year)]
        
        if filter_month != '전체':
            df = df[df['Month'] == int(filter_month)]
        
        if filter_site != '전체':
            df = df[df['Site'] == filter_site]
        
        # Populate site filter options from data
        if hasattr(self, 'cb_filter_site_monthly'):
            unique_sites = ['전체'] + sorted(self.daily_usage_df['Site'].dropna().unique().tolist())
            self.cb_filter_site_monthly['values'] = unique_sites
            if not self.cb_filter_site_monthly.get():
                self.cb_filter_site_monthly.set('전체')
        
        # Populate material filter options from data
        if hasattr(self, 'cb_filter_material_monthly'):
            # Get unique material names from materials_df based on Material IDs in daily_usage_df
            unique_mat_ids = self.daily_usage_df['Material ID'].dropna().unique()
            material_names = []
            for mat_id in unique_mat_ids:
                mat_row = self.materials_df[self.materials_df['Material ID'] == mat_id]
                if not mat_row.empty:
                    material_names.append(mat_row.iloc[0]['품목명'])
            unique_materials = ['전체'] + sorted(set(material_names))
            self.cb_filter_material_monthly['values'] = unique_materials
            if not self.cb_filter_material_monthly.get():
                self.cb_filter_material_monthly.set('전체')
        
        # Prepare aggregation dictionary for all numeric fields
        agg_dict = {'Usage': 'sum'}
        
        # Add Film Count if exists
        if 'Film Count' in df.columns:
            agg_dict['Film Count'] = 'sum'
        
        # Add RTK categories
        rtk_categories = ['RTK_센터미스', 'RTK_농도', 'RTK_마킹미스', 'RTK_필름마크', 'RTK_취급부주의', 'RTK_고객불만', 'RTK_기타']
        for cat in rtk_categories:
            if cat in df.columns:
                agg_dict[cat] = 'sum'
        
        # Add NDT materials
        ndt_materials = ['NDT_자분', 'NDT_페인트', 'NDT_침투제', 'NDT_세척제', 'NDT_현상제', 'NDT_형광']
        for mat in ndt_materials:
            if mat in df.columns:
                agg_dict[mat] = 'sum'
        
        # Group by Year, Month, Site, Material ID and aggregate
        grouped = df.groupby(['Year', 'Month', 'Site', 'Material ID']).agg(agg_dict).reset_index()
        
        # Initialize totals for cumulative sum
        total_film_count = 0.0
        total_rtk_center = 0.0
        total_rtk_density = 0.0
        total_rtk_marking = 0.0
        total_rtk_film = 0.0
        total_rtk_handling = 0.0
        total_rtk_customer = 0.0
        total_rtk_other = 0.0
        total_ndt_magnet = 0.0
        total_ndt_paint = 0.0
        total_ndt_penetrant = 0.0
        total_ndt_cleaner = 0.0
        total_ndt_developer = 0.0
        total_ndt_fluorescent = 0.0
        
        # Display aggregated entries
        for _, entry in grouped.iterrows():
            mat_id = entry['Material ID']
            mat_row = self.materials_df[self.materials_df['Material ID'] == mat_id]
            
            if not mat_row.empty:
                mat_name = mat_row.iloc[0]['품목명']
            else:
                mat_name = f"ID: {mat_id}"
            
            # Apply material filter
            if filter_material != '전체' and mat_name != filter_material:
                continue
            
            # Get aggregated values
            film_count = entry.get('Film Count', 0.0)
            
            # RTK values
            rtk_center = entry.get('RTK_센터미스', 0.0)
            rtk_density = entry.get('RTK_농도', 0.0)
            rtk_marking = entry.get('RTK_마킹미스', 0.0)
            rtk_film = entry.get('RTK_필름마크', 0.0)
            rtk_handling = entry.get('RTK_취급부주의', 0.0)
            rtk_customer = entry.get('RTK_고객불만', 0.0)
            rtk_other = entry.get('RTK_기타', 0.0)
            rtk_total = rtk_center + rtk_density + rtk_marking + rtk_film + rtk_handling + rtk_customer + rtk_other
            
            # NDT values
            ndt_magnet = entry.get('NDT_자분', 0.0)
            ndt_paint = entry.get('NDT_페인트', 0.0)
            ndt_penetrant = entry.get('NDT_침투제', 0.0)
            ndt_cleaner = entry.get('NDT_세척제', 0.0)
            ndt_developer = entry.get('NDT_현상제', 0.0)
            ndt_fluorescent = entry.get('NDT_형광', 0.0)
            
            # Accumulate totals
            total_film_count += film_count
            total_rtk_center += rtk_center
            total_rtk_density += rtk_density
            total_rtk_marking += rtk_marking
            total_rtk_film += rtk_film
            total_rtk_handling += rtk_handling
            total_rtk_customer += rtk_customer
            total_rtk_other += rtk_other
            total_ndt_magnet += ndt_magnet
            total_ndt_paint += ndt_paint
            total_ndt_penetrant += ndt_penetrant
            total_ndt_cleaner += ndt_cleaner
            total_ndt_developer += ndt_developer
            total_ndt_fluorescent += ndt_fluorescent
            
            self.monthly_usage_tree.insert('', tk.END, values=(
                int(entry['Year']),
                int(entry['Month']),
                entry.get('Site', ''),
                mat_name,
                f"{film_count:.1f}" if film_count > 0 else '',
                f"{rtk_center:.1f}" if rtk_center > 0 else '',
                f"{rtk_density:.1f}" if rtk_density > 0 else '',
                f"{rtk_marking:.1f}" if rtk_marking > 0 else '',
                f"{rtk_film:.1f}" if rtk_film > 0 else '',
                f"{rtk_handling:.1f}" if rtk_handling > 0 else '',
                f"{rtk_customer:.1f}" if rtk_customer > 0 else '',
                f"{rtk_other:.1f}" if rtk_other > 0 else '',
                f"{rtk_total:.1f}" if rtk_total > 0 else '',
                f"{ndt_magnet:.1f}" if ndt_magnet > 0 else '',
                f"{ndt_paint:.1f}" if ndt_paint > 0 else '',
                f"{ndt_penetrant:.1f}" if ndt_penetrant > 0 else '',
                f"{ndt_cleaner:.1f}" if ndt_cleaner > 0 else '',
                f"{ndt_developer:.1f}" if ndt_developer > 0 else '',
                f"{ndt_fluorescent:.1f}" if ndt_fluorescent > 0 else '',
                ''  # Empty note field
            ))
        
        # Add total row at the bottom if there's data
        if not grouped.empty:
            total_rtk_sum = total_rtk_center + total_rtk_density + total_rtk_marking + total_rtk_film + total_rtk_handling + total_rtk_customer + total_rtk_other
            
            # Configure tag for total row
            self.monthly_usage_tree.tag_configure('total', background='#E8F4F8', font=('Arial', 9, 'bold'))
            
            self.monthly_usage_tree.insert('', tk.END, values=(
                '',
                '',
                '',
                '=== 누계 ===',
                f"{total_film_count:.1f}" if total_film_count > 0 else '',
                f"{total_rtk_center:.1f}" if total_rtk_center > 0 else '',
                f"{total_rtk_density:.1f}" if total_rtk_density > 0 else '',
                f"{total_rtk_marking:.1f}" if total_rtk_marking > 0 else '',
                f"{total_rtk_film:.1f}" if total_rtk_film > 0 else '',
                f"{total_rtk_handling:.1f}" if total_rtk_handling > 0 else '',
                f"{total_rtk_customer:.1f}" if total_rtk_customer > 0 else '',
                f"{total_rtk_other:.1f}" if total_rtk_other > 0 else '',
                f"{total_rtk_sum:.1f}" if total_rtk_sum > 0 else '',
                f"{total_ndt_magnet:.1f}" if total_ndt_magnet > 0 else '',
                f"{total_ndt_paint:.1f}" if total_ndt_paint > 0 else '',
                f"{total_ndt_penetrant:.1f}" if total_ndt_penetrant > 0 else '',
                f"{total_ndt_cleaner:.1f}" if total_ndt_cleaner > 0 else '',
                f"{total_ndt_developer:.1f}" if total_ndt_developer > 0 else '',
                f"{total_ndt_fluorescent:.1f}" if total_ndt_fluorescent > 0 else '',
                ''
            ), tags=('total',))

    def export_monthly_usage_history(self):
        """Export monthly usage data displayed in the monthly_usage_tree to Excel"""
        # Build data from current treeview
        monthly_data = []
        for item in self.monthly_usage_tree.get_children():
            values = self.monthly_usage_tree.item(item, 'values')
            monthly_data.append({
                '연도': values[0],
                '월': values[1],
                '현장': values[2],
                '품목명': values[3],
                '필름매수': values[4],
                '센터미스': values[5],
                '농도': values[6],
                '마킹미스': values[7],
                '필름마크': values[8],
                '취급부주의': values[9],
                '고객불만': values[10],
                '기타': values[11],
                'RT총계': values[12],
                '자분': values[13],
                '페인트': values[14],
                '침투제': values[15],
                '세척제': values[16],
                '현상제': values[17],
                '형광': values[18],
                '비고': values[19]
            })
        
        if not monthly_data:
            messagebox.showinfo("알림", "내보낼 데이터가 없습니다.")
            return
        
        # Prepare filename
        today = datetime.datetime.now().strftime('%Y%m%d')
        filename = f"월별집계_{today}.xlsx"
        
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=filename,
            title="월별 집계 내역 저장",
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if save_path:
            try:
                monthly_df = pd.DataFrame(monthly_data)
                monthly_df.to_excel(save_path, index=False)
                messagebox.showinfo("완료", "월별 집계 내역이 저장되었습니다.")
            except Exception as e:
                messagebox.showerror("오류", f"저장 실패: {e}")

    def setup_daily_usage_tab(self):
        """Setup the daily usage entry tab"""
        # Top frame for entry form
        entry_frame = ttk.LabelFrame(self.tab_daily_usage, text="현장별 일일 사용량 기입")
        entry_frame.pack(fill='x', padx=10, pady=10)
        
        # Date selection
        ttk.Label(entry_frame, text="날짜:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.ent_daily_date = DateEntry(entry_frame, width=28, background='darkblue',
                                        foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
        self.ent_daily_date.grid(row=0, column=1, padx=5, pady=5)
        # Set default to today (DateEntry handles this by default, but we can be explicit)
        self.ent_daily_date.set_date(datetime.datetime.now())
        ttk.Label(entry_frame, text="(마우스로 선택 가능)", font=('Arial', 8)).grid(row=0, column=2, padx=5, pady=5, sticky='w')
        
        # Site selection
        ttk.Label(entry_frame, text="현장:").grid(row=1, column=0, padx=5, pady=5, sticky='w')
        site_frame = ttk.Frame(entry_frame)
        site_frame.grid(row=1, column=1, padx=5, pady=5, sticky='w')
        
        self.ent_daily_site = ttk.Combobox(site_frame, width=37, values=self.sites)
        self.ent_daily_site.pack(side='left', padx=(0, 5))
        # Auto-save new site entries
        self.ent_daily_site.bind('<FocusOut>', lambda e: self.auto_save_to_list(e, self.ent_daily_site, self.sites, 'sites'))
        self.ent_daily_site.bind('<Return>', lambda e: self.auto_save_to_list(e, self.ent_daily_site, self.sites, 'sites'))
        
        btn_del_site = ttk.Button(site_frame, text="현장 삭제", command=self.delete_selected_site, width=10)
        btn_del_site.pack(side='left')
        
        # Name/User field
        ttk.Label(entry_frame, text="담당자:").grid(row=1, column=2, padx=(20, 5), pady=5, sticky='w')
        self.cb_daily_user = ttk.Combobox(entry_frame, width=20, values=getattr(self, 'users', []))
        self.cb_daily_user.grid(row=1, column=3, padx=5, pady=5, sticky='w')
        # Auto-save new manager entries
        self.cb_daily_user.bind('<FocusOut>', lambda e: self.auto_save_to_list(e, self.cb_daily_user, self.users, 'users'))
        self.cb_daily_user.bind('<Return>', lambda e: self.auto_save_to_list(e, self.cb_daily_user, self.users, 'users'))
        
        # Material selection
        ttk.Label(entry_frame, text="품목명:").grid(row=2, column=0, padx=5, pady=5, sticky='w')
        self.cb_daily_material = ttk.Combobox(entry_frame, state="readonly", width=65)
        self.cb_daily_material.grid(row=2, column=1, padx=5, pady=5)
        
        # Film count (필름매수)
        ttk.Label(entry_frame, text="필름매수:").grid(row=3, column=0, padx=5, pady=5, sticky='w')
        self.ent_film_count = ttk.Entry(entry_frame, width=20)
        self.ent_film_count.grid(row=3, column=1, padx=5, pady=5, sticky='w')
        
        # RT 매수 (RTK Usage amounts - separate for each category)
        rtk_frame = ttk.LabelFrame(entry_frame, text="RT 매수 (RTK별 사용량)")
        rtk_frame.grid(row=4, column=0, columnspan=3, padx=5, pady=5, sticky='ew')
        
        # Create entry fields for each RTK category
        self.rtk_entries = {}
        rtk_categories = ["센터미스", "농도", "마킹미스", "필름마크", "취급부주의", "고객불만", "기타", "총계"]
        
        for i, category in enumerate(rtk_categories):
            ttk.Label(rtk_frame, text=f"{category}:").grid(row=0, column=i*2, padx=5, pady=2, sticky='w')
            entry = ttk.Entry(rtk_frame, width=10)
            entry.grid(row=0, column=i*2+1, padx=5, pady=2)
            self.rtk_entries[category] = entry
            
            # Bind change event to calculate total
            if category != "총계":
                entry.bind('<KeyRelease>', lambda e, cat=category: self.calculate_rtk_total())
                entry.bind('<FocusOut>', lambda e, cat=category: self.calculate_rtk_total())
        
        # Make total entry readonly
        self.rtk_entries["총계"].config(state='readonly', background='lightgray')
        
        # NDT Materials Usage Frame
        ndt_frame = ttk.LabelFrame(entry_frame, text="NDT 자재 사용량")
        ndt_frame.grid(row=5, column=0, columnspan=3, padx=5, pady=5, sticky='ew')
        
        # Create entry fields for NDT materials
        self.ndt_entries = {}
        ndt_materials = ["자분", "페인트", "침투제", "세척제", "현상제", "형광"]
        
        for i, material in enumerate(ndt_materials):
            ttk.Label(ndt_frame, text=f"{material}:").grid(row=0, column=i*2, padx=5, pady=5, sticky='w')
            entry = ttk.Entry(ndt_frame, width=15)
            entry.grid(row=0, column=i*2+1, padx=5, pady=5)
            self.ndt_entries[material] = entry
        
        # Usage amount (hidden, for backward compatibility)
        self.ent_daily_usage = ttk.Entry(entry_frame, width=40)
        self.ent_daily_usage.grid_forget()  # Hide the old usage field
        
        # Note
        ttk.Label(entry_frame, text="비고:").grid(row=6, column=0, padx=5, pady=5, sticky='w')
        self.ent_daily_note = ttk.Entry(entry_frame, width=100)
        self.ent_daily_note.grid(row=6, column=1, columnspan=2, padx=5, pady=5, sticky='ew')
        
        # Save button
        btn_save_daily = ttk.Button(entry_frame, text="기록 저장", command=self.add_daily_usage_entry)
        btn_save_daily.grid(row=7, column=0, columnspan=3, pady=10)
        
        # Update material combobox
        self.update_material_combo()
        
        # Add Carestream films to daily usage material combobox
        if hasattr(self, 'cb_daily_material'):
            carestream_films = [
                "Carestream AA400-3⅓*12\"",
                "Carestream AA400-3⅓*17\"", 
                "Carestream AA400-4½*12\"",
                "Carestream AA400-10*12\"",
                "Carestream M100-3⅓*12\"",
                "Carestream M100-10*12\"",
                "Carestream M100-14*17\"",
                "Carestream MX125-3⅓*6\"",
                "Carestream MX125-3⅓*12\"",
                "Carestream MX125-4½*12\"",
                "Carestream MX125-10*12\"",
                "Carestream MX125-14*17\"",
                "Carestream T200-3⅓*12\"",
                "Carestream T200-3⅓*17\"",
                "Carestream T200-4½*12\"",
                "Carestream T200-10*12\"",
                "Carestream T200-14*17\""
            ]
            
            # Get existing values and add Carestream films
            existing_vals = list(self.cb_daily_material['values']) if self.cb_daily_material['values'] else []
            # Convert all to string, remove duplicates and sort
            all_vals = [str(mat) for mat in existing_vals + carestream_films if pd.notna(mat) and str(mat).strip()]
            all_vals = list(set(all_vals))
            all_vals.sort()
            self.cb_daily_material['values'] = all_vals
        
        # Bottom frame for display
        display_frame = ttk.LabelFrame(self.tab_daily_usage, text="일일 사용량 기록 조회")
        display_frame.pack(expand=True, fill='both', padx=10, pady=10)
        
        # Filter controls
        filter_frame = ttk.Frame(display_frame)
        filter_frame.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(filter_frame, text="시작일:").pack(side='left', padx=5)
        self.ent_daily_start_date = DateEntry(filter_frame, width=12, background='darkblue',
                                              foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
        self.ent_daily_start_date.pack(side='left', padx=5)
        # Default to 7 days ago
        start_date = (datetime.datetime.now() - datetime.timedelta(days=7))
        self.ent_daily_start_date.set_date(start_date)
        
        ttk.Label(filter_frame, text="종료일:").pack(side='left', padx=5)
        self.ent_daily_end_date = DateEntry(filter_frame, width=12, background='darkblue',
                                            foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
        self.ent_daily_end_date.pack(side='left', padx=5)
        self.ent_daily_end_date.set_date(datetime.datetime.now())
        
        ttk.Label(filter_frame, text="현장:").pack(side='left', padx=5)
        self.cb_daily_filter_site = ttk.Combobox(filter_frame, width=15, state="readonly")
        self.cb_daily_filter_site.pack(side='left', padx=5)
        self.cb_daily_filter_site.set('전체')
        
        ttk.Label(filter_frame, text="품목명:").pack(side='left', padx=5)
        self.cb_daily_filter_material = ttk.Combobox(filter_frame, width=25, state="readonly")
        self.cb_daily_filter_material.pack(side='left', padx=5)
        self.cb_daily_filter_material.set('전체')
        
        btn_filter = ttk.Button(filter_frame, text="조회", command=self.update_daily_usage_view)
        btn_filter.pack(side='left', padx=10)
        
        btn_delete = ttk.Button(filter_frame, text="선택 항목 삭제", command=self.delete_daily_usage_entry)
        btn_delete.pack(side='left', padx=10)
        
        btn_export = ttk.Button(filter_frame, text="엑셀 내보내기", command=self.export_daily_usage_history)
        btn_export.pack(side='left', padx=5)
        
        btn_export_all = ttk.Button(filter_frame, text="전체 기록 내보내기", command=self.export_all_daily_usage)
        btn_export_all.pack(side='left', padx=5)
        
        # Treeview for daily usage records
        tree_frame = ttk.Frame(display_frame)
        tree_frame.pack(expand=True, fill='both', padx=5, pady=5)
        
        # Scrollbars
        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")
        
        # Treeview with RTK categories and NDT materials
        columns = ('날짜', '현장', '품목명', '필름매수', '센터미스', '농도', '마킹미스', '필름마크', '취급부주의', '고객불만', '기타', 'RT총계', '자분', '페인트', '침투제', '세척제', '현상제', '형광', '비고', '입력시간')
        self.daily_usage_tree = ttk.Treeview(tree_frame, columns=columns, show='headings',
                                              yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        vsb.config(command=self.daily_usage_tree.yview)
        hsb.config(command=self.daily_usage_tree.xview)
        
        # Column configuration
        col_widths = [100, 120, 200, 100, 70, 70, 70, 70, 70, 70, 70, 70, 80, 80, 80, 80, 80, 80, 200, 150]
        # Columns to center-align (numeric values)
        center_cols = ['필름매수', '센터미스', '농도', '마킹미스', '필름마크', '취급부주의', '고객불만', '기타', 'RT총계', '자분', '페인트', '침투제', '세척제', '현상제', '형광']
        
        for col, width in zip(columns, col_widths):
            self.daily_usage_tree.heading(col, text=col)
            self.daily_usage_tree.column(col, width=width, anchor='center')
        
        # Grid layout
        self.daily_usage_tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        # Initial view update
        self.update_daily_usage_view()
    
    def delete_selected_site(self):
        """Remove the selected site from the suggestion list"""
        site = self.ent_daily_site.get().strip()
        if not site:
            messagebox.showwarning("선택 오류", "삭제할 현장명을 선택해주세요.")
            return
            
        if site in self.sites:
            if messagebox.askyesno("삭제 확인", f"'{site}' 현장명을 목록에서 삭제하시겠습니까?\n(기존 기록은 삭제되지 않습니다.)"):
                self.sites.remove(site)
                self.ent_daily_site['values'] = self.sites
                self.ent_daily_site.set('')
                self.save_tab_config()
                messagebox.showinfo("완료", "현장명이 삭제되었습니다.")
        else:
            messagebox.showinfo("알림", "목록에 없는 현장명입니다.")

    def calculate_rtk_total(self):
        """Calculate total RTK usage"""
        try:
            total = 0
            rtk_categories = ["센터미스", "농도", "마킹미스", "필름마크", "취급부주의", "고객불만", "기타"]
            for category in rtk_categories:
                value = self.rtk_entries[category].get()
                if value:
                    total += float(value)
            self.rtk_entries["총계"].config(state='normal')
            self.rtk_entries["총계"].delete(0, tk.END)
            self.rtk_entries["총계"].insert(0, str(total))
            self.rtk_entries["총계"].config(state='readonly')
        except ValueError:
            pass  # Ignore invalid input during typing
    
    def add_daily_usage_entry(self):
        """Add a daily usage entry"""
        date_str = self.ent_daily_date.get()
        site = self.ent_daily_site.get()
        mat_name = self.cb_daily_material.get()
        film_count_str = self.ent_film_count.get()
        note = self.ent_daily_note.get()
        
        # Validate film count
        try:
            film_count = float(film_count_str) if film_count_str else 0.0
        except ValueError:
            messagebox.showwarning("입력 오류", "필름매수는 숫자여야 합니다.")
            return
        
        # Get RTK usage values
        rtk_values = {}
        total_usage = 0
        rtk_categories = ["센터미스", "농도", "마킹미스", "필름마크", "취급부주의", "고객불만", "기타"]
        
        for category in rtk_categories:
            value_str = self.rtk_entries[category].get()
            try:
                value = float(value_str) if value_str else 0.0
                rtk_values[category] = value
                total_usage += value
            except ValueError:
                messagebox.showwarning("입력 오류", f"{category} 사용량은 숫자여야 합니다.")
                return
        
        # Validation
        if not date_str:
            messagebox.showwarning("입력 오류", "날짜를 입력해주세요.")
            return
        
        if not site:
            messagebox.showwarning("입력 오류", "현장을 입력해주세요.")
            return
        
        if not mat_name:
            messagebox.showwarning("입력 오류", "품목명을 선택해주세요.")
            return

        # Find pure material name
        mat_selection = mat_name
        if " - " in mat_selection:
            mat_selection = mat_selection.split(" - ")[0]
        if " (SN: " in mat_selection:
            mat_selection = mat_selection.split(" (SN: ")[0]
        
        pure_mat_name = mat_selection
        
        # Parse date
        try:
            usage_date = datetime.datetime.strptime(date_str, '%Y-%m-%d')
        except ValueError:
            messagebox.showwarning("입력 오류", "날짜 형식이 올바르지 않습니다. (YYYY-MM-DD)")
            return
        
        # Get material ID using 재고현황's 품목명 field
        mat_rows = self.materials_df[self.materials_df['품목명'] == pure_mat_name]
        
        if mat_rows.empty:
            # Material doesn't exist in database, create it
            new_mat_id = self.materials_df['Material ID'].max() + 1 if not self.materials_df.empty else 1
            
            new_material = {
                'Material ID': new_mat_id,
                '품목명': pure_mat_name,
                '관리품번': '',
                '품목군코드': 'FILM' if 'Carestream' in pure_mat_name else 'OTHER',
                '규격': '',
                '관리단위': 'EA',
                '재고위치': '',
                '공급업체': 'Carestream' if 'Carestream' in pure_mat_name else '',
                '제조사': 'Carestream' if 'Carestream' in pure_mat_name else '',
                '재고하한': 10,
                '수량': 0
            }
            
            self.materials_df = pd.concat([self.materials_df, pd.DataFrame([new_material])], ignore_index=True)
            self.save_data()
            mat_id = new_mat_id
        else:
            mat_id = mat_rows['Material ID'].values[0]
        
        # Create new entry with RTK values
        new_entry = {
            'Date': usage_date,
            'Site': site,
            'Material ID': mat_id,
            'Film Count': film_count,
            'Usage': total_usage,
            'Note': note,
            'Entry Time': datetime.datetime.now(),
            'User': self.cb_daily_user.get() if hasattr(self, 'cb_daily_user') else ''
        }
        
        # Add RTK values to the entry
        for category, value in rtk_values.items():
            new_entry[f'RTK_{category}'] = value
        
        # Add NDT materials values to the entry
        ndt_materials = ["자분", "페인트", "침투제", "세척제", "현상제", "형광"]
        ndt_usage = {}
        for material in ndt_materials:
            value_str = self.ndt_entries[material].get()
            try:
                value = float(value_str) if value_str else 0.0
                new_entry[f'NDT_{material}'] = value
                if value > 0:
                    ndt_usage[material] = value
            except ValueError:
                new_entry[f'NDT_{material}'] = 0.0
        
        # Stock deduction via transaction
        stock_info = ""
        
        # 1. Deduct main material (max of film_count or total_usage)
        main_deduction = max(film_count, total_usage)
        
        if main_deduction > 0:
            new_transaction = {
                'Date': usage_date,
                'Material ID': mat_id,
                'Type': 'OUT',
                'Quantity': main_deduction,
                'User': '현장사용',
                'Note': f'{site} 현장 사용 (자동 차감)'
            }
            self.transactions_df = pd.concat([self.transactions_df, pd.DataFrame([new_transaction])], ignore_index=True)
            stock_info += f"\n• {pure_mat_name}: {main_deduction:.1f} 차감"

        # 2. Deduct NDT materials
        for ndt_name, qty in ndt_usage.items():
            # Find NDT material in database
            ndt_mat_rows = self.materials_df[self.materials_df['품목명'] == ndt_name]
            if not ndt_mat_rows.empty:
                ndt_mat_id = ndt_mat_rows['Material ID'].values[0]
                ndt_transaction = {
                    'Date': usage_date,
                    'Material ID': ndt_mat_id,
                    'Type': 'OUT',
                    'Quantity': qty,
                    'User': '현장사용',
                    'Note': f'{site} 현장 사용 (자동 차감)'
                }
                self.transactions_df = pd.concat([self.transactions_df, pd.DataFrame([ndt_transaction])], ignore_index=True)
                stock_info += f"\n• {ndt_name}: {qty:.1f} 차감"
            else:
                stock_info += f"\n• {ndt_name}: 재고 정보 없음 (차감 실패)"
        
        self.daily_usage_df = pd.concat([self.daily_usage_df, pd.DataFrame([new_entry])], ignore_index=True)
        self.save_data()
        self.update_daily_usage_view()
        
        # Update site list if it's a new site
        if site not in self.sites:
            self.sites.append(site)
            self.sites.sort()
            self.ent_daily_site['values'] = self.sites
            self.save_tab_config()
            stock_info += f"\n• 신규 현장 '{site}'이 목록에 저장되었습니다."
        
        # Update user list if it's a new user
        user_name = self.cb_daily_user.get() if hasattr(self, 'cb_daily_user') else ''
        if user_name and user_name not in self.users:
            self.users.append(user_name)
            self.users.sort()
            if hasattr(self, 'cb_daily_user'):
                self.cb_daily_user['values'] = self.users
            self.save_tab_config()
            stock_info += f"\n• 신규 담당자 '{user_name}'이 목록에 저장되었습니다."

        # Update stock view and material lists
        if hasattr(self, 'update_stock_view'):
            self.update_stock_view()
        self.update_material_combo()
        self.update_registration_combos()
        
        messagebox.showinfo("완료", f"{pure_mat_name} 및 사용 자재 기록이 저장되었습니다.{stock_info}")
        
        # Clear entry fields
        self.ent_daily_site.set('') # Clear site combobox
        if hasattr(self, 'cb_daily_user'):
            self.cb_daily_user.set('') # Clear user combobox
        self.ent_film_count.delete(0, tk.END)
        self.ent_daily_note.delete(0, tk.END)
        for category in rtk_categories:
            self.rtk_entries[category].delete(0, tk.END)
        for material in ndt_materials:
            self.ndt_entries[material].delete(0, tk.END)
        self.rtk_entries["총계"].config(state='normal')
        self.rtk_entries["총계"].delete(0, tk.END)
        self.rtk_entries["총계"].config(state='readonly')
        # Keep date and reset to today
        self.ent_daily_date.set_date(datetime.datetime.now())
    
    def update_daily_usage_view(self):
        """Update the daily usage treeview"""
        # Clear current view
        for item in self.daily_usage_tree.get_children():
            self.daily_usage_tree.delete(item)
        
        # Get filter values
        start_date_str = self.ent_daily_start_date.get()
        end_date_str = self.ent_daily_end_date.get()
        filter_site = self.cb_daily_filter_site.get() if hasattr(self, 'cb_daily_filter_site') else '전체'
        filter_material = self.cb_daily_filter_material.get() if hasattr(self, 'cb_daily_filter_material') else '전체'
        
        # Populate site filter options from data
        if hasattr(self, 'cb_daily_filter_site'):
            unique_sites = ['전체'] + sorted(self.daily_usage_df['Site'].dropna().unique().tolist())
            self.cb_daily_filter_site['values'] = unique_sites
            if not self.cb_daily_filter_site.get():
                self.cb_daily_filter_site.set('전체')
        
        # Populate material filter options from data
        if hasattr(self, 'cb_daily_filter_material'):
            # Use unified display names for filter options
            unique_mat_ids = self.daily_usage_df['Material ID'].dropna().unique()
            material_names = []
            for mat_id in unique_mat_ids:
                material_names.append(self.get_material_display_name(mat_id))
            unique_materials = ['전체'] + sorted(set(material_names))
            self.cb_daily_filter_material['values'] = unique_materials
            if not self.cb_daily_filter_material.get():
                self.cb_daily_filter_material.set('전체')
        
        # Parse dates
        try:
            start_date = datetime.datetime.strptime(start_date_str, '%Y-%m-%d') if start_date_str else None
            end_date = datetime.datetime.strptime(end_date_str, '%Y-%m-%d') if end_date_str else None
            if end_date:
                # Include the entire end date
                end_date = end_date + datetime.timedelta(days=1) - datetime.timedelta(seconds=1)
        except ValueError:
            messagebox.showwarning("입력 오류", "날짜 형식이 올바르지 않습니다. (YYYY-MM-DD)")
            return
        
        # Filter data
        filtered_df = self.daily_usage_df.copy()
        
        if start_date is not None:
            filtered_df = filtered_df[pd.to_datetime(filtered_df['Date']) >= start_date]
        
        if end_date is not None:
            filtered_df = filtered_df[pd.to_datetime(filtered_df['Date']) <= end_date]
        
        if filter_site != '전체':
            filtered_df = filtered_df[filtered_df['Site'] == filter_site]
        
        # Filter by material if specified
        if filter_material != '전체':
            # Check all materials for matching full display name
            matching_mat_ids = []
            if not self.materials_df.empty:
                for _, mat in self.materials_df.iterrows():
                    if self.get_material_display_name(mat['Material ID']) == filter_material:
                        matching_mat_ids.append(mat['Material ID'])
            
            if matching_mat_ids:
                filtered_df = filtered_df[filtered_df['Material ID'].isin(matching_mat_ids)]
            else:
                # If no direct match in materials_df, try basic 품목명 (for edge cases)
                matching_mat_ids = self.materials_df[self.materials_df['품목명'] == filter_material]['Material ID'].tolist()
                if matching_mat_ids:
                    filtered_df = filtered_df[filtered_df['Material ID'].isin(matching_mat_ids)]
        
        # Sort by date descending
        if not filtered_df.empty:
            filtered_df = filtered_df.sort_values('Date', ascending=False)
        
        # Display entries and calculate totals
        total_film_count = 0
        total_rtk = [0.0] * 8  # 7 categories + 1 total
        total_ndt = [0.0] * 6  # 6 materials
        
        current_date = None
        sub_film_count = 0
        sub_rtk = [0.0] * 8
        sub_ndt = [0.0] * 6
        
        # Helper to insert subtotal row
        def insert_subtotal(date):
            if date is None: return
            self.daily_usage_tree.tag_configure('subtotal', background='#FFFDE7', font=('Arial', 9, 'italic'))
            sub_values = [
                f"[{date}] 일일 소계",
                '',
                '',
                f"{sub_film_count:.1f}",
                *[f"{v:.1f}" for v in sub_rtk],
                *[f"{v:.1f}" for v in sub_ndt],
                '',
                ''
            ]
            self.daily_usage_tree.insert('', tk.END, values=sub_values, tags=('subtotal',))

        for idx, entry in filtered_df.iterrows():
            usage_date = entry.get('Date', '')
            if pd.notna(usage_date):
                usage_date = pd.to_datetime(usage_date).strftime('%Y-%m-%d')
            else:
                usage_date = "Unknown"
            
            # Check for date change to insert subtotal
            if current_date is not None and usage_date != current_date:
                insert_subtotal(current_date)
                # Reset subtotals
                sub_film_count = 0
                sub_rtk = [0.0] * 8
                sub_ndt = [0.0] * 6
            
            current_date = usage_date
            
            mat_id = entry['Material ID']
            mat_name = self.get_material_display_name(mat_id)
            
            entry_time = entry.get('Entry Time', '')
            if pd.notna(entry_time):
                entry_time = pd.to_datetime(entry_time).strftime('%Y-%m-%d %H:%M')
            
            # Get film count
            film_count = entry.get('Film Count', 0)
            if pd.notna(film_count):
                film_count_val = float(film_count)
                total_film_count += film_count_val
                sub_film_count += film_count_val
                film_count_str = f"{film_count_val:.1f}"
            else:
                film_count_str = "0.0"
            
            # Get RTK values
            rtk_values = []
            row_rtk_total = 0
            rtk_categories = ["센터미스", "농도", "마킹미스", "필름마크", "취급부주의", "고객불만", "기타"]
            
            for i, category in enumerate(rtk_categories):
                value = entry.get(f'RTK_{category}', 0)
                if pd.notna(value):
                    val_float = float(value)
                    rtk_values.append(f"{val_float:.1f}")
                    row_rtk_total += val_float
                    total_rtk[i] += val_float
                    sub_rtk[i] += val_float
                else:
                    rtk_values.append("0.0")
            
            rtk_values.append(f"{row_rtk_total:.1f}")  # Add row total
            total_rtk[7] += row_rtk_total
            sub_rtk[7] += row_rtk_total
            
            # Get NDT materials values
            ndt_values = []
            ndt_materials = ["자분", "페인트", "침투제", "세척제", "현상제", "형광"]
            
            for i, material in enumerate(ndt_materials):
                value = entry.get(f'NDT_{material}', 0)
                if pd.notna(value):
                    val_float = float(value)
                    ndt_values.append(f"{val_float:.1f}")
                    total_ndt[i] += val_float
                    sub_ndt[i] += val_float
                else:
                    ndt_values.append("0.0")
            
            # Get remark with manager name if exists
            user_val = entry.get('User', '')
            note_val = entry.get('Note', '')
            display_note = f"[{user_val}] {note_val}" if user_val else note_val
            
            # Insert with index as tag for reliable deletion
            self.daily_usage_tree.insert('', tk.END, values=(
                usage_date,
                entry.get('Site', ''),
                mat_name,
                film_count_str,
                *rtk_values,
                *ndt_values,
                display_note,
                entry_time
            ), tags=(str(idx),))
            
        # Insert last daily subtotal and final total row if data exists
        if not filtered_df.empty:
            # Last day subtotal
            if current_date is not None:
                insert_subtotal(current_date)
            
            # Final overall total
            self.daily_usage_tree.tag_configure('total', background='#E8F4F8', font=('Arial', 9, 'bold'))
            
            total_values = [
                '=== 전체 누계 ===',
                '',
                '',
                f"{total_film_count:.1f}",
                *[f"{v:.1f}" for v in total_rtk],
                *[f"{v:.1f}" for v in total_ndt],
                '',
                ''
            ]
            self.daily_usage_tree.insert('', tk.END, values=total_values, tags=('total',))
    
    def delete_daily_usage_entry(self):
        """선택된 일일 사용량 기록 삭제 (인덱스 기반으로 정확하게 삭제 및 재고 환원)"""
        selected_items = self.daily_usage_tree.selection()
        
        if not selected_items:
            messagebox.showwarning("선택 오류", "삭제할 항목을 선택해주세요.")
            return
        
        # 사용자 확인
        result = messagebox.askyesno("삭제 확인", f"{len(selected_items)}개의 기록을 삭제하시겠습니까?\n(삭제 시 차감되었던 재고도 자동으로 환원됩니다.)")
        
        if not result:
            return
        
        indices_to_delete = []
        for item in selected_items:
            tags = self.daily_usage_tree.item(item, 'tags')
            if tags:
                try:
                    df_idx = int(tags[0])
                    indices_to_delete.append(df_idx)
                except ValueError:
                    continue
        
        if not indices_to_delete:
            messagebox.showwarning("삭제 실패", "삭제할 항목의 데이터 정보를 찾을 수 없습니다.")
            return

        deleted_count = 0
        try:
            # Sort indices in descending order to avoid index shift issues during deletion
            # But wait, we are using boolean masking or actual indices in the original dataframe
            # Since we will drop them at once or one by one from the original, it's better to collect all indices.
            
            for idx in indices_to_delete:
                if idx not in self.daily_usage_df.index:
                    continue
                
                entry = self.daily_usage_df.loc[idx]
                
                # --- 재고 환원 로직 (해당 내역으로 인해 차감되었던 트랜잭션 삭제) ---
                site = entry.get('Site', '')
                usage_date = pd.to_datetime(entry.get('Date'))
                mat_id = entry.get('Material ID')
                
                # 해당 내역과 관련된 트랜잭션 찾기 (날짜, Material ID, Note가 현장 사용 자동 차감인 것)
                # Note 형식을 add_daily_usage_entry와 동일하게 구성
                note_pattern = f'{site} 현장 사용 (자동 차감)'
                
                # 필터링하여 관련 트랜잭션 삭제
                if not self.transactions_df.empty:
                    trans_mask = (
                        (pd.to_datetime(self.transactions_df['Date']).dt.date == usage_date.date()) &
                        (self.transactions_df['Material ID'] == mat_id) &
                        (self.transactions_df['Type'] == 'OUT') &
                        (self.transactions_df['Note'] == note_pattern)
                    )
                    
                    # NDT 자재 트랜잭션도 포함될 수 있음. NDT 자재들은 Material ID가 다를 수 있으나 
                    # Note는 동일하게 "현장 사용 (자동 차감)"으로 들어감.
                    # NDT 자재들의 Material ID는 entry에 NDT_xxx 필드로 명시되어 있음
                    ndt_materials = ["자분", "페인트", "침투제", "세척제", "현상제"]
                    ndt_mat_ids = []
                    for ndt_name in ndt_materials:
                        if entry.get(f'NDT_{ndt_name}', 0) > 0:
                            ndt_mat_rows = self.materials_df[self.materials_df['품목명'] == ndt_name]
                            if not ndt_mat_rows.empty:
                                ndt_mat_ids.append(ndt_mat_rows['Material ID'].values[0])
                    
                    if ndt_mat_ids:
                        trans_mask = trans_mask | (
                            (pd.to_datetime(self.transactions_df['Date']).dt.date == usage_date.date()) &
                            (self.transactions_df['Material ID'].isin(ndt_mat_ids)) &
                            (self.transactions_df['Type'] == 'OUT') &
                            (self.transactions_df['Note'] == note_pattern)
                        )

                    # 트랜잭션 삭제
                    self.transactions_df = self.transactions_df[~trans_mask]
                
                # 사용량 기록 삭제
                self.daily_usage_df = self.daily_usage_df.drop(idx)
                deleted_count += 1
                
        except Exception as e:
            messagebox.showerror("삭제 오류", f"삭제 중 오류가 발생했습니다: {e}")
            print(f"삭제 중 상세 오류: {e}")
        
        # 데이터 저장 및 화면 갱신
        if deleted_count > 0:
            self.daily_usage_df.reset_index(drop=True, inplace=True)
            self.save_data()
            self.update_daily_usage_view()
            if hasattr(self, 'update_stock_view'):
                self.update_stock_view()
            messagebox.showinfo("완료", f"{deleted_count}개의 기록 및 관련 재고 차감 내역이 삭제되었습니다.")
        else:
            messagebox.showwarning("삭제 실패", "삭제할 기록을 찾지 못했습니다.")


    def setup_keyboard_shortcuts(self):
        """Setup keyboard shortcuts for navigation"""
        # Ctrl+Tab to switch between notebook tabs (forward)
        self.root.bind('<Control-Tab>', self.next_tab)
        # Ctrl+Shift+Tab to switch between notebook tabs (backward)
        self.root.bind('<Control-Shift-Tab>', self.prev_tab)
        
        # Alt+숫자 for direct tab access
        self.root.bind('<Alt-Key-1>', lambda e: self.notebook.select(0))
        self.root.bind('<Alt-Key-2>', lambda e: self.notebook.select(1))
        self.root.bind('<Alt-Key-3>', lambda e: self.notebook.select(2))
        self.root.bind('<Alt-Key-4>', lambda e: self.notebook.select(3))
        self.root.bind('<Alt-Key-5>', lambda e: self.notebook.select(4))
        self.root.bind('<Alt-Key-6>', lambda e: self.notebook.select(5))
        
        # Right-click on notebook for tab reordering
        self.notebook.bind('<Button-3>', self.show_tab_context_menu)
        
        # Save tab config when tab selection changes
        self.notebook.bind('<<NotebookTabChanged>>', self.on_tab_changed)
    
    def next_tab(self, event=None):
        """Switch to next tab"""
        current = self.notebook.index(self.notebook.select())
        total = self.notebook.index('end')
        next_tab = (current + 1) % total
        self.notebook.select(next_tab)
        return 'break'
    
    def prev_tab(self, event=None):
        """Switch to previous tab"""
        current = self.notebook.index(self.notebook.select())
        total = self.notebook.index('end')
        prev_tab = (current - 1) % total
        self.notebook.select(prev_tab)
        return 'break'
    
    def show_tab_context_menu(self, event):
        """Show context menu for tab reordering"""
        # Identify which tab was clicked
        try:
            clicked_tab = self.notebook.index(f"@{event.x},{event.y}")
            self.notebook.select(clicked_tab)
            
            # Create context menu
            context_menu = tk.Menu(self.root, tearoff=0)
            
            # Only show "Move Left" if not the first tab
            if clicked_tab > 0:
                context_menu.add_command(label="← 탭 왼쪽으로 이동", 
                                        command=lambda: self.move_tab_left(clicked_tab))
            
            # Only show "Move Right" if not the last tab
            if clicked_tab < self.notebook.index('end') - 1:
                context_menu.add_command(label="탭 오른쪽으로 이동 →", 
                                        command=lambda: self.move_tab_right(clicked_tab))
            
            # Show menu at cursor position
            context_menu.post(event.x_root, event.y_root)
        except:
            pass
    
    def move_tab_left(self, tab_index):
        """Move tab one position to the left"""
        if tab_index > 0:
            # Get tab info
            tab = self.notebook.tabs()[tab_index]
            text = self.notebook.tab(tab_index, "text")
            
            # Remove and reinsert at new position
            self.notebook.insert(tab_index - 1, tab, text=text)
            self.notebook.select(tab_index - 1)
            
            # Save configuration immediately
            self.save_tab_config()
    
    def move_tab_right(self, tab_index):
        """Move tab one position to the right"""
        if tab_index < self.notebook.index('end') - 1:
            # Get tab info
            tab = self.notebook.tabs()[tab_index]
            text = self.notebook.tab(tab_index, "text")
            
            # Remove and reinsert at new position
            self.notebook.insert(tab_index + 2, tab, text=text)
            self.notebook.select(tab_index + 1)
            
            # Save configuration immediately
            self.save_tab_config()
    
    def on_tab_changed(self, event=None):
        """Handle tab selection change event"""
        # Save configuration when tab changes
        self.save_tab_config()
    
    def auto_save_to_list(self, event, combobox, data_list, config_key):
        """Helper to auto-save new entry from combobox to a list and update all related UI"""
        new_val = combobox.get().strip()
        if not new_val:
            return
            
        if new_val not in data_list:
            data_list.append(new_val)
            data_list.sort()
            
            # Update all comboboxes that use this list
            if config_key == 'sites':
                if hasattr(self, 'ent_daily_site'): self.ent_daily_site['values'] = self.sites
                if hasattr(self, 'cb_trans_site'): self.cb_trans_site['values'] = self.sites
                if hasattr(self, 'cb_daily_filter_site'): # Update filter if it exists
                     self.cb_daily_filter_site['values'] = ['전체'] + self.sites
            elif config_key == 'users':
                if hasattr(self, 'cb_daily_user'): self.cb_daily_user['values'] = self.users
                if hasattr(self, 'ent_user'): # Transaction tab user field is Entry, but could be upgraded
                    pass
            
            self.save_tab_config()

    def save_tab_config(self):
        """Save current tab configuration (order and selected tab)"""
        try:
            config = {
                'selected_tab': self.notebook.index(self.notebook.select()),
                'tab_order': [],
                'sites': self.sites,
                'users': getattr(self, 'users', [])
            }
            
            # Save current stock column widths
            config['stock_col_widths'] = {}
            if hasattr(self, 'stock_tree'):
                for col in self.stock_tree['columns']:
                    config['stock_col_widths'][col] = self.stock_tree.column(col, 'width')
            
            # Save tab order
            for tab in self.notebook.tabs():
                config['tab_order'].append(self.notebook.tab(tab, "text"))
            
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Failed to save tab config: {e}")
    
    def load_tab_config(self):
        """Load and restore tab configuration"""
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                
                # Get current tab order
                current_order = []
                for tab in self.notebook.tabs():
                    tab_index = self.notebook.index(tab)
                    tab_text = self.notebook.tab(tab_index, "text")
                    current_order.append((tab_text, tab))
                
                # Create mapping from text to tab widget
                tab_map = {text: tab for text, tab in current_order}
                
                # Restore tab order if saved order exists
                saved_order = config.get('tab_order', [])
                if saved_order and len(saved_order) == len(current_order):
                    # Reorder tabs according to saved order
                    for i, tab_text in enumerate(saved_order):
                        if tab_text in tab_map:
                            tab = tab_map[tab_text]
                            current_pos = self.notebook.index(tab)
                            if current_pos != i:
                                self.notebook.insert(i, tab, text=tab_text)
                
                # Restore selected tab
                selected = config.get('selected_tab', 0)
                if 0 <= selected < self.notebook.index('end'):
                    self.notebook.select(selected)
                
                # Restore sites list
                self.sites = config.get('sites', [])
                # If sites list is empty, try to populate from current daily_usage_df
                if not self.sites and not self.daily_usage_df.empty:
                    self.sites = sorted(self.daily_usage_df['Site'].dropna().unique().tolist())
                    self.sites = [str(s).strip() for s in self.sites if str(s).strip()]
                
                # Restore users list
                self.users = config.get('users', [])
                if not self.users and not self.daily_usage_df.empty and 'User' in self.daily_usage_df.columns:
                    self.users = sorted(self.daily_usage_df['User'].dropna().unique().tolist())
                    self.users = [str(u).strip() for u in self.users if str(u).strip()]
                
                # Refresh UI dropdowns with restored data
                if hasattr(self, 'ent_daily_site'): self.ent_daily_site['values'] = self.sites
                if hasattr(self, 'cb_trans_site'): self.cb_trans_site['values'] = self.sites
                if hasattr(self, 'cb_daily_filter_site'): self.cb_daily_filter_site['values'] = ['전체'] + self.sites
                
                if hasattr(self, 'cb_daily_user'): self.cb_daily_user['values'] = self.users
                if hasattr(self, 'ent_user') and not isinstance(self.ent_user, tk.Entry): 
                    self.ent_user['values'] = self.users
                
                # Restore stock column widths
                stock_col_widths = config.get('stock_col_widths', {})
                if stock_col_widths and hasattr(self, 'stock_tree'):
                    for col, width in stock_col_widths.items():
                        try:
                            self.stock_tree.column(col, width=int(width))
                        except:
                            pass
        except Exception as e:
            print(f"Failed to load tab config: {e}")
    
    def on_closing(self):
        """Handle window closing event"""
        self.save_tab_config()
        self.root.destroy()
    
    def export_stock_to_excel(self):
        """Export current stock data to Excel"""
        try:
            # Get current filtered data from treeview
            stock_data = []
            
            for item in self.stock_tree.get_children():
                values = self.stock_tree.item(item, 'values')
                stock_data.append({
                    'ID': values[0],
                    '회사코드': values[1],
                    '관리품번': values[2],
                    '품목명': values[3],
                    'SN': values[4],
                    '창고': values[5],
                    '모델명': values[6],
                    '규격': values[7],
                    '품목군코드': values[8],
                    '공급업체': values[9],
                    '제조사': values[10],
                    '제조국': values[11],
                    '가격': values[12],
                    '관리단위': values[13],
                    '수량': values[14],
                    '재고하한': values[15]
                })
            
            if not stock_data:
                messagebox.showinfo("알림", "내보낼 데이터가 없습니다.")
                return
            
            # Prepare filename with current date
            current_date = datetime.datetime.now().strftime('%Y%m%d')
            filename = f"재고현황_{current_date}.xlsx"
            
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile=filename,
                title="재고 현황 저장",
                filetypes=[("Excel files", "*.xlsx")]
            )
            
            if save_path:
                stock_df = pd.DataFrame(stock_data)
                
                # Use ExcelWriter to set column widths
                with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                    stock_df.to_excel(writer, index=False, sheet_name='재고현황')
                    workbook = writer.book
                    worksheet = writer.sheets['재고현황']
                    
                    # Get column widths from treeview and apply to Excel
                    # Note: Excel column width is roughly pixels / 7 for default font
                    if hasattr(self, 'stock_tree'):
                        for idx, col in enumerate(self.stock_tree['columns']):
                            width = self.stock_tree.column(col, 'width')
                            # Excel unit is approx 1/7 of pixel width for standard Excel font
                            excel_width = width / 7.0 
                            # Set a reasonable max/min
                            excel_width = max(5, min(excel_width, 100))
                            column_letter = chr(65 + idx) if idx < 26 else None
                            if column_letter:
                                worksheet.column_dimensions[column_letter].width = excel_width
                
                messagebox.showinfo("완료", f"재고 현황이 저장되었습니다.\n저장 위치: {save_path}")
                
        except Exception as e:
            messagebox.showerror("오류", f"내보내기 실패: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = MaterialManager(root)
    root.mainloop()
