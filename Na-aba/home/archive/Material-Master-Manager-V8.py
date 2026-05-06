import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import sys
import subprocess
import os

def install_and_import(package):
    try:
        __import__(package)
    except ImportError:
        try:
            print(f"Installing {package}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        except Exception as e:
            print(f"Failed to install {package}: {e}")
            messagebox.showerror("Dependency Error", f"Failed to install required package '{package}'.\nPlease install it manually: pip install {package}")
            sys.exit(1)

# Auto-install dependencies
try:
    import tkcalendar
except ImportError:
    install_and_import('tkcalendar')
    import tkcalendar

try:
    import pandas as pd
except ImportError:
    install_and_import('pandas')
    import pandas as pd

try:
    import openpyxl
except ImportError:
    install_and_import('openpyxl')
    import openpyxl

from tkcalendar import DateEntry
import datetime
import json
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
        self.users = [
            "주진철", "우명광", "김진환", "장승대", "김성렬", "박광복", "주영광"
        ] # Initialize worker/name list
        self.warehouses = [] # Initialize warehouse list
        self.equipments = [] # Initialize equipment list
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
        
        # Registry for draggable items {config_key: widget_instance}
        self.draggable_items = {}
        self.memos = {} # {memo_key: {'container': frame, 'text_widget': text}}
        self.layout_locked = False

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
                                                'RTK_취급부주의', 'RTK_고객불만', 'RTK_기타', '장비명', '검사방법', '검사량',
                                                '단가', '출장비', '일식', '검사비', 'User', 'User2', 'User3', 'User4', 'User5', 'User6'])
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
                
                # Force specific columns to string type and clean numeric artifacts (.0, -0.0)
                str_cols = ['회사코드', '관리품번', '품목명', 'SN', '창고', '모델명', '규격', 
                            '품목군코드', '공급업체', '제조사', '제조국', '관리단위']
                for col in str_cols:
                    if col in self.materials_df.columns:
                        self.materials_df[col] = self.materials_df[col].astype(str).replace(['nan', 'None', 'NULL', '-0.0', '0.0'], '')
                        self.materials_df[col] = self.materials_df[col].str.replace(r'\.0$', '', regex=True)
                
                self.transactions_df = pd.read_excel(self.db_path, sheet_name='Transactions')
                # Ensure it has all required columns
                for col in ['Site', 'User', 'Note', 'Material ID', 'Type', 'Quantity', 'Date']:
                    if col not in self.transactions_df.columns:
                        self.transactions_df[col] = '' if col != 'Quantity' else 0.0
                
                # Ensure Date column is datetime and Material ID is numeric
                if not self.transactions_df.empty:
                    self.transactions_df['Date'] = pd.to_datetime(self.transactions_df['Date'], errors='coerce')
                    self.transactions_df['Material ID'] = pd.to_numeric(self.transactions_df['Material ID'], errors='coerce')
                    
                    # Force string columns and clean numeric artifacts
                    for col in ['Type', 'Note', 'User', 'Site']:
                        self.transactions_df[col] = self.transactions_df[col].astype(str).replace(['nan', 'None', 'NULL', '-0.0', '0.0', 'NaN'], '')
                        self.transactions_df[col] = self.transactions_df[col].str.replace(r'\.0$', '', regex=True)
                    
                    # One-time cleanup: Remove '현장사용' and redundant model names from historical notes
                    self.transactions_df['Note'] = self.transactions_df['Note'].astype(str).str.replace('현장사용', '', regex=False).str.strip()
                    
                    # Clean up notes that are identical to model names
                    if not self.transactions_df.empty and not self.materials_df.empty:
                        # Create a map for Material ID -> Model Name
                        id_to_model = self.materials_df.set_index('Material ID')['모델명'].astype(str).to_dict()
                        
                        def clean_redundant_note(row):
                            note = str(row['Note']).strip()
                            mat_id = row['Material ID']
                            model = str(id_to_model.get(mat_id, '')).strip()
                            if note and model and note == model:
                                return ''
                            return note
                            
                        self.transactions_df['Note'] = self.transactions_df.apply(clean_redundant_note, axis=1)
                
                # Load monthly usage data
                try:
                    self.monthly_usage_df = pd.read_excel(self.db_path, sheet_name='MonthlyUsage', dtype={'Site': str, 'Note': str})
                    if not self.monthly_usage_df.empty:
                        self.monthly_usage_df['Material ID'] = pd.to_numeric(self.monthly_usage_df['Material ID'], errors='coerce')
                        self.monthly_usage_df['Entry Date'] = pd.to_datetime(self.monthly_usage_df['Entry Date'])
                        # Add Site column if it doesn't exist (for backward compatibility)
                        if 'Site' not in self.monthly_usage_df.columns:
                            self.monthly_usage_df['Site'] = ''
                except:
                    self.monthly_usage_df = pd.DataFrame(columns=['Material ID', 'Year', 'Month', 'Site', 'Usage', 'Note', 'Entry Date'])
                
                # Load daily usage data if exists
                try:
                    # Explicitly set dtypes to avoid float inference for empty columns
                    self.daily_usage_df = pd.read_excel(self.db_path, sheet_name='DailyUsage', 
                                                        dtype={'Site': str, 'Note': str, 'User': str})
                    
                    if not self.daily_usage_df.empty:
                        self.daily_usage_df['Date'] = pd.to_datetime(self.daily_usage_df['Date'])
                        self.daily_usage_df['Entry Time'] = pd.to_datetime(self.daily_usage_df['Entry Time'])
                        
                        # Fill NaNs and clean numeric artifacts in string columns
                        string_columns = ['Site', 'Note', 'User', 'User2', 'User3', 'User4', 'User5', 'User6', '장비명', '검사방법']
                        for col in string_columns:
                            if col not in self.daily_usage_df.columns:
                                self.daily_usage_df[col] = ''
                            self.daily_usage_df[col] = self.daily_usage_df[col].astype(str).replace(['nan', 'None', 'NULL', '-0.0', '0.0', 'NaN'], '')
                        
                        # Add/Fix columns for auto-calculation
                        if '검사방법' not in self.daily_usage_df.columns and '검사량' in self.daily_usage_df.columns:
                            # Migrate old string '검사량' (PAUT, UT...) to '검사방법'
                            self.daily_usage_df['검사방법'] = self.daily_usage_df['검사량'].astype(str)
                            self.daily_usage_df['검사량'] = 0.0
                        
                        for col in ['검사량', '단가', '출장비', '일식', '검사비']:
                            if col not in self.daily_usage_df.columns:
                                self.daily_usage_df[col] = 0.0
                            else:
                                self.daily_usage_df[col] = pd.to_numeric(self.daily_usage_df[col], errors='coerce').fillna(0.0)
                        
                        # Ensure Material ID is numeric
                        if 'Material ID' in self.daily_usage_df.columns:
                            self.daily_usage_df['Material ID'] = pd.to_numeric(self.daily_usage_df['Material ID'], errors='coerce')
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
                                                        'RTK_취급부주의', 'RTK_고객불만', 'RTK_기타', 'User', '장비명', '검사량',
                                                        '단가', '출장비', '일식', '검사비'])
                
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
                '모델명', '규격', '품목군코드', '공급업체', '제조사', '제조국', 
                '가격', '관리단위', '수량', '재고하한'
            ])
            self.transactions_df = pd.DataFrame(columns=['Date', 'Material ID', 'Site', 'Type', 'Quantity', 'Note', 'User'])
            self.monthly_usage_df = pd.DataFrame(columns=['Material ID', 'Year', 'Month', 'Site', 'Usage', 'Note', 'Entry Date'])
            self.daily_usage_df = pd.DataFrame(columns=['Date', 'Site', 'Material ID', 'Usage', 'Note', 'Entry Time', 
                                                'RTK_센터미스', 'RTK_농도', 'RTK_마킹미스', 'RTK_필름마크', 
                                                'RTK_취급부주의', 'RTK_고객불만', 'RTK_기타', '장비명', '검사량'])
    
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
                # Check the actual dtype of the column
                col_dtype = self.materials_df[col].dtype
                
                # Handle different data types appropriately
                if col_dtype == 'float64' or col_dtype == 'int64':
                    # Numeric columns - convert to float, empty becomes 0.0
                    try:
                        val = float(val) if str(val).strip() else 0.0
                    except (ValueError, TypeError):
                        val = 0.0
                else:
                    # String/object columns - handle empty values
                    if val == '' or val == 'nan' or pd.isna(val):
                        val = ''
                    else:
                        val = str(val)
                
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
        self.cb_trans_site.bind('<FocusOut>', lambda e: self.auto_save_to_list(e, self.cb_trans_site, self.sites, 'sites'))
        self.cb_trans_site.bind('<Return>', lambda e: self.auto_save_to_list(e, self.cb_trans_site, self.sites, 'sites'))
        
        ttk.Label(trans_frame, text="창고:").grid(row=2, column=2, padx=5, pady=2, sticky='w')
        self.cb_warehouse = ttk.Combobox(trans_frame, width=28, values=self.warehouses)
        self.cb_warehouse.grid(row=2, column=3, padx=5, pady=2, sticky='w')
        self.cb_warehouse.bind('<FocusOut>', lambda e: self.auto_save_to_list(e, self.cb_warehouse, self.warehouses, 'warehouses'))
        self.cb_warehouse.bind('<Return>', lambda e: self.auto_save_to_list(e, self.cb_warehouse, self.warehouses, 'warehouses'))
        
        ttk.Label(trans_frame, text="담당자:").grid(row=3, column=0, padx=5, pady=2, sticky='w')
        self.ent_user = ttk.Combobox(trans_frame, width=28, values=getattr(self, 'users', []))
        self.ent_user.grid(row=3, column=1, padx=5, pady=2)
        self.ent_user.bind('<FocusOut>', lambda e: self.auto_save_to_list(e, self.ent_user, self.users, 'users'))
        self.ent_user.bind('<Return>', lambda e: self.auto_save_to_list(e, self.ent_user, self.users, 'users'))
        
        ttk.Label(trans_frame, text="비고:").grid(row=3, column=2, padx=5, pady=2, sticky='w')
        self.ent_note = ttk.Entry(trans_frame, width=30)
        self.ent_note.grid(row=3, column=3, padx=5, pady=2)
        
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
        
        ttk.Label(history_ctrl_frame, text="품목명 필터:").pack(side='left', padx=(20, 5))
        self.cb_trans_filter_mat = ttk.Combobox(history_ctrl_frame, width=40, state="readonly")
        self.cb_trans_filter_mat.pack(side='left', padx=5)
        self.cb_trans_filter_mat.bind('<<ComboboxSelected>>', lambda e: self.update_transaction_view())
        
        ttk.Label(history_ctrl_frame, text="현장 필터:").pack(side='left', padx=(20, 5))
        self.cb_trans_filter_site = ttk.Combobox(history_ctrl_frame, width=15, state="readonly")
        self.cb_trans_filter_site.pack(side='left', padx=5)
        self.cb_trans_filter_site.bind('<<ComboboxSelected>>', lambda e: self.update_transaction_view())
        
        # Treeview for history
        tree_scroll_frame = ttk.Frame(history_frame)
        tree_scroll_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        inout_vsb = ttk.Scrollbar(tree_scroll_frame, orient="vertical")
        inout_hsb = ttk.Scrollbar(tree_scroll_frame, orient="horizontal")
        
        columns = ('날짜', '현장', '품목명', '모델명', '구분', '수량', '재고', '담당자', '비고')
        self.inout_tree = ttk.Treeview(tree_scroll_frame, columns=columns, show='headings', height=10,
                                       yscrollcommand=inout_vsb.set, xscrollcommand=inout_hsb.set)
        
        inout_vsb.config(command=self.inout_tree.yview)
        inout_hsb.config(command=self.inout_tree.xview)
        
        col_widths = [150, 120, 180, 150, 70, 70, 70, 100, 200]
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
            # Try numeric fallback just in case some IDs are still strings
            try:
                num_id = int(float(mat_id))
                mat_row = self.materials_df[self.materials_df['Material ID'] == num_id]
            except (ValueError, TypeError):
                pass
                
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
            
        # Sync equipment suggestions with the same list as materials + custom equipments
        if hasattr(self, 'cb_daily_equip'):
            # Combine custom equipments with the full material display list (all_vals)
            combined_equip = list(set([str(e).strip() for e in self.equipments + all_vals if pd.notna(e) and str(e).strip()]))
            combined_equip.sort()
            self.cb_daily_equip['values'] = combined_equip

        if hasattr(self, 'cb_trans_filter_mat'):
            self.cb_trans_filter_mat['values'] = ["전체"] + all_vals
            if not self.cb_trans_filter_mat.get():
                self.cb_trans_filter_mat.set("전체")
        
        if hasattr(self, 'cb_trans_filter_site'):
            self.cb_trans_filter_site['values'] = ["전체"] + sorted(self.sites)
            if not self.cb_trans_filter_site.get():
                self.cb_trans_filter_site.set("전체")

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
        """Record an IN or OUT transaction"""
        try:
            mat_selection = self.cb_material.get().strip()
            t_type = self.cb_type.get()
            user = self.ent_user.get().strip()
            
            if not mat_selection:
                messagebox.showwarning("입력 오류", "자재를 선택해주세요.")
                return
            if not t_type:
                messagebox.showwarning("입력 오류", "구분(입고/출고)을 선택해주세요.")
                return
                
            try:
                qty_str = self.ent_qty.get().strip()
                if not qty_str:
                    messagebox.showwarning("입력 오류", "수량을 입력해주세요.")
                    return
                qty = float(qty_str)
            except ValueError:
                messagebox.showwarning("입력 오류", "수량은 숫자여야 합니다.")
                return
                
            note = self.ent_note.get().strip()
            
            # Extract pure material name from selection
            mat_name = mat_selection
            if " - " in mat_name: mat_name = mat_name.split(" - ")[0]
            if " (SN: " in mat_name: mat_name = mat_name.split(" (SN: ")[0]
            pure_mat_name = mat_name.strip()
            
            # Find material ID
            # Ensure Material ID is treated consistently
            mat_rows = self.materials_df[self.materials_df['품목명'] == pure_mat_name]
            
            if mat_rows.empty:
                # If exact match fails, try a case-insensitive or stripped match
                mat_rows = self.materials_df[self.materials_df['품목명'].str.strip() == pure_mat_name]
            
            if mat_rows.empty:
                messagebox.showerror("오류", f"'{pure_mat_name}' 자재를 찾을 수 없습니다.\n먼저 자재를 등록하거나 정확한 명칭을 입력해주세요.")
                return
                
            mat_id = mat_rows['Material ID'].values[0]
            
            # Update Warehouse in materials_df
            warehouse = str(self.cb_warehouse.get()).strip()
            if warehouse:
                # Type-safe assignment
                if '창고' in self.materials_df.columns:
                    mask = self.materials_df['Material ID'] == mat_id
                    if mask.any():
                        self.materials_df.loc[mask, '창고'] = warehouse
            
            # Create transaction record
            new_trans = {
                'Date': datetime.datetime.now(),
                'Material ID': mat_id,
                'Type': t_type,
                'Quantity': qty,
                'Note': note,
                'User': user,
                'Site': self.cb_trans_site.get() if hasattr(self, 'cb_trans_site') else ''
            }
            
            # Add to dataframe
            self.transactions_df = pd.concat([self.transactions_df, pd.DataFrame([new_trans])], ignore_index=True)
            
            # Save all data (force sync)
            self.save_data()
            
            # Check if it was really added (for debug feedback)
            last_count = len(self.transactions_df)
            
            # Refresh views
            self.update_stock_view()
            self.update_transaction_view()
            self.update_material_combo() 
            
            # Auto-save site/user to lists
            site_value = new_trans['Site']
            if site_value and site_value not in self.sites:
                self.sites.append(site_value)
                self.sites.sort()
                self.save_tab_config()
            
            if user and user not in self.users:
                self.users.append(user)
                self.users.sort()
                self.save_tab_config()
            
            # Success feedback
            messagebox.showinfo("완료", f"{pure_mat_name} {t_type} 처리되었습니다.\n(전체 기록 수: {last_count}개)")
            
            # Clear UI fields
            self.ent_qty.delete(0, tk.END)
            self.ent_note.delete(0, tk.END)
            if hasattr(self, 'ent_user') and hasattr(self.ent_user, 'set'):
                self.ent_user.set('')
            elif hasattr(self, 'ent_user'):
                self.ent_user.delete(0, tk.END)
            
            if hasattr(self, 'cb_warehouse'): self.cb_warehouse.set('')
            if hasattr(self, 'cb_trans_site'): self.cb_trans_site.set('')
            
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            messagebox.showerror("저장 오류", f"기록 저장 중 기술적인 오류가 발생했습니다:\n{e}\n\n상세 정보:\n{error_details}")
        
        # Ensure view is updated regardless
        self.update_transaction_view()

    def update_transaction_view(self):
        """Populate the transaction history Treeview"""
        try:
            if not hasattr(self, 'inout_tree'):
                return
                
            # Clear current view
            for item in self.inout_tree.get_children():
                self.inout_tree.delete(item)
                
            if self.transactions_df.empty:
                return
                
            # Update frame title with count
            if hasattr(self, 'scrollable_frame'):
                # Try to find the history LabelFrame to update its text
                for child in self.scrollable_frame.winfo_children():
                    if isinstance(child, ttk.LabelFrame) and child.cget("text").startswith("최근 입출고 내역"):
                        child.configure(text=f"최근 입출고 내역 (총 {len(self.transactions_df)}건)")
            
            # Display last 500 transactions, descending by date
            df_to_show = self.transactions_df.copy()
            
            # Apply material filter if selected
            if hasattr(self, 'cb_trans_filter_mat'):
                selected_mat = self.cb_trans_filter_mat.get()
                if selected_mat and selected_mat != "전체":
                    # Find all mat_ids that match this display name
                    matching_ids = []
                    for mat_id in df_to_show['Material ID'].unique():
                        if self.get_material_display_name(mat_id) == selected_mat:
                            matching_ids.append(mat_id)
                    
                    df_to_show = df_to_show[df_to_show['Material ID'].isin(matching_ids)]
            
            # Apply site filter if selected
            if hasattr(self, 'cb_trans_filter_site'):
                selected_site = self.cb_trans_filter_site.get()
                if selected_site and selected_site != "전체":
                    df_to_show = df_to_show[df_to_show['Site'] == selected_site]

            df_sorted = df_to_show.sort_values(by='Date', ascending=False, na_position='last').head(500)
            
            for idx, row in df_sorted.iterrows():
                mat_id = row['Material ID']
                mat_name = self.get_material_display_name(mat_id)
                
                # Safe date formatting
                raw_date = row.get('Date')
                if pd.isna(raw_date):
                    date_str = "Unknown"
                elif hasattr(raw_date, 'strftime'):
                    date_str = raw_date.strftime('%Y-%m-%d %H:%M:%S')
                else:
                    date_str = str(raw_date)
                
                current_stock = self.calculate_current_stock(mat_id)
                
                # Find model name from materials_df
                model_name = ""
                mat_row = self.materials_df[self.materials_df['Material ID'] == mat_id]
                if not mat_row.empty:
                    model_name = str(mat_row.iloc[0].get('모델명', '')).replace('nan', '').replace('None', '')
                
                self.inout_tree.insert('', tk.END, values=(
                    date_str,
                    str(row.get('Site', '')).replace('nan', ''),
                    mat_name,
                    model_name,
                    row.get('Type', ''),
                    row.get('Quantity', 0),
                    current_stock,
                    str(row.get('User', '')).replace('nan', ''),
                    str(row.get('Note', '')).replace('nan', '')
                ), tags=(str(idx),))
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"Transaction View Error: {e}\n{error_details}")

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
        columns = ('연도', '월', '현장', '검사량', '단가', '출장비', '일식', '검사비', '품목명', '필름매수', '센터미스', '농도', '마킹미스', '필름마크', '취급부주의', '고객불만', '기타', 'RT총계', '형광자분', '흑색자분', '백색페인트', '침투제', '세척제', '현상제', '형광침투제', '비고')
        self.report_tree = ttk.Treeview(tree_frame, columns=columns, show='headings',
                                       yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        vsb.config(command=self.report_tree.yview)
        hsb.config(command=self.report_tree.xview)
        
        # Column configuration with same widths as monthly usage tab
        col_widths = [80, 60, 120, 100, 100, 100, 100, 100, 200, 100, 70, 70, 70, 70, 70, 70, 70, 70, 80, 80, 80, 80, 80, 80, 80, 200]
        # Columns to center-align (numeric values)
        center_cols = ['검사량', '단가', '출장비', '일식', '검사비', '필름매수', '센터미스', '농도', '마킹미스', '필름마크', '취급부주의', '고객불만', '기타', 'RT총계', '형광자분', '흑색자분', '백색페인트', '침투제', '세척제', '현상제', '형광침투제']
        
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
        
        # Add NDT materials (include both old and new names for compatibility)
        ndt_categories = ['NDT_형광자분', 'NDT_자분', 'NDT_흑색자분', 'NDT_페인트', 'NDT_백색페인트', 'NDT_침투제', 'NDT_세척제', 'NDT_현상제', 'NDT_형광', 'NDT_형광침투제']
        for mat in ndt_categories:
            if mat in df.columns:
                agg_dict[mat] = 'sum'

        # Add cost fields
        cost_categories = ['검사량', '단가', '출장비', '일식', '검사비']
        for cost_col in cost_categories:
            if cost_col in df.columns:
                agg_dict[cost_col] = 'sum'
        
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
        total_ndt_fluorescent_mag = 0.0
        total_ndt_magnet = 0.0
        total_ndt_paint = 0.0
        total_ndt_penetrant = 0.0
        total_ndt_cleaner = 0.0
        total_ndt_developer = 0.0
        total_ndt_fluorescent_pen = 0.0
        total_test_amount = 0.0
        total_unit_price = 0.0
        total_travel_cost = 0.0
        total_meal_cost = 0.0
        total_test_fee = 0.0
        
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
            
            # NDT values (summing both old and new names)
            ndt_fluorescent_mag = entry.get('NDT_형광자분', 0.0)
            ndt_magnet = entry.get('NDT_자분', 0.0) + entry.get('NDT_흑색자분', 0.0)
            ndt_paint = entry.get('NDT_페인트', 0.0) + entry.get('NDT_백색페인트', 0.0)
            ndt_penetrant = entry.get('NDT_침투제', 0.0)
            ndt_cleaner = entry.get('NDT_세척제', 0.0)
            ndt_developer = entry.get('NDT_현상제', 0.0)
            ndt_fluorescent_pen = entry.get('NDT_형광', 0.0) + entry.get('NDT_형광침투제', 0.0)

            # Cost values
            test_amount = entry.get('검사량', 0.0)
            unit_price = entry.get('단가', 0.0)
            travel_cost = entry.get('출장비', 0.0)
            meal_cost = entry.get('일식', 0.0)
            test_fee = entry.get('검사비', 0.0)
            
            # Accumulate totals
            total_film_count += film_count
            total_test_amount += test_amount
            total_unit_price += unit_price
            total_travel_cost += travel_cost
            total_meal_cost += meal_cost
            total_test_fee += test_fee
            
            total_rtk_center += rtk_center
            total_rtk_density += rtk_density
            total_rtk_marking += rtk_marking
            total_rtk_film += rtk_film
            total_rtk_handling += rtk_handling
            total_rtk_customer += rtk_customer
            total_rtk_other += rtk_other
            total_ndt_fluorescent_mag += ndt_fluorescent_mag
            total_ndt_magnet += ndt_magnet
            total_ndt_paint += ndt_paint
            total_ndt_penetrant += ndt_penetrant
            total_ndt_cleaner += ndt_cleaner
            total_ndt_developer += ndt_developer
            total_ndt_fluorescent_pen += ndt_fluorescent_pen
            
            self.report_tree.insert('', tk.END, values=(
                int(entry['Year']),
                int(entry['Month']),
                entry.get('Site', ''),
                f"{test_amount:.1f}" if test_amount > 0 else '',
                f"{unit_price:,.0f}" if unit_price > 0 else '',
                f"{travel_cost:,.0f}" if travel_cost > 0 else '',
                f"{meal_cost:,.0f}" if meal_cost > 0 else '',
                f"{test_fee:,.0f}" if test_fee > 0 else '',
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
                f"{ndt_fluorescent_mag:.1f}" if ndt_fluorescent_mag > 0 else '',
                f"{ndt_magnet:.1f}" if ndt_magnet > 0 else '',
                f"{ndt_paint:.1f}" if ndt_paint > 0 else '',
                f"{ndt_penetrant:.1f}" if ndt_penetrant > 0 else '',
                f"{ndt_cleaner:.1f}" if ndt_cleaner > 0 else '',
                f"{ndt_developer:.1f}" if ndt_developer > 0 else '',
                f"{ndt_fluorescent_pen:.1f}" if ndt_fluorescent_pen > 0 else '',
                ''  # Empty note field
            ))
        
        # Add total row at the bottom if there's data
        if not grouped.empty:
            total_rtk_sum = total_rtk_center + total_rtk_density + total_rtk_marking + total_rtk_film + total_rtk_handling + total_rtk_customer + total_rtk_other
            
            # Configure tag for total row
            self.report_tree.tag_configure('total', background='#E8F4F8', font=('Arial', 9, 'bold'))
            
            self.report_tree.insert('', tk.END, values=(
                '', # 연도
                '', # 월
                '=== 전체 누계 ===', # 현장
                f"{total_test_amount:.1f}" if total_test_amount > 0 else '',
                f"{total_unit_price:,.0f}" if total_unit_price > 0 else '',
                f"{total_travel_cost:,.0f}" if total_travel_cost > 0 else '',
                f"{total_meal_cost:,.0f}" if total_meal_cost > 0 else '',
                f"{total_test_fee:,.0f}" if total_test_fee > 0 else '',
                '', # 품목명
                f"{total_film_count:.1f}" if total_film_count > 0 else '',
                f"{total_rtk_center:.1f}" if total_rtk_center > 0 else '',
                f"{total_rtk_density:.1f}" if total_rtk_density > 0 else '',
                f"{total_rtk_marking:.1f}" if total_rtk_marking > 0 else '',
                f"{total_rtk_film:.1f}" if total_rtk_film > 0 else '',
                f"{total_rtk_handling:.1f}" if total_rtk_handling > 0 else '',
                f"{total_rtk_customer:.1f}" if total_rtk_customer > 0 else '',
                f"{total_rtk_other:.1f}" if total_rtk_other > 0 else '',
                f"{total_rtk_sum:.1f}" if total_rtk_sum > 0 else '',
                f"{total_ndt_fluorescent_mag:.1f}" if total_ndt_fluorescent_mag > 0 else '',
                f"{total_ndt_magnet:.1f}" if total_ndt_magnet > 0 else '',
                f"{total_ndt_paint:.1f}" if total_ndt_paint > 0 else '',
                f"{total_ndt_penetrant:.1f}" if total_ndt_penetrant > 0 else '',
                f"{total_ndt_cleaner:.1f}" if total_ndt_cleaner > 0 else '',
                f"{total_ndt_developer:.1f}" if total_ndt_developer > 0 else '',
                f"{total_ndt_fluorescent_pen:.1f}" if total_ndt_fluorescent_pen > 0 else '',
                ''
            ), tags=('total',))
    def clean_df_export(self, df):
        """Replace literal 0s with blanks and drop columns that are entirely empty"""
        # 1. Replace 0, 0.0, '0', '0.0' with empty strings
        df = df.replace([0, 0.0, '0', '0.0'], "")
        
        # 2. Identify columns that have at least one non-empty, non-NaN value
        # We convert to string and strip to handle cases that might be whitespace only
        def is_really_empty(col):
            # Check if all values in the column are either NaN or an empty string
            return df[col].astype(str).replace(['nan', 'None', ''], pd.NA).dropna().empty

        non_empty_cols = [col for col in df.columns if not is_really_empty(col)]
        
        # Always keep some essential columns even if empty (safety)
        essential = ['날짜', '현장', '품목명', 'Date', 'Site']
        for col in essential:
            if col in df.columns and col not in non_empty_cols:
                non_empty_cols.append(col)
                
        # Maintain original order
        final_cols = [c for c in df.columns if c in non_empty_cols]
        return df[final_cols]

    def save_df_to_excel_autofit(self, df, save_path, sheet_name='Sheet1'):
        """Save a DataFrame to Excel with automatic column width adjustment (AutoFit)"""
        with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)
            worksheet = writer.sheets[sheet_name]
            
            for idx, col in enumerate(df.columns):
                # Calculate max length of values in column + header
                # We handle Korean characters by assuming they take ~2 units of width
                def get_display_width(s):
                    width = 0
                    for char in str(s):
                        if ord(char) > 127: # Non-ASCII (Korean, etc.)
                            width += 2
                        else:
                            width += 1
                    return width

                series = df[col].astype(str)
                # Filter out empty strings/NAs for max calc
                lengths = series.apply(get_display_width)
                max_val_len = lengths.max() if not lengths.empty else 0
                header_len = get_display_width(col)
                
                # Final width with padding
                final_width = max(max_val_len, header_len) + 2
                
                # Map index to column letter
                # column_letter property is available in openpyxl cells
                col_letter = worksheet.cell(row=1, column=idx+1).column_letter
                worksheet.column_dimensions[col_letter].width = min(final_width, 100) # Cap at 100

    def export_report_to_excel(self):
        """Export current report view to Excel file using dynamic column mapping"""
        try:
            columns = self.report_tree['columns']
            report_data = []
            for item in self.report_tree.get_children():
                report_data.append(self.report_tree.item(item, 'values'))
            
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
                report_df = pd.DataFrame(report_data, columns=columns)
                report_df = self.clean_df_export(report_df)
                self.save_df_to_excel_autofit(report_df, save_path, "보고서")
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
                history_df = self.clean_df_export(history_df)
                self.save_df_to_excel_autofit(history_df, save_path, "입출고내역")
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
        
        # Treeview with columns including cost fields
        columns = ('연도', '월', '현장', '검사량', '단가', '출장비', '일식', '검사비', '품목명', '필름매수', '센터미스', '농도', '마킹미스', '필름마크', '취급부주의', '고객불만', '기타', 'RT총계', '형광자분', '흑색자분', '백색페인트', '침투제', '세척제', '현상제', '형광침투제', '비고')
        self.monthly_usage_tree = ttk.Treeview(tree_frame, columns=columns, show='headings',
                                               yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        vsb.config(command=self.monthly_usage_tree.yview)
        hsb.config(command=self.monthly_usage_tree.xview)
        
        # Column configuration
        col_widths = [80, 60, 120, 100, 100, 100, 100, 100, 200, 100, 70, 70, 70, 70, 70, 70, 70, 70, 80, 80, 80, 80, 80, 80, 80, 200]
        # Columns to center-align (numeric values)
        center_cols = ['검사량', '단가', '출장비', '일식', '검사비', '필름매수', '센터미스', '농도', '마킹미스', '필름마크', '취급부주의', '고객불만', '기타', 'RT총계', '형광자분', '흑색자분', '백색페인트', '침투제', '세척제', '현상제', '형광침투제']
        
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
        
        # Add NDT materials and cost fields
        other_agg_cols = ['NDT_형광자분', 'NDT_자분', 'NDT_흑색자분', 'NDT_페인트', 'NDT_백색페인트', 'NDT_침투제', 'NDT_세척제', 'NDT_현상제', 'NDT_형광', 'NDT_형광침투제',
                          '검사량', '단가', '출장비', '일식', '검사비']
        for col in other_agg_cols:
            if col in df.columns:
                agg_dict[col] = 'sum'
        
        # Group by Year, Month, Site, Material ID and aggregate
        grouped = df.groupby(['Year', 'Month', 'Site', 'Material ID']).agg(agg_dict).reset_index()
        
        # Initialize totals for cumulative sum
        total_film_count = 0.0
        total_test_amount = 0.0
        total_unit_price = 0.0
        total_travel_cost = 0.0
        total_meal_cost = 0.0
        total_test_fee = 0.0
        total_rtk_center = 0.0
        total_rtk_density = 0.0
        total_rtk_marking = 0.0
        total_rtk_film = 0.0
        total_rtk_handling = 0.0
        total_rtk_customer = 0.0
        total_rtk_other = 0.0
        total_ndt_fluorescent_mag = 0.0
        total_ndt_magnet = 0.0
        total_ndt_paint = 0.0
        total_ndt_penetrant = 0.0
        total_ndt_cleaner = 0.0
        total_ndt_developer = 0.0
        total_ndt_fluorescent_pen = 0.0
        
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
            test_amount = entry.get('검사량', 0.0)
            unit_price = entry.get('단가', 0.0)
            travel_cost = entry.get('출장비', 0.0)
            meal_cost = entry.get('일식', 0.0)
            test_fee = entry.get('검사비', 0.0)
            
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
            ndt_fluorescent_mag = entry.get('NDT_형광자분', 0.0)
            ndt_magnet = entry.get('NDT_자분', 0.0) + entry.get('NDT_흑색자분', 0.0)
            ndt_paint = entry.get('NDT_페인트', 0.0) + entry.get('NDT_백색페인트', 0.0)
            ndt_penetrant = entry.get('NDT_침투제', 0.0)
            ndt_cleaner = entry.get('NDT_세척제', 0.0)
            ndt_developer = entry.get('NDT_현상제', 0.0)
            ndt_fluorescent_pen = entry.get('NDT_형광', 0.0) + entry.get('NDT_형광침투제', 0.0)
            
            # Accumulate totals
            total_film_count += film_count
            total_test_amount += test_amount
            total_unit_price += unit_price
            total_travel_cost += travel_cost
            total_meal_cost += meal_cost
            total_test_fee += test_fee
            
            total_rtk_center += rtk_center
            total_rtk_density += rtk_density
            total_rtk_marking += rtk_marking
            total_rtk_film += rtk_film
            total_rtk_handling += rtk_handling
            total_rtk_customer += rtk_customer
            total_rtk_other += rtk_other
            total_ndt_fluorescent_mag += ndt_fluorescent_mag
            total_ndt_magnet += ndt_magnet
            total_ndt_paint += ndt_paint
            total_ndt_penetrant += ndt_penetrant
            total_ndt_cleaner += ndt_cleaner
            total_ndt_developer += ndt_developer
            total_ndt_fluorescent_pen += ndt_fluorescent_pen
            
            self.monthly_usage_tree.insert('', tk.END, values=(
                int(entry['Year']),
                int(entry['Month']),
                entry.get('Site', ''),
                f"{test_amount:.1f}" if test_amount > 0 else '',
                f"{unit_price:,.0f}" if unit_price > 0 else '',
                f"{travel_cost:,.0f}" if travel_cost > 0 else '',
                f"{meal_cost:,.0f}" if meal_cost > 0 else '',
                f"{test_fee:,.0f}" if test_fee > 0 else '',
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
                f"{ndt_fluorescent_mag:.1f}" if ndt_fluorescent_mag > 0 else '',
                f"{ndt_magnet:.1f}" if ndt_magnet > 0 else '',
                f"{ndt_paint:.1f}" if ndt_paint > 0 else '',
                f"{ndt_penetrant:.1f}" if ndt_penetrant > 0 else '',
                f"{ndt_cleaner:.1f}" if ndt_cleaner > 0 else '',
                f"{ndt_developer:.1f}" if ndt_developer > 0 else '',
                f"{ndt_fluorescent_pen:.1f}" if ndt_fluorescent_pen > 0 else '',
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
                '=== 전체 누계 ===',
                f"{total_test_amount:.1f}" if total_test_amount > 0 else '',
                f"{total_unit_price:,.0f}" if total_unit_price > 0 else '',
                f"{total_travel_cost:,.0f}" if total_travel_cost > 0 else '',
                f"{total_meal_cost:,.0f}" if total_meal_cost > 0 else '',
                f"{total_test_fee:,.0f}" if total_test_fee > 0 else '',
                '', # 품목명
                f"{total_film_count:.1f}" if total_film_count > 0 else '',
                f"{total_rtk_center:.1f}" if total_rtk_center > 0 else '',
                f"{total_rtk_density:.1f}" if total_rtk_density > 0 else '',
                f"{total_rtk_marking:.1f}" if total_rtk_marking > 0 else '',
                f"{total_rtk_film:.1f}" if total_rtk_film > 0 else '',
                f"{total_rtk_handling:.1f}" if total_rtk_handling > 0 else '',
                f"{total_rtk_customer:.1f}" if total_rtk_customer > 0 else '',
                f"{total_rtk_other:.1f}" if total_rtk_other > 0 else '',
                f"{total_rtk_sum:.1f}" if total_rtk_sum > 0 else '',
                f"{total_ndt_fluorescent_mag:.1f}" if total_ndt_fluorescent_mag > 0 else '',
                f"{total_ndt_magnet:.1f}" if total_ndt_magnet > 0 else '',
                f"{total_ndt_paint:.1f}" if total_ndt_paint > 0 else '',
                f"{total_ndt_penetrant:.1f}" if total_ndt_penetrant > 0 else '',
                f"{total_ndt_cleaner:.1f}" if total_ndt_cleaner > 0 else '',
                f"{total_ndt_developer:.1f}" if total_ndt_developer > 0 else '',
                f"{total_ndt_fluorescent_pen:.1f}" if total_ndt_fluorescent_pen > 0 else '',
                ''
            ), tags=('total',))

    def export_monthly_usage_history(self):
        """Export monthly usage data displayed in the monthly_usage_tree to Excel using dynamic mapping"""
        try:
            columns = self.monthly_usage_tree['columns']
            monthly_data = []
            for item in self.monthly_usage_tree.get_children():
                monthly_data.append(self.monthly_usage_tree.item(item, 'values'))
            
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
                monthly_df = pd.DataFrame(monthly_data, columns=columns)
                monthly_df = self.clean_df_export(monthly_df)
                self.save_df_to_excel_autofit(monthly_df, save_path, "월별집계")
                messagebox.showinfo("완료", "월별 집계 내역이 저장되었습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"저장 실패: {e}")

    def create_draggable_container(self, parent, label_text, widget_class, config_key, manage_list_key=None, **widget_kwargs):
        """Create a draggable container with a label and a widget"""
        container = ttk.Frame(parent)
        
        # Header container for label and buttons
        hdr = ttk.Frame(container)
        hdr.pack(side='left', padx=(0, 5))
        
        # Label
        lbl = ttk.Label(hdr, text=label_text)
        lbl.pack(side='left')
        
        # Rename icon
        btn_rename = ttk.Label(hdr, text="✏️", font=('Arial', 7), cursor='hand2')
        btn_rename.pack(side='left', padx=(2, 0))
        btn_rename.bind('<Button-1>', lambda e: self.rename_widget_label(config_key))
        
        # Clone icon
        btn_clone = ttk.Label(hdr, text="📋", font=('Arial', 7), cursor='hand2')
        btn_clone.pack(side='left', padx=(2, 0))
        btn_clone.bind('<Button-1>', lambda e: self.clone_widget(config_key))

        # Delete icon (X)
        btn_del = ttk.Label(hdr, text="❌", font=('Arial', 7), cursor='hand2')
        btn_del.pack(side='left', padx=(2, 0))
        btn_del.bind('<Button-1>', lambda e: self.remove_box(config_key))

        # Manage List Icon (Gear) - if it's a list-based widget
        if manage_list_key:
            btn_manage = ttk.Label(hdr, text="⚙️", font=('Arial', 7), cursor='hand2')
            btn_manage.pack(side='left', padx=(2, 0))
            btn_manage.bind('<Button-1>', lambda e: self.open_list_management_dialog(manage_list_key))
        
        # Widget
        widget = widget_class(container, **widget_kwargs)
        widget.pack(side='left')
        
        # Make container draggable
        self.make_draggable(container, config_key)
        
        # Register
        self.draggable_items[config_key] = container
        # Store label widget for renaming and tracking
        container._label_widget = lbl
        container._widget = widget
        container._widget_class = widget_class
        container._widget_kwargs = widget_kwargs
        container._manage_list_key = manage_list_key
        
        return container, widget

    def open_list_management_dialog(self, config_key):
        """Open a generic dialog to manage (edit/delete) items in a data list"""
        if self.layout_locked: return
        
        # Map key to actual list
        data_map = {
            'sites': ('현장 목록 관리', self.sites),
            'users': ('담당자 목록 관리', getattr(self, 'users', [])),
            'equipments': ('장비 목록 관리', getattr(self, 'equipments', []))
        }
        
        if config_key not in data_map: return
        title, data_list = data_map[config_key]
        
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry("300x400")
        dialog.transient(self.root)
        dialog.grab_set()
        
        frame = ttk.Frame(dialog, padding=10)
        frame.pack(fill='both', expand=True)
        
        listbox = tk.Listbox(frame, font=('Arial', 10))
        listbox.pack(fill='both', expand=True, side='left')
        
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=listbox.yview)
        scrollbar.pack(fill='y', side='right')
        listbox.config(yscrollcommand=scrollbar.set)
        
        for item in sorted(data_list):
            listbox.insert('end', item)
            
        btn_frame = ttk.Frame(dialog, padding=5)
        btn_frame.pack(fill='x')
        
        def refresh_list():
            listbox.delete(0, 'end')
            for item in sorted(data_list):
                listbox.insert('end', item)
            # Trigger app-wide update of related comboboxes
            self.refresh_ui_for_list_change(config_key)

        def edit_item():
            sel = listbox.curselection()
            if not sel: return
            idx = sel[0]
            old_val = listbox.get(idx)
            new_val = simpledialog.askstring("수정", "새 이름을 입력하세요:", initialvalue=old_val)
            if new_val and new_val.strip() and new_val != old_val:
                data_list.remove(old_val)
                data_list.append(new_val.strip())
                self.save_tab_config()
                refresh_list()

        def delete_item():
            sel = listbox.curselection()
            if not sel: return
            idx = sel[0]
            val = listbox.get(idx)
            if messagebox.askyesno("삭제 확인", f"'{val}'을 목록에서 삭제하시겠습니까?"):
                data_list.remove(val)
                self.save_tab_config()
                refresh_list()

        ttk.Button(btn_frame, text="수정", command=edit_item).pack(side='left', padx=5, expand=True)
        ttk.Button(btn_frame, text="삭제", command=delete_item).pack(side='left', padx=5, expand=True)
        ttk.Button(btn_frame, text="닫기", command=dialog.destroy).pack(side='left', padx=5, expand=True)

    def refresh_ui_for_list_change(self, config_key):
        """Update all related UI elements after a list (sites, users, etc) has changed"""
        # Dictionary mapping config keys to their current values
        list_map = {
            'sites': self.sites,
            'users': self.users,
            'warehouses': self.warehouses,
            'equipments': self.equipments
        }
        
        if config_key not in list_map:
            return
            
        current_vals = list_map[config_key]
        sorted_vals = sorted(current_vals)
        
        # 1. Update standard widgets
        if config_key == 'sites':
            if hasattr(self, 'ent_daily_site'): self.ent_daily_site['values'] = sorted_vals
            if hasattr(self, 'cb_trans_site'): self.cb_trans_site['values'] = sorted_vals
            if hasattr(self, 'cb_daily_filter_site'):
                self.cb_daily_filter_site['values'] = ['전체'] + sorted_vals
        elif config_key == 'users':
            for i in range(1, 7):
                attr = 'cb_daily_user' if i == 1 else f'cb_daily_user{i}'
                if hasattr(self, attr):
                    try: getattr(self, attr)['values'] = sorted_vals
                    except: pass
            if hasattr(self, 'ent_user') and hasattr(self.ent_user, 'config'):
                try: self.ent_user['values'] = sorted_vals
                except: pass
        elif config_key == 'equipments':
            self.update_material_combo()

        # 2. Update ALL draggable widgets (clones) that depend on this list
        for key, container in self.draggable_items.items():
            # Heuristic: if manage_list_key is missing but label suggests it's a worker/user/site
            m_key = getattr(container, '_manage_list_key', None)
            if not m_key:
                if hasattr(container, '_label_widget'):
                    lbl_text = container._label_widget.cget('text').lower()
                    if config_key == 'users' and any(x in lbl_text for x in ['작업자', '담당자', 'user', 'worker']):
                        m_key = 'users'
                        container._manage_list_key = 'users'
                    elif config_key == 'sites' and any(x in lbl_text for x in ['현장', 'site']):
                        m_key = 'sites'
                        container._manage_list_key = 'sites'
                    elif config_key == 'equipments' and any(x in lbl_text for x in ['장비', 'equip']):
                        m_key = 'equipments'
                        container._manage_list_key = 'equipments'

            if m_key == config_key:
                if hasattr(container, '_widget') and hasattr(container._widget, 'config'):
                    try:
                        container._widget['values'] = sorted_vals
                    except:
                        pass

    def remove_box(self, key):
        """Intelligently remove a box: hide standard ones, destroy custom ones"""
        if self.layout_locked: return
        
        if key.startswith('memo_') or key.startswith('clone_'):
            # Permanent deletion for custom/dynamic items
            self.destroy_custom_widget(key)
        else:
            # Hiding for standard items (can be restored via Reset All)
            widget = self.draggable_items.get(key)
            if widget:
                # We reuse the existing hide_widget logic
                self.hide_widget(None, widget=widget)

    def destroy_custom_widget(self, key):
        """Destroy and remove from config any dynamic widget (clone or memo)"""
        if self.layout_locked: return
        
        widget = self.draggable_items.get(key)
        if widget:
            widget.destroy()
            if key in self.draggable_items:
                del self.draggable_items[key]
            if key in self.memos:
                del self.memos[key]
            
            # Clean from config immediately to prevent resurrection on next load
            try:
                import json
                if os.path.exists(self.config_path):
                    with open(self.config_path, 'r', encoding='utf-8') as f:
                        config = json.load(f)
                    
                    if 'draggable_geometries' in config and key in config['draggable_geometries']:
                        del config['draggable_geometries'][key]
                        
                    with open(self.config_path, 'w', encoding='utf-8') as f:
                        json.dump(config, f, ensure_ascii=False, indent=2)
            except:
                pass
            
            self.save_tab_config()

    def rename_widget_label(self, key):
        """Show a simple dialog to rename a widget's label"""
        if self.layout_locked: return
        
        widget = self.draggable_items.get(key)
        if not widget: return
        
        current_text = ""
        if hasattr(widget, '_label_widget'):
            current_text = widget._label_widget.cget('text')
        elif key in self.memos:
            current_text = self.memos[key]['title_entry'].get()
            
        new_name = simpledialog.askstring("이름 변경", "새 이름을 입력하세요:", initialvalue=current_text)
        if new_name is not None:
            if hasattr(widget, '_label_widget'):
                widget._label_widget.config(text=new_name)
            elif key in self.memos:
                self.memos[key]['title_entry'].delete(0, 'end')
                self.memos[key]['title_entry'].insert(0, new_name)
            self.save_tab_config()

    def clone_widget(self, key):
        """Create a clone of an existing widget as a new custom box"""
        if self.layout_locked: return
        
        orig = self.draggable_items.get(key)
        if not orig: return
        
        import time
        new_key = f"clone_{int(time.time() * 1000)}"
        
        label_text = ""
        if hasattr(orig, '_label_widget'):
            label_text = orig._label_widget.cget('text')
        
        if hasattr(orig, '_widget_class'):
            # It's a container created via create_draggable_container
            cont, w = self.create_draggable_container(
                self.entry_inner_frame, 
                label_text, 
                orig._widget_class, 
                new_key, 
                manage_list_key=getattr(orig, '_manage_list_key', None), # Pass manage_list_key
                **orig._widget_kwargs
            )
            cont.place(x=50, y=50) # Start position
            self.save_tab_config()
        elif key in self.memos:
            # It's a memo
            content = self.memos[key]['text_widget'].get('1.0', 'end-1c')
            title = self.memos[key]['title_entry'].get()
            self.add_new_memo(initial_text=content, initial_title=title, key=new_key)
            self.save_tab_config()

    def _bind_recursive(self, widget, target_container):
        """Recursively bind drag events to widget and its children"""
        from functools import partial
        
        # Bind events to this widget, targeting the container
        # Note: We use add=True to avoid overwriting existing bindings if possible, 
        # but for drag we usually want exclusive control or at least priority.
        # Here we just bind standard.
        
        # We need to capture the target_container in the callback
        widget.bind("<Button-3>", partial(self.on_drag_start, widget=target_container))
        widget.bind("<Shift-Button-3>", partial(self.on_resize_start, widget=target_container))
        widget.bind("<B3-Motion>", partial(self.on_mouse_motion, widget=target_container))
        widget.bind("<Double-Button-3>", partial(self.reset_widget_position, widget=target_container))
        widget.bind("<Control-Button-3>", partial(self.hide_widget, widget=target_container))
        widget.bind("<ButtonRelease-3>", partial(self.on_drag_stop, widget=target_container))
        
        # Recurse for children
        try:
            for child in widget.winfo_children():
                self._bind_recursive(child, target_container)
        except:
            pass

    def hide_widget(self, event, widget=None):
        """Hide a widget (Ctrl + Right Click)"""
        if self.layout_locked:
            return "break"
            
        if widget is None:
            widget = event.widget
        
        # Remove from place layout
        if widget.winfo_manager() == 'place':
            widget.place_forget()
        elif widget.winfo_manager() == 'grid':
            widget.grid_forget()
            
        # Remove placeholder if exists
        self._remove_placeholder(widget)
        
        # Mark as hidden in config
        if hasattr(widget, '_config_key') and widget._config_key:
            try:
                import json
                if os.path.exists(self.config_path):
                    with open(self.config_path, 'r', encoding='utf-8') as f:
                        config = json.load(f)
                else:
                    config = {}
                
                if 'draggable_geometries' not in config:
                    config['draggable_geometries'] = {}
                
                if widget._config_key not in config['draggable_geometries']:
                    config['draggable_geometries'][widget._config_key] = {}
                
                config['draggable_geometries'][widget._config_key]['hidden'] = True
                
                with open(self.config_path, 'w', encoding='utf-8') as f:
                    json.dump(config, f, ensure_ascii=False, indent=2)
            except Exception as e:
                print(f"Failed to save hide status: {e}")
        return "break"

    def make_draggable(self, widget, config_key=None):
        """Make a widget draggable and resizable with right mouse button"""
        widget._config_key = config_key
        # Save original grid info for reset/placeholder
        # We need to do this immediately while it's still in the grid
        widget._original_grid_info = widget.grid_info()
        
        # Recursively bind to ensure clicking anywhere works
        self._bind_recursive(widget, widget)
        
    def reset_widget_position(self, event, widget=None):
        """Reset widget to original grid position"""
        if self.layout_locked:
            return "break"
            
        if widget is None:
            widget = event.widget
        
        # Remove from place layout
        widget.place_forget()
        
        # Remove placeholder if exists
        self._remove_placeholder(widget)
            
        # Restore to grid
        if hasattr(widget, '_original_grid_info'):
            widget.grid(**widget._original_grid_info)
        
        # Reset size variables if any
        if hasattr(widget, '_start_width'): del widget._start_width
        if hasattr(widget, '_start_height'): del widget._start_height
        
        # Remove from config
        if hasattr(widget, '_config_key') and widget._config_key:
            try:
                import json
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                
                if 'draggable_geometries' in config and widget._config_key in config['draggable_geometries']:
                    del config['draggable_geometries'][widget._config_key]
                
                # Backward compatibility: also check top level
                if widget._config_key in config:
                    del config[widget._config_key]
                    
                with open(self.config_path, 'w', encoding='utf-8') as f:
                    json.dump(config, f, ensure_ascii=False, indent=2)
                    
            except Exception as e:
                print(f"Failed to reset config: {e}")

    def reset_all_widgets_layout(self):
        """Reset all widgets to their original grid slots, unhide them, and reset paned window sashes"""
        if messagebox.askyesno("초기화", "모든 항목의 위치와 크기를 초기화하고 숨겨진 항목을 다시 표시하시겠습니까?\n(창 분할 위치도 초기화됩니다.)"):
            # 1. Reset draggable widgets
            for key, widget in list(self.draggable_items.items()):
                self.reset_widget_position(None, widget=widget)
            
            # 2. Reset sash positions (splitters)
            # We use half/half or sensible defaults
            try:
                if hasattr(self, 'daily_usage_paned'):
                    self.daily_usage_paned.sashpos(0, 300) # Default vertical split
                if hasattr(self, 'daily_history_paned'):
                    # For horizontal split, we want the list to be larger
                    total_w = self.daily_history_paned.winfo_width()
                    if total_w > 100:
                        self.daily_history_paned.sashpos(0, int(total_w * 0.7))
                    else:
                        self.daily_history_paned.sashpos(0, 800)
            except:
                pass

            self.save_tab_config()
            messagebox.showinfo("완료", "레이아웃과 창 분할 위치가 초기화되었습니다.")
            
    def toggle_layout_lock(self):
        """Toggle layout locking state"""
        self.layout_locked = not self.layout_locked
        if hasattr(self, 'btn_lock_layout'):
            if self.layout_locked:
                self.btn_lock_layout.config(text="🔒 배치 고정됨")
                self.style.configure("Lock.TButton", foreground="black")
            else:
                self.btn_lock_layout.config(text="🔓 배치 수정 중")
                self.style.configure("Lock.TButton", foreground="red")
        
        # Save lock state
        self.save_tab_config()

    def on_drag_stop(self, event, widget=None):
        """Handle end of dragging or resizing and auto-save"""
        if widget is None:
            widget = event.widget
        if hasattr(widget, '_interaction_mode'):
            del widget._interaction_mode
            # Auto-save layout
            self.save_tab_config()

    def add_new_memo(self, initial_text="", initial_title="메모", key=None):
        """Create a new movable/editable/copyable memo box with editable title"""
        import time
        if key is None:
            key = f"memo_{int(time.time() * 1000)}"
        
        # Create container
        memo_container = ttk.LabelFrame(self.entry_inner_frame)
        
        # Controls frame (top of memo)
        ctrl_frame = ttk.Frame(memo_container)
        ctrl_frame.pack(fill='x', side='top')
        
        # Editable Title Entry
        title_entry = ttk.Entry(ctrl_frame, font=('Arial', 9, 'bold'), width=15)
        title_entry.pack(side='left', padx=2)
        title_entry.insert(0, initial_title)
        title_entry.bind('<FocusOut>', lambda e: self.save_tab_config())
        
        btn_copy = ttk.Button(ctrl_frame, text="📋", width=3, command=lambda: self.duplicate_memo(key))
        btn_copy.pack(side='right')
        
        btn_del = ttk.Button(ctrl_frame, text="❌", width=3, command=lambda: self.remove_box(key))
        btn_del.pack(side='right')
        
        # Text area
        text_area = tk.Text(memo_container, wrap='word', height=5, width=30, font=('Arial', 10))
        text_area.pack(fill='both', expand=True, padx=2, pady=2)
        text_area.insert('1.0', initial_text)
        
        # Bind text change to auto-save
        text_area.bind('<FocusOut>', lambda e: self.save_tab_config())
        
        # Make draggable
        self.make_draggable(memo_container, key)
        self.draggable_items[key] = memo_container
        self.memos[key] = {
            'container': memo_container, 
            'text_widget': text_area,
            'title_entry': title_entry
        }
        
        # Initial placement if not loaded from config (will be handled by load_tab_config if exists)
        if key not in getattr(self, '_loading_memos', []):
            memo_container.place(x=50, y=50) # Default start position
        
        return memo_container

    def duplicate_memo(self, key):
        """Duplicate an existing memo with its title and content"""
        if self.layout_locked: return
        if key in self.memos:
            content = self.memos[key]['text_widget'].get('1.0', 'end-1c')
            title = self.memos[key]['title_entry'].get()
            self.add_new_memo(initial_text=content, initial_title=title)
            self.save_tab_config()


    def on_drag_start(self, event, widget=None):
        """Begin dragging widget"""
        if self.layout_locked:
            return "break" # Prevent movement and stop propagation
            
        if widget is None:
            widget = event.widget
        widget._interaction_mode = 'move'
        
        # Save absolute start position of mouse
        widget._drag_start_root_x = event.x_root
        widget._drag_start_root_y = event.y_root
        
        # Save initial widget position relative to parent
        widget._drag_start_pos_x = widget.winfo_x()
        widget._drag_start_pos_y = widget.winfo_y()
        
        # Ensure we have grid info (redundant but safe)
        if not hasattr(widget, '_original_grid_info') and widget.grid_info():
            widget._original_grid_info = widget.grid_info()
        
    def on_resize_start(self, event, widget=None):
        """Begin resizing widget"""
        if self.layout_locked:
            return "break"
            
        if widget is None:
            widget = event.widget
        widget._interaction_mode = 'resize'
        
        # Save absolute start position of mouse
        widget._drag_start_root_x = event.x_root
        widget._drag_start_root_y = event.y_root
        
        # Save initial size
        widget._start_width = widget.winfo_width()
        widget._start_height = widget.winfo_height()
        
        # Ensure we have grid info
        if not hasattr(widget, '_original_grid_info') and widget.grid_info():
            widget._original_grid_info = widget.grid_info()
        return "break"
        
    def on_mouse_motion(self, event, widget=None):
        """Handle dragging or resizing motion"""
        if self.layout_locked:
            return "break"
            
        if widget is None:
            widget = event.widget
        
        if not hasattr(widget, '_interaction_mode'):
            return
        
        # Calculate TOTAL delta from start
        dx = event.x_root - widget._drag_start_root_x
        dy = event.y_root - widget._drag_start_root_y
        
        # Get parent dimensions for clamping
        parent = widget.master
        parent_w = parent.winfo_width()
        parent_h = parent.winfo_height()
        
        if widget._interaction_mode == 'move':
            # Calculate new position based on INITIAL position + delta
            # This prevents accumulation errors
            new_x = widget._drag_start_pos_x + dx
            new_y = widget._drag_start_pos_y + dy
            
            # Constrain to parent bounds
            widget_w = widget.winfo_width()
            widget_h = widget.winfo_height()
            
            # Clamp X
            if new_x < 0: 
                new_x = 0
            elif new_x + widget_w > parent_w:
                new_x = max(0, parent_w - widget_w)
                
            # Clamp Y
            if new_y < 0:
                new_y = 0
            elif new_y + widget_h > parent_h:
                new_y = max(0, parent_h - widget_h)
            
            if widget.winfo_manager() != 'place':
                # Switching to absolute positioning
                self._ensure_placeholder(widget)
                widget.lift()
                widget.place(x=widget._drag_start_pos_x, y=widget._drag_start_pos_y, width=widget_w, height=widget_h)
                
            widget.place(x=new_x, y=new_y)
            
        elif widget._interaction_mode == 'resize':
            new_width = max(50, widget._start_width + dx)
            new_height = max(20, widget._start_height + dy)
            
            # Constrain resize to parent bounds
            current_x = widget.winfo_x()
            current_y = widget.winfo_y()
            
            if current_x + new_width > parent_w:
                new_width = max(50, parent_w - current_x)
            
            if current_y + new_height > parent_h:
                new_height = max(20, parent_h - current_y)
            
            if widget.winfo_manager() != 'place':
                # Switching to absolute positioning
                self._ensure_placeholder(widget)
                widget.lift()
                widget.place(x=widget.winfo_x(), y=widget.winfo_y(), width=widget.winfo_width(), height=widget.winfo_height())
                
            widget.place(width=new_width, height=new_height)
        
        return "break"

    def _ensure_placeholder(self, widget, width=None, height=None):
        """Ensure a placeholder exists in the grid where the widget used to be"""
        # Map widget to a placeholder attribute name dynamically
        # We can use the widget id or a dictionary, but simpler to attach to widget
        if not hasattr(widget, '_placeholder'):
            # Use provided dims or current widget dims
            w = width if width is not None else widget.winfo_width()
            h = height if height is not None else widget.winfo_height()
            
            # Create a frame to hold the space
            widget._placeholder = ttk.Frame(widget.master, width=w, height=h)
            
            # Grid it at the original position
            if hasattr(widget, '_original_grid_info'):
                widget._placeholder.grid(**widget._original_grid_info)
                # Ensure the placeholder doesn't shrink and holds its size
                widget._placeholder.grid_propagate(False)
                
    def _remove_placeholder(self, widget):
        """Remove placeholder for a widget"""
        if hasattr(widget, '_placeholder'):
            widget._placeholder.destroy()
            del widget._placeholder

        

    def setup_daily_usage_tab(self):
        """Setup the daily usage entry tab"""
        # Top frame for entry form
        # We use a canvas or large frame to allow free movement? 
        # Actually, we keep the entry_frame as the parent but use grid for initial layout
        # The user can then move them out of grid into place
        
        # Create PanedWindow for resizable frames
        self.daily_usage_paned = ttk.Panedwindow(self.tab_daily_usage, orient='vertical')
        self.daily_usage_paned.pack(fill='both', expand=True, padx=10, pady=10)
        
        entry_frame = ttk.LabelFrame(self.daily_usage_paned, text="현장별 일일 사용량 기입")
        self.daily_usage_paned.add(entry_frame, weight=1)
        
        # Header for buttons to keep them separate from draggable area
        header_frame = ttk.Frame(entry_frame)
        header_frame.pack(fill='x', padx=5, pady=2)
        
        btn_reset_all = ttk.Button(header_frame, text="전체 레이아웃 초기화", command=self.reset_all_widgets_layout)
        btn_reset_all.pack(side='right', padx=5)
        
        self.btn_lock_layout = ttk.Button(header_frame, text="🔓 배치 수정 중", command=self.toggle_layout_lock, style="Lock.TButton")
        self.btn_lock_layout.pack(side='right', padx=5)
        self.style.configure("Lock.TButton", foreground="red")
        
        btn_add_memo = ttk.Button(header_frame, text="➕ 메모 추가", command=self.add_new_memo)
        btn_add_memo.pack(side='right', padx=5)

        # Stable inner frame for draggable widgets - 0,0 is now truly top-left of content
        self.entry_inner_frame = ttk.Frame(entry_frame)
        self.entry_inner_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        # 1. Date
        # Special case because DateEntry needs specific init
        date_container = ttk.Frame(self.entry_inner_frame)
        ttk.Label(date_container, text="날짜:").pack(side='left', padx=(0, 5))
        self.ent_daily_date = DateEntry(date_container, width=20, background='darkblue',
                                        foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
        self.ent_daily_date.pack(side='left')
        self.ent_daily_date.set_date(datetime.datetime.now())
        
        self.make_draggable(date_container, 'date_box_geometry')
        self.draggable_items['date_box_geometry'] = date_container
        
        # Initial grid pos
        date_container.grid(row=0, column=0, padx=5, pady=5, sticky='w')
        
        # 2. Site
        # Special case for site because of the delete button
        site_container = ttk.Frame(self.entry_inner_frame)
        
        # Site header buttons (Refactored to match draggable container style if possible, 
        # but since Site has special logic, we manually add icons)
        site_hdr = ttk.Frame(site_container)
        site_hdr.pack(side='left', padx=(0, 5))
        
        ttk.Label(site_hdr, text="현장:").pack(side='left')
        
        # Rename icon
        btn_rename_site = ttk.Label(site_hdr, text="✏️", font=('Arial', 7), cursor='hand2')
        btn_rename_site.pack(side='left', padx=(2, 0))
        btn_rename_site.bind('<Button-1>', lambda e: self.rename_widget_label('site_box_geometry'))
        
        # Clone icon
        btn_clone_site = ttk.Label(site_hdr, text="📋", font=('Arial', 7), cursor='hand2')
        btn_clone_site.pack(side='left', padx=(2, 0))
        btn_clone_site.bind('<Button-1>', lambda e: self.clone_widget('site_box_geometry'))

        # Manage List Icon (Gear)
        btn_manage_site = ttk.Label(site_hdr, text="⚙️", font=('Arial', 7), cursor='hand2')
        btn_manage_site.pack(side='left', padx=(2, 0))
        btn_manage_site.bind('<Button-1>', lambda e: self.open_list_management_dialog('sites'))

        # Delete icon
        btn_del_site = ttk.Label(site_hdr, text="❌", font=('Arial', 7), cursor='hand2')
        btn_del_site.pack(side='left', padx=(2, 0))
        btn_del_site.bind('<Button-1>', lambda e: self.remove_box('site_box_geometry'))

        self.ent_daily_site = ttk.Combobox(site_container, width=30, values=self.sites)
        self.ent_daily_site.pack(side='left')
        self.ent_daily_site.bind('<FocusOut>', lambda e: self.auto_save_to_list(e, self.ent_daily_site, self.sites, 'sites'))
        self.ent_daily_site.bind('<Return>', lambda e: self.auto_save_to_list(e, self.ent_daily_site, self.sites, 'sites'))
        
        self.make_draggable(site_container, 'site_box_geometry')
        self.draggable_items['site_box_geometry'] = site_container
        self.daily_site_frame = site_container
        site_container._label_widget = site_hdr.winfo_children()[0] # The Label widget
        site_container._widget = self.ent_daily_site # Store the actual widget
        site_container._manage_list_key = 'sites' # Track the list it manages
        
        site_container.grid(row=1, column=0, padx=5, pady=5, sticky='w')

        # 3. Worker 1
        user_container, self.cb_daily_user = self.create_draggable_container(
            self.entry_inner_frame, "작업자1:", ttk.Combobox, 'user_box_geometry', 
            manage_list_key='users', width=18, values=getattr(self, 'users', [])
        )
        self.cb_daily_user.bind('<FocusOut>', lambda e: self.auto_save_to_list(e, self.cb_daily_user, self.users, 'users'))
        self.cb_daily_user.bind('<Return>', lambda e: self.auto_save_to_list(e, self.cb_daily_user, self.users, 'users'))
        user_container.grid(row=1, column=1, padx=5, pady=5, sticky='w')

        # Worker 2
        user2_container, self.cb_daily_user2 = self.create_draggable_container(
            self.entry_inner_frame, "작업자2:", ttk.Combobox, 'user2_box_geometry', 
            manage_list_key='users', width=18, values=getattr(self, 'users', [])
        )
        self.cb_daily_user2.bind('<FocusOut>', lambda e: self.auto_save_to_list(e, self.cb_daily_user2, self.users, 'users'))
        self.cb_daily_user2.bind('<Return>', lambda e: self.auto_save_to_list(e, self.cb_daily_user2, self.users, 'users'))
        user2_container.grid(row=1, column=2, padx=5, pady=5, sticky='w')

        # Worker 3
        user3_container, self.cb_daily_user3 = self.create_draggable_container(
            self.entry_inner_frame, "작업자3:", ttk.Combobox, 'user3_box_geometry', 
            manage_list_key='users', width=18, values=getattr(self, 'users', [])
        )
        self.cb_daily_user3.bind('<FocusOut>', lambda e: self.auto_save_to_list(e, self.cb_daily_user3, self.users, 'users'))
        self.cb_daily_user3.bind('<Return>', lambda e: self.auto_save_to_list(e, self.cb_daily_user3, self.users, 'users'))
        user3_container.grid(row=1, column=3, padx=5, pady=5, sticky='w')

        # Worker 4
        user4_container, self.cb_daily_user4 = self.create_draggable_container(
            self.entry_inner_frame, "작업자4:", ttk.Combobox, 'user4_box_geometry', 
            manage_list_key='users', width=18, values=getattr(self, 'users', [])
        )
        self.cb_daily_user4.bind('<FocusOut>', lambda e: self.auto_save_to_list(e, self.cb_daily_user4, self.users, 'users'))
        self.cb_daily_user4.bind('<Return>', lambda e: self.auto_save_to_list(e, self.cb_daily_user4, self.users, 'users'))
        user4_container.grid(row=2, column=2, padx=5, pady=5, sticky='w')

        # Worker 5
        user5_container, self.cb_daily_user5 = self.create_draggable_container(
            self.entry_inner_frame, "작업자5:", ttk.Combobox, 'user5_box_geometry', 
            manage_list_key='users', width=18, values=getattr(self, 'users', [])
        )
        self.cb_daily_user5.bind('<FocusOut>', lambda e: self.auto_save_to_list(e, self.cb_daily_user5, self.users, 'users'))
        self.cb_daily_user5.bind('<Return>', lambda e: self.auto_save_to_list(e, self.cb_daily_user5, self.users, 'users'))
        user5_container.grid(row=2, column=3, padx=5, pady=5, sticky='w')

        # Worker 6
        user6_container, self.cb_daily_user6 = self.create_draggable_container(
            self.entry_inner_frame, "작업자6:", ttk.Combobox, 'user6_box_geometry', 
            manage_list_key='users', width=18, values=getattr(self, 'users', [])
        )
        self.cb_daily_user6.bind('<FocusOut>', lambda e: self.auto_save_to_list(e, self.cb_daily_user6, self.users, 'users'))
        self.cb_daily_user6.bind('<Return>', lambda e: self.auto_save_to_list(e, self.cb_daily_user6, self.users, 'users'))
        user6_container.grid(row=3, column=2, padx=5, pady=5, sticky='w')

        # 4. Equipment
        equip_container, self.cb_daily_equip = self.create_draggable_container(
            self.entry_inner_frame, "장비명:", ttk.Combobox, 'equip_box_geometry', 
            manage_list_key='equipments', width=18, values=getattr(self, 'equipments', [])
        )
        self.cb_daily_equip.bind('<FocusOut>', lambda e: self.auto_save_to_list(e, self.cb_daily_equip, self.equipments, 'equipments'))
        self.cb_daily_equip.bind('<Return>', lambda e: self.auto_save_to_list(e, self.cb_daily_equip, self.equipments, 'equipments'))
        equip_container.grid(row=2, column=0, padx=5, pady=5, sticky='w')

        # 5. Material
        mat_container, self.cb_daily_material = self.create_draggable_container(
            self.entry_inner_frame, "품목명:", ttk.Combobox, 'mat_box_geometry', width=35
        )
        mat_container.grid(row=2, column=1, padx=5, pady=5, sticky='w')
        
        # 6. Test Method
        method_container, self.cb_daily_test_method = self.create_draggable_container(
            self.entry_inner_frame, "검사방법:", ttk.Combobox, 'method_box_geometry', width=10, values=["PAUT", "UT", "MT", "PT", "PMI"]
        )
        method_container.grid(row=3, column=0, padx=5, pady=5, sticky='w')

        # 7. Test Amount
        amount_container, self.ent_daily_test_amount = self.create_draggable_container(
            self.entry_inner_frame, "검사량:", ttk.Entry, 'amount_box_geometry', width=10
        )
        self.ent_daily_test_amount.insert(0, "0")
        amount_container.grid(row=3, column=1, padx=5, pady=5, sticky='w')
        
        # 8. Costs (Unit Price, Travel, Meal, Test Fee)
        # Unit Price
        u_price_container, self.ent_daily_unit_price = self.create_draggable_container(
            self.entry_inner_frame, "단가:", ttk.Entry, 'u_price_box_geometry', width=12
        )
        self.ent_daily_unit_price.insert(0, "0")
        u_price_container.grid(row=4, column=0, padx=5, pady=5, sticky='w')
        
        # Travel Cost
        travel_container, self.ent_daily_travel_cost = self.create_draggable_container(
            self.entry_inner_frame, "출장비:", ttk.Entry, 'travel_box_geometry', width=12
        )
        self.ent_daily_travel_cost.insert(0, "0")
        travel_container.grid(row=4, column=1, padx=5, pady=5, sticky='w')

        # Meal Cost
        meal_container, self.ent_daily_meal_cost = self.create_draggable_container(
            self.entry_inner_frame, "일식:", ttk.Entry, 'meal_box_geometry', width=12
        )
        self.ent_daily_meal_cost.insert(0, "0")
        meal_container.grid(row=5, column=0, padx=5, pady=5, sticky='w')
        
        # Test Fee
        fee_container, self.ent_daily_test_fee = self.create_draggable_container(
            self.entry_inner_frame, "검사비:", ttk.Entry, 'fee_box_geometry', width=12
        )
        self.ent_daily_test_fee.insert(0, "0")
        fee_container.grid(row=5, column=1, padx=5, pady=5, sticky='w')
        


        # 9. Film count
        film_container, self.ent_film_count = self.create_draggable_container(
            self.entry_inner_frame, "필름매수:", ttk.Entry, 'film_box_geometry', width=20
        )
        self.ent_film_count.insert(0, "0")
        film_container.grid(row=5, column=2, padx=5, pady=5, sticky='w')

        # Add calculation bindings
        calc_trigger = lambda e: self.update_daily_test_fee_calc()
        self.ent_daily_test_amount.bind('<KeyRelease>', calc_trigger)
        self.ent_daily_unit_price.bind('<KeyRelease>', calc_trigger)
        self.ent_daily_travel_cost.bind('<KeyRelease>', calc_trigger)
        self.ent_daily_meal_cost.bind('<KeyRelease>', calc_trigger)
        
        # NDT Materials Usage Frame
        ndt_frame = ttk.LabelFrame(self.entry_inner_frame, text="NDT 자재 사용량")
        ndt_frame.grid(row=6, column=0, columnspan=3, padx=5, pady=5, sticky='ew')
        
        # Create entry fields for NDT materials
        self.ndt_entries = {}
        ndt_materials = ["형광자분", "흑색자분", "백색페인트", "침투제", "세척제", "현상제", "형광침투제"]
        
        for i, material in enumerate(ndt_materials):
            r = i // 4
            c = (i % 4) * 2
            ttk.Label(ndt_frame, text=f"{material}:").grid(row=r, column=c, padx=5, pady=2, sticky='w')
            entry = ttk.Entry(ndt_frame, width=8)
            entry.grid(row=r, column=c+1, padx=5, pady=2)
            self.ndt_entries[material] = entry
        
        # RT 매수 (RTK Usage amounts - separate for each category)
        rtk_frame = ttk.LabelFrame(self.entry_inner_frame, text="RT 매수 (RTK별 사용량)")
        rtk_frame.grid(row=7, column=0, columnspan=3, padx=5, pady=5, sticky='ew')
        
        # Create entry fields for each RTK category
        self.rtk_entries = {}
        rtk_categories = ["센터미스", "농도", "마킹미스", "필름마크", "취급부주의", "고객불만", "기타", "총계"]
        
        for i, category in enumerate(rtk_categories):
            r = i // 4
            c = (i % 4) * 2
            ttk.Label(rtk_frame, text=f"{category}:").grid(row=r, column=c, padx=5, pady=2, sticky='w')
            entry = ttk.Entry(rtk_frame, width=10)
            entry.grid(row=r, column=c+1, padx=5, pady=2)
            self.rtk_entries[category] = entry
            
            # Bind change event to calculate total
            if category != "총계":
                entry.bind('<KeyRelease>', lambda e, cat=category: self.calculate_rtk_total())
                entry.bind('<FocusOut>', lambda e, cat=category: self.calculate_rtk_total())
        
        # Make total entry readonly
        self.rtk_entries["총계"].config(state='readonly', background='lightgray')
        
        # 10. Note
        note_container, self.ent_daily_note = self.create_draggable_container(
            self.entry_inner_frame, "비고:", ttk.Entry, 'note_box_geometry', width=100
        )
        note_container.grid(row=8, column=1, columnspan=2, padx=5, pady=5, sticky='ew')
        
        # Save button
        save_btn_container = ttk.Frame(self.entry_inner_frame)
        btn_save_daily = ttk.Button(save_btn_container, text="기록 저장", command=self.add_daily_usage_entry)
        btn_save_daily.pack()
        
        self.make_draggable(save_btn_container, 'save_btn_geometry')
        self.draggable_items['save_btn_geometry'] = save_btn_container
        
        save_btn_container.grid(row=9, column=0, columnspan=3, pady=10)
        
        # Update material and equipment comboboxes
        self.update_material_combo()
        
        # Bottom frame for display
        display_frame = ttk.LabelFrame(self.daily_usage_paned, text="일일 사용량 기록 조회")
        self.daily_usage_paned.add(display_frame, weight=2)
        
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
        tree_container = ttk.Frame(display_frame)
        tree_container.pack(expand=True, fill='both', padx=5, pady=5)
        
        # Horizontal PanedWindow for List vs Details
        self.daily_history_paned = ttk.Panedwindow(tree_container, orient='horizontal')
        self.daily_history_paned.pack(fill='both', expand=True)

        list_frame = ttk.Frame(self.daily_history_paned)
        self.daily_history_paned.add(list_frame, weight=3)

        # Scrollbars
        vsb = ttk.Scrollbar(list_frame, orient="vertical")
        hsb = ttk.Scrollbar(list_frame, orient="horizontal")
        
        # Treeview with RTK categories and NDT materials
        # Note: Workers 1-6 columns are kept in the 'columns' tuple for data storage,
        # but we will only show a consolidated '작업자' in 'displaycolumns'.
        columns = ('날짜', '현장', '작업자', '장비명', '검사방법', '검사량', '필름매수', '단가', '출장비', '일식', '검사비', '품목명', '센터미스', '농도', '마킹미스', '필름마크', '취급부주의', '고객불만', '기타', 'RT총계', '형광자분', '흑색자분', '백색페인트', '침투제', '세척제', '현상제', '형광침투제', '비고', '입력시간')
        self.daily_usage_tree = ttk.Treeview(list_frame, columns=columns, show='headings',
                                              yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        vsb.config(command=self.daily_usage_tree.yview)
        hsb.config(command=self.daily_usage_tree.xview)
        
        # Column configuration
        # Column configuration
        col_widths = [100, 120, 100, 100, 100, 100, 100, 100, 120, 100, 100, 100, 100, 100, 100, 100, 200, 70, 70, 70, 70, 70, 70, 70, 70, 80, 80, 80, 80, 80, 80, 80, 200, 150]
        # Columns to center-align (numeric values)
        center_cols = ['검사방법', '검사량', '필름매수', '단가', '출장비', '일식', '검사비', '센터미스', '농도', '마킹미스', '필름마크', '취급부주의', '고객불만', '기타', 'RT총계', '형광자분', '흑색자분', '백색페인트', '침투제', '세척제', '현상제', '형광침투제']
        
        # Column configuration
        # '작업자' replaces '작업자1'...'작업자6' for primary display
        col_widths = {
            '날짜': 100, '현장': 120, '작업자': 150, '장비명': 120, '검사방법': 80, 
            '검사량': 70, '필름매수': 70, '단가': 80, '출장비': 80, '일식': 70, 
            '검사비': 90, '품목명': 180, '센터미스': 60, '농도': 60, '마킹미스': 60, 
            '필름마크': 60, '취급부주의': 60, '고객불만': 60, '기타': 60, 'RT총계': 70, 
            '형광자분': 70, '흑색자분': 70, '백색페인트': 70, '침투제': 70, '세척제': 70, 
            '현상제': 70, '형광침투제': 70, '비고': 200, '입력시간': 130
        }
        
        for col in columns:
            self.daily_usage_tree.heading(col, text=col)
            width = col_widths.get(col, 100)
            self.daily_usage_tree.column(col, width=width, anchor='center')
        
        # Grid layout for list_frame
        self.daily_usage_tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)

        # Right side: Details Panel
        detail_frame = ttk.LabelFrame(self.daily_history_paned, text="상세 정보")
        self.daily_history_paned.add(detail_frame, weight=1)

        self.daily_detail_text = tk.Text(detail_frame, wrap='word', width=35, background='#fdfdfd', font=('Pretendard', 10))
        self.daily_detail_text.pack(fill='both', expand=True, padx=5, pady=5)
        self.daily_detail_text.config(state='disabled')
        
        # Selection binding
        self.daily_usage_tree.bind('<<TreeviewSelect>>', self.on_daily_usage_tree_select)
        
        # Set a default sash position after UI is drawn if not loaded from config
        self.root.after(100, self._ensure_sash_visible)
        
        # Initial view update
        self.update_daily_usage_view()

    def _ensure_sash_visible(self):
        """Ensure the history details panel is visible by default if not set by config"""
        try:
            if hasattr(self, 'daily_history_paned'):
                self.root.update_idletasks() # Synchronize layout for accurate width
                total_w = self.daily_history_paned.winfo_width()
                if total_w < 50: # If window is still being drawn, try again shortly
                    self.root.after(200, self._ensure_sash_visible)
                    return

                try:
                    current_pos = self.daily_history_paned.sashpos(0)
                except:
                    current_pos = 0 # Assume hidden if errored

                # Only set if it's currently at 0 (hidden) or near the end (effectively hidden)
                if current_pos < 50 or current_pos > total_w - 50:
                    new_pos = int(total_w * 0.75) if total_w > 200 else 600
                    self.daily_history_paned.sashpos(0, new_pos)
        except Exception as e:
            print(f"Error ensuring sash visibility: {e}")
    
    def on_daily_usage_tree_select(self, event):
        """Update the side details panel when a row is selected using full DataFrame data"""
        selected = self.daily_usage_tree.selection()
        if not selected:
            return
            
        item = self.daily_usage_tree.item(selected[0])
        tags = item.get('tags', [])
        if not tags or tags[0] == 'total':
            # Clear if it's a total row or no tags
            self.daily_detail_text.config(state='normal')
            self.daily_detail_text.delete('1.0', tk.END)
            self.daily_detail_text.config(state='disabled')
            return
            
        try:
            df_idx = int(tags[0])
            if df_idx not in self.daily_usage_df.index:
                return
            
            entry = self.daily_usage_df.loc[df_idx]
            
            self.daily_detail_text.config(state='normal')
            self.daily_detail_text.delete('1.0', tk.END)
            
            # Header
            self.daily_detail_text.insert(tk.END, "=== 상세 내용 ===\n\n", "header")
            
            # Key mappings for better display labels
            display_map = [
                ('Date', '날짜'), ('Site', '현장'),
                ('User', '작업자1'), ('User2', '작업자2'), ('User3', '작업자3'),
                ('User4', '작업자4'), ('User5', '작업자5'), ('User6', '작업자6'),
                ('장비명', '장비명'), ('검사방법', '검사방법'), ('검사량', '검사량'),
                ('Film Count', '필름매수'), ('단가', '단가'), ('출장비', '출장비'),
                ('일식', '일식'), ('검사비', '검사비'), ('Note', '비고'),
                ('Entry Time', '입력시간')
            ]
            
            # Main fields
            for col_id, label in display_map:
                val = entry.get(col_id, '')
                if pd.isna(val) or str(val).strip() == '' or str(val).lower() == 'nan':
                    continue
                if col_id == 'Date' or col_id == 'Entry Time':
                    val = pd.to_datetime(val).strftime('%Y-%m-%d %H:%M') if pd.notna(val) else val
                
                self.daily_detail_text.insert(tk.END, f"• {label}: ", "label")
                self.daily_detail_text.insert(tk.END, f"{val}\n", "value")
            
            # Material Name (special lookup)
            mat_id = entry.get('Material ID')
            if pd.notna(mat_id):
                mat_name = self.get_material_display_name(mat_id)
                self.daily_detail_text.insert(tk.END, f"• 품목명: ", "label")
                self.daily_detail_text.insert(tk.END, f"{mat_name}\n", "value")

            # RTK Values
            rtk_cats = ["센터미스", "농도", "마킹미스", "필름마크", "취급부주의", "고객불만", "기타"]
            rtk_found = False
            for cat in rtk_cats:
                val = entry.get(f'RTK_{cat}', 0)
                if pd.notna(val) and float(val) > 0:
                    if not rtk_found:
                        self.daily_detail_text.insert(tk.END, "\n[RT 매수]\n", "section")
                        rtk_found = True
                    self.daily_detail_text.insert(tk.END, f"  - {cat}: {val}\n", "value")

            # NDT Values
            ndt_mats = ["형광자분", "흑색자분", "백색페인트", "침투제", "세척제", "현상제", "형광침투제"]
            ndt_found = False
            for mat in ndt_mats:
                val = entry.get(f'NDT_{mat}', 0)
                # Fallback check
                if pd.isna(val):
                    if mat == "흑색자분": val = entry.get('NDT_자분', 0)
                    elif mat == "백색페인트": val = entry.get('NDT_페인트', 0)
                    elif mat == "형광침투제": val = entry.get('NDT_형광', 0)
                if pd.notna(val) and float(val) > 0:
                    if not ndt_found:
                        self.daily_detail_text.insert(tk.END, "\n[NDT 자재]\n", "section")
                        ndt_found = True
                    self.daily_detail_text.insert(tk.END, f"  - {mat}: {val}\n", "value")
            
            self.daily_detail_text.tag_configure("header", font=('Arial', 11, 'bold'), foreground='blue')
            self.daily_detail_text.tag_configure("section", font=('Arial', 10, 'bold'), foreground='green')
            self.daily_detail_text.tag_configure("label", font=('Arial', 10, 'bold'))
            self.daily_detail_text.tag_configure("value", font=('Arial', 10))
            
            self.daily_detail_text.config(state='disabled')
        except Exception as e:
            print(f"Error updating details: {e}")
            self.daily_detail_text.config(state='disabled')

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
    
    def update_daily_test_fee_calc(self):
        """Auto-calculate Inspection Fee = (Amount * Unit Price) + Travel Expense + Meal Cost"""
        try:
            amount = float(self.ent_daily_test_amount.get().strip()) if self.ent_daily_test_amount.get().strip() else 0.0
            price = float(self.ent_daily_unit_price.get().strip()) if self.ent_daily_unit_price.get().strip() else 0.0
            travel = float(self.ent_daily_travel_cost.get().strip()) if self.ent_daily_travel_cost.get().strip() else 0.0
            meal = float(self.ent_daily_meal_cost.get().strip()) if self.ent_daily_meal_cost.get().strip() else 0.0
            
            calc_fee = (amount * price) + travel + meal
            
            # Update the fee field
            self.ent_daily_test_fee.delete(0, tk.END)
            self.ent_daily_test_fee.insert(0, f"{calc_fee:.0f}")
        except:
            pass

    def add_daily_usage_entry(self):
        """Add a daily usage entry"""
        try:
            date_str = self.ent_daily_date.get()
            site = self.ent_daily_site.get()
            mat_name = self.cb_daily_material.get()
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

            # Robust Material ID Lookup
            # Instead of splitting which fails if name contains " - ", 
            # we find the ID by reconstructing the display name same way as in the combo.
            mat_id = None
            pure_mat_name = ""
            
            for _, m_row in self.materials_df.iterrows():
                m_name = str(m_row.get('품목명', ''))
                m_model = str(m_row.get('모델명', '')).replace('nan', '').replace('None', '')
                m_sn = str(m_row.get('SN', '')).replace('nan', '').replace('None', '')
                
                m_display = m_name
                if m_model: m_display += f" - {m_model}"
                if m_sn: m_display += f" (SN: {m_sn})"
                
                if m_display == mat_name:
                    mat_id = m_row['Material ID']
                    pure_mat_name = m_name
                    break
            
            # Fallback for manual entry if exact match not found
            if mat_id is None:
                pure_mat_name = mat_name
                if " - " in pure_mat_name:
                    pure_mat_name = pure_mat_name.split(" - ")[0]
                if " (SN: " in pure_mat_name:
                    pure_mat_name = pure_mat_name.split(" (SN: ")[0]
                
                # Try finding by name only
                mat_rows = self.materials_df[self.materials_df['품목명'] == pure_mat_name]
                if not mat_rows.empty:
                    mat_id = mat_rows['Material ID'].values[0]
            
            # Parse date and combine with current time for transaction accuracy
            try:
                selected_date = datetime.datetime.strptime(date_str, '%Y-%m-%d')
                # Use current time to provide full timestamp for In/Out records
                usage_datetime = datetime.datetime.combine(selected_date.date(), datetime.datetime.now().time())
            except ValueError:
                messagebox.showwarning("입력 오류", "날짜 형식이 올바르지 않습니다. (YYYY-MM-DD)")
                return
            
            if mat_id is None:
                # Material doesn't exist in database, create it
                new_mat_id = self.materials_df['Material ID'].max() + 1 if not self.materials_df.empty else 1
                
                # More inclusive brand-to-category mapping for new materials
                default_cat = 'OTHER'
                name_up = pure_mat_name.upper()
                if any(k in name_up for k in ['CARESTREAM', 'AGFA', 'FUJI', 'KODAK', 'STRUCTURIX', 'FILM', '필름']):
                    default_cat = 'FILM'
                elif any(k in name_up for k in ['침투제', '세척제', '현상제', '자분', '페인트', 'NABAKEM', 'MAGNAFLUX']):
                    default_cat = 'NDT_CHEM'
                
                new_material = {
                    'Material ID': new_mat_id,
                    '품목명': pure_mat_name,
                    '관리품번': '',
                    '품목군코드': default_cat,
                    '규격': '',
                    '관리단위': 'EA',
                    '재고위치': '',
                    '공급업체': '',
                    '제조사': '',
                    '재고하한': 10,
                    '수량': 0
                }
                
                self.materials_df = pd.concat([self.materials_df, pd.DataFrame([new_material])], ignore_index=True)
                self.save_data()
                mat_id = new_mat_id
            else:
                # Material ID is already set from the lookup above
                pass
            
            manager_val = self.cb_daily_user.get().strip() if hasattr(self, 'cb_daily_user') else ''
            user2_val = self.cb_daily_user2.get().strip() if hasattr(self, 'cb_daily_user2') else ''
            user3_val = self.cb_daily_user3.get().strip() if hasattr(self, 'cb_daily_user3') else ''
            user4_val = self.cb_daily_user4.get().strip() if hasattr(self, 'cb_daily_user4') else ''
            user5_val = self.cb_daily_user5.get().strip() if hasattr(self, 'cb_daily_user5') else ''
            user6_val = self.cb_daily_user6.get().strip() if hasattr(self, 'cb_daily_user6') else ''
            
            user_names = [manager_val, user2_val, user3_val, user4_val, user5_val, user6_val]

            equip_name = self.cb_daily_equip.get().strip() if hasattr(self, 'cb_daily_equip') else ''
            test_method = self.cb_daily_test_method.get().strip() if hasattr(self, 'cb_daily_test_method') else ''
            
            try:
                test_amount = float(self.ent_daily_test_amount.get().strip()) if self.ent_daily_test_amount.get().strip() else 0.0
            except ValueError: test_amount = 0.0
            
            # Cost fields
            try:
                unit_price = float(self.ent_daily_unit_price.get().strip()) if self.ent_daily_unit_price.get().strip() else 0.0
            except ValueError: unit_price = 0.0
            try:
                travel_cost = float(self.ent_daily_travel_cost.get().strip()) if self.ent_daily_travel_cost.get().strip() else 0.0
            except ValueError: travel_cost = 0.0
            try:
                meal_cost = float(self.ent_daily_meal_cost.get().strip()) if self.ent_daily_meal_cost.get().strip() else 0.0
            except ValueError: meal_cost = 0.0
            try:
                test_fee = float(self.ent_daily_test_fee.get().strip()) if self.ent_daily_test_fee.get().strip() else 0.0
            except ValueError: test_fee = 0.0

            new_entry = {
                'Date': selected_date,
                'Site': site,
                'Material ID': mat_id,
                '장비명': equip_name,
                '검사방법': test_method,
                '검사량': test_amount,
                '단가': unit_price,
                '출장비': travel_cost,
                '일식': meal_cost,
                '검사비': test_fee,
                '검사비': test_fee,
                'Film Count': film_count,
                'Usage': total_usage,
                'Usage': total_usage,
                'Note': note,
                'Entry Time': datetime.datetime.now(),
                'User': manager_val,
                'User2': user2_val,
                'User3': user3_val,
                'User4': user4_val,
                'User5': user5_val,
                'User6': user6_val
            }
            
            # Add RTK values to the entry
            for category, value in rtk_values.items():
                new_entry[f'RTK_{category}'] = value
            
            # Add NDT materials values to the entry
            ndt_materials = ["형광자분", "흑색자분", "백색페인트", "침투제", "세척제", "현상제", "형광침투제"]
            
            for i, material in enumerate(ndt_materials):
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
            
            # Keyword mapping for NDT materials (PT/MT Chemicals)
            ndt_keywords = {
                "형광자분": ["형광자분", "FLUOR", "MAGNETIC", "GLO"],
                "흑색자분": ["흑색", "자분", "BLACK", "MAGNETIC", "7HF"],
                "백색페인트": ["백색", "페인트", "WHITE", "PAINT", "CONTRAST", "WCP"],
                "침투제": ["침투제", "PENETRANT", "VP", "SKL", "NP"],
                "세척제": ["세척제", "CLEANER", "SKC", "CLEAN", "NC"],
                "현상제": ["현상제", "DEVELOPER", "SKD", "DEV", "ND"],
                "형광침투제": ["형광", "FLUOR", "GLO"]
            }
            
            # 1. Deduct main selection (FILM or Item based on Inspection Amount)
            # UNIFIED LOGIC: Total Deduction = Film Count + RTK Total (total_usage)
            # Inspection Quantity (test_amount) is EXCLUDED from deduction as per user request ("검사량 빼줘").
            total_deduction = film_count + total_usage
            if total_deduction > 0:
                is_film = False
                mat_row = self.materials_df[self.materials_df['Material ID'] == mat_id]
                category = ""
                if not mat_row.empty:
                    category = str(mat_row.iloc[0].get('품목군코드', '')).upper()
                
                mat_name_upper = pure_mat_name.upper()
                
                # Very broad FILM detection
                if any(k in category for k in ['FILM', '필름', 'FLM', 'RT']):
                    is_film = True
                elif any(k in mat_name_upper for k in ['CARESTREAM', 'AGFA', 'FUJI', 'KODAK', 'STRUCTURIX', 'FILM', '필름', 'IX100', 'IX80', 'IX50', 'FOMA']):
                    is_film = True
                elif test_method == "RT":
                    is_film = True
                
                # Create transaction for total deduction
                all_workers = ", ".join([u for u in user_names if u])
                new_transaction = {
                    'Date': usage_datetime,
                    'Material ID': mat_id,
                    'Type': 'OUT',
                    'Quantity': total_deduction,
                    'User': all_workers,
                    'Site': site,
                    'Note': f'{site} 현장 사용 (자동 차감)'
                }
                self.transactions_df = pd.concat([self.transactions_df, pd.DataFrame([new_transaction])], ignore_index=True)
                
                # Build summary message
                parts = []
                if film_count > 0: parts.append(f"{film_count:.1f}(필름)")
                # if test_amount > 0: parts.append(f"{test_amount:.1f}(검사량)") # Excluded from deduction
                if total_usage > 0: parts.append(f"{total_usage:.1f}(재촬영)")
                summary = "+".join(parts)
                
                stock_info += f"\n• {pure_mat_name}: {summary} = {total_deduction:.1f} 차감"
            else:
                 # No deduction if total is 0
                 pass

            # 2. Deduct NDT materials (PT/MT Chemicals)
            for input_name, qty in ndt_usage.items():
                # Fallback mapping for renamed items
                search_names = [input_name]
                if input_name == "흑색자분": search_names.append("자분")
                elif input_name == "백색페인트": search_names.append("페인트")
                elif input_name == "형광침투제": search_names.append("형광")
                
                ndt_mat_rows = pd.DataFrame()
                
                # Step 1: Try exact match for any of the search names (품목명 or 모델명)
                for s_name in search_names:
                    matches = self.materials_df[
                        (self.materials_df['품목명'] == s_name) | 
                        (self.materials_df['모델명'] == s_name)
                    ]
                    if not matches.empty:
                        ndt_mat_rows = matches
                        break
                
                # Step 2: Try finding using keywords from ndt_keywords mapping
                if ndt_mat_rows.empty:
                    keywords = ndt_keywords.get(input_name, [input_name])
                    for kw in keywords:
                        matches = self.materials_df[
                            self.materials_df['품목명'].astype(str).str.contains(kw, na=False, case=False) |
                            self.materials_df['모델명'].astype(str).str.contains(kw, na=False, case=False)
                        ]
                        if not matches.empty:
                            ndt_mat_rows = matches
                            break
                
                # Step 3: Fallback to partial match of original name
                if ndt_mat_rows.empty:
                    for s_name in search_names:
                        matches = self.materials_df[
                            self.materials_df['품목명'].astype(str).str.contains(s_name, na=False, case=False) |
                            self.materials_df['모델명'].astype(str).str.contains(s_name, na=False, case=False)
                        ]
                        if not matches.empty:
                            ndt_mat_rows = matches
                            break
                    
                if not ndt_mat_rows.empty:
                    # Check category for PT/MT chemicals
                    ndt_mat_row = ndt_mat_rows.iloc[0]
                    ndt_mat_id = ndt_mat_row['Material ID']
                    actual_mat_name = ndt_mat_row['품목명']
                    category = str(ndt_mat_row.get('품목군코드', '')).upper()
                    
                    # Check if it belongs to PT or MT chemicals
                    is_chemical = False
                    category_upper = category.upper()
                    name_upper = str(actual_mat_name).upper()
                    
                    if any(k in category_upper for k in ['PT', 'MT', '약품', '소모품', 'CHEM', 'NDT']):
                        is_chemical = True
                    # Fallback: specific names if category is vague
                    elif any(k in name_upper for k in ['침투제', '세척제', '현상제', '자분', '페인트', '흑색자분', '백색페인트', '형광', 'DEVELOPER', 'CLEANER', 'PENETRANT', 'MT', 'PT', 'NABAKEM', 'MAGNAFLUX']):
                        is_chemical = True
                    
                    if is_chemical:
                        ndt_transaction = {
                            'Date': usage_datetime,
                            'Material ID': ndt_mat_id,
                            'Type': 'OUT',
                            'Quantity': qty,
                            'User': all_workers, # Changed from new_entry.get('User', '') to all_workers
                            'Site': site,
                            'Note': f'{site} 현장 사용 (자동 차감)'
                        }
                        self.transactions_df = pd.concat([self.transactions_df, pd.DataFrame([ndt_transaction])], ignore_index=True)
                        stock_info += f"\n• {actual_mat_name}: {qty:.1f} 차감"
                    else:
                        stock_info += f"\n• {actual_mat_name}: 재고 차감 대상 아님 (구분:'{category}', 약품 아님)"
                else:
                    stock_info += f"\n• {input_name}: 재고 정보 없음 (차감 실패 - 재고 등록 권장)"
            
            self.daily_usage_df = pd.concat([self.daily_usage_df, pd.DataFrame([new_entry])], ignore_index=True)
            
            # Count transactions effectively added in this call
            # (This is a bit hard with the current structure, but we can check if self.transactions_df length increased)
            # Actually, let's just make sure we save and refresh.
            
            self.save_data()
            
            # Set Daily Usage filters to the entry date to show relevant columns only
            if hasattr(self, 'ent_daily_start_date'): 
                self.ent_daily_start_date.delete(0, tk.END)
                self.ent_daily_start_date.insert(0, date_str)
            if hasattr(self, 'ent_daily_end_date'): 
                self.ent_daily_end_date.delete(0, tk.END)
                self.ent_daily_end_date.insert(0, date_str)
            if hasattr(self, 'cb_daily_filter_site'): self.cb_daily_filter_site.set("전체")
            if hasattr(self, 'cb_daily_filter_material'): self.cb_daily_filter_material.set("전체")
            
            self.update_daily_usage_view()
            
            # Reset filters in In/Out tab to ensure the new transaction is visible
            if hasattr(self, 'cb_trans_filter_mat'):
                self.cb_trans_filter_mat.set("전체")
            if hasattr(self, 'cb_trans_filter_site'):
                self.cb_trans_filter_site.set("전체")
                
            self.update_transaction_view() # Ensure In/Out tab is refreshed
            
            # Update site list if it's a new site
            if site not in self.sites:
                self.sites.append(site)
                self.sites.sort()
                self.ent_daily_site['values'] = self.sites
                self.save_tab_config()
                stock_info += f"\n• 신규 현장 '{site}'이 목록에 저장되었습니다."
            
            # Update user list if any new users entered
            for i in range(1, 7):
                cb_attr = f'cb_daily_user{i}' if i > 1 else 'cb_daily_user'
                if hasattr(self, cb_attr):
                    u_name = getattr(self, cb_attr).get().strip()
                    if u_name and u_name not in self.users:
                        self.users.append(u_name)
                        self.users.sort()
                        self.refresh_ui_for_list_change('users')
                        self.save_tab_config()
                        stock_info += f"\n• 신규 담당자 '{u_name}'이 목록에 저장되었습니다."

            # Update stock view and material lists
            self.update_stock_view()
            self.update_material_combo()
            self.update_registration_combos()
            
            final_msg = f"{site}의 {pure_mat_name} 사용 기록이 저장되었습니다."
            if stock_info:
                final_msg += f"\n\n[재고 차감 내역]:{stock_info}"
            else:
                final_msg += "\n\n(재고 차감 대상이 없어 입출고 내역이 생성되지 않았습니다.)"
                
            messagebox.showinfo("완료", final_msg)
            self.update_daily_usage_view()
            
            # Switch to In/Out Management tab (index 1) to see the transactions
            try:
                self.notebook.select(1)
            except:
                pass
                
            # Clear fields
            self.ent_daily_site.set('') # Clear site combobox
            for i in range(1, 7):
                cb_attr = f'cb_daily_user{i}' if i > 1 else 'cb_daily_user'
                if hasattr(self, cb_attr):
                    getattr(self, cb_attr).set('') # Clear user combobox
            self.ent_film_count.delete(0, tk.END)
            self.ent_film_count.insert(0, "0")
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

        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            messagebox.showerror("저장 오류", f"기록 저장 중 오류가 발생했습니다:\n{e}\n\n상세 정보:\n{error_details}")
    
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
        
        # Define RTK categories
        rtk_categories = ["센터미스", "농도", "마킹미스", "필름마크", "취급부주의", "고객불만", "기타", "총계"]
        
        # Display entries and calculate totals
        # Display entries and calculate totals
        total_film_count = 0
        total_rtk = [0.0] * len(rtk_categories)
        total_ndt = [0.0] * 7 # Increased from 6 to 7
        total_test_amount = 0.0
        total_unit_price = 0.0
        total_travel_cost = 0.0
        total_meal_cost = 0.0
        total_test_fee = 0.0
        
        
        current_date = None
        

        for idx, entry in filtered_df.iterrows():
            usage_date = entry.get('Date', '')
            if pd.notna(usage_date):
                usage_date = pd.to_datetime(usage_date).strftime('%Y-%m-%d')
            else:
                usage_date = "Unknown"
            
            current_date = usage_date
            
            mat_id = entry['Material ID']
            mat_name = self.get_material_display_name(mat_id)
            
            entry_time = entry.get('Entry Time', '')
            if pd.notna(entry_time):
                entry_time = pd.to_datetime(entry_time).strftime('%Y-%m-%d %H:%M')
            
            # Consolidate workers
            def clean_str(val):
                return str(val).replace('nan', '').replace('None', '').strip()

            all_users = [
                clean_str(entry.get('User', '')),
                clean_str(entry.get('User2', '')),
                clean_str(entry.get('User3', '')),
                clean_str(entry.get('User4', '')),
                clean_str(entry.get('User5', '')),
                clean_str(entry.get('User6', ''))
            ]
            consolidated_workers = ", ".join([u for u in all_users if u])
            
            # Get film count
            film_count = entry.get('Film Count', 0)
            if pd.notna(film_count):
                film_count_val = float(film_count)
                total_film_count += film_count_val
                film_count_str = f"{film_count_val:.1f}"
            else:
                film_count_str = "0.0"
            
            # Accumulate cost totals
            total_test_amount += entry.get('검사량', 0.0)
            total_unit_price += entry.get('단가', 0.0)
            total_travel_cost += entry.get('출장비', 0.0)
            total_meal_cost += entry.get('일식', 0.0)
            total_test_fee += entry.get('검사비', 0.0)
            
            # Get RTK values
            rtk_values = []
            row_rtk_total = 0
            # Use categories without '총계' for individual columns
            rtk_cats_only = rtk_categories[:-1]
            
            for i, category in enumerate(rtk_cats_only):
                value = entry.get(f'RTK_{category}', 0)
                if pd.notna(value):
                    val_float = float(value)
                    rtk_values.append(f"{val_float:.1f}")
                    row_rtk_total += val_float
                    total_rtk[i] += val_float
                else:
                    rtk_values.append("0.0")
            
            rtk_values.append(f"{row_rtk_total:.1f}")  # Add row total
            total_rtk[7] += row_rtk_total
            
            # Get NDT materials values
            ndt_values = []
            ndt_materials = ["형광자분", "흑색자분", "백색페인트", "침투제", "세척제", "현상제", "형광침투제"]
            
            for i, material in enumerate(ndt_materials):
                # Map new names to old keys for backward compatibility if needed
                col_name = f'NDT_{material}'
                value = entry.get(col_name)
                # Fallback for old data
                if pd.isna(value):
                    if material == "흑색자분": value = entry.get('NDT_자분', 0)
                    elif material == "백색페인트": value = entry.get('NDT_페인트', 0)
                    elif material == "형광침투제": value = entry.get('NDT_형광', 0)
                    else: value = 0
                if pd.notna(value):
                    val_float = float(value)
                    ndt_values.append(f"{val_float:.1f}")
                    total_ndt[i] += val_float
                else:
                    ndt_values.append("0.0")
            
            # Get remark with manager name if exists
            user_val = entry.get('User', '')
            if pd.isna(user_val) or str(user_val).lower() == 'nan': user_val = ''
            
            note_val = entry.get('Note', '')
            if pd.isna(note_val) or str(note_val).lower() == 'nan': note_val = ''
            
            display_note = f"[{user_val}] {note_val}" if user_val else note_val
            
            # Insert with index as tag for reliable deletion
            def clean_str(val):
                return str(val).replace('nan', '').replace('None', '').strip()

            self.daily_usage_tree.insert('', tk.END, values=(
                usage_date,
                entry.get('Site', ''),
                consolidated_workers,
                entry.get('장비명', ''),
                entry.get('검사방법', ''),
                f"{entry.get('검사량', 0.0):.1f}",
                film_count_str,
                f"{entry.get('단가', 0.0):,.0f}",
                f"{entry.get('출장비', 0.0):,.0f}",
                f"{entry.get('일식', 0.0):,.0f}",
                f"{entry.get('검사비', 0.0):,.0f}",
                mat_name,
                *rtk_values,
                *ndt_values,
                note_val, 
                entry_time
            ), tags=(str(idx),))
            
        # Insert last daily subtotal and final total row if data exists
        if not filtered_df.empty:
            
            # Final overall total
            self.daily_usage_tree.tag_configure('total', background='#E8F4F8', font=('Arial', 9, 'bold'))
            
            total_values = [
                '=== 전체 누계 ===',
                '', # 현장
                '', # 작업자
                '', # 장비명
                '', # 검사방법
                f"{total_test_amount:.1f}",
                f"{total_film_count:.1f}",
                f"{total_unit_price:,.0f}",
                f"{total_travel_cost:,.0f}",
                f"{total_meal_cost:,.0f}",
                f"{total_test_fee:,.0f}",
                '', # 품목명
                *[f"{v:.1f}" for v in total_rtk],
                *[f"{v:.1f}" for v in total_ndt],
                '',   # 비고
                ''    # 입력시간
            ]
            self.daily_usage_tree.insert('', tk.END, values=total_values, tags=('total',))
            
            # --- Dynamic Column Hiding ---
            mandatory_cols = ['날짜', '현장', '작업자', '장비명', '검사방법', '품목명', '비고', '입력시간']
            
            # Map dynamic columns to their total values
            # Index positions in total_rtk and total_ndt must match setup_daily_usage_tab's columns
            dynamic_col_status = {
                '검사량': total_test_amount > 0,
                '필름매수': total_film_count > 0,
                '단가': total_unit_price > 0,
                '출장비': total_travel_cost > 0,
                '일식': total_meal_cost > 0,
                '검사비': total_test_fee > 0,
            }
            
            # RTK Columns (11-17) and RT총계 (18)
            rtk_col_names = ["센터미스", "농도", "마킹미스", "필름마크", "취급부주의", "고객불만", "기타"]
            for i, col_name in enumerate(rtk_col_names):
                dynamic_col_status[col_name] = total_rtk[i] > 0
            dynamic_col_status['RT총계'] = total_rtk[7] > 0
            
            # NDT Columns (19-25)
            ndt_col_names = ["형광자분", "흑색자분", "백색페인트", "침투제", "세척제", "현상제", "형광침투제"]
            for i, col_name in enumerate(ndt_col_names):
                dynamic_col_status[col_name] = total_ndt[i] > 0
                
            # Build final visible columns list based on the order in self.daily_usage_tree['columns']
            all_cols = self.daily_usage_tree['columns']
            visible_cols = [col for col in all_cols if col in mandatory_cols or dynamic_col_status.get(col, False)]
            
            # Apply to treeview
            self.daily_usage_tree['displaycolumns'] = visible_cols
        else:
            # If empty, show mandatory columns including combined worker column
            self.daily_usage_tree['displaycolumns'] = ['날짜', '현장', '작업자', '장비명', '검사방법', '품목명', '비고', '입력시간']
    
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
                    ndt_materials = ["형광자분", "흑색자분", "백색페인트", "침투제", "세척제", "현상제", "형광"]
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

    def export_daily_usage_history(self):
        """Export current filtered daily usage history to Excel"""
        try:
            data = []
            for item in self.daily_usage_tree.get_children():
                data.append(self.daily_usage_tree.item(item, 'values'))
            
            if not data:
                messagebox.showinfo("알림", "내보낼 데이터가 없습니다.")
                return
            
            columns = self.daily_usage_tree['columns']
            df = pd.DataFrame(data, columns=columns)
            
            filename = f"일일사용내역_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=filename, filetypes=[("Excel files", "*.xlsx")])
            
            if save_path:
                df = pd.DataFrame(data, columns=columns)
                df = self.clean_df_export(df)
                self.save_df_to_excel_autofit(df, save_path, "일일사용내역")
                messagebox.showinfo("완료", f"데이터가 엑셀로 저장되었습니다.\n{save_path}")
        except Exception as e:
            messagebox.showerror("오류", f"엑셀 내보내기 중 오류가 발생했습니다: {e}")

    def export_all_daily_usage(self):
        """Export all daily usage records to Excel"""
        try:
            if self.daily_usage_df.empty:
                messagebox.showinfo("알림", "기록된 데이터가 없습니다.")
                return
            
            filename = f"전체사용기록_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=filename, filetypes=[("Excel files", "*.xlsx")])
            
            if save_path:
                export_df = self.clean_df_export(self.daily_usage_df.copy())
                self.save_df_to_excel_autofit(export_df, save_path, "전체기록")
                messagebox.showinfo("완료", f"전체 기록이 엑셀로 저장되었습니다.\n{save_path}")
        except Exception as e:
            messagebox.showerror("오류", f"엑셀 내보내기 중 오류가 발생했습니다: {e}")


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
        
        # If switching to daily usage tab, ensure the details panel is visible
        current_tab_idx = self.notebook.index("current")
        tab_text = self.notebook.tab(current_tab_idx, "text")
        if tab_text == '현장별 일일 사용량 기입':
            self._ensure_sash_visible()
    
    def auto_save_to_list(self, event, combobox, data_list, config_key):
        """Helper to auto-save new entry from combobox to a list and update all related UI"""
        new_val = combobox.get().strip()
        if not new_val:
            return
            
        if new_val not in data_list:
            data_list.append(new_val)
            data_list.sort()
            
            # Trigger app-wide update
            self.refresh_ui_for_list_change(config_key)
            self.save_tab_config()

    def save_tab_config(self):
        """Save current tab configuration (order and selected tab)"""
        try:
            # Force update to ensure all coordinates and sizes are accurate
            self.root.update_idletasks()
            config = {
                'selected_tab': self.notebook.index(self.notebook.select()),
                'tab_order': [],
                'sites': self.sites,
                'users': getattr(self, 'users', []),
                'warehouses': getattr(self, 'warehouses', []),
                'equipments': getattr(self, 'equipments', []),
                'layout_locked': getattr(self, 'layout_locked', False),
                'daily_usage_sash_pos': self.daily_usage_paned.sashpos(0) if hasattr(self, 'daily_usage_paned') else None,
                'daily_history_sash_pos': self.daily_history_paned.sashpos(0) if hasattr(self, 'daily_history_paned') else None
            }
            
            # Save current stock column widths
            config['stock_col_widths'] = {}
            if hasattr(self, 'stock_tree'):
                for col in self.stock_tree['columns']:
                    config['stock_col_widths'][col] = self.stock_tree.column(col, 'width')
            
            # Save tab order
            for tab in self.notebook.tabs():
                config['tab_order'].append(self.notebook.tab(tab, "text"))
            
            # Save Date Box position if modified (moved/resized)
            # Iterate through all registered draggable items
            if 'draggable_geometries' not in config:
                config['draggable_geometries'] = {}
            
            for key, widget in self.draggable_items.items():
                if widget.winfo_manager() == 'place':
                    config['draggable_geometries'][key] = {
                        'x': widget.winfo_x(),
                        'y': widget.winfo_y(),
                        'width': widget.winfo_width(),
                        'height': widget.winfo_height(),
                        'hidden': False
                    }
                    
                    # Save custom label if exists
                    if hasattr(widget, '_label_widget'):
                        config['draggable_geometries'][key]['custom_label'] = widget._label_widget.cget('text')
                    
                    # Save manage list key if exists
                    if hasattr(widget, '_manage_list_key') and widget._manage_list_key:
                        config['draggable_geometries'][key]['manage_list_key'] = widget._manage_list_key

                    # Save cloning info if it's a clone
                    if key.startswith('clone_'):
                        config['draggable_geometries'][key]['is_clone'] = True
                        config['draggable_geometries'][key]['widget_class_name'] = widget._widget_class.__name__
                        # Clean kwargs to avoid saving large stale lists like 'values'
                        saved_kwargs = widget._widget_kwargs.copy()
                        if 'values' in saved_kwargs:
                            del saved_kwargs['values']
                        config['draggable_geometries'][key]['widget_kwargs'] = saved_kwargs

                    # Save memo content and title
                    if key in self.memos:
                        config['draggable_geometries'][key]['text'] = self.memos[key]['text_widget'].get('1.0', 'end-1c')
                        config['draggable_geometries'][key]['memo_title'] = self.memos[key]['title_entry'].get()
                elif hasattr(widget, 'winfo_manager') and widget.winfo_manager() == '': # Hidden
                    # If it's already in config as hidden, keep it. 
                    # If it was just hidden, it won't have a manager.
                    if key in config['draggable_geometries']:
                        config['draggable_geometries'][key]['hidden'] = True
                    else:
                        config['draggable_geometries'][key] = {'hidden': True}
            
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
                    # Force update after selection so the tab is rendered and children computed
                    self.root.update_idletasks()
                
                # Restore sites list
                self.sites = config.get('sites', [])
                # If sites list is empty, try to populate from current daily_usage_df
                if not self.sites and not self.daily_usage_df.empty:
                    self.sites = sorted(self.daily_usage_df['Site'].dropna().unique().tolist())
                    self.sites = [str(s).strip() for s in self.sites if str(s).strip()]
                
                # Restore users list
                self.users = config.get('users', [])
                
                # Requested default workers
                default_workers = ["주진철", "우명광", "김진환", "장승대", "김성렬", "박광복", "주영광"]
                for w in default_workers:
                    if w not in self.users:
                        self.users.append(w)
                self.users.sort()
                
                if not self.users and not self.daily_usage_df.empty:
                    found_users = set()
                    for col in ['User', 'User2', 'User3', 'User4', 'User5', 'User6']:
                        if col in self.daily_usage_df.columns:
                            found_users.update(self.daily_usage_df[col].dropna().unique().tolist())
                    self.users = sorted([str(u).strip() for u in found_users if str(u).strip()])
                
                # Restore warehouses list
                self.warehouses = config.get('warehouses', [])
                if not self.warehouses and not self.materials_df.empty:
                    self.warehouses = sorted(self.materials_df['창고'].dropna().unique().tolist())
                    self.warehouses = [str(w).strip() for w in self.warehouses if str(w).strip()]
                
                # Equipment list handled by refresh_ui_for_list_change later
                self.equipments = config.get('equipments', [])
                if not self.equipments and not self.daily_usage_df.empty and '장비명' in self.daily_usage_df.columns:
                    self.equipments = sorted(self.daily_usage_df['장비명'].dropna().unique().tolist())
                    self.equipments = [str(e).strip() for e in self.equipments if str(e).strip()]
                
                # We will call refresh_ui_for_list_change for everything at the end
                # to ensure both standard and cloned/heuristic widgets are updated.
                
                # Restore stock column widths
                stock_col_widths = config.get('stock_col_widths', {})
                if stock_col_widths and hasattr(self, 'stock_tree'):
                    for col, width in stock_col_widths.items():
                        try:
                            self.stock_tree.column(col, width=int(width))
                        except:
                            pass
                if hasattr(self, '_loading_memos'):
                    del self._loading_memos
                    
                # Restore layout lock state
                self.layout_locked = config.get('layout_locked', False)
                if hasattr(self, 'btn_lock_layout'):
                    if self.layout_locked:
                        self.btn_lock_layout.config(text="🔒 배치 고정됨")
                        self.style.configure("Lock.TButton", foreground="black")
                    else:
                        self.btn_lock_layout.config(text="🔓 배치 수정 중")
                        self.style.configure("Lock.TButton", foreground="red")

                # Tab selection handled already at line 4531
                
                # Restore draggable items positions
                draggable_geos = config.get('draggable_geometries', {})
                
                # Recreate Memos and Clones first
                self._loading_memos = []
                # Map class names to actual classes for recreation
                class_map = {'Entry': ttk.Entry, 'Combobox': ttk.Combobox}
                
                for key, geo in draggable_geos.items():
                    if key.startswith('memo_'):
                        self._loading_memos.append(key)
                        self.add_new_memo(
                            initial_text=geo.get('text', ""), 
                            initial_title=geo.get('memo_title', "메모"), 
                            key=key
                        )
                    elif key.startswith('clone_'):
                        self._loading_memos.append(key)
                        cls_name = geo.get('widget_class_name', 'Entry')
                        cls = class_map.get(cls_name, ttk.Entry)
                        label = geo.get('custom_label', "복제항목")
                        kwargs = geo.get('widget_kwargs', {})
                        m_list_key = geo.get('manage_list_key')
                        cont, w = self.create_draggable_container(self.entry_inner_frame, label, cls, key, manage_list_key=m_list_key, **kwargs)
                        # Updates will be handled broad-scale at the end of load_tab_config
                
                # Apply custom labels to standard widgets
                for key, geo in draggable_geos.items():
                    if key in self.draggable_items and geo.get('custom_label'):
                        widget = self.draggable_items[key]
                        if hasattr(widget, '_label_widget'):
                            widget._label_widget.config(text=geo['custom_label'])
                
                # Backward compatibility for old keys if they exist and aren't in new dict
                if 'date_box_geometry' in config and 'date_box_geometry' not in draggable_geos:
                    draggable_geos['date_box_geometry'] = config['date_box_geometry']
                if 'site_box_geometry' in config and 'site_box_geometry' not in draggable_geos:
                    draggable_geos['site_box_geometry'] = config['site_box_geometry']
                
                for key, geo in draggable_geos.items():
                    if key in self.draggable_items:
                        widget = self.draggable_items[key]
                        try:
                            if geo.get('hidden'):
                                if widget.winfo_manager() == 'grid':
                                    widget.grid_forget()
                                elif widget.winfo_manager() == 'place':
                                    widget.place_forget()
                                continue
                                
                            # Ensure placeholder exists to prevent layout shift
                            # Use saved dimensions so placeholder has correct size
                            self._ensure_placeholder(widget, width=geo['width'], height=geo['height'])
                            
                            # Ensure it's on top and visible
                            widget.lift()
                            widget.place(
                                x=geo['x'],
                                y=geo['y'],
                                width=geo['width'],
                                height=geo['height']
                            )
                        except Exception as e:
                            print(f"Failed to restore position for {key}: {e}")
                
                # Restore daily_usage_paned (vertical) sash position
                daily_sash = config.get('daily_usage_sash_pos')
                if daily_sash is not None and hasattr(self, 'daily_usage_paned'):
                    try:
                        self.daily_usage_paned.sashpos(0, int(daily_sash))
                    except:
                        pass
                
                # Restore daily_history_paned (horizontal) sash position
                history_sash = config.get('daily_history_sash_pos')
                if history_sash is not None and hasattr(self, 'daily_history_paned'):
                    try:
                        self.daily_history_paned.sashpos(0, int(history_sash))
                    except:
                        pass

                # FINAL STEP: Refresh ALL list-based dropdowns (standard & clones & heuristics)
                # This ensures everyone is in sync with the current self.users, self.sites, etc.
                for l_key in ['users', 'sites', 'equipments', 'warehouses']:
                    self.refresh_ui_for_list_change(l_key)
                
                # Special cases for compound updates
                self.update_material_combo()
                
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
                stock_df = self.clean_df_export(stock_df)
                self.save_df_to_excel_autofit(stock_df, save_path, "재고현황")
                messagebox.showinfo("완료", f"재고 현황이 저장되었습니다.\n저장 위치: {save_path}")
                
        except Exception as e:
            messagebox.showerror("오류", f"내보내기 실패: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = MaterialManager(root)
    root.mainloop()
