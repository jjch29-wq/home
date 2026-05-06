import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import sys
import subprocess
import os
import time

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
import re

class WorkerCompositeWidget(ttk.Frame):
    """
    Composite widget for Worker selection: [Shift] + [Name]
    """
    def __init__(self, parent, shift_width=5, enable_autocomplete=False, user_list=None, **kwargs):
        super().__init__(parent)
        
        # Shift selection
        self.cb_shift = ttk.Combobox(self, values=["주간", "야간", "휴일"], width=shift_width, state="readonly")
        self.cb_shift.pack(side='left', padx=(0, 2))
        self.cb_shift.set("주간") # Default
        
        # Worker Name selection
        # Extract width from kwargs if present to apply to name combo
        name_width = kwargs.pop('width', 15)
        self.cb_name = ttk.Combobox(self, width=name_width, **kwargs)
        self.cb_name.pack(side='left', fill='x', expand=True)
        
        # Enable autocomplete if requested
        if enable_autocomplete and user_list is not None:
            def on_keyrelease(event):
                typed = event.widget.get()
                if not typed:
                    event.widget['values'] = user_list
                else:
                    filtered = [v for v in user_list if typed.lower() in str(v).lower()]
                    event.widget['values'] = filtered
            
            self.cb_name.bind('<KeyRelease>', on_keyrelease)
            
            def on_focus_in(event):
                typed = event.widget.get()
                if not typed:
                    event.widget['values'] = user_list
            
            self.cb_name.bind('<FocusIn>', on_focus_in)
        
    def get(self):
        """Return combined string: (Shift) Name"""
        shift = self.cb_shift.get()
        name = self.cb_name.get().strip()
        if not name:
            return ""
        return f"({shift}) {name}"

    def set(self, value):
        """Parse string '(Shift) Name' and set widgets"""
        if not value:
            self.cb_name.set("")
            self.cb_shift.set("주간")
            return
            
        import re
        match = re.match(r"\((주간|야간|휴일)\)\s*(.*)", value)
        if match:
            self.cb_shift.set(match.group(1))
            self.cb_name.set(match.group(2))
        else:
            # Fallback for old format (just name)
            self.cb_shift.set("주간")
            self.cb_name.set(value)

    def bind(self, sequence=None, func=None, add=None):
        """Forward binding to the name combobox (for focus out / return auto-save)"""
        self.cb_name.bind(sequence, func, add)

    def current(self, newindex=None):
        return self.cb_name.current(newindex)
        
    def config(self, **kwargs):
        self.cb_name.config(**kwargs)

    def __setitem__(self, key, value):
        self.cb_name[key] = value

    def __getitem__(self, key):
        return self.cb_name[key]

class WorkerDataGroup(ttk.Frame):
    """
    Unified widget for a worker's record: [Shift + Name] [WorkTime] [OT]
    """
    def __init__(self, parent, worker_index, users_list, time_list=None, enable_autocomplete=False, **kwargs):
        super().__init__(parent, padding=2) # Reduced padding for compact layout
        self.worker_index = worker_index
        
        # 1. Shift + Name (WorkerCompositeWidget)
        self.composite = WorkerCompositeWidget(
            self, width=12, values=users_list, 
            enable_autocomplete=enable_autocomplete, 
            user_list=users_list
        )
        self.composite.pack(side='left', padx=(0, 2))
        self.cb_name = self.composite.cb_name
        self.cb_shift = self.composite.cb_shift
        
        # 2. Work Time (Changed to Combobox for mouse selection)
        ttk.Label(self, text="시간:").pack(side='left', padx=(1, 0))
        self.ent_worktime = ttk.Combobox(self, width=12, values=time_list or [])
        self.ent_worktime.pack(side='left', padx=(0, 2))
        
        # 3. OT
        ttk.Label(self, text="OT:").pack(side='left', padx=(1, 0))
        self.ent_ot = ttk.Entry(self, width=12)
        self.ent_ot.pack(side='left')

    def get_worker(self): return self.composite.get()
    def set_worker(self, val): self.composite.set(val)
    def get_time(self): return self.ent_worktime.get()
    def set_time(self, val):
        self.ent_worktime.set(val)

    def get_ot(self): return self.ent_ot.get()
    def set_ot(self, val):
        self.ent_ot.delete(0, tk.END)
        self.ent_ot.insert(0, val)

    def bind_name(self, seq, func): self.cb_name.bind(seq, func)
    def bind_time(self, seq, func): 
        self.ent_worktime.bind(seq, func)
        if 'FocusOut' in seq or 'Return' in seq:
            # Also trigger on selection from dropdown
            self.ent_worktime.bind('<<ComboboxSelected>>', func, add='+')
            
    def bind_ot(self, seq, func): self.ent_ot.bind(seq, func)

class ColumnSelectionDialog(tk.Toplevel):
    """Dialog to select columns for Excel export"""
    def __init__(self, parent, columns, title="엑셀 출력 컬럼 선택"):
        super().__init__(parent)
        self.title(title)
        self.geometry("500x700") # Larger for better visibility
        self.transient(parent)
        self.grab_set()
        
        self.result = None
        self.vars = {}
        
        # Header
        lbl = ttk.Label(self, text="출력할 컬럼을 선택하세요:", font=('Malgun Gothic', 11, 'bold'))
        lbl.pack(pady=10)
        
        # Scrollable area for checkboxes
        container = ttk.Frame(self)
        container.pack(fill='both', expand=True, padx=20)
        
        canvas = tk.Canvas(container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Mousewheel scrolling
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        canvas.bind_all("<MouseWheel>", on_mousewheel)
        
        # Unbind when closed to avoid errors
        def on_close():
            canvas.unbind_all("<MouseWheel>")
            self.destroy()
        
        self.protocol("WM_DELETE_WINDOW", on_close)
            
        # Create checkboxes
        for col in columns:
            var = tk.BooleanVar(value=True)
            self.vars[col] = var
            cb = ttk.Checkbutton(scrollable_frame, text=col, variable=var)
            cb.pack(anchor='w', pady=2)
            
        # Buttons area
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill='x', pady=15, padx=20)
        
        def select_all():
            for v in self.vars.values(): v.set(True)
        def deselect_all():
            for v in self.vars.values(): v.set(False)
            
        ttk.Button(btn_frame, text="전체 선택", command=select_all).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="전체 해제", command=deselect_all).pack(side='left', padx=5)
        
        # Bottom controls
        bottom_frame = ttk.Frame(self)
        bottom_frame.pack(fill='x', pady=10)
        
        def on_ok():
            canvas.unbind_all("<MouseWheel>") # Unbind on OK too
            self.result = [col for col, var in self.vars.items() if var.get()]
            self.destroy()
            
        def on_cancel():
            on_close()
            
        ttk.Button(bottom_frame, text="확인", command=on_ok).pack(side='right', padx=10)
        ttk.Button(bottom_frame, text="취소", command=on_close).pack(side='right', padx=5)
        
        # self.wait_window() removed from here to allow caller to set vars before blocking
        # self.wait_window() removed from here to allow caller to set vars before blocking

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
        try:
            self.root.state('zoomed') # Maximize on Windows
        except:
            pass
        
        # Configure overall style
        self.style = ttk.Style()
        try:
            self.style.theme_use('clam') # Use 'clam' theme for better grid line visibility
        except:
            pass
            
        self.style.configure(".", font=('Malgun Gothic', 12))
        self.style.configure("Treeview.Heading", font=('Malgun Gothic', 12, 'bold'))
        self.style.configure("Treeview", font=('Malgun Gothic', 12), rowheight=35) # Increased row height for "boxed" look
        
        # Detect system background color for tk widgets (Canvas, Text)
        self.theme_bg = self.root.cget('bg')
        if not self.theme_bg or self.theme_bg == 'SystemButtonFace':
             # Fallback/Modern check for Windows
             try:
                 self.theme_bg = self.style.lookup('TFrame', 'background')
             except:
                 self.theme_bg = '#f0f0f0' # Standard Windows gray
        
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
        self.worktimes = [] # Initialize worktimes list
        self.ot_times = [] # Initialize ot_times list
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
        self.checklists = {} # {checklist_key: {'container': frame, 'title_entry': entry, 'item_frame': frame, 'items': []}}
        self.layout_locked = False
        self._last_motion_time = 0 # For performance throttling
        self.is_ready = False  # Suppress saves until fully loaded

        self.create_widgets()
        self.update_registration_combos()
        
        # Enable keyboard navigation
        self.setup_keyboard_shortcuts()

    def _ensure_canvas_scroll_region(self):
        """Canvas scroll region no longer needed - using direct frame"""
        # No canvas scroll region needed since we removed canvas
        # Frame will expand naturally without scrollbars
        pass

    def _on_daily_usage_sash_changed(self, event=None):
        """Handle sash position change to save ratio"""
        try:
            # Skip saving if locked
            if hasattr(self, 'daily_usage_sash_locked') and self.daily_usage_sash_locked:
                return

            if hasattr(self, 'daily_usage_paned'):
                self.daily_usage_paned.update_idletasks()
                total_h = self.daily_usage_paned.winfo_height()
                if total_h > 0:
                    sash_pos = self.daily_usage_paned.sashpos(0)
                    ratio = sash_pos / total_h
                    
                    if not hasattr(self, 'tab_config'):
                        self.tab_config = {}
                    
                    self.tab_config['daily_usage_sash_ratio'] = ratio
                    self.tab_config['daily_usage_sash_pos'] = sash_pos
                    
                    self.save_tab_config()
                    print(f"Sash ratio saved: {ratio:.3f}")
        except Exception as e:
            print(f"Error saving sash position: {e}")

    def toggle_resolution_lock(self):
        """Toggle window resolution lock"""
        try:
            self.resolution_locked = not self.resolution_locked
            
            if self.resolution_locked:
                self.locked_width = self.root.winfo_width()
                self.locked_height = self.root.winfo_height()
                self.root.resizable(False, False)
                if hasattr(self, 'btn_resolution_lock'):
                    self.btn_resolution_lock.config(text="🔒 해상도 고정됨")
                
                if not hasattr(self, 'tab_config'):
                    self.tab_config = {}
                self.tab_config['resolution_locked'] = True
                self.tab_config['locked_width'] = self.locked_width
                self.tab_config['locked_height'] = self.locked_height
                print(f"Resolution locked at: {self.locked_width}x{self.locked_height}")
            else:
                self.root.resizable(True, True)
                if hasattr(self, 'btn_resolution_lock'):
                    self.btn_resolution_lock.config(text="🔓 해상도 고정")
                
                if not hasattr(self, 'tab_config'):
                    self.tab_config = {}
                self.tab_config['resolution_locked'] = False
                print("Resolution unlocked")
                
            self.save_tab_config()
            self.force_save_config()
        except Exception as e:
            print(f"Error toggling resolution lock: {e}")

    def force_save_config(self):
        """Force immediate save of config to file"""
        try:
            import json
            
            # Load existing config first
            existing_config = {}
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    existing_config = json.load(f)
            
            # Merge tab_config into existing config
            if hasattr(self, 'tab_config') and self.tab_config:
                existing_config.update(self.tab_config)
            
            # Save merged config
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(existing_config, f, ensure_ascii=False, indent=2)
            
            print(f"Config forcefully saved to {self.config_path}")
            print(f"Saved keys: {list(existing_config.keys())}")
        except Exception as e:
            print(f"Error force saving config: {e}")

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
        import re
        def normalize_cols(df):
            if df is not None and not df.empty:
                df.columns = [re.sub(r'\s+', '', str(c)) for c in df.columns]
            return df

        try:
            print(f"DEBUG: Loading data from {self.db_path}...")
            
            # Check if database exists in app_dir. If not, try to restore from bundle_dir
            if not os.path.exists(self.db_path):
                bundled_db = os.path.join(self.bundle_dir, 'Material_Inventory.xlsx')
                print(f"DEBUG: Main DB not found. Trying to restore from bundle: {bundled_db}")
                if os.path.exists(bundled_db):
                    import shutil
                    try:
                        shutil.copy2(bundled_db, self.db_path)
                        print("DEBUG: Restored DB from bundle.")
                        # Also try to copy config if it exists in bundle but not in app_dir
                        bundled_config = os.path.join(self.bundle_dir, 'Material_Manager_Config.json')
                        if os.path.exists(bundled_config) and not os.path.exists(self.config_path):
                            shutil.copy2(bundled_config, self.config_path)
                            print("DEBUG: Restored Config from bundle.")
                    except Exception as e:
                        print(f"Failed to restore data from bundle: {e}")

            if not os.path.exists(self.db_path):
                print("DEBUG: DB still not found. Initializing new DataFrames.")
                # Initialize with new schema if still not found
                self.materials_df = pd.DataFrame(columns=[
                    'MaterialID', '회사코드', '관리품번', '품목명', 'SN', '창고',
                    '모델명', '규격', '품목군코드', '공급업체', '제조사', '제조국', 
                    '가격', '관리단위', '수량', '재고하한'
                ])
                self.transactions_df = pd.DataFrame(columns=['Date', 'MaterialID', 'Site', 'Type', 'Quantity', 'Note', 'User'])
                self.monthly_usage_df = pd.DataFrame(columns=['MaterialID', 'Year', 'Month', 'Site', 'Usage', 'Note', 'Entry Date'])
                self.daily_usage_df = pd.DataFrame(columns=['Date', 'Site', 'MaterialID', 'Usage', 'Note', 'EntryTime', 
                                                'RTK_센터미스', 'RTK_농도', 'RTK_마킹미스', 'RTK_필름마크', 
                                                'RTK_취급부주의', 'RTK_고객불만', 'RTK_기타', '장비명', '검사량',
                                                '단가', '출장비', '일식', '검사비', 'User', 'User2', 'User3', 'User4', 'User5', 'User6'])
            else:
                print("DEBUG: DB found. Reading Excel...")
                # 1. Materials
                self.materials_df = pd.read_excel(self.db_path, sheet_name='Materials')
                self.materials_df = normalize_cols(self.materials_df)
                
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
                
                # 2. Transactions
                self.transactions_df = pd.read_excel(self.db_path, sheet_name='Transactions')
                self.transactions_df = normalize_cols(self.transactions_df)
                
                # Ensure it has all required columns
                for col in ['Site', 'User', 'Note', 'MaterialID', 'Type', 'Quantity', 'Date']:
                    if col not in self.transactions_df.columns:
                        self.transactions_df[col] = '' if col != 'Quantity' else 0.0
                
                # Ensure Date column is datetime and MaterialID is numeric
                if not self.transactions_df.empty:
                    self.transactions_df['Date'] = pd.to_datetime(self.transactions_df['Date'], errors='coerce')
                    self.transactions_df['MaterialID'] = pd.to_numeric(self.transactions_df['MaterialID'], errors='coerce')
                    
                    # Force string columns and clean numeric artifacts
                    for col in ['Type', 'Note', 'User', 'Site']:
                        self.transactions_df[col] = self.transactions_df[col].astype(str).replace(['nan', 'None', 'NULL', '-0.0', '0.0', 'NaN'], '')
                        self.transactions_df[col] = self.transactions_df[col].str.replace(r'\.0$', '', regex=True)
                    
                    # One-time cleanup: Remove '현장사용' and redundant model names from historical notes
                    self.transactions_df['Note'] = self.transactions_df['Note'].astype(str).str.replace('현장사용', '', regex=False).str.strip()
                    
                    # Clean up notes that are identical to model names
                    if not self.transactions_df.empty and not self.materials_df.empty:
                        # Create a map for MaterialID -> Model Name
                        id_to_model = self.materials_df.set_index('MaterialID')['모델명'].astype(str).to_dict()
                        
                        def clean_redundant_note(row):
                            note = str(row['Note']).strip()
                            mat_id = row['MaterialID']
                            model = str(id_to_model.get(mat_id, '')).strip()
                            if note and model and note == model:
                                return ''
                            return note
                            
                        self.transactions_df['Note'] = self.transactions_df.apply(clean_redundant_note, axis=1)
                
                # 3. Monthly Usage
                try:
                    self.monthly_usage_df = pd.read_excel(self.db_path, sheet_name='MonthlyUsage', dtype={'Site': str, 'Note': str})
                    self.monthly_usage_df = normalize_cols(self.monthly_usage_df)
                    
                    if not self.monthly_usage_df.empty:
                        self.monthly_usage_df['MaterialID'] = pd.to_numeric(self.monthly_usage_df['MaterialID'], errors='coerce')
                        self.monthly_usage_df['EntryDate'] = pd.to_datetime(self.monthly_usage_df['EntryDate'])
                        # Add Site column if it doesn't exist (for backward compatibility)
                        if 'Site' not in self.monthly_usage_df.columns:
                            self.monthly_usage_df['Site'] = ''
                except Exception as e:
                    print(f"DEBUG: Failed to load MonthlyUsage: {e}")
                    self.monthly_usage_df = pd.DataFrame(columns=['MaterialID', 'Year', 'Month', 'Site', 'Usage', 'Note', 'EntryDate'])
                
                # 4. Daily Usage
                try:
                    # Explicitly set dtypes to avoid float inference for empty columns
                    self.daily_usage_df = pd.read_excel(self.db_path, sheet_name='DailyUsage', 
                                                        dtype={'Site': str, 'Note': str, 'User': str})
                    self.daily_usage_df = normalize_cols(self.daily_usage_df)
                                                        
                    if not self.daily_usage_df.empty:
                        self.daily_usage_df['Date'] = pd.to_datetime(self.daily_usage_df['Date'])
                        self.daily_usage_df['EntryTime'] = pd.to_datetime(self.daily_usage_df['EntryTime'])
                        
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
                        
                        # Ensure MaterialID is numeric
                        if 'MaterialID' in self.daily_usage_df.columns:
                            self.daily_usage_df['MaterialID'] = pd.to_numeric(self.daily_usage_df['MaterialID'], errors='coerce')
                          # Add RTK columns if they don't exist (for backward compatibility)
                        rtk_columns = ['RTK_센터미스', 'RTK_농도', 'RTK_마킹미스', 'RTK_필름마크', 
                                      'RTK_취급부주의', 'RTK_고객불만', 'RTK_기타']
                        for col in rtk_columns:
                            if col not in self.daily_usage_df.columns:
                                self.daily_usage_df[col] = 0.0
                        # Remove old RTK Category column if exists
                        if 'RTK Category' in self.daily_usage_df.columns:
                            self.daily_usage_df = self.daily_usage_df.drop('RTK Category', axis=1)
                except Exception as e:
                    print(f"DEBUG: Failed to load DailyUsage: {e}")
                    self.daily_usage_df = pd.DataFrame(columns=['Date', 'Site', 'MaterialID', 'Usage', 'Note', 'EntryTime', 
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
            import traceback
            traceback.print_exc()
            self.show_error_dialog("Error", f"데이터를 불러오는데 실패했습니다: {e}")
            self.materials_df = pd.DataFrame(columns=[
                'MaterialID', '회사코드', '관리품번', '품목명', 'SN', '창고',
                '모델명', '규격', '품목군코드', '공급업체', '제조사', '제조국', 
                '가격', '관리단위', '수량', '재고하한'
            ])
            self.transactions_df = pd.DataFrame(columns=['Date', 'MaterialID', 'Site', 'Type', 'Quantity', 'Note', 'User'])
            self.monthly_usage_df = pd.DataFrame(columns=['MaterialID', 'Year', 'Month', 'Site', 'Usage', 'Note', 'Entry Date'])
            self.daily_usage_df = pd.DataFrame(columns=['Date', 'Site', 'MaterialID', 'Usage', 'Note', 'EntryTime', 
                                                'RTK_센터미스', 'RTK_농도', 'RTK_마킹미스', 'RTK_필름마크', 
                                                'RTK_취급부주의', 'RTK_고객불만', 'RTK_기타', '장비명', '검사량'])
        
        # Global header cleanup: remove ALL internal and edge whitespace for permanent stability
        for df_attr in ['materials_df', 'transactions_df', 'daily_usage_df']:
            if hasattr(self, df_attr):
                df = getattr(self, df_attr)
                if df is not None and not df.empty:
                    df.columns = [re.sub(r'\s+', '', str(c)) for c in df.columns]
                    setattr(self, df_attr, df)

    
    def migrate_old_schema(self):
        """Migrate data from old schema to new schema"""
        old_df = self.materials_df.copy()
        self.materials_df = pd.DataFrame(columns=[
            'MaterialID', '회사코드', '관리품번', '품목명', 'SN', '창고',
            '모델명', '규격', '품목군코드', '제조사', '제조국', 
            '가격', '관리단위', '수량'
        ])
        
        # Headers are already cleaned by the global cleanup in load_data/init

        for _, row in old_df.iterrows():
            new_row = {
                'MaterialID': row.get('MaterialID', ''),
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
    
    def enable_autocomplete(self, combobox, values_list_attr=None, values_list=None):
        """
        Enable autocomplete/autosuggestion on a Combobox widget with proactive dropdown.
        """

        def on_keyrelease(event):
            # Ignore navigation and special keys
            if event.keysym in ("Left", "Right", "Up", "Down", "Return", "Escape", "Tab", "Shift_L", "Shift_R", "Control_L", "Control_R"):
                return
                
            # Get current text typed by user
            typed = event.widget.get()
            
            # Get the source values list
            if values_list_attr:
                all_values = getattr(self, values_list_attr, [])
            elif values_list is not None:
                all_values = values_list
            else:
                all_values = []
            
            if not typed:
                # Show all options if input is empty
                event.widget['values'] = all_values
            else:
                # Filter options based on what's typed (case-insensitive, partial match)
                filtered = [v for v in all_values if typed.lower() in str(v).lower()]
                event.widget['values'] = filtered
                
                # [PROACTIVE] Force dropdown to open so user sees suggestions
                if filtered and len(typed) > 0:
                    try:
                        # This opens the dropdown list in most Tkinter environments
                        event.widget.event_generate('<Down>')
                    except:
                        pass
        
        # Bind the keyrelease event to trigger filtering
        combobox.bind('<KeyRelease>', on_keyrelease)
        
        # Also bind FocusIn to update the values list when combobox gets focus
        def on_focus_in(event):
            typed = event.widget.get()
            if values_list_attr:
                all_values = getattr(self, values_list_attr, [])
            elif values_list is not None:
                all_values = values_list
            else:
                all_values = []
            
            if not typed:
                event.widget['values'] = all_values
        
        combobox.bind('<FocusIn>', on_focus_in)


    def apply_autocomplete_to_all_comboboxes(self):
        """Map specific comboboxes to their data lists for autocomplete"""
        # Mapping: {widget_attr_name: values_list_attr_name}
        mappings = {
            'ent_daily_site': 'sites',
            'cb_daily_equip': 'equipment_suggestions',
            'cb_material': 'materials_display_list',
            'cb_daily_material': 'materials_display_list',
            'cb_trans_filter_mat': 'materials_display_list',
            'cb_trans_filter_site': 'sites',
            'cb_daily_filter_site': 'sites',
            'cb_daily_filter_material': 'materials_display_list',
            'cb_daily_filter_worker': 'users',
            'cb_co_code': 'co_code_list',
            'cb_eq_code': 'eq_code_list',
            'cb_item_name': 'item_name_list',
            'cb_class': 'class_list',
            'cb_spec': 'spec_list',
            'cb_unit': 'unit_list',
            'cb_mfr': 'mfr_list',
            'cb_origin': 'origin_list'
        }
        
        for widget_attr, list_attr in mappings.items():
            if hasattr(self, widget_attr):
                widget = getattr(self, widget_attr)
                if isinstance(widget, ttk.Combobox):
                    # For filtering, we might want to include '전체' if it's already there
                    current_values = list(widget['values'])
                    if '전체' in current_values:
                        # Create a temporary list that includes '전체'
                        source_list = ['전체'] + getattr(self, list_attr, [])
                        self.enable_autocomplete(widget, values_list=source_list)
                    else:
                        self.enable_autocomplete(widget, values_list_attr=list_attr)

        # Handle cb_daily_test_method with static list
        if hasattr(self, 'cb_daily_test_method'):
            self.enable_autocomplete(self.cb_daily_test_method, values_list=["RT", "PAUT", "UT", "MT", "PT", "PMI"])

    def _safe_format_datetime(self, val, format_str='%Y-%m-%d %H:%M'):
        """Safely format a datetime value for display"""
        if pd.isna(val) or val == '':
            return ''
        try:
            dt = pd.to_datetime(val)
            if pd.isna(dt):
                return ''
            return dt.strftime(format_str)
        except:
            return str(val)

    def save_data(self):
        try:
            print(f"DEBUG: Saving data to {self.db_path}...")
            print(f"DEBUG: DataFrames size - Materials: {len(self.materials_df)}, Transactions: {len(self.transactions_df)}, Monthly: {len(self.monthly_usage_df)}, Daily: {len(self.daily_usage_df)}")
            
            with pd.ExcelWriter(self.db_path, engine='openpyxl') as writer:
                self.materials_df.to_excel(writer, sheet_name='Materials', index=False)
                self.transactions_df.to_excel(writer, sheet_name='Transactions', index=False)
                self.monthly_usage_df.to_excel(writer, sheet_name='MonthlyUsage', index=False)
                self.daily_usage_df.to_excel(writer, sheet_name='DailyUsage', index=False)
            
            print("DEBUG: Save successful.")
            return True
        except Exception as e:
            print(f"DEBUG: Save FAILED: {e}")
            self.show_error_dialog("Error", f"데이터를 저장하는데 실패했습니다: {e}")
            return False

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
        
        # Bind tab change event to ensure visibility
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)
        
        # Bind main window resize to maintain sash ratios
        self.root.bind("<Configure>", self._on_main_window_resize)
        
        # Load and restore all configuration (geometry, locks, sashes, draggable items)
        self.root.after(100, self.load_tab_config)
        
        # Apply autocomplete to all comboboxes
        self.root.after(200, self.apply_autocomplete_to_all_comboboxes)
        
        # Save tab config on window close
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

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
        col_widths = [60, 80, 100, 180, 90, 90, 120, 120, 90, 120, 120, 80, 80, 80, 80, 80]
        for col, width in zip(columns, col_widths):
            self.stock_tree.heading(col, text=col)
            self.stock_tree.column(col, width=width, minwidth=50, stretch=True, anchor='center')
        
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
            current = self.calculate_current_stock(mat['MaterialID'])
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
            
        # Get MaterialIDs to delete
        mat_ids_to_remove = []
        for item in selected_items:
            values = self.stock_tree.item(item, 'values')
            if values:
                # Ensure we match the type of MaterialID in the dataframe
                mat_ids_to_remove.append(type(self.materials_df['MaterialID'].iloc[0])(values[0]))
        
        # Remove from materials_df
        initial_count = len(self.materials_df)
        self.materials_df = self.materials_df[~self.materials_df['MaterialID'].isin(mat_ids_to_remove)]
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
        mat_data = self.materials_df[self.materials_df['MaterialID'] == mat_id].iloc[0]
        
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
        idx = self.materials_df.index[self.materials_df['MaterialID'] == mat_id].tolist()
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
        mat_trans = self.transactions_df[self.transactions_df['MaterialID'] == mat_id]
        in_qty = mat_trans[mat_trans['Type'] == 'IN']['Quantity'].sum()
        out_qty = mat_trans[mat_trans['Type'] == 'OUT']['Quantity'].sum()
        
        # Get the current stored quantity from materials_df
        mat = self.materials_df[self.materials_df['MaterialID'] == mat_id]
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
            mat_id = mat['MaterialID']
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
        # Create PanedWindow for better space utilization
        self.inout_paned = ttk.Panedwindow(self.tab_inout, orient='vertical')
        self.inout_paned.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Save sash position on adjustment
        self.inout_paned.bind("<ButtonRelease-1>", lambda e: self.save_tab_config())
        
        # Top frame for registration
        reg_container = ttk.Frame(self.inout_paned)
        self.inout_paned.add(reg_container, weight=1)
        
        # Scrollable frame for registration (to handle many fields)
        reg_canvas = tk.Canvas(reg_container, highlightthickness=0)
        reg_scrollbar = ttk.Scrollbar(reg_container, orient="vertical", command=reg_canvas.yview)
        reg_scrollable_frame = ttk.Frame(reg_canvas)
        
        reg_scrollable_frame.bind(
            "<Configure>",
            lambda e: reg_canvas.configure(scrollregion=reg_canvas.bbox("all"))
        )
        
        reg_canvas.create_window((0, 0), window=reg_scrollable_frame, anchor="nw")
        reg_canvas.configure(yscrollcommand=reg_scrollbar.set)
        
        reg_canvas.pack(side='left', fill='both', expand=True)
        reg_scrollbar.pack(side='right', fill='y')
        
        # Frame for Registration
        reg_frame = ttk.LabelFrame(reg_scrollable_frame, text="자재 신규 등록")
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
        
        # Bottom frame for transaction and history
        trans_container = ttk.Frame(self.inout_paned)
        self.inout_paned.add(trans_container, weight=2)  # Give more weight to transaction area
        
        # Frame for In/Out Transaction
        trans_frame = ttk.LabelFrame(trans_container, text="입출고 기록")
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
        history_frame = ttk.LabelFrame(trans_container, text="최근 입출고 내역")
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
            self.inout_tree.column(col, width=width, minwidth=50, stretch=True, anchor='center')
        
        self.inout_tree.grid(row=0, column=0, sticky='nsew')
        inout_vsb.grid(row=0, column=1, sticky='ns')
        inout_hsb.grid(row=1, column=0, sticky='ew')
        
        tree_scroll_frame.grid_rowconfigure(0, weight=1)
        tree_scroll_frame.grid_columnconfigure(0, weight=1)
        
        # Initial populate
        self.update_transaction_view()
        
        # Set initial sash position for better balance
        self.inout_paned.after(200, self._ensure_inout_sash_visibility)

    def _ensure_inout_sash_visibility(self):
        """Ensure the inout sash position is properly set"""
        try:
            if hasattr(self, 'inout_paned'):
                self.inout_paned.update_idletasks()
                total_h = self.inout_paned.winfo_height()
                if total_h > 200:
                    # Set sash to 30% of total height (registration area smaller)
                    new_pos = int(total_h * 0.3)
                    self.inout_paned.sashpos(0, new_pos)
                    print(f"Set inout sash to {new_pos} (total height: {total_h})")
        except Exception as e:
            print(f"Error ensuring inout sash visibility: {e}")


    def get_material_display_name(self, mat_id):
        """Get formatted material name as '품목명 (SN: SN번호) - 규격'"""
        if self.materials_df.empty:
            return f"ID: {mat_id}"
            
        mat_row = self.materials_df[self.materials_df['MaterialID'] == mat_id]
        if mat_row.empty:
            # Try numeric fallback just in case some IDs are still strings
            try:
                num_id = int(float(mat_id))
                mat_row = self.materials_df[self.materials_df['MaterialID'] == num_id]
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
                display = self.get_material_display_name(mat['MaterialID'])
                mat_list.append(display)
        
        # Merge database items with centralized films list, unique and sort
        all_vals = list(set([str(m) for m in mat_list + self.carestream_films if pd.notna(m) and str(m).strip()]))
        all_vals.sort()
        
        # Update ComboBoxes
        if hasattr(self, 'cb_material'):
            self.cb_material['values'] = all_vals
        if hasattr(self, 'cb_daily_material'):
            self.cb_daily_material['values'] = all_vals
            
        if hasattr(self, 'cb_daily_equip'):
            # Combine custom equipments with the full material display list (all_vals)
            combined_equip = list(set([str(e).strip() for e in self.equipments + all_vals if pd.notna(e) and str(e).strip()]))
            combined_equip.sort()
            self.cb_daily_equip['values'] = combined_equip
            self.equipment_suggestions = combined_equip
        
        self.materials_display_list = all_vals


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
        
        # Store lists for autocomplete
        self.co_code_list = []
        self.eq_code_list = []
        self.item_name_list = []
        self.class_list = []
        self.spec_list = []
        self.unit_list = []
        self.mfr_list = []
        self.origin_list = []

        attr_mapping = {
            '회사코드': 'co_code_list',
            '관리품번': 'eq_code_list',
            '품목명': 'item_name_list',
            '품목군코드': 'class_list',
            '규격': 'spec_list',
            '관리단위': 'unit_list',
            '제조사': 'mfr_list',
            '제조국': 'origin_list'
        }

        for col, combo in fields.items():
            vals = []
            if not self.materials_df.empty and col in self.materials_df.columns:
                unique_vals = self.materials_df[col].dropna().unique()
                vals = sorted([str(v).strip() for v in unique_vals if v and str(v).strip()])
            
            # Update instance attributes for autocomplete
            if col in attr_mapping:
                setattr(self, attr_mapping[col], vals)
            
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
        
        # Generate MaterialID
        if self.materials_df.empty:
            mat_id = 1
        else:
            mat_id = self.materials_df['MaterialID'].max() + 1
        
        # Extract SN from Model Name if present
        model_name = '' # Default empty
        new_model, new_sn = self.extract_sn_from_model(model_name, sn)
        # Note: In register_material, model_name is currently not an input field, but SN is.
        # If the user adds model name to registration later, this will handle it.
        
        new_row = {
            'MaterialID': mat_id,
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
            
            # Find MaterialID
            # Ensure MaterialID is treated consistently
            mat_rows = self.materials_df[self.materials_df['품목명'] == pure_mat_name]
            
            if mat_rows.empty:
                # If exact match fails, try a case-insensitive or stripped match
                mat_rows = self.materials_df[self.materials_df['품목명'].str.strip() == pure_mat_name]
            
            if mat_rows.empty:
                messagebox.showerror("오류", f"'{pure_mat_name}' 자재를 찾을 수 없습니다.\n먼저 자재를 등록하거나 정확한 명칭을 입력해주세요.")
                return
                
            mat_id = mat_rows['MaterialID'].values[0]
            
            # Update Warehouse in materials_df
            warehouse = str(self.cb_warehouse.get()).strip()
            if warehouse:
                # Type-safe assignment
                if '창고' in self.materials_df.columns:
                    mask = self.materials_df['MaterialID'] == mat_id
                    if mask.any():
                        self.materials_df.loc[mask, '창고'] = warehouse
            
            # Create transaction record
            new_trans = {
                'Date': datetime.datetime.now(),
                'MaterialID': mat_id,
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
            self.show_error_dialog("저장 오류", f"기록 저장 중 기술적인 오류가 발생했습니다:\n{e}\n\n상세 정보:\n{error_details}")
        
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
                    for mat_id in df_to_show['MaterialID'].unique():
                        if self.get_material_display_name(mat_id) == selected_mat:
                            matching_ids.append(mat_id)
                    
                    df_to_show = df_to_show[df_to_show['MaterialID'].isin(matching_ids)]
            
            # Apply site filter if selected
            if hasattr(self, 'cb_trans_filter_site'):
                selected_site = self.cb_trans_filter_site.get()
                if selected_site and selected_site != "전체":
                    df_to_show = df_to_show[df_to_show['Site'] == selected_site]

            df_sorted = df_to_show.sort_values(by='Date', ascending=False, na_position='last').head(500)
            
            for idx, row in df_sorted.iterrows():
                mat_id = row['MaterialID']
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
                mat_row = self.materials_df[self.materials_df['MaterialID'] == mat_id]
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
        col_widths = [90, 70, 140, 110, 110, 110, 110, 110, 220, 110, 80, 80, 80, 80, 80, 80, 80, 80, 90, 90, 90, 90, 90, 90, 90, 220]
        # Columns to center-align (numeric values)
        center_cols = ['검사량', '단가', '출장비', '일식', '검사비', '필름매수', '센터미스', '농도', '마킹미스', '필름마크', '취급부주의', '고객불만', '기타', 'RT총계', '형광자분', '흑색자분', '백색페인트', '침투제', '세척제', '현상제', '형광침투제']
        
        for col, width in zip(columns, col_widths):
            self.report_tree.heading(col, text=col)
            self.report_tree.column(col, width=width, minwidth=50, stretch=True, anchor='center')
        
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
            # Get unique material names from materials_df based on MaterialIDs in daily_usage_df
            unique_mat_ids = self.daily_usage_df['MaterialID'].dropna().unique()
            material_names = []
            for mat_id in unique_mat_ids:
                mat_row = self.materials_df[self.materials_df['MaterialID'] == mat_id]
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
        
        # Group by Year, Month, Site, MaterialID and aggregate
        grouped = df.groupby(['Year', 'Month', 'Site', 'MaterialID']).agg(agg_dict).reset_index()
        
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
            mat_id = entry['MaterialID']
            mat_row = self.materials_df[self.materials_df['MaterialID'] == mat_id]
            
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
            self.report_tree.tag_configure('total', background='#E8F4F8', font=('Arial', 12, 'bold'))
            
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
        ttk.Label(import_frame, text="형식: MaterialID, 회사코드, 관리품번, 품목명, 창고, 모델명, 규격, 품목군코드, 제조사, 제조국, 가격, 관리단위, 수량", 
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
                        new_mat_id = self.materials_df['MaterialID'].max() + 1 if not self.materials_df.empty else 1
                        data_row['MaterialID'] = new_mat_id
                        self.materials_df = pd.concat([self.materials_df, pd.DataFrame([data_row])], ignore_index=True)
                        count_new += 1
                else:
                    # Add as new
                    new_mat_id = self.materials_df['MaterialID'].max() + 1 if not self.materials_df.empty else 1
                    data_row['MaterialID'] = new_mat_id
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
            mat_id = mat['MaterialID']
            
            # Calculate monthly usage
            month_usage = monthly_trans[monthly_trans['MaterialID'] == mat_id]['Quantity'].sum()
            
            # Calculate cumulative usage (year-to-date)
            cumulative_usage = cumulative_trans[cumulative_trans['MaterialID'] == mat_id]['Quantity'].sum()
            
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
            mat_id = mat['MaterialID']
            
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
                month_mask = (yearly_trans['MaterialID'] == mat_id) & \
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
            mat_id = mat['MaterialID']
            total_usage = monthly_trans[monthly_trans['MaterialID'] == mat_id]['Quantity'].sum()
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
        
        # Treeview with columns including worker, work time, and OT fields
        columns = ('연도', '월', '현장', '작업자', '작업시간', 'OT1', 'OT2', 'OT3', 'OT4', 'OT5', 'OT6', 'OT7', 'OT8', 'OT9', 'OT10', 
                   '검사량', '단가', '출장비', '일식', '검사비', '품목명', '필름매수', '센터미스', '농도', '마킹미스', '필름마크', 
                   '취급부주의', '고객불만', '기타', 'RT총계', '형광자분', '흑색자분', '백색페인트', '침투제', '세척제', '현상제', '형광침투제', '비고')
        self.monthly_usage_tree = ttk.Treeview(tree_frame, columns=columns, show='headings',
                                               yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        vsb.config(command=self.monthly_usage_tree.yview)
        hsb.config(command=self.monthly_usage_tree.xview)
        
        # Column configuration with added OT columns
        col_widths = {
            '연도': 90, '월': 70, '현장': 140, '작업자': 100, '작업시간': 80,
            'OT1': 100, 'OT2': 100, 'OT3': 100, 'OT4': 100, 'OT5': 100,
            'OT6': 100, 'OT7': 100, 'OT8': 100, 'OT9': 100, 'OT10': 100,
            '검사량': 110, '단가': 110, '출장비': 110, '일식': 110, '검사비': 110,
            '품목명': 220, '필름매수': 110, '센터미스': 80, '농도': 80, '마킹미스': 80,
            '필름마크': 80, '취급부주의': 80, '고객불만': 80, '기타': 80, 'RT총계': 80,
            '형광자분': 90, '흑색자분': 90, '백색페인트': 90, '침투제': 90, '세척제': 90,
            '현상제': 90, '형광침투제': 90, '비고': 220
        }
        
        for col in columns:
            self.monthly_usage_tree.heading(col, text=col)
            width = col_widths.get(col, 100)
            self.monthly_usage_tree.column(col, width=width, minwidth=50, stretch=True, anchor='center')
        
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
        
        # Normalize column names - remove ALL types of whitespace using regex
        import re
        df.columns = [re.sub(r'\s+', '', str(c)) for c in df.columns]
        
        df['Year'] = pd.to_datetime(df['Date']).dt.year
        df['Month'] = pd.to_datetime(df['Date']).dt.month
        
        # DEBUG: Check what columns exist
        print(f"DEBUG: Normalized daily_usage_df columns: {df.columns.tolist()}")
        print(f"DEBUG: Number of rows before filter: {len(df)}")
        
        # Apply filters
        if filter_year != '전체':
            df = df[df['Year'] == int(filter_year)]
        
        if filter_month != '전체':
            df = df[df['Month'] == int(filter_month)]
        
        if filter_site != '전체':
            df = df[df['Site'] == filter_site]
        
        print(f"DEBUG: Number of rows after filter: {len(df)}")
        
        # Populate site filter options from data
        if hasattr(self, 'cb_filter_site_monthly'):
            unique_sites = ['전체'] + sorted(self.daily_usage_df['Site'].dropna().unique().tolist())
            self.cb_filter_site_monthly['values'] = unique_sites
            if not self.cb_filter_site_monthly.get():
                self.cb_filter_site_monthly.set('전체')
        
        # Populate material filter options from data
        if hasattr(self, 'cb_filter_material_monthly'):
            # Get unique material names from materials_df based on MaterialIDs in daily_usage_df
            # Note: MaterialID column name itself might have spaces in self.daily_usage_df
            m_id_col = 'MaterialID' if 'MaterialID' in self.daily_usage_df.columns else 'MaterialID'
            unique_mat_ids = self.daily_usage_df[m_id_col].dropna().unique()
            material_names = []
            for mat_id in unique_mat_ids:
                mat_row = self.materials_df[self.materials_df['MaterialID'] == mat_id]
                if not mat_row.empty:
                    material_names.append(mat_row.iloc[0]['품목명'])
            unique_materials = ['전체'] + sorted(set(material_names))
            self.cb_filter_material_monthly['values'] = unique_materials
            if not self.cb_filter_material_monthly.get():
                self.cb_filter_material_monthly.set('전체')
        
        # Prepare aggregation dictionary for all numeric fields
        agg_dict = {'Usage': 'sum'}
        
        # Helper for joining workers - clean internal spaces to avoid "11시간" vs "11 시간" mismatch
        def join_unique_non_empty(series):
            # Strip outer spaces and compress internal spaces for consistency
            vals = [" ".join(str(v).split()) for v in series if pd.notna(v) and str(v).strip()]
            return ", ".join(sorted(set(vals)))
        
        # Also define a sum helper that handles potential type issues
        def safe_sum(series):
            return pd.to_numeric(series, errors='coerce').sum()

        # Add worker info - combine unique values from User, User2, ..., User10
        worker_cols = ['User', 'User2', 'User3', 'User4', 'User5', 'User6', 'User7', 'User8', 'User9', 'User10']
        for col in worker_cols:
            if col in df.columns:
                agg_dict[col] = join_unique_non_empty
        
        # Add work time - join unique values
        worktime_cols = ['WorkTime', 'WorkTime2', 'WorkTime3', 'WorkTime4', 'WorkTime5', 'WorkTime6', 'WorkTime7', 'WorkTime8', 'WorkTime9', 'WorkTime10']
        for col in worktime_cols:
            if col in df.columns:
                agg_dict[col] = join_unique_non_empty
        
        # Add OT - join unique values
        ot_cols = ['OT', 'OT2', 'OT3', 'OT4', 'OT5', 'OT6', 'OT7', 'OT8', 'OT9', 'OT10']
        for col in ot_cols:
            if col in df.columns:
                agg_dict[col] = join_unique_non_empty
        
        # Add Film Count if exists
        if 'FilmCount' in df.columns:
            agg_dict['FilmCount'] = safe_sum
        
        # Add RTK categories
        rtk_categories = ['RTK_센터미스', 'RTK_농도', 'RTK_마킹미스', 'RTK_필름마크', 'RTK_취급부주의', 'RTK_고객불만', 'RTK_기타']
        for cat in rtk_categories:
            if cat in df.columns:
                agg_dict[cat] = safe_sum
        
        # Add NDT materials and cost fields
        other_agg_cols = ['NDT_형광자분', 'NDT_자분', 'NDT_흑색자분', 'NDT_페인트', 'NDT_백색페인트', 'NDT_침투제', 'NDT_세척제', 'NDT_현상제', 'NDT_형광', 'NDT_형광침투제',
                          '검사량', '단가', '출장비', '일식', '검사비']
        for col in other_agg_cols:
            if col in df.columns:
                agg_dict[col] = safe_sum
        
        # Group by Year, Month, Site, MaterialID and aggregate
        grouped = df.groupby(['Year', 'Month', 'Site', 'MaterialID']).agg(agg_dict).reset_index()
        
        print(f"DEBUG: Number of grouped rows: {len(grouped)}")
        if len(grouped) > 0:
            print(f"DEBUG: First grouped row: {grouped.iloc[0].to_dict()}")
        
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
            mat_id = entry['MaterialID']
            mat_row = self.materials_df[self.materials_df['MaterialID'] == mat_id]
            
            if not mat_row.empty:
                mat_name = mat_row.iloc[0]['품목명']
            else:
                mat_name = f"ID: {mat_id}"
            
            # Apply material filter
            if filter_material != '전체' and mat_name != filter_material:
                continue
            
            # Get aggregated values from clean columns
            film_count = entry.get('FilmCount', 0.0)
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
            
            # Extract worker names from User, User2, ..., User10
            all_workers = []
            worker_cols = ['User', 'User2', 'User3', 'User4', 'User5', 'User6', 'User7', 'User8', 'User9', 'User10']
            for col in worker_cols:
                val = str(entry.get(col, '')).strip()
                if val and val != 'nan' and val != '0.0':
                    # Split in case it was already joined in aggregation
                    all_workers.extend([v.strip() for v in val.split(',') if v.strip()])
            worker_str = ", ".join(sorted(set(all_workers)))
            
            # Concatenate work times
            all_worktimes = []
            worktime_cols = ['WorkTime', 'WorkTime2', 'WorkTime3', 'WorkTime4', 'WorkTime5', 'WorkTime6', 'WorkTime7', 'WorkTime8', 'WorkTime9', 'WorkTime10']
            for col in worktime_cols:
                val = str(entry.get(col, '')).strip()
                if val and val != 'nan' and val != '0.0':
                    all_worktimes.extend([v.strip() for v in val.split(',') if v.strip()])
            worktime_str = ", ".join(sorted(set(all_worktimes)))
            
            # Extract OT amounts
            ot_values = []
            ot_cols = ['OT', 'OT2', 'OT3', 'OT4', 'OT5', 'OT6', 'OT7', 'OT8', 'OT9', 'OT10']
            for col in ot_cols:
                val = str(entry.get(col, '')).strip()
                if val and val != 'nan' and val != '0.0':
                    # Handle multiple OTs if joined
                    sub_vals = [v.strip() for v in val.split(',') if v.strip()]
                    parsed_ots = []
                    for v_str in sub_vals:
                        if '(' in v_str and '원)' in v_str:
                            try:
                                amount_str = v_str.split('(')[1].split('원')[0].replace(',', '').strip()
                                amount = int(amount_str)
                                parsed_ots.append(f"{amount:,}")
                            except:
                                parsed_ots.append(v_str)
                        else:
                            parsed_ots.append(v_str)
                    ot_values.append(", ".join(parsed_ots))
                else:
                    ot_values.append('')
            
            self.monthly_usage_tree.insert('', tk.END, values=(
                int(entry['Year']),
                int(entry['Month']),
                entry.get('Site', ''),
                worker_str,  # 작업자
                worktime_str,  # 작업시간 (concatenated strings)
                *ot_values,  # OT1-10 (parsed amounts)
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
            self.monthly_usage_tree.tag_configure('total', background='#E8F4F8', font=('Arial', 12, 'bold'))
            
            self.monthly_usage_tree.insert('', tk.END, values=(
                '',
                '',
                '=== 전체 누계 ===',
                '',  # 작업자
                '',  # 작업시간
                '', '', '', '', '', '', '', '', '', '',  # OT1-10
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

    def create_draggable_container(self, parent, label_text, widget_class, config_key, manage_list_key=None, grid_info=None, **widget_kwargs):
        """Create a draggable container with a label and a widget, styled as a visible box"""
        # Use ttk.Frame for automatic background matching
        container = ttk.Frame(parent, relief="solid", borderwidth=1)
        
        # Header container for label and buttons
        hdr = ttk.Frame(container) 
        hdr.pack(side='top', fill='x', padx=0, pady=0)
        
        # Drag handle icon
        lbl_drag = ttk.Label(hdr, text="✥", font=('Arial', 9), cursor='fleur')
        lbl_drag.pack(side='left', padx=1)
        self.make_header_draggable(lbl_drag, container)
        
        # Label
        lbl = ttk.Label(hdr, text=label_text, font=('Malgun Gothic', 9, 'bold'))
        lbl.pack(side='left', padx=1)
        self.make_header_draggable(lbl, container)
        
        # Icons container on the right
        btn_box = ttk.Frame(hdr)
        btn_box.pack(side='right', padx=1)

        # Rename icon
        btn_rename = ttk.Label(btn_box, text="✏️", font=('Arial', 8), cursor='hand2')
        btn_rename.pack(side='left', padx=1)
        btn_rename.bind('<Button-1>', lambda e: self.rename_widget_label(config_key))
        
        # Clone icon
        btn_clone = ttk.Label(btn_box, text="📋", font=('Arial', 8), cursor='hand2')
        btn_clone.pack(side='left', padx=1)
        btn_clone.bind('<Button-1>', lambda e: self.clone_widget(config_key))

        # Delete icon (X)
        btn_del = ttk.Label(btn_box, text="❌", font=('Arial', 8), cursor='hand2')
        btn_del.pack(side='left', padx=1)
        btn_del.bind('<Button-1>', lambda e: self.remove_box(config_key))

        # Manage List Icon (Gear) - if it's a list-based widget
        if manage_list_key:
            btn_manage = ttk.Label(btn_box, text="⚙️", font=('Arial', 8), cursor='hand2')
            btn_manage.pack(side='left', padx=1)
            btn_manage.bind('<Button-1>', lambda e: self.open_list_management_dialog(manage_list_key))
        
        # Internal Content Area
        content_area = ttk.Frame(container, padding=(1, 0))
        content_area.pack(side='top', fill='both', expand=True)

        # Widget
        widget = widget_class(content_area, **widget_kwargs)
        widget.pack(side='left', fill='both', expand=True)
        
        # If the widget is a basic tk widget (Text, Canvas), set its background
        if hasattr(widget, 'config') and 'bg' in widget.keys():
            try: widget.config(bg=self.theme_bg)
            except: pass
        
        # NEW: Handle grid placement first if grid_info provided
        if grid_info:
            container.grid(**grid_info)
        
        # Track for layout reset/config
        container._config_key = config_key
        container._label_widget = lbl
        container._widget = widget
        container._widget_class = widget_class
        container._widget_kwargs = widget_kwargs
        container._manage_list_key = manage_list_key
        
        # Register and make draggable (this now captures the CORRECT grid info)
        self.draggable_items[config_key] = container
        self.make_draggable(container, config_key)
        
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
            # Updated to match current attribute names if needed, 
            # but usually it's cb_daily_user for 1, and cb_daily_user{i} for 2-10
            for i in range(1, 11):
                attr = 'cb_daily_user' if i == 1 else f'cb_daily_user{i}'
                if hasattr(self, attr):
                    widget = getattr(self, attr)
                    if isinstance(widget, WorkerCompositeWidget):
                         widget.cb_name['values'] = sorted_vals
                    elif hasattr(widget, 'configure'):
                        try: widget['values'] = sorted_vals
                        except: pass
            
            if hasattr(self, 'cb_trans_user'): self.cb_trans_user['values'] = sorted_vals
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
        
        if key.startswith('memo_') or key.startswith('clone_') or key.startswith('checklist_'):
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
            if key in self.checklists:
                del self.checklists[key]
            
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
        elif key in self.checklists:
            current_text = self.checklists[key]['title_entry'].get()
            
        new_name = simpledialog.askstring("이름 변경", "새 이름을 입력하세요:", initialvalue=current_text)
        if new_name is not None:
            if hasattr(widget, '_label_widget'):
                widget._label_widget.config(text=new_name)
            elif key in self.memos:
                self.memos[key]['title_entry'].delete(0, 'end')
                self.memos[key]['title_entry'].insert(0, new_name)
            elif key in self.checklists:
                self.checklists[key]['title_entry'].delete(0, 'end')
                self.checklists[key]['title_entry'].insert(0, new_name)
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
            
            # Copy value from original widget
            if hasattr(orig, '_widget'):
                try:
                    current_val = str(orig._widget.get()) # Ensure string
                    
                    # Try generic Entry-like setting (works for Entry and Combobox text area)
                    if hasattr(w, 'delete') and hasattr(w, 'insert'):
                        try:
                            w.delete(0, 'end')
                            w.insert(0, current_val)
                        except:
                            # Readonly comboboxes might fail delete/insert
                            pass
                            
                    # Try specific set method (Combobox, Scale, etc)
                    if hasattr(w, 'set'):
                        w.set(current_val)
                        
                except Exception as e:
                    print(f"Failed to copy value: {e}")
            
            cont.place(x=50, y=50) # Start position
            self.save_tab_config()

        elif key in self.memos:
            # It's a memo
            content = self.memos[key]['text_widget'].get('1.0', 'end-1c')
            title = self.memos[key]['title_entry'].get()
            self.add_new_memo(initial_text=content, initial_title=title, key=new_key)
            self.save_tab_config()
        elif key in self.checklists:
            # It's a checklist
            self.duplicate_checklist(key)
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

    def make_header_draggable(self, widget, target_container):
        """Make a specific widget (header/label) draggable with Left Mouse Button targeting a container"""
        from functools import partial
        widget.bind("<Button-1>", partial(self.on_drag_start, widget=target_container))
        widget.bind("<B1-Motion>", partial(self.on_mouse_motion, widget=target_container))
        widget.bind("<ButtonRelease-1>", partial(self.on_drag_stop, widget=target_container))
        
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
                # Explicitly unhide if it was hidden
                if widget.winfo_manager() == '':
                    # For grid, we just grid it back
                    if hasattr(widget, '_original_grid_info'):
                         widget.grid(**widget._original_grid_info)
                
                self.reset_widget_position(None, widget=widget)
            
            # 1.1 Restore parent propagation safely and reset size
            if hasattr(self, 'entry_inner_frame'):
                self.entry_inner_frame.pack_propagate(True)
                self.entry_inner_frame.grid_propagate(True)
                self.entry_inner_frame.config(height=0, width=0) # Let it auto-size for reset
                
                # Perform a layout update pass
                self.root.update_idletasks()
                
                # Re-apply stability locks
                self._adjust_parent_height(self.entry_inner_frame, force=True)
                self.entry_inner_frame.pack_propagate(False)
                self.entry_inner_frame.grid_propagate(False)
                
                # Refresh scrollregion if in canvas
                if hasattr(self, 'entry_canvas'):
                    self.entry_canvas.configure(scrollregion=self.entry_canvas.bbox("all"))
            
            # 2. Reset sash positions (splitters)
            try:
                if hasattr(self, 'daily_usage_paned'):
                    # Give more space to the entry form by default (500px)
                    self.daily_usage_paned.sashpos(0, 500) 
                if hasattr(self, 'daily_history_paned'):
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
        
        # Save layout lock state to config immediately
        if not hasattr(self, 'tab_config'):
            self.tab_config = {}
        self.tab_config['layout_locked'] = self.layout_locked
        self.save_tab_config()
        
        # Force immediate save to file
        self.force_save_config()
        
        print(f"Layout lock {'enabled' if self.layout_locked else 'disabled'}")
        
        # Save lock state
        self.save_tab_config()

    def on_drag_stop(self, event, widget=None):
        """Handle end of dragging or resizing and auto-save"""
        if widget is None:
            widget = event.widget
        if hasattr(widget, '_interaction_mode'):
            mode = getattr(widget, '_interaction_mode')
            del widget._interaction_mode
            
            # Update parent height if something moved or resized
            self._adjust_parent_height(widget.master, force=True)
            
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
        text_area = tk.Text(memo_container, wrap='word', height=5, width=30, font=('Arial', 10), bg=self.theme_bg, highlightthickness=0)
        text_area.pack(fill='both', expand=True, padx=2, pady=2)
        text_area.insert('1.0', initial_text)
        
        # Bind text change to auto-save
        text_area.bind('<FocusOut>', lambda e: self.save_tab_config())
        
        # Make draggable (Right Click)
        self.make_draggable(memo_container, key)
        
        # Make header draggable (Left Click)
        self.make_header_draggable(ctrl_frame, memo_container)
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

    def add_new_checklist(self, initial_data=None, initial_title="체크리스트", key=None):
        """Create a new movable/editable checklist box"""
        import time
        if key is None:
            key = f"checklist_{int(time.time() * 1000)}"
        
        # Create container
        check_container = ttk.LabelFrame(self.entry_inner_frame)
        
        # Controls frame (top of checklist)
        ctrl_frame = ttk.Frame(check_container)
        ctrl_frame.pack(fill='x', side='top')
        
        # Editable Title Entry
        title_entry = ttk.Entry(ctrl_frame, font=('Arial', 9, 'bold'), width=15)
        title_entry.pack(side='left', padx=2)
        title_entry.insert(0, initial_title)
        title_entry.bind('<FocusOut>', lambda e: self.save_tab_config())
        
        btn_copy = ttk.Button(ctrl_frame, text="📋", width=3, command=lambda: self.duplicate_checklist(key))
        btn_copy.pack(side='right')
        
        btn_del = ttk.Button(ctrl_frame, text="❌", width=3, command=lambda: self.remove_box(key))
        btn_del.pack(side='right')
        
        # Add Item Area
        add_frame = ttk.Frame(check_container)
        add_frame.pack(fill='x', padx=2, pady=2)
        
        new_item_var = tk.StringVar()
        entry_new = ttk.Entry(add_frame, textvariable=new_item_var, width=20)
        entry_new.pack(side='left', fill='x', expand=True)
        
        def add_item(event=None):
             text = new_item_var.get().strip()
             if text:
                 self.add_checklist_item(item_frame, text, False, key)
                 new_item_var.set("")
                 self.save_tab_config()
        
        entry_new.bind('<Return>', add_item)
        btn_add = ttk.Button(add_frame, text="➕", width=3, command=add_item)
        btn_add.pack(side='right')

        # Scrollable Frame for Items
        canvas_frame = ttk.Frame(check_container)
        canvas_frame.pack(fill='both', expand=True, padx=2, pady=2)
        
        canvas = tk.Canvas(canvas_frame, height=100, width=200, bg=self.theme_bg, highlightthickness=0) # Match theme
        scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
        item_frame = ttk.Frame(canvas)
        
        item_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=item_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # [STABILITY FIX] Use Enter/Leave to localize scrolling and prevent conflict with other lists
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            
        def _bind_scroll(event):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
        def _unbind_scroll(event):
            canvas.unbind_all("<MouseWheel>")
            
        canvas.bind('<Enter>', _bind_scroll)
        canvas.bind('<Leave>', _unbind_scroll)

        # Make draggable
        self.make_draggable(check_container, key)
        self.draggable_items[key] = check_container
        self.checklists[key] = {
            'container': check_container,
            'title_entry': title_entry,
            'item_frame': item_frame,
            'items': [] # List of item widgets/vars will be managed dynamically via children of item_frame
        }
        
        # Add initial items if provided
        if initial_data:
            for item in initial_data:
                self.add_checklist_item(item_frame, item.get('text', ''), item.get('checked', False), key)
        
        # Initial placement
        if key not in getattr(self, '_loading_memos', []): # Reuse loading flag or logic
            check_container.place(x=50, y=150)
            
        return check_container

    def add_checklist_item(self, parent_frame, text, checked, checklist_key):
        """Add a single item row to the checklist"""
        row_frame = ttk.Frame(parent_frame)
        row_frame.pack(fill='x', pady=1)
        
        var = tk.BooleanVar(value=checked)
        cb = ttk.Checkbutton(row_frame, variable=var, command=lambda: self.save_tab_config())
        cb.pack(side='left')
        
        entry = ttk.Entry(row_frame)
        entry.insert(0, text)
        entry.pack(side='left', fill='x', expand=True)
        entry.bind('<FocusOut>', lambda e: self.save_tab_config())
        
        def delete_this_item():
            row_frame.destroy()
            self.save_tab_config()
            
        btn_del_item = ttk.Label(row_frame, text="❌", font=('Arial', 8), cursor='hand2', foreground='gray')
        btn_del_item.pack(side='right', padx=2)
        btn_del_item.bind('<Button-1>', lambda e: delete_this_item())
        
        # Store refs in widget for retrieval during save
        row_frame._checklist_data = {'var': var, 'entry': entry}

    def duplicate_checklist(self, key):
        """Duplicate an existing checklist"""
        if self.layout_locked: return
        if key in self.checklists:
            # get current data
            original_items = []
            item_frame = self.checklists[key]['item_frame']
            for child in item_frame.winfo_children():
                if hasattr(child, '_checklist_data'):
                    data = child._checklist_data
                    original_items.append({
                        'text': data['entry'].get(),
                        'checked': data['var'].get()
                    })
            
            title = self.checklists[key]['title_entry'].get()
            self.add_new_checklist(initial_data=original_items, initial_title=title)
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
        """Handle dragging or resizing motion with performance throttling"""
        if self.layout_locked:
            return "break"
            
        if widget is None:
            widget = event.widget
        
        if not hasattr(widget, '_interaction_mode'):
            return
        
        # PERFORMANCE THROTTLE: Limit updates to ~60fps (16ms)
        curr_time = time.time()
        if curr_time - self._last_motion_time < 0.016:
            # Still update the physical position of the widget being interacted with
            # or the user will feel lag in the initial drag/resize itself.
            self._update_widget_position(event, widget)
            return
        
        self._last_motion_time = curr_time
        
        # Apply positioning only (No collision, No auto-resize)
        self._update_widget_position(event, widget)

    def _update_widget_position(self, event, widget):
        """Internal helper to calculate and set widget position/size during interaction"""
        dx = event.x_root - widget._drag_start_root_x
        dy = event.y_root - widget._drag_start_root_y
        parent = widget.master
        parent_w = parent.winfo_width()
        parent_h = parent.winfo_height()
        
        if widget._interaction_mode == 'move':
            new_x = widget._drag_start_pos_x + dx
            new_y = widget._drag_start_pos_y + dy
            widget_w = widget.winfo_width()
            widget_h = widget.winfo_height()
            
            # Clamp
            if new_x < 0: new_x = 0
            elif new_x + widget_w > parent_w: new_x = max(0, parent_w - widget_w)
            
            # Loosen Y clamping: allow moving slightly below parent to trigger growth
            if new_y < 0: new_y = 0
            # Instead of strict clamp at parent_h, allow a small overflow to trigger resize logic
            elif new_y + widget_h > parent_h + 100: 
                new_y = parent_h + 100 - widget_h

            
            if widget.winfo_manager() != 'place':
                self._ensure_placeholder(widget)
                widget.place(width=widget_w, height=widget_h)
                
            widget.place(x=new_x, y=new_y)
            widget.lift()
            
        elif widget._interaction_mode == 'resize':
            new_width = max(50, widget._start_width + dx)
            new_height = max(20, widget._start_height + dy)
            current_x = widget.winfo_x()
            current_y = widget.winfo_y()
            
            # Clamp
            if current_x + new_width > parent_w: new_width = max(50, parent_w - current_x)
            if current_y + new_height > parent_h: new_height = max(20, parent_h - current_y)
            
            if widget.winfo_manager() != 'place':
                self._ensure_placeholder(widget)
                widget.place(x=current_x, y=current_y)
                
            widget.place(width=new_width, height=new_height)
            widget.lift()
        
        return "break"
        
    def _apply_push_down_logic(self, dragged_widget):
        """Recursively push widgets down if they overlap with the dragged widget"""
        try:
            # Current bounding box of the dragging widget
            x1 = dragged_widget.winfo_x()
            y1 = dragged_widget.winfo_y()
            w1 = dragged_widget.winfo_width()
            h1 = dragged_widget.winfo_height()
            
            padding = 10
            
            # Parent of the widgets
            parent = dragged_widget.master
            
            # OPTIMIZATION: Early filter candidates by parent to avoid massive iterations
            candidates = [w for w in self.draggable_items.values() if w.master == parent]
            
            for other in candidates:
                if other == dragged_widget:
                    continue
                
                # [STABILITY FIX] Skip if the 'other' widget's BOTTOM is ABOVE the dragged widget's TOP.
                # Strictly ignore anything physically above the current interaction zone.
                if (other.winfo_y() + other.winfo_height()) <= y1:
                    continue

                # Skip if not in the same parent
                if other.master != parent:
                    continue
                
                # Skip if hidden
                if not other.winfo_ismapped():
                    continue
                
                # Get current pos/size of the 'other' widget
                x2 = other.winfo_x()
                y2 = other.winfo_y()
                w2 = other.winfo_width()
                h2 = other.winfo_height()
                
                # 1. Horizontal Overlap Check
                if (x1 < x2 + w2) and (x1 + w1 > x2):
                    # 2. Vertical Collision Check
                    # If dragged_widget is hitting 'other' from the top
                    # We check if the bottom of dragged_widget is below the top of other
                    if y1 < y2 and (y1 + h1) > y2:
                        # New target Y for 'other'
                        new_y2 = y1 + h1 + padding
                        
                        # Only move if it's actually pushing it DOWN
                        if y2 < new_y2:
                            if other.winfo_manager() == 'grid':
                                # Convert grid to place
                                self._ensure_placeholder(other)
                                other.lift()
                                other.place(x=x2, y=new_y2, width=w2, height=h2)
                            else:
                                # Just update place
                                other.place(y=new_y2)
                            
                            # 3. Recursive Push
                            self._apply_push_down_logic(other)
        except Exception as e:
            # Silent fail to avoid interrupting drag
            print(f"Error in push-down logic: {e}")

    def _adjust_parent_height(self, parent, force=False):
        """Adjust parent frame height with performance check"""
        try:
            # Only update idletasks if forced or we're not in the middle of a high-speed interaction
            # This is the single biggest cause of UI stutter.
            if force:
                parent.update_idletasks()
            
            # 2. Start with the bounding box of all GRIDDED items
            try:
                # grid_bbox returns (x, y, width, height) of the grid
                bbox = parent.grid_bbox()
                required_h = bbox[1] + bbox[3] if bbox[3] > 0 else 0
                required_w = bbox[0] + bbox[2] if bbox[2] > 0 else 0
            except:
                required_h = 0
                required_w = 0

            # 3. Handle PLACE items as well (dragged/custom items)
            for child in parent.winfo_children():
                try:
                    manager = child.winfo_manager()
                    if manager == 'place':
                        info = child.place_info()
                        # Get relative or absolute y + height
                        y = child.winfo_y()
                        h = child.winfo_height()
                        x = child.winfo_x()
                        w = child.winfo_width()
                        required_h = max(required_h, y + h)
                        required_w = max(required_w, x + w)
                except:
                    continue
            
            # Add some padding
            new_height = required_h + 30
            
            # [STABILITY FIX] Guard against collapse but allow growth
            # Width is now dictated by canvas parent in responsive mode.
            if new_height < 500: new_height = 800
            
            # Only resize if the required dimensions are significantly different
            current_h = parent.winfo_height()
            if abs(current_h - new_height) > 10:
                 parent.config(height=new_height)
        except Exception as e:
            print(f"Error adjusting parent height: {e}")


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
        self.daily_usage_paned.pack(fill='both', expand=True, padx=5, pady=5)  # Reduced padding
        
        # Save sash position on adjustment and lock it
        self.daily_usage_paned.bind("<ButtonRelease-1>", self._on_daily_usage_sash_changed)
        self.daily_usage_paned.bind("<Configure>", self._on_daily_usage_resize)
        self.daily_usage_sash_locked = False
        
        # Set initial sash position to ensure visibility (30% for top frame, 70% for bottom)
        self.daily_usage_paned.after(200, self._ensure_daily_usage_sash_visibility)
        self.daily_usage_paned.after(500, self._ensure_daily_usage_sash_visibility)
        self.daily_usage_paned.after(1000, self._ensure_daily_usage_sash_visibility)
        self.daily_usage_paned.after(1200, self._ensure_canvas_scroll_region)

        
        entry_frame = ttk.LabelFrame(self.daily_usage_paned, text="현장별 일일 사용량 기입")
        self.daily_usage_paned.add(entry_frame, weight=1) # Changed from weight=3 to weight=1
        
        # Header for buttons to keep them separate from draggable area
        header_frame = ttk.Frame(entry_frame)
        header_frame.pack(fill='x', padx=2, pady=1)  # Reduced padding
        
        btn_reset_all = ttk.Button(header_frame, text="전체 레이아웃 초기화", command=self.reset_all_widgets_layout)
        btn_reset_all.pack(side='right', padx=5)
        
        self.btn_lock_layout = ttk.Button(header_frame, text="🔓 배치 수정 중", command=self.toggle_layout_lock, style="Lock.TButton")
        self.btn_lock_layout.pack(side='right', padx=5)
        self.style.configure("Lock.TButton", foreground="red")
        
        # Add sash lock toggle button to header
        self.btn_sash_lock = ttk.Button(header_frame, text="🔓 경계 잠금", command=self.toggle_sash_lock)
        self.btn_sash_lock.pack(side='right', padx=5)
        
        # Add resolution display
        self.resolution_label = ttk.Label(header_frame, text="", font=('Malgun Gothic', 8))
        self.resolution_label.pack(side='right', padx=10)
        
        # Add resolution lock button
        self.btn_resolution_lock = ttk.Button(header_frame, text="🔓 해상도 고정", command=self.toggle_resolution_lock)
        self.btn_resolution_lock.pack(side='right', padx=5)
        self.resolution_locked = False
        
        # Update resolution display
        self.update_resolution_display()
        
        
        btn_add_memo = ttk.Button(header_frame, text="➕ 메모 추가", command=self.add_new_memo)
        btn_add_memo.pack(side='right', padx=5)

        btn_add_checklist = ttk.Button(header_frame, text="☑️ 목록 추가", command=self.add_new_checklist)
        btn_add_checklist.pack(side='right', padx=5)

        # [STABILITY FIX] Use a scrollable Canvas to hold the entry form
        # This prevents clipping when content grows or sash is small.
        # Create scrollable frame directly instead of canvas
        canvas_parent = ttk.Frame(entry_frame)
        canvas_parent.pack(fill='both', expand=True, padx=2, pady=1)  # Reduced padding
        
        # Create the entry frame directly without canvas
        self.entry_inner_frame = ttk.Frame(canvas_parent)
        self.entry_inner_frame.pack(fill='both', expand=True)
        
        # No canvas, no scrollbars needed - frame will expand naturally
        self.entry_canvas = None  # Disable canvas-related functionality
        
        # Enable propagation to allow content to dictate size
        self.entry_inner_frame.pack_propagate(True)
        self.entry_inner_frame.grid_propagate(True)
        
        def _on_entry_config(e):
            # Update frame width as content changes
            self.entry_inner_frame.config(width=e.width)
        
        # No canvas - no initial scroll region update needed
        # self.entry_canvas.update_idletasks()
        # self.entry_canvas.configure(scrollregion=self.entry_canvas.bbox("all"))
        
        self.entry_inner_frame.bind("<Configure>", _on_entry_config)
        
        # No canvas - remove all canvas-related code
        # def _on_canvas_wheel(event):
        #     self.entry_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        # 
        # def _bind_canvas_scroll(event):
        #     self.entry_canvas.bind_all("<MouseWheel>", _on_canvas_wheel)
        # def _unbind_canvas_scroll(event):
        #     self.entry_canvas.unbind_all("<MouseWheel>")
        # 
        # self.entry_canvas.bind('<Enter>', _bind_canvas_scroll)
        # self.entry_canvas.bind('<Leave>', _unbind_canvas_scroll)
        
        # Explicitly fix all possible grid rows to weight 0
        for r in range(100):
            self.entry_inner_frame.grid_rowconfigure(r, weight=0)
        
        # Configure the main 2-column container for Split-Pane
        self.entry_inner_frame.grid_columnconfigure(0, weight=1) # Left Pane (Form)
        self.entry_inner_frame.grid_columnconfigure(1, weight=1) # Right Pane (Workers)
        self.entry_inner_frame.grid_rowconfigure(0, weight=1)    # Allow panes to grow vertically

        self.left_pane = ttk.Frame(self.entry_inner_frame)
        self.left_pane.grid(row=0, column=0, sticky='nsew', padx=0, pady=0)  # No padding
        self.right_pane = ttk.Frame(self.entry_inner_frame)
        self.right_pane.grid(row=0, column=1, sticky='nsew', padx=0, pady=0)  # No padding


        # Configure Left Pane columns (let's use 6 columns internally)
        for c in range(6): self.left_pane.grid_columnconfigure(c, weight=0)
        
        # Configure Right Pane columns (2 worker columns)
        for c in range(2): self.right_pane.grid_columnconfigure(c, weight=1)

        # 1. Left Pane: Form Fields
        date_container, self.ent_daily_date = self.create_draggable_container(
            self.left_pane, "날짜:", DateEntry, 'date_box_geometry', 
            width=10, date_pattern='yyyy-mm-dd',
            grid_info={'row': 0, 'column': 0, 'padx': 0, 'pady': 0, 'sticky': 'w'}  # Reduced padding
        )
        self.ent_daily_date.set_date(datetime.datetime.now())
        
        # Create container for site with compact header
        site_container = ttk.Frame(self.left_pane, relief="solid", borderwidth=1)
        site_hdr = ttk.Frame(site_container)
        site_hdr.pack(side='top', fill='x', padx=0, pady=0)
        ttk.Label(site_hdr, text="✥", font=('Arial', 9), cursor='fleur').pack(side='left', padx=1)
        ttk.Label(site_hdr, text="현장:", font=('Malgun Gothic', 9, 'bold')).pack(side='left', padx=1)
        site_btn_box = ttk.Frame(site_hdr)
        site_btn_box.pack(side='right', padx=1)
        ttk.Label(site_btn_box, text="⚙️", font=('Arial', 8), cursor='hand2').pack(side='left', padx=1)
        site_btn_box.winfo_children()[0].bind('<Button-1>', lambda e: self.open_list_management_dialog('sites'))
        ttk.Label(site_btn_box, text="❌", font=('Arial', 8), cursor='hand2').pack(side='left', padx=1)
        site_btn_box.winfo_children()[1].bind('<Button-1>', lambda e: self.remove_box('site_box_geometry'))
        site_content = ttk.Frame(site_container, padding=(1, 0))
        site_content.pack(side='top', fill='both', expand=True)
        self.ent_daily_site = ttk.Combobox(site_content, width=15, values=self.sites)
        self.ent_daily_site.pack(side='left')
        site_container.grid(row=0, column=1, columnspan=2, padx=1, pady=1, sticky='w')
        self.make_draggable(site_container, 'site_box_geometry')
        self.draggable_items['site_box_geometry'] = site_container

        # Row 1: Equip & Material & Sync
        equip_container, self.cb_daily_equip = self.create_draggable_container(
            self.left_pane, "장비:", ttk.Combobox, 'equip_box_geometry', 
            manage_list_key='equipments', width=12, values=getattr(self, 'equipments', []),
            grid_info={'row': 1, 'column': 0, 'padx': 1, 'pady': 1, 'sticky': 'w'}
        )
        
        mat_container, self.cb_daily_material = self.create_draggable_container(
            self.left_pane, "품목:", ttk.Combobox, 'mat_box_geometry', width=20,
            grid_info={'row': 1, 'column': 1, 'columnspan': 2, 'padx': 1, 'pady': 1, 'sticky': 'w'}
        )

        sync_container = ttk.Frame(self.left_pane, relief="solid", borderwidth=1)
        sync_btn_box = ttk.Frame(sync_container)
        sync_btn_box.pack(side='top', fill='x')
        ttk.Label(sync_btn_box, text="✥", font=('Arial', 8)).pack(side='left', padx=1)
        sync_content = ttk.Frame(sync_container, padding=(1, 0))
        sync_content.pack(side='top', fill='both')
        ttk.Button(sync_content, text="일괄", command=self.sync_worker_times, width=4).pack()
        sync_container.grid(row=1, column=3, padx=1, pady=1, sticky='w')
        self.make_draggable(sync_container, 'sync_box_geometry')
        self.draggable_items['sync_box_geometry'] = sync_container

        # Row 2: Consolidated Results
        method_container, self.cb_daily_test_method = self.create_draggable_container(
            self.left_pane, "방법:", ttk.Combobox, 'method_box_geometry', width=6, 
            values=["RT", "PAUT", "UT", "MT", "PT", "PMI"],
            grid_info={'row': 2, 'column': 0, 'padx': 1, 'pady': 1, 'sticky': 'w'}
        )
        amount_container, self.ent_daily_test_amount = self.create_draggable_container(
            self.left_pane, "량:", ttk.Entry, 'amount_box_geometry', width=6,
            grid_info={'row': 2, 'column': 1, 'padx': 1, 'pady': 1, 'sticky': 'w'}
        )
        self.ent_daily_test_amount.insert(0, "0")
        u_price_container, self.ent_daily_unit_price = self.create_draggable_container(
            self.left_pane, "단가:", ttk.Entry, 'u_price_box_geometry', width=8,
            grid_info={'row': 2, 'column': 2, 'padx': 1, 'pady': 1, 'sticky': 'w'}
        )
        self.ent_daily_unit_price.insert(0, "0")
        travel_container, self.ent_daily_travel_cost = self.create_draggable_container(
            self.left_pane, "출장:", ttk.Entry, 'travel_box_geometry', width=8,
            grid_info={'row': 2, 'column': 3, 'padx': 1, 'pady': 1, 'sticky': 'w'}
        )
        self.ent_daily_travel_cost.insert(0, "0")
        meal_container, self.ent_daily_meal_cost = self.create_draggable_container(
            self.left_pane, "일식:", ttk.Entry, 'meal_box_geometry', width=8,
            grid_info={'row': 2, 'column': 4, 'padx': 1, 'pady': 1, 'sticky': 'w'}
        )
        self.ent_daily_meal_cost.insert(0, "0")
        fee_container, self.ent_daily_test_fee = self.create_draggable_container(
            self.left_pane, "검사비:", ttk.Entry, 'fee_box_geometry', width=8,
            grid_info={'row': 2, 'column': 5, 'padx': 1, 'pady': 1, 'sticky': 'w'}
        )
        self.ent_daily_test_fee.insert(0, "0")
        film_container, self.ent_film_count = self.create_draggable_container(
            self.left_pane, "필름:", ttk.Entry, 'film_box_geometry', width=4,
            grid_info={'row': 2, 'column': 6, 'padx': 1, 'pady': 1, 'sticky': 'w'}
        )
        self.ent_film_count.insert(0, "0")

        # Row 3: NDT, Row 4: RTK (Left Pane - narrower than workers)
        ndt_container = ttk.Frame(self.left_pane, relief="solid", borderwidth=1)
        ndt_hdr = ttk.Frame(ndt_container)
        ndt_hdr.pack(fill='x', side='top'); ttk.Label(ndt_hdr, text="✥", font=('Arial', 9)).pack(side='left', padx=1)
        lbl_ndt_title = ttk.Label(ndt_hdr, text="NDT 자재", font=('Malgun Gothic', 9, 'bold'))
        lbl_ndt_title.pack(side='left', padx=1); ndt_content = ttk.Frame(ndt_container, padding=(1, 0))
        ndt_content.pack(fill='both', expand=True); ndt_grid = ttk.LabelFrame(ndt_content)
        ndt_grid.pack(fill='both', expand=True, padx=1, pady=1)
        
        # Add a specific resize handle to the NDT header as well
        btn_ndt_resize = ttk.Label(ndt_hdr, text="⚙️", font=('Arial', 8), cursor='hand2')
        btn_ndt_resize.pack(side='right', padx=2)
        from functools import partial
        btn_ndt_resize.bind("<Button-1>", partial(self.on_resize_start, widget=ndt_container))
        btn_ndt_resize.bind("<B1-Motion>", partial(self.on_mouse_motion, widget=ndt_container))
        btn_ndt_resize.bind("<ButtonRelease-1>", partial(self.on_drag_stop, widget=ndt_container))

        self.ndt_entries = {}
        ndt_materials = ["형광자분", "흑색자분", "백색페인트", "침투제", "세척제", "현상제", "형광침투제"]
        for i, mat in enumerate(ndt_materials):
            r = i // 4; col = (i % 4) * 2
            ttk.Label(ndt_grid, text=f"{mat}:", font=('Arial', 8)).grid(row=r, column=col, padx=1, pady=1, sticky='w')
            e = ttk.Entry(ndt_grid, width=6); e.grid(row=r, column=col+1, padx=1, pady=1, sticky='ew'); self.ndt_entries[mat] = e
        ndt_container.grid(row=3, column=0, columnspan=7, padx=1, pady=1, sticky='nw')
        self.make_draggable(ndt_container, 'ndt_usage_box_geometry'); self.draggable_items['ndt_usage_box_geometry'] = ndt_container

        rtk_container = ttk.Frame(self.left_pane, relief="solid", borderwidth=1)
        rtk_hdr = ttk.Frame(rtk_container)
        rtk_hdr.pack(fill='x', side='top'); ttk.Label(rtk_hdr, text="✥", font=('Arial', 9)).pack(side='left', padx=1)
        lbl_rtk_title = ttk.Label(rtk_hdr, text="RT 매수", font=('Malgun Gothic', 9, 'bold'))
        lbl_rtk_title.pack(side='left', padx=1); rtk_content = ttk.Frame(rtk_container, padding=(1, 0))
        rtk_content.pack(fill='both', expand=True); rtk_grid = ttk.LabelFrame(rtk_content)
        rtk_grid.pack(fill='both', expand=True, padx=1, pady=1)
        
        # Add a specific resize handle to the header
        btn_rtk_resize = ttk.Label(rtk_hdr, text="⚙️", font=('Arial', 8), cursor='hand2')
        btn_rtk_resize.pack(side='right', padx=2)
        from functools import partial
        btn_rtk_resize.bind("<Button-1>", partial(self.on_resize_start, widget=rtk_container))
        btn_rtk_resize.bind("<B1-Motion>", partial(self.on_mouse_motion, widget=rtk_container))
        btn_rtk_resize.bind("<ButtonRelease-1>", partial(self.on_drag_stop, widget=rtk_container))

        self.rtk_entries = {}
        rtk_cats = ["센터미스", "농도", "마킹미스", "필름마크", "취급부주의", "고객불만", "기타", "총계"]
        for i, cat in enumerate(rtk_cats):
            r = i // 4; col = (i % 4) * 2
            ttk.Label(rtk_grid, text=f"{cat}:", font=('Arial', 8)).grid(row=r, column=col, padx=1, pady=1, sticky='w')
            e = ttk.Entry(rtk_grid, width=6)
            e.grid(row=r, column=col+1, padx=1, pady=1, sticky='ew')
            self.rtk_entries[cat] = e
            
            # Bind auto-calculation
            if cat != "총계":
                e.bind('<KeyRelease>', lambda e: self.calculate_rtk_total())
        
        self.rtk_entries["총계"].config(state='readonly')
        rtk_container.grid(row=4, column=0, columnspan=7, padx=1, pady=1, sticky='nw')
        self.make_draggable(rtk_container, 'rtk_usage_box_geometry'); self.draggable_items['rtk_usage_box_geometry'] = rtk_container

        # Row 5: Note & Save
        note_container, self.ent_daily_note = self.create_draggable_container(
            self.left_pane, "비고:", ttk.Entry, 'note_box_geometry', width=45,
            grid_info={'row': 5, 'column': 0, 'columnspan': 4, 'padx': 0, 'pady': 0, 'sticky': 'nw'}  # Reduced padding
        )
        save_btn_container = ttk.Frame(self.left_pane)
        ttk.Button(save_btn_container, text="저장", command=self.add_daily_usage_entry, width=8, style='Big.TButton').pack()
        save_btn_container.grid(row=5, column=4, columnspan=2, padx=0, pady=0, sticky='nw')  # Reduced padding
        self.make_draggable(save_btn_container, 'save_btn_geometry'); self.draggable_items['save_btn_geometry'] = save_btn_container

        # 2. Right Pane: Worker Groups (1-10)
        time_presets = [
            "08:00~17:00", "08:00~18:00", "08:00~19:00", 
            "18:00~21:00", "18:00~22:00", "18:00~23:00", "18:00~24:00",
            "18:00~01:00", "18:00~02:00", "18:00~03:00"
        ]
        
        def setup_worker_group(idx, row, col):
            config_key = f'worker_group{idx}_geometry'
            container, group = self.create_draggable_container(
                self.right_pane, f"작업자 {idx}:", WorkerDataGroup, config_key,
                worker_index=idx, users_list=getattr(self, 'users', []),
                enable_autocomplete=True, # Enable autocomplete for worker names
                time_list=time_presets,
                manage_list_key='users', grid_info={'row': row, 'column': col, 'padx': 0, 'pady': 0, 'sticky': 'w'}  # Reduced padding
            )
            # Store global references for all workers (1-10)
            if idx == 1:
                self.cb_daily_user = group.composite
                self.ent_worktime1 = group.ent_worktime
                self.ent_ot1 = group.ent_ot
            
            # Always set index-specific attributes for sync/data access
            setattr(self, f'cb_daily_user{idx}', group.composite)
            setattr(self, f'ent_worktime{idx}', group.ent_worktime)
            setattr(self, f'ent_ot{idx}', group.ent_ot)
            
            # Bindings for auto-save & auto-OT
            group.bind_name('<FocusOut>', lambda e: self.auto_save_to_list(e, group.cb_name, self.users, 'users'))
            group.bind_name('<Return>', lambda e: self.auto_save_to_list(e, group.cb_name, self.users, 'users'))
            
            group.bind_time('<FocusOut>', lambda e: self.auto_save_worktime(e, group.ent_worktime, 'worktimes'))
            group.bind_time('<Return>', lambda e: self.auto_save_worktime(e, group.ent_worktime, 'worktimes'))
            
            group.bind_ot('<FocusOut>', lambda e: self.auto_save_ot(e, group.ent_ot, 'ot_times'))
            group.bind_ot('<Return>', lambda e: self.auto_save_ot(e, group.ent_ot, 'ot_times'))
            
            return container, group

        # 5 rows x 2 cols = 10 workers
        for i in range(1, 6): setup_worker_group(i, i-1, 0)
        for i in range(6, 11): setup_worker_group(i, i-6, 1)

        # Bindings & Finalization
        calc_trigger = lambda e: self.update_daily_test_fee_calc()
        
        def on_qty_change(e):
            self.update_daily_test_fee_calc()
            self.sync_film_with_quantity()
            
        self.ent_daily_test_amount.bind('<KeyRelease>', on_qty_change)
        self.cb_daily_test_method.bind('<<ComboboxSelected>>', self.sync_film_with_quantity)
        
        self.ent_daily_unit_price.bind('<KeyRelease>', calc_trigger)
        self.ent_daily_travel_cost.bind('<KeyRelease>', calc_trigger)
        self.ent_daily_meal_cost.bind('<KeyRelease>', calc_trigger)
        
        self.update_material_combo()
        self._adjust_parent_height(self.entry_inner_frame, force=True)
        
        display_frame = ttk.LabelFrame(self.daily_usage_paned, text="일일 사용량 기록 조회")
        self.daily_usage_paned.add(display_frame, weight=1) # Less weight for the list
        
        # Filter controls
        filter_frame = ttk.Frame(display_frame)
        filter_frame.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(filter_frame, text="시작일:").pack(side='left', padx=5)
        self.ent_daily_start_date = DateEntry(filter_frame, width=12, date_pattern='yyyy-mm-dd')
        self.ent_daily_start_date.pack(side='left', padx=5)
        # Default to 7 days ago
        start_date = (datetime.datetime.now() - datetime.timedelta(days=7))
        self.ent_daily_start_date.set_date(start_date)
        
        ttk.Label(filter_frame, text="종료일:").pack(side='left', padx=5)
        self.ent_daily_end_date = DateEntry(filter_frame, width=12, date_pattern='yyyy-mm-dd')
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

        ttk.Label(filter_frame, text="작업자:").pack(side='left', padx=5)
        self.cb_daily_filter_worker = ttk.Combobox(filter_frame, width=15, state="readonly")
        self.cb_daily_filter_worker.pack(side='left', padx=5)
        self.cb_daily_filter_worker.set('전체')
        
        btn_filter = ttk.Button(filter_frame, text="조회", command=self.update_daily_usage_view)
        btn_filter.pack(side='left', padx=10)
        
        btn_delete = ttk.Button(filter_frame, text="선택 항목 삭제", command=self.delete_daily_usage_entry)
        btn_delete.pack(side='left', padx=10)
        
        btn_export = ttk.Button(filter_frame, text="엑셀 내보내기", command=self.export_daily_usage_history)
        btn_export.pack(side='left', padx=5)
        
        btn_export_all = ttk.Button(filter_frame, text="전체 기록 내보내기", command=self.export_all_daily_usage)
        btn_export_all.pack(side='left', padx=5)
        
        btn_col_manage = ttk.Button(filter_frame, text="컬럼 관리", command=self.show_column_visibility_dialog)
        btn_col_manage.pack(side='left', padx=10)
        
        # Treeview for daily usage records
        tree_container = ttk.Frame(display_frame)
        tree_container.pack(expand=True, fill='both', padx=5, pady=5)
        
        # Horizontal PanedWindow for List vs Details
        self.daily_history_paned = ttk.Panedwindow(tree_container, orient='horizontal')
        self.daily_history_paned.pack(fill='both', expand=True)
        
        # Save sash position on adjustment
        self.daily_history_paned.bind("<ButtonRelease-1>", lambda e: self.save_tab_config())


        list_frame = ttk.Frame(self.daily_history_paned)
        self.daily_history_paned.add(list_frame, weight=3)

        # Scrollbars
        vsb = ttk.Scrollbar(list_frame, orient="vertical")
        hsb = ttk.Scrollbar(list_frame, orient="horizontal")
        
        # Treeview with RTK categories and NDT materials
        # Note: Workers 1-10 columns are kept in the 'columns' tuple for data storage,
        # but we will only show a consolidated '작업자' in 'displaycolumns'.
        columns = ('날짜', '현장', '작업자', '작업시간', 'OT1', 'OT2', 'OT3', 'OT4', 'OT5', 'OT6', 'OT7', 'OT8', 'OT9', 'OT10', '장비명', '검사방법', '검사량', '필름매수', '단가', '출장비', '일식', '검사비', 'OT시간', 'OT금액', '품목명', '센터미스', '농도', '마킹미스', '필름마크', '취급부주의', '고객불만', '기타', 'RT총계', '형광자분', '흑색자분', '백색페인트', '침투제', '세척제', '현상제', '형광침투제', '비고', '입력시간')
        self.daily_usage_tree = ttk.Treeview(list_frame, columns=columns, show='headings',
                                              yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        vsb.config(command=self.daily_usage_tree.yview)
        hsb.config(command=self.daily_usage_tree.xview)
        
        # Column configuration
        col_widths = {
            '날짜': 160, '현장': 130, '작업자': 170, '작업시간': 120,
            'OT1': 140, 'OT2': 140, 'OT3': 140, 'OT4': 140, 'OT5': 140, 'OT6': 140,
            'OT7': 140, 'OT8': 140, 'OT9': 140, 'OT10': 140,
            '장비명': 130, '검사방법': 90, 
            '검사량': 80, '필름매수': 80, '단가': 90, '출장비': 90, '일식': 80, 
            '검사비': 100, 'OT시간': 80, 'OT금액': 100, '품목명': 210, '센터미스': 70, '농도': 70, '마킹미스': 70, 
            '필름마크': 70, '취급부주의': 70, '고객불만': 70, '기타': 70, 'RT총계': 80, 
            '형광자분': 80, '흑색자분': 80, '백색페인트': 80, '침투제': 80, '세척제': 80, 
            '현상제': 80, '형광침투제': 80, '비고': 230, '입력시간': 300
        }
        
        for col in columns:
            self.daily_usage_tree.heading(col, text=col, command=lambda c=col: self.treeview_sort_column(self.daily_usage_tree, c, False))
            width = col_widths.get(col, 100)
            self.daily_usage_tree.column(col, width=width, minwidth=20, stretch=False, anchor='center')
        
        # Grid layout for list_frame
        self.daily_usage_tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)
        
        # Auto-save column widths when user resizes columns
        def save_column_widths(event=None):
            self.save_tab_config()
        
        self.daily_usage_tree.bind('<ButtonRelease-1>', save_column_widths)

        # Right side: Details Panel
        detail_frame = ttk.LabelFrame(self.daily_history_paned, text="상세 정보")
        self.daily_history_paned.add(detail_frame, weight=1)

        self.daily_detail_text = tk.Text(detail_frame, wrap='word', width=35, font=('Pretendard', 14), bg=self.theme_bg, highlightthickness=0)
        self.daily_detail_text.pack(fill='both', expand=True, padx=5, pady=5)
        self.daily_detail_text.config(state='disabled')
        
        # Selection binding
        self.daily_usage_tree.bind('<<TreeviewSelect>>', self.on_daily_usage_tree_select)
        
        # Optimize treeview scroll region to prevent unnecessary scrollbars
        def optimize_treeview_scroll_region():
            try:
                if hasattr(self, 'daily_usage_tree'):
                    self.daily_usage_tree.update_idletasks()
                    
                    # Get treeview content dimensions
                    total_items = len(self.daily_usage_tree.get_children())
                    if total_items > 0:
                        # Calculate approximate content height
                        item_height = 25  # Approximate height per row
                        content_height = total_items * item_height + 50  # Add some padding
                        
                        # Get treeview dimensions
                        tree_width = self.daily_usage_tree.winfo_width()
                        tree_height = self.daily_usage_tree.winfo_height()
                        
                        # Set scroll region to content size if smaller than treeview
                        if content_height < tree_height:
                            self.daily_usage_tree.configure(yscrollcommand=(0, 0, tree_width, content_height))
                            print(f"Treeview scroll region optimized: (0, 0, {tree_width}, {content_height})")
                        else:
                            # Use full treeview dimensions
                            self.daily_usage_tree.configure(yscrollcommand=(0, 0, tree_width, tree_height))
                            print(f"Treeview scroll region set to full size: (0, 0, {tree_width}, {tree_height})")
            except Exception as e:
                print(f"Error optimizing treeview scroll region: {e}")
        
        # Apply optimization after a short delay
        self.root.after(500, optimize_treeview_scroll_region)
        
        # Set a default sash position after UI is drawn if not loaded from config
        self.root.after(100, self._ensure_sash_visible)
        
        # Initial view update
        self.update_daily_usage_view()

    def _ensure_sash_visible(self):
        """Ensure the history details panel is visible by default if not set by config"""
        # If we already restored a valid position from config, don't force it again
        if getattr(self, '_sash_restored', False):
            return

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

                    current_pos = 0 # Assume hidden if errored

                # Only set if it's currently at 0 (hidden) or near the end (effectively hidden)
                if current_pos < 50 or current_pos > total_w - 50:
                    new_pos = int(total_w * 0.75) if total_w > 200 else 600
                    self.daily_history_paned.sashpos(0, new_pos)
        except Exception as e:
            print(f"Error ensuring sash visibility: {e}")
    
    def _ensure_daily_usage_sash_visibility(self):
        """Ensure sash position is visible and properly sized using ratio"""
        try:
            if not hasattr(self, 'daily_usage_paned'):
                return
                
            # [STABILITY] Skip if we are still in the initial loading/restoration phase
            # load_tab_config will handle the initial placement.
            if not getattr(self, 'is_ready', False):
                return

            # Force multiple updates to get correct dimensions
            for _ in range(3):
                self.daily_usage_paned.update_idletasks()
            
            # Get current paned window dimensions
            total_h = self.daily_usage_paned.winfo_height()
            
            if total_h <= 0:
                print("PanedWindow height is 0, retrying...")
                self.root.after(100, self._ensure_daily_usage_sash_visibility)
                return
            
            # If locked, prioritize absolute position if valid
            if hasattr(self, 'daily_usage_sash_locked') and self.daily_usage_sash_locked:
                if hasattr(self, 'tab_config') and 'daily_usage_sash_pos' in self.tab_config:
                    target_pos = int(self.tab_config['daily_usage_sash_pos'])
                    # Ensure within bounds
                    if 50 < target_pos < total_h - 50:
                        self.daily_usage_paned.sashpos(0, target_pos)
                        print(f"Restored locked absolute position: {target_pos}")
                        return

            # Try to restore from saved ratio first
            target_ratio = 0.25  # Default ratio
            if hasattr(self, 'tab_config') and 'daily_usage_sash_ratio' in self.tab_config:
                target_ratio = self.tab_config['daily_usage_sash_ratio']
            
            # Calculate target position
            target_pos = int(total_h * target_ratio)
            
            # Apply relaxed bounds (10% to 90% of total height)
            min_pos = int(total_h * 0.1)
            max_pos = int(total_h * 0.9)
            
            if target_pos < min_pos:
                target_pos = min_pos
            elif target_pos > max_pos:
                target_pos = max_pos
            
            # Set sash position
            self.daily_usage_paned.sashpos(0, target_pos)
            
            # Verify position was set correctly
            actual_pos = self.daily_usage_paned.sashpos(0)
            actual_ratio = actual_pos / total_h if total_h > 0 else target_ratio
            
            # Update config
            if hasattr(self, 'is_ready') and self.is_ready:
                if not hasattr(self, 'tab_config'):
                    self.tab_config = {}
                self.tab_config['daily_usage_sash_ratio'] = actual_ratio
                self.tab_config['daily_usage_sash_pos'] = actual_pos
            
            # Update canvas scroll region
            self.root.after(100, self._ensure_canvas_scroll_region)
            
        except Exception as e:
            print(f"Error ensuring sash visibility: {e}")
            self.root.after(500, self._ensure_daily_usage_sash_visibility)
    
    
    
    def update_resolution_display(self):
        """Update the resolution display"""
        try:
            if hasattr(self, 'resolution_label'):
                # Get window dimensions
                window_width = self.root.winfo_width()
                window_height = self.root.winfo_height()
                
                # Get screen dimensions
                screen_width = self.root.winfo_screenwidth()
                screen_height = self.root.winfo_screenheight()
                
                # Get daily usage paned window dimensions
                if hasattr(self, 'daily_usage_paned'):
                    self.daily_usage_paned.update_idletasks()
                    pane_height = self.daily_usage_paned.winfo_height()
                    pane_width = self.daily_usage_paned.winfo_width()
                    
                    # Get current sash position and ratio
                    try:
                        sash_pos = self.daily_usage_paned.sashpos(0)
                        ratio = (sash_pos / pane_height * 100) if pane_height > 0 else 0
                        resolution_text = f"창: {window_width}x{window_height} | 화면: {screen_width}x{screen_height} | 패널: {pane_width}x{pane_height} | 경계: {ratio:.1f}%"
                    except:
                        resolution_text = f"창: {window_width}x{window_height} | 화면: {screen_width}x{screen_height} | 패널: {pane_width}x{pane_height}"
                else:
                    resolution_text = f"창: {window_width}x{window_height} | 화면: {screen_width}x{screen_height}"
                
                self.resolution_label.config(text=resolution_text)
                
                # Schedule next update
                self.root.after(500, self.update_resolution_display)
        except Exception as e:
            print(f"Error updating resolution display: {e}")
            # Still schedule next update even if there's an error
            if hasattr(self, 'resolution_label'):
                self.root.after(1000, self.update_resolution_display)
    
    def _restore_locked_position(self):
        """Restore sash to locked position"""
        try:
            if not hasattr(self, 'daily_usage_paned'): return
            
            if hasattr(self, 'tab_config'):
                total_h = self.daily_usage_paned.winfo_height()
                if total_h <= 0: return
                
                # Prioritize absolute position when locked
                if 'daily_usage_sash_pos' in self.tab_config:
                    pos = int(self.tab_config['daily_usage_sash_pos'])
                    # Only apply if it doesn't hide the bottom area completely
                    if 50 < pos < total_h - 50:
                        self.daily_usage_paned.sashpos(0, pos)
                        return
                
                # Fallback to ratio
                if 'daily_usage_sash_ratio' in self.tab_config:
                    ratio = self.tab_config['daily_usage_sash_ratio']
                    locked_pos = int(total_h * ratio)
                    self.daily_usage_paned.sashpos(0, locked_pos)
        except Exception as e:
            print(f"Error restoring locked position: {e}")
    
    def _on_daily_usage_resize(self, event):
        """Handle window resize to maintain sash ratio or absolute position if locked"""
        try:
            if not hasattr(self, 'daily_usage_paned'): return
            
            # If locked, maintain absolute position from top
            if hasattr(self, 'daily_usage_sash_locked') and self.daily_usage_sash_locked:
                self._restore_locked_position()
                return

            # Otherwise maintain ratio
            if hasattr(self, 'tab_config') and 'daily_usage_sash_ratio' in self.tab_config:
                ratio = self.tab_config['daily_usage_sash_ratio']
                total_h = self.daily_usage_paned.winfo_height()
                
                if total_h > 200:
                    new_pos = int(total_h * ratio)
                    min_pos, max_pos = 50, total_h - 50
                    new_pos = max(min_pos, min(new_pos, max_pos))
                    
                    self.daily_usage_paned.sashpos(0, new_pos)
        except Exception as e:
            print(f"Error handling resize: {e}")
    
    def toggle_sash_lock(self):
        """Toggle sash lock state"""
        try:
            self.daily_usage_sash_locked = not self.daily_usage_sash_locked
            
            if self.daily_usage_sash_locked:
                # If just locked, save current state
                sash_pos = self.daily_usage_paned.sashpos(0)
                total_height = self.daily_usage_paned.winfo_height()
                
                if not hasattr(self, 'tab_config'):
                    self.tab_config = {}
                
                # Calculate and save ratio (percentage)
                ratio = sash_pos / total_height if total_height > 0 else 0.2
                self.tab_config['daily_usage_sash_ratio'] = ratio
                self.tab_config['daily_usage_sash_pos'] = sash_pos  # Keep as backup
                self.tab_config['daily_usage_sash_locked'] = True
                
                # Save configuration immediately
                self.save_tab_config()
                self.force_save_config()
                
                if hasattr(self, 'btn_sash_lock'):
                    self.btn_sash_lock.config(text="🔒 경계 고정됨")
                    self.btn_sash_lock.configure(style="SashLock.TButton")
                    self.style.configure("SashLock.TButton", foreground="red")
                
                print(f"Daily usage sash position LOCKED at ratio: {ratio:.3f}")
                # Start periodic monitoring
                self._start_sash_monitor()
            else:
                if not hasattr(self, 'tab_config'):
                    self.tab_config = {}
                self.tab_config['daily_usage_sash_locked'] = False
                self.save_tab_config()
                self.force_save_config()

                if hasattr(self, 'btn_sash_lock'):
                    self.btn_sash_lock.config(text="🔓 경계 고정")
                    self.btn_sash_lock.configure(style="TButton")
                print("Daily usage sash UNLOCKED")
                # Stop periodic monitoring
                self._stop_sash_monitor()
                
        except Exception as e:
            print(f"Error toggling sash lock: {e}")
    
    def _start_sash_monitor(self):
        """Start periodic monitoring of sash position"""
        if hasattr(self, '_sash_monitor_job'):
            self.root.after_cancel(self._sash_monitor_job)
        
        def check_sash():
            if hasattr(self, 'daily_usage_sash_locked') and self.daily_usage_sash_locked:
                self._restore_locked_position()
                # Schedule next check
                self._sash_monitor_job = self.root.after(200, check_sash)
        
        # Start monitoring
        self._sash_monitor_job = self.root.after(200, check_sash)
        print("Started sash position monitoring")
    
    def _stop_sash_monitor(self):
        """Stop periodic monitoring of sash position"""
        if hasattr(self, '_sash_monitor_job'):
            self.root.after_cancel(self._sash_monitor_job)
            delattr(self, '_sash_monitor_job')
            print("Stopped sash position monitoring")
    
    def _on_main_window_resize(self, event):
        """Handle main window resize to maintain all sash ratios"""
        try:
            # Only process for actual window resize, not widget events
            if event.widget == self.root:
                # Check if daily usage tab exists and has saved ratio
                if hasattr(self, 'daily_usage_paned') and hasattr(self, 'tab_config') and 'daily_usage_sash_ratio' in self.tab_config:
                    self.root.after(100, self._ensure_daily_usage_sash_visibility)
                
                # Check if inout tab exists and has saved ratio
                if hasattr(self, 'inout_paned') and hasattr(self, 'tab_config') and 'inout_sash_ratio' in self.tab_config:
                    self.root.after(100, self._ensure_inout_sash_visibility)
                    
        except Exception as e:
            print(f"Error handling main window resize: {e}")
    
    def on_tab_changed(self, event):
        """Handle tab change events to ensure proper visibility"""
        try:
            current_tab = self.notebook.select()
            if current_tab and str(current_tab) == str(self.tab_daily_usage):
                print("Daily usage tab selected - ensuring visibility")
                # Force multiple updates when tab is selected
                self.root.after(100, self._ensure_daily_usage_sash_visibility)
                self.root.after(300, self._ensure_daily_usage_sash_visibility)
                self.root.after(500, self._ensure_canvas_scroll_region)
        except Exception as e:
            print(f"Error in tab change handler: {e}")
    
    def show_error_dialog(self, title, message):
        """Show a custom error dialog with draggable text"""
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry("600x400")
        dialog.resizable(True, True)
        
        # Make dialog modal
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (600 // 2)
        y = (dialog.winfo_screenheight() // 2) - (400 // 2)
        dialog.geometry(f"600x400+{x}+{y}")
        
        # Main frame
        main_frame = ttk.Frame(dialog)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Title label
        title_label = ttk.Label(main_frame, text=title, font=('Malgun Gothic', 12, 'bold'))
        title_label.pack(pady=(0, 10))
        
        # Text widget with scrollbar for error message
        text_frame = ttk.Frame(main_frame)
        text_frame.pack(fill='both', expand=True)
        
        text_widget = tk.Text(text_frame, wrap='word', font=('Malgun Gothic', 10))
        scrollbar = ttk.Scrollbar(text_frame, orient='vertical', command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        text_widget.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # Insert error message
        text_widget.insert('1.0', message)
        text_widget.configure(state='normal')  # Allow selection and copying
        
        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill='x', pady=(10, 0))
        
        # Copy button
        def copy_text():
            dialog.clipboard_clear()
            selected_text = text_widget.get('sel.first', 'sel.last') if text_widget.tag_ranges('sel') else text_widget.get('1.0', 'end')
            dialog.clipboard_append(selected_text)
        
        copy_btn = ttk.Button(button_frame, text="복사하기", command=copy_text)
        copy_btn.pack(side='left', padx=5)
        
        # Close button
        close_btn = ttk.Button(button_frame, text="닫기", command=dialog.destroy)
        close_btn.pack(side='right', padx=5)
        
        # Focus on text widget
        text_widget.focus_set()
        
        # Bind Escape key to close
        dialog.bind('<Escape>', lambda e: dialog.destroy())
        
        # Wait for dialog to close
        dialog.wait_window()

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
                ('일식', '일식'), ('검사비', '검사비'),
                ('User', '작업자1'), ('WorkTime', '작업시간1'), ('OT', 'OT1'),
                ('User2', '작업자2'), ('WorkTime2', '작업시간2'), ('OT2', 'OT2'),
                ('User3', '작업자3'), ('WorkTime3', '작업시간3'), ('OT3', 'OT3'),
                ('User4', '작업자4'), ('WorkTime4', '작업시간4'), ('OT4', 'OT4'),
                ('User5', '작업자5'), ('WorkTime5', '작업시간5'), ('OT5', 'OT5'),
                ('User6', '작업자6'), ('WorkTime6', '작업시간6'), ('OT6', 'OT6'),
                ('User7', '작업자7'), ('WorkTime7', '작업시간7'), ('OT7', 'OT7'),
                ('User8', '작업자8'), ('WorkTime8', '작업시간8'), ('OT8', 'OT8'),
                ('User9', '작업자9'), ('WorkTime9', '작업시간9'), ('OT9', 'OT9'),
                ('User10', '작업자10'), ('WorkTime10', '작업시간10'), ('OT10', 'OT10'),
                ('Note', '비고'), ('Entry Time', '입력시간')
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
            mat_id = entry.get('MaterialID')
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
            
            self.daily_detail_text.tag_configure("header", font=('Arial', 16, 'bold'), foreground='blue')
            self.daily_detail_text.tag_configure("section", font=('Arial', 15, 'bold'), foreground='green')
            self.daily_detail_text.tag_configure("label", font=('Arial', 14, 'bold'))
            self.daily_detail_text.tag_configure("value", font=('Arial', 14))
            
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
            # Format as integer if possible, else float
            if total == int(total):
                self.rtk_entries["총계"].insert(0, str(int(total)))
            else:
                self.rtk_entries["총계"].insert(0, f"{total:.1f}")
            self.rtk_entries["총계"].config(state='readonly')
        except ValueError:
            pass  # Ignore invalid input during typing
    
    def sync_film_with_quantity(self, event=None):
        """Sync film count with quantity if method is RT"""
        try:
            method = self.cb_daily_test_method.get().strip().upper()
            if "RT" in method:
                qty = self.ent_daily_test_amount.get().strip()
                self.ent_film_count.delete(0, tk.END)
                self.ent_film_count.insert(0, qty)
        except Exception as e:
            print(f"Error syncing film with quantity: {e}")

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

    def sync_worker_times(self):
        """작업자 1의 작업시간/OT를 모든 작업자와 동기화"""
        try:
            wt1 = self.ent_worktime1.get().strip()
            ot1 = self.ent_ot1.get().strip()
            
            if not wt1 and not ot1:
                messagebox.showwarning("입력 필요", "작업자 1의 작업시간이나 OT를 입력해주세요.")
                return

            for i in range(2, 11):
                user_attr = f'cb_daily_user{i}'
                wt_attr = f'ent_worktime{i}'
                ot_attr = f'ent_ot{i}'
                
                # Check if this worker slot exists
                if not hasattr(self, user_attr):
                    continue
                    
                worker_widget = getattr(self, user_attr)
                
                # [USER REQUEST] Only apply to workers with names selected
                target_name = worker_widget.cb_name.get().strip()
                if not target_name:
                    continue

                # Sync Shift
                worker_widget.cb_shift.set(self.cb_daily_user.cb_shift.get())

                # Sync Work Time
                if hasattr(self, wt_attr):
                    getattr(self, wt_attr).delete(0, tk.END)
                    getattr(self, wt_attr).insert(0, wt1)
                
                # Sync OT
                if hasattr(self, ot_attr):
                    getattr(self, ot_attr).delete(0, tk.END)
                    getattr(self, ot_attr).insert(0, ot1)
            
            messagebox.showinfo("완료", "작업자 1의 시간/OT가 모든 작업자에게 적용되었습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"동기화 중 오류가 발생했습니다: {e}")

    def add_daily_usage_entry(self):
        """Add a daily usage entry"""
        try:
            date_str = self.ent_daily_date.get()
            site = self.ent_daily_site.get()
            mat_name = self.cb_daily_material.get()
            film_count_str = self.ent_film_count.get()
            note = self.ent_daily_note.get()
            ndt_usage = {} # Initialize NDT usage dictionary
            
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

            # Robust MaterialID Lookup
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
                    mat_id = m_row['MaterialID']
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
                    mat_id = mat_rows['MaterialID'].values[0]
            
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
                new_mat_id = self.materials_df['MaterialID'].max() + 1 if not self.materials_df.empty else 1
                
                # More inclusive brand-to-category mapping for new materials
                default_cat = 'OTHER'
                name_up = pure_mat_name.upper()
                if any(k in name_up for k in ['CARESTREAM', 'AGFA', 'FUJI', 'KODAK', 'STRUCTURIX', 'FILM', '필름']):
                    default_cat = 'FILM'
                elif any(k in name_up for k in ['침투제', '세척제', '현상제', '자분', '페인트', 'NABAKEM', 'MAGNAFLUX']):
                    default_cat = 'NDT_CHEM'
                
                new_material = {
                    'MaterialID': new_mat_id,
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
                # MaterialID is already set from the lookup above
                pass
            
            manager_val = self.cb_daily_user.get().strip() if hasattr(self, 'cb_daily_user') else ''
            # WorkTime and OT fields (keeping current attribute references for worker 1)
            wt1_val = self.ent_worktime1.get().strip() if hasattr(self, 'ent_worktime1') else ''
            ot1_val = self.ent_ot1.get().strip() if hasattr(self, 'ent_ot1') else ''
            
            user_names = [manager_val]
            
            # Access other workers (2-10) dynamically if they exist
            for idx in range(2, 11):
                cb_attr = f'cb_daily_user{idx}'
                if hasattr(self, cb_attr):
                    u_val = getattr(self, cb_attr).get().strip()
                    user_names.append(u_val)
                else:
                    user_names.append('')
            
            # Unpack for the dictionary below
            user2_val = getattr(self, 'cb_daily_user2').get().strip() if hasattr(self, 'cb_daily_user2') else ''
            wt2_val = getattr(self, 'ent_worktime2').get().strip() if hasattr(self, 'ent_worktime2') else ''
            ot2_val = getattr(self, 'ent_ot2').get().strip() if hasattr(self, 'ent_ot2') else ''
            
            user3_val = getattr(self, 'cb_daily_user3').get().strip() if hasattr(self, 'cb_daily_user3') else ''
            wt3_val = getattr(self, 'ent_worktime3').get().strip() if hasattr(self, 'ent_worktime3') else ''
            ot3_val = getattr(self, 'ent_ot3').get().strip() if hasattr(self, 'ent_ot3') else ''
            
            user4_val = getattr(self, 'cb_daily_user4').get().strip() if hasattr(self, 'cb_daily_user4') else ''
            wt4_val = getattr(self, 'ent_worktime4').get().strip() if hasattr(self, 'ent_worktime4') else ''
            ot4_val = getattr(self, 'ent_ot4').get().strip() if hasattr(self, 'ent_ot4') else ''
            
            user5_val = getattr(self, 'cb_daily_user5').get().strip() if hasattr(self, 'cb_daily_user5') else ''
            wt5_val = getattr(self, 'ent_worktime5').get().strip() if hasattr(self, 'ent_worktime5') else ''
            ot5_val = getattr(self, 'ent_ot5').get().strip() if hasattr(self, 'ent_ot5') else ''
            
            user6_val = getattr(self, 'cb_daily_user6').get().strip() if hasattr(self, 'cb_daily_user6') else ''
            wt6_val = getattr(self, 'ent_worktime6').get().strip() if hasattr(self, 'ent_worktime6') else ''
            ot6_val = getattr(self, 'ent_ot6').get().strip() if hasattr(self, 'ent_ot6') else ''
            
            user7_val = getattr(self, 'cb_daily_user7').get().strip() if hasattr(self, 'cb_daily_user7') else ''
            wt7_val = getattr(self, 'ent_worktime7').get().strip() if hasattr(self, 'ent_worktime7') else ''
            ot7_val = getattr(self, 'ent_ot7').get().strip() if hasattr(self, 'ent_ot7') else ''
            
            user8_val = getattr(self, 'cb_daily_user8').get().strip() if hasattr(self, 'cb_daily_user8') else ''
            wt8_val = getattr(self, 'ent_worktime8').get().strip() if hasattr(self, 'ent_worktime8') else ''
            ot8_val = getattr(self, 'ent_ot8').get().strip() if hasattr(self, 'ent_ot8') else ''
            
            user9_val = getattr(self, 'cb_daily_user9').get().strip() if hasattr(self, 'cb_daily_user9') else ''
            wt9_val = getattr(self, 'ent_worktime9').get().strip() if hasattr(self, 'ent_worktime9') else ''
            ot9_val = getattr(self, 'ent_ot9').get().strip() if hasattr(self, 'ent_ot9') else ''
            
            user10_val = getattr(self, 'cb_daily_user10').get().strip() if hasattr(self, 'cb_daily_user10') else ''
            wt10_val = getattr(self, 'ent_worktime10').get().strip() if hasattr(self, 'ent_worktime10') else ''
            ot10_val = getattr(self, 'ent_ot10').get().strip() if hasattr(self, 'ent_ot10') else ''
            
            all_workers = ", ".join([u for u in user_names if u])

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
                'Site': site.strip(),
                'MaterialID': mat_id,
                '장비명': equip_name,
                '검사방법': test_method,
                '검사량': test_amount,
                '단가': unit_price,
                '출장비': travel_cost,
                '일식': meal_cost,
                '검사비': test_fee,
                'FilmCount': film_count,
                'Usage': total_usage,
                'Note': note,
                'EntryTime': datetime.datetime.now(),
                'User': manager_val,
                'WorkTime': wt1_val,
                'OT': ot1_val,
                'User2': user2_val,
                'WorkTime2': wt2_val,
                'OT2': ot2_val,
                'User3': user3_val,
                'WorkTime3': wt3_val,
                'OT3': ot3_val,
                'User4': user4_val,
                'WorkTime4': wt4_val,
                'OT4': ot4_val,
                'User5': user5_val,
                'WorkTime5': wt5_val,
                'OT5': ot5_val,
                'User6': user6_val,
                'WorkTime6': wt6_val,
                'OT6': ot6_val,
                'User7': user7_val,
                'WorkTime7': wt7_val,
                'OT7': ot7_val,
                'User8': user8_val,
                'WorkTime8': wt8_val,
                'OT8': ot8_val,
                'User9': user9_val,
                'WorkTime9': wt9_val,
                'OT9': ot9_val,
                'User10': user10_val,
                'WorkTime10': wt10_val,
                'OT10': ot10_val
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
                    # Remove all whitespace from material name for the key
                    m_key = "".join(material.split())
                    new_entry[f'NDT_{m_key}'] = value
                    if value > 0:
                        ndt_usage[material] = value
                except ValueError:
                    m_key = "".join(material.split())
                    new_entry[f'NDT_{m_key}'] = 0.0
            
            # Final key cleanup for the entire entry to be 100% safe
            # Ensure all keys are strings and values are properly formatted
            final_entry = {}
            for k, v in new_entry.items():
                clean_key = str(k).strip()
                if isinstance(v, (int, float)):
                    final_entry[clean_key] = float(v)
                else:
                    final_entry[clean_key] = str(v) if v is not None else ""
            
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
                mat_row = self.materials_df[self.materials_df['MaterialID'] == mat_id]
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
                new_transaction = {
                    'Date': usage_datetime,
                    'MaterialID': mat_id,
                    'Type': 'OUT',
                    'Quantity': total_deduction,
                    'User': all_workers,
                    'Site': site,
                    'Note': f'{site} 현장 사용 (자동 차감)'
                }
                self.transactions_df = pd.concat([self.transactions_df, pd.DataFrame([new_transaction])], ignore_index=True)
                # Ensure columns stay consistent
                self.transactions_df.columns = [str(c).strip() for c in self.transactions_df.columns]
                
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
                    ndt_mat_id = ndt_mat_row['MaterialID']
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
                            'MaterialID': ndt_mat_id,
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
            
            self.daily_usage_df = pd.concat([self.daily_usage_df, pd.DataFrame([final_entry])], ignore_index=True)
            # Maintain clean headers
            self.daily_usage_df.columns = [str(c).strip() for c in self.daily_usage_df.columns]
            
            # Ensure Date column is consistently formatted
            if 'Date' in self.daily_usage_df.columns:
                self.daily_usage_df['Date'] = pd.to_datetime(self.daily_usage_df['Date'], errors='coerce').dt.strftime('%Y-%m-%d')
            
            # Count transactions effectively added in this call
            # (This is a bit hard with the current structure, but we can check if self.transactions_df length increased)
            # Actually, let's just make sure we save and refresh.
            
            if self.save_data():
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
                for i in range(1, 11):
                    cb_attr = f'cb_daily_user{i}' if i > 1 else 'cb_daily_user'
                    if hasattr(self, cb_attr):
                        u_name = getattr(self, cb_attr).get().strip()
                        # Extract actual name from "(Shift) Name" format
                        import re
                        actual_name = u_name
                        match = re.match(r"\((주간|야간|휴일)\)\s*(.*)", u_name)
                        if match:
                            # Extract only the name part (group 2)
                            actual_name = match.group(2).strip()
                        
                        # Only add if there's a real name (not empty or just shift)
                        if actual_name and actual_name not in self.users:
                            self.users.append(actual_name)
                            self.users.sort()
                            self.refresh_ui_for_list_change('users')
                            self.save_tab_config()
                            stock_info += f"\n• 신규 담당자 '{actual_name}'이 목록에 저장되었습니다."
                
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
                for i in range(1, 11):
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
                
                # Clear WorkTime and OT fields
                for i in range(1, 11):
                    wt_attr = f'ent_worktime{i}' if i > 1 else 'ent_worktime1'
                    ot_attr = f'ent_ot{i}' if i > 1 else 'ent_ot1'
                    if hasattr(self, wt_attr): getattr(self, wt_attr).delete(0, tk.END)
                    if hasattr(self, ot_attr): getattr(self, ot_attr).delete(0, tk.END)
                self.rtk_entries["총계"].config(state='normal')
                self.rtk_entries["총계"].delete(0, tk.END)
                self.rtk_entries["총계"].config(state='readonly')
                # Keep date and reset to today
                self.ent_daily_date.set_date(datetime.datetime.now())
            else:
                # Save returned False - do not clear fields
                pass # The error message is already shown by save_data

        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            self.show_error_dialog("저장 오류", f"기록 저장 중 오류가 발생했습니다:\n{e}\n\n상세 정보:\n{error_details}")
    
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
        filter_worker = self.cb_daily_filter_worker.get() if hasattr(self, 'cb_daily_filter_worker') else '전체'
        
        # Populate site filter options from data
        if hasattr(self, 'cb_daily_filter_site'):
            unique_sites = ['전체'] + sorted(self.daily_usage_df['Site'].dropna().unique().tolist())
            self.cb_daily_filter_site['values'] = unique_sites
            if not self.cb_daily_filter_site.get():
                self.cb_daily_filter_site.set('전체')
        
        # Populate material filter options from data
        if hasattr(self, 'cb_daily_filter_material'):
            # Use unified display names for filter options
            unique_mat_ids = self.daily_usage_df['MaterialID'].dropna().unique()
            material_names = []
            for mat_id in unique_mat_ids:
                material_names.append(self.get_material_display_name(mat_id))
            unique_materials = ['전체'] + sorted(set(material_names))
            self.cb_daily_filter_material['values'] = unique_materials
            if not self.cb_daily_filter_material.get():
                self.cb_daily_filter_material.set('전체')

        # Populate worker filter options from data
        if hasattr(self, 'cb_daily_filter_worker'):
            worker_cols = ['User', 'User2', 'User3', 'User4', 'User5', 'User6', 'User7', 'User8', 'User9', 'User10']
            all_workers = set()
            for col in worker_cols:
                if col in self.daily_usage_df.columns:
                    all_workers.update(self.daily_usage_df[col].dropna().unique().tolist())
            unique_workers = ['전체'] + sorted([w for w in all_workers if w])
            self.cb_daily_filter_worker['values'] = unique_workers
            if not self.cb_daily_filter_worker.get():
                self.cb_daily_filter_worker.set('전체')
        
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
                    if self.get_material_display_name(mat['MaterialID']) == filter_material:
                        matching_mat_ids.append(mat['MaterialID'])
            
            if matching_mat_ids:
                filtered_df = filtered_df[filtered_df['MaterialID'].isin(matching_mat_ids)]
            else:
                # If no direct match in materials_df, try basic 품목명 (for edge cases)
                matching_mat_ids = self.materials_df[self.materials_df['품목명'] == filter_material]['MaterialID'].tolist()
                if matching_mat_ids:
                    filtered_df = filtered_df[filtered_df['MaterialID'].isin(matching_mat_ids)]

        # Filter by worker if specified
        if filter_worker != '전체':
            worker_cols = ['User', 'User2', 'User3', 'User4', 'User5', 'User6', 'User7', 'User8', 'User9', 'User10']
            # Build a mask that is True if the worker is in any of the columns
            mask = filtered_df[worker_cols[0]] == filter_worker
            for col in worker_cols[1:]:
                if col in filtered_df.columns:
                    mask = mask | (filtered_df[col] == filter_worker)
            filtered_df = filtered_df[mask]
        
        # Sort by date descending - ensure Date column is consistent type
        if not filtered_df.empty:
            # Convert Date column to datetime for consistent sorting
            filtered_df['Date'] = pd.to_datetime(filtered_df['Date'], errors='coerce')
            filtered_df = filtered_df.sort_values('Date', ascending=False)
        
        # Define RTK categories
        rtk_categories = ["센터미스", "농도", "마킹미스", "필름마크", "취급부주의", "고객불만", "기타", "총계"]
        
        # Display entries and calculate totals
        # Display entries and calculate totals
        total_film_count = 0
        total_rtk = [0.0] * len(rtk_categories)
        total_ndt = [0.0] * 7
        total_test_amount = 0.0
        total_unit_price = 0.0
        total_travel_cost = 0.0
        total_meal_cost = 0.0
        total_test_fee = 0.0
        total_ot_hours = 0.0
        total_ot_amount = 0
        total_work_hours = 0.0 # Added for Total Working Time
        total_indiv_ot_hours = [0.0] * 10
        total_indiv_ot_amounts = [0] * 10
        
        
        current_date = None
        

        for idx, entry in filtered_df.iterrows():
            usage_date = self._safe_format_datetime(entry.get('Date', ''), '%Y-%m-%d')
            if not usage_date:
                usage_date = "Unknown"
            
            current_date = usage_date
            
            mat_id = entry['MaterialID']
            mat_name = self.get_material_display_name(mat_id)
            
            entry_time = self._safe_format_datetime(entry.get('EntryTime', ''), '%Y-%m-%d %H:%M:%S')
            
            # Consolidate workers
            def clean_str(val):
                return str(val).replace('nan', '').replace('None', '').strip()

            all_users = [
                clean_str(entry.get('User', '')),
                clean_str(entry.get('User2', '')),
                clean_str(entry.get('User3', '')),
                clean_str(entry.get('User4', '')),
                clean_str(entry.get('User5', '')),
                clean_str(entry.get('User6', '')),
                clean_str(entry.get('User7', '')),
                clean_str(entry.get('User8', '')),
                clean_str(entry.get('User9', '')),
                clean_str(entry.get('User10', ''))
            ]
            consolidated_workers = ", ".join([u for u in all_users if u])
            
            # Get film count - consistently use FilmCount
            film_count = entry.get('FilmCount', 0)
            if pd.notna(film_count):
                film_count_val = float(film_count)
                total_film_count += film_count_val
                film_count_str = f"{film_count_val:.1f}"
            else:
                film_count_str = "0.0"
            
            # Accumulate cost totals
            total_test_fee += entry.get('검사비', 0.0)
            
            # Calculate row OT total
            row_ot_hours = 0.0
            row_ot_amount = 0
            row_ot_parts = []
            
            
            # Calculate total work hours for all workers (1-10)
            # Calculate total work hours for all workers (1-10)
            for i in range(1, 11):
                # Check user for this slot
                user_key = 'User' if i == 1 else f'User{i}'
                user_val = clean_str(entry.get(user_key, ''))
                
                # If filtering by worker, skip if this slot is not the worker
                if filter_worker != '전체' and user_val != filter_worker:
                    continue
                    
                wt_key = 'WorkTime' if i == 1 else f'WorkTime{i}'
                wt_str = clean_str(entry.get(wt_key, ''))
                if wt_str and '~' in wt_str:
                    try:
                        start_str, end_str = wt_str.split('~')
                        sh, sm = map(int, start_str.split(':'))
                        eh, em = map(int, end_str.split(':'))
                        
                        start_min = sh * 60 + sm
                        end_min = eh * 60 + em
                        
                        if end_min < start_min: # Cross midnight
                             end_min += 24 * 60
                             
                        duration_hours = (end_min - start_min) / 60.0
                        total_work_hours += duration_hours
                    except:
                        pass

            for i in range(1, 11):
                # Check worker filter for this slot
                user_key = 'User' if i == 1 else f'User{i}'
                user_val = str(entry.get(user_key, '')).replace('nan', '').replace('None', '').strip()
                
                if filter_worker != '전체' and user_val != filter_worker:
                    continue

                ot_str = clean_str(entry.get(f'OT{i}' if i > 1 else 'OT', ''))
                if ot_str:
                    try:
                        # Extract hours and amount from "X.X시간 (X,XXX원)"
                        h_part = ot_str.split('시간')[0]
                        row_ot_hours += float(h_part)
                        if '(' in ot_str and '원)' in ot_str:
                            a_part = ot_str.split('(')[1].split('원')[0].replace(',', '')
                            row_ot_amount += int(a_part)
                    except:
                        pass
            
            total_ot_hours += row_ot_hours
            total_ot_amount += row_ot_amount
            ot_sum_display = f"{row_ot_hours:.1f}시간" if row_ot_hours > 0 else "0"
            if row_ot_amount > 0:
                ot_sum_display += f" ({row_ot_amount:,}원)"
            
            # Individual OT values for columns (Simplified: Amount only)
            row_ots = []
            for i in range(1, 11):
                # Check worker filter for this slot
                user_key = 'User' if i == 1 else f'User{i}'
                user_val = str(entry.get(user_key, '')).replace('nan', '').replace('None', '').strip()
                
                if filter_worker != '전체' and user_val != filter_worker:
                    row_ots.append("") # Empty string looks cleaner than "0" for hidden cols
                    continue

                ot_str = clean_str(entry.get(f'OT{i}' if i > 1 else 'OT', ''))
                if ot_str:
                    try:
                        h_part = ot_str.split('시간')[0]
                        a_part = ot_str.split('(')[1].split('원')[0].replace(',', '')
                        total_indiv_ot_hours[i-1] += float(h_part)
                        total_indiv_ot_amounts[i-1] += int(a_part)
                        row_ots.append(f"{int(a_part):,}")
                    except:
                        row_ots.append(ot_str)
                else:
                    row_ots.append("0")
            
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
                entry.get('WorkTime', ''),
                *row_ots,
                entry.get('장비명', ''),
                entry.get('검사방법', ''),
                f"{entry.get('검사량', 0.0):.1f}",
                film_count_str,
                f"{entry.get('단가', 0.0):,.0f}",
                f"{entry.get('출장비', 0.0):,.0f}",
                f"{entry.get('일식', 0.0):,.0f}",
                f"{entry.get('검사비', 0.0):,.0f}",
                f"{row_ot_hours:.1f}", # OT시간
                f"{row_ot_amount:,}",  # OT금액
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
                f"{total_work_hours:.1f} Hrs", # 작업시간 (Calculated Total)
                # Individual OT Totals (Simplified: Amount only)
                *[f"{a:,}" if a > 0 else "0" for a in total_indiv_ot_amounts],
                '', # 장비명
                '', # 검사방법
                f"{total_test_amount:.1f}",
                f"{total_film_count:.1f}",
                f"{total_unit_price:,.0f}",
                f"{total_travel_cost:,.0f}",
                f"{total_meal_cost:,.0f}",
                f"{total_test_fee:,.0f}",
                f"{total_ot_hours:.1f}", # OT시간 합계
                f"{total_ot_amount:,}",  # OT금액 합계
                '', # 품목명
                *[f"{v:.1f}" for v in total_rtk],
                *[f"{v:.1f}" for v in total_ndt],
                '',   # 비고
                ''    # 입력시간
            ]
            self.daily_usage_tree.tag_configure('total', background='#E8F4F8', font=('Arial', 12, 'bold'))
            self.daily_usage_tree.insert('', tk.END, values=total_values, tags=('total',))
            
            # --- Dynamic Column Hiding ---
            mandatory_cols = ['날짜', '현장', '작업자', '작업시간', 'OT1', 'OT2', 'OT3', 'OT4', 'OT5', 'OT6', 'OT7', 'OT8', 'OT9', 'OT10', '장비명', '검사방법', '품목명', '비고', '입력시간']
            
            # Map dynamic columns to their total values
            # Index positions in total_rtk and total_ndt must match setup_daily_usage_tab's columns
            dynamic_col_status = {
                '검사량': total_test_amount > 0,
                '필름매수': total_film_count > 0,
                '단가': total_unit_price > 0,
                '출장비': total_travel_cost > 0,
                '일식': total_meal_cost > 0,
                '검사비': total_test_fee > 0,
                'OT시간': total_ot_hours > 0,
                'OT금액': total_ot_amount > 0
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
            
            # If user has manually selected columns, use those
            if hasattr(self, 'manual_visible_cols') and self.manual_visible_cols:
                visible_cols = [col for col in self.manual_visible_cols if col in all_cols]
                # CRITICAL: Always ensure mandatory columns are included
                for m_col in mandatory_cols:
                    if m_col in all_cols and m_col not in visible_cols:
                        # Find original index to keep order if possible
                        try:
                            orig_idx = list(all_cols).index(m_col)
                            # Insert at correct relative position
                            inserted = False
                            for i, v_col in enumerate(visible_cols):
                                if list(all_cols).index(v_col) > orig_idx:
                                    visible_cols.insert(i, m_col)
                                    inserted = True
                                    break
                            if not inserted:
                                visible_cols.append(m_col)
                        except:
                            visible_cols.append(m_col)
            else:
                visible_cols = [col for col in all_cols if col in mandatory_cols or dynamic_col_status.get(col, False)]
            
            # Apply visible columns to treeview
            self.daily_usage_tree['displaycolumns'] = visible_cols

            # Ensure stretch=False and minwidth is relaxed for ALL displayed columns
            # This is critical for making all columns, especially the last one, resizable
            for col in visible_cols:
                self.daily_usage_tree.column(col, stretch=False, minwidth=20)

            # Re-apply saved column widths if available
            if hasattr(self, 'tab_config') and 'daily_usage_col_widths' in self.tab_config:
                saved_widths = self.tab_config['daily_usage_col_widths']
                for col in visible_cols:
                    if col in saved_widths:
                        try:
                            self.daily_usage_tree.column(col, width=int(saved_widths[col]), stretch=False)
                            # Enforce minimums for high-precision cols to prevent truncation
                            if col == '날짜':
                                self.daily_usage_tree.column(col, width=max(int(saved_widths[col]), 160), stretch=False)
                            elif col == '입력시간':
                                self.daily_usage_tree.column(col, width=max(int(saved_widths[col]), 300), stretch=False)
                        except: pass
            
            # Ensure Total Row stays at bottom
            self.daily_usage_tree.detach(self.daily_usage_tree.get_children()[-1])
            self.daily_usage_tree.insert('', tk.END, values=total_values, tags=('total',))
        else:
            # If empty, check manual visibility first
            if hasattr(self, 'manual_visible_cols') and self.manual_visible_cols:
                visible_cols = [col for col in self.manual_visible_cols if col in self.daily_usage_tree['columns']]
                self.daily_usage_tree['displaycolumns'] = visible_cols
            else:
                visible_cols = ['날짜', '현장', '작업자', '작업시간', 'OT1', 'OT2', 'OT3', 'OT4', 'OT5', 'OT6', 'OT7', 'OT8', 'OT9', 'OT10', '장비명', '검사방법', '품목명', '비고', '입력시간']
                self.daily_usage_tree['displaycolumns'] = visible_cols
            
            # Ensure stretch=False and minwidth is relaxed for all displayed columns
            for col in visible_cols:
                self.daily_usage_tree.column(col, stretch=False, minwidth=20)

            # Re-apply saved column widths if available
            if hasattr(self, 'tab_config') and 'daily_usage_col_widths' in self.tab_config:
                saved_widths = self.tab_config['daily_usage_col_widths']
                for col in visible_cols:
                    if col in saved_widths:
                        try:
                            self.daily_usage_tree.column(col, width=int(saved_widths[col]), stretch=False)
                            # Enforce minimums for high-precision cols
                            if col == '날짜':
                                self.daily_usage_tree.column(col, width=max(int(saved_widths[col]), 160), stretch=False)
                            elif col == '입력시간':
                                self.daily_usage_tree.column(col, width=max(int(saved_widths[col]), 300), stretch=False)
                        except: pass
    
    def _auto_adjust_tree_columns(self, tree, expand_only=False):
        """Automatically adjust column widths to fit content"""
        import tkinter.font as tkfont
        
        # Use the actual font used in the Treeview (12pt Malgun Gothic)
        # This ensures measurement matches display perfectly.
        content_font = tkfont.Font(family="Malgun Gothic", size=12)
        heading_font = tkfont.Font(family="Malgun Gothic", size=12, weight="bold")
        
        # We only care about visible columns
        visible_cols = tree['displaycolumns']
        if not visible_cols or visible_cols == ('#all'):
             visible_cols = tree['columns']

        for col in visible_cols:
            # Maximum Overkill: Character Count Heuristic
            # Assume 40px per character + 200px buffer
            # This GUARANTEES visibility in any environment
            
            # Measure heading
            w = len(col) * 40 + 200
            
            # Measure content
            col_index = list(tree['columns']).index(col)
            for item in tree.get_children():
                val = tree.item(item, 'values')
                if val and col_index < len(val):
                    text_len = len(str(val[col_index]))
                    content_w = text_len * 40 + 200
                    if content_w > w:
                        w = content_w
            
            # Apply specific constraints
            # Ultra-Safe: No max limit for potentially long text columns
            if col in ['품목명', '작업자', '비고']:
                w = max(w, 300) # Minimum base
                if col == '품목명': w = max(w, 700)
                elif col == '작업자': w = max(w, 300)
                elif col == '비고': w = max(w, 600)
            else:
                # Other columns get a relaxed global max
                min_w = 200
                if col == '날짜': min_w = 400
                elif col == '현장': min_w = 300
                elif col == '장비명': min_w = 250
                elif col == '검사방법': min_w = 200
                
                w = max(min_w, min(w, 2000)) 
            
            # Smart Auto-Expand: Only grow if content requires it
            if expand_only:
                current_w = int(tree.column(col, 'width'))
                if w <= current_w:
                    continue  # Keep user's manual adjustment
            
            # All columns are now user-resizable
            tree.column(col, width=w, minwidth=20, stretch=False, anchor='center')
        
        # Force layout update
        tree.update_idletasks()

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
                mat_id = entry.get('MaterialID')
                
                # 해당 내역과 관련된 트랜잭션 찾기 (날짜, MaterialID, Note가 현장 사용 자동 차감인 것)
                # Note 형식을 add_daily_usage_entry와 동일하게 구성
                note_pattern = f'{site} 현장 사용 (자동 차감)'
                
                # 필터링하여 관련 트랜잭션 삭제
                if not self.transactions_df.empty:
                    trans_mask = (
                        (pd.to_datetime(self.transactions_df['Date']).dt.date == usage_date.date()) &
                        (self.transactions_df['MaterialID'] == mat_id) &
                        (self.transactions_df['Type'] == 'OUT') &
                        (self.transactions_df['Note'] == note_pattern)
                    )
                    
                    # NDT 자재 트랜잭션도 포함될 수 있음. NDT 자재들은 MaterialID가 다를 수 있으나 
                    # Note는 동일하게 "현장 사용 (자동 차감)"으로 들어감.
                    # NDT 자재들의 MaterialID는 entry에 NDT_xxx 필드로 명시되어 있음
                    ndt_materials = ["형광자분", "흑색자분", "백색페인트", "침투제", "세척제", "현상제", "형광"]
                    ndt_mat_ids = []
                    for ndt_name in ndt_materials:
                        if entry.get(f'NDT_{ndt_name}', 0) > 0:
                            ndt_mat_rows = self.materials_df[self.materials_df['품목명'] == ndt_name]
                            if not ndt_mat_rows.empty:
                                ndt_mat_ids.append(ndt_mat_rows['MaterialID'].values[0])
                    
                    if ndt_mat_ids:
                        trans_mask = trans_mask | (
                            (pd.to_datetime(self.transactions_df['Date']).dt.date == usage_date.date()) &
                            (self.transactions_df['MaterialID'].isin(ndt_mat_ids)) &
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
            
            # Export columns currently visible in the Treeview
            display_cols = self.daily_usage_tree['displaycolumns']
            if not display_cols or display_cols == ('#all'):
                selected_cols = list(self.daily_usage_tree['columns'])
            else:
                selected_cols = list(display_cols)
            
            # Map selected column names to their indices in the Treeview values
            all_cols = list(self.daily_usage_tree['columns'])
            col_indices = [all_cols.index(col) for col in selected_cols]
            
            # Reconstruct data
            filtered_data = []
            for item_values in data:
                row = [item_values[i] for i in col_indices if i < len(item_values)]
                filtered_data.append(row)
            
            filename = f"일일사용내역_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=filename, filetypes=[("Excel files", "*.xlsx")])
            
            if save_path:
                df = pd.DataFrame(filtered_data, columns=selected_cols)
                df = self.clean_df_export(df)
                self.save_df_to_excel_autofit(df, save_path, "일일사용내역")
                messagebox.showinfo("완료", f"데이터가 엑셀로 저장되었습니다.\n{save_path}")
        except Exception as e:
            messagebox.showerror("오류", f"엑셀 내보내기 중 오류가 발생했습니다: {e}")

    def show_column_visibility_dialog(self):
        """Open dialog to manually show/hide columns in the history view"""
        all_cols = list(self.daily_usage_tree['columns'])
        # Pre-select columns that are currently visible
        active_cols = self.daily_usage_tree['displaycolumns']
        if not active_cols or active_cols == ('#all'):
             active_cols = all_cols
        
        dialog = ColumnSelectionDialog(self.root, all_cols, title="표시 컬럼 관리")
        # Overwrite vars with current visibility state
        for col, var in dialog.vars.items():
            var.set(col in active_cols)
            
        dialog.wait_window()
        
        if dialog.result is not None:
            self.manual_visible_cols = dialog.result
            self.daily_usage_tree['displaycolumns'] = dialog.result
            # self._auto_adjust_tree_columns(self.daily_usage_tree)
            # Save configuration if possible
            self.save_tab_config()

    def export_all_daily_usage(self):
        """Export all daily usage records to Excel"""
        try:
            if self.daily_usage_df.empty:
                messagebox.showinfo("알림", "기록된 데이터가 없습니다.")
                return
            
            # Export all columns from the tree structure
            selected_cols = list(self.daily_usage_tree['columns'])
            
            filename = f"전체사용기록_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=filename, filetypes=[("Excel files", "*.xlsx")])
            
            if save_path:
                # Ensure we only try to export columns that exist in the dataframe
                valid_cols = [c for c in selected_cols if c in self.daily_usage_df.columns]
                # If dataframe has extra columns not in tree (like raw data), include them? 
                # Better to stick to what the user sees/defines in tree interface + critical data
                # Actually, export_all usually implies backing up the raw dataframe. 
                # But previous logic used tree columns. Let's stick to tree columns if they exist in DF.
                
                # Fallback: if tree columns are not in DF (e.g. calculated columns), we might have issues.
                # The daily_usage_df has raw data. The tree has formatted data.
                # daily_usage_df columns: Date, Site, MaterialID, Usage, Note, etc. + RTK_..., NDT_...
                # Tree columns: 날짜, 현장, 작업자... OT합계...
                
                # If we want "All Daily Usage", we should probably export the RAW dataframe for backup purposes, 
                # OR the formatted view for reporting.
                # The previous code seemed to try to export based on tree columns but mapped from DF?
                # "export_df = self.daily_usage_df[selected_cols].copy()"
                # This implies 'selected_cols' MUST exist in daily_usage_df.
                # But 'daily_usage_tree['columns']' includes 'OT합계', 'RT총계' which are likely NOT in daily_usage_df (calculated).
                
                # Let's check daily_usage_df columns again.
                # It has 'Date', 'Site', 'MaterialID'...
                # It does NOT have '날짜', '현장' (Korean names).
                # So the previous code 'self.daily_usage_df[selected_cols]' would have FAILED if selected_cols came from tree columns!
                
                # WAIT. The user said "remove column selection". 
                # The previous working code (before my changes) probably did:
                # export_df = self.daily_usage_df.copy()
                # Let's revert to a safe "Backup" style export for "All Data".
                
                export_df = self.daily_usage_df.copy()
                
                export_df = self.clean_df_export(export_df)
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

    def auto_save_worktime(self, event, entry, config_key):
        """Helper to auto-save worktime values and support copy functionality"""
        worktime_value = entry.get().strip()
        if not worktime_value:
            return
            
        # Initialize worktimes list if not exists
        if not hasattr(self, 'worktimes'):
            self.worktimes = []
            
        # Check if this value already exists
        if worktime_value not in self.worktimes:
            self.worktimes.append(worktime_value)
            self.worktimes.sort()
            
            # Save to config with error handling - only save when list changes
            try:
                self.save_tab_config()
            except Exception as e:
                print(f"Failed to save worktime config: {e}")
        
        # Calculate and update corresponding OT field
        self.calculate_and_update_ot(worktime_value, entry)
        
        # Add copy functionality - select all text on focus and bind Ctrl+C (only once)
        entry.select_range(0, tk.END)
        if not hasattr(entry, '_copy_bound'):
            entry.bind('<Control-c>', lambda e: self.copy_worktime(entry))
            entry._copy_bound = True
        
        # Store current value for copy functionality
        entry.last_value = worktime_value

    def calculate_and_update_ot(self, worktime_value, worktime_entry):
        """Calculate OT amount based on worktime and update corresponding OT field"""
        try:
            if '~' not in worktime_value:
                return

            start_time_str, end_time_str = worktime_value.split('~')
            start_hour = int(start_time_str.split(':')[0])
            start_min = int(start_time_str.split(':')[1]) if ':' in start_time_str else 0
            end_hour = int(end_time_str.split(':')[0])
            end_min = int(end_time_str.split(':')[1]) if ':' in end_time_str else 0
            
            # Start and end in float hours (e.g., 18:30 -> 18.5)
            start_f = start_hour + start_min / 60.0
            end_f = end_hour + end_min / 60.0
            
            # Handle overnight shifts
            if end_f < start_f:
                end_f += 24
            
            total_duration = end_f - start_f
            if total_duration <= 0:
                return

            # Get date to determine day of week
            current_date = self.ent_daily_date.get_date()
            weekday = current_date.weekday() # 0=Mon, 4=Fri, 5=Sat, 6=Sun
            is_holiday = weekday >= 5
            is_friday = (weekday == 4)

            ot_hours = 0
            amount = 0

            if is_holiday:
                # Holiday: All hours are OT at 7,500 KRW
                ot_hours = total_duration
                amount = ot_hours * 7500
            else:
                # Weekday: OT starts after 18:00
                if end_f <= 18:
                    ot_hours = 0
                    amount = 0
                else:
                    # Effective OT start is max of actual start and 18:00
                    ot_start = max(start_f, 18.0)
                    ot_hours = end_f - ot_start
                    
                    # Split into 3 buckets: [18-22], [22-24], [24+]
                    # 1. Evening (18:00 ~ 22:00) @ 4,000
                    evening_end = min(end_f, 22.0)
                    evening_hours = max(0, evening_end - ot_start)
                    
                    # 2. Night (22:00 ~ 24:00) @ 5,000
                    night_start = max(ot_start, 22.0)
                    night_end = min(end_f, 24.0)
                    night_hours = max(0, night_end - night_start)
                    
                    # 3. Dawn (24:00+)
                    dawn_start = max(ot_start, 24.0)
                    dawn_hours = max(0, end_f - dawn_start)
                    
                    # Rate for Dawn: 7,500 if started on Friday, else 5,000
                    dawn_rate = 7500 if is_friday else 5000
                    
                    amount = (evening_hours * 4000) + (night_hours * 5000) + (dawn_hours * dawn_rate)

            # Update OT field
            ot_entry = self.get_corresponding_ot_field(worktime_entry)
            if ot_entry:
                if ot_hours > 0:
                    # Round hours to 2 decimal places for display if needed
                    display_hours = round(ot_hours, 2)
                    if display_hours == int(display_hours): display_hours = int(display_hours)
                    
                    ot_display = f"{display_hours}시간 ({int(amount):,}원)"
                    ot_entry.delete(0, tk.END)
                    ot_entry.insert(0, ot_display)
                else:
                    ot_entry.delete(0, tk.END)
        except Exception as e:
            print(f"Error calculating OT: {e}")

    def get_corresponding_ot_field(self, worktime_entry):
        """Get corresponding OT field for a worktime entry"""
        # Map worktime entries to OT entries
        worktime_to_ot = {
            'ent_worktime1': 'ent_ot1', 'ent_worktime2': 'ent_ot2', 
            'ent_worktime3': 'ent_ot3', 'ent_worktime4': 'ent_ot4',
            'ent_worktime5': 'ent_ot5', 'ent_worktime6': 'ent_ot6',
            'ent_worktime7': 'ent_ot7', 'ent_worktime8': 'ent_ot8',
            'ent_worktime9': 'ent_ot9', 'ent_worktime10': 'ent_ot10'
        }
            
        # Find worktime entry name by checking object IDs
        worktime_id = id(worktime_entry)
            
        worktime_attrs = [f'ent_worktime{i}' for i in range(1, 11)]
        for attr_name in worktime_attrs:
            if hasattr(self, attr_name):
                attr_value = getattr(self, attr_name)
                if id(attr_value) == worktime_id:
                    ot_attr_name = worktime_to_ot.get(attr_name)
                    if ot_attr_name and hasattr(self, ot_attr_name):
                        return getattr(self, ot_attr_name)
                    break
            
        return None

    def copy_worktime(self, entry):
        """Copy worktime value to clipboard"""
        try:
            worktime_value = entry.get().strip()
            if worktime_value:
                self.root.clipboard_clear()
                self.root.clipboard_append(worktime_value)
                # Show brief feedback
                entry.delete(0, tk.END)
                entry.insert(0, worktime_value)
        except Exception as e:
            pass  # Silently handle clipboard errors

    def auto_save_ot(self, event, entry, config_key):
        """Helper to auto-save OT values and support copy functionality"""
        ot_value = entry.get().strip()
        if not ot_value:
            return
                
        # Initialize ot_times list if not exists
        if not hasattr(self, 'ot_times'):
            self.ot_times = []
            
        # Calculate OT amount for any input
        calculated_amount = self.calculate_ot_amount(ot_value)
        if calculated_amount:
            # Update entry with calculated amount
            entry.delete(0, tk.END)
            display_value = f"{ot_value} ({calculated_amount:,}원)"
            entry.insert(0, display_value)
            ot_value = display_value
            print(f"OT calculated: {ot_value}")  # Debug print
            
        # Check if this value already exists
        if ot_value not in self.ot_times:
            self.ot_times.append(ot_value)
            self.ot_times.sort()
            
        # Save to config with error handling - only save when list changes
        try:
            self.save_tab_config()
        except Exception as e:
            print(f"Failed to save OT config: {e}")
            
        # Add copy functionality - select all text on focus and bind Ctrl+C (only once)
        entry.select_range(0, tk.END)
        if not hasattr(entry, '_copy_bound'):
            entry.bind('<Control-c>', lambda e: self.copy_ot(entry))
            entry._copy_bound = True
            
        # Store current value for copy functionality
        entry.last_value = ot_value

    def calculate_ot_amount(self, ot_value):
        """Calculate OT amount based on time and rates"""
        try:
            print(f"Calculating OT for: {ot_value}")  # Debug print
                
            # Extract hours from different formats
            hours = 0
            start_hour = 18  # Default start time for OT
            end_hour = 22    # Default end time for OT
                
            if ':' in ot_value:
                # Format like "2:30" or "18:00-22:00" or "18:00~22:00"
                separator = '-' if '-' in ot_value else '~'
                if separator in ot_value:
                    # Time range format "18:00-22:00" or "18:00~22:00"
                    start_time, end_time = ot_value.split(separator)
                    start_hour = int(start_time.split(':')[0])
                    end_hour = int(end_time.split(':')[0])
                    
                    # Handle overnight shifts
                    if end_hour < start_hour:
                        end_hour += 24
                    
                    start_min = int(start_time.split(':')[1]) if ':' in start_time else 0
                    end_min = int(end_time.split(':')[1]) if ':' in end_time else 0
                        
                    # Calculate total hours and minutes
                    total_minutes = (end_hour * 60 + end_min) - (start_hour * 60 + start_min)
                    hours = total_minutes / 60
                    print(f"Time range: {start_hour}:{start_min} to {end_hour}:{end_min} = {hours} hours")
                else:
                    # Simple time format "2:30" (2 hours 30 minutes)
                    time_parts = ot_value.split(':')
                    hours = int(time_parts[0]) + int(time_parts[1]) / 60 if len(time_parts) > 1 else int(time_parts[0])
                    print(f"Simple time: {hours} hours")
            elif '시간' in ot_value:
                # Format like "2시간" or "2.5시간"
                hours = float(ot_value.replace('시간', '').strip())
                print(f"Hours format: {hours} hours")
            else:
                # Simple number
                hours = float(ot_value)
                print(f"Number format: {hours} hours")
                
            # Calculate amount based on time ranges
            amount = 0
                
            if '-' in ot_value:
                # Specific time range calculation
                current_time = start_hour
                remaining_hours = hours
                    
                while remaining_hours > 0 and current_time < 24:
                    if current_time >= 18 and current_time < 22:
                        # 18:00-22:00 rate: 4,000원/hour
                        rate = 4000
                    elif current_time >= 22:
                        # 22:00+ rate: 5,000원/hour
                        rate = 5000
                    else:
                        # Daytime OT (if any): 3,000원/hour
                        rate = 3000
                        
                    # Calculate hours in this rate period
                    next_rate_change = 22 if current_time < 22 else end_hour
                    hours_in_period = min(remaining_hours, next_rate_change - current_time)
                        
                    amount += hours_in_period * rate
                    remaining_hours -= hours_in_period
                    current_time = next_rate_change
                    print(f"Period {current_time}: {hours_in_period}h × {rate} = {amount}")
            else:
                # Simple duration - assume evening OT (18:00-22:00 range primarily)
                # If hours exceed 4 hours, assume some are after 22:00
                evening_hours = min(hours, 4)  # 18:00-22:00 max 4 hours
                night_hours = max(0, hours - 4)  # 22:00+ hours
                    
                amount = evening_hours * 4000 + night_hours * 5000
                print(f"Simple duration: {evening_hours}h × 4000 + {night_hours}h × 5000 = {amount}")
                
            print(f"Final amount: {amount}")
            return int(amount)
                
        except Exception as e:
            print(f"Error calculating OT amount: {e}")
            import traceback
            traceback.print_exc()
            return None

    def copy_ot(self, entry):
        """Copy OT value to clipboard"""
        try:
            ot_value = entry.get().strip()
            if ot_value:
                self.root.clipboard_clear()
                self.root.clipboard_append(ot_value)
                # Show brief feedback
                entry.delete(0, tk.END)
                entry.insert(0, ot_value)
        except Exception as e:
            pass  # Silently handle clipboard errors


    def save_tab_config(self):
        """Save current tab configuration (order and selected tab)"""
        if not getattr(self, 'is_ready', False):
            return  # Prevent saving during startup/loading
            
        try:
            # Force update to ensure all coordinates and sizes are accurate
            self.root.update_idletasks()
            
            # Initialize or keep existing tab_config
            if not hasattr(self, 'tab_config'):
                self.tab_config = {}
                
            # Update with current core values
            self.tab_config.update({
                'selected_tab': self.notebook.index(self.notebook.select()),
                'tab_order': [],
                'sites': self.sites,
                'users': getattr(self, 'users', []),
                'warehouses': getattr(self, 'warehouses', []),
                'equipments': getattr(self, 'equipments', []),
                'worktimes': getattr(self, 'worktimes', []),
                'ot_times': getattr(self, 'ot_times', []),
                'layout_locked': getattr(self, 'layout_locked', False),
                'resolution_locked': getattr(self, 'resolution_locked', False),
                'locked_width': getattr(self, 'locked_width', 1200),
                'locked_height': getattr(self, 'locked_height', 800),
                'daily_usage_sash_locked': getattr(self, 'daily_usage_sash_locked', False),
                'daily_usage_sash_pos': self.daily_usage_paned.sashpos(0) if hasattr(self, 'daily_usage_paned') else None,
                'daily_history_sash_pos': self.daily_history_paned.sashpos(0) if hasattr(self, 'daily_history_paned') else None,
                'entry_inner_frame_height': self.entry_inner_frame.winfo_height() if hasattr(self, 'entry_inner_frame') else None,
                'history_visible_cols': getattr(self, 'manual_visible_cols', []),
                'window_state': self.root.state(),
                'window_width': self.root.winfo_width(),
                'window_height': self.root.winfo_height()
            })
            
            # Save current stock column widths
            self.tab_config['stock_col_widths'] = {}
            if hasattr(self, 'stock_tree'):
                for col in self.stock_tree['columns']:
                    self.tab_config['stock_col_widths'][col] = self.stock_tree.column(col, 'width')

            # Save daily usage column widths
            self.tab_config['daily_usage_col_widths'] = {}
            if hasattr(self, 'daily_usage_tree'):
                for col in self.daily_usage_tree['columns']:
                    self.tab_config['daily_usage_col_widths'][col] = self.daily_usage_tree.column(col, 'width')
            
            # Save tab order
            self.tab_config['tab_order'] = []
            for tab in self.notebook.tabs():
                self.tab_config['tab_order'].append(self.notebook.tab(tab, "text"))
            
            # Save Draggable items
            if 'draggable_geometries' not in self.tab_config:
                self.tab_config['draggable_geometries'] = {}
            
            for key, widget in self.draggable_items.items():
                if widget.winfo_manager() == 'place':
                    self.tab_config['draggable_geometries'][key] = {
                        'x': widget.winfo_x(),
                        'y': widget.winfo_y(),
                        'width': widget.winfo_width(),
                        'height': widget.winfo_height(),
                        'hidden': False
                    }
                    
                    if hasattr(widget, '_label_widget'):
                        self.tab_config['draggable_geometries'][key]['custom_label'] = widget._label_widget.cget('text')
                    
                    if hasattr(widget, '_manage_list_key') and widget._manage_list_key:
                        self.tab_config['draggable_geometries'][key]['manage_list_key'] = widget._manage_list_key

                    if key.startswith('clone_'):
                        self.tab_config['draggable_geometries'][key]['is_clone'] = True
                        self.tab_config['draggable_geometries'][key]['widget_class_name'] = widget._widget_class.__name__
                        saved_kwargs = widget._widget_kwargs.copy()
                        if 'values' in saved_kwargs: del saved_kwargs['values']
                        self.tab_config['draggable_geometries'][key]['widget_kwargs'] = saved_kwargs

                    if key.startswith('memo_'):
                        self.tab_config['draggable_geometries'][key]['text'] = self.memos[key]['text_widget'].get('1.0', 'end-1c')
                        self.tab_config['draggable_geometries'][key]['memo_title'] = self.memos[key]['title_entry'].get()
                    
                    if key in self.checklists:
                        self.tab_config['draggable_geometries'][key]['checklist_title'] = self.checklists[key]['title_entry'].get()
                        items_data = []
                        for child in self.checklists[key]['item_frame'].winfo_children():
                            if hasattr(child, '_checklist_data'):
                                items_data.append({
                                    'text': child._checklist_data['entry'].get(),
                                    'checked': child._checklist_data['var'].get()
                                })
                        self.tab_config['draggable_geometries'][key]['checklist_items'] = items_data
                elif hasattr(widget, 'winfo_manager') and widget.winfo_manager() == '': # Hidden
                    if key in self.tab_config['draggable_geometries']:
                        self.tab_config['draggable_geometries'][key]['hidden'] = True

            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(self.tab_config, f, ensure_ascii=False, indent=2)
            
            print(f"Configuration saved to {self.config_path}")
        except Exception as e:
            print(f"Failed to save tab config: {e}")

    
    def load_tab_config(self):
        """Load and restore tab configuration"""
        try:
            if not hasattr(self, 'tab_config'):
                self.tab_config = {}
                
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    self.tab_config = json.load(f)
                
                config = self.tab_config
                
                # Apply high-level lock states immediately to internal variables
                self.layout_locked = config.get('layout_locked', False)
                self.resolution_locked = config.get('resolution_locked', False)
                self.daily_usage_sash_locked = config.get('daily_usage_sash_locked', False)
                self.locked_width = config.get('locked_width', 1200)
                self.locked_height = config.get('locked_height', 800)

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
                    try:
                        self.notebook.select(selected)
                    except: pass
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
                self.worktimes = config.get('worktimes', [])
                self.ot_times = config.get('ot_times', [])

                # [CLEANUP] Clean up users list from shift markers
                import re
                cleaned_users = []
                for user in self.users:
                    if re.match(r"^\((주간|야간|휴일)\)$", user.strip()):
                        continue
                    match = re.match(r"\((주간|야간|휴일)\)\s*(.*)", user.strip())
                    if match:
                        actual_name = match.group(2).strip()
                        if actual_name and actual_name not in cleaned_users:
                            cleaned_users.append(actual_name)
                    elif user.strip() and user.strip() not in cleaned_users:
                        cleaned_users.append(user.strip())
                
                if len(cleaned_users) != len(self.users):
                    self.users = cleaned_users
                    # Don't save yet, let the normal flow handle it
                
                # Restore history visibility if saved
                self.manual_visible_cols = config.get('history_visible_cols', [])
                # [MIGRATION] If 'OT합계' is in saved cols, replace with 'OT시간', 'OT금액'
                if 'OT합계' in self.manual_visible_cols:
                    index = self.manual_visible_cols.index('OT합계')
                    self.manual_visible_cols.pop(index)
                    if 'OT시간' not in self.manual_visible_cols:
                        self.manual_visible_cols.insert(index, 'OT시간')
                    if 'OT금액' not in self.manual_visible_cols:
                        self.manual_visible_cols.insert(index + 1, 'OT금액')
                
                # [MIGRATION] Translate English names to Korean
                translations = {
                    'Date': '날짜', 'Site': '현장', 'User': '작업자', 'WorkTime': '작업시간',
                    'Equipment': '장비명', 'Method': '검사방법', 'Remark': '비고', 
                    'Note': '비고', 'EntryTime': '입력시간', 'MaterialName': '품목명'
                }
                for i, col in enumerate(self.manual_visible_cols):
                    if col in translations:
                        self.manual_visible_cols[i] = translations[col]
                
                if self.manual_visible_cols and hasattr(self, 'daily_usage_tree'):
                    try:
                        self.daily_usage_tree['displaycolumns'] = [c for c in self.manual_visible_cols if c in self.daily_usage_tree['columns']]
                    except: pass
                
                # Restore warehouses list
                self.warehouses = config.get('warehouses', [])
                if not self.warehouses and not self.materials_df.empty:
                    self.warehouses = sorted(self.materials_df['창고'].dropna().unique().tolist())
                    self.warehouses = [str(w).strip() for w in self.warehouses if str(w).strip()]
                
                # Equipment list
                self.equipments = config.get('equipments', [])
                if not self.equipments and not self.daily_usage_df.empty and '장비명' in self.daily_usage_df.columns:
                    self.equipments = sorted(self.daily_usage_df['장비명'].dropna().unique().tolist())
                    self.equipments = [str(e).strip() for e in self.equipments if str(e).strip()]
                
                # Restore stock column widths
                stock_col_widths = config.get('stock_col_widths', {})
                if stock_col_widths and hasattr(self, 'stock_tree'):
                    for col, width in stock_col_widths.items():
                        try:
                            self.stock_tree.column(col, width=int(width))
                        except:
                            pass
                
                # Restore daily usage column widths
                daily_usage_col_widths = config.get('daily_usage_col_widths', {})
                if daily_usage_col_widths and hasattr(self, 'daily_usage_tree'):
                    for col, width in daily_usage_col_widths.items():
                        try:
                            # Only apply if width is reasonable (e.g. > 10)
                            w = int(width)
                            
                            # All columns are now user-resizable
                            if w > 10:
                                # Enforce minimums for high-precision cols
                                if col == '날짜': w = max(w, 160)
                                elif col == '입력시간': w = max(w, 300)
                                self.daily_usage_tree.column(col, width=w, minwidth=20, stretch=False)
                        except:
                            pass
                if hasattr(self, '_loading_memos'):
                    del self._loading_memos
                    
                # Restore layout lock state
                if hasattr(self, 'btn_lock_layout'):
                    if self.layout_locked:
                        self.btn_lock_layout.config(text="🔒 배치 고정됨")
                        self.style.configure("Lock.TButton", foreground="black")
                    else:
                        self.btn_lock_layout.config(text="🔓 배치 수정 중")
                        self.style.configure("Lock.TButton", foreground="red")

                if config.get('daily_usage_sash_locked', False):
                    self.daily_usage_sash_locked = True
                    
                    if hasattr(self, 'btn_sash_lock'):
                        self.btn_sash_lock.config(text="🔒 경계 잠금됨")
                    
                    print("LOADED: Restored sash lock state")

                # Tab selection handled already at line 4531
                
                # Recreate Memos and Clones first (these must exist before they can be placed)
                self._loading_memos = []
                # Map class names to actual classes for recreation
                class_map = {'Entry': ttk.Entry, 'Combobox': ttk.Combobox}
                
                draggable_geos = config.get('draggable_geometries', {})
                for key, geo in draggable_geos.items():
                    if key.startswith('memo_'):
                        self._loading_memos.append(key)
                        self.add_new_memo(
                            initial_text=geo.get('text', ""), 
                            initial_title=geo.get('memo_title', "메모"), 
                            key=key
                        )
                    elif key.startswith('checklist_'):
                        self._loading_memos.append(key)
                        self.add_new_checklist(
                            initial_data=geo.get('checklist_items', []),
                            initial_title=geo.get('checklist_title', "체크리스트"),
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

                # 2. DELAYED RESTORATION: Restore complex states after UI mapped
                def delayed_restore():
                    self.root.update_idletasks()
                    
                    # 1. Restore window state and size
                    try:
                        if config.get('window_state'):
                            if config['window_state'] == 'zoomed':
                                self.root.state('zoomed')
                            else:
                                try:
                                    self.root.state(config['window_state'])
                                except: pass
                        
                        w = config.get('window_width')
                        h = config.get('window_height')
                        if w and h:
                            # Only set geometry if not zoomed
                            if self.root.state() != 'zoomed':
                                self.root.geometry(f"{w}x{h}")
                    except: pass

                    # 2. Restore resolution lock if active
                    if self.resolution_locked:
                        self.root.resizable(False, False)
                        if hasattr(self, 'btn_resolution_lock'):
                            self.btn_resolution_lock.config(text="🔒 해상도 고정됨")

                    # 3. Restore draggable components
                    draggable_geos = config.get('draggable_geometries', {})
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
                                    
                                if geo.get('custom_label') and hasattr(widget, '_label_widget'):
                                    widget._label_widget.config(text=geo['custom_label'])
                                    
                                self._ensure_placeholder(widget, width=geo['width'], height=geo['height'])
                                widget.lift()
                                widget.place(x=geo['x'], y=geo['y'], width=geo['width'], height=geo['height'])
                            except: pass

                    # 4. Restore sash positions
                    def apply_sashes():
                        try:
                            # Daily usage sash
                            daily_sash = config.get('daily_usage_sash_pos')
                            if daily_sash is not None and hasattr(self, 'daily_usage_paned'):
                                self.daily_usage_paned.sashpos(0, int(daily_sash))
                            
                            # History sash
                            history_sash = config.get('daily_history_sash_pos')
                            if history_sash is not None and hasattr(self, 'daily_history_paned'):
                                self.daily_history_paned.sashpos(0, int(history_sash))
                                
                            # If sash lock is active
                            if self.daily_usage_sash_locked:
                                if hasattr(self, 'btn_sash_lock'):
                                    self.btn_sash_lock.config(text="🔒 경계 잠금됨")
                                self._start_sash_monitor()
                        except: pass

                    self.root.after(100, apply_sashes)
                    self.root.after(800, apply_sashes)

                    # 5. Restore entry_inner_frame height
                    inner_h = config.get('entry_inner_frame_height')
                    if inner_h and hasattr(self, 'entry_inner_frame') and int(inner_h) > 100:
                        self.entry_inner_frame.config(height=int(inner_h))
                    
                    # 6. Final UI refresh
                    for l_key in ['users', 'sites', 'equipments', 'warehouses']:
                        self.refresh_ui_for_list_change(l_key)
                        
                    # 7. Mark as ready for future saves
                    self.is_ready = True
                    print("APP READY: State restoration complete.")
                
                # Execute delayed restoration
                self.root.after(300, delayed_restore)

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


    def treeview_sort_column(self, tv, col, reverse):
        """Sort treeview contents when a column header is clicked"""
        l = [(tv.set(k, col), k) for k in tv.get_children('')]
        
        # Separate the 'Total' row (tagged 'total') from sorting
        data_rows = []
        total_row = None
        
        for val, k in l:
            # Check tags for this item
            tags = tv.item(k, 'tags')
            # If tags is a tuple/list, check if 'total' is in it. 
            # If it's a string (though usually tuple), check equality or contains.
            if tags and ('total' in tags or tags == 'total'):
                total_row = (val, k)
            else:
                data_rows.append((val, k))
                
        # Helper for numeric conversion
        def convert(val):
            try:
                # Remove common non-numeric chars
                s = str(val).replace(',', '').replace('시간', '').replace('Hrs', '').replace('원', '').replace(' ', '').replace('(', '').replace(')', '')
                if not s: return 0.0
                return float(s)
            except ValueError:
                return str(val).lower() # Default to string sort

        try:
            data_rows.sort(key=lambda t: convert(t[0]), reverse=reverse)
        except Exception:
            data_rows.sort(key=lambda t: t[0].lower(), reverse=reverse)

        # Rearrange items in sorted positions
        for index, (val, k) in enumerate(data_rows):
            tv.move(k, '', index)
            
        # Ensure Total row is always at the bottom
        if total_row:
             tv.move(total_row[1], '', 'end')

        # Reverse sort next time
        tv.heading(col, command=lambda: self.treeview_sort_column(tv, col, not reverse))

if __name__ == "__main__":
    root = tk.Tk()
    app = MaterialManager(root)
    root.mainloop()
