import tkinter as tk
from tkinter import ttk, simpledialog
import pandas as pd
from tkcalendar import DateEntry
import traceback
import json
from site_apps.central.src.utils.helpers import normalize_id, enable_column_resize

class DraggableMessagebox:
    """A custom draggable replacements for standard tkinter.messagebox"""
    @staticmethod
    def _show(type, title, message):
        root = tk._default_root
        if not root:
            # Fallback to standard if no root exists yet
            if type == "error": return messagebox.showerror_orig(title, message)
            if type == "warning": return messagebox.showwarning_orig(title, message)
            return messagebox.showinfo_orig(title, message)

        dialog = tk.Toplevel(root)
        dialog.overrideredirect(True) # Remove standard title bar for better drag control
        dialog.attributes("-topmost", True)
        
        # Style the custom dialog
        dialog.config(background="#f3f4f6", highlightthickness=1, highlightbackground="#d1d5db")

        # Custom Title Bar
        title_bar = tk.Frame(dialog, background="#ffffff", height=30, cursor="fleur")
        title_bar.pack(side="top", fill="x")
        
        title_lbl = tk.Label(title_bar, text=title, font=("Malgun Gothic", 9, "bold"), background="#ffffff", padx=10)
        title_lbl.pack(side="left")
        
        def close_dialog():
            dialog.grab_release()
            dialog.destroy()

        btn_close = tk.Label(title_bar, text="✕", font=("Malgun Gothic", 10), background="#ffffff", padx=10, cursor="hand2")
        btn_close.pack(side="right")
        btn_close.bind("<Button-1>", lambda e: close_dialog())
        btn_close.bind("<Enter>", lambda e: btn_close.config(background="#ef4444", foreground="white"))
        btn_close.bind("<Leave>", lambda e: btn_close.config(background="#ffffff", foreground="black"))

        # Disable main window
        dialog.transient(root)
        dialog.grab_set()

        # Dragging logic
        def start_drag(event):
            dialog._drag_start_x = event.x
            dialog._drag_start_y = event.y

        def do_drag(event):
            x = dialog.winfo_x() + event.x - dialog._drag_start_x
            y = dialog.winfo_y() + event.y - dialog._drag_start_y
            dialog.geometry(f"+{x}+{y}")

        title_bar.bind("<Button-1>", start_drag)
        title_bar.bind("<B1-Motion>", do_drag)
        title_lbl.bind("<Button-1>", start_drag)
        title_lbl.bind("<B1-Motion>", do_drag)

        # Content
        main_frame = tk.Frame(dialog, background="#f3f4f6", padx=20, pady=20)
        main_frame.pack(expand=True, fill='both')

        icon_char = "ℹ" if type == "info" else "⚠" if type == "warning" else "❌"
        icon_color = "#0078d7" if type == "info" else "#f59e0b" if type == "warning" else "#ef4444"
        
        lbl_icon = tk.Label(main_frame, text=icon_char, font=("Malgun Gothic", 24), fg=icon_color, background="#f3f4f6")
        lbl_icon.pack(side="left", anchor="n", padx=(0, 15))

        # Use a wider wraplength for longer messages
        lbl_msg = tk.Label(main_frame, text=message, font=("Malgun Gothic", 10), justify="left", wraplength=480, background="#f3f4f6")
        lbl_msg.pack(side="left", fill="both", expand=True)

        # Auto-size: compute proper width/height after all widgets are built
        def _auto_position():
            dialog.update_idletasks()
            req_w = max(440, dialog.winfo_reqwidth() + 20)
            req_h = max(200, dialog.winfo_reqheight() + 20)
            x = root.winfo_x() + (root.winfo_width() // 2) - (req_w // 2)
            y = root.winfo_y() + (root.winfo_height() // 2) - (req_h // 2)
            dialog.geometry(f"{req_w}x{req_h}+{x}+{y}")
        dialog.after(10, _auto_position)

        btn_frame = tk.Frame(dialog, background="#f3f4f6", pady=10)
        btn_frame.pack(side="bottom", fill='x')
        
        # Standard tk.Button for crisp rectangular (사각형) look
        btn_ok = tk.Button(btn_frame, text="확인", font=("Malgun Gothic", 10, "bold"), 
                           command=close_dialog, background="#ffffff", activebackground="#e5e7eb",
                           relief="raised", borderwidth=1, padx=30, pady=5, width=10)
        btn_ok.pack(pady=10)
        btn_ok.focus_set()
        
        dialog.bind("<Return>", lambda e: close_dialog())
        dialog.bind("<Escape>", lambda e: close_dialog())

        dialog.lift()
        dialog.focus_force()
        root.wait_window(dialog)

    @staticmethod
    def showerror(title, message): DraggableMessagebox._show("error", title, message)
    @staticmethod
    def showwarning(title, message): DraggableMessagebox._show("warning", title, message)
    @staticmethod
    def showinfo(title, message): DraggableMessagebox._show("info", title, message)


class WorkerCompositeWidget(ttk.Frame):
    """
    Composite widget for Worker selection: [Name] with Autocomplete
    """
    def __init__(self, parent, enable_autocomplete=False, user_list=None, **kwargs):
        super().__init__(parent)
        
        # Worker Name selection
        name_width = kwargs.pop('width', 15)
        self.cb_name = ttk.Combobox(self, width=name_width, **kwargs)
        self.cb_name.pack(side='left', fill='x', expand=True)
        
    def get(self):
        """Return clean name"""
        return self.cb_name.get().strip()

    def set(self, value):
        """Set name, cleaning off any (Shift) prefixes if present"""
        if not value:
            self.cb_name.set("")
            return
            
        import re
        # Progressively migrate: if data still has (Shift) prefix, strip it for the name field
        match = re.match(r"\((주간|야간|휴일|주야간)\)\s*(.*)", str(value))
        if match:
            self.cb_name.set(match.group(2).strip())
        else:
            self.cb_name.set(str(value).strip())

    def bind(self, sequence=None, func=None, add=None):
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
    Unified widget for a worker's record: [Name] [Shift] [WorkTime] [OT]
    """
    def __init__(self, parent, worker_index, users_list, time_list=None, enable_autocomplete=False, **kwargs):
        super().__init__(parent, padding=2) # Reduced padding for compact layout
        self.worker_index = worker_index
        
        # 1. Name selection (WorkerCompositeWidget now handles only name)
        mod_users = [''] + list(users_list) if users_list else ['']
        self.composite = WorkerCompositeWidget(
            self, width=10, values=mod_users, 
            enable_autocomplete=enable_autocomplete, 
            user_list=mod_users
        )
        self.composite.pack(side='left', padx=(0, 2), fill='x', expand=True)
        self.cb_name = self.composite.cb_name
        
        # 2. Shift selection (Moved here from WorkerCompositeWidget)
        self.cb_shift = ttk.Combobox(self, values=["주간", "야간", "휴일", "주야간"], width=5, state="readonly")
        self.cb_shift.pack(side='left', padx=(1, 2))
        self.cb_shift.set("") # Default empty
        
        # 3. Work Time (Changed to Combobox for mouse selection)
        ttk.Label(self, text="시간:").pack(side='left', padx=(1, 0))
        mod_time = [''] + list(time_list) if time_list else ['']
        self.ent_worktime = ttk.Combobox(self, width=12, values=mod_time)
        self.ent_worktime.pack(side='left', padx=(0, 2), fill='x', expand=True)
        self.ent_worktime.set("") # Default empty
        
        # 4. OT
        ttk.Label(self, text="OT:").pack(side='left', padx=(1, 0))
        self.ent_ot = ttk.Entry(self, width=10)
        self.ent_ot.pack(side='left', fill='x', expand=True)
        
        # 5. 일비 (Meal/Per Diem)
        ttk.Label(self, text="일비:").pack(side='left', padx=(5, 0))
        self.ent_meal = ttk.Entry(self, width=10)
        self.ent_meal.pack(side='left', fill='x', expand=True)


    def get_worker(self): return self.composite.get()
    def set_worker(self, val): self.composite.set(val)
    
    def get_time(self): 
        """Return combined string: (Shift) Time"""
        shift = self.cb_shift.get()
        time = self.ent_worktime.get().strip()
        if not time:
            return ""
        return f"({shift}) {time}"
        
    def set_time(self, val):
        """Parse string '(Shift) Time' and set widgets"""
        if not val:
            self.ent_worktime.set("")
            self.cb_shift.set("주간")
            return
            
        import re
        match = re.match(r"\((주간|야간|휴일|주야간)\)\s*(.*)", str(val))
        if match:
            self.cb_shift.set(match.group(1))
            self.ent_worktime.set(match.group(2).strip())
        else:
            # Fallback for old format (just time)
            self.cb_shift.set("주간")
            self.ent_worktime.set(str(val).strip())

    def get_ot(self): return self.ent_ot.get()
    def set_ot(self, val):
        self.ent_ot.delete(0, tk.END)
        self.ent_ot.insert(0, val)

    def get_meal(self): return self.ent_meal.get()
    def set_meal(self, val):
        self.ent_meal.delete(0, tk.END)
        self.ent_meal.insert(0, val)

    def bind_name(self, seq, func): self.cb_name.bind(seq, func)
    def bind_time(self, seq, func): 
        self.ent_worktime.bind(seq, func)
        # Shift selection should also trigger any auto-save bindings
        self.cb_shift.bind('<<ComboboxSelected>>', func, add='+')
        if 'FocusOut' in seq or 'Return' in seq:
            # Also trigger on selection from dropdown
            self.ent_worktime.bind('<<ComboboxSelected>>', func, add='+')
            
    def bind_ot(self, seq, func): self.ent_ot.bind(seq, func)

    def update_time_list(self, new_list):
        """Refresh the combobox values with a new list"""
        if hasattr(self, 'ent_worktime'):
            self.ent_worktime['values'] = [''] + list(new_list) if new_list else ['']


class VehicleInspectionWidget(ttk.Frame):
    """
    Dedicated widget for vehicle inspection records with scrollable content
    """
    def __init__(self, parent, theme_bg='#f0f0f0', vehicle_list=None, **kwargs):
        super().__init__(parent, padding=0)
        
        # Create Canvas and Scrollbar for internal scrolling
        self.canvas = tk.Canvas(self, highlightthickness=0, bg=theme_bg)
        self.scrollbar_y = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollbar_x = ttk.Scrollbar(self, orient="horizontal", command=self.canvas.xview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar_y.set, xscrollcommand=self.scrollbar_x.set)

        self.scrollbar_x.pack(side="bottom", fill="x")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar_y.pack(side="right", fill="y")
        
        # Scroll binding is handled globally in MaterialManager.__init__

        def _on_canvas_configure(event):
            # Allow inner frame to be at least its required width, enabling horizontal scroll if needed
            target_w = max(event.width, self.scrollable_frame.winfo_reqwidth())
            self.canvas.itemconfig(self.canvas_window, width=target_w)
        self.canvas.bind("<Configure>", _on_canvas_configure)

        # 1. Inspection Items (2x4 Table based on user request)
        chk_frame = ttk.LabelFrame(self.scrollable_frame, text="차량관리 및 안전관리 점검")
        chk_frame.pack(fill='x', padx=2, pady=2)
        
        # Headers (Table Column Labels) - Shortened to prevent cutoff
        headers = ["구분", "외부상태", "내부청결", "내부청소", "이동함 시건"]
        for i, h in enumerate(headers):
            colspan = 1 if i == 0 else 2
            col_idx = 0 if i == 0 else 1 + (i-1)*2
            tk.Label(chk_frame, text=h, font=('Malgun Gothic', 8, 'bold'), background=theme_bg).grid(row=0, column=col_idx, columnspan=colspan, padx=1, pady=1)

        self.vars = {}
        # Define Categories and Options
        rows = [("출차시", "out"), ("입차시", "in")]
        categories = [
            ("exterior", ["양호", "불량"]),
            ("cleanliness", ["양호", "불량"]),
            ("cleaning", ["함", "안함"]), # [RESTORED] Changed back to 함/안함 as requested
            ("locking", ["잠금", "안함"])   # [FIX] Changed to 잠금 to match template
        ]

        for r_idx, (r_label, r_key) in enumerate(rows, 1):
            tk.Label(chk_frame, text=r_label, font=('Malgun Gothic', 8), background=theme_bg).grid(row=r_idx, column=0, padx=2, pady=1)
            for c_idx, (c_key, options) in enumerate(categories):
                var = tk.StringVar(value="")
                self.vars[f"{r_key}_{c_key}"] = var
                
                start_col = 1 + (c_idx * 2)
                for o_idx, opt in enumerate(options):
                    # Custom Checkbutton logic to simulate radio behavior (one per category/row)
                    cb = ttk.Checkbutton(chk_frame, text=opt, variable=var, 
                                       onvalue=opt, offvalue="", 
                                       command=lambda v=var, o=opt: self._ensure_exclusive(v, o))
                    cb.grid(row=r_idx, column=start_col + o_idx, padx=1, pady=1, sticky='w')

        # 2. Input Fields
        # 2. Input Fields
        input_frame = ttk.Frame(self.scrollable_frame)
        input_frame.pack(fill='x', padx=2, pady=0)
        
        ttk.Label(input_frame, text="차량정보:").grid(row=0, column=0, padx=2, pady=1, sticky='e')
        self.cb_vehicle_info = ttk.Combobox(input_frame, width=15)
        self.cb_vehicle_info.grid(row=0, column=1, padx=2, pady=1, sticky='w')
        if vehicle_list is not None:
            self.update_vehicle_list(vehicle_list)
            
        ttk.Label(input_frame, text="주행거리 (km):").grid(row=1, column=0, padx=2, pady=1, sticky='e')
        self.ent_mileage = ttk.Entry(input_frame, width=15)
        self.ent_mileage.grid(row=1, column=1, padx=2, pady=1, sticky='w')
        
        ttk.Label(input_frame, text="비고:").grid(row=2, column=0, padx=2, pady=1, sticky='e')
        self.ent_remarks = ttk.Entry(input_frame)
        self.ent_remarks.grid(row=2, column=1, padx=2, pady=1, sticky='ew')
        input_frame.grid_columnconfigure(1, weight=1)

        # Focus transitions
        def move_to_mileage(e): self.ent_mileage.focus_set()
        self.cb_vehicle_info.bind('<<ComboboxSelected>>', move_to_mileage)
        self.cb_vehicle_info.bind('<Return>', move_to_mileage)
        
        def on_mileage_return(e):
            self.format_mileage()
            self.ent_remarks.focus_set()
            
        self.ent_mileage.bind('<Return>', on_mileage_return)
        self.ent_mileage.bind('<FocusOut>', lambda e: self.format_mileage())

        # [NEW] Add Standalone Save Button
        self.on_save_callback = kwargs.get('on_save')
        if self.on_save_callback:
            btn_save = ttk.Button(self.scrollable_frame, text="🚗 차량 정보만 개별 저장", 
                                 command=self.trigger_save, style='Accent.TButton' if 'Accent.TButton' in getattr(parent, 'style_names', []) else 'TButton')
            btn_save.pack(fill='x', padx=5, pady=5)
            
        # 3. Photo Section [NEW]
        photo_frame = ttk.LabelFrame(self.scrollable_frame, text="차량 사진 (클릭 시 크게 보기)")
        photo_frame.pack(fill='x', padx=5, pady=2)
        
        # [LARGER] Increased preview area size
        self.photo_canvas = tk.Label(photo_frame, text="사진 없음", background="#e0e0e0", width=60, height=15, cursor='hand2')
        self.photo_canvas.pack(pady=5)
        self.photo_canvas.bind("<Button-1>", lambda e: self.open_full_photo())
        
        self.photo_path = ""
        self.photo_image = None # Keep reference
        
        btn_photo_f = ttk.Frame(photo_frame)
        btn_photo_f.pack(fill='x', pady=2)
        
        ttk.Button(btn_photo_f, text="📸 사진 추가/변경", command=self.add_photo).pack(side='left', expand=True, fill='x', padx=2)
        ttk.Button(btn_photo_f, text="🗑️ 사진 삭제", command=self.clear_photo).pack(side='left', expand=True, fill='x', padx=2)

    def _ensure_exclusive(self, var, current_val):
        """Helper to ensure only one value is selected if needed (though StringVar handles it naturally)"""
        # This is primarily to handle clicking an already selected check to uncheck it if desired,
        # but ttk.Checkbutton with variable already does this for on/off values.
        pass

    def trigger_save(self):
        """Invoke the save callback provided by MaterialManager"""
        if self.on_save_callback:
            self.on_save_callback(self)

    def update_vehicle_list(self, new_list):
        """Update the combobox values (native dropdown)"""
        self.cb_vehicle_info['values'] = new_list

    def format_mileage(self, event=None):
        """Autoformat mileage with commas"""
        try:
            val = self.ent_mileage.get().strip().replace(',', '')
            if not val: return
            # Handle possible float if someone type 123.4
            if '.' in val:
                num = float(val)
                formatted = f"{num:,.1f}"
                if formatted.endswith('.0'): formatted = formatted[:-2]
            else:
                num = int(val)
                formatted = f"{num:,}"
            
            self.ent_mileage.delete(0, tk.END)
            self.ent_mileage.insert(0, formatted)
        except:
            pass

    def get_data(self):
        # Collect all inspection variables
        data = {key: var.get() for key, var in self.vars.items()}
        data['vehicle_info'] = self.cb_vehicle_info.get().strip()
        data['mileage'] = self.ent_mileage.get().strip()
        data['remarks'] = self.ent_remarks.get().strip()
        
        data['_raw_mileage'] = data['mileage'].replace(',', '')
        data['photo_path'] = self.photo_path
        return data

    def add_photo(self):
        """Open file dialog and copy photo to storage"""
        file_path = filedialog.askopenfilename(
            title="차량 사진 선택",
            filetypes=[("Image files", "*.jpg *.jpeg *.png *.bmp *.gif")]
        )
        if not file_path: return
        
        try:
            # Create directory if missing
            photo_dir = r"c:\Users\-\OneDrive\바탕 화면\home\data\vehicle_photos"
            if not os.path.exists(photo_dir):
                os.makedirs(photo_dir)
                
            v_no = self.cb_vehicle_info.get().strip()
            if not v_no:
                messagebox.showwarning("경고", "사진을 저장하기 전에 차량정보(차량번호)를 먼저 입력해주세요.")
                return
                
            # Clean filename
            clean_v_no = "".join(c for c in v_no if c.isalnum() or c in (' ', '-', '_')).strip()
            timestamp = int(time.time())
            ext = os.path.splitext(file_path)[1]
            new_filename = f"{clean_v_no}_{timestamp}{ext}"
            new_path = os.path.join(photo_dir, new_filename)
            
            # Copy file
            import shutil
            shutil.copy2(file_path, new_path)
            
            # Update and display
            self.photo_path = new_path
            self.display_photo()
            
        except Exception as e:
            messagebox.showerror("오류", f"사진을 저장하는 중 오류가 발생했습니다: {e}")

    def display_photo(self):
        """Load and display the photo in the widget"""
        if not self.photo_path or not os.path.exists(self.photo_path):
            self.photo_canvas.config(image='', text="사진 없음")
            return
            
        try:
            img = Image.open(self.photo_path)
            # [LARGER] Increased thumbnail size for better visibility
            img.thumbnail((600, 400))
            self.photo_image = ImageTk.PhotoImage(img)
            self.photo_canvas.config(image=self.photo_image, text="")
        except Exception as e:
            print(f"DEBUG: Error displaying photo: {e}")
            self.photo_canvas.config(image='', text="이미지 로드 오류")

    def open_full_photo(self):
        """Open the photo in a separate large window"""
        if not self.photo_path or not os.path.exists(self.photo_path): return
        
        full_win = tk.Toplevel(self)
        full_win.title("차량 사진 크게 보기")
        full_win.geometry("1000x800")
        
        try:
            img = Image.open(self.photo_path)
            # Maximize for the window
            img.thumbnail((1200, 900))
            full_photo = ImageTk.PhotoImage(img)
            
            # Keep reference to avoid GC
            full_win.full_photo = full_photo 
            
            canvas = tk.Canvas(full_win, bg='black')
            canvas.pack(fill='both', expand=True)
            
            canvas.create_image(500, 400, image=full_photo, anchor='center')
            
            ttk.Button(full_win, text="닫기", command=full_win.destroy).pack(pady=5)
        except Exception as e:
            messagebox.showerror("오류", f"이미지를 여는 중 오류가 발생했습니다: {e}")

    def clear_photo(self):
        """Remove photo from widget (doesn't delete file)"""
        self.photo_path = ""
        self.photo_image = None
        self.photo_canvas.config(image='', text="사진 없음")

    def reset_fields(self):
        """Clear all input fields for the next vehicle entry and set focus"""
        try:
            self.cb_vehicle_info.set('')
        except Exception as e: print(f"Error clearing cb_vehicle_info: {e}")
        
        try:
            self.ent_mileage.delete(0, tk.END)
        except Exception as e: print(f"Error clearing ent_mileage: {e}")
        
        try:
            self.ent_remarks.delete(0, tk.END)
        except Exception as e: print(f"Error clearing ent_remarks: {e}")
        
        for var in self.vars.values():
            try:
                var.set("") # Clear StringVar
            except Exception as e: print(f"Error clearing var: {e}")
            
        try:
            self.cb_vehicle_info.focus_set()
        except: pass
        
    def set_data(self, data):
        if not data: return
        for key, val in data.items():
            if key in self.vars:
                self.vars[key].set(val)
            elif key == 'vehicle_info':
                self.cb_vehicle_info.set(val)
            elif key == 'mileage':
                self.ent_mileage.delete(0, tk.END)
                self.ent_mileage.insert(0, val)
            elif key == 'remarks':
                self.ent_remarks.delete(0, tk.END)
                self.ent_remarks.insert(0, val)
            elif key == 'photo_path':
                self.photo_path = val
                self.display_photo()


class LaborCostDetailWidget(ttk.Frame):
    """
    Detailed labor cost calculation widget with three sections: 
    1) Regular Work (정시근무), 2) Special Work (특별근무), and 3) Base Salary Reference (기준급여)
    """
    def __init__(self, parent, on_change_callback=None, cost_mode='planned', **kwargs):
        super().__init__(parent, **kwargs)
        self.on_change_callback = on_change_callback
        self.cost_mode = cost_mode
        
        # Rankings for Table 1
        self.ranks = ["이사", "부장", "차장", "과장", "대리", "계장", "주임", "기사"]
        # Shift types for Table 2
        self.special_types = ["연장근무", "야간근무", "휴일근무"]
        
        # Base Salaries for Reference (Table 3) - Dynamically loaded from master
        self.base_salaries = {}
        
        # Resolve MaterialManager to get rates
        master = parent
        while master and not hasattr(master, 'get_base_salaries'):
            master = getattr(master, 'master', None)
        
        if master:
            self.base_salaries = master.get_base_salaries()
        else:
             self.base_salaries = {
                "이사": 55250000, "부장": 55250000, "차장": 47670000, "과장": 41170000,
                "대리": 37920000, "계장": 34670000, "주임": 31420000, "기사": 29250000
            }
        
        self.entries = {} # Key -> Rank/Type -> Column -> Entry
        self.totals = {}  # Key -> Rank/Type -> Label
        
        self._create_widgets()

    def get_total_cost(self):
        """[FINAL_FIX] Robustly get total cost from label at class level"""
        try:
            raw_text = getattr(self, 'lbl_grand_total', None)
            if raw_text:
                txt = raw_text.cget('text')
                val = "".join(c for c in txt if c.isdigit() or c == '.')
                return float(val or 0)
            return 0.0
        except:
            return 0.0

    def _create_widgets(self):
        style = ttk.Style()
        style.configure("LaborHeader.TLabel", font=('Malgun Gothic', 10, 'bold'), background='#e0e0e0', relief='solid')
        style.configure("LaborTotal.TLabel", font=('Malgun Gothic', 10, 'bold'), background='#fff9c4', relief='solid')
        style.configure("RefTitle.TLabel", font=('Malgun Gothic', 10, 'bold'), foreground='red', background='#fce4ec', relief='solid')
        
        # Main Container
        main_container = ttk.Frame(self)
        main_container.pack(fill='both', expand=True)
        
        # Left Side (Calculation)
        calc_frame = ttk.Frame(main_container)
        calc_frame.pack(side='left', fill='both', expand=True)
        
        # Right Side (Reference Table)
        ref_frame = ttk.Frame(main_container, padding=(20, 40, 0, 0))
        ref_frame.pack(side='right', fill='y')

        # --- Section 1: Regular Work (Left) ---
        ttk.Label(calc_frame, text="1) 정시근무 (240일/년)", font=('Malgun Gothic', 11, 'bold')).pack(anchor='w', pady=(10, 5))
        
        table1_frame = ttk.Frame(calc_frame)
        table1_frame.pack(fill='x')
        
        cost_header = "사후원가가액" if self.cost_mode == 'actual' else "사전원가가액"
        headers1 = ["구분", "직급", "투입인원(명)", "투입일수/인(일)", "단가/일", cost_header]
        for j, h in enumerate(headers1):
            lbl = ttk.Label(table1_frame, text=h, style="LaborHeader.TLabel", padding=5, anchor='center')
            lbl.grid(row=0, column=j, sticky='nsew')
            table1_frame.grid_columnconfigure(j, weight=1)
        enable_column_resize(table1_frame, len(headers1))

        merge_lbl = ttk.Label(table1_frame, text="정시근무\n(240일/년)", relief='solid', anchor='center', padding=10)
        merge_lbl.grid(row=1, column=0, rowspan=len(self.ranks), sticky='nsew')

        for i, rank in enumerate(self.ranks):
            row = i + 1
            ttk.Label(table1_frame, text=rank, relief='solid', anchor='center', padding=5).grid(row=row, column=1, sticky='nsew')
            
            self.entries[rank] = {}
            ent_personnel = ttk.Entry(table1_frame, width=10, justify='center')
            ent_personnel.grid(row=row, column=2, sticky='nsew')
            ent_personnel.bind("<KeyRelease>", lambda e, r=rank: self._on_input_change(r))
            self.entries[rank]['personnel'] = ent_personnel
            
            ent_days = ttk.Entry(table1_frame, width=10, justify='center')
            ent_days.grid(row=row, column=3, sticky='nsew')
            ent_days.bind("<KeyRelease>", lambda e, r=rank: self._on_input_change(r))
            self.entries[rank]['period'] = ent_days
            
            ent_price = ttk.Entry(table1_frame, width=15, justify='right')
            ent_price.grid(row=row, column=4, sticky='nsew')
            ent_price.bind("<KeyRelease>", lambda e, r=rank: self._on_input_change(r))
            self.entries[rank]['unit_price'] = ent_price
            
            # [NEW] Default value from base salary / 240
            daily_rate = round(self.base_salaries.get(rank, 0) / 240)
            if daily_rate > 0:
                ent_price.insert(0, f"{daily_rate:,.0f}")
            
            lbl_subtotal = ttk.Label(table1_frame, text="0", relief='solid', anchor='e', padding=5)
            lbl_subtotal.grid(row=row, column=5, sticky='nsew')
            self.totals[rank] = lbl_subtotal

        # Table 1 Subtotal Row
        row_t1_sum = len(self.ranks) + 1
        ttk.Label(table1_frame, text="소계", style="LaborHeader.TLabel", anchor='center', padding=5).grid(row=row_t1_sum, column=0, columnspan=2, sticky='nsew')
        self.lbl_t1_personnel_sum = ttk.Label(table1_frame, text="0 명", style="LaborHeader.TLabel", anchor='center')
        self.lbl_t1_personnel_sum.grid(row=row_t1_sum, column=2, sticky='nsew')
        self.lbl_t1_days_sum = ttk.Label(table1_frame, text="0 일", style="LaborHeader.TLabel", anchor='center')
        self.lbl_t1_days_sum.grid(row=row_t1_sum, column=3, sticky='nsew')
        ttk.Label(table1_frame, text="", style="LaborHeader.TLabel").grid(row=row_t1_sum, column=4, sticky='nsew')
        self.lbl_t1_cost_sum = ttk.Label(table1_frame, text="0", style="LaborHeader.TLabel", anchor='e', padding=5)
        self.lbl_t1_cost_sum.grid(row=row_t1_sum, column=5, sticky='nsew')

        # --- Section 2: Special Work (Left) ---
        ttk.Label(calc_frame, text="2) 특별근무", font=('Malgun Gothic', 11, 'bold')).pack(anchor='w', pady=(20, 5))
        
        table2_frame = ttk.Frame(calc_frame)
        table2_frame.pack(fill='x')
        
        headers2 = ["구분", "형태", "투입인원(명)", "투입시간/인(시간)", "단가", cost_header]
        for j, h in enumerate(headers2):
            lbl = ttk.Label(table2_frame, text=h, style="LaborHeader.TLabel", padding=5, anchor='center')
            lbl.grid(row=0, column=j, sticky='nsew')
            table2_frame.grid_columnconfigure(j, weight=1)
        enable_column_resize(table2_frame, len(headers2))

        merge_lbl2 = ttk.Label(table2_frame, text="특별근무", relief='solid', anchor='center', padding=10)
        merge_lbl2.grid(row=1, column=0, rowspan=len(self.special_types), sticky='nsew')

        for i, stype in enumerate(self.special_types):
            row = i + 1
            ttk.Label(table2_frame, text=stype, relief='solid', anchor='center', padding=5).grid(row=row, column=1, sticky='nsew')
            
            self.entries[stype] = {}
            ent_personnel = ttk.Entry(table2_frame, width=10, justify='center')
            ent_personnel.grid(row=row, column=2, sticky='nsew')
            ent_personnel.bind("<KeyRelease>", lambda e, s=stype: self._on_input_change(s))
            self.entries[stype]['personnel'] = ent_personnel
            
            ent_hours = ttk.Entry(table2_frame, width=10, justify='center')
            ent_hours.grid(row=row, column=3, sticky='nsew')
            ent_hours.bind("<KeyRelease>", lambda e, s=stype: self._on_input_change(s))
            self.entries[stype]['period'] = ent_hours
            
            ent_price = ttk.Entry(table2_frame, width=15, justify='right')
            ent_price.grid(row=row, column=4, sticky='nsew')
            ent_price.bind("<KeyRelease>", lambda e, s=stype: self._on_input_change(s))
            self.entries[stype]['unit_price'] = ent_price
            
            # [NEW] Default OT Rates
            ot_rates = {"연장근무": 4000, "야간근무": 5000, "휴일근무": 7500}
            if stype in ot_rates:
                ent_price.insert(0, f"{ot_rates[stype]:,.0f}")
            
            lbl_subtotal = ttk.Label(table2_frame, text="0", relief='solid', anchor='e', padding=5)
            lbl_subtotal.grid(row=row, column=5, sticky='nsew')
            self.totals[stype] = lbl_subtotal

        # Table 2 Subtotal Row
        row_t2_sum = len(self.special_types) + 1
        ttk.Label(table2_frame, text="소계", style="LaborHeader.TLabel", anchor='center', padding=5).grid(row=row_t2_sum, column=0, columnspan=2, sticky='nsew')
        self.lbl_t2_personnel_sum = ttk.Label(table2_frame, text="0 명", style="LaborHeader.TLabel", anchor='center')
        self.lbl_t2_personnel_sum.grid(row=row_t2_sum, column=2, sticky='nsew')
        self.lbl_t2_hours_sum = ttk.Label(table2_frame, text="0 시간", style="LaborHeader.TLabel", anchor='center')
        self.lbl_t2_hours_sum.grid(row=row_t2_sum, column=3, sticky='nsew')
        ttk.Label(table2_frame, text="", style="LaborHeader.TLabel").grid(row=row_t2_sum, column=4, sticky='nsew')
        self.lbl_t2_cost_sum = ttk.Label(table2_frame, text="0", style="LaborHeader.TLabel", anchor='e', padding=5)
        self.lbl_t2_cost_sum.grid(row=row_t2_sum, column=5, sticky='nsew')

        # --- Section 3: Base Salary Reference (Right) ---
        ttk.Label(ref_frame, text="고정금액(변경불가)", style="RefTitle.TLabel", padding=5).pack(fill='x')
        
        ref_table = ttk.Frame(ref_frame)
        ref_table.pack(fill='x')
        
        ttk.Label(ref_table, text="직급", style="LaborHeader.TLabel", padding=5, width=10, anchor='center').grid(row=0, column=0, sticky='nsew')
        ttk.Label(ref_table, text="기준급여", style="LaborHeader.TLabel", padding=5, width=15, anchor='center').grid(row=0, column=1, sticky='nsew')
        
        # Display ranks (skipping 이사 as it's not in the reference list image, but we have the data)
        display_ranks = ["부장", "차장", "과장", "대리", "계장", "주임", "기사"]
        for i, rank in enumerate(display_ranks):
            row = i + 1
            ttk.Label(ref_table, text=rank, relief='solid', padding=5, anchor='center').grid(row=row, column=0, sticky='nsew')
            salary = self.base_salaries.get(rank, 0)
            ttk.Label(ref_table, text=f"{salary:,.0f}", relief='solid', padding=5, anchor='e').grid(row=row, column=1, sticky='nsew')

        ttk.Button(ref_frame, text="기준급여 일괄 적용", command=self.apply_base_salaries).pack(pady=10, fill='x')

        # --- Grand Total ---
        total_frame = ttk.Frame(calc_frame)
        total_frame.pack(fill='x', pady=(10, 0))
        ttk.Label(total_frame, text="인건비 합계", style="LaborTotal.TLabel", anchor='center', padding=10).pack(side='left', fill='x', expand=True)
        self.lbl_grand_total = ttk.Label(total_frame, text="0", style="LaborTotal.TLabel", anchor='e', font=('Malgun Gothic', 12, 'bold'), padding=10)
        self.lbl_grand_total.pack(side='right', fill='x', expand=True)

    def apply_base_salaries(self):
        """Reset Daily Unit Price to the standard (Base Salary / 240)"""
        for rank, salary in self.base_salaries.items():
            if rank in self.entries:
                daily_rate = round(salary / 240)
                self.entries[rank]['unit_price'].delete(0, tk.END)
                self.entries[rank]['unit_price'].insert(0, f"{daily_rate:,.0f}")
                self._on_input_change(rank)
        messagebox.showinfo("적용 완료", "기준급여에 따른 일일 단가가 적용되었습니다.")

    def get_total_cost(self):
        """Retrieve the final calculated labor cost as a float"""
        try:
            val = self.lbl_grand_total.cget('text').replace('₩', '').replace(',', '').replace(' ', '').strip()
            return float(val or 0)
        except:
            return 0.0

    def _to_f(self, val):
        try:
            return float(str(val).replace(',', '') or 0)
        except:
            return 0.0

    def _on_input_change(self, key):
        # Calculate row total
        personnel = self._to_f(self.entries[key]['personnel'].get())
        period = self._to_f(self.entries[key]['period'].get())
        price = self._to_f(self.entries[key]['unit_price'].get())
        
        row_total = personnel * period * price
        self.totals[key].config(text=f"{row_total:,.0f}")
        
        self.calculate_all()

    def calculate_all(self):
        # Table 1 Totals
        t1_personnel = 0
        t1_days = 0
        t1_cost = 0
        for rank in self.ranks:
            p = self._to_f(self.entries[rank]['personnel'].get())
            d = self._to_f(self.entries[rank]['period'].get())
            c = self._to_f(self.totals[rank].cget('text'))
            t1_personnel += p
            t1_days += d
            t1_cost += c
        
        self.lbl_t1_personnel_sum.config(text=f"{t1_personnel:g} 명")
        self.lbl_t1_days_sum.config(text=f"{t1_days:g} 일")
        self.lbl_t1_cost_sum.config(text=f"{t1_cost:,.0f}")

        # Table 2 Totals
        t2_personnel = 0
        t2_hours = 0
        t2_cost = 0
        for stype in self.special_types:
            p = self._to_f(self.entries[stype]['personnel'].get())
            h = self._to_f(self.entries[stype]['period'].get())
            c = self._to_f(self.totals[stype].cget('text'))
            t2_personnel = max(t2_personnel, p)  # [FIX] 투입인원은 동일 인원이 연장/야간/휴일을 병행하므로 합계(sum) 대신 최대값(max) 적용
            t2_hours += h
            t2_cost += c
            
        self.lbl_t2_personnel_sum.config(text=f"{t2_personnel:g} 명")
        self.lbl_t2_hours_sum.config(text=f"{t2_hours:g} 시간")
        self.lbl_t2_cost_sum.config(text=f"{t2_cost:,.0f}")

        grand_total = t1_cost + t2_cost
        self.lbl_grand_total.config(text=f"₩ {grand_total:,.0f}")
        
        if self.on_change_callback:
            self.on_change_callback(grand_total)

    def get_data(self):
        """Export all entry values as a dictionary"""
        data = {}
        for key, widgets in self.entries.items():
            data[key] = {
                'personnel': widgets['personnel'].get(),
                'period': widgets['period'].get(),
                'unit_price': widgets['unit_price'].get()
            }
        return data

    def set_data(self, data):
        """Populate entries from a dictionary"""
        if not data or not isinstance(data, dict):
            self.reset()
            return
            
        for key, values in data.items():
            if key in self.entries:
                self.entries[key]['personnel'].delete(0, tk.END); self.entries[key]['personnel'].insert(0, values.get('personnel', ''))
                self.entries[key]['period'].delete(0, tk.END); self.entries[key]['period'].insert(0, values.get('period', ''))
                self.entries[key]['unit_price'].delete(0, tk.END); self.entries[key]['unit_price'].insert(0, values.get('unit_price', ''))
        
        # Trigger all row calculations
        for key in list(self.ranks) + list(self.special_types):
            self._on_input_change(key)

    def reset(self):
        """Clear all entries"""
        for key, widgets in self.entries.items():
            widgets['personnel'].delete(0, tk.END)
            widgets['period'].delete(0, tk.END)
            widgets['unit_price'].delete(0, tk.END)
        self.calculate_all()


class MaterialCostDetailWidget(ttk.Frame):
    """
    Detailed material cost calculation widget.
    Columns: Item (품목), Spec (사양), Quantity (수량), Unit (규격), Price (단가), Amount (사전원가가액)
    """
    def __init__(self, parent, on_change_callback=None, **kwargs):
        super().__init__(parent, **kwargs)
        self.on_change_callback = on_change_callback
        
        # Default Items from settings_df or fallback
        self.default_items = []
        
        # Resolve MaterialManager to get rates
        master = parent
        while master and not hasattr(master, 'get_material_defaults'):
            master = getattr(master, 'master', None)
        
        if master:
            self.default_items = master.get_material_defaults()
        else:
            self.default_items = [
                ("PT 약품", "세척제", "CAN", 1500), ("PT 약품", "침투제", "CAN", 2300),
                ("PT 약품", "현상제", "CAN", 2000), ("MT 약품", "백색페인트", "CAN", 2350),
                ("MT 약품", "흑색자분", "CAN", 1800), ("방사선투과검사 필름", "MX125", "매", 990),
                ("글리세린", "20L", "통", 100000), ("필름 현상액", "3L", "통", 16500),
                ("필름 정착액", "3L", "통", 16500), ("수적방지액", "200mL", "통", 2500)
            ]
        
        self.entries = [] # List of dicts for each row: {'item': lbl, 'spec': lbl, 'qty': ent, 'unit': lbl, 'price': ent, 'amount': lbl}
        self._create_widgets()

    def _create_widgets(self):
        style = ttk.Style()
        style.configure("MatHeader.TLabel", font=('Malgun Gothic', 10, 'bold'), background='#e0e0e0', relief='solid')
        style.configure("MatTotal.TLabel", font=('Malgun Gothic', 10, 'bold'), background='#ffff00', relief='solid') # Yellow as in image
        style.configure("MatRef.TLabel", font=('Malgun Gothic', 9), background='#f5f5f5', relief='solid')

        # Main Container
        main_container = ttk.Frame(self)
        main_container.pack(fill='both', expand=True)

        # Left Side (Calc)
        calc_frame = ttk.Frame(main_container)
        calc_frame.pack(side='left', fill='both', expand=True)

        # Right Side (Ref)
        ref_frame = ttk.Frame(main_container, padding=(20, 40, 0, 0))
        ref_frame.pack(side='right', fill='y')

        # --- Section 1: Tables (Left) ---
        ttk.Label(calc_frame, text="2) 재료비", font=('Malgun Gothic', 11, 'bold')).pack(anchor='w', pady=(10, 5))
        
        table_frame = ttk.Frame(calc_frame)
        table_frame.pack(fill='x')
        
        headers = ["품목", "사양", "수량", "규격", "단가", "사전원가가액"]
        widths = [20, 15, 10, 8, 15, 20]
        for j, (h, w) in enumerate(zip(headers, widths)):
            lbl = ttk.Label(table_frame, text=h, style="MatHeader.TLabel", padding=5, anchor='center', width=w)
            lbl.grid(row=0, column=j, sticky='nsew')
            table_frame.grid_columnconfigure(j, weight=1 if j in [0, 5] else 0)
        enable_column_resize(table_frame, len(headers))

        for i, (item, spec, unit, price) in enumerate(self.default_items):
            row = i + 1
            # Item
            ttk.Label(table_frame, text=item, relief='solid', padding=5, anchor='center').grid(row=row, column=0, sticky='nsew')
            # Spec
            ttk.Label(table_frame, text=spec, relief='solid', padding=5, anchor='center').grid(row=row, column=1, sticky='nsew')
            
            row_widgets = {}
            # Quantity
            ent_qty = ttk.Entry(table_frame, width=10, justify='center')
            ent_qty.grid(row=row, column=2, sticky='nsew')
            ent_qty.bind("<KeyRelease>", lambda e, idx=i: self._on_input_change(idx))
            row_widgets['qty'] = ent_qty
            
            # Unit
            ttk.Label(table_frame, text=unit, relief='solid', padding=5, anchor='center').grid(row=row, column=3, sticky='nsew')
            
            # Unit Price
            ent_price = ttk.Entry(table_frame, width=15, justify='right')
            ent_price.grid(row=row, column=4, sticky='nsew')
            ent_price.insert(0, f"{price:,.0f}")
            ent_price.bind("<KeyRelease>", lambda e, idx=i: self._on_input_change(idx))
            row_widgets['price'] = ent_price
            
            # Amount
            lbl_amount = ttk.Label(table_frame, text="0", relief='solid', anchor='e', padding=5)
            lbl_amount.grid(row=row, column=5, sticky='nsew')
            row_widgets['amount'] = lbl_amount
            
            self.entries.append(row_widgets)

        # Footer 1: Total
        row_footer = len(self.default_items) + 1
        ttk.Label(table_frame, text="재료비 합계", style="MatTotal.TLabel", anchor='center', padding=5).grid(row=row_footer, column=0, columnspan=5, sticky='nsew')
        self.lbl_mat_total = ttk.Label(table_frame, text="₩ 0", style="MatTotal.TLabel", anchor='e', padding=5)
        self.lbl_mat_total.grid(row=row_footer, column=5, sticky='nsew')

        # Footer 2: VAT
        row_vat = len(self.default_items) + 2
        ttk.Label(table_frame, text="부가세", relief='solid', anchor='center', padding=5).grid(row=row_vat, column=4, sticky='nsew')
        self.lbl_mat_vat = ttk.Label(table_frame, text="0", relief='solid', anchor='e', padding=5)
        self.lbl_mat_vat.grid(row=row_vat, column=5, sticky='nsew')

        # --- Section 2: Ref Info (Right) ---
        ref_table = ttk.Frame(ref_frame)
        ref_table.pack(fill='x')

        # Headers for reference
        ref_data = [
            ("지에스켐", "30M"),
            ("지에스켐", ""),
            ("지에스켐", ""),
            ("지에스켐", "32M"),
            ("지에스켐", ""),
            ("한스", "3 1/3 x 12\" (1850 매)"),
            ("한스", "3 1/3 x 6\" (313 매)"),
            ("한스", "총 매수: 2,163"),
            ("경도", "500매 기준 / 1회 4통 / 5회 교환"),
            ("나우", "500매 기준 / 1회 2동 / 5회 교환")
        ]

        for i, (title, val) in enumerate(ref_data):
            row = i
            ttk.Label(ref_table, text=title, width=10, anchor='center', style="MatRef.TLabel", padding=3).grid(row=row, column=0, sticky='nsew')
            ttk.Label(ref_table, text=val, width=25, anchor='w', style="MatRef.TLabel", padding=3).grid(row=row, column=1, sticky='nsew')

    def _on_input_change(self, idx):
        widgets = self.entries[idx]
        qty = self._to_f(widgets['qty'].get())
        price = self._to_f(widgets['price'].get())
        
        amount = qty * price
        widgets['amount'].config(text=f"{amount:,.0f}")
        
        self.calculate_all()

    def calculate_all(self):
        total_mat = 0.0
        for widgets in self.entries:
            qty = self._to_f(widgets['qty'].get())
            price = self._to_f(widgets['price'].get())
            amount = qty * price
            widgets['amount'].config(text=f"{amount:,.0f}")
            total_mat += amount
            
        vat = total_mat * 0.1
        self.lbl_mat_total.config(text=f"₩ {total_mat:,.0f}")
        self.lbl_mat_vat.config(text=f"{vat:,.0f}")
        
        if self.on_change_callback:
            self.on_change_callback(total_mat)

    def _to_f(self, val):
        try:
            return float(str(val).replace(',', '') or 0)
        except:
            return 0.0

    def get_data(self):
        """Export entry values"""
        data = []
        for widgets in self.entries:
            data.append({
                'qty': widgets['qty'].get(),
                'price': widgets['price'].get()
            })
        return data

    def set_data(self, data):
        """Populate entries"""
        if not data or not isinstance(data, list):
            self.reset()
            return
            
        for i, val in enumerate(data):
            if i < len(self.entries):
                self.entries[i]['qty'].delete(0, tk.END); self.entries[i]['qty'].insert(0, val.get('qty', ''))
                self.entries[i]['price'].delete(0, tk.END); self.entries[i]['price'].insert(0, val.get('price', ''))
        
        self.calculate_all()

    def reset(self):
        """Clear quantities and restore default prices"""
        for i, widgets in enumerate(self.entries):
            widgets['qty'].delete(0, tk.END)
            # Restore default price
            default_price = self.default_items[i][3]
            widgets['price'].delete(0, tk.END); widgets['price'].insert(0, f"{default_price:,.0f}")
            
        self.calculate_all()


class ExpenseProfitDetailWidget(ttk.Frame):
    """
    Comprehensive expense and profit calculation widget.
    Sections: 1) Site Expenses, 2) Rental, 3) Outsource, 4) Insurance, 5) Depreciation, 6) Indirect Cost, 7) Profit
    """
    def __init__(self, parent, on_change_callback=None, get_labor_total_func=None, get_material_total_func=None, get_revenue_func=None, budget_mode=None, **kwargs):
        super().__init__(parent, **kwargs)
        self.on_change_callback = on_change_callback
        self.get_labor_total = get_labor_total_func
        self.get_material_total = get_material_total_func
        self.get_revenue = get_revenue_func
        self.budget_mode = budget_mode
        
        self.entries = {
            'site_expense': [], # list of dicts
            'rental': [],
            'outsource': [],
            'depreciation': []
        }
        
        # Resolve MaterialManager to get rates
        self.master_app = parent
        while self.master_app and not hasattr(self.master_app, 'get_expense_defaults'):
            self.master_app = getattr(self.master_app, 'master', None)
            
        self._create_widgets()

    def _create_widgets(self):
        style = ttk.Style()
        style.configure("ExpHeader.TLabel", font=('Malgun Gothic', 10, 'bold'), background='#e0e0e0', relief='solid')
        style.configure("ExpTotal.TLabel", font=('Malgun Gothic', 10, 'bold'), background='#ffff00', relief='solid')
        style.configure("Margin.TLabel", font=('Malgun Gothic', 10, 'bold'), background='#90ee90', relief='solid') # Light green for profit

        # --- Section 1: Site Expenses ---
        ttk.Label(self, text="3) 경비", font=('Malgun Gothic', 11, 'bold')).pack(anchor='w', pady=(10, 5))
        
        s1_frame = ttk.LabelFrame(self, text="(1) 현장 경비")
        s1_frame.pack(fill='x', pady=5)
        
        self.s1_table = ttk.Frame(s1_frame)
        self.s1_table.pack(fill='x')
        
        headers = ["구분", "내용", "인원수", "수량", "규격", "단가", "금액(원)"]
        widths = [15, 20, 8, 8, 8, 15, 20]
        for j, (h, w) in enumerate(zip(headers, widths)):
            ttk.Label(self.s1_table, text=h, style="ExpHeader.TLabel", padding=5, anchor='center', width=w).grid(row=0, column=j, sticky='nsew')
            self.s1_table.grid_columnconfigure(j, weight=1 if j in [1, 6] else 0)
        enable_column_resize(self.s1_table, len(headers))

        # Default Site Expenses - Dynamically loaded from master
        defaults_s1 = []
        if self.master_app:
            defaults_s1 = self.master_app.get_expense_defaults()
        else:
            defaults_s1 = [
                ("차량유지비", "주유, 수리, 통행, 주차 등", "N/A", 1, "일", 5000),
                ("소모품비", "장갑,일회용 작업복외", "N/A", 1, "일", 500),
                ("복리후생비", "생수, 음료 외 기타", "N/A", 1, "일", 1667),
                ("Se-175", "방사성동위원소 구매", "N/A", 1, "일", 47619)
            ]
        
        for i, (cat, cont, ppl, qty, unit, price) in enumerate(defaults_s1):
            if self.budget_mode == 'actual':
                actual_rates = {
                    '차량유지비': ('일', 5000),
                    '소모품비': ('일', 500),
                    '복리후생비': ('일', 1667),
                    'Se-175': ('일', 47619),
                }
                actual_unit, actual_price = actual_rates.get(str(cat).strip(), (unit, price))
                self._add_row_s1(cat, cont, ppl, '', actual_unit, actual_price)
            else:
                self._add_row_s1(cat, cont, ppl, qty, unit, price)

        # --- Section 2: Rental Costs ---
        s2_frame = ttk.LabelFrame(self, text="(2) 장비/차량 임차료")
        s2_frame.pack(fill='x', pady=5)
        
        self.s2_table = ttk.Frame(s2_frame)
        self.s2_table.pack(fill='x')
        
        headers2 = ["구분", "사양", "수량", "사용기간", "기간단위", "단가/월,대", "금액(원)"]
        for j, h in enumerate(headers2):
            ttk.Label(self.s2_table, text=h, style="ExpHeader.TLabel", padding=5, anchor='center', width=widths[j]).grid(row=0, column=j, sticky='nsew')
            self.s2_table.grid_columnconfigure(j, weight=1 if j in [1, 6] else 0)
        enable_column_resize(self.s2_table, len(headers2))
            
        # Add 3 empty rows by default
        for _ in range(3):
            self._add_row_s2()

        # --- Section 3: Outsource Costs ---
        s3_frame = ttk.LabelFrame(self, text="(3) 외주비/잡급")
        s3_frame.pack(fill='x', pady=5)
        
        self.s3_table = ttk.Frame(s3_frame)
        self.s3_table.pack(fill='x')
        
        headers3 = ["구분", "작업내용", "공수", "단가", "금액(원)"]
        widths3 = [15, 30, 10, 15, 20]
        for j, h in enumerate(headers3):
            ttk.Label(self.s3_table, text=h, style="ExpHeader.TLabel", padding=5, anchor='center', width=widths3[j]).grid(row=0, column=j, sticky='nsew')
            self.s3_table.grid_columnconfigure(j, weight=1 if j in [1, 4] else 0)
        enable_column_resize(self.s3_table, len(headers3))
            
        # Outsource defaults from master
        outsource_defaults = []
        if self.master_app:
            outsource_defaults = self.master_app.get_outsource_defaults()
        else:
            outsource_defaults = [
                ("케이엔디이",     "방사선투과검사", 0, 15000),
                ("고려검사",       "방사선투과검사", 0, 13000),
                ("한국기계검사소", "방사선투과검사", 0, 15000),
            ]

        for cat, content, qty, price in outsource_defaults:
            if self.budget_mode == 'actual':
                initial_count = {
                    '케이엔디이': 3918,
                    '고려검사': 651,
                    '한국기계검사소': 453,
                }.get(str(cat).strip(), '')
            else:
                initial_count = '' if self.budget_mode == 'actual' else qty
            self._add_row_s3(cat, content, initial_count, price)

        for _ in range(2): self._add_row_s3()

        # --- Section 4: Social Insurance ---
        s4_frame = ttk.Frame(self)
        s4_frame.pack(fill='x', pady=5)
        ttk.Label(s4_frame, text="(4) 4대 보험료", font=('Malgun Gothic', 10, 'bold')).pack(side='left', padx=5)
        
        insurance_table = ttk.Frame(self)
        insurance_table.pack(fill='x')
        headers4 = ["구분", "산출 기준", "산출 인건비", "단가(요율)", "금액(원)"]
        for j, h in enumerate(headers4):
            ttk.Label(insurance_table, text=h, style="ExpHeader.TLabel", padding=5, anchor='center', width=widths[j] if j < len(widths) else 20).grid(row=0, column=j, sticky='nsew')
            insurance_table.grid_columnconfigure(j, weight=1 if j in [1, 4] else 0)
        enable_column_resize(insurance_table, len(headers4))
            
        ttk.Label(insurance_table, text="4대 보험료", relief='solid', padding=5, anchor='center').grid(row=1, column=0, sticky='nsew')
        ttk.Label(insurance_table, text="산출인건비 X 요율(2024.7.1 기준)", relief='solid', padding=5, anchor='w').grid(row=1, column=1, sticky='nsew')
        self.lbl_insurance_base = ttk.Label(insurance_table, text="₩ 0", relief='solid', padding=5, anchor='e')
        self.lbl_insurance_base.grid(row=1, column=2, sticky='nsew')
        ttk.Label(insurance_table, text="10.6661%", relief='solid', padding=5, anchor='center').grid(row=1, column=3, sticky='nsew')
        self.lbl_insurance_amount = ttk.Label(insurance_table, text="0", relief='solid', padding=5, anchor='e')
        self.lbl_insurance_amount.grid(row=1, column=4, sticky='nsew')

        # --- Section 5: Depreciation ---
        s5_frame = ttk.LabelFrame(self, text="(5) 감가상각비 (Depreciation)")
        s5_frame.pack(fill='x', pady=5)
        
        self.s5_table = ttk.Frame(s5_frame)
        self.s5_table.pack(fill='x')
        
        headers5 = ["장비명", "사양", "내용년수", "수량", "사용일수", "감가비/일", "금액(원)"]
        for j, h in enumerate(headers5):
            ttk.Label(self.s5_table, text=h, style="ExpHeader.TLabel", padding=5, anchor='center', width=widths[j]).grid(row=0, column=j, sticky='nsew')
            self.s5_table.grid_columnconfigure(j, weight=1 if j in [1, 6] else 0)
        enable_column_resize(self.s5_table, len(headers5))
            
        defaults_s5 = [
            ("PAUT 장비", "", 5, 1, 0, 44444),
            ("PAUT SCANNER (MANUAL)", "", 5, 1, 0, 5556),
            ("PAUT SCANNER (COBRA)", "", 5, 1, 0, 16667),
            ("YOKE", "", 5, 1, 0, 222),
            ("PMI 장비", "", 5, 1, 0, 12778),
            ("UT 장비", "", 5, 1, 0, 7778),
            ("현상용 탑차(5년간 보험비 포함)", "현장별 차량기입시 탑차 구분 기입", 5, 1, 0, 16667),
            ("스타렉스(5년간 보험비 포함)", "현장별 차량기입시 스타렉스 구분 기입", 5, 1, 0, 16667)
        ]
        for item, spec, life, qty, days, rate in defaults_s5:
            self._add_row_s5(item, spec, life, qty, days, rate)

        # --- TOTALS SUMMARY ---
        summary_frame = ttk.Frame(self, padding=10)
        summary_frame.pack(fill='x', pady=10)
        
        # Row: Expense Total (1~5)
        ttk.Label(summary_frame, text="경비 합계 : (1)~(5) 합계", style="ExpTotal.TLabel", anchor='center', padding=8).grid(row=0, column=0, columnspan=4, sticky='nsew')
        self.lbl_exp_total = ttk.Label(summary_frame, text="₩ 0", style="ExpTotal.TLabel", anchor='e', padding=8)
        self.lbl_exp_total.grid(row=0, column=4, sticky='nsew')
        
        # Row: Expense VAT
        ttk.Label(summary_frame, text="경비 부가세 합계", relief='solid', anchor='center', padding=5).grid(row=1, column=3, sticky='nsew')
        self.lbl_exp_vat = ttk.Label(summary_frame, text="0", relief='solid', anchor='e', padding=5)
        self.lbl_exp_vat.grid(row=1, column=4, sticky='nsew')

        # Row: Total Direct Cost (Sales Cost)
        ttk.Label(summary_frame, text="매출원가 총계 : 1), 2), 3) 합계", font=('Malgun Gothic', 10, 'bold'), background='#00ffff', relief='solid', anchor='center', padding=8).grid(row=2, column=0, columnspan=4, sticky='nsew')
        self.lbl_sales_cost_total = ttk.Label(summary_frame, text="₩ 0", font=('Malgun Gothic', 10, 'bold'), background='#00ffff', relief='solid', anchor='e', padding=8)
        self.lbl_sales_cost_total.grid(row=2, column=4, sticky='nsew')

        # Row: Indirect Cost (판관비)
        ttk.Label(summary_frame, text="3. 간접비(판관비)", font=('Malgun Gothic', 10, 'bold'), anchor='w').grid(row=3, column=0, pady=(10, 0))
        
        indirect_table = ttk.Frame(summary_frame)
        indirect_table.grid(row=4, column=0, columnspan=5, sticky='ew')
        headers_ind = ["구분", "산출 기준", "산출직접비", "간접비율(%)", "사전 간접비 합계"]
        for j, h in enumerate(headers_ind):
             ttk.Label(indirect_table, text=h, style="ExpHeader.TLabel", padding=5, anchor='center', width=widths[j] if j < len(widths) else 25).grid(row=0, column=j, sticky='nsew')
             indirect_table.grid_columnconfigure(j, weight=1 if j in [1, 4] else 0)
        enable_column_resize(indirect_table, len(headers_ind))
        
        ttk.Label(indirect_table, text="간접비", relief='solid', padding=5, anchor='center').grid(row=1, column=0, sticky='nsew')
        ttk.Label(indirect_table, text="산출직접비 x 간접비율(2024년 기준)", relief='solid', padding=5, anchor='w').grid(row=1, column=1, sticky='nsew')
        self.lbl_indirect_base = ttk.Label(indirect_table, text="₩ 0", relief='solid', padding=5, anchor='e')
        self.lbl_indirect_base.grid(row=1, column=2, sticky='nsew')
        ttk.Label(indirect_table, text="14%", relief='solid', padding=5, anchor='center').grid(row=1, column=3, sticky='nsew')
        self.lbl_indirect_total = ttk.Label(indirect_table, text="₩ 0", font=('Malgun Gothic', 10, 'bold'), background='#00ffff', relief='solid', anchor='e', padding=5)
        self.lbl_indirect_total.grid(row=1, column=4, sticky='nsew')

        # Row: Total Cost (Direct + Indirect)
        ttk.Label(summary_frame, text="4. 총원가(매출원가+간접비)", style="ExpTotal.TLabel", anchor='center', padding=10).grid(row=5, column=0, columnspan=4, sticky='nsew', pady=(10, 0))
        self.lbl_grand_total_cost = ttk.Label(summary_frame, text="₩ 0", style="ExpTotal.TLabel", anchor='e', padding=10)
        self.lbl_grand_total_cost.grid(row=5, column=4, sticky='nsew', pady=(10, 0))

        # Row: Operating Profit Section
        ttk.Label(summary_frame, text="5. 영업이익, 영업이익률", font=('Malgun Gothic', 10, 'bold'), anchor='w').grid(row=6, column=0, pady=(10, 0))
        
        profit_table = ttk.Frame(summary_frame)
        profit_table.grid(row=7, column=0, columnspan=5, sticky='ew')
        headers_prof = ["구분", "매출(수입)", "총원가", "영업이익", "영업이익률", "기준"]
        widths_prof = [15, 20, 20, 20, 15, 10]
        for j, h in enumerate(headers_prof):
             ttk.Label(profit_table, text=h, style="ExpHeader.TLabel", padding=5, anchor='center', width=widths_prof[j]).grid(row=0, column=j, sticky='nsew')
             profit_table.grid_columnconfigure(j, weight=1 if j != 5 else 0)
        enable_column_resize(profit_table, len(headers_prof))

        # Row 1: Budget (사전)
        ttk.Label(profit_table, text="사전(예산)", relief='solid', padding=5, anchor='center').grid(row=1, column=0, sticky='nsew')
        self.lbl_prof_revenue = ttk.Label(profit_table, text="₩ 0", relief='solid', padding=5, anchor='e')
        self.lbl_prof_revenue.grid(row=1, column=1, sticky='nsew')
        self.lbl_prof_total_cost = ttk.Label(profit_table, text="₩ 0", relief='solid', padding=5, anchor='e')
        self.lbl_prof_total_cost.grid(row=1, column=2, sticky='nsew')
        self.lbl_prof_op_profit = ttk.Label(profit_table, text="₩ 0", style="Margin.TLabel", padding=5, anchor='center')
        self.lbl_prof_op_profit.grid(row=1, column=3, sticky='nsew')
        self.lbl_prof_margin = ttk.Label(profit_table, text="0.00%", relief='solid', padding=5, anchor='center')
        self.lbl_prof_margin.grid(row=1, column=4, sticky='nsew')
        ttk.Label(profit_table, text="부가세 별도", relief='solid', padding=5, anchor='center').grid(row=1, column=5, sticky='nsew')

    def _add_row_s1(self, cat="", cont="", ppl="", qty="", unit="", price=0):
        row = len(self.entries['site_expense']) + 1
        widgets = {}
        ent_cat = ttk.Entry(self.s1_table, width=15, justify='center'); ent_cat.insert(0, cat); ent_cat.grid(row=row, column=0, sticky='nsew')
        ent_cont = ttk.Entry(self.s1_table, width=20); ent_cont.insert(0, cont); ent_cont.grid(row=row, column=1, sticky='nsew')
        ent_ppl = ttk.Entry(self.s1_table, width=8, justify='center'); ent_ppl.insert(0, str(ppl)); ent_ppl.grid(row=row, column=2, sticky='nsew')
        ent_qty = ttk.Entry(self.s1_table, width=8, justify='center'); ent_qty.insert(0, str(qty)); ent_qty.grid(row=row, column=3, sticky='nsew')
        ent_unit = ttk.Entry(self.s1_table, width=8, justify='center'); ent_unit.insert(0, unit); ent_unit.grid(row=row, column=4, sticky='nsew')
        ent_price = ttk.Entry(self.s1_table, width=15, justify='right'); ent_price.insert(0, f"{price:,.0f}"); ent_price.grid(row=row, column=5, sticky='nsew')
        lbl_amt = ttk.Label(self.s1_table, text="0", relief='solid', anchor='e', padding=5); lbl_amt.grid(row=row, column=6, sticky='nsew')
        
        widgets = {'cat': ent_cat, 'cont': ent_cont, 'ppl': ent_ppl, 'qty': ent_qty, 'unit': ent_unit, 'price': ent_price, 'amount': lbl_amt}
        for w in [ent_ppl, ent_qty, ent_price]: w.bind("<KeyRelease>", lambda e: self.calculate_all())
        self.entries['site_expense'].append(widgets)

    def _add_row_s2(self, cat="", spec="", qty="", period="", unit="", price=0):
        row = len(self.entries['rental']) + 1
        widgets = {}
        e1 = ttk.Entry(self.s2_table, width=15, justify='center'); e1.insert(0, cat); e1.grid(row=row, column=0, sticky='nsew')
        e2 = ttk.Entry(self.s2_table, width=20); e2.insert(0, spec); e2.grid(row=row, column=1, sticky='nsew')
        e3 = ttk.Entry(self.s2_table, width=8, justify='center'); e3.insert(0, str(qty)); e3.grid(row=row, column=2, sticky='nsew')
        e4 = ttk.Entry(self.s2_table, width=8, justify='center'); e4.insert(0, str(period)); e4.grid(row=row, column=3, sticky='nsew')
        e5 = ttk.Entry(self.s2_table, width=8, justify='center'); e5.insert(0, unit); e5.grid(row=row, column=4, sticky='nsew')
        e6 = ttk.Entry(self.s2_table, width=15, justify='right'); e6.insert(0, f"{price:,.0f}"); e6.grid(row=row, column=5, sticky='nsew')
        lbl = ttk.Label(self.s2_table, text="0", relief='solid', anchor='e', padding=5); lbl.grid(row=row, column=6, sticky='nsew')
        
        widgets = {'cat': e1, 'spec': e2, 'qty': e3, 'period': e4, 'unit': e5, 'price': e6, 'amount': lbl}
        for w in [e3, e4, e6]: w.bind("<KeyRelease>", lambda e: self.calculate_all())
        self.entries['rental'].append(widgets)

    def _add_row_s3(self, cat="", work="", count=0, price=0):
        row = len(self.entries['outsource']) + 1
        widgets = {}
        e1 = ttk.Entry(self.s3_table, width=15, justify='center'); e1.insert(0, cat); e1.grid(row=row, column=0, sticky='nsew')
        e2 = ttk.Entry(self.s3_table, width=30); e2.insert(0, work); e2.grid(row=row, column=1, sticky='nsew')
        e3 = ttk.Entry(self.s3_table, width=10, justify='center'); e3.insert(0, str(count)); e3.grid(row=row, column=2, sticky='nsew')
        e4 = ttk.Entry(self.s3_table, width=15, justify='right'); e4.insert(0, f"{price:,.0f}"); e4.grid(row=row, column=3, sticky='nsew')
        lbl = ttk.Label(self.s3_table, text="0", relief='solid', anchor='e', padding=5); lbl.grid(row=row, column=4, sticky='nsew')
        
        widgets = {'cat': e1, 'work': e2, 'count': e3, 'price': e4, 'amount': lbl}
        for w in [e3, e4]: w.bind("<KeyRelease>", lambda e: self.calculate_all())
        self.entries['outsource'].append(widgets)

    def _add_row_s5(self, item="", spec="", life=5, qty=1, days=0, rate=0):
        row = len(self.entries['depreciation']) + 1
        widgets = {}
        e1 = ttk.Entry(self.s5_table, width=20); e1.insert(0, item); e1.grid(row=row, column=0, sticky='nsew')
        e2 = ttk.Entry(self.s5_table, width=15); e2.insert(0, spec); e2.grid(row=row, column=1, sticky='nsew')
        e3 = ttk.Entry(self.s5_table, width=8, justify='center'); e3.insert(0, str(life)); e3.grid(row=row, column=2, sticky='nsew')
        e4 = ttk.Entry(self.s5_table, width=8, justify='center'); e4.insert(0, str(qty)); e4.grid(row=row, column=3, sticky='nsew')
        e5 = ttk.Entry(self.s5_table, width=8, justify='center'); e5.insert(0, str(days)); e5.grid(row=row, column=4, sticky='nsew')
        e6 = ttk.Entry(self.s5_table, width=15, justify='right'); e6.insert(0, f"{rate:,.0f}"); e6.grid(row=row, column=5, sticky='nsew')
        lbl = ttk.Label(self.s5_table, text="0", relief='solid', anchor='e', padding=5); lbl.grid(row=row, column=6, sticky='nsew')
        
        widgets = {'item': e1, 'spec': e2, 'life': e3, 'qty': e4, 'days': e5, 'rate': e6, 'amount': lbl}
        for w in [e4, e5, e6]: w.bind("<KeyRelease>", lambda e: self.calculate_all())
        self.entries['depreciation'].append(widgets)

    def calculate_all(self, event=None):
        # 1. Site Expenses
        t1 = 0.0
        for w in self.entries['site_expense']:
            amt = self._to_f(w['qty'].get()) * self._to_f(w['price'].get())
            w['amount'].config(text=f"{amt:,.0f}")
            t1 += amt
            
        # 2. Rentals
        t2 = 0.0
        for w in self.entries['rental']:
            amt = self._to_f(w['qty'].get()) * self._to_f(w['period'].get()) * self._to_f(w['price'].get())
            w['amount'].config(text=f"{amt:,.0f}")
            t2 += amt
            
        # 3. Outsource
        t3 = 0.0
        for w in self.entries['outsource']:
            amt = self._to_f(w['count'].get()) * self._to_f(w['price'].get())
            w['amount'].config(text=f"{amt:,.0f}")
            t3 += amt
            
        # 4. Insurance
        labor_total = self.get_labor_total() if self.get_labor_total else 0.0
        t4 = labor_total * 0.106661
        self.lbl_insurance_base.config(text=f"₩ {labor_total:,.0f}")
        self.lbl_insurance_amount.config(text=f"{t4:,.0f}")
        
        # 5. Depreciation
        t5 = 0.0
        for w in self.entries['depreciation']:
            amt = self._to_f(w['qty'].get()) * self._to_f(w['days'].get()) * self._to_f(w['rate'].get())
            w['amount'].config(text=f"{amt:,.0f}")
            t5 += amt
            
        # 실행 경비와 외주비는 요약 화면에서 별도 항목으로 관리한다.
        # 외주비(t3)를 경비에 포함하면 영업이익 계산 시 외주비가 두 번 차감된다.
        exp_total = t1 + t2 + t4 + t5

        # 롯데 전용 화면은 생성 시 planned/actual 모드를 명시적으로 전달한다.
        # Tk 위젯의 master 체인만으로는 MaterialManager 컨트롤러에 도달하지 못할 수 있다.
        lotte_rules = self.budget_mode in ('planned', 'actual') or bool(getattr(self.master_app, 'lotte_budget_rules', False))
        if lotte_rules:
            # 롯데 사전원가 시트의 VAT 기준:
            # 현장경비는 차량유지비, 임차료는 첫 행, 외주비는 케이엔디이를 제외한다.
            site_vat_base = sum(
                self._to_f(w['qty'].get()) * self._to_f(w['price'].get())
                for w in self.entries['site_expense']
                if '차량유지비' not in str(w['cat'].get())
            )
            rental_vat_base = sum(
                self._to_f(w['qty'].get()) * self._to_f(w['period'].get()) * self._to_f(w['price'].get())
                for i, w in enumerate(self.entries['rental']) if i != 0
            )
            outsource_vat_base = sum(
                self._to_f(w['count'].get()) * self._to_f(w['price'].get())
                for w in self.entries['outsource']
                if str(w['cat'].get()).strip() != '케이엔디이'
            )
            exp_vat = (site_vat_base + rental_vat_base + outsource_vat_base) * 0.1
        else:
            exp_vat = (t1 + t2 + t3) * 0.1
        self.lbl_exp_total.config(text=f"₩ {exp_total:,.0f}")
        self.lbl_exp_vat.config(text=f"{exp_vat:,.0f}")
        
        # 6. Sales Cost (Labor + Material + Expense + Outsource)
        mat_total = self.get_material_total() if self.get_material_total else 0.0
        direct_cost = labor_total + mat_total + exp_total + t3
        self.lbl_sales_cost_total.config(text=f"₩ {direct_cost:,.0f}")
        
        # 7. Indirect Cost (14%)
        self.lbl_indirect_base.config(text=f"₩ {direct_cost:,.0f}")
        indirect_cost = (direct_cost - t3) * 0.14 if lotte_rules else direct_cost * 0.14
        self.lbl_indirect_total.config(text=f"₩ {indirect_cost:,.0f}")
        
        grand_total_cost = direct_cost + indirect_cost
        self.lbl_grand_total_cost.config(text=f"₩ {grand_total_cost:,.0f}")
        
        # 8. Profit
        revenue = self.get_revenue() if self.get_revenue else 0.0
        op_profit = revenue - grand_total_cost
        margin = (op_profit / revenue * 100) if revenue > 0 else 0.0
        
        self.lbl_prof_revenue.config(text=f"₩ {revenue:,.0f}")
        self.lbl_prof_total_cost.config(text=f"₩ {grand_total_cost:,.0f}")
        self.lbl_prof_op_profit.config(text=f"₩ {op_profit:,.0f}")
        self.lbl_prof_margin.config(text=f"{margin:.2f}%")
        
        # Update main form "Expense" and "Outsource" fields
        if self.on_change_callback:
            # We pass (Expense, Outsource, TotalProfit) or something?
            # Expense excludes outsource; outsource is passed separately as t3.
            self.on_change_callback(exp_total, t3, op_profit)

    def _to_f(self, val):
        try:
            return float(str(val).replace(',', '') or 0)
        except:
            return 0.0

    def get_total_cost(self):
        """[FINAL_FIX] Robustly get total expense cost (including depreciation)"""
        try:
            raw_text = self.lbl_exp_total.cget('text')
            val = "".join(c for c in raw_text if c.isdigit() or c == '.')
            return float(val or 0)
        except:
            return 0.0

    def get_data(self):
        data = {
            'site_expense': [{k: v.get() if hasattr(v, 'get') else v.cget('text') for k, v in row.items()} for row in self.entries['site_expense']],
            'rental': [{k: v.get() if hasattr(v, 'get') else v.cget('text') for k, v in row.items()} for row in self.entries['rental']],
            'outsource': [{k: v.get() if hasattr(v, 'get') else v.cget('text') for k, v in row.items()} for row in self.entries['outsource']],
            'depreciation': [{k: v.get() if hasattr(v, 'get') else v.cget('text') for k, v in row.items()} for row in self.entries['depreciation']]
        }
        return data

    def get_total_cost(self):
        """Retrieve the total site expense cost (items 1~5) as a float"""
        try:
            val = self.lbl_exp_total.cget('text').replace('₩', '').replace(',', '').replace(' ', '').strip()
            return float(val or 0)
        except:
            return 0.0

    def set_data(self, data):
        if not data or not isinstance(data, dict):
            self.reset()
            return
            
        def fill(entry_list, data_list):
            for i, d in enumerate(data_list):
                if i < len(entry_list):
                    for k, v in d.items():
                        if k in entry_list[i] and hasattr(entry_list[i][k], 'delete'):
                            entry_list[i][k].delete(0, tk.END); entry_list[i][k].insert(0, str(v))

        def fill_depreciation(data_list):
            """Restore depreciation rows by equipment name, not legacy row index."""
            rows_by_name = {
                str(row['item'].get()).strip().upper(): row
                for row in self.entries['depreciation']
            }
            restored_names = set()
            for saved in data_list:
                name = str(saved.get('item', '')).strip().upper()
                row = rows_by_name.get(name)
                if row is None or name in restored_names:
                    continue
                restored_names.add(name)
                for key, value in saved.items():
                    widget = row.get(key)
                    if widget is not None and hasattr(widget, 'delete'):
                        # Legacy saved rows may contain an empty depreciation rate.
                        # Keep the current equipment default instead of erasing it.
                        normalized = str(value).strip().replace(',', '')
                        if key in ('qty', 'rate') and normalized in ('', '0', '0.0', 'None', 'nan'):
                            continue
                        if key == 'days' and normalized in ('', 'None', 'nan'):
                            continue
                        widget.delete(0, tk.END)
                        widget.insert(0, str(value))
        
        site_expenses = data.get('site_expense', [])
        if self.budget_mode in ('planned', 'actual'):
            for exp in site_expenses:
                name = str(exp.get('cat', ''))
                if self.budget_mode == 'planned':
                    if '차량' in name:
                        exp['qty'], exp['unit'], exp['price'] = '12', '개월', '150,000'
                    elif '소모' in name:
                        exp['qty'], exp['unit'], exp['price'] = '12', '개월', '15,000'
                    elif '복리' in name or '후생' in name:
                        exp['qty'], exp['unit'], exp['price'] = '12', '개월', '50,000'
                    elif 'Se' in name or '175' in name:
                        exp['qty'], exp['unit'], exp['price'] = '1', 'EA', '10,000,000'
                elif self.budget_mode == 'actual':
                    # 사후원가는 현장별 실제 투입일수 × 일일 단가로 계산한다.
                    if '차량' in name:
                        exp['unit'], exp['price'] = '일', '5,000'
                    elif '소모' in name:
                        exp['unit'], exp['price'] = '일', '500'
                    elif '복리' in name or '후생' in name:
                        exp['unit'], exp['price'] = '일', '1,667'
                    elif 'Se' in name or '175' in name:
                        exp['unit'], exp['price'] = '일', '47,619'

        fill(self.entries['site_expense'], site_expenses)
        fill(self.entries['rental'], data.get('rental', []))
        fill(self.entries['outsource'], data.get('outsource', []))
        fill_depreciation(data.get('depreciation', []))
        
        self.calculate_all()

    def reset(self):
        def clear(entry_list):
            for row in entry_list:
                for k, v in row.items():
                    if hasattr(v, 'delete'): v.delete(0, tk.END)
        
        clear(self.entries['site_expense'])
        clear(self.entries['rental'])
        clear(self.entries['outsource'])
        clear(self.entries['depreciation'])
        self.calculate_all()


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
        
        # Unbind when closed to avoid errors
        def on_close():
            self.destroy()
        
        self.protocol("WM_DELETE_WINDOW", on_close)
            
        # Create checkboxes
        for item in columns:
            if isinstance(item, (tuple, list)) and len(item) >= 2:
                col = item[0]
                provided_display_text = item[1]
            else:
                col = item
                provided_display_text = None
            var = tk.BooleanVar(value=True)
            self.vars[col] = var
            # 사용자에게는 내부 컬럼명 대신 표시용 한글명을 보여줌
            display_map = {
                '검사량': '수량',
                'Date': '날짜',
                'Site': '현장',
                'EntryTime': '입력시간',
                'MaterialID': '자재ID',
                'FilmCount': '수량'
            }
            display_text = provided_display_text or display_map.get(col, col)
            cb = ttk.Checkbutton(scrollable_frame, text=display_text, variable=var)
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


