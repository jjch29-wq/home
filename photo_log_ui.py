import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import xlsxwriter
import os
import glob
from PIL import Image, ImageOps
import datetime
import math
import threading
import json

class PhotoLogApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Photo Log Generator v33:00")
        self.root.geometry("600x650")
        self.root.configure(background="#f3f4f6")
        
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Header Mapping
        self.header_map = {
            "PAUT": "REPORT OF PHASED ARRAY UT EXAMINATION (위 상 배 열 초 음 파 탐 상 검 사 보 고 서)",
            "RT": "REPORT OF RADIOGRAPHIC EXAMINATION (방 사 선 투 과 검 사 보 고 서)",
            "PT": "REPORT OF LIQUID PENETRANT EXAMINATION (침 투 탐 상 검 사 보 고 서)",
            "MT": "REPORT OF MAGNETIC PARTICLE EXAMINATION (자 분 탐 상 검 사 보 고 서)",
            "PMI": "REPORT OF POSITIVE MATERIAL IDENTIFICATION (재 질 성 분 분 석 검 사 보 고 서)",
            "UT": "REPORT OF ULTRASONIC EXAMINATION (초 음 파 탐 상 검 사 보 고 서)",
            "NDT": "REPORT OF NON-DESTRUCTIVE EXAMINATION (비 파 괴 검 사 보 고 서)",
            "기타 (직접 입력)": ""
        }
        
        # UI Variables
        self.orderer = tk.StringVar(value="서울에너지공사")
        self.report_no = tk.StringVar(value="SIT/GI-SE-PAUT-TNTFJPWJ001")
        self.inspect_date = tk.StringVar(value=datetime.datetime.now().strftime("%Y-%m-%d"))
        self.inspect_type = tk.StringVar(value="PAUT")
        self.report_title = tk.StringVar(value=self.header_map["PAUT"])
        
        self.cols_per_row = tk.StringVar(value="2")
        self.keep_aspect = tk.BooleanVar(value=True)
        
        self.logo_path = tk.StringVar(value=os.path.join(os.getcwd(), "logo.png"))
        self.output_name = tk.StringVar(value="NDT_Photo_Log_Final.xlsx")
        self.logo_width_var = tk.StringVar(value="80") # Custom Logo Width
        self.logo_x_var = tk.StringVar(value="2")      # Logo X Offset
        self.logo_y_var = tk.StringVar(value="0")      # Logo Y Offset
        self.selected_files = [] # Store individual file paths
        
        # New: Customizable Cell Dimensions
        self.cell_width_var = tk.StringVar(value="53.0")
        self.cell_height_var = tk.StringVar(value="178.0")
        
        # New: Page Layout Settings
        self.margin_top_var = tk.StringVar(value="0.5")
        self.margin_bottom_var = tk.StringVar(value="0.5")
        self.margin_left_var = tk.StringVar(value="0.4")
        self.margin_right_var = tk.StringVar(value="0.4")
        self.print_scale_var = tk.StringVar(value="100")
        self.desc_height_var = tk.StringVar(value="20.0")
        self.photo_align_var = tk.StringVar(value="중앙 정렬")
        self.fit_width_var = tk.BooleanVar(value=True)
        
        self.settings_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "photo_log_settings.json")
        
        self.create_widgets()
        self.load_settings() # Restore previous session data
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
    def create_widgets(self):
        # Main Container
        main_frame = tk.Frame(self.root, background="#f3f4f6", padx=20, pady=20)
        main_frame.pack(fill="both", expand=True)
        
        # Title
        tk.Label(main_frame, text="📷 Photo Log Generator", font=("Malgun Gothic", 18, "bold"), 
                 background="#f3f4f6", foreground="#1e3a8a").pack(side="top", pady=(0, 20))
        
        # 1. Buttons Frame (Bottom)
        btn_frame = tk.Frame(main_frame, background="#f3f4f6")
        btn_frame.pack(side="bottom", fill="x", pady=(10, 0))
        
        ttk.Button(btn_frame, text="리포트 생성 시작", command=self.start_generation).pack(side="right", padx=5)
        ttk.Button(btn_frame, text="종료", command=self.root.quit).pack(side="right", padx=5)

        # 2. Progress Bar
        self.progress = ttk.Progressbar(main_frame, orient="horizontal", mode="determinate")
        self.progress.pack(side="bottom", fill="x", pady=(10, 0))

        # 3. Input Section (Top)
        input_frame = ttk.LabelFrame(main_frame, text=" 리포트 정보 (Report Info) ", padding=15)
        input_frame.pack(side="top", fill="x", pady=5)
        
        # Inspection Type & Title
        tk.Label(input_frame, text="검사 항목:", font=("Malgun Gothic", 10), anchor="w").grid(row=0, column=0, sticky="ew", pady=2)
        type_combo = ttk.Combobox(input_frame, textvariable=self.inspect_type, values=list(self.header_map.keys()), state="readonly", font=("Malgun Gothic", 10))
        type_combo.grid(row=0, column=1, sticky="ew", padx=(10, 0), pady=2)
        type_combo.bind("<<ComboboxSelected>>", self.on_type_change)
        
        tk.Label(input_frame, text="리포트 제목:", font=("Malgun Gothic", 10), anchor="w").grid(row=1, column=0, sticky="ew", pady=2)
        ttk.Entry(input_frame, textvariable=self.report_title, font=("Malgun Gothic", 9)).grid(row=1, column=1, sticky="ew", padx=(10, 0), pady=2)

        self._add_input_row(input_frame, "발주처", self.orderer, 2)
        self._add_input_row(input_frame, "REPORT NO:", self.report_no, 3)
        self._add_input_row(input_frame, "검사일자", self.inspect_date, 4)
        input_frame.columnconfigure(1, weight=1)
        
        # 4. Image Options Section
        opt_frame = ttk.LabelFrame(main_frame, text=" 사진 레이아웃 설정 (Image Options) ", padding=15)
        opt_frame.pack(side="top", fill="x", pady=5)
        
        tk.Label(opt_frame, text="한 줄당 사진 개수:", font=("Malgun Gothic", 10)).grid(row=0, column=0, sticky="w", pady=2)
        layout_combo = ttk.Combobox(opt_frame, textvariable=self.cols_per_row, values=["1", "2", "3"], state="readonly", width=5)
        layout_combo.grid(row=0, column=1, sticky="w", padx=10, pady=2)
        
        ttk.Checkbutton(opt_frame, text="가로세로 비율 유지 및 중앙 정렬", variable=self.keep_aspect).grid(row=0, column=2, padx=20, sticky="w")

        # New: Cell Dimension Inputs
        tk.Label(opt_frame, text="사진 칸 너비 (엑셀):", font=("Malgun Gothic", 10)).grid(row=1, column=0, sticky="w", pady=2)
        ttk.Entry(opt_frame, textvariable=self.cell_width_var, width=10).grid(row=1, column=1, sticky="w", padx=10, pady=2)
        
        tk.Label(opt_frame, text="사진 칸 높이 (포인트):", font=("Malgun Gothic", 10)).grid(row=1, column=2, sticky="w", pady=2)
        ttk.Entry(opt_frame, textvariable=self.cell_height_var, width=10).grid(row=1, column=3, sticky="w", padx=10, pady=2)

        # New: Margin & Scale Section
        tk.Label(opt_frame, text="여백(Top/Bottom):", font=("Malgun Gothic", 10)).grid(row=2, column=0, sticky="w", pady=2)
        m_tb_frame = tk.Frame(opt_frame, background="#f3f4f6")
        m_tb_frame.grid(row=2, column=1, sticky="w", padx=10)
        ttk.Entry(m_tb_frame, textvariable=self.margin_top_var, width=5).pack(side="left")
        ttk.Entry(m_tb_frame, textvariable=self.margin_bottom_var, width=5).pack(side="left", padx=2)

        tk.Label(opt_frame, text="여백(Left/Right):", font=("Malgun Gothic", 10)).grid(row=2, column=2, sticky="w", pady=2)
        m_lr_frame = tk.Frame(opt_frame, background="#f3f4f6")
        m_lr_frame.grid(row=2, column=3, sticky="w", padx=10)
        ttk.Entry(m_lr_frame, textvariable=self.margin_left_var, width=5).pack(side="left")
        ttk.Entry(m_lr_frame, textvariable=self.margin_right_var, width=5).pack(side="left", padx=2)

        tk.Label(opt_frame, text="인쇄 배율 (%):", font=("Malgun Gothic", 10)).grid(row=3, column=0, sticky="w", pady=2)
        ttk.Entry(opt_frame, textvariable=self.print_scale_var, width=10).grid(row=3, column=1, sticky="w", padx=10, pady=2)

        tk.Label(opt_frame, text="설명 줄 높이 (포인트):", font=("Malgun Gothic", 10)).grid(row=3, column=2, sticky="w", pady=2)
        ttk.Entry(opt_frame, textvariable=self.desc_height_var, width=10).grid(row=3, column=3, sticky="w", padx=10, pady=2)

        tk.Label(opt_frame, text="로고 너비 (픽셀):", font=("Malgun Gothic", 10)).grid(row=4, column=0, sticky="w", pady=2)
        ttk.Entry(opt_frame, textvariable=self.logo_width_var, width=10).grid(row=4, column=1, sticky="w", padx=10, pady=2)

        tk.Label(opt_frame, text="로고 X/Y 오프셋", font=("Malgun Gothic", 10)).grid(row=4, column=2, sticky="w", pady=2)
        logo_xy_frame = tk.Frame(opt_frame, background="#f3f4f6")
        logo_xy_frame.grid(row=4, column=3, sticky="w", padx=10)
        ttk.Entry(logo_xy_frame, textvariable=self.logo_x_var, width=5).pack(side="left")
        ttk.Entry(logo_xy_frame, textvariable=self.logo_y_var, width=5).pack(side="left", padx=2)

        # New: Photo Alignment & Fit Width
        tk.Label(opt_frame, text="사진 정렬:", font=("Malgun Gothic", 10)).grid(row=5, column=0, sticky="w", pady=2)
        align_combo = ttk.Combobox(opt_frame, textvariable=self.photo_align_var, values=["중앙 정렬", "좌측 정렬"], state="readonly", width=10)
        align_combo.grid(row=5, column=1, sticky="w", padx=10, pady=2)

        ttk.Checkbutton(opt_frame, text="가로 폭 맞춤 (Fit to Width)", variable=self.fit_width_var).grid(row=5, column=2, columnspan=2, padx=20, sticky="w")

        # 5. Path & File Management Section (Replaced Folder Entry with Listbox)
        path_frame = ttk.LabelFrame(main_frame, text=" 사진 파일 관리 (Selection Management) ", padding=15)
        path_frame.pack(side="top", fill="both", expand=True, pady=5)
        
        # File Listbox
        list_grid_frame = tk.Frame(path_frame, background="#ffffff")
        list_grid_frame.pack(side="top", fill="both", expand=True)
        
        self.file_listbox = tk.Listbox(list_grid_frame, height=8, font=("Consolas", 9), selectmode="extended", borderwidth=0)
        self.file_listbox.pack(side="left", fill="both", expand=True)
        
        list_scroll = ttk.Scrollbar(list_grid_frame, orient="vertical", command=self.file_listbox.yview)
        list_scroll.pack(side="right", fill="y")
        self.file_listbox.config(yscrollcommand=list_scroll.set)
        
        # Management Tool Buttons
        m_btn_frame = tk.Frame(path_frame)
        m_btn_frame.pack(side="top", fill="x", pady=5)
        ttk.Button(m_btn_frame, text="파일 개별 추가", command=self.add_files).pack(side="left", padx=2)
        ttk.Button(m_btn_frame, text="폴더 전체 추가", command=self.add_folder).pack(side="left", padx=2)
        ttk.Button(m_btn_frame, text="선택 항목 제거", command=self.remove_selected).pack(side="right", padx=2)
        ttk.Button(m_btn_frame, text="전체 비우기", command=self.clear_all).pack(side="right", padx=2)

        # Logo and Output Name Section
        misc_frame = tk.Frame(path_frame)
        misc_frame.pack(side="top", fill="x", pady=5)
        self._add_path_row(misc_frame, "로고 파일:", self.logo_path, 0)
        self._add_input_row(misc_frame, "파일명", self.output_name, 1)
        
        # 6. Log Section
        log_frame = ttk.LabelFrame(main_frame, text=" 진행 로그 (Log) ", padding=10)
        log_frame.pack(side="top", fill="both", expand=True, pady=10)
        
        self.log_text = tk.Text(log_frame, height=5, font=("Consolas", 9), state="disabled", background="#ffffff")
        self.log_text.pack(side="left", fill="both", expand=True)
        
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

    def _add_input_row(self, parent, label, var, row):
        tk.Label(parent, text=label, font=("Malgun Gothic", 10), anchor="w").grid(row=row, column=0, sticky="ew", pady=2)
        ttk.Entry(parent, textvariable=var, font=("Malgun Gothic", 10)).grid(row=row, column=1, sticky="ew", padx=(10, 0), pady=2)
        parent.columnconfigure(1, weight=1)

    def _add_path_row(self, parent, label, var, row, is_dir=False):
        tk.Label(parent, text=label, font=("Malgun Gothic", 10), anchor="w").grid(row=row, column=0, sticky="ew", pady=2)
        ttk.Entry(parent, textvariable=var, font=("Malgun Gothic", 10)).grid(row=row, column=1, sticky="ew", padx=(10, 0), pady=2)
        cmd = lambda: self._browse(var, is_dir)
        ttk.Button(parent, text="찾기", width=5, command=cmd).grid(row=row, column=2, padx=5, pady=2)
        parent.columnconfigure(1, weight=1)

    def on_type_change(self, event):
        new_type = self.inspect_type.get()
        if new_type in self.header_map:
            self.report_title.set(self.header_map[new_type])

    def add_files(self):
        files = filedialog.askopenfilenames(filetypes=[("Image files", "*.png;*.jpg;*.jpeg;*.bmp")])
        if files:
            logo_f = os.path.normpath(self.logo_path.get())
            for f in files:
                f_norm = os.path.normpath(f)
                if f_norm == logo_f: continue # Skip logo file
                if f_norm not in self.selected_files:
                    self.selected_files.append(f_norm)
                    self.file_listbox.insert(tk.END, f_norm)
            self.log(f"{len(files)}개 파일 추가 시도")

    def add_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            added_count = 0
            logo_f = os.path.normpath(self.logo_path.get())
            for ext in ('*.png', '*.jpg', '*.jpeg', '*.bmp'):
                for f in glob.glob(os.path.join(folder, ext)):
                    f_norm = os.path.normpath(f)
                    if f_norm == logo_f: continue # Skip logo file
                    if f_norm not in self.selected_files:
                        self.selected_files.append(f_norm)
                        self.file_listbox.insert(tk.END, f_norm)
                        added_count += 1
            self.log(f"폴더에서 {added_count}개 파일 추가 완료 ({folder})")

    def remove_selected(self):
        idxs = list(self.file_listbox.curselection())
        for i in reversed(idxs):
            path = self.file_listbox.get(i)
            if path in self.selected_files:
                self.selected_files.remove(path)
            self.file_listbox.delete(i)

    def clear_all(self):
        self.selected_files.clear()
        self.file_listbox.delete(0, tk.END)
        self.log("파일 리스트가 초기화되었습니다.")

    def _browse(self, var, is_dir):
        if is_dir:
            path = filedialog.askdirectory()
        else:
            path = filedialog.askopenfilename(filetypes=[("Image files", "*.png;*.jpg;*.jpeg;*.bmp")])
        if path:
            var.set(os.path.normpath(path))

    def _get_unique_path(self, path):
        """Returns a non-existing file path by appending _1, _2, etc."""
        if not os.path.exists(path):
            return path
        base, ext = os.path.splitext(path)
        counter = 1
        while True:
            new_path = f"{base}_{counter}{ext}"
            if not os.path.exists(new_path):
                return new_path
            counter += 1

    def on_closing(self):
        self.save_settings()
        self.root.destroy()

    def save_settings(self):
        try:
            data = {
                "orderer": self.orderer.get(),
                "report_no": self.report_no.get(),
                "inspect_date": self.inspect_date.get(),
                "inspect_type": self.inspect_type.get(),
                "report_title": self.report_title.get(),
                "cols_per_row": self.cols_per_row.get(),
                "keep_aspect": self.keep_aspect.get(),
                "logo_path": self.logo_path.get(),
                "output_name": self.output_name.get(),
                "cell_width": self.cell_width_var.get(),
                "cell_height": self.cell_height_var.get(),
                "logo_width": self.logo_width_var.get(),
                "logo_x": self.logo_x_var.get(),
                "logo_y": self.logo_y_var.get(),
                "margin_top": self.margin_top_var.get(),
                "margin_bottom": self.margin_bottom_var.get(),
                "margin_left": self.margin_left_var.get(),
                "margin_right": self.margin_right_var.get(),
                "print_scale": self.print_scale_var.get(),
                "desc_height": self.desc_height_var.get(),
                "photo_align": self.photo_align_var.get(),
                "fit_width": self.fit_width_var.get(),
                "selected_files": self.selected_files
            }
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"Error saving settings: {e}")

    def load_settings(self):
        if not os.path.exists(self.settings_file):
            return
        try:
            with open(self.settings_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if "orderer" in data: self.orderer.set(data["orderer"])
                if "report_no" in data: self.report_no.set(data["report_no"])
                if "inspect_date" in data: self.inspect_date.set(data["inspect_date"])
                if "inspect_type" in data: self.inspect_type.set(data["inspect_type"])
                if "report_title" in data: self.report_title.set(data["report_title"])
                if "cols_per_row" in data: self.cols_per_row.set(data["cols_per_row"])
                if "keep_aspect" in data: self.keep_aspect.set(data["keep_aspect"])
                if "logo_path" in data: self.logo_path.set(data["logo_path"])
                if "output_name" in data: self.output_name.set(data["output_name"])
                if "cell_width" in data:
                    val = data["cell_width"]
                    if val in ["38.0", "50.0", "55.0", "60.0", "48.0"]:
                        self.cell_width_var.set("53.0")
                    else:
                        self.cell_width_var.set(val)
                if "cell_height" in data: 175.0
                if "logo_width" in data: self.logo_width_var.set(data["logo_width"])
                if "logo_x" in data: self.logo_x_var.set(data["logo_x"])
                if "logo_y" in data: self.logo_y_var.set(data["logo_y"])
                if "margin_top" in data: self.margin_top_var.set(data["margin_top"])
                if "margin_bottom" in data: self.margin_bottom_var.set(data["margin_bottom"])
                if "margin_left" in data: self.margin_left_var.set(data["margin_left"])
                if "margin_right" in data: self.margin_right_var.set(data["margin_right"])
                if "print_scale" in data: self.print_scale_var.set(data["print_scale"])
                if "desc_height" in data: self.desc_height_var.set(data["desc_height"])
                if "photo_align" in data: self.photo_align_var.set(data["photo_align"])
                if "fit_width" in data: self.fit_width_var.set(data["fit_width"])
                
                if "selected_files" in data:
                    self.selected_files = data["selected_files"]
                    self.file_listbox.delete(0, tk.END)
                    for f_path in self.selected_files:
                        self.file_listbox.insert(tk.END, f_path)
        except Exception as e:
            print(f"Error loading settings: {e}")

    def log(self, message):
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")
        self.root.update_idletasks()

    def start_generation(self):
        threading.Thread(target=self.generate_report, daemon=True).start()

    def generate_report(self):
        try:
            self.save_settings() # Auto-save before starting
            self.progress["value"] = 0
            self.log("작업을 시작합니다..")
            
            if not self.selected_files:
                self.log("오류: 선택된 파일이 없습니다.")
                messagebox.showwarning("경고", "리포트에 포함할 이미지를 먼저 선택해주세요.")
                return

            logo_f = os.path.normpath(self.logo_path.get())
            image_files = sorted([f for f in self.selected_files if f != logo_f])
            
            if not image_files:
                self.log("오류: 선택된 이미지가 없습니다 (또는 로고만 선택됨).")
                return
            # Use the directory of the first image as default output location
            img_folder = os.path.dirname(image_files[0])
            
            output_name_val = self.output_name.get()
            if not output_name_val.endswith(".xlsx"):
                output_name_val += ".xlsx"
                
            output_path = os.path.join(os.path.dirname(img_folder), output_name_val)
            output_path = self._get_unique_path(output_path) # Auto-versioning support
            
            self.log(f"파일 생성 중: {os.path.basename(output_path)}")
            workbook = xlsxwriter.Workbook(output_path)
            worksheet = workbook.add_worksheet()

            # Page Setup
            worksheet.set_paper(9) # A4
            worksheet.set_portrait()
            worksheet.center_horizontally()
            
            try:
                m_t = max(0.5, float(self.margin_top_var.get()))
                m_b = max(0.35, float(self.margin_bottom_var.get()))
                m_l = max(0.25, float(self.margin_left_var.get()))
                m_r = max(0.25, float(self.margin_right_var.get()))
                worksheet.set_margins(left=m_l, right=m_r, top=m_t, bottom=m_b)
            except:
                worksheet.set_margins(left=0.4, right=0.4, top=0.5, bottom=0.5)
            
            worksheet.set_footer('&C&P / &N')
            worksheet.repeat_rows(0, 4)  # 1~5행(제목+회사정보+PHOTO LOG) 2페이지부터 반복
            
            # Layout
            num_cols = int(self.cols_per_row.get())
            photos_per_page = 4 if num_cols == 1 else (8 if num_cols == 2 else 12)
            total_pages = math.ceil(len(image_files) / photos_per_page)
            worksheet.fit_to_pages(1, total_pages)

            # Formats
            title_format = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'shrink': True})
            company_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_size': 9, 'text_wrap': True})
            center_border = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_size': 10})
            bold_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
            desc_format = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1, 'font_size': 10, 'shrink': True, 'text_wrap': False, 'indent': 1})

            # Layout Calculation
            num_cols = int(self.cols_per_row.get())
            # Standard Excel pixel mapping (Optimized for 96 DPI / Calibri 11)
            COL_UNIT_TO_PX = 7.03 # Golden ratio for full A4 coverage
            ROW_PT_TO_PX = 1.33333
            MARGIN_1MM_PX = 0

            # Fixed 6-column Grid System for consistent layout
            GRID_COLS = 6
            # Total width derived from user-defined unit (default 38)
            unit_per_grid = (float(self.cell_width_var.get()) * 2) / GRID_COLS
            worksheet.set_column(0, GRID_COLS - 1, unit_per_grid)

            # User-controlled cell height from UI
            CELL_ROW_HEIGHT = float(self.cell_height_var.get())
            
            # Precise Column spans and width calculations based on grid units
            if num_cols == 1:
                photo_col_spans = [(0, GRID_COLS - 1)]
                CELL_WIDTH_PX = (unit_per_grid * 6 * 7.0) + (6 * 5)
            elif num_cols == 2:
                photo_col_spans = [(0, 2), (3, 5)]
                CELL_WIDTH_PX = (unit_per_grid * 3 * 7.0) + (3 * 5)
            else: # 3 Columns
                photo_col_spans = [(0, 1), (2, 3), (4, 5)]
                CELL_WIDTH_PX = (unit_per_grid * 2 * 7.0) + (2 * 5)

            last_col_idx = GRID_COLS - 1
            last_col_letter = chr(65 + last_col_idx)
            header_range = f"A1:{last_col_letter}1"
            worksheet.set_row(0, 30) # Explicitly set title row height
            worksheet.merge_range(header_range, self.report_title.get(), title_format)
            
            # Company Info (Left side, takes 3 grid units - symmetry)
            company_info_text = "서   울   檢   査   株   式   會   社\nSEOUL INSPECTION & TESTING Co., Ltd.\nTEL : (02) 552-1112   FAX : (02) 2058-0720"
            worksheet.merge_range(1, 0, 3, 2, company_info_text, company_format)

            # Logo Placement within Company Info range
            logo_p = self.logo_path.get()
            if os.path.exists(logo_p):
                try:
                    total_header_h = 45 # 15 * 3 rows
                    for r in range(1, 4): worksheet.set_row(r, 15) 
                    with Image.open(logo_p) as img:
                        w, h = img.size
                        # Dynamic logo width based on user input (default 80)
                        try:
                            max_w_logo = float(self.logo_width_var.get())
                        except:
                            max_w_logo = 80.0
                        
                        try:
                            manual_x = float(self.logo_x_var.get())
                            manual_y = float(self.logo_y_var.get())
                        except:
                            manual_x = 2.0
                            manual_y = 0.0

                        scale = min(max_w_logo/w, 42/h) * 0.95 
                        logo_h = h * scale
                        # Auto-centered Y plus manual adjustment
                        y_offset = (total_header_h - logo_h) / 2 + manual_y
                        
                        worksheet.insert_image('A2', logo_p, {
                            'x_scale': scale, 
                            'y_scale': scale, 
                            'x_offset': manual_x, 
                            'y_offset': y_offset, 
                            'object_position': 1
                        })
                except: pass

            # Info (Right side, takes 3 grid units - symmetry)
            # Merge info across columns 3, 4, 5
            worksheet.merge_range(1, 3, 1, 5, f"발주처: {self.orderer.get()}", center_border)
            worksheet.merge_range(2, 3, 2, 5, f"REPORT NO: {self.report_no.get()}", center_border)
            worksheet.merge_range(3, 3, 3, 5, f"검사일자: {self.inspect_date.get()}", center_border)
            
            # Photo Log Header (Row 5 - Row idx 4)
            worksheet.set_row(4, 25)
            worksheet.merge_range(f"A5:{last_col_letter}5", "PHOTO LOG (사진 대장)", bold_format)

            # Image Insertion Logic
            row = 5
            col_ptr = 0
            page_breaks = []
            photos_per_page = 4 if num_cols == 1 else (8 if num_cols == 2 else 12)
            try:
                DESC_ROW_HEIGHT = float(self.desc_height_var.get())
            except:
                DESC_ROW_HEIGHT = 22.0

            CELL_HEIGHT_PX = (CELL_ROW_HEIGHT * ROW_PT_TO_PX) - 2
            total = len(image_files)
            for i, img_path in enumerate(image_files):
                worksheet.set_row(row, CELL_ROW_HEIGHT)
                try:
                    with Image.open(img_path) as img:
                        # Auto Rotate based on EXIF
                        img = ImageOps.exif_transpose(img)
                        img_w, img_h = img.size
                        # Determine insertion col and merge based on grid
                        c_start, c_end = photo_col_spans[col_ptr]
                        if c_start != c_end:
                            worksheet.merge_range(row, c_start, row, c_end, "", center_border)
                        
                        target_w_px = CELL_WIDTH_PX
                        
                        if self.fit_width_var.get():
                            # Fit to width, but cap at height to avoid covering description
                            scale = min(target_w_px / img_w, CELL_HEIGHT_PX / img_h)
                        elif self.keep_aspect.get():
                            # Smart Fitting (Fit to box)
                            scale = min(target_w_px / img_w, CELL_HEIGHT_PX / img_h)
                        else:
                            # Force Stretch to full box
                            scale = 1.0 # Placeholder, will use explicit x/y scale below

                        if not self.keep_aspect.get() and not self.fit_width_var.get():
                            # Force Stretch
                            x_scale = target_w_px / img_w
                            y_scale = CELL_HEIGHT_PX / img_h
                            x_offset = MARGIN_1MM_PX
                            y_offset = MARGIN_1MM_PX
                        else:
                            x_scale = y_scale = scale
                            # Horizontal Alignment
                            if self.photo_align_var.get() == "좌측 정렬":
                                x_offset = MARGIN_1MM_PX
                            else: # 중앙 정렬
                                x_offset = (target_w_px - (img_w * scale)) / 2 + MARGIN_1MM_PX
                            
                            # Vertical Centering (always maintained for non-stretched)
                            y_offset = (CELL_HEIGHT_PX - (img_h * scale)) / 2 + MARGIN_1MM_PX

                        worksheet.insert_image(row, c_start, img_path, {
                            'x_scale': x_scale, 'y_scale': y_scale, 
                            'x_offset': x_offset, 'y_offset': y_offset, 
                            'object_position': 1,
                            'image_data': None
                        })
                except Exception as e:
                    self.log(f"이미지 오류({os.path.basename(img_path)}): {e}")

                name_only = os.path.splitext(os.path.basename(img_path))[0]
                worksheet.set_row(row + 1, DESC_ROW_HEIGHT)
                
                c_start, c_end = photo_col_spans[col_ptr]
                if c_start != c_end:
                    worksheet.merge_range(row+1, c_start, row+1, c_end, f"설명: {name_only}", desc_format)
                else:
                    worksheet.write(row + 1, c_start, f"설명: {name_only}", desc_format)
                
                if num_cols == 1:
                    row += 2
                else:
                    col_ptr += 1
                    if col_ptr >= num_cols:
                        col_ptr = 0
                        row += 2
                
                # Calculate if a page break is needed
                # For 1 col: every 1 photo row (2 excel rows) * 4 photos
                # For 2/3 cols: after every 'photos_per_page' images are processed
                if (i + 1) % photos_per_page == 0 and (i + 1) < total:
                    # The current 'row' is already pointing to the next available row index
                    page_breaks.append(row)

                self.progress["value"] = ((i + 1) / total) * 100
                self.log(f"처리 중.. ({i+1}/{total})")

            # Fill remaining empty cells in last row if incomplete
            if num_cols > 1 and col_ptr > 0:
                worksheet.set_row(row, CELL_ROW_HEIGHT)
                worksheet.set_row(row + 1, DESC_ROW_HEIGHT)
                while col_ptr < num_cols:
                    c_start, c_end = photo_col_spans[col_ptr]
                    if c_start != c_end:
                        worksheet.merge_range(row, c_start, row, c_end, "", center_border)
                        worksheet.merge_range(row+1, c_start, row+1, c_end, "", desc_format)
                    else:
                        worksheet.write(row, c_start, "", center_border)
                        worksheet.write(row+1, c_start, "", desc_format)
                    col_ptr += 1

            if page_breaks:
                worksheet.set_h_pagebreaks(page_breaks)

            workbook.close()
            self.log(f"성공! 파일 저장 완료: {output_path}")
            messagebox.showinfo("성공", f"리포트가 생성되었습니다.\n{output_path}")

        except PermissionError:
            self.log("오류: 파일을 저장할 수 없습니다. 프로그램에서 닫아주세요.")
            messagebox.showerror("오류", "파일이 프로그램에서 열려 있습니다. 닫고 다시 시도해주세요.")
        except Exception as e:
            self.log(f"치명적 오류: {e}")
            messagebox.showerror("오류", f"작업 중 오류가 발생했습니다:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = PhotoLogApp(root)
    root.mainloop()

