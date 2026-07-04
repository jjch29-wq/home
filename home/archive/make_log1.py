import os
import json
import io
import glob
import sys
import math
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import xlsxwriter
from PIL import Image

class PhotoLogApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PAUT 사진 대장 엑셀 자동화 프로그램")
        self.root.geometry("550x520")
        
        self.base_path = os.path.dirname(os.path.abspath(__file__))
        self.config_path = os.path.join(self.base_path, 'make_log1_config.json')
        self.load_config()
        self.create_widgets()
        
    def load_config(self):
        # 기본값
        img_dir = os.path.join(self.base_path, 'images')
        out_dir = os.path.join(self.base_path, "NDT_Photo_Log_Final.xlsx")
        client = "서울에너지공사"
        report_no = "SIT/GI-SE-PAUT-TNTFJPWJ001"
        date = "2024년 10월 24일"
        
        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    img_dir   = config.get('img_folder', img_dir)
                    out_dir   = config.get('out_file', out_dir)
                    client    = config.get('client', client)
                    report_no = config.get('report_no', report_no)
                    date      = config.get('date', date)
                    img_w     = config.get('img_w', 280)
                    img_h     = config.get('img_h', 150)
                    offset_x  = config.get('offset_x', 15)
                    offset_y  = config.get('offset_y', 15)
            except: 
                img_w, img_h, offset_x, offset_y = 280, 150, 15, 15
        else:
            img_w, img_h, offset_x, offset_y = 280, 150, 15, 15
            
        self.img_folder_var = tk.StringVar(value=img_dir)
        self.client_var = tk.StringVar(value=client)
        self.report_no_var = tk.StringVar(value=report_no)
        self.date_var = tk.StringVar(value=date)
        self.out_file_var = tk.StringVar(value=out_dir)
        self.img_w_var = tk.IntVar(value=img_w)
        self.img_h_var = tk.IntVar(value=img_h)
        self.offset_x_var = tk.IntVar(value=offset_x)
        self.offset_y_var = tk.IntVar(value=offset_y)
        
        # 스핀박스 직접 입력 시 숫자가 튕기는 버그를 방지하기 위해 실시간 자동 저장을 뺍니다.
        # 설정은 변환을 실행할 때만 저장됩니다.
        
    def save_config(self):
        config = {
            'img_folder': self.img_folder_var.get(),
            'out_file': self.out_file_var.get(),
            'client': self.client_var.get(),
            'report_no': self.report_no_var.get(),
            'date': self.date_var.get(),
            'img_w': self.img_w_var.get(),
            'img_h': self.img_h_var.get(),
            'offset_x': self.offset_x_var.get(),
            'offset_y': self.offset_y_var.get()
        }
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=4)
        except: pass
        
    def create_widgets(self):
        # 1. Image Folder Selection
        folder_frame = tk.LabelFrame(self.root, text="1. 사진 폴더 선택", padx=10, pady=10)
        folder_frame.pack(fill="x", padx=10, pady=5)
        
        tk.Entry(folder_frame, textvariable=self.img_folder_var, state="readonly").pack(side="left", fill="x", expand=True, padx=(0, 5))
        tk.Button(folder_frame, text="폴더 찾기", command=self.browse_folder).pack(side="right")
        
        # 2. Report Information
        info_frame = tk.LabelFrame(self.root, text="2. 보고서 정보 입력", padx=10, pady=10)
        info_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(info_frame, text="발주처:").grid(row=0, column=0, sticky="e", pady=2, padx=5)
        ttk.Entry(info_frame, textvariable=self.client_var, width=40).grid(row=0, column=1, sticky="w", pady=2)
        
        ttk.Label(info_frame, text="REPORT NO:").grid(row=1, column=0, sticky="e", pady=2, padx=5)
        ttk.Entry(info_frame, textvariable=self.report_no_var, width=40).grid(row=1, column=1, sticky="w", pady=2)
        
        ttk.Label(info_frame, text="검사일자:").grid(row=2, column=0, sticky="e", pady=2, padx=5)
        ttk.Entry(info_frame, textvariable=self.date_var, width=40).grid(row=2, column=1, sticky="w", pady=2)

        ttk.Label(info_frame, text="저장 위치:").grid(row=3, column=0, sticky="e", pady=2, padx=5)
        save_frame = tk.Frame(info_frame)
        save_frame.grid(row=3, column=1, sticky="w", pady=2)
        ttk.Entry(save_frame, textvariable=self.out_file_var, width=32).pack(side="left")
        tk.Button(save_frame, text="찾아보기", command=self.browse_save).pack(side="left", padx=5)
        
        # 3. Size Adjustment
        size_frame = tk.LabelFrame(self.root, text="3. 사진 크기 및 위치 완벽 제어 (드래그 또는 숫자 입력)", padx=10, pady=10)
        size_frame.pack(fill="x", padx=10, pady=5)
        
        # 사진 가로 크기
        tk.Label(size_frame, text="사진 가로폭 (기본 280):").grid(row=0, column=0, sticky="e", pady=2)
        w_scale = tk.Scale(size_frame, variable=self.img_w_var, from_=100, to=800, orient="horizontal", length=400, resolution=1, showvalue=False)
        w_scale.grid(row=0, column=1, sticky="w", padx=(5, 5))
        w_spin = ttk.Spinbox(size_frame, textvariable=self.img_w_var, from_=100, to=800, width=5)
        w_spin.grid(row=0, column=2, sticky="w")
        
        # 좌우 위치 조절
        tk.Label(size_frame, text="좌우 위치 조절 (왼쪽 여백):").grid(row=1, column=0, sticky="e", pady=2)
        x_scale = tk.Scale(size_frame, variable=self.offset_x_var, from_=-50, to=300, orient="horizontal", length=400, resolution=1, showvalue=False)
        x_scale.grid(row=1, column=1, sticky="w", padx=(5, 5))
        x_spin = ttk.Spinbox(size_frame, textvariable=self.offset_x_var, from_=-50, to=300, width=5)
        x_spin.grid(row=1, column=2, sticky="w")
        
        # 사진 세로 크기
        tk.Label(size_frame, text="사진 높이 (기본 150):").grid(row=2, column=0, sticky="e", pady=2)
        h_scale = tk.Scale(size_frame, variable=self.img_h_var, from_=50, to=600, orient="horizontal", length=200, resolution=1, showvalue=False)
        h_scale.grid(row=2, column=1, sticky="w", padx=(5, 5))
        h_spin = ttk.Spinbox(size_frame, textvariable=self.img_h_var, from_=50, to=600, width=5)
        h_spin.grid(row=2, column=2, sticky="w")
        
        # 상하 위치 조절
        tk.Label(size_frame, text="상하 위치 조절 (위쪽 여백):").grid(row=3, column=0, sticky="e", pady=2)
        y_scale = tk.Scale(size_frame, variable=self.offset_y_var, from_=-50, to=300, orient="horizontal", length=400, resolution=1, showvalue=False)
        y_scale.grid(row=3, column=1, sticky="w", padx=(5, 5))
        y_spin = ttk.Spinbox(size_frame, textvariable=self.offset_y_var, from_=-50, to=300, width=5)
        y_spin.grid(row=3, column=2, sticky="w")
        
        # 투명 눈금자 & 자동 맞춤 버튼
        btn_frame = tk.Frame(size_frame)
        btn_frame.grid(row=0, column=3, rowspan=4, sticky="nsew", padx=(15, 0))
        
        measure_btn = tk.Button(btn_frame, text="📏 투명 자 띄우기\n(직접 재기)", bg="lightyellow", font=("Arial", 9), command=self.open_measurer)
        measure_btn.pack(side="top", fill="both", expand=True, pady=(2, 2))
        
        def auto_fit():
            self.img_w_var.set(381)
            self.offset_x_var.set(0)
            
        auto_btn = tk.Button(btn_frame, text="✨ 엑셀 셀 크기에\n완벽하게 자동 맞춤", bg="#FFD0D0", font=("Arial", 9, "bold"), command=auto_fit)
        auto_btn.pack(side="top", fill="both", expand=True, pady=(2, 2))
        
        # 4. Execution
        exec_frame = tk.Frame(self.root)
        exec_frame.pack(fill="x", padx=10, pady=15)
        
        tk.Button(exec_frame, text="엑셀 변환 실행", font=("Arial", 12, "bold"), bg="lightblue", command=self.generate_excel).pack(fill="x", ipady=5)
        
    def open_measurer(self):
        """반투명 투명 자: 창 내부 크기 = 실제 픽셀 (테두리 없음)"""
        m = tk.Toplevel(self.root)
        m.overrideredirect(True)
        m.attributes('-alpha', 0.6)
        m.attributes('-topmost', True)
        m.configure(bg='red')   # 빨간 바깥 테두리 역할

        try:
            import ctypes
            dpi = ctypes.windll.user32.GetDpiForSystem()
            self.scale_factor = dpi / 96.0
        except Exception:
            self.scale_factor = 1.0

        W0 = self.img_w_var.get()
        H0 = self.img_h_var.get()
        screen_W = int(W0 * self.scale_factor)
        screen_H = int(H0 * self.scale_factor)
        m.geometry(f"{screen_W+6}x{screen_H+6}+200+150")

        # 안쪽 초록 영역 (= 실제 사진이 들어갈 크기)
        inner = tk.Frame(m, bg='#00AA44')
        inner.place(x=3, y=3, relwidth=1.0, relheight=1.0, width=-6, height=-6)

        # 안내 텍스트
        size_lbl = tk.Label(inner, text=f"{W0} x {H0} px",
                            bg='#00AA44', fg='white', font=("Arial", 13, "bold"))
        size_lbl.place(relx=0.5, rely=0.35, anchor="center")

        hint_lbl = tk.Label(inner, text="● 안쪽 드래그 → 이동\n■ 우하단 드래그 → 크기조절\nEsc·더블클릭 → 닫기",
                            bg='#00AA44', fg='white', font=("Arial", 9), justify="center")
        hint_lbl.place(relx=0.5, rely=0.65, anchor="center")

        # 우하단 크기조절 핸들
        hdl = tk.Label(inner, text="⬛", bg='red', fg='white',
                       cursor='size_nw_se', font=("Arial", 13))
        hdl.place(relx=1.0, rely=1.0, anchor='se')

        # ── 이동 (inner / size_lbl / hint_lbl 드래그) ────────────────────
        _mv = {}
        def mv_press(e):
            _mv['x'] = e.x_root
            _mv['y'] = e.y_root
        def mv_drag(e):
            dx = e.x_root - _mv['x']
            dy = e.y_root - _mv['y']
            m.geometry(f"+{m.winfo_x()+dx}+{m.winfo_y()+dy}")
            _mv['x'] = e.x_root
            _mv['y'] = e.y_root
        for w in (inner, size_lbl, hint_lbl):
            w.bind('<ButtonPress-1>', mv_press)
            w.bind('<B1-Motion>', mv_drag)

        # ── 크기 조절 (우하단 ⬛ 드래그) ─────────────────────────────────
        _rs = {}
        def rs_press(e):
            _rs['x'] = e.x_root
            _rs['y'] = e.y_root
            _rs['w'] = m.winfo_width()
            _rs['h'] = m.winfo_height()
        def rs_drag(e):
            nw = max(90,  _rs['w'] + e.x_root - _rs['x'])
            nh = max(60,  _rs['h'] + e.y_root - _rs['y'])
            m.geometry(f"{int(nw)}x{int(nh)}")
            
            # 투명 자의 물리적 픽셀 크기를 다시 원래 배율로 나눠서 UI 숫자에 반영합니다.
            inner_screen_w = int(nw) - 6
            inner_screen_h = int(nh) - 6
            
            inner_w = int(inner_screen_w / self.scale_factor)
            inner_h = int(inner_screen_h / self.scale_factor)
            
            self.img_w_var.set(inner_w)
            self.img_h_var.set(inner_h)
            size_lbl.config(text=f"{inner_w} x {inner_h} px")
        hdl.bind('<ButtonPress-1>', rs_press)
        hdl.bind('<B1-Motion>', rs_drag)

        # ── 닫기 ─────────────────────────────────────────────────────────
        for w in (m, inner, size_lbl, hint_lbl):
            w.bind('<Escape>', lambda e: m.destroy())
            w.bind('<Double-Button-1>', lambda e: m.destroy())



    def browse_folder(self):
        current_dir = self.img_folder_var.get()
        if not os.path.exists(current_dir):
            current_dir = self.base_path
            
        folder = filedialog.askdirectory(initialdir=current_dir, title="사진이 있는 폴더를 선택하세요")
        if folder:
            self.img_folder_var.set(folder)
            self.save_config()
            
    def browse_save(self):
        current_file = self.out_file_var.get()
        current_dir = os.path.dirname(current_file) if current_file else self.base_path
        current_name = os.path.basename(current_file) if current_file else "NDT_Photo_Log_Final.xlsx"
        
        if not os.path.exists(current_dir):
            current_dir = self.base_path
            
        file_path = filedialog.asksaveasfilename(
            initialdir=current_dir,
            initialfile=current_name,
            title="엑셀 파일 저장 위치 선택",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if file_path:
            self.out_file_var.set(file_path)
            self.save_config()
            
    def generate_excel(self):
        image_folder = self.img_folder_var.get()
        client = self.client_var.get().strip()
        report_no = self.report_no_var.get().strip()
        inspect_date = self.date_var.get().strip()
        output_filename = self.out_file_var.get().strip()
        
        if not output_filename:
            messagebox.showwarning("경고", "저장 위치를 지정해 주세요.")
            return
            
        if not output_filename.lower().endswith('.xlsx'):
            output_filename += '.xlsx'
        logo_filename = 'logo.png'
        
        if not os.path.exists(image_folder):
            os.makedirs(image_folder)
            messagebox.showinfo("안내", f"'{image_folder}' 폴더가 없어 새로 생성했습니다.\n해당 폴더에 사진 파일을 넣은 뒤 다시 실행해 주세요.")
            return
            
        all_files = sorted(glob.glob(os.path.join(image_folder, '*.[jJ][pP][gG]')) + 
                           glob.glob(os.path.join(image_folder, '*.[pP][nN][gG]')) +
                           glob.glob(os.path.join(image_folder, '*.[jJ][pP][eE][gG]')) +
                           glob.glob(os.path.join(image_folder, '*.[bB][mM][pP]')))
                           
        image_files = [f for f in all_files if os.path.splitext(os.path.basename(f))[0].lower() != 'logo']
        
        if len(image_files) == 0:
            messagebox.showwarning("경고", f"선택한 폴더에 사진이 없습니다.\n경로: {image_folder}")
            return
            
        try:
            workbook = xlsxwriter.Workbook(output_filename)
            worksheet = workbook.add_worksheet()
            
            worksheet.set_paper(9)
            worksheet.set_portrait()
            # 마진을 균형있게 설정하고 페이지의 정중앙에 출력되도록 합니다.
            worksheet.set_margins(left=0.2, right=0.2, top=0.5, bottom=0.5)
            worksheet.center_vertically()
            worksheet.center_horizontally()
            # 높이는 나중에 동적으로 계산해서 fit_to_pages를 적용합니다.
            
            worksheet.repeat_rows(0, 4)
            worksheet.set_footer('&C&P / &N')
            
            offset_x = self.offset_x_var.get()
            offset_y = self.offset_y_var.get()
            # 사용자가 설정한 여백만 사용 (사진 폭은 xlsxwriter가 열 너비에 맞게 비율 계산)
            # 열 너비는 fit_to_pages와 함께 A4에 딱 맞게 자동 조정됩니다.
            worksheet.set_column('A:B', 40)
            
            title_format = workbook.add_format({'font_name': '맑은 고딕', 'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'shrink': True})
            company_format = workbook.add_format({'font_name': '맑은 고딕', 'align': 'right', 'valign': 'vcenter', 'border': 1, 'font_size': 9, 'text_wrap': True})
            center_border = workbook.add_format({'font_name': '맑은 고딕', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_size': 10})
            bold_format = workbook.add_format({'font_name': '맑은 고딕', 'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
            desc_format = workbook.add_format({'font_name': '맑은 고딕', 'align': 'left', 'valign': 'vcenter', 'border': 1, 'font_size': 10, 'shrink': True, 'text_wrap': False, 'indent': 1})
            image_cell_format = workbook.add_format({'border': 1})
            
            worksheet.merge_range('A1:B1', "REPORT OF PHASED ARRAY UT EXAMINATION (위 상 배 열 초 음 파 탐 상 검 사 보 고 서)", title_format)
            company_info_text = "서   울   檢   査   株   式   會   社\nSEOUL INSPECTION & TESTING Co., Ltd.\n서울특별시 서초구 바우뫼로 41길 54\nTEL : (02) 552-1112   FAX : (02) 2058-0720"
            worksheet.merge_range('A2:A4', company_info_text, company_format)
            
            logo_path = os.path.join(self.base_path, logo_filename)
            if not os.path.exists(logo_path):
                logo_path = os.path.join(image_folder, logo_filename)
            if not os.path.exists(logo_path):
                parent_dir = os.path.dirname(self.base_path)
                logo_path = os.path.join(parent_dir, 'resources', logo_filename)
                
            if os.path.exists(logo_path):
                try:
                    for r in range(1, 4): worksheet.set_row(r, 15)
                    with Image.open(logo_path) as img:
                        w, h = img.size
                        scale = min(285/w, 60/h) * 0.8
                        y_pos = (60 - (h * scale)) / 2
                        worksheet.insert_image('A2', logo_path, {'x_scale': scale, 'y_scale': scale, 'x_offset': 15, 'y_offset': y_pos, 'object_position': 1})
                except: pass
                
            worksheet.write('B2', f"발주처: {client}", center_border)
            worksheet.write('B3', f"REPORT NO: {report_no}", center_border)
            worksheet.write('B4', f"검사일자: {inspect_date}", center_border)
            worksheet.merge_range('A5:B5', "PHOTO LOG (사진 대장)", bold_format)
            
            row = 5
            col = 0
            
            target_w = self.img_w_var.get()
            target_h = self.img_h_var.get()
            offset_x = self.offset_x_var.get()
            offset_y = self.offset_y_var.get()

            # 셀 너비 동적 랩핑: A4 절반 사이즈(386)를 최소한으로 보장하여 레이아웃 축소를 막고,
            # 사용자가 사진을 386 이상으로 키울 경우 사진이 셀 테두리를 넘어 엑셀 버그가 발생하는 것을 막기 위해
            # 사진 크기에 맞춰 셀 너비가 자동으로 함께 늘어나도록 방어 코드를 추가합니다. (+5px 안전여백)
            safe_cell_width = max(386, target_w + (offset_x * 2) + 5)
            worksheet.set_column_pixels('A:B', safe_cell_width)

            DESC_ROW_HEIGHT = 25
            h_breaks = []
            
            for image_path in image_files:
                # 행 높이 자동 맞춤: (사진 세로 크기 + 위아래 여백)
                worksheet.set_row_pixels(row, target_h + (offset_y * 2))
                worksheet.write(row, col, "", image_cell_format)
                try:
                    with Image.open(image_path) as img:
                        if img.mode not in ('RGB', 'RGBA'):
                            img = img.convert('RGB')
                            
                        # 엑셀의 줌 배율 및 숨은 DPI 메타데이터 버그를 원천 차단하기 위해
                        # 파이썬에서 이미지를 지정된 픽셀로 아예 "물리적 리사이징" 해버립니다.
                        resample_filter = getattr(Image, 'Resampling', Image).LANCZOS
                        resized_img = img.resize((target_w, target_h), resample_filter)
                        
                        img_byte_arr = io.BytesIO()
                        # 모든 원본 메타데이터(DPI 등)를 날리고 순수 96 DPI로 저장
                        resized_img.save(img_byte_arr, format='PNG', dpi=(96, 96))
                        img_byte_arr.seek(0)
                        
                    worksheet.insert_image(row, col, f"img_{row}_{col}.png", {
                        'image_data': img_byte_arr,
                        'x_scale': 1.0,
                        'y_scale': 1.0,
                        'x_offset': offset_x,
                        'y_offset': offset_y,
                        'object_position': 3
                    })
                except Exception as e:
                    worksheet.set_row_pixels(row, 200)
                    worksheet.write(row, col, "", image_cell_format)
                    print(f"이미지 오류({os.path.basename(image_path)}): {e}")
                    
                name_only = os.path.splitext(os.path.basename(image_path))[0]
                worksheet.set_row(row + 1, DESC_ROW_HEIGHT)
                worksheet.write(row + 1, col, f"설명: {name_only}", desc_format)
                
                if col == 0: 
                    col = 1
                else: 
                    col = 0
                    row += 2 
                    # 4줄(8장)마다 페이지 나누기
                    if (row - 5) % 8 == 0:
                        h_breaks.append(row)
                
            if col == 1:
                worksheet.write(row, 1, "", image_cell_format)
                
            if h_breaks:
                worksheet.set_h_pagebreaks(h_breaks)
                
            # 데이터가 있는 곳까지만 정확히 인쇄 영역으로 설정 (빈 페이지 방지)
            worksheet.print_area(0, 0, row + 1, 1)
            
            # 전체 사진 수를 바탕으로 총 페이지 수(4줄=8장 당 1페이지) 계산 후 엑셀의 강제 맞춤 기능 적용
            # 주의: fit_to_pages가 적용되어도 엑셀 '기본 보기'에서는 100% 기준의 점선이 보여 페이지가 늘어난 것처럼 보일 수 있으나,
            # 실제 인쇄 미리보기(Ctrl+P)에서는 무조건 이 total_pages 숫자에 맞춰서 출력됩니다.
            total_pages = math.ceil(len(image_files) / 8)
            if total_pages > 0:
                worksheet.fit_to_pages(1, total_pages)
                
            workbook.close()
            self.save_config()
            messagebox.showinfo("성공", f"엑셀 파일이 성공적으로 생성되었습니다!\n저장 위치: {output_filename}")
            
        except Exception as e:
            messagebox.showerror("오류", f"엑셀 생성 중 오류가 발생했습니다:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = PhotoLogApp(root)
    root.mainloop()