import os
import io
import math
import tkinter as tk
from tkinter import filedialog, messagebox
import xlsxwriter
from PIL import Image

class SimplePhotoLogApp:
    def __init__(self, root):
        self.root = root
        self.root.title("A4 8장 사진대장 자동 생성기 (완벽 맞춤판)")
        self.root.geometry("500x250")
        self.root.resizable(False, False)
        
        self.img_folder_var = tk.StringVar()
        
        # UI 구성
        tk.Label(root, text="A4 1장에 사진 8장(2x4)이 완벽하게 들어가는 사진대장을 만듭니다.", font=("Arial", 10, "bold")).pack(pady=20)
        
        frame = tk.Frame(root)
        frame.pack(pady=10, fill="x", padx=20)
        
        tk.Label(frame, text="사진 폴더:").pack(side="left")
        tk.Entry(frame, textvariable=self.img_folder_var, state="readonly", width=40).pack(side="left", padx=5)
        tk.Button(frame, text="폴더 선택", command=self.browse).pack(side="left")
        
        tk.Button(root, text="✨ 엑셀 변환 실행", font=("Arial", 12, "bold"), bg="#E0FFE0", command=self.generate, width=30, height=2).pack(pady=20)
        
    def browse(self):
        folder = filedialog.askdirectory(title="사진이 있는 폴더를 선택하세요")
        if folder:
            self.img_folder_var.set(folder)
            
    def generate(self):
        folder = self.img_folder_var.get()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("경고", "사진 폴더를 선택해주세요.")
            return
            
        image_files = [os.path.join(folder, f) for f in os.listdir(folder) if f.lower().endswith(('.png', '.jpg', '.jpeg'))]
        if not image_files:
            messagebox.showwarning("경고", "폴더에 사진 파일이 없습니다.")
            return
            
        out_file = os.path.join(folder, "사진대장_A4_8장_완벽맞춤.xlsx")
        
        try:
            workbook = xlsxwriter.Workbook(out_file)
            worksheet = workbook.add_worksheet('Photo Log')
            
            # A4 기준 완벽 핏 설정 (셀: 400px, 사진 최대 크기: 396px)
            cell_w = 400
            img_w = 396
            img_h = 164
            
            worksheet.set_column_pixels('A:B', cell_w)
            
            # 서식 지정
            title_format = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter'})
            desc_format = workbook.add_format({'font_size': 11, 'align': 'left', 'valign': 'vcenter', 'border': 1})
            image_cell_format = workbook.add_format({'border': 1})
            
            # 제목 줄
            worksheet.merge_range('A1:B1', "PHOTO LOG (사진 대장)", title_format)
            worksheet.set_row_pixels(0, 40)
            
            row = 1
            col = 0
            h_breaks = []
            
            for i, img_path in enumerate(image_files):
                # 이미지 물리적 리사이징 (원본 비율 완벽 유지 & 찌그러짐 방지)
                with Image.open(img_path) as img:
                    if img.mode not in ('RGB', 'RGBA'):
                        img = img.convert('RGB')
                    resample_filter = getattr(Image, 'Resampling', Image).LANCZOS
                    
                    # 억지로 찌그러뜨리지 않고, 원본 비율을 유지하면서 박스 안에 맞춤
                    img.thumbnail((img_w, img_h), resample_filter)
                    new_w, new_h = img.size
                    
                    img_byte_arr = io.BytesIO()
                    img.save(img_byte_arr, format='PNG', dpi=(96, 96))
                    img_byte_arr.seek(0)
                
                # 사진 셀 설정 (항상 일정한 높이 유지)
                worksheet.set_row_pixels(row, img_h + 10) # 위아래 여백 5px씩
                worksheet.write(row, col, "", image_cell_format)
                
                # 비율이 달라 남는 공간을 계산하여 사진을 셀의 정중앙에 배치
                center_x = 2 + (img_w - new_w) // 2
                center_y = 5 + (img_h - new_h) // 2
                
                worksheet.insert_image(row, col, f"img_{i}.png", {
                    'image_data': img_byte_arr,
                    'x_offset': center_x,
                    'y_offset': center_y,
                    'x_scale': 1.0,
                    'y_scale': 1.0,
                    'object_position': 3
                })
                
                # 설명 셀 설정
                name_only = os.path.basename(img_path)
                worksheet.set_row_pixels(row + 1, 30)
                worksheet.write(row + 1, col, f"설명: {name_only}", desc_format)
                
                # 2열(A, B) 배치 로직
                if col == 0:
                    col = 1
                else:
                    col = 0
                    row += 2
                    
                    # 4줄(사진 8장)마다 페이지 나누기
                    if (row - 1) % 8 == 0:
                        h_breaks.append(row)
                        
            # 페이지 나누기 적용 (마지막 사진이 홀수개일 때)
            if col == 1:
                worksheet.write(row, 1, "", image_cell_format)
                worksheet.write(row+1, 1, "", desc_format)
                
            if h_breaks:
                worksheet.set_h_pagebreaks(h_breaks)
                
            # 인쇄 설정 (A4 한 장에 무조건 맞춤)
            worksheet.print_area(0, 0, row + 1, 1)
            worksheet.set_margins(left=0.3, right=0.3, top=0.5, bottom=0.5)
            
            total_pages = math.ceil(len(image_files) / 8)
            if total_pages > 0:
                worksheet.fit_to_pages(1, total_pages)
                
            workbook.close()
            messagebox.showinfo("완료", f"엑셀 파일이 생성되었습니다!\n저장위치: {out_file}")
            
        except Exception as e:
            messagebox.showerror("오류", f"엑셀 생성 중 오류 발생:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = SimplePhotoLogApp(root)
    root.mainloop()
