import os
import sys
import traceback
import tkinter as tk
from tkinter import messagebox, filedialog, ttk
from PIL import Image, ImageTk

# OCR Libraries Check
IMPORT_ERRORS = []

try:
    import easyocr
    EASYOCR_AVAILABLE = True
except Exception as e:
    EASYOCR_AVAILABLE = False
    IMPORT_ERRORS.append(f"EasyOCR Error: {e}")

try:
    import pytesseract
    # Tesseract specific path (common for Windows)
    tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    if os.path.exists(tesseract_cmd):
        pytesseract.pytesseract.tesseract_cmd = tesseract_cmd
    PYTESSERACT_AVAILABLE = True
except Exception as e:
    PYTESSERACT_AVAILABLE = False
    IMPORT_ERRORS.append(f"PyTesseract Error: {e}")

class OCRExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("OCR Text Extractor - 숫지/글자/기호 추출기")
        self.root.geometry("1000x700")
        
        self.image_path = None
        self.reader = None # EasyOCR reader instance
        
        self.setup_ui()
        self.check_initial_errors()
        
    def setup_ui(self):
        # Top Control Frame
        control_frame = ttk.Frame(self.root, padding=10)
        control_frame.pack(side="top", fill="x")
        
        ttk.Button(control_frame, text="이미지 선택 (Select Image)", command=self.load_image).pack(side="left", padx=5)
        
        self.engine_var = tk.StringVar(value="EasyOCR" if EASYOCR_AVAILABLE else "Tesseract")
        ttk.Label(control_frame, text="OCR 엔진:").pack(side="left", padx=(20, 5))
        engine_combo = ttk.Combobox(control_frame, textvariable=self.engine_var, values=["EasyOCR", "Tesseract"], state="readonly", width=10)
        engine_combo.pack(side="left", padx=5)
        
        ttk.Button(control_frame, text="텍스트 추출 하기 (Extract)", command=self.run_ocr).pack(side="left", padx=20)
        ttk.Button(control_frame, text="클립보드 복사 (Copy)", command=self.copy_to_clipboard).pack(side="right", padx=5)

        # Paned Window for Preview and Results
        paned = ttk.PanedWindow(self.root, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=10, pady=5)
        
        # Left: Image Preview
        preview_frame = ttk.LabelFrame(paned, text="이미지 미리보기 (Preview)")
        paned.add(preview_frame, weight=1)
        
        self.canvas = tk.Canvas(preview_frame, bg="gray")
        self.canvas.pack(fill="both", expand=True)
        
        # Right: Result Text
        result_frame = ttk.LabelFrame(paned, text="추출 결과 (Results)")
        paned.add(result_frame, weight=1)
        
        self.result_text = tk.Text(result_frame, font=("Malgun Gothic", 12), wrap="word")
        self.result_text.pack(fill="both", expand=True)
        
        # Status Bar
        self.status_var = tk.StringVar(value="대기 중...")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief="sunken", anchor="w")
        status_bar.pack(side="bottom", fill="x")
        
    def check_initial_errors(self):
        if IMPORT_ERRORS:
            error_msg = "프로그램 시작 중 일부 모듈에서 오류가 발생했습니다:\n\n" + "\n".join(IMPORT_ERRORS)
            self.update_results(error_msg + "\n\n(위 내용을 복사하여 담당자에게 전달해주세요.)")
            self.status_var.set("모듈 로드 오류 발생")

    def load_image(self):
        path = filedialog.askopenfilename(filetypes=[("Image Files", "*.png *.jpg *.jpeg *.bmp *.tiff")])
        if path:
            self.image_path = path
            self.display_image(path)
            self.status_var.set(f"파일 로드됨: {os.path.basename(path)}")
            
    def display_image(self, path):
        try:
            img = Image.open(path)
            # Resize for preview
            img.thumbnail((500, 600))
            self.tk_img = ImageTk.PhotoImage(img)
            self.canvas.delete("all")
            self.canvas.create_image(0, 0, anchor="nw", image=self.tk_img)
        except Exception as e:
            self.show_detailed_error(f"이미지 표시 중 오류: {e}")

    def run_ocr(self):
        if not self.image_path:
            messagebox.showwarning("오류", "먼저 이미지를 선택해주세요.")
            return
            
        engine = self.engine_var.get()
        self.status_var.set(f"Processing with {engine}...")
        self.root.update_idletasks()
        
        try:
            if engine == "EasyOCR":
                if not EASYOCR_AVAILABLE:
                    messagebox.showerror("엔진 오류", "EasyOCR 패키지가 설치되어 있지 않습니다.\npip install easyocr 를 실행해주세요.")
                    return
                self.extract_easyocr()
            else:
                if not PYTESSERACT_AVAILABLE:
                    messagebox.showerror("엔진 오류", "pytesseract 패키지가 설치되어 있지 않거나 Tesseract 엔진이 없습니다.")
                    return
                self.extract_tesseract()
        except Exception as e:
            self.show_detailed_error(f"작업 중 오류 발생: {e}")
            self.status_var.set("오류 발생")

    def extract_easyocr(self):
        if self.reader is None:
            # Initialize reader (Korean + English)
            self.reader = easyocr.Reader(['ko', 'en'])
        
        results = self.reader.readtext(self.image_path)
        # Sort results by vertical position then horizontal
        results.sort(key=lambda x: (x[0][0][1], x[0][0][0]))
        
        full_text = ""
        last_y = -1
        for res in results:
            text = res[1]
            # Simple heuristic for new line
            current_y = res[0][0][1]
            if last_y != -1 and abs(current_y - last_y) > 10:
                full_text += "\n"
            full_text += text + " "
            last_y = current_y
            
        self.update_results(full_text)
        self.status_var.set("추출 완료 (EasyOCR)")

    def extract_tesseract(self):
        # Use lang='kor+eng'
        text = pytesseract.image_to_string(Image.open(self.image_path), lang='kor+eng')
        self.update_results(text)
        self.status_var.set("추출 완료 (Tesseract)")

    def update_results(self, text):
        self.result_text.delete("1.0", tk.END)
        self.result_text.insert(tk.END, text.strip())
        
    def show_detailed_error(self, message):
        """오류 내용을 결과창에 출력하여 복사 가능하게 함"""
        err_details = traceback.format_exc()
        full_msg = f"!!! 오류 발생 (ERROR Details) !!!\n\n{message}\n\n--- 상세 로그 (Traceback) ---\n{err_details}"
        self.update_results(full_msg)
        messagebox.showerror("오류", f"{message}\n\n상세 내용은 결과창을 확인하여 복사해주세요.")

    def copy_to_clipboard(self):
        content = self.result_text.get("1.0", tk.END).strip()
        if content:
            self.root.clipboard_clear()
            self.root.clipboard_append(content)
            self.root.update()
            messagebox.showinfo("복사 완료", "텍스트 또는 오류 로그가 클립보드에 복사되었습니다.")

if __name__ == "__main__":
    root = tk.Tk()
    app = OCRExtractorApp(root)
    root.mainloop()
