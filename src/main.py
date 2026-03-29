import os
import sys
import subprocess
import traceback
import json
import base64
import tkinter as tk
from tkinter import messagebox, filedialog, ttk

# ========== AUTO-INSTALL ==========
def _ensure_package(import_name, pip_name=None):
    if pip_name is None:
        pip_name = import_name
    try:
        __import__(import_name)
    except ImportError:
        print(f"[AUTO-INSTALL] {pip_name} 설치 중...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", pip_name],
                              stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

for _imp, _pip in [("PIL", "pillow"), ("google.genai", "google-genai"), ("easyocr", "easyocr"), ("fitz", "pymupdf")]:
    try:
        _ensure_package(_imp, _pip)
    except Exception:
        pass

from PIL import Image, ImageTk
from google import genai

# ========== CONFIG ==========
CONFIG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "config")
CONFIG_FILE = os.path.join(CONFIG_DIR, "gemini_config.json")

def load_api_key():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            return json.load(f).get("api_key", "")
    return ""

def save_api_key(key):
    os.makedirs(CONFIG_DIR, exist_ok=True)
    with open(CONFIG_FILE, "w") as f:
        json.dump({"api_key": key}, f)


class TextExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("📝 사진/스캔 문자 추출기 (Google AI)")
        self.root.geometry("1100x750")
        
        self.image_path = None
        self.api_key = load_api_key()
        self.easyocr_reader = None  # Lazy load reader
        
        self.setup_ui()
        
    def setup_ui(self):
        control_frame = ttk.Frame(self.root, padding=10)
        control_frame.pack(side="top", fill="x")
        
        ttk.Button(control_frame, text="📁 파일 선택", command=self.load_file).pack(side="left", padx=5)
        ttk.Button(control_frame, text="🔍 텍스트 추출", command=self.run_ocr).pack(side="left", padx=15)
        ttk.Button(control_frame, text="📋 복사", command=self.copy_to_clipboard).pack(side="right", padx=5)
        
        # 엔진 선택 및 기타 설정
        engine_frame = ttk.LabelFrame(control_frame, text="분석 엔진")
        engine_frame.pack(side="left", padx=10)
        self.engine_var = tk.StringVar(value="Gemini")
        ttk.Radiobutton(engine_frame, text="Gemini (Cloud)", variable=self.engine_var, value="Gemini").pack(side="left", padx=5)
        ttk.Radiobutton(engine_frame, text="EasyOCR (Local)", variable=self.engine_var, value="EasyOCR").pack(side="left", padx=5)

        ttk.Button(control_frame, text="⚙ API 설정", command=self.setup_api_key).pack(side="left", padx=5)
        ttk.Button(control_frame, text="🗑 지우기", command=self.clear_results).pack(side="left", padx=5)

        paned = ttk.PanedWindow(self.root, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=10, pady=5)
        
        preview_frame = ttk.LabelFrame(paned, text="이미지 미리보기")
        paned.add(preview_frame, weight=1)
        self.canvas = tk.Canvas(preview_frame, bg="#2d2d2d")
        self.canvas.pack(fill="both", expand=True)
        
        result_frame = ttk.LabelFrame(paned, text="추출 결과")
        paned.add(result_frame, weight=1)
        self.result_text = tk.Text(result_frame, font=("Malgun Gothic", 12), wrap="word")
        scrollbar = ttk.Scrollbar(result_frame, command=self.result_text.yview)
        self.result_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.result_text.pack(fill="both", expand=True)
        
        self.status_var = tk.StringVar(value="파일을 선택해주세요")
        ttk.Label(self.root, textvariable=self.status_var, relief="sunken", anchor="w").pack(side="bottom", fill="x")

    def setup_api_key(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Google Gemini API 키 설정")
        dialog.geometry("500x150")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="API 키 (aistudio.google.com/apikey 에서 발급):").pack(pady=(15, 5), padx=15, anchor="w")
        key_entry = ttk.Entry(dialog, width=60)
        key_entry.pack(padx=15, fill="x")
        key_entry.insert(0, self.api_key)
        
        def save():
            key = key_entry.get().strip()
            if key:
                save_api_key(key)
                self.api_key = key
                messagebox.showinfo("저장", "API 키가 저장되었습니다!")
                dialog.destroy()
        
        ttk.Button(dialog, text="저장", command=save).pack(pady=15)

    def load_file(self):
        path = filedialog.askopenfilename(filetypes=[
            ("이미지 및 PDF 파일", "*.png *.jpg *.jpeg *.bmp *.tiff *.tif *.pdf"),
            ("모든 파일", "*.*")
        ])
        if path:
            self.image_path = path
            self.display_image(path)
            self.status_var.set(f"파일 로드됨: {os.path.basename(path)}")
            
    def display_image(self, path):
        try:
            if path.lower().endswith(".pdf"):
                import fitz
                doc = fitz.open(path)
                page = doc.load_page(0)
                pix = page.get_pixmap()
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                doc.close()
            else:
                img = Image.open(path)
                
            img.thumbnail((500, 650))
            self.tk_img = ImageTk.PhotoImage(img)
            self.canvas.delete("all")
            self.canvas.create_image(0, 0, anchor="nw", image=self.tk_img)
        except Exception as e:
            messagebox.showerror("오류", f"파일 표시 오류: {e}")

    def run_ocr(self):
        if not self.image_path:
            messagebox.showwarning("알림", "먼저 파일을 선택해주세요.")
            return
        
        if self.engine_var.get() == "Gemini":
            self.run_gemini_ocr()
        else:
            self.run_local_ocr()

    def run_gemini_ocr(self):
        if not self.api_key:
            messagebox.showwarning("API 키 필요", "먼저 '⚙ API 키 설정'에서 Gemini API 키를 입력해주세요.")
            self.setup_api_key()
            return
        
        try:
            from google import genai
        except ImportError:
            messagebox.showerror("오류", "google-genai 패키지가 설치되지 않았습니다.")
            return

        client = genai.Client(api_key=self.api_key)
        img = Image.open(self.image_path)
        
        prompt = (
            "이 이미지에서 보이는 모든 텍스트를 정확하게 추출해주세요. "
            "손글씨, 인쇄체 모두 포함합니다. "
            "원본의 줄바꿈과 레이아웃을 최대한 유지해주세요. "
            "텍스트만 출력하고 다른 설명은 하지 마세요."
        )
        
        # 여러 모델 시도 (할당량 분산)
        models = ["gemini-2.0-flash-lite", "gemini-2.0-flash"]
        
        for model_name in models:
            try:
                self.status_var.set(f"{model_name} 으로 분석 중...")
                self.root.update()
                
                if self.image_path.lower().endswith(".pdf"):
                    # PDF 파일 API 업로드
                    self.status_var.set(f"PDF 파일 업로드 중 ({model_name})...")
                    self.root.update()
                    uploaded_file = client.files.upload(file=self.image_path)
                    
                    # 업로드 완료 대기
                    import time
                    while uploaded_file.state == "PROCESSING":
                        time.sleep(1)
                        uploaded_file = client.files.get(name=uploaded_file.name)
                    
                    if uploaded_file.state == "FAILED":
                        raise Exception("Gemini 파일 업로드 분석 실패")
                        
                    response = client.models.generate_content(
                        model=model_name,
                        contents=[prompt, uploaded_file]
                    )
                else:
                    img = Image.open(self.image_path)
                    response = client.models.generate_content(
                        model=model_name,
                        contents=[prompt, img]
                    )
                
                result_text = response.text.strip()
                if result_text:
                    self.update_results(result_text)
                    self.status_var.set(f"추출 완료 ({model_name})")
                    return
                    
            except Exception as e:
                err = str(e)
                if "429" in err or "quota" in err.lower() or "ResourceExhausted" in err:
                    self.status_var.set(f"{model_name} 할당량 초과, 다른 모델 시도 중...")
                    self.root.update()
                    import time
                    time.sleep(3)
                    continue
                elif "404" in err or "not found" in err.lower():
                    continue
                else:
                    self.update_results(f"오류 발생:\n{err}")
                    self.status_var.set("오류 발생")
                    return

        
        # 모든 모델 실패
        self.update_results(
            "⚠️ 모든 모델의 무료 할당량이 소진되었습니다.\n\n"
            "해결 방법:\n"
            "1. 잠시 후(약 1분) 다시 시도해주세요\n"
            "2. 내일 다시 시도하면 할당량이 초기화됩니다\n"
            "3. https://aistudio.google.com 에서 새 API 키를 발급받아 사용할 수도 있습니다"
        )
        self.status_var.set("할당량 소진 - 잠시 후 재시도")


    def run_local_ocr(self):
        try:
            import easyocr
            import numpy as np
            
            if self.easyocr_reader is None:
                self.status_var.set("EasyOCR 모델 로딩 중 (첫 실행 시 시간이 걸릴 수 있습니다)...")
                self.root.update()
                # 'ko' (Korean), 'en' (English)
                self.easyocr_reader = easyocr.Reader(['ko', 'en'], gpu=True) # GPU available check is internal to EasyOCR
            
            self.status_var.set("로컬 엔진으로 분석 중...")
            self.root.update()
            
            all_text_parts = []
            
            if self.image_path.lower().endswith(".pdf"):
                import fitz
                doc = fitz.open(self.image_path)
                total_pages = len(doc)
                
                for i in range(total_pages):
                    self.status_var.set(f"로컬 분석 중 (페이지 {i+1}/{total_pages})...")
                    self.root.update()
                    
                    page = doc.load_page(i)
                    pix = page.get_pixmap()
                    img_pil = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    img_np = np.array(img_pil)
                    
                    results = self.easyocr_reader.readtext(img_np, detail=0)
                    all_text_parts.append(f"--- [Page {i+1}] ---\n" + "\n".join(results))
                doc.close()
                result_text = "\n\n".join(all_text_parts)
            else:
                # PIL을 사용하여 이미지를 로드 (유니코드 경로 지원) 후 numpy 배열로 변환
                img_pil = Image.open(self.image_path).convert('RGB')
                img_np = np.array(img_pil)
                results = self.easyocr_reader.readtext(img_np, detail=0)
                result_text = "\n".join(results)
            
            if result_text.strip():
                self.update_results(result_text)
                self.status_var.set("추출 완료 (EasyOCR)")
            else:
                self.update_results("텍스트를 찾을 수 없습니다.")
                self.status_var.set("추출 실패")
                
        except Exception as e:
            err = str(e)
            self.update_results(f"로컬 OCR 오류 발생:\n{err}")
            self.status_var.set("오류 발생")


    def update_results(self, text):
        self.result_text.delete("1.0", tk.END)
        self.result_text.insert(tk.END, text.strip())

    def clear_results(self):
        self.result_text.delete("1.0", tk.END)
        self.status_var.set("결과 초기화됨")

    def copy_to_clipboard(self):
        content = self.result_text.get("1.0", tk.END).strip()
        if content:
            self.root.clipboard_clear()
            self.root.clipboard_append(content)
            self.root.update()
            messagebox.showinfo("복사 완료", "클립보드에 복사되었습니다.")
        else:
            messagebox.showinfo("알림", "복사할 내용이 없습니다.")

if __name__ == "__main__":
    root = tk.Tk()
    app = TextExtractorApp(root)
    root.mainloop()
