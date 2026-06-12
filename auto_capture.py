import os
import time
import subprocess
import threading
import pyautogui
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from pathlib import Path

class AutoCaptureApp:
    def __init__(self, root):
        self.root = root
        self.root.title("OmniPC Auto Capture")
        self.root.geometry("500x550")
        
        # 기본 경로 설정
        self.default_data_dir = r"C:\Users\jjch2\Desktop\data"
        self.default_capture_dir = r"C:\Users\jjch2\Desktop\captuer"
        self.default_shortcut = r"C:\Users\jjch2\Desktop\OmniPC 6.0.lnk"
        
        self.create_widgets()
        
    def create_widgets(self):
        # --- Data Directory ---
        tk.Label(self.root, text="데이터 폴더 (.opd 파일 위치):").pack(anchor="w", padx=10, pady=(10, 0))
        frame1 = tk.Frame(self.root)
        frame1.pack(fill="x", padx=10)
        self.data_var = tk.StringVar(value=self.default_data_dir)
        tk.Entry(frame1, textvariable=self.data_var).pack(side="left", fill="x", expand=True, padx=(0, 5))
        tk.Button(frame1, text="찾아보기", command=lambda: self.browse_dir(self.data_var)).pack(side="right")
        
        # --- Capture Directory ---
        tk.Label(self.root, text="캡처 저장 폴더:").pack(anchor="w", padx=10, pady=(10, 0))
        frame2 = tk.Frame(self.root)
        frame2.pack(fill="x", padx=10)
        self.capture_var = tk.StringVar(value=self.default_capture_dir)
        tk.Entry(frame2, textvariable=self.capture_var).pack(side="left", fill="x", expand=True, padx=(0, 5))
        tk.Button(frame2, text="찾아보기", command=lambda: self.browse_dir(self.capture_var)).pack(side="right")
        
        # --- Shortcut Path ---
        tk.Label(self.root, text="OmniPC 바로가기 경로:").pack(anchor="w", padx=10, pady=(10, 0))
        frame3 = tk.Frame(self.root)
        frame3.pack(fill="x", padx=10)
        self.shortcut_var = tk.StringVar(value=self.default_shortcut)
        tk.Entry(frame3, textvariable=self.shortcut_var).pack(side="left", fill="x", expand=True, padx=(0, 5))
        tk.Button(frame3, text="찾아보기", command=lambda: self.browse_file(self.shortcut_var)).pack(side="right")
        
        # --- Delay Time ---
        frame4 = tk.Frame(self.root)
        frame4.pack(fill="x", padx=10, pady=10)
        tk.Label(frame4, text="대기 시간(초): ").pack(side="left")
        self.delay_var = tk.IntVar(value=12)
        tk.Entry(frame4, textvariable=self.delay_var, width=5).pack(side="left")
        tk.Label(frame4, text=" (파일 열고 캡처할 때까지 대기)").pack(side="left")
        
        # --- Start Button ---
        self.start_btn = tk.Button(self.root, text="▶ 자동 캡처 시작", bg="lightblue", font=("Arial", 11, "bold"), command=self.start_capture_thread)
        self.start_btn.pack(pady=10, fill="x", padx=50)
        
        # --- Log Text ---
        tk.Label(self.root, text="진행 상황:").pack(anchor="w", padx=10)
        self.log_text = scrolledtext.ScrolledText(self.root, height=12)
        self.log_text.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
    def browse_dir(self, var):
        folder = filedialog.askdirectory(initialdir=var.get())
        if folder:
            var.set(folder)
            
    def browse_file(self, var):
        file = filedialog.askopenfilename(initialdir=os.path.dirname(var.get()), filetypes=[("Shortcut", "*.lnk"), ("Executable", "*.exe"), ("All files", "*.*")])
        if file:
            var.set(file)
            
    def log(self, msg):
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.root.update()

    def start_capture_thread(self):
        self.start_btn.config(state="disabled", text="진행 중...")
        self.log_text.delete(1.0, tk.END)
        # UI가 멈추지 않도록 별도의 쓰레드에서 캡처 작업 실행
        threading.Thread(target=self.run_capture, daemon=True).start()

    def run_capture(self):
        try:
            data_dir = Path(self.data_var.get())
            capture_dir = Path(self.capture_var.get())
            shortcut_path = self.shortcut_var.get()
            delay = self.delay_var.get()
            
            if not data_dir.exists():
                self.log(f"에러: 데이터 폴더를 찾을 수 없습니다. ({data_dir})")
                return
                
            capture_dir.mkdir(parents=True, exist_ok=True)
            
            opd_files = list(data_dir.glob("*.opd"))
            self.log(f"총 {len(opd_files)}개의 .opd 파일을 찾았습니다.\n")
            
            if len(opd_files) == 0:
                self.log("작업을 취소합니다. 폴더 내에 .opd 파일이 없습니다.")
                return

            for idx, opd_file in enumerate(opd_files, 1):
                self.log(f"[{idx}/{len(opd_files)}] {opd_file.name} 처리 중...")
                
                try:
                    # 1. 프로그램 실행
                    cmd = ['powershell', '-Command', f'Start-Process "{shortcut_path}" -ArgumentList "{opd_file}"']
                    subprocess.run(cmd)
                    
                    # 2. 대기 (카운트다운 로그 표시)
                    for i in range(delay, 0, -1):
                        self.log(f"  -> 화면 로딩 대기 중... {i}초 남음")
                        time.sleep(1)
                        
                    # 3. 캡처
                    screenshot_path = capture_dir / f"{opd_file.stem}_capture.png"
                    pyautogui.screenshot(str(screenshot_path))
                    self.log(f"  -> 📸 캡처 완료: {screenshot_path.name}")
                    
                except Exception as e:
                    self.log(f"  -> ❌ 에러 발생: {e}")
                    
                finally:
                    # 4. 프로그램 강제 종료
                    subprocess.run(['taskkill', '/F', '/IM', 'OmniPC.exe'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                    self.log("  -> 다음 파일을 위해 프로그램을 닫습니다.\n")
                    time.sleep(2)
                    
            self.log("✅ 모든 캡처 작업이 성공적으로 완료되었습니다!")
            messagebox.showinfo("완료", "모든 캡처 작업이 완료되었습니다!")
            
        except Exception as e:
            self.log(f"치명적 에러 발생: {e}")
        finally:
            self.start_btn.config(state="normal", text="▶ 자동 캡처 시작")

if __name__ == "__main__":
    root = tk.Tk()
    app = AutoCaptureApp(root)
    root.mainloop()
