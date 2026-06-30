import os
import time
import subprocess
import threading
import ctypes
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except:
        pass
import pyautogui
pyautogui.FAILSAFE = False
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from pathlib import Path
import json

class SnippingToolOverlay:
    def __init__(self, parent, callback):
        self.parent = parent
        self.callback = callback
        self.top = tk.Toplevel(parent)
        self.top.attributes('-alpha', 0.4)
        self.top.attributes('-fullscreen', True)
        self.top.attributes('-topmost', True)
        self.top.configure(background='black')
        self.top.config(cursor="cross")
        
        self.canvas = tk.Canvas(self.top, cursor="cross", bg="black")
        self.canvas.pack(fill="both", expand=True)
        
        self.start_x = None
        self.start_y = None
        self.rect = None
        
        self.canvas.bind("<ButtonPress-1>", self.on_button_press)
        self.canvas.bind("<B1-Motion>", self.on_move_press)
        self.canvas.bind("<ButtonRelease-1>", self.on_button_release)
        self.top.bind("<Escape>", lambda e: self.top.destroy())
        
    def on_button_press(self, event):
        self.start_x = event.x
        self.start_y = event.y
        self.rect = self.canvas.create_rectangle(self.start_x, self.start_y, 1, 1, outline='red', width=3, fill="white")
        
    def on_move_press(self, event):
        curX, curY = (event.x, event.y)
        self.canvas.coords(self.rect, self.start_x, self.start_y, curX, curY)
        
    def on_button_release(self, event):
        end_x, end_y = (event.x, event.y)
        x1 = min(self.start_x, end_x)
        y1 = min(self.start_y, end_y)
        w = abs(end_x - self.start_x)
        h = abs(end_y - self.start_y)
        
        # 실제 화면의 절대 좌표 구하기 (다중 모니터 대응)
        root_x = self.top.winfo_rootx()
        root_y = self.top.winfo_rooty()
        abs_x1 = root_x + x1
        abs_y1 = root_y + y1
        
        self.top.destroy()
        if w > 10 and h > 10:
            self.callback([abs_x1, abs_y1, w, h])

class AutoCaptureApp:
    def __init__(self, root):
        self.root = root
        self.root.title("OmniPC Auto Capture")
        self.root.geometry("550x850")
        
        # 설정 파일 경로 (현재 폴더)
        self.config_file = Path(__file__).parent / "auto_capture_config.json"
        
        # 기본 설정값
        self.config = {
            "data_dir": r"C:\Users\jjch2\Desktop\data",
            "capture_dir": r"C:\Users\jjch2\Desktop\captuer",
            "shortcut": r"C:\Users\jjch2\Desktop\OmniPC 6.0.lnk",
            "delay": 12,
            "use_click": True,
            "click1_x": 0, "click1_y": 0,
            "click2_x": 0, "click2_y": 0,
            "click3_x": 0, "click3_y": 0,
            "use_region": False,
            "region": [0, 0, 0, 0] # x, y, w, h
        }
        self.load_config()
        self.create_widgets()
        
        # Ensure config is saved when window is closed
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
    def on_closing(self):
        self.save_config()
        self.root.destroy()
        
    def load_config(self):
        if self.config_file.exists():
            try:
                with open(self.config_file, "r", encoding="utf-8") as f:
                    loaded = json.load(f)
                    self.config.update(loaded)
            except:
                pass
                
    def save_config(self):
        # 현재 UI의 값들을 config에 업데이트 후 저장
        self.config["data_dir"] = self.data_var.get()
        self.config["capture_dir"] = self.capture_var.get()
        self.config["shortcut"] = self.shortcut_var.get()
        self.config["delay"] = self.delay_var.get()
        self.config["use_click"] = self.use_click_var.get()
        self.config["click1_x"] = self.c1_x.get()
        self.config["click1_y"] = self.c1_y.get()
        self.config["click2_x"] = self.c2_x.get()
        self.config["click2_y"] = self.c2_y.get()
        self.config["click3_x"] = self.c3_x.get()
        self.config["click3_y"] = self.c3_y.get()
        self.config["use_region"] = self.use_region_var.get()
        try:
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(self.config, f, indent=4)
        except Exception as e:
            print("설정 저장 실패:", e)

    def create_widgets(self):
        # --- Data Directory ---
        tk.Label(self.root, text="데이터 폴더 (.opd 파일 위치):").pack(anchor="w", padx=10, pady=(10, 0))
        frame1 = tk.Frame(self.root)
        frame1.pack(fill="x", padx=10)
        self.data_var = tk.StringVar(value=self.config["data_dir"])
        tk.Entry(frame1, textvariable=self.data_var).pack(side="left", fill="x", expand=True, padx=(0, 5))
        tk.Button(frame1, text="찾아보기", command=lambda: self.browse_dir(self.data_var)).pack(side="right")
        
        # --- Capture Directory ---
        tk.Label(self.root, text="캡처 저장 폴더:").pack(anchor="w", padx=10, pady=(10, 0))
        frame2 = tk.Frame(self.root)
        frame2.pack(fill="x", padx=10)
        self.capture_var = tk.StringVar(value=self.config["capture_dir"])
        tk.Entry(frame2, textvariable=self.capture_var).pack(side="left", fill="x", expand=True, padx=(0, 5))
        tk.Button(frame2, text="찾아보기", command=lambda: self.browse_dir(self.capture_var)).pack(side="right")
        
        # --- Shortcut Path ---
        tk.Label(self.root, text="OmniPC 바로가기 경로:").pack(anchor="w", padx=10, pady=(10, 0))
        frame3 = tk.Frame(self.root)
        frame3.pack(fill="x", padx=10)
        self.shortcut_var = tk.StringVar(value=self.config["shortcut"])
        tk.Entry(frame3, textvariable=self.shortcut_var).pack(side="left", fill="x", expand=True, padx=(0, 5))
        tk.Button(frame3, text="찾아보기", command=lambda: self.browse_file(self.shortcut_var)).pack(side="right")
        
        # --- Delay Time ---
        frame4 = tk.Frame(self.root)
        frame4.pack(fill="x", padx=10, pady=10)
        tk.Label(frame4, text="대기 시간(초): ").pack(side="left")
        self.delay_var = tk.IntVar(value=self.config["delay"])
        tk.Entry(frame4, textvariable=self.delay_var, width=5).pack(side="left")
        tk.Label(frame4, text=" (파일 열고 캡처할 때까지 대기)").pack(side="left")
        
        # --- Capture Region Settings Frame ---
        rf = tk.LabelFrame(self.root, text="캡처 영역 설정")
        rf.pack(fill="x", padx=10, pady=5)
        
        self.use_region_var = tk.BooleanVar(value=self.config.get("use_region", False))
        tk.Checkbutton(rf, text="전체 화면 대신 지정된 영역만 캡처", variable=self.use_region_var).grid(row=0, column=0, columnspan=2, sticky="w", padx=5)
        
        tk.Button(rf, text="영역 드래그로 선택하기", command=self.start_region_select, bg="#ffdddd").grid(row=1, column=0, padx=10, pady=5)
        self.region_label = tk.Label(rf, text=self.format_region_text())
        self.region_label.grid(row=1, column=1, sticky="w")
        
        # --- Click Settings Frame ---
        lf = tk.LabelFrame(self.root, text="자동 클릭 위치 설정 (캡처 전 메뉴 조작)")
        lf.pack(fill="x", padx=10, pady=5)
        
        self.use_click_var = tk.BooleanVar(value=self.config["use_click"])
        tk.Checkbutton(lf, text="캡처 전 자동 클릭 기능 사용", variable=self.use_click_var).grid(row=0, column=0, columnspan=4, sticky="w", padx=5)
        
        # Click 1
        tk.Label(lf, text="1. Single/Multiple 토글:").grid(row=1, column=0, sticky="e", padx=5)
        self.c1_x = tk.IntVar(value=self.config["click1_x"])
        self.c1_y = tk.IntVar(value=self.config["click1_y"])
        tk.Entry(lf, textvariable=self.c1_x, width=5).grid(row=1, column=1)
        tk.Entry(lf, textvariable=self.c1_y, width=5).grid(row=1, column=2)
        tk.Button(lf, text="위치 지정(7초)", command=lambda: self.get_pos(self.c1_x, self.c1_y)).grid(row=1, column=3, padx=5, pady=2)
        
        # Click 2
        tk.Label(lf, text="2. Layouts 메뉴 클릭:").grid(row=2, column=0, sticky="e", padx=5)
        self.c2_x = tk.IntVar(value=self.config["click2_x"])
        self.c2_y = tk.IntVar(value=self.config["click2_y"])
        tk.Entry(lf, textvariable=self.c2_x, width=5).grid(row=2, column=1)
        tk.Entry(lf, textvariable=self.c2_y, width=5).grid(row=2, column=2)
        tk.Button(lf, text="위치 지정(7초)", command=lambda: self.get_pos(self.c2_x, self.c2_y)).grid(row=2, column=3, padx=5, pady=2)
        
        # Click 3
        tk.Label(lf, text="3. A-C-S (PA) 항목 클릭:").grid(row=3, column=0, sticky="e", padx=5)
        self.c3_x = tk.IntVar(value=self.config["click3_x"])
        self.c3_y = tk.IntVar(value=self.config["click3_y"])
        tk.Entry(lf, textvariable=self.c3_x, width=5).grid(row=3, column=1)
        tk.Entry(lf, textvariable=self.c3_y, width=5).grid(row=3, column=2)
        tk.Button(lf, text="위치 지정(7초)", command=lambda: self.get_pos(self.c3_x, self.c3_y)).grid(row=3, column=3, padx=5, pady=2)
        
        # --- Start Button ---
        self.start_btn = tk.Button(self.root, text="▶ 자동 캡처 시작", bg="lightblue", font=("Arial", 11, "bold"), command=self.start_capture_thread)
        self.start_btn.pack(pady=10, fill="x", padx=50)
        
        # --- Log Text ---
        tk.Label(self.root, text="진행 상황:").pack(anchor="w", padx=10)
        self.log_text = scrolledtext.ScrolledText(self.root, height=10)
        self.log_text.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
    def format_region_text(self):
        r = self.config.get("region", [0,0,0,0])
        if r == [0,0,0,0]:
            return "지정되지 않음 (전체 화면)"
        return f"현재 영역: {r[2]}x{r[3]} (시작: {r[0]},{r[1]})"
        
    def start_region_select(self):
        messagebox.showinfo("영역 지정", "확인을 누르시면 화면 전체가 까맣게 변합니다.\n마우스로 화면에서 캡처할 영역을 주욱 드래그해 주세요!\n(취소하려면 키보드의 Esc 키를 누르세요)")
        SnippingToolOverlay(self.root, self.on_region_selected)
        
    def on_region_selected(self, rect):
        self.config["region"] = rect
        self.use_region_var.set(True)
        self.region_label.config(text=self.format_region_text())
        self.save_config()
        self.log(f"✅ 캡처 영역이 지정되었습니다! {self.format_region_text()}")

    def browse_dir(self, var):
        folder = filedialog.askdirectory(initialdir=var.get())
        if folder:
            var.set(folder)
            self.save_config()
            
    def browse_file(self, var):
        file = filedialog.askopenfilename(initialdir=os.path.dirname(var.get()), filetypes=[("Shortcut", "*.lnk"), ("Executable", "*.exe"), ("All files", "*.*")])
        if file:
            var.set(file)
            self.save_config()
            
    def log(self, msg):
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.root.update()

    def get_pos(self, var_x, var_y):
        self.start_btn.config(state="disabled")
        self.log("\n[위치 저장] 7초 뒤 현재 마우스 위치가 저장됩니다.")
        self.log("여유있게 마우스를 원하는 버튼 위로 옮겨두세요!")
        for i in range(7, 0, -1):
            self.log(f"{i}초 전...")
            self.root.update()
            time.sleep(1)
        x, y = pyautogui.position()
        var_x.set(x)
        var_y.set(y)
        self.log(f"✅ 위치 저장 완료: X={x}, Y={y}\n")
        self.start_btn.config(state="normal")
        self.save_config()

    def start_capture_thread(self):
        self.save_config()
        self.start_btn.config(state="disabled", text="진행 중...")
        self.log_text.delete(1.0, tk.END)
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
            nde_files = list(data_dir.glob("*.nde"))
            all_files = opd_files + nde_files
            self.log(f"총 {len(all_files)}개의 데이터 파일(.opd {len(opd_files)}개, .nde {len(nde_files)}개)을 찾았습니다.\n")
            
            if len(all_files) == 0:
                self.log("작업을 취소합니다. 폴더 내에 처리할 데이터 파일이 없습니다.")
                return

            for idx, target_file in enumerate(all_files, 1):
                self.log(f"[{idx}/{len(all_files)}] {target_file.name} 처리 중...")
                
                try:
                    # 1. 프로그램 실행 (경로에 공백이 있어도 안전하게 실행되도록 start 명령어 사용)
                    os.system(f'start "" "{shortcut_path}" "{target_file}"')
                    
                    # 2. 대기 (카운트다운 로그 표시)
                    for i in range(delay, 0, -1):
                        self.log(f"  -> 화면 로딩 대기 중... {i}초 남음")
                        time.sleep(1)
                        
                    # 3. 자동 클릭 수행 (체크된 경우이고 .opd 파일일 때만)
                    is_opd = target_file.suffix.lower() == '.opd'
                    
                    if self.use_click_var.get():
                        if is_opd:
                            self.log("  -> 창을 최대화하고 설정된 메뉴 위치를 클릭합니다 (.opd 파일)")
                            
                            # 창 최대화 보장
                            try:
                                import pygetwindow as gw
                                windows = gw.getWindowsWithTitle('OmniPC')
                                for w in windows:
                                    if w.title == 'OmniPC' or 'OmniPC' in w.title:
                                        if not w.isMaximized:
                                            w.maximize()
                                        w.activate()
                                time.sleep(1.0)
                            except Exception as e:
                                self.log(f"  -> 창 상태 변경 실패 (무시됨): {e}")

                            # 클릭 과정이 눈에 보이도록 마우스를 부드럽게 이동
                            pyautogui.moveTo(self.c1_x.get(), self.c1_y.get(), duration=0.5)
                            pyautogui.click()
                            time.sleep(1.0)
                            
                            pyautogui.moveTo(self.c2_x.get(), self.c2_y.get(), duration=0.5)
                            pyautogui.click()
                            time.sleep(1.0)
                            
                            pyautogui.moveTo(self.c3_x.get(), self.c3_y.get(), duration=0.5)
                            pyautogui.click()
                            time.sleep(1.0)
                        else:
                            self.log("  -> .nde 파일이므로 자동 클릭 과정을 생략합니다.")
                        
                    # 4. 캡처
                    screenshot_path = capture_dir / f"{target_file.stem}_capture.png"
                    
                    if self.use_region_var.get() and self.config.get("region") != [0,0,0,0]:
                        reg = self.config["region"]
                        pyautogui.screenshot(str(screenshot_path), region=(reg[0], reg[1], reg[2], reg[3]))
                        self.log(f"  -> 📸 지정된 영역 캡처 완료: {screenshot_path.name}")
                    else:
                        pyautogui.screenshot(str(screenshot_path))
                        self.log(f"  -> 📸 전체 화면 캡처 완료: {screenshot_path.name}")
                    
                except Exception as e:
                    self.log(f"  -> ❌ 에러 발생: {e}")
                    
                finally:
                    # 5. 프로그램 강제 종료
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
