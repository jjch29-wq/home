import customtkinter as ctk
import math
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import warnings

# 불필요한 Matplotlib 경고 숨김
warnings.filterwarnings("ignore", message="Ignoring fixed y limits")
warnings.filterwarnings("ignore", message="Ignoring fixed x limits")

# 테마 설정
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("PAUT 용접부 개선각 위치 계산기")
        self.geometry("1200x700")
        
        # 메인 레이아웃 설정 (1행 2열)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=0) # 좌측 입력부 고정
        self.grid_columnconfigure(1, weight=1) # 우측 그래프부 확장
        
        # ---------------- 좌측 프레임 (입력부) ----------------
        self.input_frame = ctk.CTkFrame(self, width=350, corner_radius=0)
        self.input_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self.input_frame.grid_propagate(False) # 고정 너비
        
        # 타이틀
        title_label = ctk.CTkLabel(self.input_frame, text="용접부 파라미터 설정", font=ctk.CTkFont(size=20, weight="bold"))
        title_label.pack(pady=(20, 20))
        
        # 입력 필드들 생성
        self.entries = {}
        self.create_input_field("전체 두께 (T) [mm]:", "thickness", "15.88")
        self.create_input_field("루트면 높이 [mm]:", "root_face", "1.60")
        self.create_input_field("루트간격 절반 [mm]:", "root_gap_half", "1.50")
        self.create_input_field("개선각 (편각) [°]:", "bevel_angle_deg", "37.5")
        
        # 계산/그리기 버튼
        self.calc_btn = ctk.CTkButton(self.input_frame, text="그래프 그리기 및 갱신", command=self.update_plot, height=40, font=ctk.CTkFont(weight="bold"))
        self.calc_btn.pack(pady=(20, 30), padx=20, fill="x")
        
        # 구분선
        sep = ctk.CTkFrame(self.input_frame, height=2, fg_color="gray50")
        sep.pack(fill="x", padx=20, pady=10)
        
        # 결함 위치 계산기 섹션
        defect_title = ctk.CTkLabel(self.input_frame, text="결함 위치(X) 역산 시뮬레이터", font=ctk.CTkFont(size=18, weight="bold"))
        defect_title.pack(pady=(10, 10))
        
        self.create_input_field("결함 시작 깊이(Z_start) [mm]:", "defect_start_depth", "5.0")
        self.create_input_field("결함 끝 깊이(Z_end) [mm]:", "defect_end_depth", "9.0")
        
        self.defect_btn = ctk.CTkButton(self.input_frame, text="선형 결함 중심거리(X) 계산", command=self.calculate_defect, fg_color="#28a745", hover_color="#218838", height=35)
        self.defect_btn.pack(pady=(10, 10), padx=20, fill="x")
        
        self.result_box = ctk.CTkTextbox(self.input_frame, height=100, font=ctk.CTkFont(size=13))
        self.result_box.pack(pady=(10, 20), padx=20, fill="x")
        self.result_box.insert("0.0", "결과가 여기에 표시됩니다.\n")
        self.result_box.configure(state="disabled")
        
        # ---------------- 우측 프레임 (그래프부) ----------------
        self.plot_frame = ctk.CTkFrame(self)
        self.plot_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
        
        # Matplotlib Figure 준비
        self.fig, self.ax = plt.subplots(figsize=(8, 6))
        self.fig.patch.set_facecolor('#2b2b2b') # 다크모드 배경
        self.ax.set_facecolor('#2b2b2b')
        self.ax.tick_params(colors='white')
        self.ax.xaxis.label.set_color('white')
        self.ax.yaxis.label.set_color('white')
        self.ax.title.set_color('white')
        
        # 캔버스 위젯 생성
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.plot_frame)
        self.canvas_widget = self.canvas.get_tk_widget()
        self.canvas_widget.pack(fill="both", expand=True)
        
        # 초기 그래프 그리기
        self.update_plot()

    def create_input_field(self, label_text, key, default_val):
        frame = ctk.CTkFrame(self.input_frame, fg_color="transparent")
        frame.pack(fill="x", padx=20, pady=8)
        
        lbl = ctk.CTkLabel(frame, text=label_text, width=150, anchor="w")
        lbl.pack(side="left")
        
        entry = ctk.CTkEntry(frame, width=120)
        entry.pack(side="right")
        entry.insert(0, default_val)
        
        self.entries[key] = entry
        
    def get_params(self):
        try:
            return {
                "thickness": float(self.entries["thickness"].get()),
                "root_face": float(self.entries["root_face"].get()),
                "root_gap_half": float(self.entries["root_gap_half"].get()),
                "bevel_angle_deg": float(self.entries["bevel_angle_deg"].get()),
            }
        except ValueError:
            return None

    def update_plot(self):
        p = self.get_params()
        if not p:
            self.show_result("입력값이 올바르지 않습니다. 숫자를 입력해주세요.")
            return
            
        t = p["thickness"]
        r_f = p["root_face"]
        r_g = p["root_gap_half"]
        ang = p["bevel_angle_deg"]
        
        z_top = 0.0
        z_root_top = t - r_f
        z_bottom = t
        
        bevel_angle_rad = math.radians(ang)
        calc_top_bevel_x = r_g + (z_root_top * math.tan(bevel_angle_rad))
        
        points = {
            "Bevel_Top": (calc_top_bevel_x, z_top),
            "Root_Top": (r_g, z_root_top),
            "Root_Bottom": (r_g, z_bottom)
        }
        
        # 그래프 초기화 및 다시 그리기
        self.ax.clear()
        
        self.ax.axvline(0, color='gray', linestyle='-.', linewidth=1, label='Centerline')
        self.ax.axhline(0, color='gray', linestyle='--', linewidth=1)
        self.ax.axhline(t, color='gray', linestyle='--', linewidth=1)
        
        # 우측/좌측 선
        x_right = [points["Bevel_Top"][0], points["Root_Top"][0], points["Root_Bottom"][0]]
        z_right = [points["Bevel_Top"][1], points["Root_Top"][1], points["Root_Bottom"][1]]
        x_left = [-x for x in x_right]
        
        self.ax.plot(x_right, z_right, 'dodgerblue', linewidth=2, label='V-Groove Profile')
        self.ax.plot(x_left, z_right, 'dodgerblue', linewidth=2)
        self.ax.plot([x_left[-1], x_right[-1]], [z_bottom, z_bottom], 'dodgerblue', linewidth=2)
        
        # 모재 핑크색 영역
        bm_right = patches.Polygon([(points["Bevel_Top"][0], 0), (max(20, calc_top_bevel_x+10), 0), 
                                    (max(20, calc_top_bevel_x+10), t), (points["Root_Bottom"][0], t), 
                                    (points["Root_Top"][0], z_root_top)], closed=True, color='pink', alpha=0.3)
        bm_left = patches.Polygon([(-points["Bevel_Top"][0], 0), (-max(20, calc_top_bevel_x+10), 0), 
                                   (-max(20, calc_top_bevel_x+10), t), (-points["Root_Bottom"][0], t), 
                                   (-points["Root_Top"][0], z_root_top)], closed=True, color='pink', alpha=0.3)
        self.ax.add_patch(bm_right)
        self.ax.add_patch(bm_left)
        
        # 좌표 마킹
        for name, (x, z) in points.items():
            self.ax.plot(x, z, 'ro')
            self.ax.text(x + 0.5, z - 0.5 if z > t/2 else z + 0.8, f"({x:.2f}, {z:.2f})", color='white', fontsize=9)
            
        # 결함 위치가 입력되어 있다면 그리기 (선분형 결함)
        try:
            defect_start_z = float(self.entries["defect_start_depth"].get())
            defect_end_z = float(self.entries["defect_end_depth"].get())
            
            if 0 <= defect_start_z <= t and 0 <= defect_end_z <= t:
                # 시작점 X 계산
                if defect_start_z >= z_root_top:
                    defect_start_x = r_g
                else:
                    defect_start_x = r_g + ((z_root_top - defect_start_z) * math.tan(bevel_angle_rad))
                
                # 끝점 X 계산
                if defect_end_z >= z_root_top:
                    defect_end_x = r_g
                else:
                    defect_end_x = r_g + ((z_root_top - defect_end_z) * math.tan(bevel_angle_rad))
                    
                # 선으로 그리기 (빨간색 굵은 선)
                self.ax.plot([defect_start_x, defect_end_x], [defect_start_z, defect_end_z], 'r-', linewidth=4, label='Defect (Line)')
                self.ax.plot([-defect_start_x, -defect_end_x], [defect_start_z, defect_end_z], 'r-', linewidth=4)
                
                # 위아래 끝점 별표 마킹
                self.ax.plot(defect_start_x, defect_start_z, 'y*', markersize=10)
                self.ax.plot(defect_end_x, defect_end_z, 'y*', markersize=10)
                self.ax.plot(-defect_start_x, defect_start_z, 'y*', markersize=10)
                self.ax.plot(-defect_end_x, defect_end_z, 'y*', markersize=10)
        except ValueError:
            pass

        self.ax.invert_yaxis()
        self.ax.set_xlabel('X Position (mm)', color='white')
        self.ax.set_ylabel('Z Depth (mm)', color='white')
        self.ax.set_title('PAUT Single V-Groove Profile', color='white', fontweight='bold')
        self.ax.grid(True, linestyle=':', alpha=0.3)
        
        # 범례 텍스트 색상 수정
        legend = self.ax.legend(loc='upper right')
        for text in legend.get_texts():
            text.set_color("black")
            
        self.ax.set_aspect('equal', adjustable='datalim')
        self.ax.set_xlim(-max(20, calc_top_bevel_x+5), max(20, calc_top_bevel_x+5))
        self.ax.set_ylim(t + 5, -5)
        
        self.canvas.draw()
        
        # 결과 요약 출력
        msg = f"[그래프 갱신 완료]\n상단 개선각 너비(X) 계산 결과: {calc_top_bevel_x:.3f} mm\n루트면 시작 깊이(Z): {z_root_top:.2f} mm"
        self.show_result(msg)

    def calculate_defect(self):
        p = self.get_params()
        if not p:
            self.show_result("도면 파라미터 입력값이 올바르지 않습니다.")
            return
            
        try:
            defect_start_z = float(self.entries["defect_start_depth"].get())
            defect_end_z = float(self.entries["defect_end_depth"].get())
        except ValueError:
            self.show_result("결함 깊이(Z)에는 숫자를 입력해주세요.")
            return
            
        t = p["thickness"]
        r_f = p["root_face"]
        r_g = p["root_gap_half"]
        ang = p["bevel_angle_deg"]
        z_root_top = t - r_f
        
        if defect_start_z < 0 or defect_end_z > t or defect_start_z > defect_end_z:
            self.show_result(f"결함 깊이 범위가 올바르지 않습니다.\n(시작 깊이 <= 끝 깊이 조건 불만족)")
            return
            
        bevel_angle_rad = math.radians(ang)
        
        # 시작점 계산
        if defect_start_z >= z_root_top:
            start_x = r_g
        else:
            start_x = r_g + ((z_root_top - defect_start_z) * math.tan(bevel_angle_rad))
            
        # 끝점 계산
        if defect_end_z >= z_root_top:
            end_x = r_g
        else:
            end_x = r_g + ((z_root_top - defect_end_z) * math.tan(bevel_angle_rad))
            
        msg = f"[선형 결함 구간: {defect_start_z:.1f}mm ~ {defect_end_z:.1f}mm]\n"
        msg += f"▶ 윗단(Upper) X: 중심선에서 {start_x:.2f} mm\n"
        msg += f"▶ 아랫단(Lower) X: 중심선에서 {end_x:.2f} mm\n"
        msg += f"(결함의 높이(Z-extent): {defect_end_z - defect_start_z:.1f} mm)"
            
        # 계산 후 그래프에 결함 위치(별표) 업데이트 (먼저 호출해야 덮어쓰지 않음)
        self.update_plot()
        
        # 그래프 업데이트가 끝난 뒤에 결함 계산 결과 메시지를 텍스트 박스에 표시
        self.show_result(msg)
        
    def show_result(self, text):
        self.result_box.configure(state="normal")
        self.result_box.delete("0.0", "end")
        self.result_box.insert("0.0", text)
        self.result_box.configure(state="disabled")

if __name__ == "__main__":
    app = App()
    app.mainloop()
