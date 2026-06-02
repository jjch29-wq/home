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
        
        # 결함 시뮬레이터 섹션
        defect_title = ctk.CTkLabel(self.input_frame, text="결함 시뮬레이터", font=ctk.CTkFont(size=18, weight="bold"))
        defect_title.pack(pady=(10, 10))
        
        self.create_input_field("결함 시작 깊이(Z_start) [mm]:", "defect_start_depth", "5.0")
        self.create_input_field("결함 끝 깊이(Z_end) [mm]:", "defect_end_depth", "9.0")
        self.create_input_field("결함 X 오프셋 [mm]:", "defect_x_offset", "0.0")
        self.create_input_field("결함 폭(Width) [mm]:", "defect_width", "2.0")
        self.create_input_field("결함 회전각(Angle) [°]:", "defect_angle", "0.0")

        # 마우스 조작 안내
        hint_lbl = ctk.CTkLabel(self.input_frame, text="*좌클릭 드래그: 이동 | 우클릭 드래그: 회전*", text_color="orange", font=ctk.CTkFont(size=12))
        hint_lbl.pack(pady=(0, 5))

        # 결함 모양 콤보박스
        shape_frame = ctk.CTkFrame(self.input_frame, fg_color="transparent")
        shape_frame.pack(fill="x", padx=20, pady=8)
        ctk.CTkLabel(shape_frame, text="결함 형상:", width=150, anchor="w").pack(side="left")
        self.shape_var = ctk.StringVar(value="타원형(Ellipse)")
        self.shape_menu = ctk.CTkOptionMenu(shape_frame, values=["원형(Circle)", "타원형(Ellipse)", "사각형(Rectangle)", "선(Line)"], variable=self.shape_var, width=120)
        self.shape_menu.pack(side="right")

        # 결함 발생 방향 콤보박스
        side_frame = ctk.CTkFrame(self.input_frame, fg_color="transparent")
        side_frame.pack(fill="x", padx=20, pady=8)
        ctk.CTkLabel(side_frame, text="결함 방향:", width=150, anchor="w").pack(side="left")
        self.side_var = ctk.StringVar(value="우측(Right)")
        self.side_menu = ctk.CTkOptionMenu(side_frame, values=["우측(Right)", "좌측(Left)", "양측(Both)"], variable=self.side_var, width=120)
        self.side_menu.pack(side="right")
        
        self.defect_btn = ctk.CTkButton(self.input_frame, text="결함 위치 및 형상 적용", command=self.calculate_defect, fg_color="#28a745", hover_color="#218838", height=35)
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
        
        # 드래그 이벤트 관련 변수
        self.dragging = False
        self.drag_start_x = None
        self.drag_start_z = None
        
        # 캔버스 이벤트 바인딩
        self.canvas.mpl_connect('button_press_event', self.on_press)
        self.canvas.mpl_connect('motion_notify_event', self.on_motion)
        self.canvas.mpl_connect('button_release_event', self.on_release)
        
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
            
        # 결함 위치가 입력되어 있다면 그리기
        try:
            defect_start_z_input = float(self.entries["defect_start_depth"].get())
            defect_end_z_input = float(self.entries["defect_end_depth"].get())
            defect_x_offset = float(self.entries["defect_x_offset"].get())
            defect_width = float(self.entries["defect_width"].get())
            
            try:
                defect_angle_offset = float(self.entries["defect_angle"].get())
            except ValueError:
                defect_angle_offset = 0.0
                
            shape = self.shape_var.get() if hasattr(self, 'shape_var') else "선(Line)"
            side = self.side_var.get() if hasattr(self, 'side_var') else "우측(Right)"
            
            if 0 <= defect_start_z_input <= t and 0 <= defect_end_z_input <= t:
                # 시작점 기본 X 계산
                if defect_start_z_input >= z_root_top:
                    base_start_x = r_g
                else:
                    base_start_x = r_g + ((z_root_top - defect_start_z_input) * math.tan(bevel_angle_rad))
                
                # 끝점 기본 X 계산
                if defect_end_z_input >= z_root_top:
                    base_end_x = r_g
                else:
                    base_end_x = r_g + ((z_root_top - defect_end_z_input) * math.tan(bevel_angle_rad))
                    
                dx_base = base_end_x - base_start_x
                dz_base = defect_end_z_input - defect_start_z_input
                length = math.hypot(dx_base, dz_base)
                if length == 0: length = 0.1
                
                # 중심점 및 회전 계산
                center_z = (defect_start_z_input + defect_end_z_input) / 2
                if center_z >= z_root_top:
                    base_center_x = r_g
                else:
                    base_center_x = r_g + ((z_root_top - center_z) * math.tan(bevel_angle_rad))
                    
                center_x_right = base_center_x + defect_x_offset
                center_x_left = -center_x_right
                
                base_angle_rad = math.atan2(dz_base, dx_base)
                final_angle_deg = math.degrees(base_angle_rad) + defect_angle_offset
                final_angle_rad = math.radians(final_angle_deg)
                
                final_angle_deg_l = 180 - final_angle_deg
                final_angle_rad_l = math.radians(final_angle_deg_l)
                
                hx = math.cos(final_angle_rad) * length / 2
                hz = math.sin(final_angle_rad) * length / 2
                p_start_r = (center_x_right - hx, center_z - hz)
                p_end_r = (center_x_right + hx, center_z + hz)
                
                hx_l = math.cos(final_angle_rad_l) * length / 2
                hz_l = math.sin(final_angle_rad_l) * length / 2
                p_start_l = (center_x_left - hx_l, center_z - hz_l)
                p_end_l = (center_x_left + hx_l, center_z + hz_l)
                
                if shape == "선(Line)":
                    if side in ["우측(Right)", "양측(Both)"]:
                        self.ax.plot([p_start_r[0], p_end_r[0]], [p_start_r[1], p_end_r[1]], 'r-', linewidth=4, label='Defect')
                        self.ax.plot(p_start_r[0], p_start_r[1], 'y*', markersize=10)
                        self.ax.plot(p_end_r[0], p_end_r[1], 'y*', markersize=10)
                    if side in ["좌측(Left)", "양측(Both)"]:
                        lbl = 'Defect' if side == "좌측(Left)" else None
                        self.ax.plot([p_start_l[0], p_end_l[0]], [p_start_l[1], p_end_l[1]], 'r-', linewidth=4, label=lbl)
                        self.ax.plot(p_start_l[0], p_start_l[1], 'y*', markersize=10)
                        self.ax.plot(p_end_l[0], p_end_l[1], 'y*', markersize=10)
                elif shape == "원형(Circle)":
                    if side in ["우측(Right)", "양측(Both)"]:
                        circ_right = patches.Ellipse((center_x_right, center_z), width=length, height=length, angle=0, 
                                                    edgecolor='red', facecolor='red', alpha=0.7, label='Defect')
                        self.ax.add_patch(circ_right)
                    if side in ["좌측(Left)", "양측(Both)"]:
                        lbl = 'Defect' if side == "좌측(Left)" else None
                        circ_left = patches.Ellipse((center_x_left, center_z), width=length, height=length, angle=0, 
                                                   edgecolor='red', facecolor='red', alpha=0.7, label=lbl)
                        self.ax.add_patch(circ_left)
                elif shape == "타원형(Ellipse)":
                    if side in ["우측(Right)", "양측(Both)"]:
                        ell_right = patches.Ellipse((center_x_right, center_z), width=length, height=defect_width, angle=final_angle_deg, 
                                                    edgecolor='red', facecolor='red', alpha=0.7, label='Defect')
                        self.ax.add_patch(ell_right)
                    if side in ["좌측(Left)", "양측(Both)"]:
                        lbl = 'Defect' if side == "좌측(Left)" else None
                        ell_left = patches.Ellipse((center_x_left, center_z), width=length, height=defect_width, angle=final_angle_deg_l, 
                                                   edgecolor='red', facecolor='red', alpha=0.7, label=lbl)
                        self.ax.add_patch(ell_left)
                elif shape == "사각형(Rectangle)":
                    dx_r = p_end_r[0] - p_start_r[0]
                    dz_r = p_end_r[1] - p_start_r[1]
                    ux_r, uz_r = dx_r/length, dz_r/length
                    nx_r, nz_r = -uz_r, ux_r
                    hw = defect_width / 2
                    
                    p1_r = (p_start_r[0] + nx_r*hw, p_start_r[1] + nz_r*hw)
                    p2_r = (p_start_r[0] - nx_r*hw, p_start_r[1] - nz_r*hw)
                    p3_r = (p_end_r[0] - nx_r*hw, p_end_r[1] - nz_r*hw)
                    p4_r = (p_end_r[0] + nx_r*hw, p_end_r[1] + nz_r*hw)
                    
                    dx_l = p_end_l[0] - p_start_l[0]
                    dz_l = p_end_l[1] - p_start_l[1]
                    ux_l, uz_l = dx_l/length, dz_l/length
                    nx_l, nz_l = -uz_l, ux_l
                    
                    p1_l = (p_start_l[0] + nx_l*hw, p_start_l[1] + nz_l*hw)
                    p2_l = (p_start_l[0] - nx_l*hw, p_start_l[1] - nz_l*hw)
                    p3_l = (p_end_l[0] - nx_l*hw, p_end_l[1] - nz_l*hw)
                    p4_l = (p_end_l[0] + nx_l*hw, p_end_l[1] + nz_l*hw)
                    
                    if side in ["우측(Right)", "양측(Both)"]:
                        rect_right = patches.Polygon([p1_r, p2_r, p3_r, p4_r], closed=True, edgecolor='red', facecolor='red', alpha=0.7, label='Defect')
                        self.ax.add_patch(rect_right)
                    if side in ["좌측(Left)", "양측(Both)"]:
                        lbl = 'Defect' if side == "좌측(Left)" else None
                        rect_left = patches.Polygon([p1_l, p2_l, p3_l, p4_l], closed=True, edgecolor='red', facecolor='red', alpha=0.7, label=lbl)
                        self.ax.add_patch(rect_left)
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
            defect_x_offset = float(self.entries["defect_x_offset"].get())
            defect_width = float(self.entries["defect_width"].get())
        except ValueError:
            self.show_result("결함 관련 입력값에는 숫자를 입력해주세요.")
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
        
        # 시작점 기본 X 계산
        if defect_start_z >= z_root_top:
            base_start_x = r_g
        else:
            base_start_x = r_g + ((z_root_top - defect_start_z) * math.tan(bevel_angle_rad))
            
        # 끝점 기본 X 계산
        if defect_end_z >= z_root_top:
            base_end_x = r_g
        else:
            base_end_x = r_g + ((z_root_top - defect_end_z) * math.tan(bevel_angle_rad))
            
        try:
            defect_angle_offset = float(self.entries["defect_angle"].get())
        except ValueError:
            defect_angle_offset = 0.0
            
        dx_base = base_end_x - base_start_x
        dz_base = defect_end_z - defect_start_z
        length = math.hypot(dx_base, dz_base)
        if length == 0: length = 0.1
        
        center_z = (defect_start_z + defect_end_z) / 2
        if center_z >= z_root_top:
            base_center_x = r_g
        else:
            base_center_x = r_g + ((z_root_top - center_z) * math.tan(bevel_angle_rad))
            
        center_x_r = base_center_x + defect_x_offset
        
        base_angle_rad = math.atan2(dz_base, dx_base)
        final_angle_deg = math.degrees(base_angle_rad) + defect_angle_offset
        final_angle_rad = math.radians(final_angle_deg)
        
        final_angle_deg_l = 180 - final_angle_deg
        final_angle_rad_l = math.radians(final_angle_deg_l)
        
        hx = math.cos(final_angle_rad) * length / 2
        hz = math.sin(final_angle_rad) * length / 2
        start_x_r = center_x_r - hx
        end_x_r = center_x_r + hx
        
        hx_l = math.cos(final_angle_rad_l) * length / 2
        hz_l = math.sin(final_angle_rad_l) * length / 2
        start_x_l = -center_x_r - hx_l
        end_x_l = -center_x_r + hx_l
        
        shape = self.shape_var.get() if hasattr(self, 'shape_var') else "선(Line)"
        side = self.side_var.get() if hasattr(self, 'side_var') else "우측(Right)"
            
        msg = f"[{shape} 결함 구간(원래 깊이): {defect_start_z:.1f}mm ~ {defect_end_z:.1f}mm]\n"
        msg += f"▶ 방향: {side} | 오프셋: {defect_x_offset:.2f} mm | 폭: {defect_width:.2f} mm | 회전: {defect_angle_offset:.1f}°\n"
        if side in ["우측(Right)", "양측(Both)"]:
            msg += f"▶ 우측 윗단 X, Z: {start_x_r:.2f}, {center_z - hz:.2f} | 아랫단 X, Z: {end_x_r:.2f}, {center_z + hz:.2f}\n"
        if side in ["좌측(Left)", "양측(Both)"]:
            msg += f"▶ 좌측 윗단 X, Z: {start_x_l:.2f}, {center_z - hz_l:.2f} | 아랫단 X, Z: {end_x_l:.2f}, {center_z + hz_l:.2f}\n"
        msg += f"(투영 높이(Z-extent): {abs(hz*2):.1f} mm, 길이: {length:.1f} mm)"
            
        # 계산 후 그래프에 결함 위치(별표) 업데이트 (먼저 호출해야 덮어쓰지 않음)
        self.update_plot()
        
        # 그래프 업데이트가 끝난 뒤에 결함 계산 결과 메시지를 텍스트 박스에 표시
        self.show_result(msg)
        
    def show_result(self, text):
        self.result_box.configure(state="normal")
        self.result_box.delete("0.0", "end")
        self.result_box.insert("0.0", text)
        self.result_box.configure(state="disabled")

    def on_press(self, event):
        if event.inaxes != self.ax: return
        
        try:
            self.orig_z_start = float(self.entries["defect_start_depth"].get())
            self.orig_z_end = float(self.entries["defect_end_depth"].get())
            self.orig_x_offset = float(self.entries["defect_x_offset"].get())
            try:
                self.orig_angle = float(self.entries["defect_angle"].get())
            except ValueError:
                self.orig_angle = 0.0
        except ValueError:
            return
            
        p = self.get_params()
        if not p: return
        
        t = p["thickness"]
        r_f = p["root_face"]
        r_g = p["root_gap_half"]
        ang = p["bevel_angle_deg"]
        z_root_top = t - r_f
        bevel_angle_rad = math.radians(ang)
        
        center_z = (self.orig_z_start + self.orig_z_end) / 2
        
        if center_z >= z_root_top:
            base_center_x = r_g
        else:
            base_center_x = r_g + ((z_root_top - center_z) * math.tan(bevel_angle_rad))
            
        side = self.side_var.get() if hasattr(self, 'side_var') else "우측(Right)"
        right_cx = base_center_x + self.orig_x_offset
        left_cx = -(base_center_x + self.orig_x_offset)
        
        cx, cz = event.xdata, event.ydata
        if cx is None or cz is None: return
        
        try:
            defect_width = float(self.entries["defect_width"].get())
        except ValueError:
            defect_width = 2.0
            
        dz = abs(cz - center_z)
        height_half = max(2.0, abs(self.orig_z_end - self.orig_z_start) / 2)
        width_margin = max(2.0, defect_width * 1.5)
        
        hit = False
        self.drag_side = None
        self.drag_mode = None
        
        if side in ["우측(Right)", "양측(Both)"]:
            dx = abs(cx - right_cx)
            if dx <= width_margin and dz <= height_half:
                hit = True
                self.drag_side = "우측(Right)"
                
        if not hit and side in ["좌측(Left)", "양측(Both)"]:
            dx = abs(cx - left_cx)
            if dx <= width_margin and dz <= height_half:
                hit = True
                self.drag_side = "좌측(Left)"
                
        if hit:
            self.dragging = True
            self.drag_start_x = cx
            self.drag_start_z = cz
            
            if self.drag_side == "우측(Right)":
                self.orig_cx = right_cx
            else:
                self.orig_cx = left_cx
                
            self.orig_cz = center_z
            self.mouse_offset_x = cx - self.orig_cx
            self.mouse_offset_z = cz - self.orig_cz
            
            if event.button == 1: # 좌클릭
                self.drag_mode = "translate"
            elif event.button == 3: # 우클릭
                self.drag_mode = "rotate"
                self.start_mouse_angle = math.degrees(math.atan2(cz - self.orig_cz, cx - self.orig_cx))
            else:
                self.dragging = False
                return
            
            # 양측(Both) 모드일 때 하나를 클릭해 드래그하면, 해당 결함만 움직이도록 모드를 전환
            if side == "양측(Both)":
                self.side_var.set(self.drag_side)

    def on_motion(self, event):
        if not getattr(self, 'dragging', False) or event.inaxes != self.ax: return
        if event.xdata is None or event.ydata is None: return
        
        if self.drag_mode == "translate":
            desired_cz = event.ydata - self.mouse_offset_z
            desired_cx = event.xdata - self.mouse_offset_x
            
            dz = desired_cz - self.orig_cz
            
            new_z_start = self.orig_z_start + dz
            new_z_end = self.orig_z_end + dz
            
            # 클램핑 (0과 thickness 사이)
            p = self.get_params()
            if p:
                t = p["thickness"]
                if new_z_start < 0:
                    shift = -new_z_start
                    new_z_start += shift
                    new_z_end += shift
                    desired_cz += shift
                elif new_z_end > t:
                    shift = new_z_end - t
                    new_z_start -= shift
                    new_z_end -= shift
                    desired_cz -= shift
                    
            # 새로운 Z 높이에서의 기준 X값 계산
            new_center_z = (new_z_start + new_z_end) / 2
            r_f = p["root_face"]
            r_g = p["root_gap_half"]
            ang = p["bevel_angle_deg"]
            z_root_top = t - r_f
            bevel_angle_rad = math.radians(ang)
            
            if new_center_z >= z_root_top:
                new_base_center_x = r_g
            else:
                new_base_center_x = r_g + ((z_root_top - new_center_z) * math.tan(bevel_angle_rad))
                
            # 드래그한 위치(desired_cx)에 맞춰 X 오프셋 역산
            if self.drag_side == "우측(Right)":
                new_x_offset = desired_cx - new_base_center_x
            else:
                new_x_offset = (-desired_cx) - new_base_center_x
                
            self.entries["defect_start_depth"].delete(0, "end")
            self.entries["defect_start_depth"].insert(0, f"{new_z_start:.2f}")
            
            self.entries["defect_end_depth"].delete(0, "end")
            self.entries["defect_end_depth"].insert(0, f"{new_z_end:.2f}")
            
            self.entries["defect_x_offset"].delete(0, "end")
            self.entries["defect_x_offset"].insert(0, f"{new_x_offset:.2f}")
            
        elif self.drag_mode == "rotate":
            current_mouse_angle = math.degrees(math.atan2(event.ydata - self.orig_cz, event.xdata - self.orig_cx))
            angle_diff = current_mouse_angle - self.start_mouse_angle
            
            if self.drag_side == "좌측(Left)":
                angle_diff = -angle_diff
                
            new_angle = self.orig_angle + angle_diff
            self.entries["defect_angle"].delete(0, "end")
            self.entries["defect_angle"].insert(0, f"{new_angle:.1f}")
            
        self.calculate_defect()

    def on_release(self, event):
        self.dragging = False
        self.drag_start_x = None
        self.drag_start_z = None

if __name__ == "__main__":
    app = App()
    app.mainloop()
