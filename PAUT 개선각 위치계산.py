import customtkinter as ctk
import math
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import warnings
import json
from tkinter import filedialog, messagebox

# 불필요한 Matplotlib 경고 숨김
warnings.filterwarnings("ignore", message="Ignoring fixed y limits")
warnings.filterwarnings("ignore", message="Ignoring fixed x limits")
warnings.filterwarnings("ignore", message=".*Tight layout not applied.*")

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
        self.input_frame = ctk.CTkScrollableFrame(self, width=350, corner_radius=0)
        self.input_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        
        # 타이틀
        title_label = ctk.CTkLabel(self.input_frame, text="용접부 파라미터 설정", font=ctk.CTkFont(size=20, weight="bold"))
        title_label.pack(pady=(20, 20))
        
        # 입력 필드들 생성
        self.entries = {}
        self.create_input_field("전체 두께 (T) [mm]:", "thickness", "15.88")
        self.create_input_field("루트면 높이 [mm]:", "root_face", "1.60")
        self.create_input_field("루트간격 절반 [mm]:", "root_gap_half", "1.50")
        self.create_input_field("개선각 (편각) [°]:", "bevel_angle_deg", "37.5")
        self.create_input_field("시편 X축 길이 (스캔) [mm]:", "specimen_length", "320.0")
        self.create_input_field("시편 Y축 폭 (인덱스) [mm]:", "specimen_width", "610.0")
        
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
        self.create_input_field("결함 Y 위치 (중심선 기준) [mm]:", "defect_y_center", "0.0")
        self.create_input_field("결함 폭/직경(W/D) [mm]:", "defect_width", "2.5")
        self.create_input_field("결함 회전각(Angle) [°]:", "defect_angle", "0.0")
        self.create_input_field("결함 X 위치 (스캔) [mm]:", "defect_y_pos", "160.0")
        self.create_input_field("결함 길이(L) [mm]:", "defect_y_length", "10.0")

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
        
        # 표시 화면 콤보박스
        scan_frame = ctk.CTkFrame(self.input_frame, fg_color="transparent")
        scan_frame.pack(fill="x", padx=20, pady=8)
        ctk.CTkLabel(scan_frame, text="표시 화면:", width=150, anchor="w").pack(side="left")
        self.scan_var = ctk.StringVar(value="Front B-Scan")
        self.scan_menu = ctk.CTkOptionMenu(scan_frame, values=["Front B-Scan", "Back B-Scan", "양쪽 화면(Both)"], variable=self.scan_var, width=120)
        self.scan_menu.pack(side="right")
        
        self.create_input_field("라벨 접두사 (ex: SDH, 빈칸가능):", "defect_custom_label", "SDH")
        
        btn_frame = ctk.CTkFrame(self.input_frame, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=(10, 10))
        
        self.add_front_btn = ctk.CTkButton(btn_frame, text="+ Front 뷰 결함", command=self.add_front_defect, width=90)
        self.add_front_btn.pack(side="left", expand=True, padx=(0, 2))
        
        self.add_back_btn = ctk.CTkButton(btn_frame, text="+ Back 뷰 결함", command=self.add_back_defect, width=90, fg_color="#17a2b8", hover_color="#138496")
        self.add_back_btn.pack(side="left", expand=True, padx=(2, 5))
        
        self.del_btn = ctk.CTkButton(btn_frame, text="- 선택 삭제", command=self.delete_defect, width=60, fg_color="#dc3545", hover_color="#c82333")
        self.del_btn.pack(side="right", expand=True, padx=(5, 0))
        
        self.defect_btn = ctk.CTkButton(self.input_frame, text="현재 값으로 결함 속성 적용", command=self.apply_defect_properties, fg_color="#28a745", hover_color="#218838", height=35)
        self.defect_btn.pack(pady=(0, 10), padx=20, fill="x")
        
        self.defects = []
        self.selected_defect_idx = -1
        
        self.result_box = ctk.CTkTextbox(self.input_frame, height=100, font=ctk.CTkFont(size=13))
        self.result_box.pack(pady=(10, 20), padx=20, fill="x")
        self.result_box.insert("0.0", "결과가 여기에 표시됩니다.\n")
        self.result_box.configure(state="disabled")
        
        save_load_frame = ctk.CTkFrame(self.input_frame, fg_color="transparent")
        save_load_frame.pack(fill="x", padx=20, pady=(0, 20))
        
        self.save_btn = ctk.CTkButton(save_load_frame, text="설정 저장", command=self.save_project, width=100)
        self.save_btn.pack(side="left", expand=True, padx=(0, 5))
        
        self.load_btn = ctk.CTkButton(save_load_frame, text="설정 불러오기", command=self.load_project, width=100, fg_color="#17a2b8", hover_color="#138496")
        self.load_btn.pack(side="right", expand=True, padx=(5, 0))
        
        # ---------------- 우측 프레임 (그래프부) ----------------
        self.plot_frame = ctk.CTkFrame(self)
        self.plot_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
        
        # Matplotlib Figure 준비
        self.fig, (self.ax_side_back, self.ax_top, self.ax_side) = plt.subplots(3, 1, figsize=(8, 12), gridspec_kw={'height_ratios': [1, 1.2, 1]})
        self.fig.patch.set_facecolor('#2b2b2b') # 다크모드 배경
        for ax in (self.ax_top, self.ax_side, self.ax_side_back):
            ax.set_facecolor('#2b2b2b')
            ax.tick_params(colors='white')
            ax.xaxis.label.set_color('white')
            ax.yaxis.label.set_color('white')
            ax.title.set_color('white')
        
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
        
        # 초기 그래프 그리기 (결함 하나 추가하면서 자동 갱신됨)
        self.add_defect()
    def create_input_field(self, label_text, key, default_val):
        frame = ctk.CTkFrame(self.input_frame, fg_color="transparent")
        frame.pack(fill="x", padx=20, pady=8)
        
        lbl = ctk.CTkLabel(frame, text=label_text, width=150, anchor="w")
        lbl.pack(side="left")
        
        entry = ctk.CTkEntry(frame, width=120)
        entry.pack(side="right")
        entry.insert(0, default_val)
        
        # 엔터키 누르면 속성 적용되도록 바인딩
        entry.bind("<Return>", lambda e: self.apply_defect_properties() if key.startswith("defect_") else self.update_plot())
        
        self.entries[key] = entry
        
    def get_params(self):
        try:
            return {
                "thickness": float(self.entries["thickness"].get()),
                "root_face": float(self.entries["root_face"].get()),
                "root_gap_half": float(self.entries["root_gap_half"].get()),
                "bevel_angle_deg": float(self.entries["bevel_angle_deg"].get()),
                "specimen_length": float(self.entries["specimen_length"].get()),
                "specimen_width": float(self.entries["specimen_width"].get())
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
        self.ax_side.clear()
        self.ax_side_back.clear()
        self.ax_top.clear()
        
        specimen_width = p["specimen_width"]
        half_width = specimen_width / 2.0
        
        # == Side View (B-Scan) ==
        for ax in (self.ax_side, self.ax_side_back):
            is_back = (ax == self.ax_side_back)
            mult = -1 if is_back else 1
            
            ax.axvline(0, color='gray', linestyle='-.', linewidth=1, label='Centerline')
            ax.axhline(0, color='gray', linestyle='--', linewidth=1)
            ax.axhline(t, color='gray', linestyle='--', linewidth=1)
            
            x_right = [mult * points["Bevel_Top"][0], mult * points["Root_Top"][0], mult * points["Root_Bottom"][0]]
            z_right = [points["Bevel_Top"][1], points["Root_Top"][1], points["Root_Bottom"][1]]
            x_left = [-x for x in x_right]
            
            ax.plot(x_right, z_right, 'dodgerblue', linewidth=2, label='V-Groove Profile')
            ax.plot(x_left, z_right, 'dodgerblue', linewidth=2)
            ax.plot([x_left[-1], x_right[-1]], [z_bottom, z_bottom], 'dodgerblue', linewidth=2)
            
            bm_right = patches.Polygon([(mult * points["Bevel_Top"][0], 0), (mult * half_width, 0), 
                                        (mult * half_width, t), (mult * points["Root_Bottom"][0], t), 
                                        (mult * points["Root_Top"][0], z_root_top)], closed=True, color='pink', alpha=0.3)
            bm_left = patches.Polygon([(-mult * points["Bevel_Top"][0], 0), (-mult * half_width, 0), 
                                       (-mult * half_width, t), (-mult * points["Root_Bottom"][0], t), 
                                       (-mult * points["Root_Top"][0], z_root_top)], closed=True, color='pink', alpha=0.3)
            ax.add_patch(bm_right)
            ax.add_patch(bm_left)
            
            for name, (x, z) in points.items():
                # 하부 텍스트 삭제 및 빨간 점 완전 제거
                if "Root" in name: continue
                
                px = mult * x
                
                # 상부 텍스트(Bevel_Top)만 바깥쪽으로 밀어서 표기합니다 (빨간 점 없음)
                offset_x = 3.0 if px >= 0 else -3.0
                ha = 'left' if px >= 0 else 'right'
                
                offset_z = -2.0
                va = 'bottom'
                
                ax.text(px + offset_x, z + offset_z, f"({x:.1f}, {z:.1f})", color='white', fontsize=9, ha=ha, va=va)
            
        # == Top View (C-Scan) ==
        self.ax_top.axvline(0, color='gray', linestyle='-.', linewidth=1, label='Centerline')
        self.ax_top.axvline(calc_top_bevel_x, color='dodgerblue', linestyle='-', linewidth=2, label='Cap Edge')
        self.ax_top.axvline(-calc_top_bevel_x, color='dodgerblue', linestyle='-', linewidth=2)
        self.ax_top.axvline(r_g, color='lightblue', linestyle='--', linewidth=1, label='Root Edge')
        self.ax_top.axvline(-r_g, color='lightblue', linestyle='--', linewidth=1)
        
        self.ax_top.axvspan(calc_top_bevel_x, half_width, color='pink', alpha=0.3)
        self.ax_top.axvspan(-half_width, -calc_top_bevel_x, color='pink', alpha=0.3)

        # 결함 리스트 그리기
        grp_side_r, grp_side_l, grp_top_r, grp_top_l = {}, {}, {}, {}
        first_group_side_r, first_group_side_l = {}, {}
        sel_side_r_key, sel_side_l_key, sel_top_r_key, sel_top_l_key = None, None, None, None
        for idx, dfct in enumerate(getattr(self, 'defects', [])):
            cnum = chr(0x2460 + idx) if idx < 20 else str(idx + 1)
            shp, w, yl, yp, cx, cz, s = dfct["shape"], round(dfct["width"], 1), round(dfct.get("y_length", 10.0), 1), round(dfct.get("y_pos", 0.0), 1), round(dfct["y_center"], 1), round((dfct["z_start"] + dfct["z_end"])/2, 1), dfct["side"]
            if s in ["우측(Right)", "양측(Both)"]:
                k_sr, k_tr = (shp, w, cx, cz), (shp, w, yl, cx, yp)
                grp_side_r.setdefault(k_sr, []).append(cnum)
                grp_top_r.setdefault(k_tr, []).append(cnum)
                if (shp, w) not in first_group_side_r: first_group_side_r[(shp, w)] = k_sr
                if idx == self.selected_defect_idx: sel_side_r_key, sel_top_r_key = k_sr, k_tr
            if s in ["좌측(Left)", "양측(Both)"]:
                k_sl, k_tl = (shp, w, -cx, cz), (shp, w, yl, -cx, yp)
                grp_side_l.setdefault(k_sl, []).append(cnum)
                grp_top_l.setdefault(k_tl, []).append(cnum)
                if (shp, w) not in first_group_side_l: first_group_side_l[(shp, w)] = k_sl
                if idx == self.selected_defect_idx: sel_side_l_key, sel_top_l_key = k_sl, k_tl

        seen_side_r = set()
        seen_side_l = set()
        seen_top_r = set()
        seen_top_l = set()
        seen_y_pos = set()
        
        dim_count_side_r = {}
        dim_count_side_l = {}
        dim_count_top_len_r = {}
        dim_count_top_len_l = {}
        lbl_count_side_r = {}
        lbl_count_side_l = {}
        lbl_count_top_r = {}
        lbl_count_top_l = {}
        
        for idx, dfct in enumerate(getattr(self, 'defects', [])):
            defect_start_z_input = dfct["z_start"]
            defect_end_z_input = dfct["z_end"]
            defect_y_center = dfct["y_center"]
            defect_width = dfct["width"]
            defect_angle_offset = dfct["angle"]
            shape = dfct["shape"]
            side = dfct["side"]
            is_selected = (idx == self.selected_defect_idx)
            
            color = 'magenta' if is_selected else 'red'
            edge_color = 'yellow' if is_selected else color
            
            if defect_start_z_input >= z_root_top:
                base_start_x = r_g
            else:
                base_start_x = r_g + ((z_root_top - defect_start_z_input) * math.tan(bevel_angle_rad))
            
            if defect_end_z_input >= z_root_top:
                base_end_x = r_g
            else:
                base_end_x = r_g + ((z_root_top - defect_end_z_input) * math.tan(bevel_angle_rad))
                
            dx_base = base_end_x - base_start_x
            dz_base = defect_end_z_input - defect_start_z_input
            length = math.hypot(dx_base, dz_base)
            if length == 0: length = 0.1
            
            center_z = (defect_start_z_input + defect_end_z_input) / 2
                
            center_x_right = defect_y_center
            center_x_left = -defect_y_center
            
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
            
            perp_angle_rad = final_angle_rad + math.pi / 2
            dx_w = math.cos(perp_angle_rad) * defect_width / 2
            dz_w = math.sin(perp_angle_rad) * defect_width / 2
            
            poly_r = [
                (p_start_r[0] + dx_w, p_start_r[1] + dz_w),
                (p_end_r[0] + dx_w, p_end_r[1] + dz_w),
                (p_end_r[0] - dx_w, p_end_r[1] - dz_w),
                (p_start_r[0] - dx_w, p_start_r[1] - dz_w)
            ]
            
            perp_angle_rad_l = final_angle_rad_l + math.pi / 2
            dx_w_l = math.cos(perp_angle_rad_l) * defect_width / 2
            dz_w_l = math.sin(perp_angle_rad_l) * defect_width / 2
            
            poly_l = [
                (p_start_l[0] + dx_w_l, p_start_l[1] + dz_w_l),
                (p_end_l[0] + dx_w_l, p_end_l[1] + dz_w_l),
                (p_end_l[0] - dx_w_l, p_end_l[1] - dz_w_l),
                (p_start_l[0] - dx_w_l, p_start_l[1] - dz_w_l)
            ]
            
            def draw_stars(ax, xs, ys, color):
                ax.plot(xs, ys, marker='*', color=color, markersize=10, linestyle='None')
            
            for ax in (self.ax_side, self.ax_side_back):
                is_front = (ax == self.ax_side)
                mult = 1 if is_front else -1
                
                # Apply mult to all X coordinates for this ax
                p_start_r_m = (mult * p_start_r[0], p_start_r[1])
                p_end_r_m = (mult * p_end_r[0], p_end_r[1])
                p_start_l_m = (mult * p_start_l[0], p_start_l[1])
                p_end_l_m = (mult * p_end_l[0], p_end_l[1])
                
                scan_view = dfct.get("scan_view", "Front B-Scan")
                if is_front and scan_view not in ["Front B-Scan", "양쪽 화면(Both)"]:
                    continue
                if not is_front and scan_view not in ["Back B-Scan", "양쪽 화면(Both)"]:
                    continue
                    
                if shape == "선(Line)":
                    if side in ["우측(Right)", "양측(Both)"]:
                        ax.plot([p_start_r_m[0], p_end_r_m[0]], [p_start_r_m[1], p_end_r_m[1]], color=color, linewidth=4)
                        if is_selected: draw_stars(ax, [p_start_r_m[0], p_end_r_m[0]], [p_start_r_m[1], p_end_r_m[1]], 'yellow')
                    if side in ["좌측(Left)", "양측(Both)"]:
                        ax.plot([p_start_l_m[0], p_end_l_m[0]], [p_start_l_m[1], p_end_l_m[1]], color=color, linewidth=4)
                        if is_selected: draw_stars(ax, [p_start_l_m[0], p_end_l_m[0]], [p_start_l_m[1], p_end_l_m[1]], 'yellow')
                elif shape in ["원형(Circle)", "타원형(Ellipse)"]:
                    if side in ["우측(Right)", "양측(Both)"]:
                        ax.add_patch(patches.Ellipse((mult * center_x_right, center_z), width=length if shape == "타원형(Ellipse)" else defect_width, height=defect_width, angle=(final_angle_deg if mult == 1 else -final_angle_deg) if shape == "타원형(Ellipse)" else 0, edgecolor=edge_color, facecolor=color, alpha=0.7, linewidth=0.5))
                    if side in ["좌측(Left)", "양측(Both)"]:
                        ax.add_patch(patches.Ellipse((mult * center_x_left, center_z), width=length if shape == "타원형(Ellipse)" else defect_width, height=defect_width, angle=(final_angle_deg_l if mult == 1 else -final_angle_deg_l) if shape == "타원형(Ellipse)" else 0, edgecolor=edge_color, facecolor=color, alpha=0.7, linewidth=0.5))
                elif shape == "사각형(Rectangle)":
                    poly_r_m = [(mult * px, pz) for (px, pz) in poly_r]
                    poly_l_m = [(mult * px, pz) for (px, pz) in poly_l]
                    
                    if side in ["우측(Right)", "양측(Both)"]:
                        ax.add_patch(patches.Polygon(poly_r_m, closed=True, edgecolor=edge_color, facecolor=color, alpha=0.7, linewidth=0.5))
                    if side in ["좌측(Left)", "양측(Both)"]:
                        ax.add_patch(patches.Polygon(poly_l_m, closed=True, edgecolor=edge_color, facecolor=color, alpha=0.7, linewidth=0.5))
            y_pos = dfct.get("y_pos", 10.0)
            y_len = dfct.get("y_length", 10.0)
            
            circle_num = chr(0x2460 + idx) if idx < 20 else str(idx + 1)
            circle_num = chr(0x2460 + idx) if idx < 20 else str(idx + 1)
            custom_lbl = dfct.get("custom_label", "").strip()
            
            key_r = (shape, round(defect_width, 1), round(center_x_right, 1), round(center_z, 1))
            key_l = (shape, round(defect_width, 1), round(center_x_left, 1), round(center_z, 1))
            
            cnums_side_r = ",".join(grp_side_r.get(key_r, [circle_num]))
            cnums_side_l = ",".join(grp_side_l.get(key_l, [circle_num]))
            
            rest_r = ""
            rest_l = ""
            
            if shape == "원형(Circle)":
                suffix_r = f" {custom_lbl}{defect_width:.1f}MM" if custom_lbl else f" {defect_width:.1f}MM"
                suffix_l = f" {custom_lbl}{defect_width:.1f}MM" if custom_lbl else f" {defect_width:.1f}MM"
                
                if key_r == first_group_side_r.get((shape, round(defect_width, 1))):
                    text_r = f"{cnums_side_r}{suffix_r}"
                    rest_r = suffix_r
                else:
                    text_r = f"{cnums_side_r}"
                    
                if key_l == first_group_side_l.get((shape, round(defect_width, 1))):
                    text_l = f"{cnums_side_l}{suffix_l}"
                    rest_l = suffix_l
                else:
                    text_l = f"{cnums_side_l}"
            else:
                text_r = f"{cnums_side_r} {custom_lbl}" if custom_lbl else f"{cnums_side_r}"
                rest_r = f" {custom_lbl}" if custom_lbl else ""
                
                text_l = f"{cnums_side_l} {custom_lbl}" if custom_lbl else f"{cnums_side_l}"
                rest_l = f" {custom_lbl}" if custom_lbl else ""
            
            if is_selected:
                extra_r = f"\nY:{center_x_right:.1f} X:{y_pos:.1f}\nd:{center_z:.1f} W:{defect_width:.1f}"
                extra_l = f"\nY:{abs(center_x_left):.1f} X:{y_pos:.1f}\nd:{center_z:.1f} W:{defect_width:.1f}"
                text_r += extra_r
                text_l += extra_l
                if rest_r: rest_r += extra_r
                if rest_l: rest_l += extra_l
                
            font_color = 'yellow' if is_selected else 'white'
            font_weight = 'bold' if is_selected else 'normal'
            
            skip_text_r = False
            skip_text_l = False
            
            if shape in ["원형(Circle)", "타원형(Ellipse)"]:
                actual_top_z = center_z - defect_width/2
                actual_bottom_z = center_z + defect_width/2
                h_val = defect_width
            else:
                actual_top_z = min(defect_start_z_input, defect_end_z_input)
                actual_bottom_z = max(defect_start_z_input, defect_end_z_input)
                h_val = actual_bottom_z - actual_top_z
                if h_val < 0.1:
                    actual_top_z = center_z - defect_width/2
                    actual_bottom_z = center_z + defect_width/2
                    h_val = defect_width
                    
            
            if shape == "원형(Circle)":
                if key_r == sel_side_r_key and not is_selected:
                    skip_text_r = True
                elif key_r in seen_side_r and not is_selected:
                    skip_text_r = True
                else:
                    seen_side_r.add(key_r)
                    
                if key_l == sel_side_l_key and not is_selected:
                    skip_text_l = True
                elif key_l in seen_side_l and not is_selected:
                    skip_text_l = True
                else:
                    seen_side_l.add(key_l)
            
            scan_view = dfct.get("scan_view", "Front B-Scan")
            
            if side in ["우측(Right)", "양측(Both)"] and not skip_text_r:
                offset_pt = len(cnums_side_r) * 15.0 if cnums_side_r else 0
                coord_r = (round(center_x_right, 1), round(center_z, 1))
                lbl_offset_y = lbl_count_side_r.get(coord_r, 0) * -15
                lbl_count_side_r[coord_r] = lbl_count_side_r.get(coord_r, 0) + 1
                
                if scan_view in ["Front B-Scan", "양쪽 화면(Both)"]:
                    self.ax_side.annotate(cnums_side_r, xy=(center_x_right, actual_bottom_z + 2), xytext=(2, lbl_offset_y), textcoords='offset points', color=font_color, fontsize=16, ha='left', va='top', fontweight=font_weight)
                    if rest_r: self.ax_side.annotate(rest_r, xy=(center_x_right, actual_bottom_z + 2), xytext=(2 + offset_pt, lbl_offset_y), textcoords='offset points', color=font_color, fontsize=12, ha='left', va='top', fontweight=font_weight)
                if scan_view in ["Back B-Scan", "양쪽 화면(Both)"]:
                    self.ax_side_back.annotate(cnums_side_r, xy=(-center_x_right, actual_bottom_z + 2), xytext=(-2, lbl_offset_y), textcoords='offset points', color=font_color, fontsize=16, ha='right', va='top', fontweight=font_weight)
                    if rest_r: self.ax_side_back.annotate(rest_r, xy=(-center_x_right, actual_bottom_z + 2), xytext=(-2 - offset_pt, lbl_offset_y), textcoords='offset points', color=font_color, fontsize=12, ha='right', va='top', fontweight=font_weight)
            if side in ["좌측(Left)", "양측(Both)"] and not skip_text_l:
                offset_pt_l = len(cnums_side_l) * 15.0 if cnums_side_l else 0
                coord_l = (round(center_x_left, 1), round(center_z, 1))
                lbl_offset_y_l = lbl_count_side_l.get(coord_l, 0) * -15
                lbl_count_side_l[coord_l] = lbl_count_side_l.get(coord_l, 0) + 1
                
                if scan_view in ["Front B-Scan", "양쪽 화면(Both)"]:
                    self.ax_side.annotate(cnums_side_l, xy=(center_x_left, actual_bottom_z + 2), xytext=(-2, lbl_offset_y_l), textcoords='offset points', color=font_color, fontsize=16, ha='right', va='top', fontweight=font_weight)
                    if rest_l: self.ax_side.annotate(rest_l, xy=(center_x_left, actual_bottom_z + 2), xytext=(-2 - offset_pt_l, lbl_offset_y_l), textcoords='offset points', color=font_color, fontsize=12, ha='right', va='top', fontweight=font_weight)
                if scan_view in ["Back B-Scan", "양쪽 화면(Both)"]:
                    self.ax_side_back.annotate(cnums_side_l, xy=(-center_x_left, actual_bottom_z + 2), xytext=(2, lbl_offset_y_l), textcoords='offset points', color=font_color, fontsize=16, ha='left', va='top', fontweight=font_weight)
                    if rest_l: self.ax_side_back.annotate(rest_l, xy=(-center_x_left, actual_bottom_z + 2), xytext=(2 + offset_pt_l, lbl_offset_y_l), textcoords='offset points', color=font_color, fontsize=12, ha='left', va='top', fontweight=font_weight)

            # == Draw defect in Top View ==
            y_pos = dfct.get("y_pos", 0.0)
            y_len = dfct.get("y_length", 10.0)
            
            # Projected X-span from side view depends on shape
            if shape == "원형(Circle)":
                x_span = defect_width
            elif shape == "선(Line)":
                x_span = abs(p_end_r[0] - p_start_r[0])
            else: # 타원형, 사각형
                x_span = abs(p_end_r[0] - p_start_r[0]) + (defect_width * abs(math.cos(final_angle_rad)))
            
            if x_span < 1.0: x_span = 1.0
            
            if shape == "선(Line)":
                if side in ["우측(Right)", "양측(Both)"]:
                    self.ax_top.plot([center_x_right, center_x_right], [y_pos, y_pos + y_len], color=color, linewidth=4)
                if side in ["좌측(Left)", "양측(Both)"]:
                    self.ax_top.plot([center_x_left, center_x_left], [y_pos, y_pos + y_len], color=color, linewidth=4)
            elif shape == "원형(Circle)":
                if side in ["우측(Right)", "양측(Both)"]:
                    self.ax_top.add_patch(patches.Rectangle((center_x_right - defect_width/2, y_pos), defect_width, y_len, edgecolor=edge_color, facecolor=color, alpha=1.0, linewidth=2.0))
                if side in ["좌측(Left)", "양측(Both)"]:
                    self.ax_top.add_patch(patches.Rectangle((center_x_left - defect_width/2, y_pos), defect_width, y_len, edgecolor=edge_color, facecolor=color, alpha=1.0, linewidth=2.0))
            elif shape == "타원형(Ellipse)":
                if side in ["우측(Right)", "양측(Both)"]:
                    self.ax_top.add_patch(patches.Ellipse((center_x_right, y_pos + y_len/2), width=x_span, height=y_len, angle=0, edgecolor=edge_color, facecolor=color, alpha=1.0, linewidth=2.0))
                if side in ["좌측(Left)", "양측(Both)"]:
                    self.ax_top.add_patch(patches.Ellipse((center_x_left, y_pos + y_len/2), width=x_span, height=y_len, angle=0, edgecolor=edge_color, facecolor=color, alpha=1.0, linewidth=2.0))
            else: # 사각형(Rectangle)
                if side in ["우측(Right)", "양측(Both)"]:
                    self.ax_top.add_patch(patches.Rectangle((center_x_right - x_span/2, y_pos), x_span, y_len, edgecolor=edge_color, facecolor=color, alpha=1.0, linewidth=2.0))
                if side in ["좌측(Left)", "양측(Both)"]:
                    self.ax_top.add_patch(patches.Rectangle((center_x_left - x_span/2, y_pos), x_span, y_len, edgecolor=edge_color, facecolor=color, alpha=1.0, linewidth=2.0))
            
            key_top_r = (shape, round(defect_width, 1), round(y_len, 1), round(center_x_right, 1), round(y_pos, 1))
            key_top_l = (shape, round(defect_width, 1), round(y_len, 1), round(center_x_left, 1), round(y_pos, 1))
            
            cnums_top_r = ",".join(grp_top_r.get(key_top_r, [circle_num]))
            cnums_top_l = ",".join(grp_top_l.get(key_top_l, [circle_num]))
            
            if shape == "원형(Circle)":
                top_rest_r = ""
                top_rest_l = ""
            else:
                top_rest_r = rest_r
                top_rest_l = rest_l
            
            skip_top_r = False
            skip_top_l = False
            
            if key_top_r == sel_top_r_key and not is_selected:
                skip_top_r = True
            elif key_top_r in seen_top_r and not is_selected: skip_top_r = True
            else: seen_top_r.add(key_top_r)
            
            if key_top_l == sel_top_l_key and not is_selected:
                skip_top_l = True
            elif key_top_l in seen_top_l and not is_selected: skip_top_l = True
            else: seen_top_l.add(key_top_l)


            # --- Draw Dimensions (B) ---
            # 1. Top View Y position (from Y=0 to y_pos) on the left side
            skip_y_pos = False
            round_y_pos = round(y_pos, 1)
            if round_y_pos in seen_y_pos and not is_selected:
                skip_y_pos = True
            else:
                seen_y_pos.add(round_y_pos)
                
            if not skip_y_pos:
                dim_x_ypos = -calc_top_bevel_x - 15 - (len(seen_y_pos) * 15)
                self.ax_top.plot([0, dim_x_ypos - 2], [0, 0], color='red', lw=0.5, linestyle='--')
                self.ax_top.plot([0, dim_x_ypos - 2], [y_pos, y_pos], color='red', lw=0.5, linestyle='--')
                self.ax_top.annotate('', xy=(dim_x_ypos, y_pos), xytext=(dim_x_ypos, 0), arrowprops=dict(arrowstyle='|-|', color='red', lw=1, shrinkA=0, shrinkB=0))
                self.ax_top.text(dim_x_ypos - 1, y_pos / 2, f"{y_pos:.0f}", color='red', fontsize=11, ha='right', va='center', rotation=90, bbox=dict(facecolor='#2b2b2b', edgecolor='none', pad=1))
            
            # 2. Top View Length (y_len) near the defect
            if side in ["우측(Right)", "양측(Both)"] and not skip_top_r:
                ypos_round = round(y_pos, 1)
                offset_len = dim_count_top_len_r.get(ypos_round, 0) * 20
                dim_count_top_len_r[ypos_round] = dim_count_top_len_r.get(ypos_round, 0) + 1
                dim_x_len = center_x_right + x_span/2 + 15 + offset_len
                
                self.ax_top.plot([center_x_right, dim_x_len + 2], [y_pos, y_pos], color='red', lw=0.5, linestyle='--')
                self.ax_top.plot([center_x_right, dim_x_len + 2], [y_pos + y_len, y_pos + y_len], color='red', lw=0.5, linestyle='--')
                self.ax_top.annotate('', xy=(dim_x_len, y_pos), xytext=(dim_x_len, y_pos + y_len), arrowprops=dict(arrowstyle='|-|', color='red', lw=1, shrinkA=0, shrinkB=0))
                self.ax_top.text(dim_x_len + 1, y_pos + y_len/2, f"L:{y_len:.0f}", color='red', fontsize=11, ha='left', va='center', bbox=dict(facecolor='#2b2b2b', edgecolor='none', pad=1))
                
                offset_top_pt = len(cnums_top_r) * 15.0 if cnums_top_r else 0
                lbl_x_r = dim_x_len + 25
                if cnums_top_r: self.ax_top.annotate(cnums_top_r, xy=(lbl_x_r, y_pos + y_len/2), xytext=(0, 0), textcoords='offset points', color=font_color, fontsize=16, ha='left', va='center', fontweight=font_weight)
                if top_rest_r: self.ax_top.annotate(top_rest_r, xy=(lbl_x_r, y_pos + y_len/2), xytext=(offset_top_pt, 0), textcoords='offset points', color=font_color, fontsize=12, ha='left', va='center', fontweight=font_weight)
                
            if side in ["좌측(Left)", "양측(Both)"] and not skip_top_l:
                ypos_round = round(y_pos, 1)
                offset_len_l = dim_count_top_len_l.get(ypos_round, 0) * 20
                dim_count_top_len_l[ypos_round] = dim_count_top_len_l.get(ypos_round, 0) + 1
                dim_x_len_l = center_x_left - x_span/2 - 15 - offset_len_l
                
                self.ax_top.plot([center_x_left, dim_x_len_l - 2], [y_pos, y_pos], color='red', lw=0.5, linestyle='--')
                self.ax_top.plot([center_x_left, dim_x_len_l - 2], [y_pos + y_len, y_pos + y_len], color='red', lw=0.5, linestyle='--')
                self.ax_top.annotate('', xy=(dim_x_len_l, y_pos), xytext=(dim_x_len_l, y_pos + y_len), arrowprops=dict(arrowstyle='|-|', color='red', lw=1, shrinkA=0, shrinkB=0))
                self.ax_top.text(dim_x_len_l - 1, y_pos + y_len/2, f"L:{y_len:.0f}", color='red', fontsize=11, ha='right', va='center', bbox=dict(facecolor='#2b2b2b', edgecolor='none', pad=1))
                
                offset_top_pt_l = len(top_rest_l.split('\n')[0]) * 10.0 if top_rest_l else 0
                lbl_x_l = dim_x_len_l - 25
                if top_rest_l: self.ax_top.annotate(top_rest_l, xy=(lbl_x_l, y_pos + y_len/2), xytext=(0, 0), textcoords='offset points', color=font_color, fontsize=12, ha='right', va='center', fontweight=font_weight)
                if cnums_top_l: self.ax_top.annotate(cnums_top_l, xy=(lbl_x_l, y_pos + y_len/2), xytext=(-offset_top_pt_l, 0), textcoords='offset points', color=font_color, fontsize=16, ha='right', va='center', fontweight=font_weight)

            # 3. Side View Z depth (from Z=0 to actual_top_z) & 4. Height (H)
            scan_view = dfct.get("scan_view", "Front B-Scan")
            
            if side in ["우측(Right)", "양측(Both)"] and not skip_text_r:
                cx_round = round(center_x_right, 1)
                offset = dim_count_side_r.get(cx_round, 0) * 35
                dim_count_side_r[cx_round] = dim_count_side_r.get(cx_round, 0) + 1
                
                dim_x_z = center_x_right + defect_width/2 + 10 + offset
                dim_x_h = center_x_right + defect_width/2 + 25 + offset
                
                if scan_view in ["Front B-Scan", "양쪽 화면(Both)"]:
                    # Depth
                    self.ax_side.plot([center_x_right, dim_x_z + 2], [0, 0], color='red', lw=0.5, linestyle='--')
                    self.ax_side.plot([center_x_right, dim_x_z + 2], [actual_top_z, actual_top_z], color='red', lw=0.5, linestyle='--')
                    self.ax_side.annotate('', xy=(dim_x_z, actual_top_z), xytext=(dim_x_z, 0), arrowprops=dict(arrowstyle='|-|', color='red', lw=1, shrinkA=0, shrinkB=0))
                    self.ax_side.text(dim_x_z + 1, actual_top_z + 2.5, f"{actual_top_z:.1f}", color='red', fontsize=11, ha='left', va='center', bbox=dict(facecolor='#2b2b2b', edgecolor='none', pad=1))
                    # Height
                    if shape != "원형(Circle)":
                        self.ax_side.plot([center_x_right, dim_x_h + 2], [actual_bottom_z, actual_bottom_z], color='red', lw=0.5, linestyle='--')
                        self.ax_side.annotate('', xy=(dim_x_h, actual_top_z), xytext=(dim_x_h, actual_bottom_z), arrowprops=dict(arrowstyle='|-|', color='red', lw=1, shrinkA=0, shrinkB=0))
                        self.ax_side.text(dim_x_h + 1, center_z, f"H:{h_val:.1f}", color='red', fontsize=11, ha='left', va='center', bbox=dict(facecolor='#2b2b2b', edgecolor='none', pad=1))
                if scan_view in ["Back B-Scan", "양쪽 화면(Both)"]:
                    dim_x_z_m = -center_x_right - defect_width/2 - 10 - offset
                    dim_x_h_m = -center_x_right - defect_width/2 - 25 - offset
                    self.ax_side_back.plot([-center_x_right, dim_x_z_m - 2], [0, 0], color='red', lw=0.5, linestyle='--')
                    self.ax_side_back.plot([-center_x_right, dim_x_z_m - 2], [actual_top_z, actual_top_z], color='red', lw=0.5, linestyle='--')
                    self.ax_side_back.annotate('', xy=(dim_x_z_m, actual_top_z), xytext=(dim_x_z_m, 0), arrowprops=dict(arrowstyle='|-|', color='red', lw=1, shrinkA=0, shrinkB=0))
                    self.ax_side_back.text(dim_x_z_m - 1, actual_top_z + 2.5, f"{actual_top_z:.1f}", color='red', fontsize=11, ha='right', va='center', bbox=dict(facecolor='#2b2b2b', edgecolor='none', pad=1))
                    if shape != "원형(Circle)":
                        self.ax_side_back.plot([-center_x_right, dim_x_h_m - 2], [actual_bottom_z, actual_bottom_z], color='red', lw=0.5, linestyle='--')
                        self.ax_side_back.annotate('', xy=(dim_x_h_m, actual_top_z), xytext=(dim_x_h_m, actual_bottom_z), arrowprops=dict(arrowstyle='|-|', color='red', lw=1, shrinkA=0, shrinkB=0))
                        self.ax_side_back.text(dim_x_h_m - 1, center_z, f"H:{h_val:.1f}", color='red', fontsize=11, ha='right', va='center', bbox=dict(facecolor='#2b2b2b', edgecolor='none', pad=1))
                
            if side in ["좌측(Left)", "양측(Both)"] and not skip_text_l:
                cxl_round = round(center_x_left, 1)
                offset_l = dim_count_side_l.get(cxl_round, 0) * 35
                dim_count_side_l[cxl_round] = dim_count_side_l.get(cxl_round, 0) + 1
                
                dim_x_z_l = center_x_left - defect_width/2 - 10 - offset_l
                dim_x_h_l = center_x_left - defect_width/2 - 25 - offset_l
                
                if scan_view in ["Front B-Scan", "양쪽 화면(Both)"]:
                    self.ax_side.plot([center_x_left, dim_x_z_l - 2], [0, 0], color='red', lw=0.5, linestyle='--')
                    self.ax_side.plot([center_x_left, dim_x_z_l - 2], [actual_top_z, actual_top_z], color='red', lw=0.5, linestyle='--')
                    self.ax_side.annotate('', xy=(dim_x_z_l, actual_top_z), xytext=(dim_x_z_l, 0), arrowprops=dict(arrowstyle='|-|', color='red', lw=1, shrinkA=0, shrinkB=0))
                    self.ax_side.text(dim_x_z_l - 1, actual_top_z + 2.5, f"{actual_top_z:.1f}", color='red', fontsize=11, ha='right', va='center', bbox=dict(facecolor='#2b2b2b', edgecolor='none', pad=1))
                    if shape != "원형(Circle)":
                        self.ax_side.plot([center_x_left, dim_x_h_l - 2], [actual_bottom_z, actual_bottom_z], color='red', lw=0.5, linestyle='--')
                        self.ax_side.annotate('', xy=(dim_x_h_l, actual_top_z), xytext=(dim_x_h_l, actual_bottom_z), arrowprops=dict(arrowstyle='|-|', color='red', lw=1, shrinkA=0, shrinkB=0))
                        self.ax_side.text(dim_x_h_l - 1, center_z, f"H:{h_val:.1f}", color='red', fontsize=11, ha='right', va='center', bbox=dict(facecolor='#2b2b2b', edgecolor='none', pad=1))
                if scan_view in ["Back B-Scan", "양쪽 화면(Both)"]:
                    dim_x_z_l_m = -center_x_left + defect_width/2 + 10 + offset_l
                    dim_x_h_l_m = -center_x_left + defect_width/2 + 25 + offset_l
                    self.ax_side_back.plot([-center_x_left, dim_x_z_l_m + 2], [0, 0], color='red', lw=0.5, linestyle='--')
                    self.ax_side_back.plot([-center_x_left, dim_x_z_l_m + 2], [actual_top_z, actual_top_z], color='red', lw=0.5, linestyle='--')
                    self.ax_side_back.annotate('', xy=(dim_x_z_l_m, actual_top_z), xytext=(dim_x_z_l_m, 0), arrowprops=dict(arrowstyle='|-|', color='red', lw=1, shrinkA=0, shrinkB=0))
                    self.ax_side_back.text(dim_x_z_l_m + 1, actual_top_z + 2.5, f"{actual_top_z:.1f}", color='red', fontsize=11, ha='left', va='center', bbox=dict(facecolor='#2b2b2b', edgecolor='none', pad=1))
                    if shape != "원형(Circle)":
                        self.ax_side_back.plot([-center_x_left, dim_x_h_l_m + 2], [actual_bottom_z, actual_bottom_z], color='red', lw=0.5, linestyle='--')
                        self.ax_side_back.annotate('', xy=(dim_x_h_l_m, actual_top_z), xytext=(dim_x_h_l_m, actual_bottom_z), arrowprops=dict(arrowstyle='|-|', color='red', lw=1, shrinkA=0, shrinkB=0))
                        self.ax_side_back.text(dim_x_h_l_m + 1, center_z, f"H:{h_val:.1f}", color='red', fontsize=11, ha='left', va='center', bbox=dict(facecolor='#2b2b2b', edgecolor='none', pad=1))

        # Side View 설정
        for ax, title, invert_x in [(self.ax_side, 'Side View (Front B-Scan)', False), 
                                    (self.ax_side_back, 'Side View (Back B-Scan)', True)]:
            ax.invert_yaxis()
            ax.set_xlabel('Y Position (mm)', color='white')
            ax.set_ylabel('Z Depth (mm)', color='white')
            ax.set_title(title, color='white', fontweight='bold')
            ax.grid(True, linestyle=':', alpha=0.3)
            legend_side = ax.legend(loc='upper right')
            if legend_side:
                for text in legend_side.get_texts(): text.set_color("black")
            
            from matplotlib.ticker import ScalarFormatter
            ax.set_xlim(-half_width, half_width)
            ax.set_xscale('symlog', linthresh=30, linscale=0.56)
            ax.xaxis.set_major_formatter(ScalarFormatter())
            
            ticks = [-half_width, -100, -30, 0, 30, 100, half_width]
            ax.set_xticks(ticks)
            if invert_x:
                ax.set_xticklabels([-t for t in ticks])
                
            ax.set_ylim(t + 15, -10) # 결함이 잘리지 않도록 여백을 넉넉히 줍니다.
        
        # Top View 설정
        self.ax_top.set_xlabel('Y Position (mm)', color='white')
        self.ax_top.set_ylabel('X Position (mm)', color='white')
        self.ax_top.set_title('Top View (C-Scan)', color='white', fontweight='bold')
        self.ax_top.grid(True, linestyle=':', alpha=0.3)
        legend_top = self.ax_top.legend(loc='upper right')
        if legend_top:
            for text in legend_top.get_texts(): text.set_color("black")
        from matplotlib.ticker import ScalarFormatter
        self.ax_top.set_xlim(-half_width, half_width)
        self.ax_top.set_xscale('symlog', linthresh=30, linscale=0.56)
        self.ax_top.xaxis.set_major_formatter(ScalarFormatter())
        
        # 가로축(Y Position) 틱 동적 생성
        self.ax_top.set_xticks([-half_width, -100, -30, 0, 30, 100, half_width])
        
        # 세로축(X Position) 동적 설정
        specimen_length = p["specimen_length"]
        self.ax_top.set_ylim(0, specimen_length)
        
        # 세로축 틱 자동 계산 (50단위로, 마지막은 specimen_length 포함)
        yticks = list(range(0, int(specimen_length) + 1, 50))
        if yticks[-1] != specimen_length:
            yticks.append(specimen_length)
        self.ax_top.set_yticks(yticks)
        
        self.fig.tight_layout(pad=2.0)
        self.canvas.draw()
        
        # 결과 요약 출력
        msg = f"[그래프 갱신 완료]\n상단 개선각 너비(X) 계산 결과: {calc_top_bevel_x:.3f} mm\n루트면 시작 깊이(Z): {z_root_top:.2f} mm"
        self.show_result(msg)

    def show_result(self, text):
        self.result_box.configure(state="normal")
        self.result_box.delete("0.0", "end")
        self.result_box.insert("0.0", text + "\n")
        self.result_box.configure(state="disabled")

    def add_front_defect(self):
        self.add_defect(scan_view="Front B-Scan")
        
    def add_back_defect(self):
        self.add_defect(scan_view="Back B-Scan")

    def add_defect(self, side=None, scan_view=None):
        try:
            w = float(self.entries["defect_width"].get())
        except:
            w = 2.5
            
        try:
            p = self.get_params()
            t = p["thickness"]
            r_f = p["root_face"]
            r_g = p["root_gap_half"]
            ang = p["bevel_angle_deg"]
            z_root_top = t - r_f
            if 7.0 <= z_root_top:
                y_c = r_g + (z_root_top - 7.0) * math.tan(math.radians(ang))
            else:
                y_c = r_g
        except:
            y_c = 10.0
            
        selected_shape = self.shape_var.get()
        selected_side = side if side else self.side_var.get()
        selected_scan = scan_view if scan_view else self.scan_var.get()
        new_dfct = {
            "z_start": 5.0, "z_end": 9.0, "y_center": round(y_c, 1), 
            "width": w, "angle": 0.0, "shape": selected_shape, "side": selected_side, "scan_view": selected_scan, "y_pos": 160.0, "y_length": 10.0,
            "custom_label": "SDH" if selected_shape == "원형(Circle)" else ""
        }
        self.defects.append(new_dfct)
        self.selected_defect_idx = len(self.defects) - 1
        self.update_ui_from_selected()
        self.update_plot()
        self.show_result(f"새 결함이 추가되었습니다. (총 {len(self.defects)}개)")
        
    def delete_defect(self):
        if 0 <= self.selected_defect_idx < len(self.defects):
            self.defects.pop(self.selected_defect_idx)
            if len(self.defects) > 0:
                self.selected_defect_idx = len(self.defects) - 1
                self.update_ui_from_selected()
                self.update_plot()
                self.show_result(f"결함이 삭제되었습니다. (남은 결함: {len(self.defects)}개)")
            else:
                self.selected_defect_idx = -1
                self.ax_side.clear()
                self.ax_top.clear()
                self.update_plot()
                self.show_result("모든 결함이 삭제되었습니다.")
                
    def update_ui_from_selected(self):
        if 0 <= self.selected_defect_idx < len(self.defects):
            dfct = self.defects[self.selected_defect_idx]
            self.entries["defect_start_depth"].delete(0, "end")
            self.entries["defect_start_depth"].insert(0, f"{dfct['z_start']:.2f}")
            self.entries["defect_end_depth"].delete(0, "end")
            self.entries["defect_end_depth"].insert(0, f"{dfct['z_end']:.2f}")
            self.entries["defect_y_center"].delete(0, "end")
            self.entries["defect_y_center"].insert(0, f"{dfct['y_center']:.2f}")
            self.entries["defect_width"].delete(0, "end")
            self.entries["defect_width"].insert(0, f"{dfct['width']:.2f}")
            self.entries["defect_angle"].delete(0, "end")
            self.entries["defect_angle"].insert(0, f"{dfct['angle']:.1f}")
            self.entries["defect_y_pos"].delete(0, "end")
            self.entries["defect_y_pos"].insert(0, f"{dfct.get('y_pos', 0.0):.1f}")
            self.entries["defect_y_length"].delete(0, "end")
            self.entries["defect_y_length"].insert(0, f"{dfct.get('y_length', 10.0):.1f}")
            self.entries["defect_custom_label"].delete(0, "end")
            self.entries["defect_custom_label"].insert(0, dfct.get("custom_label", ""))
            if hasattr(self, 'shape_var'): self.shape_var.set(dfct["shape"])
            if hasattr(self, 'side_var'): self.side_var.set(dfct["side"])
            if hasattr(self, 'scan_var'): self.scan_var.set(dfct.get("scan_view", "Front B-Scan"))
            
    def apply_defect_properties(self):
        if self.selected_defect_idx < 0 or self.selected_defect_idx >= len(self.defects): return
        try:
            self.defects[self.selected_defect_idx]["z_start"] = float(self.entries["defect_start_depth"].get())
            self.defects[self.selected_defect_idx]["z_end"] = float(self.entries["defect_end_depth"].get())
            self.defects[self.selected_defect_idx]["y_center"] = float(self.entries["defect_y_center"].get())
            self.defects[self.selected_defect_idx]["width"] = float(self.entries["defect_width"].get())
            self.defects[self.selected_defect_idx]["angle"] = float(self.entries["defect_angle"].get())
            self.defects[self.selected_defect_idx]["y_pos"] = float(self.entries["defect_y_pos"].get())
            self.defects[self.selected_defect_idx]["y_length"] = float(self.entries["defect_y_length"].get())
            self.defects[self.selected_defect_idx]["custom_label"] = self.entries["defect_custom_label"].get()
            self.defects[self.selected_defect_idx]["shape"] = self.shape_var.get()
            self.defects[self.selected_defect_idx]["side"] = self.side_var.get()
            self.defects[self.selected_defect_idx]["scan_view"] = self.scan_var.get()
        except ValueError:
            self.show_result("입력값이 올바르지 않습니다.")
            return
        self.update_plot()
        msg = f"{self.selected_defect_idx + 1}번 결함 속성 적용 완료!\n"
        msg += f"- 현재 형상: {self.defects[self.selected_defect_idx]['shape']}\n"
        msg += f"- 현재 폭(Width): {self.defects[self.selected_defect_idx]['width']} mm"
        self.show_result(msg)

    def on_press(self, event):
        if event.inaxes not in [self.ax_side, self.ax_side_back, self.ax_top]: return
        self.active_ax = event.inaxes
        
        p = self.get_params()
        if not p: return
        t = p["thickness"]
        r_f = p["root_face"]
        r_g = p["root_gap_half"]
        ang = p["bevel_angle_deg"]
        z_root_top = t - r_f
        bevel_angle_rad = math.radians(ang)
        
        cx, cz = event.xdata, event.ydata
        if cx is None or cz is None: return
        
        hit_idx = -1
        hit_side = None
        orig_center_x = None
        orig_center_z = None
        
        for i in range(len(getattr(self, 'defects', []))-1, -1, -1):
            dfct = self.defects[i]
            center_z = (dfct["z_start"] + dfct["z_end"]) / 2
                
            right_cx = dfct["y_center"]
            left_cx = -dfct["y_center"]
            
            dz = abs(cz - center_z)
            height_half = max(2.0, abs(dfct["z_end"] - dfct["z_start"]) / 2)
            width_margin = max(2.0, dfct["width"] * 1.5)
            
            side = dfct["side"]
            y_pos = dfct.get("y_pos", 0.0)
            y_half = max(2.0, dfct.get("y_length", 10.0) / 2)
            
            is_hit = False
            hit_right = False
            hit_left = False
            
            scan_view = dfct.get("scan_view", "Front B-Scan")
            
            if self.active_ax == self.ax_side:
                if scan_view in ["Front B-Scan", "양쪽 화면(Both)"]:
                    hit_right = (abs(cx - right_cx) <= width_margin and dz <= height_half)
                    hit_left = (abs(cx - left_cx) <= width_margin and dz <= height_half)
            elif self.active_ax == self.ax_side_back:
                if scan_view in ["Back B-Scan", "양쪽 화면(Both)"]:
                    hit_right = (abs(cx - (-right_cx)) <= width_margin and dz <= height_half)
                    hit_left = (abs(cx - (-left_cx)) <= width_margin and dz <= height_half)
            else:
                hit_right = (abs(cx - right_cx) <= width_margin and abs(cz - y_pos) <= y_half)
                hit_left = (abs(cx - left_cx) <= width_margin and abs(cz - y_pos) <= y_half)
                
            if side in ["우측(Right)", "양측(Both)"] and hit_right:
                hit_idx = i
                hit_side = "우측(Right)"
                if self.active_ax == self.ax_side:
                    orig_center_x = right_cx
                    orig_center_z = center_z
                elif self.active_ax == self.ax_side_back:
                    orig_center_x = -right_cx
                    orig_center_z = center_z
                elif self.active_ax == self.ax_top:
                    orig_center_x = right_cx
                    orig_center_z = y_pos
                break
            
            if side in ["좌측(Left)", "양측(Both)"] and hit_left:
                hit_idx = i
                hit_side = "좌측(Left)"
                if self.active_ax == self.ax_side:
                    orig_center_x = left_cx
                    orig_center_z = center_z
                elif self.active_ax == self.ax_side_back:
                    orig_center_x = -left_cx
                    orig_center_z = center_z
                elif self.active_ax == self.ax_top:
                    orig_center_x = left_cx
                    orig_center_z = y_pos
                break
                    
        if hit_idx != -1:
            self.selected_defect_idx = hit_idx
            self.update_ui_from_selected()
            self.update_plot()
            
            self.dragging = True
            self.drag_start_x = cx
            self.drag_start_z = cz
            self.drag_side = hit_side
            self.orig_cx = orig_center_x
            self.orig_cz = orig_center_z
            self.mouse_offset_x = cx - orig_center_x
            self.mouse_offset_z = cz - orig_center_z
            
            self.orig_z_start = self.defects[hit_idx]["z_start"]
            self.orig_z_end = self.defects[hit_idx]["z_end"]
            self.orig_y_center = self.defects[hit_idx]["y_center"]
            self.orig_angle = self.defects[hit_idx]["angle"]
            
            if event.button == 1:
                self.drag_mode = "translate"
            elif event.button == 3:
                self.drag_mode = "rotate"
                self.start_mouse_angle = math.degrees(math.atan2(cz - self.orig_cz, cx - self.orig_cx))
            else:
                self.dragging = False
                return
                
            if self.defects[hit_idx]["side"] == "양측(Both)":
                self.defects[hit_idx]["side"] = hit_side
                self.update_ui_from_selected()
        else:
            self.selected_defect_idx = -1
            self.update_plot()

    def on_motion(self, event):
        if not getattr(self, 'dragging', False) or event.inaxes != self.active_ax: return
        if event.xdata is None or event.ydata is None: return
        if self.selected_defect_idx < 0 or self.selected_defect_idx >= len(self.defects): return
        
        if self.drag_mode == "translate":
            desired_cz = event.ydata - self.mouse_offset_z
            desired_cx = event.xdata - self.mouse_offset_x
            
            if self.active_ax in [self.ax_side, self.ax_side_back]:
                dz = desired_cz - self.orig_cz
                new_z_start = self.orig_z_start + dz
                new_z_end = self.orig_z_end + dz
            else:
                new_z_start = self.orig_z_start
                new_z_end = self.orig_z_end
                self.defects[self.selected_defect_idx]["y_pos"] = desired_cz
            
            # 클램핑: 결함의 '중심(Center)'이 0부터 t 사이를 벗어나지 못하도록 제한합니다.
            p = self.get_params()
            if p:
                t = p["thickness"]
                new_center_z = (new_z_start + new_z_end) / 2
                if new_center_z < 0:
                    shift = -new_center_z
                    new_z_start += shift
                    new_z_end += shift
                    desired_cz += shift
                elif new_center_z > t:
                    shift = new_center_z - t
                    new_z_start -= shift
                    new_z_end -= shift
                    desired_cz -= shift
                    
            if self.active_ax in [self.ax_side, self.ax_side_back]:
                if self.drag_side == "우측(Right)":
                    new_y_center = desired_cx
                else: # 좌측
                    # On ax_side_back, desired_cx is visually positive when on the right side of the screen
                    new_y_center = desired_cx if self.active_ax == self.ax_side_back else -desired_cx
            else: # ax_top
                if self.drag_side == "우측(Right)":
                    new_y_center = desired_cx
                else:
                    new_y_center = -desired_cx
                
            self.defects[self.selected_defect_idx]["z_start"] = new_z_start
            self.defects[self.selected_defect_idx]["z_end"] = new_z_end
            self.defects[self.selected_defect_idx]["y_center"] = new_y_center
            
        elif self.drag_mode == "rotate":
            current_mouse_angle = math.degrees(math.atan2(event.ydata - self.orig_cz, event.xdata - self.orig_cx))
            angle_diff = current_mouse_angle - self.start_mouse_angle
            
            if self.active_ax == self.ax_side_back:
                angle_diff = -angle_diff
                
            if self.drag_side == "좌측(Left)":
                angle_diff = -angle_diff
                
            new_angle = self.orig_angle + angle_diff
            self.defects[self.selected_defect_idx]["angle"] = new_angle
            
        self.update_ui_from_selected()
        self.update_plot()

    def on_release(self, event):
        self.dragging = False
        self.drag_start_x = None
        self.drag_start_z = None

    def save_project(self):
        filepath = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON Files", "*.json")], title="프로젝트 저장")
        if not filepath:
            return
            
        data = {
            "inputs": {k: v.get() for k, v in self.entries.items()},
            "defects": self.defects
        }
        
        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
            messagebox.showinfo("저장 완료", "설정이 성공적으로 저장되었습니다.")
        except Exception as e:
            messagebox.showerror("저장 오류", f"저장 중 오류가 발생했습니다:\n{str(e)}")
            
    def load_project(self):
        filepath = filedialog.askopenfilename(filetypes=[("JSON Files", "*.json")], title="프로젝트 불러오기")
        if not filepath:
            return
            
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)
                
            if "inputs" in data:
                for k, v in data["inputs"].items():
                    if k in self.entries:
                        self.entries[k].delete(0, 'end')
                        self.entries[k].insert(0, str(v))
                        
            if "defects" in data:
                self.defects = data["defects"]
                
            self.selected_defect_idx = 0 if len(self.defects) > 0 else -1
            self.update_plot()
            messagebox.showinfo("불러오기 완료", "설정을 성공적으로 불러왔습니다.")
        except Exception as e:
            messagebox.showerror("불러오기 오류", f"파일을 불러오는 중 오류가 발생했습니다:\n{str(e)}")

if __name__ == "__main__":
    app = App()
    app.mainloop()
