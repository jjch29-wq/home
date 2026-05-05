
import tkinter as tk
from tkinter import ttk, messagebox
import math

class PipeAreaCalculator:
    def __init__(self, root):
        self.root = root
        self.root.title("Pipe Surface Area Calculator")
        self.root.geometry("480x580")
        self.root.configure(bg="#1e1e1e")
        self.root.resizable(False, False)

        self.setup_styles()
        self.create_widgets()

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        # Dark Theme Configuration
        style.configure("TFrame", background="#1e1e1e")
        style.configure("TLabel", background="#1e1e1e", foreground="#ffffff", font=("Segoe UI", 11))
        style.configure("Header.TLabel", background="#1e1e1e", foreground="#4cc9f0", font=("Segoe UI", 16, "bold"))
        style.configure("Result.TLabel", background="#2d2d2d", foreground="#f72585", font=("Segoe UI", 24, "bold"), padding=10)
        style.configure("Unit.TLabel", background="#1e1e1e", foreground="#aaaaaa", font=("Segoe UI", 9))
        
        style.configure("TEntry", fieldbackground="#333333", foreground="#ffffff", insertcolor="#ffffff")
        style.configure("TButton", font=("Segoe UI", 12, "bold"), background="#4cc9f0", foreground="#ffffff", borderwidth=0)
        style.map("TButton", background=[('active', '#3a0ca3')])
        
        # Notebook (Tabs) Styling
        style.configure("TNotebook", background="#1e1e1e", borderwidth=0)
        style.configure("TNotebook.Tab", background="#333333", foreground="#ffffff", padding=[20, 5], font=("Segoe UI", 10))
        style.map("TNotebook.Tab", background=[("selected", "#4cc9f0")], foreground=[("selected", "#ffffff")])

    def create_widgets(self):
        # Header
        header_frame = ttk.Frame(self.root, padding=(30, 20, 30, 10))
        header_frame.pack(fill="x")
        ttk.Label(header_frame, text="파이프 계산기", style="Header.TLabel").pack()

        # Tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        # Tab 1: Circular (Pipe)
        self.tab_circ = ttk.Frame(self.notebook, padding=20)
        self.notebook.add(self.tab_circ, text="원형 (Pipe/Tube)")
        self.create_circ_tab()

        # Tab 2: Rectangle (Duct)
        self.tab_rect = ttk.Frame(self.notebook, padding=20)
        self.notebook.add(self.tab_rect, text="사각덕트 (Duct)")
        self.create_rect_tab()

        # Tab 3: Plate (Flat)
        self.tab_plate = ttk.Frame(self.notebook, padding=20)
        self.notebook.add(self.tab_plate, text="철판/평면 (Plate)")
        self.create_plate_tab()

    def create_circ_tab(self):
        # Calc Mode
        self.c_model_var = tk.StringVar(value="area")
        mode_frame = ttk.Frame(self.tab_circ)
        mode_frame.grid(row=0, column=0, sticky="ew", pady=(0, 15))
        ttk.Radiobutton(mode_frame, text="면적(m²) 구하기", variable=self.c_model_var, value="area", command=self.update_circ_ui).pack(side="left", padx=5)
        ttk.Radiobutton(mode_frame, text="길이(m) 구하기", variable=self.c_model_var, value="len", command=self.update_circ_ui).pack(side="left", padx=5)

        # Diameter
        ttk.Label(self.tab_circ, text="파이프 외경 (Diameter):").grid(row=1, column=0, sticky="w", pady=(0, 5))
        self.c_dia_var = tk.StringVar()
        ttk.Entry(self.tab_circ, textvariable=self.c_dia_var, font=("Segoe UI", 12)).grid(row=2, column=0, sticky="ew")
        
        c_dia_unit_frame = ttk.Frame(self.tab_circ)
        c_dia_unit_frame.grid(row=3, column=0, sticky="e", pady=(2, 10))
        self.c_dia_unit = tk.StringVar(value="mm")
        ttk.Radiobutton(c_dia_unit_frame, text="mm", variable=self.c_dia_unit, value="mm").pack(side="left")
        ttk.Radiobutton(c_dia_unit_frame, text="inch", variable=self.c_dia_unit, value="inch").pack(side="left")

        # Variable Input Label (Length or Area)
        self.c_var_label = ttk.Label(self.tab_circ, text="파이프 길이 (Length, m):")
        self.c_var_label.grid(row=4, column=0, sticky="w", pady=(10, 5))
        self.c_var_input = tk.StringVar()
        ttk.Entry(self.tab_circ, textvariable=self.c_var_input, font=("Segoe UI", 12)).grid(row=5, column=0, sticky="ew")

        # Calculate Button
        self.c_btn = ttk.Button(self.tab_circ, text="계산 (Calculate)", command=self.calc_circ)
        self.c_btn.grid(row=6, column=0, pady=25, sticky="ew")

        # Result Display
        self.c_res_title = ttk.Label(self.tab_circ, text="결과 (Area):")
        self.c_res_title.grid(row=7, column=0, sticky="w")
        self.c_result_label = tk.Label(self.tab_circ, text="0.0000 m²", bg="#2d2d2d", fg="#f72585", font=("Segoe UI", 24, "bold"), pady=10)
        self.c_result_label.grid(row=8, column=0, sticky="ew", pady=10)
        
        self.tab_circ.columnconfigure(0, weight=1)

    def update_circ_ui(self):
        if self.c_model_var.get() == "area":
            self.c_var_label.config(text="파이프 길이 (Length, m):")
            self.c_res_title.config(text="결과 (Surface Area):")
            self.c_result_label.config(text="0.0000 m²", fg="#f72585")
        else:
            self.c_var_label.config(text="목표 면적 (Target Area, m²):")
            self.c_res_title.config(text="결과 (Required Length):")
            self.c_result_label.config(text="0.0000 m", fg="#4cc9f0")

    def create_rect_tab(self):
        # Calc Mode
        self.r_model_var = tk.StringVar(value="area")
        mode_frame = ttk.Frame(self.tab_rect)
        mode_frame.grid(row=0, column=0, sticky="ew", pady=(0, 15))
        ttk.Radiobutton(mode_frame, text="면적(m²) 구하기", variable=self.r_model_var, value="area", command=self.update_rect_ui).pack(side="left", padx=5)
        ttk.Radiobutton(mode_frame, text="길이(m) 구하기", variable=self.r_model_var, value="len", command=self.update_rect_ui).pack(side="left", padx=5)

        # Width
        ttk.Label(self.tab_rect, text="가로 (Width, mm):").grid(row=1, column=0, sticky="w", pady=(0, 5))
        self.rect_w_var = tk.StringVar()
        ttk.Entry(self.tab_rect, textvariable=self.rect_w_var, font=("Segoe UI", 12)).grid(row=2, column=0, sticky="ew")

        # Height
        ttk.Label(self.tab_rect, text="세로 (Height, mm):").grid(row=3, column=0, sticky="w", pady=(10, 5))
        self.rect_h_var = tk.StringVar()
        ttk.Entry(self.tab_rect, textvariable=self.rect_h_var, font=("Segoe UI", 12)).grid(row=4, column=0, sticky="ew")

        # Variable Input Label (Length or Area)
        self.rect_var_label = ttk.Label(self.tab_rect, text="길이 (Length, m):")
        self.rect_var_label.grid(row=5, column=0, sticky="w", pady=(10, 5))
        self.rect_var_input = tk.StringVar()
        ttk.Entry(self.tab_rect, textvariable=self.rect_var_input, font=("Segoe UI", 12)).grid(row=6, column=0, sticky="ew")

        # Calculate Button
        ttk.Button(self.tab_rect, text="계산 (Calculate)", command=self.calc_rect).grid(row=7, column=0, pady=25, sticky="ew")

        # Result Display
        self.rect_res_title = ttk.Label(self.tab_rect, text="결과 (Area):")
        self.rect_res_title.grid(row=8, column=0, sticky="w")
        self.rect_result_label = tk.Label(self.tab_rect, text="0.0000 m²", bg="#2d2d2d", fg="#ffca3a", font=("Segoe UI", 24, "bold"), pady=10)
        self.rect_result_label.grid(row=9, column=0, sticky="ew", pady=10)
        
        self.tab_rect.columnconfigure(0, weight=1)

    def update_rect_ui(self):
        if self.r_model_var.get() == "area":
            self.rect_var_label.config(text="길이 (Length, m):")
            self.rect_res_title.config(text="결과 (Surface Area):")
            self.rect_result_label.config(text="0.0000 m²", fg="#ffca3a")
        else:
            self.rect_var_label.config(text="목표 면적 (Target Area, m²):")
            self.rect_res_title.config(text="결과 (Required Length):")
            self.rect_result_label.config(text="0.0000 m", fg="#4cc9f0")

    def create_plate_tab(self):
        # Calc Mode
        self.p_model_var = tk.StringVar(value="area")
        mode_frame = ttk.Frame(self.tab_plate)
        mode_frame.grid(row=0, column=0, sticky="ew", pady=(0, 15))
        ttk.Radiobutton(mode_frame, text="면적(m²) 구하기", variable=self.p_model_var, value="area", command=self.update_plate_ui).pack(side="left", padx=5)
        ttk.Radiobutton(mode_frame, text="길이(m) 구하기", variable=self.p_model_var, value="len", command=self.update_plate_ui).pack(side="left", padx=5)

        # Width
        ttk.Label(self.tab_plate, text="가로 (Width, mm):").grid(row=1, column=0, sticky="w", pady=(0, 5))
        self.plate_w_var = tk.StringVar()
        ttk.Entry(self.tab_plate, textvariable=self.plate_w_var, font=("Segoe UI", 12)).grid(row=2, column=0, sticky="ew")

        # Variable Input Label (Height or Area)
        self.plate_var_label = ttk.Label(self.tab_plate, text="세로 (Height, mm):")
        self.plate_var_label.grid(row=3, column=0, sticky="w", pady=(15, 5))
        self.plate_var_input = tk.StringVar()
        ttk.Entry(self.tab_plate, textvariable=self.plate_var_input, font=("Segoe UI", 12)).grid(row=4, column=0, sticky="ew")

        # Calculate Button
        ttk.Button(self.tab_plate, text="계산 (Calculate)", command=self.calc_plate).grid(row=5, column=0, pady=40, sticky="ew")

        # Result Display
        self.plate_res_title = ttk.Label(self.tab_plate, text="결과 (Area):")
        self.plate_res_title.grid(row=6, column=0, sticky="w")
        self.plate_result_label = tk.Label(self.tab_plate, text="0.0000 m²", bg="#2d2d2d", fg="#a2d149", font=("Segoe UI", 24, "bold"), pady=15)
        self.plate_result_label.grid(row=7, column=0, sticky="ew", pady=10)
        
        self.tab_plate.columnconfigure(0, weight=1)

    def update_plate_ui(self):
        if self.p_model_var.get() == "area":
            self.plate_var_label.config(text="세로 (Height, mm):")
            self.plate_res_title.config(text="결과 (Surface Area):")
            self.plate_result_label.config(text="0.0000 m²", fg="#a2d149")
        else:
            self.plate_var_label.config(text="목표 면적 (Target Area, m²):")
            self.plate_res_title.config(text="결과 (Required Height, mm):")
            self.plate_result_label.config(text="0.0000 mm", fg="#4cc9f0")

    def calc_circ(self):
        try:
            d = float(self.c_dia_var.get())
            val = float(self.c_var_input.get())
            d_m = (d / 1000.0) if self.c_dia_unit.get() == "mm" else (d * 0.0254)
            
            if self.c_model_var.get() == "area":
                # L -> Area
                area = math.pi * d_m * val
                self.c_result_label.config(text=f"{area:.4f} m²")
            else:
                # Area -> L
                if d_m == 0: raise ZeroDivisionError
                length = val / (math.pi * d_m)
                self.c_result_label.config(text=f"{length:.4f} m")
        except ValueError:
            messagebox.showerror("입력 오류", "정확한 숫자를 입력해 주세요.")
        except ZeroDivisionError:
            messagebox.showerror("오류", "외경은 0일 수 없습니다.")

    def calc_rect(self):
        try:
            w_mm = float(self.rect_w_var.get())
            h_mm = float(self.rect_h_var.get())
            val = float(self.rect_var_input.get())
            
            w_m = w_mm / 1000.0
            h_m = h_mm / 1000.0
            circumference = 2 * (w_m + h_m)
            
            if self.r_model_var.get() == "area":
                # L -> Area
                area = circumference * val
                self.rect_result_label.config(text=f"{area:.4f} m²")
            else:
                # Area -> L
                if circumference == 0: raise ZeroDivisionError
                length = val / circumference
                self.rect_result_label.config(text=f"{length:.4f} m")
        except ValueError:
            messagebox.showerror("입력 오류", "정확한 숫자를 입력해 주세요.")
        except ZeroDivisionError:
            messagebox.showerror("오류", "둘레가 0일 수 없습니다.")

    def calc_plate(self):
        try:
            w_mm = float(self.plate_w_var.get())
            val = float(self.plate_var_input.get())
            
            w_m = w_mm / 1000.0
            
            if self.p_model_var.get() == "area":
                # Height(mm) -> Area
                h_m = val / 1000.0
                area = w_m * h_m
                self.plate_result_label.config(text=f"{area:.4f} m²")
            else:
                # Area -> Height(mm)
                if w_m == 0: raise ZeroDivisionError
                # H_m = Area / W_m
                h_m = val / w_m
                h_mm = h_m * 1000.0
                self.plate_result_label.config(text=f"{h_mm:.1f} mm")
        except ValueError:
            messagebox.showerror("입력 오류", "정확한 숫자를 입력해 주세요.")
        except ZeroDivisionError:
            messagebox.showerror("오류", "가로는 0일 수 없습니다.")

if __name__ == "__main__":
    root = tk.Tk()
    root.option_add("*TRadiobutton.background", "#1e1e1e")
    root.option_add("*TRadiobutton.foreground", "#ffffff")
    app = PipeAreaCalculator(root)
    root.mainloop()
