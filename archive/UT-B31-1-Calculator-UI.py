import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
from datetime import datetime

# ==========================================
# ⚙️ Core Evaluation Logic (ASME B31.1)
# ==========================================
def evaluate_paut_flaw(t, h, l, depth, flaw_nature, flaw_location=None):
    """
    ASME B31.1 (2024) - UT Acceptance Criteria Logic
    Automatically determines location if depth is provided.
    Includes special rules for 6mm <= t < 13mm.
    """
    try:
        t = float(t)
        h = float(h)
        l = float(l)
        d = float(depth)
    except ValueError:
        return "Error (Invalid Dimension)", "Unknown"

    # 0. Automatic Location Determination
    s_top = d
    s_bottom = t - (d + h)
    s = min(s_top, s_bottom)
    s_limit = 0.4 * (h / 2)
    
    loc = "Surface" if s <= s_limit else "Subsurface"

    # 1. Immediate Rejection (Crack, LOF, IP)
    unacceptable_types = ['crack', 'lof', 'lack of fusion', 'ip', 'incomplete penetration']
    if str(flaw_nature).strip().lower() in unacceptable_types:
        return "Reject (Fatal Flaw Type)", loc
    
    if l <= 0 or h <= 0 or t <= 0:
        return "Error (Zero/Negative Value)", loc

    # 1.1 Special Rules for 6mm <= t < 13mm
    if 6 <= t < 13:
        # Length limit (Common for all 6-13mm)
        if l > 6.4:
            return f"Reject (L: {l} > 6.4mm)", loc
            
        # Height limits (Stepped based on thickness)
        if t < 10:
            h_surf_max, h_sub_max = 0.95, 0.48 * 2  # 0.96
        elif t < 12:
            h_surf_max, h_sub_max = 1.04, 0.52 * 2  # 1.04
        else: # 12 <= t < 13
            h_surf_max, h_sub_max = 1.13, 0.57 * 2  # 1.14
            
        limit = h_surf_max if loc == "Surface" else h_sub_max
        if h > limit:
            return f"Reject ({loc} h: {h} > {limit}mm)", loc
        else:
            return "Accept", loc

    # 1.2 Special Rules for 13mm <= t < 25.4mm
    if 13 <= t < 25.4:
        # Length limit
        if l > 6.4:
            return f"Reject (L: {l} > 6.4mm)", loc
        
        # Height ratio limits (fixed h/t based on user preference)
        # Surface max h/t = 0.087, Subsurface max h/t = 0.143
        actual_h_t = h / t
        allowed_h_t = 0.087 if loc == "Surface" else 0.143 
        
        if actual_h_t > allowed_h_t:
            return f"Reject ({loc} h/t: {actual_h_t:.3f} > {allowed_h_t:.3f})", loc
        else:
            return "Accept", loc
            
    # 2. Aspect Ratio (a/l) logic for t >= 25.4mm
    # For Surface: a = h, For Subsurface: a = h/2 (per user confirmation)
    a_val = h if loc == "Surface" else h / 2
    aspect_ratio_a_l = a_val / l
    
    # New Master Table (25.4mm <= t <= 64mm)
    # [a/l Limit, Surface a/t Limit, Subsurface a/t Limit]
    master_table = [
        (0.00, 0.031, 0.034),
        (0.05, 0.033, 0.038),
        (0.10, 0.036, 0.043),
        (0.15, 0.041, 0.054),
        (0.20, 0.047, 0.066),
        (0.25, 0.055, 0.078),
        (0.30, 0.064, 0.090),
        (0.35, 0.074, 0.103),
        (0.40, 0.083, 0.116),
        (0.45, 0.085, 0.129),
        (0.50, 0.087, 0.143)
    ]
    
    # Find a/t Limit from table
    allowed_a_t = 0
    for ar_limit, surf_a_t, sub_a_t in master_table:
        if aspect_ratio_a_l <= ar_limit:
            allowed_a_t = surf_a_t if loc == 'Surface' else sub_a_t
            break
            
    if allowed_a_t == 0:
        allowed_a_t = master_table[-1][1] if loc == 'Surface' else master_table[-1][2]

    # Final Evaluation (a/t comparison)
    actual_a_t = a_val / t
    
    if actual_a_t <= allowed_a_t:
        return "Accept", loc
    else:
        return f"Reject ({loc} a/t: {actual_a_t:.3f} > {allowed_a_t:.3f})", loc

# ==========================================
# 🖥️ GUI Application
# ==========================================
class UTApp:
    def __init__(self, root):
        self.root = root
        self.root.title("UT B31.1 Acceptance Criteria Tool")
        self.root.geometry("950x750")
        self.root.configure(bg="#f8f9fa")
        
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Define modern styles
        self.style.configure("TFrame", background="#f8f9fa")
        self.style.configure("TLabel", background="#f8f9fa", font=("Segoe UI", 10))
        self.style.configure("Header.TLabel", font=("Segoe UI", 18, "bold"), foreground="#2c3e50")
        self.style.configure("TButton", font=("Segoe UI", 10, "bold"), padding=5)
        
        self.df = None  
        self.create_widgets()

    def create_widgets(self):
        # Header
        header_frame = ttk.Frame(self.root, style="TFrame")
        header_frame.pack(fill="x", padx=20, pady=20)
        
        header_label = ttk.Label(header_frame, text="PAUT 합부 판정 시스템 (ASME B31.1)", style="Header.TLabel")
        header_label.pack(side="left")
        
        sub_label = ttk.Label(header_frame, text="UT Acceptance Criteria Tool v1.1", foreground="#7f8c8d")
        sub_label.pack(side="right", pady=10)

        # Tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=20, pady=10)
        
        self.tab1 = ttk.Frame(self.notebook)
        self.tab2 = ttk.Frame(self.notebook)
        
        self.notebook.add(self.tab1, text="  개별 판정 (Manual)  ")
        self.notebook.add(self.tab2, text="  일괄 판정 (Batch)  ")
        
        self.setup_tab1()
        self.setup_tab2()

    def setup_tab1(self):
        # Input Section
        input_frame = tk.LabelFrame(self.tab1, text=" 검사 데이터 입력 ", font=("Segoe UI", 10, "bold"), padx=20, pady=20, bg="white")
        input_frame.pack(fill="x", padx=40, pady=10)
        
        input_labels = [
            ("모재 두께 (t) (mm):", "t"),
            ("결함 높이 (h) (mm):", "h"),
            ("결함 길이 (l) (mm):", "l"),
            ("결함 깊이 (d) (mm):", "d"),  # New Field
            ("결함 위치 (자동):", "location"),
            ("결함 종류:", "nature")
        ]
        self.entries = {}
        
        for i, (label_text, key) in enumerate(input_labels):
            ttk.Label(input_frame, text=label_text, background="white").grid(row=i, column=0, sticky="w", pady=8)
            
            if key == "location":
                self.loc_var = tk.StringVar(value="Subsurface")
                lbl = ttk.Label(input_frame, textvariable=self.loc_var, foreground="#2980b9", font=("Segoe UI", 10, "bold"), background="white")
                lbl.grid(row=i, column=1, padx=20, sticky="w")
                self.entries[key] = self.loc_var
            elif key == "nature":
                var = tk.StringVar(value="Slag")
                cb = ttk.Combobox(input_frame, textvariable=var, values=["Crack", "LOF", "IP", "Slag", "Porosity", "Others"], width=27)
                cb.grid(row=i, column=1, padx=20, sticky="w")
                self.entries[key] = var
            else:
                entry = ttk.Entry(input_frame, width=30)
                entry.grid(row=i, column=1, padx=20, sticky="w")
                self.entries[key] = entry
                # Add auto-update trace
                entry.bind("<KeyRelease>", self.update_auto_location)

        # Action Button
        eval_btn = ttk.Button(self.tab1, text="판정 시작", command=self.manual_evaluate)
        eval_btn.pack(pady=10)

        # Result Display
        self.result_frame = tk.Frame(self.tab1, bg="#ecf0f1", height=130)
        self.result_frame.pack(fill="x", padx=40, pady=10)
        self.result_frame.pack_propagate(False)
        
        self.result_label = tk.Label(self.result_frame, text="데이터를 입력하고 '판정 시작' 버튼을 누르세요.", font=("Segoe UI", 13), bg="#ecf0f1", fg="#7f8c8d")
        self.result_label.pack(expand=True)

    def update_auto_location(self, event=None):
        try:
            t = float(self.entries["t"].get())
            h = float(self.entries["h"].get())
            d = float(self.entries["d"].get())
            
            s_top = d
            s_bottom = t - (d + h)
            s = min(s_top, s_bottom)
            s_limit = 0.4 * (h / 2)
            
            loc = "Surface" if s <= s_limit else "Subsurface"
            self.loc_var.set(loc)
        except:
            self.loc_var.set("-")

    def manual_evaluate(self):
        try:
            t = self.entries["t"].get()
            h = self.entries["h"].get()
            l = self.entries["l"].get()
            d = self.entries["d"].get()
            nat = self.entries["nature"].get()
            
            if not all([t, h, l, d]):
                messagebox.showwarning("입력 누락", "모든 치수 및 깊이 정보를 입력해 주세요.")
                return
                
            res, loc = evaluate_paut_flaw(t, h, l, d, nat)
            self.loc_var.set(loc)
            
            if res == "Accept":
                self.result_label.config(text=f"● ACCEPT (합격) - [{loc}]", fg="white", bg="#27ae60", font=("Segoe UI", 18, "bold"))
                self.result_frame.config(bg="#27ae60")
            elif "Reject" in res:
                self.result_label.config(text=f"● REJECT (불합격) - [{loc}]\n{res}", fg="white", bg="#e74c3c", font=("Segoe UI", 14, "bold"))
                self.result_frame.config(bg="#e74c3c")
            else:
                self.result_label.config(text=res, fg="black", bg="#f1c40f")
                self.result_frame.config(bg="#f1c40f")
                
        except Exception as e:
            messagebox.showerror("에러", f"판정 중 오류가 발생했습니다: {e}")

    def setup_tab2(self):
        top_frame = ttk.Frame(self.tab2)
        top_frame.pack(fill="x", padx=20, pady=10)
        
        ttk.Button(top_frame, text="엑셀 파일 불러오기", command=self.load_excel).pack(side="left", padx=5)
        ttk.Button(top_frame, text="결과 저장 (Excel)", command=self.save_excel).pack(side="left", padx=5)
        
        self.progress_label = ttk.Label(top_frame, text="파일을 선택해 주세요.")
        self.progress_label.pack(side="right", padx=10)

        table_frame = ttk.Frame(self.tab2)
        table_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        columns = ("Index", "ISO", "Joint", "t", "h", "l", "Depth", "Location", "Nature", "Result")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings")
        
        for col in columns:
            self.tree.heading(col, text=col)
            width = 40 if col == "Index" else (100 if col in ["ISO", "Result"] else 80)
            self.tree.column(col, width=width, anchor="center")
            
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def load_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if not file_path:
            return
            
        try:
            self.df = pd.read_excel(file_path)
            self.progress_label.config(text=f"불러온 파일: {os.path.basename(file_path)}")
            
            cols = self.df.columns.tolist()
            mapping = {
                "t": next((c for c in cols if "THK" in c.upper() or "두께" in c or "Thickness" in c), None),
                "h": next((c for c in cols if "HEIGHT" in c.upper() or "높이" in c or "Flaw_Height" in c), None),
                "l": next((c for c in cols if "LENGTH" in c.upper() or "길이" in c or "Flaw_Length" in c), None),
                "d": next((c for c in cols if "DEPTH" in c.upper() or "깊이" in c or "Flaw_Depth" in c), None),
                "nat": next((c for c in cols if "NAT" in c.upper() or "종류" in c or "Flaw_Nature" in c), None),
                "iso": next((c for c in cols if "ISO" in c.upper() or "DWG" in c.upper()), "ISO"),
                "joint": next((c for c in cols if "JOINT" in c.upper() or "WELD" in c.upper()), "Joint")
            }
            
            if not all([mapping["t"], mapping["h"], mapping["l"], mapping["d"]]):
                messagebox.showerror("매핑 오류", "필요한 컬럼(두께, 높이, 길이, 깊이)을 모두 찾을 수 없습니다.")
                return

            for item in self.tree.get_children():
                self.tree.delete(item)
                
            results = []
            final_locations = []
            for i, row in self.df.iterrows():
                res, loc = evaluate_paut_flaw(
                    row.get(mapping["t"], 0),
                    row.get(mapping["h"], 0),
                    row.get(mapping["l"], 0),
                    row.get(mapping["d"], 0),
                    str(row.get(mapping["nat"], "Slag"))
                )
                results.append(res)
                final_locations.append(loc)
                
                self.tree.insert("", "end", values=(
                    i + 1,
                    row.get(mapping["iso"], "-"),
                    row.get(mapping["joint"], "-"),
                    row.get(mapping["t"], 0),
                    row.get(mapping["h"], 0),
                    row.get(mapping["l"], 0),
                    row.get(mapping["d"], 0),
                    loc,
                    row.get(mapping["nat"], "-"),
                    res
                ))
            
            self.df['Determined_Location'] = final_locations
            self.df['Evaluation_Result'] = results
            self.progress_label.config(text=f"처리 완료: {len(self.df)}건")
            
        except Exception as e:
            messagebox.showerror("에러", f"파일 처리 중 오류: {e}")

    def save_excel(self):
        if self.df is None or 'Evaluation_Result' not in self.df.columns:
            messagebox.showwarning("데이터 없음", "먼저 판정을 완료해 주세요.")
            return
            
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if not file_path:
            return
            
        try:
            self.df.to_excel(file_path, index=False)
            messagebox.showinfo("저장 완료", f"저장되었습니다: {os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("에러", f"저장 실패: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = UTApp(root)
    root.mainloop()
