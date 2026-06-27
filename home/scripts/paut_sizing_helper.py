import math

def calculate_h(z1, z2):
    """
    Calculates height 'h' between two depth points.
    """
    return abs(z1 - z2)

def calculate_l(x1, x2):
    """
    Calculates length 'l' between two position points.
    """
    return abs(x1 - x2)

def show_6db_rule(peak_fsh):
    """
    Displays the 6dB drop rule for a given peak FSH.
    """
    drop_point = peak_fsh / 2
    print(f"--- PAUT 6dB Drop Rule ---")
    print(f"Peak Amplitude: {peak_fsh}% FSH")
    print(f"6dB Drop Point (-6dB): {drop_point}% FSH")
    print(f"Definition: Height 'h' is the distance between the upper and lower {drop_point}% FSH points.")
    print(f"--------------------------")

if __name__ == "__main__":
    import argparse
    import sys
    
    parser = argparse.ArgumentParser(description="PAUT 6dB Drop Sizing Helper")
    parser.add_argument("--z1", type=float, help="Depth point 1 (mm)")
    parser.add_argument("--z2", type=float, help="Depth point 2 (mm)")
    parser.add_argument("--x1", type=float, help="Position point 1 (mm)")
    parser.add_argument("--x2", type=float, help="Position point 2 (mm)")
    parser.add_argument("--peak", type=float, help="Peak Amplitude (%% FSH) to see 6dB drop point")

    args = parser.parse_args()

    # If no arguments provided, launch GUI
    if len(sys.argv) == 1:
        import tkinter as tk
        from tkinter import messagebox

        def update_target(*args):
            try:
                p_val = peak_var.get()
                db_val = db_var.get()
                if p_val and db_val:
                    peak = float(p_val)
                    db_drop = float(db_val)
                    drop_factor = pow(10, -db_drop / 20)
                    drop_point = peak * drop_factor
                    lbl_target.config(text=f"강하 목표: {drop_point:.1f}% FSH")
                else:
                    lbl_target.config(text="강하 목표: - %")
            except:
                lbl_target.config(text="강하 목표: - %")

        def calculate():
            try:
                peak = float(peak_var.get())
                db_drop = float(db_var.get())
                z1 = float(ent_z1.get()) if ent_z1.get() else 0
                z2 = float(ent_z2.get()) if ent_z2.get() else 0
                
                drop_factor = pow(10, -db_drop / 20)
                drop_point = peak * drop_factor
                h = abs(z1 - z2)
                
                res_text = f"--- 계산 결과 ---\n"
                res_text += f"피크 높이: {peak}% FSH\n"
                res_text += f"{db_drop}dB 강하 지점: {drop_point:.1f}% FSH\n\n"
                res_text += f"결함 높이 (h): {h:.2f} mm"
                
                lbl_res.config(text=res_text)
            except ValueError:
                messagebox.showerror("오류", "유효한 숫자를 입력해 주세요.")

        root = tk.Tk()
        root.title("PAUT Sizing Helper")
        root.geometry("350x450")
        
        peak_var = tk.StringVar(value="80")
        db_var = tk.StringVar(value="6")
        peak_var.trace_add("write", update_target)
        db_var.trace_add("write", update_target)

        tk.Label(root, text="PAUT Sizing Helper", font=("Arial", 14, "bold")).pack(pady=10)
        
        # Inputs
        tk.Label(root, text="피크 에코 높이 (% FSH):").pack()
        ent_peak = tk.Entry(root, textvariable=peak_var); ent_peak.pack()
        
        tk.Label(root, text="강하 dB 값 (기본 6):").pack(pady=(5,0))
        ent_db = tk.Entry(root, textvariable=db_var); ent_db.pack()
        
        lbl_target = tk.Label(root, text="강하 목표: 40.0% FSH", font=("Arial", 10, "bold"), foreground="#ef4444")
        lbl_target.pack(pady=5)
        
        tk.Label(root, text="상측(Upper) 지점 깊이 (mm):").pack(pady=(15,0))
        ent_z1 = tk.Entry(root); ent_z1.pack()
        
        tk.Label(root, text="하측(Lower) 지점 깊이 (mm):").pack()
        ent_z2 = tk.Entry(root); ent_z2.pack()
        
        tk.Button(root, text="계산하기", command=calculate, bg="#6366f1", fg="white", font=("Arial", 10, "bold"), width=15).pack(pady=20)
        
        lbl_res = tk.Label(root, text="", justify="left", font=("Consolas", 11), bg="#f0f0f0", padx=10, pady=10, relief="sunken")
        lbl_res.pack(fill="x", padx=20)
        
        root.mainloop()
        sys.exit()

    if args.peak:
        show_6db_rule(args.peak)
    
    if args.z1 is not None and args.z2 is not None:
        h = calculate_h(args.z1, args.z2)
        print(f"Calculated Height (h): {h:.2f} mm")
        
    if args.x1 is not None and args.x2 is not None:
        l = calculate_l(args.x1, args.x2)
        print(f"Calculated Length (l): {l:.2f} mm")
