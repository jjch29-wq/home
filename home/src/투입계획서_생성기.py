import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import os
import glob
from datetime import datetime
import re
import openpyxl
from openpyxl.styles import Font, Alignment

CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "personnel_equipment_db.json")
TEMPLATE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates")

def calculate_experience(date_string):
    if not date_string or date_string.strip() == "-":
        return "-"
    # Extract all dates in format YYYY.MM.DD
    dates = re.findall(r"(\d{4})\.(\d{2})\.(\d{2})", date_string)
    if not dates:
        return "-"
    
    earliest_date = min(datetime(int(y), int(m), int(d)) for y, m, d in dates)
    now = datetime.now()
    
    diff_months = (now.year - earliest_date.year) * 12 + now.month - earliest_date.month
    if now.day < earliest_date.day:
        diff_months -= 1
        
    years = diff_months // 12
    months = diff_months % 12
    
    if years > 0 and months > 0:
        return f"{years}년 {months}개월"
    elif years > 0:
        return f"{years}년"
    elif months > 0:
        return f"{months}개월"
    else:
        return "1개월 미만"

class DeploymentApp:
    def __init__(self, root):
        self.root = root
        self.root.title("인원 및 장비투입계획서 자동 생성기")
        self.root.geometry("800x900")
        
        self.load_db()
        self.templates = glob.glob(os.path.join(TEMPLATE_DIR, "양식_*.xlsx"))
        
        self.setup_ui()
        
    def load_db(self):
        self.db = {"personnel": [], "equipment": []}
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    self.db = json.load(f)
            except:
                pass
        self.personnel_names = [p["name"] for p in self.db["personnel"]]
        
    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill='both', expand=True)
        
        ttk.Label(main_frame, text="인원 및 장비투입계획서 생성", font=('Malgun Gothic', 16, 'bold')).pack(pady=(0, 10))
        
        # Template selection
        tmpl_frame = ttk.LabelFrame(main_frame, text="양식(템플릿) 선택", padding=10)
        tmpl_frame.pack(fill='x', pady=5)
        self.tmpl_var = tk.StringVar()
        tmpl_cb = ttk.Combobox(tmpl_frame, textvariable=self.tmpl_var, state='readonly', width=50)
        tmpl_cb['values'] = [os.path.basename(t) for t in self.templates]
        if tmpl_cb['values']:
            tmpl_cb.current(0)
        tmpl_cb.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(tmpl_frame, text="DB 관리 (인원/장비)", command=self.manage_db).pack(side=tk.RIGHT, padx=5)
        
        # Notebook for Tabs
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill='both', expand=True, pady=5)
        
        # --- TAB 1: Personnel ---
        pers_tab = ttk.Frame(notebook)
        notebook.add(pers_tab, text="인력 배치")
        
        pers_frame = ttk.Frame(pers_tab, padding=10)
        pers_frame.pack(fill='both', expand=True)
        
        self.personnel_vars = []
        self.role_vars = []
        self.personnel_comboboxes = []
        
        headers = ["순번", "성명", "담당업무"]
        for i, h in enumerate(headers):
            ttk.Label(pers_frame, text=h, font=('Malgun Gothic', 10, 'bold')).grid(row=0, column=i, padx=5, pady=5)
            
        for i in range(15):
            ttk.Label(pers_frame, text=str(i+1)).grid(row=i+1, column=0, padx=5, pady=2)
            p_var = tk.StringVar()
            cb = ttk.Combobox(pers_frame, textvariable=p_var, state='readonly', values=self.personnel_names, width=15)
            cb.grid(row=i+1, column=1, padx=5, pady=2)
            self.personnel_vars.append(p_var)
            self.personnel_comboboxes.append(cb)
            
            r_var = tk.StringVar()
            rcb = ttk.Combobox(pers_frame, textvariable=r_var, width=25)
            rcb['values'] = ["현장대리인", "방사선안전관리자", "RT 팀장", "PAUT 팀장", "PT 팀장", "MT 팀장", "RT 검사자", "PAUT 검사자", "MT/PT 검사자", "검사보조"]
            rcb.grid(row=i+1, column=2, padx=5, pady=2)
            self.role_vars.append(r_var)
            
        if len(self.personnel_vars) > 0: self.personnel_vars[0].set("주진철"); self.role_vars[0].set("현장대리인")
        if len(self.personnel_vars) > 1: self.personnel_vars[1].set("진병학"); self.role_vars[1].set("방사선안전관리자")
        if len(self.personnel_vars) > 2: self.personnel_vars[2].set("김춘호"); self.role_vars[2].set("RT 팀장")
        
        # --- TAB 2: Equipment ---
        equip_tab = ttk.Frame(notebook)
        notebook.add(equip_tab, text="장비 배치")
        
        # Create scrollable canvas for equipment
        canvas = tk.Canvas(equip_tab)
        scrollbar = ttk.Scrollbar(equip_tab, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        headers_eq = ["선택(수량입력)", "구분", "품명", "규격"]
        for i, h in enumerate(headers_eq):
            ttk.Label(scrollable_frame, text=h, font=('Malgun Gothic', 10, 'bold')).grid(row=0, column=i, padx=5, pady=5)
            
        self.equip_vars = []
        for i, eq in enumerate(self.db.get("equipment", [])):
            qty_var = tk.StringVar(value=eq.get("qty", "0"))
            cat_var = tk.StringVar(value=eq.get("category", ""))
            name_var = tk.StringVar(value=eq.get("name", ""))
            spec_var = tk.StringVar(value=eq.get("spec", ""))
            
            ent_q = ttk.Entry(scrollable_frame, textvariable=qty_var, width=5)
            ent_q.grid(row=i+1, column=0, padx=5, pady=2)
            
            ent_c = ttk.Entry(scrollable_frame, textvariable=cat_var, width=10)
            ent_c.grid(row=i+1, column=1, padx=5, pady=2)
            
            ent_n = ttk.Entry(scrollable_frame, textvariable=name_var, width=20)
            ent_n.grid(row=i+1, column=2, padx=5, pady=2)
            
            ent_s = ttk.Entry(scrollable_frame, textvariable=spec_var, width=25)
            ent_s.grid(row=i+1, column=3, padx=5, pady=2)
            
            self.equip_vars.append({
                "qty_var": qty_var,
                "cat_var": cat_var,
                "name_var": name_var,
                "spec_var": spec_var
            })
            
        # Buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill='x', pady=10)
        
        ttk.Button(btn_frame, text="엑셀 생성하기", command=self.generate_excel, style='Accent.TButton').pack(side=tk.RIGHT, padx=5)
        
    def manage_db(self):
        top = tk.Toplevel(self.root)
        top.title("데이터베이스 관리")
        top.geometry("700x600")
        
        notebook = ttk.Notebook(top)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # --- Personnel Tab ---
        p_tab = ttk.Frame(notebook)
        notebook.add(p_tab, text="인력 관리")
        
        columns = ("name", "qual", "date")
        tree = ttk.Treeview(p_tab, columns=columns, show="headings", height=15)
        tree.heading("name", text="이름")
        tree.heading("qual", text="자격사항")
        tree.heading("date", text="취득일(형식:YYYY.MM.DD(RT))")
        tree.column("name", width=80)
        tree.column("qual", width=200)
        tree.column("date", width=250)
        tree.pack(fill='both', expand=True, padx=10, pady=10)
        
        def refresh_tree():
            for item in tree.get_children(): tree.delete(item)
            for p in self.db["personnel"]: tree.insert("", "end", values=(p["name"], p.get("qualifications", ""), p.get("date", "")))
        refresh_tree()
        
        p_input = ttk.Frame(p_tab)
        p_input.pack(fill='x', padx=10, pady=10)
        
        ttk.Label(p_input, text="이름:").grid(row=0, column=0, padx=5)
        name_ent = ttk.Entry(p_input, width=8)
        name_ent.grid(row=0, column=1, padx=5)
        ttk.Label(p_input, text="자격:").grid(row=0, column=2, padx=5)
        qual_ent = ttk.Entry(p_input, width=15)
        qual_ent.grid(row=0, column=3, padx=5)
        ttk.Label(p_input, text="취득일:").grid(row=0, column=4, padx=5)
        date_ent = ttk.Entry(p_input, width=20)
        date_ent.grid(row=0, column=5, padx=5)
        
        def add_p():
            n = name_ent.get().strip()
            if not n: return
            found = False
            for p in self.db["personnel"]:
                if p["name"] == n:
                    p["qualifications"] = qual_ent.get().strip()
                    p["date"] = date_ent.get().strip()
                    found = True; break
            if not found: self.db["personnel"].append({"name": n, "qualifications": qual_ent.get().strip(), "date": date_ent.get().strip()})
            self.save_db(); refresh_tree()
            self.personnel_names = [p["name"] for p in self.db["personnel"]]
            for cb in self.personnel_comboboxes: cb['values'] = self.personnel_names
        
        def del_p():
            selected = tree.selection()
            if not selected: return
            name_to_del = tree.item(selected[0])['values'][0]
            if messagebox.askyesno("삭제", f"{name_to_del} 님을 삭제하시겠습니까?", parent=top):
                self.db["personnel"] = [p for p in self.db["personnel"] if p["name"] != name_to_del]
                self.save_db(); refresh_tree()
                self.personnel_names = [p["name"] for p in self.db["personnel"]]
                for cb in self.personnel_comboboxes: cb['values'] = self.personnel_names
                
        ttk.Button(p_input, text="저장", command=add_p).grid(row=0, column=6, padx=5)
        ttk.Button(p_input, text="삭제", command=del_p).grid(row=0, column=7, padx=5)
        
        def on_p_sel(e):
            sel = tree.selection()
            if sel:
                item = tree.item(sel[0])['values']
                name_ent.delete(0, tk.END); name_ent.insert(0, item[0])
                qual_ent.delete(0, tk.END); qual_ent.insert(0, item[1])
                date_ent.delete(0, tk.END); date_ent.insert(0, item[2])
        tree.bind('<<TreeviewSelect>>', on_p_sel)
        
        # --- Equipment Tab ---
        e_tab = ttk.Frame(notebook)
        notebook.add(e_tab, text="장비 관리")
        
        e_cols = ("cat", "name", "spec", "qty")
        e_tree = ttk.Treeview(e_tab, columns=e_cols, show="headings", height=15)
        e_tree.heading("cat", text="구분")
        e_tree.heading("name", text="품명")
        e_tree.heading("spec", text="규격")
        e_tree.heading("qty", text="기본수량")
        e_tree.column("cat", width=80)
        e_tree.column("name", width=150)
        e_tree.column("spec", width=200)
        e_tree.column("qty", width=60)
        e_tree.pack(fill='both', expand=True, padx=10, pady=10)
        
        def refresh_e_tree():
            for item in e_tree.get_children(): e_tree.delete(item)
            for eq in self.db.get("equipment", []):
                e_tree.insert("", "end", values=(eq.get("category",""), eq.get("name",""), eq.get("spec",""), eq.get("qty","")))
        refresh_e_tree()
        
        e_input = ttk.Frame(e_tab)
        e_input.pack(fill='x', padx=10, pady=10)
        
        ttk.Label(e_input, text="품명:").grid(row=0, column=0, padx=5)
        e_name_ent = ttk.Entry(e_input, width=15)
        e_name_ent.grid(row=0, column=1, padx=5)
        ttk.Label(e_input, text="구분:").grid(row=0, column=2, padx=5)
        e_cat_ent = ttk.Entry(e_input, width=10)
        e_cat_ent.grid(row=0, column=3, padx=5)
        ttk.Label(e_input, text="규격:").grid(row=0, column=4, padx=5)
        e_spec_ent = ttk.Entry(e_input, width=15)
        e_spec_ent.grid(row=0, column=5, padx=5)
        ttk.Label(e_input, text="수량:").grid(row=0, column=6, padx=5)
        e_qty_ent = ttk.Entry(e_input, width=5)
        e_qty_ent.grid(row=0, column=7, padx=5)
        
        def add_e():
            n = e_name_ent.get().strip()
            if not n: return
            found = False
            for eq in self.db.get("equipment", []):
                if eq["name"] == n:
                    eq["category"] = e_cat_ent.get().strip()
                    eq["spec"] = e_spec_ent.get().strip()
                    eq["qty"] = e_qty_ent.get().strip()
                    found = True; break
            if not found:
                self.db.setdefault("equipment", []).append({
                    "name": n, "category": e_cat_ent.get().strip(),
                    "spec": e_spec_ent.get().strip(), "qty": e_qty_ent.get().strip()
                })
            self.save_db(); refresh_e_tree()
            messagebox.showinfo("저장 완료", "장비 정보가 저장되었습니다. (메인 화면 갱신은 프로그램 재시작 필요)", parent=top)
            
        def del_e():
            selected = e_tree.selection()
            if not selected: return
            name_to_del = e_tree.item(selected[0])['values'][1]
            if messagebox.askyesno("삭제", f"{name_to_del} 장비를 삭제하시겠습니까?", parent=top):
                self.db["equipment"] = [eq for eq in self.db.get("equipment", []) if eq["name"] != name_to_del]
                self.save_db(); refresh_e_tree()
                
        ttk.Button(e_input, text="저장", command=add_e).grid(row=0, column=8, padx=5)
        ttk.Button(e_input, text="삭제", command=del_e).grid(row=0, column=9, padx=5)
        
        def on_e_sel(e):
            sel = e_tree.selection()
            if sel:
                item = e_tree.item(sel[0])['values']
                e_cat_ent.delete(0, tk.END); e_cat_ent.insert(0, item[0])
                e_name_ent.delete(0, tk.END); e_name_ent.insert(0, item[1])
                e_spec_ent.delete(0, tk.END); e_spec_ent.insert(0, item[2])
                e_qty_ent.delete(0, tk.END); e_qty_ent.insert(0, item[3])
        e_tree.bind('<<TreeviewSelect>>', on_e_sel)

    def save_db(self):
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(self.db, f, indent=2, ensure_ascii=False)

    def get_person_info(self, name):
        for p in self.db["personnel"]:
            if p["name"] == name:
                return p
        return None

    def generate_excel(self):
        selected_tmpl = self.tmpl_var.get()
        if not selected_tmpl:
            messagebox.showerror("오류", "템플릿을 선택하세요.")
            return
            
        tmpl_path = os.path.join(TEMPLATE_DIR, selected_tmpl)
        
        # 파일 저장 위치 묻기
        default_name = f"생성_{selected_tmpl}"
        out_path = filedialog.asksaveasfilename(
            title="엑셀 파일 저장 위치 선택",
            initialfile=default_name,
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        
        if not out_path: # 사용자가 저장을 취소한 경우
            return
            
        try:
            wb = openpyxl.load_workbook(tmpl_path)
            
            # Prepare selected personnel data
            personnel_data = []
            for i in range(15):
                name = self.personnel_vars[i].get()
                role = self.role_vars[i].get()
                if name:
                    p_info = self.get_person_info(name)
                    if p_info:
                        exp = calculate_experience(p_info.get("date", ""))
                        personnel_data.append({
                            "name": p_info["name"],
                            "role": role,
                            "exp": exp,
                            "qual": p_info.get("qualifications", ""),
                            "date": p_info.get("date", "")
                        })

            # Fill personnel list sheets
            for sheet_name in wb.sheetnames:
                if "인력투입계획서" in sheet_name or "인력투입변경" in sheet_name:
                    ws = wb[sheet_name]
                    # Find starting row
                    start_row = -1
                    name_col = -1
                    for row in ws.iter_rows(min_row=1, max_row=20):
                        for cell in row:
                            if cell.value == "성명":
                                start_row = cell.row + 1
                                name_col = cell.column
                                break
                        if start_row != -1: break
                    
                    if start_row != -1:
                        # 자격사항과 취득년월일 열 위치를 먼저 찾습니다. (양식마다 다름)
                        qual_col = -1
                        date_col = -1
                        for cell in ws[start_row-1]:
                            if cell.value and "자격" in str(cell.value): qual_col = cell.column
                            if cell.value and ("취득" in str(cell.value) or "년월일" in str(cell.value)): date_col = cell.column
                            
                        # 1. 먼저 기존 데이터(최대 20줄)를 싹 지웁니다.
                        for r in range(start_row, start_row + 20):
                            ws.cell(row=r, column=name_col-1).value = ""
                            ws.cell(row=r, column=name_col).value = ""
                            ws.cell(row=r, column=name_col+1).value = ""
                            ws.cell(row=r, column=name_col+2).value = ""
                            if qual_col != -1: ws.cell(row=r, column=qual_col).value = ""
                            if date_col != -1: ws.cell(row=r, column=date_col).value = ""
                            
                        # 2. 선택된 인원 데이터만 새로 씁니다.
                        for idx, p in enumerate(personnel_data):
                            row = start_row + idx
                            ws.cell(row=row, column=name_col-1).value = idx + 1 # 순번
                            ws.cell(row=row, column=name_col).value = p["name"]
                            ws.cell(row=row, column=name_col+1).value = p["role"]
                            ws.cell(row=row, column=name_col+2).value = p["exp"]
                            
                            wrap_alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
                            
                            # 자격사항이 길면 슬래시(/)를 줄바꿈으로 변경
                            qual_text = p["qual"].replace("/", "\n")
                            date_text = p["date"].replace("/", "\n")
                            
                            # 엑셀 고유의 '행 높이 자동 맞춤' 기능이 작동하도록 기존 고정 높이를 해제합니다.
                            ws.row_dimensions[row].height = None
                            
                            
                            if qual_col != -1:
                                cell = ws.cell(row=row, column=qual_col)
                                cell.value = qual_text
                                cell.alignment = wrap_alignment
                            if date_col != -1:
                                cell = ws.cell(row=row, column=date_col)
                                cell.value = date_text
                                cell.alignment = wrap_alignment

                # Fill org chart sheets
                if "조직도" in sheet_name:
                    ws = wb[sheet_name]
                    
                    role_keywords = ["대리인", "안전관리자", "책임", "팀장", "검사원", "검사보조", "PAUT", "RT", "MT", "PT", "PMI", "품질"]
                    
                    def is_name_cell(val_str):
                        # 조직도에서 직책 바로 밑칸에 적힌 값(이름)인지 판별
                        if not val_str: return False
                        vs = str(val_str).replace(" ", "")
                        # 직책 키워드가 포함되어 있으면 이름칸이 아님
                        if any(k in vs for k in ["대리인", "관리자", "팀장", "검사원", "검사자"]):
                            return False
                        return True
                        
                    def match_role(ui_role, cell_val):
                        ur = ui_role.replace(" ", "")
                        cv = cell_val.replace(" ", "")
                        if ur in cv or cv in ur: return True
                        
                        # 템플릿의 '책임기술자'와 UI의 '팀장' 매핑
                        if "팀장" in ur and "책임" in cv:
                            prefix = ur.replace("팀장", "")
                            if prefix and prefix in cv: return True
                            if "책임기술자" == cv: return True # 접두어가 없어도 허용
                            
                        # 템플릿의 '검사원'과 UI의 '검사자' 매핑
                        if "검사자" in ur and "검사원" in cv:
                            prefix = ur.replace("검사자", "")
                            if prefix and prefix in cv: return True
                            
                        return False

                    # 1. 먼저 조직도 내에 있는 기존 이름들을 모두 지웁니다 (직책 키워드 기준)
                    for row in ws.iter_rows():
                        for cell in row:
                            if cell.value and isinstance(cell.value, str):
                                val = cell.value.strip()
                                
                                # '사업책임기술자'를 '현장대리인'으로 텍스트 자동 변경
                                if "사업책임기술자" in val:
                                    val = val.replace("사업책임기술자", "현장대리인")
                                    cell.value = val
                                
                                if any(kw in val for kw in role_keywords):
                                    target_cell = ws.cell(row=cell.row+1, column=cell.column)
                                    if target_cell.value and is_name_cell(target_cell.value):
                                        target_cell.value = ""

                    # 2. 선택된 인원의 이름을 찾아 꽂아넣습니다.
                    for row in ws.iter_rows():
                        for cell in row:
                            if cell.value and isinstance(cell.value, str):
                                val = cell.value.strip()
                                for p in personnel_data:
                                    if match_role(p["role"], val):
                                        target_cell = ws.cell(row=cell.row+1, column=cell.column)
                                        
                                        # 이미 누군가 배정되었으면(글자가 있으면) 건너뛰기
                                        if target_cell.value and str(target_cell.value).strip():
                                            continue
                                            
                                        formatted_name = p["name"]
                                        if len(formatted_name) == 3:
                                            formatted_name = f"{formatted_name[0]}   {formatted_name[1]}   {formatted_name[2]}"
                                        target_cell.value = formatted_name
                                        break
                                        
                # Fill equipment sheets
                if "장비" in sheet_name:
                    ws = wb[sheet_name]
                    
                    # Gather selected equipment from UI
                    equip_data = []
                    for ev in self.equip_vars:
                        try:
                            q = int(ev["qty_var"].get())
                            if q > 0:
                                equip_data.append({
                                    "category": ev["cat_var"].get(),
                                    "name": ev["name_var"].get(),
                                    "spec": ev["spec_var"].get(),
                                    "qty": q
                                })
                        except ValueError:
                            pass
                            
                    start_row = -1
                    name_col = -1
                    
                    # Find header row containing "품명" or "규격"
                    for row in ws.iter_rows(min_row=1, max_row=20):
                        for cell in row:
                            if cell.value and "품명" in str(cell.value):
                                start_row = cell.row + 1
                                name_col = cell.column
                                break
                        if start_row != -1: break
                        
                    if start_row != -1 and name_col != -1:
                        cat_col = name_col - 1
                        spec_col = name_col + 1
                        qty_col = name_col + 2
                        time_col = name_col + 3
                        
                        # 1. Clear existing data (up to 30 rows)
                        for r in range(start_row, start_row + 30):
                            ws.cell(row=r, column=cat_col).value = ""
                            ws.cell(row=r, column=name_col).value = ""
                            ws.cell(row=r, column=spec_col).value = ""
                            ws.cell(row=r, column=qty_col).value = ""
                            ws.cell(row=r, column=time_col).value = ""
                            
                        # 2. Insert new data
                        for idx, eq in enumerate(equip_data):
                            row = start_row + idx
                            ws.cell(row=row, column=cat_col).value = eq["category"]
                            ws.cell(row=row, column=name_col).value = eq["name"]
                            ws.cell(row=row, column=spec_col).value = eq["spec"]
                            ws.cell(row=row, column=qty_col).value = str(eq["qty"])
                            ws.cell(row=row, column=time_col).value = "착공시 ~ 준공시"
                            
                            wrap_alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
                            for c in [cat_col, name_col, spec_col, qty_col, time_col]:
                                ws.cell(row=row, column=c).alignment = wrap_alignment
                                        
            wb.save(out_path)
            messagebox.showinfo("완료", f"성공적으로 생성되었습니다.\n저장 위치: {out_path}")
        except Exception as e:
            messagebox.showerror("오류", f"생성 중 오류 발생:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    
    # Custom styles
    style = ttk.Style()
    if 'clam' in style.theme_names():
        style.theme_use('clam')
    style.configure('Accent.TButton', font=('Malgun Gothic', 10, 'bold'), background='#0052cc', foreground='white')
    
    app = DeploymentApp(root)
    root.mainloop()
