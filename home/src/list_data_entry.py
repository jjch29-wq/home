import tkinter as tk
from tkinter import ttk, messagebox
import openpyxl
from openpyxl.styles import Border, Side, Alignment
import os
import re

class DataEntryUI(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("⚡ 종합 리스트 초고속 입력기")
        self.geometry("900x650")
        self.attributes('-topmost', True)
        
        self.excel_path = r'C:\Users\jjch2\Desktop\중앙지사 종합 리스트.xlsx'
        
        # State variables
        self.selected_sheet = tk.StringVar(value="PAUT")
        
        # Fields common to all
        self.company_var = tk.StringVar(value="세경")
        self.type_var = tk.StringVar()
        self.line_var = tk.StringVar()
        self.joint_var = tk.StringVar()
        self.welder_var = tk.StringVar()
        self.size_var = tk.StringVar()
        
        # Standard only
        self.length_var = tk.StringVar()
        self.result_var = tk.StringVar(value="합격")
        
        # RT only
        self.date_var = tk.StringVar()
        self.section_var = tk.StringVar()
        self.shoot_part_var = tk.StringVar(value="1")
        self.film_size_var = tk.StringVar(value='3 1/3 x 12"')
        self.film_count_var = tk.StringVar(value="4")
        self.rt_r_var = tk.StringVar()
        self.rt_result_var = tk.StringVar(value="합격")
        self.rt_remark_var = tk.StringVar()
        
        # Sticky checkboxes
        self.sticky_company = tk.BooleanVar(value=True)
        self.sticky_type = tk.BooleanVar(value=True)
        self.sticky_line = tk.BooleanVar(value=True)
        self.sticky_welder = tk.BooleanVar(value=True)
        self.sticky_size = tk.BooleanVar(value=True)
        self.sticky_date = tk.BooleanVar(value=True)
        self.sticky_section = tk.BooleanVar(value=True)
        self.sticky_film = tk.BooleanVar(value=True)
        
        self._build_ui()
        self._update_form_layout()
        
    def _build_ui(self):
        # 1. Top Bar: Sheet selection
        top_frame = ttk.LabelFrame(self, text="1. 입력할 시트 선택")
        top_frame.pack(fill='x', padx=10, pady=5)
        
        for sheet in ["PAUT", "RT", "PT", "MT"]:
            ttk.Radiobutton(top_frame, text=sheet, value=sheet, variable=self.selected_sheet, command=self._update_form_layout).pack(side='left', padx=10, pady=5)
            
        # 2. Main Form Area
        self.form_frame = ttk.LabelFrame(self, text="2. 데이터 입력 (기본값 고정 체크시 저장 후에도 값이 유지됩니다)")
        self.form_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        # We will dynamically pack frames into form_frame based on sheet
        self.std_frame = ttk.Frame(self.form_frame)
        self.rt_frame = ttk.Frame(self.form_frame)
        
        self._build_std_form()
        self._build_rt_form()
        
        # 3. Bottom Action Bar
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill='x', padx=10, pady=10)
        
        save_btn = ttk.Button(btn_frame, text="💾 엑셀에 저장 및 다음 입력 (Enter)", command=self._save_data)
        save_btn.pack(side='right', ipadx=20, ipady=10)
        self.bind('<Return>', lambda e: self._save_data())
        
        # 4. Preview Treeview
        preview_frame = ttk.LabelFrame(self, text="3. 최근 입력 기록 (미리보기)")
        preview_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        self.tree = ttk.Treeview(preview_frame, columns=("Sheet", "Line", "Joint", "Result"), show='headings', height=5)
        self.tree.heading("Sheet", text="시트")
        self.tree.heading("Line", text="Line No.")
        self.tree.heading("Joint", text="Joint")
        self.tree.heading("Result", text="결과")
        self.tree.pack(fill='both', expand=True, padx=5, pady=5)

    def _add_field(self, parent, row, label, var, sticky_var=None, is_combo=False, values=None, width=20):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky='e', padx=5, pady=5)
        if is_combo:
            w = ttk.Combobox(parent, textvariable=var, values=values, width=width)
        else:
            w = ttk.Entry(parent, textvariable=var, width=width)
        w.grid(row=row, column=1, sticky='w', padx=5, pady=5)
        
        if sticky_var is not None:
            ttk.Checkbutton(parent, text="고정", variable=sticky_var).grid(row=row, column=2, sticky='w')
            
        return w

    def _build_std_form(self):
        # Fields: 제조사, 구분, Line No, Joint, 용접사번호, 관경, 용접법, 결과
        self._add_field(self.std_frame, 0, "제조사명:", self.company_var, self.sticky_company)
        self._add_field(self.std_frame, 1, "구분:", self.type_var, self.sticky_type)
        self.std_line_entry = self._add_field(self.std_frame, 2, "Line No.:", self.line_var, self.sticky_line, width=40)
        self.std_joint_entry = self._add_field(self.std_frame, 3, "Joint No.:", self.joint_var) # Never sticky
        self._add_field(self.std_frame, 4, "용접사 번호:", self.welder_var, self.sticky_welder)
        self._add_field(self.std_frame, 5, "관경 (Size):", self.size_var, self.sticky_size, is_combo=True, values=["300A", "250A", "200A", "150A", "100A", "80A", "50A"])
        self._add_field(self.std_frame, 6, "검사 길이:", self.length_var)
        self._add_field(self.std_frame, 7, "판정 결과(합부):", self.result_var, is_combo=True, values=["합격", "불합격"])

    def _build_rt_form(self):
        # RT Fields: 제조사, 구분, 촬영일자, Section, Line, Joint, 용접사, 촬영구간, 관경, 필름규격, 장수, R, 촬영결과, 비고
        self._add_field(self.rt_frame, 0, "제조사명:", self.company_var, self.sticky_company)
        self._add_field(self.rt_frame, 1, "구분:", self.type_var, self.sticky_type)
        self._add_field(self.rt_frame, 2, "촬영일자:", self.date_var, self.sticky_date)
        self._add_field(self.rt_frame, 3, "Section:", self.section_var, self.sticky_section)
        
        self.rt_line_entry = self._add_field(self.rt_frame, 4, "Line No.:", self.line_var, self.sticky_line, width=40)
        self.rt_joint_entry = self._add_field(self.rt_frame, 5, "Joint No.:", self.joint_var)
        self._add_field(self.rt_frame, 6, "용접사 번호:", self.welder_var, self.sticky_welder)
        
        self._add_field(self.rt_frame, 7, "촬영구간 (1~8):", self.shoot_part_var)
        self._add_field(self.rt_frame, 8, "관경 (Size):", self.size_var, self.sticky_size, is_combo=True, values=["300A", "250A", "200A", "150A", "100A", "80A", "50A"])
        self._add_field(self.rt_frame, 9, "필름 규격:", self.film_size_var, self.sticky_film, is_combo=True, values=['3 1/3 x 12"', '4 1/2 x 17"'])
        self._add_field(self.rt_frame, 10, "필름 장수:", self.film_count_var)
        self._add_field(self.rt_frame, 11, "R (수정횟수):", self.rt_r_var)
        self._add_field(self.rt_frame, 12, "판정 결과:", self.rt_result_var, is_combo=True, values=["합격", "불합격"])
        self._add_field(self.rt_frame, 13, "비고:", self.rt_remark_var)

    def _update_form_layout(self):
        self.std_frame.pack_forget()
        self.rt_frame.pack_forget()
        
        if self.selected_sheet.get() == "RT":
            self.rt_frame.pack(fill='both', expand=True, padx=20, pady=10)
            self.rt_joint_entry.focus()
        else:
            self.std_frame.pack(fill='both', expand=True, padx=20, pady=10)
            self.std_joint_entry.focus()

    def _save_data(self):
        sheet_name = self.selected_sheet.get()
        if not os.path.exists(self.excel_path):
            messagebox.showerror("오류", f"엑셀 파일이 없습니다:\n{self.excel_path}")
            return
            
        try:
            wb = openpyxl.load_workbook(self.excel_path)
            if sheet_name not in wb.sheetnames:
                messagebox.showerror("오류", f"{sheet_name} 시트가 없습니다.")
                return
            ws = wb[sheet_name]
            
            # Find last empty row
            insert_row = 3
            while insert_row <= 2000:
                val1 = ws.cell(row=insert_row, column=2).value # 제조사
                val2 = ws.cell(row=insert_row, column=4).value # Line (PAUT) or 촬영일자 (RT)
                val3 = ws.cell(row=insert_row, column=5).value # Joint or Section
                if not val1 and not val2 and not val3:
                    break
                insert_row += 1
                
            # Prepare row data
            if sheet_name == "RT":
                # ['순번', '제조사', '구분', '촬영\n일자', 'Section', 'Line No.', 'Joint', '용접사\n번호', '촬영구간1~8', '관경', '필름규격', '장수', 'R', '촬영결과', '비고']
                # Columns: A=1(순번), B=2(제조사), C=3(구분), D=4(일자), E=5(Section), F=6(Line), G=7(Joint), H=8(용접사)
                # I=9(구간1), Q=17(관경), R=18(필름), S=19(장수), T=20(R), U=21(결과), V=22(비고)
                
                ws.cell(row=insert_row, column=1).value = insert_row - 2
                ws.cell(row=insert_row, column=2).value = self.company_var.get()
                ws.cell(row=insert_row, column=3).value = self.type_var.get()
                ws.cell(row=insert_row, column=4).value = self.date_var.get()
                ws.cell(row=insert_row, column=5).value = self.section_var.get()
                ws.cell(row=insert_row, column=6).value = self.line_var.get()
                ws.cell(row=insert_row, column=7).value = self.joint_var.get()
                ws.cell(row=insert_row, column=8).value = self.welder_var.get()
                
                # 촬영구간
                part = self.shoot_part_var.get()
                ws.cell(row=insert_row, column=9).value = part
                
                ws.cell(row=insert_row, column=17).value = self.size_var.get()
                ws.cell(row=insert_row, column=18).value = self.film_size_var.get()
                ws.cell(row=insert_row, column=19).value = self.film_count_var.get()
                ws.cell(row=insert_row, column=20).value = self.rt_r_var.get()
                ws.cell(row=insert_row, column=21).value = self.rt_result_var.get()
                ws.cell(row=insert_row, column=22).value = self.rt_remark_var.get()
                
            else:
                # PAUT, PT, MT
                # A=1(순번), B=2(제조사), C=3(구분), D=4(Line), E=5(Joint), F=6(용접사), G=7(관경), H=8(합부), I=9(검사길이)
                ws.cell(row=insert_row, column=1).value = insert_row - 2
                ws.cell(row=insert_row, column=2).value = self.company_var.get()
                ws.cell(row=insert_row, column=3).value = self.type_var.get()
                ws.cell(row=insert_row, column=4).value = self.line_var.get()
                ws.cell(row=insert_row, column=5).value = self.joint_var.get()
                ws.cell(row=insert_row, column=6).value = self.welder_var.get()
                ws.cell(row=insert_row, column=7).value = self.size_var.get()
                ws.cell(row=insert_row, column=8).value = self.result_var.get()
                ws.cell(row=insert_row, column=9).value = self.length_var.get()

            # Ensure alignment and borders if we exceed max row (which was 59)
            max_c = 22 if sheet_name == "RT" else 9
            hair_side = Side(style='hair')
            thin_side = Side(style='thin')
            
            for c in range(1, max_c + 1):
                cell = ws.cell(row=insert_row, column=c)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                
                # If we exceeded the previous bottom border (row 59), we must expand it
                if insert_row >= 59:
                    left = thin_side if c == 1 else hair_side
                    right = thin_side if c == max_c else hair_side
                    top = hair_side
                    bottom = thin_side # Move the solid bottom border down!
                    
                    cell.border = Border(left=left, right=right, top=top, bottom=bottom)
                    
                    # Fix previous row's bottom border to hair
                    prev_cell = ws.cell(row=insert_row - 1, column=c)
                    pcb = prev_cell.border
                    prev_cell.border = Border(left=pcb.left, right=pcb.right, top=pcb.top, bottom=hair_side)

            wb.save(self.excel_path)
            
            # Update Treeview
            res = self.rt_result_var.get() if sheet_name == "RT" else self.result_var.get()
            self.tree.insert("", 0, values=(sheet_name, self.line_var.get(), self.joint_var.get(), res))
            
            # Auto-increment Joint No.
            self._auto_increment_joint()
            
            # Clear non-sticky fields
            self._clear_fields()
            
        except PermissionError:
            messagebox.showerror("접근 거부", "엑셀 파일이 열려있습니다. 닫고 다시 시도해주세요!")
        except Exception as e:
            messagebox.showerror("오류", f"저장 실패:\n{e}")

    def _auto_increment_joint(self):
        curr_joint = self.joint_var.get()
        # Find ending number in W01, J-12, etc.
        match = re.search(r'(\d+)$', curr_joint)
        if match:
            num_str = match.group(1)
            next_num = str(int(num_str) + 1).zfill(len(num_str))
            next_joint = curr_joint[:match.start()] + next_num
            self.joint_var.set(next_joint)
        else:
            self.joint_var.set("") # Clear if no number found
            
    def _clear_fields(self):
        if not self.sticky_company.get(): self.company_var.set("")
        if not self.sticky_type.get(): self.type_var.set("")
        if not self.sticky_line.get(): self.line_var.set("")
        if not self.sticky_welder.get(): self.welder_var.set("")
        if not self.sticky_size.get(): self.size_var.set("")
        self.length_var.set("")
        
        # RT
        if not self.sticky_date.get(): self.date_var.set("")
        if not self.sticky_section.get(): self.section_var.set("")
        if not self.sticky_film.get(): self.film_size_var.set("")

def open_data_entry_ui(parent):
    DataEntryUI(parent)
