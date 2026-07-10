import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import datetime
import win32com.client as win32
try:
    from tkcalendar import DateEntry
except ImportError:
    DateEntry = None

class TBMFormTab(ttk.Frame):
    def __init__(self, parent, main_app=None):
        super().__init__(parent)
        self.main_app = main_app
        self.create_widgets()
        
    def create_widgets(self):
        canvas = tk.Canvas(self)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)
        
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        self.build_ui(scroll_frame)

    def build_ui(self, parent):
        lbl_title = tk.Label(parent, text="작업 전 안전점검회의(TBM) 회의록 입력", font=("맑은 고딕", 16, "bold"))
        lbl_title.pack(pady=10)
        
        # 1. 기본 정보
        f_basic = ttk.LabelFrame(parent, text="1. 기본 정보", padding=10)
        f_basic.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(f_basic, text="TBM 일자:").grid(row=0, column=0, sticky=tk.W, pady=2)
        if DateEntry:
            self.ent_date = DateEntry(f_basic, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
            self.ent_date.grid(row=0, column=1, sticky=tk.W, pady=2)
        else:
            self.ent_date = ttk.Entry(f_basic, width=15)
            self.ent_date.insert(0, datetime.datetime.now().strftime("%Y-%m-%d"))
            self.ent_date.grid(row=0, column=1, sticky=tk.W, pady=2)
            
        ttk.Label(f_basic, text="시작 시간:").grid(row=0, column=2, sticky=tk.W, padx=(10,0))
        self.ent_start_time = ttk.Entry(f_basic, width=10)
        self.ent_start_time.insert(0, "08:00")
        self.ent_start_time.grid(row=0, column=3, sticky=tk.W)
        
        ttk.Label(f_basic, text="종료 시간:").grid(row=0, column=4, sticky=tk.W, padx=(10,0))
        self.ent_end_time = ttk.Entry(f_basic, width=10)
        self.ent_end_time.insert(0, "08:15")
        self.ent_end_time.grid(row=0, column=5, sticky=tk.W)
        
        self.var_same_date = tk.StringVar(value="예")
        ttk.Label(f_basic, text="작업날짜와 동일함:").grid(row=0, column=6, sticky=tk.W, padx=(20,0))
        ttk.Radiobutton(f_basic, text="예", variable=self.var_same_date, value="예").grid(row=0, column=7)
        ttk.Radiobutton(f_basic, text="아니오", variable=self.var_same_date, value="아니오").grid(row=0, column=8)
        
        ttk.Label(f_basic, text="작 업 명:").grid(row=1, column=0, sticky=tk.W, pady=2)
        
        f_work_name = ttk.Frame(f_basic)
        f_work_name.grid(row=1, column=1, columnspan=7, sticky=tk.W, pady=2)
        
        self.var_wn_rt = tk.BooleanVar(value=True)
        self.var_wn_ut = tk.BooleanVar()
        self.var_wn_pt = tk.BooleanVar()
        
        def update_work_name():
            names = []
            if self.var_wn_rt.get(): names.append("방사선투과검사")
            if self.var_wn_ut.get(): names.append("초음파탐상검사")
            if self.var_wn_pt.get(): names.append("침투탐상검사")
            self.ent_work_name.delete(0, tk.END)
            self.ent_work_name.insert(0, ", ".join(names))
            
        ttk.Checkbutton(f_work_name, text="RT", variable=self.var_wn_rt, command=update_work_name).pack(side=tk.LEFT, padx=(0,5))
        ttk.Checkbutton(f_work_name, text="UT", variable=self.var_wn_ut, command=update_work_name).pack(side=tk.LEFT, padx=5)
        ttk.Checkbutton(f_work_name, text="PT", variable=self.var_wn_pt, command=update_work_name).pack(side=tk.LEFT, padx=(5,10))
        
        self.ent_work_name = ttk.Entry(f_work_name, width=40)
        self.ent_work_name.pack(side=tk.LEFT)
        update_work_name()
        
        ttk.Label(f_basic, text="작업내용:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.ent_work_content = ttk.Entry(f_basic, width=50)
        self.ent_work_content.grid(row=2, column=1, columnspan=5, sticky=tk.W, pady=2)
        
        ttk.Label(f_basic, text="TBM 장소:").grid(row=3, column=0, sticky=tk.W, pady=2)
        self.ent_location = ttk.Entry(f_basic, width=50)
        self.ent_location.grid(row=3, column=1, columnspan=3, sticky=tk.W, pady=2)
        
        self.var_risk_eval = tk.StringVar(value="예")
        ttk.Label(f_basic, text="위험성평가 실시여부:").grid(row=3, column=4, columnspan=2, sticky=tk.W, padx=(10,0))
        ttk.Radiobutton(f_basic, text="예", variable=self.var_risk_eval, value="예").grid(row=3, column=6)
        ttk.Radiobutton(f_basic, text="아니오", variable=self.var_risk_eval, value="아니오").grid(row=3, column=7)
        
        # 2. 위험 요인
        f_risk = ttk.LabelFrame(parent, text="2. 잠재/중점 위험요인 및 대책", padding=10)
        f_risk.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(f_risk, text="잠재위험요인", font=("", 9, "bold")).grid(row=0, column=0, sticky=tk.W)
        ttk.Label(f_risk, text="대책", font=("", 9, "bold")).grid(row=0, column=1, sticky=tk.W)
        
        self.hazards = []
        default_hazards = [
            "(방사선 피폭) 방사선 투과검사 중 피폭",
            "(추락) 지상 2m 이상 배관 위 검사",
            "(질식) 배관 내부 진입 시 산소 결핍"
        ]
        default_counters = [
            "콜리메이터 사용, 통제구역 설정/감시자 배치",
            "고소작업 시 2인 1조 필수, 안전대 체결",
            "배관내부 인원 진입 금지 (크롤러 대체)"
        ]
        
        for i in range(3):
            ent_h = ttk.Entry(f_risk, width=40)
            ent_h.insert(0, default_hazards[i])
            ent_c = ttk.Entry(f_risk, width=40)
            ent_c.insert(0, default_counters[i])
            
            ent_h.grid(row=i+1, column=0, pady=2, padx=2)
            ent_c.grid(row=i+1, column=1, pady=2, padx=2)
            self.hazards.append((ent_h, ent_c))

        ttk.Label(f_risk, text="중점위험요인 선정:").grid(row=4, column=0, sticky=tk.W, pady=(10,2))
        self.ent_key_hazard = ttk.Entry(f_risk, width=40)
        self.ent_key_hazard.insert(0, "(방사선 피폭) 방사선 투과검사 중 피폭")
        self.ent_key_hazard.grid(row=5, column=0, sticky=tk.W, padx=2)

        ttk.Label(f_risk, text="중점위험 대책:").grid(row=4, column=1, sticky=tk.W, pady=(10,2))
        self.ent_key_counter = ttk.Entry(f_risk, width=40)
        self.ent_key_counter.insert(0, "콜리메이터 사용, 통제구역 설정 및 감시자 배치")
        self.ent_key_counter.grid(row=5, column=1, sticky=tk.W, padx=2)
        
        # 3. 리더 및 점검
        f_leader = ttk.LabelFrame(parent, text="3. 리더 확인 및 안전점검 결과", padding=10)
        f_leader.pack(fill=tk.X, padx=10, pady=5)
        
        f_leader_sub = ttk.Frame(f_leader)
        f_leader_sub.pack(fill=tk.X, pady=2)
        ttk.Label(f_leader_sub, text="[TBM 리더]  소속:").pack(side=tk.LEFT)
        self.ent_leader_dept = ttk.Entry(f_leader_sub, width=15)
        self.ent_leader_dept.pack(side=tk.LEFT, padx=5)
        ttk.Label(f_leader_sub, text="직책:").pack(side=tk.LEFT)
        self.ent_leader_title = ttk.Entry(f_leader_sub, width=10)
        self.ent_leader_title.pack(side=tk.LEFT, padx=5)
        ttk.Label(f_leader_sub, text="성명:").pack(side=tk.LEFT)
        self.ent_leader_name = ttk.Entry(f_leader_sub, width=15)
        self.ent_leader_name.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(f_leader, text="작업 전 일일 안전점검 시행 결과:").pack(anchor=tk.W, pady=(10,2))
        self.ent_daily_check = tk.Text(f_leader, height=3, width=80)
        self.ent_daily_check.pack(anchor=tk.W)
        self.ent_daily_check.insert("1.0", "특이사항 없음.")
        
        ttk.Label(f_leader, text="작업 후 종료 미팅 (중점대책 실효성 등):").pack(anchor=tk.W, pady=(10,2))
        self.ent_end_meeting = tk.Text(f_leader, height=3, width=80)
        self.ent_end_meeting.pack(anchor=tk.W)
        self.ent_end_meeting.insert("1.0", "안전하게 작업 종료함.")
        
        # 4. 참석자 명단
        f_attend = ttk.LabelFrame(parent, text="4. 참석자 명단", padding=10)
        f_attend.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(f_attend, text="참석자 이름 (쉼표로 구분하여 입력, 최대 18명):").pack(anchor=tk.W)
        self.ent_attendees = ttk.Entry(f_attend, width=80)
        self.ent_attendees.pack(anchor=tk.W, pady=5)
        self.ent_attendees.insert(0, "홍길동, 김철수, 이영희")
        
        # 5. 출력 버튼
        f_btn = ttk.Frame(parent)
        f_btn.pack(fill=tk.X, padx=10, pady=20)
        
        btn_export = ttk.Button(f_btn, text="엑셀 양식 출력하기", command=self.export_excel)
        btn_export.pack(ipady=5, ipadx=20)

    def export_excel(self):
        try:
            excel = win32.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            wb = excel.Workbooks.Add()
            ws = wb.ActiveSheet
            ws.Name = "TBM 회의록"
            
            # --- Page Setup ---
            ws.PageSetup.PaperSize = 9 # xlPaperA4
            ws.PageSetup.Orientation = 1 # Portrait
            ws.PageSetup.Zoom = False
            ws.PageSetup.FitToPagesWide = 1
            ws.PageSetup.FitToPagesTall = 1
            ws.PageSetup.LeftMargin = excel.InchesToPoints(0.5)
            ws.PageSetup.RightMargin = excel.InchesToPoints(0.5)
            ws.PageSetup.TopMargin = excel.InchesToPoints(0.5)
            ws.PageSetup.BottomMargin = excel.InchesToPoints(0.5)
            
            # --- Column Widths (total ~ 80) ---
            widths = [2, 10, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 2]
            for i, w in enumerate(widths, 1):
                ws.Columns(i).ColumnWidth = w
                
            # Header
            ws.Range("B1:P1").Merge()
            ws.Range("B1").Value = "첨부1. 작업 전 안전점검회의(TBM) 회의록"
            ws.Range("B1").Font.Size = 11
            
            ws.Range("B3:P4").Merge()
            ws.Range("B3").Value = "작업 전 안전점검회의(TBM) 회의록"
            ws.Range("B3").Font.Size = 18
            ws.Range("B3").Font.Bold = True
            ws.Range("B3").Font.Underline = True
            ws.Range("B3").HorizontalAlignment = -4108 # Center
            ws.Range("B3").VerticalAlignment = -4108
            
            # Base table border range B6:Q28 roughly
            row = 6
            
            # Row 1: TBM 일시
            ws.Range(f"B{row}:C{row}").Merge()
            ws.Cells(row, 2).Value = "TBM 일시"
            
            d_val = self.ent_date.get().replace("-", "")
            y = d_val[:4] if len(d_val)>=4 else "202 "
            m = d_val[4:6] if len(d_val)>=6 else "  "
            d = d_val[6:8] if len(d_val)>=8 else "  "
            
            ws.Range(f"D{row}:J{row}").Merge()
            ws.Cells(row, 4).Value = f"{y} 년  {m} 월  {d} 일   {self.ent_start_time.get()} ~ {self.ent_end_time.get()}"
            
            ws.Range(f"K{row}:Q{row}").Merge()
            same_d = self.var_same_date.get()
            y_box = "☑" if same_d == "예" else "☐"
            n_box = "☑" if same_d == "아니오" else "☐"
            ws.Cells(row, 11).Value = f"작업날짜와 동일함 ({y_box}예, {n_box}아니오)"
            row += 1
            
            # Row 2: 작업명
            ws.Range(f"B{row}:C{row}").Merge()
            ws.Cells(row, 2).Value = "작 업 명"
            ws.Range(f"D{row}:Q{row}").Merge()
            ws.Cells(row, 4).Value = self.ent_work_name.get()
            row += 1
            
            # Row 3: 작업내용
            ws.Range(f"B{row}:C{row}").Merge()
            ws.Cells(row, 2).Value = "작업내용"
            ws.Range(f"D{row}:Q{row}").Merge()
            ws.Cells(row, 4).Value = self.ent_work_content.get()
            row += 1
            
            # Row 4: 장소 & 위험성평가
            ws.Range(f"B{row}:C{row}").Merge()
            ws.Cells(row, 2).Value = "TBM 장소"
            ws.Range(f"D{row}:K{row}").Merge()
            ws.Cells(row, 4).Value = self.ent_location.get()
            ws.Range(f"L{row}:N{row}").Merge()
            ws.Cells(row, 12).Value = "위험성평가 실시여부"
            ws.Range(f"O{row}:Q{row}").Merge()
            r_eval = self.var_risk_eval.get()
            y_box = "☑" if r_eval == "예" else "☐"
            n_box = "☑" if r_eval == "아니오" else "☐"
            ws.Cells(row, 15).Value = f"예 {y_box}  아니오 {n_box}"
            row += 1
            
            # Hazard Headers
            ws.Range(f"B{row}:I{row}").Merge()
            ws.Cells(row, 2).Value = "잠재위험요인"
            ws.Range(f"J{row}:Q{row}").Merge()
            ws.Cells(row, 10).Value = "대책 (※ 제거 → 대체 → 통제 순서 고려)"
            row += 1
            
            # Hazards
            for ent_h, ent_c in self.hazards:
                ws.Range(f"B{row}:I{row}").Merge()
                ws.Cells(row, 2).Value = f"☐ {ent_h.get()}"
                ws.Range(f"J{row}:Q{row}").Merge()
                ws.Cells(row, 10).Value = f"☐ {ent_c.get()}"
                row += 1
                
            # Key Hazard
            ws.Range(f"B{row}:C{row+1}").Merge()
            ws.Cells(row, 2).Value = "중점위험\n요인"
            ws.Cells(row, 4).Value = "선정"
            ws.Range(f"E{row}:Q{row}").Merge()
            ws.Cells(row, 5).Value = f"※ {self.ent_key_hazard.get()}"
            row += 1
            ws.Cells(row, 4).Value = "대책"
            ws.Range(f"E{row}:Q{row}").Merge()
            ws.Cells(row, 5).Value = f"{self.ent_key_counter.get()}"
            row += 1
            
            # Leader
            ws.Range(f"B{row}:D{row}").Merge()
            ws.Cells(row, 2).Value = "TBM 리더 확인"
            ws.Range(f"E{row}:G{row}").Merge()
            ws.Cells(row, 5).Value = f"소속 : {self.ent_leader_dept.get()}"
            ws.Range(f"H{row}:K{row}").Merge()
            ws.Cells(row, 8).Value = f"직책: {self.ent_leader_title.get()}"
            ws.Range(f"L{row}:Q{row}").Merge()
            ws.Cells(row, 12).Value = f"성명: {self.ent_leader_name.get()}          (서명)"
            row += 1
            
            # Pre-work measures (copying potential hazards for simplicity)
            ws.Range(f"B{row}:Q{row}").Merge()
            ws.Cells(row, 2).Value = "▣ 작업 전 안전조치 확인 ※ 위 잠재위험요인(중점위험 포함) 안전조치 여부 재확인"
            ws.Cells(row, 2).Font.Bold = True
            row += 1
            
            ws.Range(f"B{row}:I{row}").Merge()
            ws.Cells(row, 2).Value = "잠재위험요소(중점위험 포함)"
            ws.Range(f"J{row}:M{row}").Merge()
            ws.Cells(row, 10).Value = "조치여부"
            ws.Range(f"N{row}:Q{row}").Merge()
            ws.Cells(row, 14).Value = "'아니오' 인 경우 조치 내용"
            row += 1
            
            for ent_h, ent_c in self.hazards:
                ws.Range(f"B{row}:I{row}").Merge()
                h_val = ent_h.get()
                ws.Cells(row, 2).Value = f"☐ {h_val}" if h_val else "☐"
                ws.Range(f"J{row}:M{row}").Merge()
                ws.Cells(row, 10).Value = "예 ☑, 아니오 ☐" if h_val else "예 ☐, 아니오 ☐"
                ws.Range(f"N{row}:Q{row}").Merge()
                row += 1
                
            # Daily Result
            ws.Range(f"B{row}:Q{row}").Merge()
            ws.Cells(row, 2).Value = "▣ 작업 전 일일 안전점검 시행 결과"
            ws.Cells(row, 2).Font.Bold = True
            row += 1
            ws.Range(f"B{row}:Q{row+2}").Merge()
            ws.Cells(row, 2).Value = self.ent_daily_check.get("1.0", tk.END).strip()
            ws.Cells(row, 2).VerticalAlignment = -4160 # Top
            row += 3
            
            # Post meeting
            ws.Range(f"B{row}:Q{row}").Merge()
            ws.Cells(row, 2).Value = "▣ 작업 후 종료 미팅(중점대책의 실효성)"
            ws.Cells(row, 2).Font.Bold = True
            row += 1
            ws.Range(f"B{row}:Q{row+2}").Merge()
            ws.Cells(row, 2).Value = self.ent_end_meeting.get("1.0", tk.END).strip()
            ws.Cells(row, 2).VerticalAlignment = -4160
            row += 3
            
            # Attendees Header
            ws.Range(f"B{row}:Q{row}").Merge()
            ws.Cells(row, 2).Value = "▣ 참석자 확인 ※ TBM에 참여하지 않은 작업자를 확인하여 미팅 참석 유도"
            ws.Cells(row, 2).Font.Bold = True
            row += 1
            
            # Attendees Table
            a_row_start = row
            for c in range(2, 17, 3):
                ws.Range(ws.Cells(row, c), ws.Cells(row, c+1)).Merge()
                ws.Cells(row, c).Value = "이름"
                ws.Cells(row, c+2).Value = "서명"
            row += 1
            
            att_list = [x.strip() for x in self.ent_attendees.get().split(',') if x.strip()]
            att_idx = 0
            for r in range(3):
                for c in range(2, 17, 3):
                    ws.Range(ws.Cells(row, c), ws.Cells(row, c+1)).Merge()
                    if att_idx < len(att_list):
                        ws.Cells(row, c).Value = att_list[att_idx]
                        att_idx += 1
                row += 1
                
            # Formatting styles
            for r in range(6, row):
                for c in range(2, 18):
                    ws.Cells(r, c).Borders.LineStyle = 1
                    
            # Set Row Heights
            for r in range(6, row):
                ws.Rows(r).RowHeight = 25
            ws.Rows(15).RowHeight = 40 # Daily check
            ws.Rows(19).RowHeight = 40 # End meeting
            
            # Center alignments for most
            ws.Range(f"B6:Q{row-1}").VerticalAlignment = -4108
            
            # Footer
            row += 1
            ws.Range(f"B{row}:E{row}").Merge()
            ws.Cells(row, 2).Value = "양식 C 312-1"
            ws.Range(f"F{row}:M{row}").Merge()
            ws.Cells(row, 6).Value = "서울검사(주)"
            ws.Cells(row, 6).HorizontalAlignment = -4108
            ws.Range(f"N{row}:Q{row}").Merge()
            ws.Cells(row, 14).Value = "A4(210X297)"
            ws.Cells(row, 14).HorizontalAlignment = -4152 # Right
            
            default_name = f"TBM회의록_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx"
            filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=default_name, filetypes=[("Excel File", "*.xlsx")], title="TBM 엑셀 저장")
            
            if filepath:
                filepath = filepath.replace("/", "\\")
                wb.SaveAs(filepath)
                wb.Close()
                excel.Quit()
                messagebox.showinfo("저장 완료", f"TBM 회의록이 성공적으로 생성되었습니다.\n{filepath}")
                os.startfile(filepath)
            else:
                wb.Close(False)
                excel.Quit()
                
        except Exception as e:
            messagebox.showerror("오류", f"엑셀 생성 중 오류가 발생했습니다.\n{e}")
            try: excel.Quit()
            except: pass
