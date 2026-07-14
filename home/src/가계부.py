import tkinter as tk
from tkinter import ttk, messagebox
import datetime
from 가계부_DB import HouseholdDB
import matplotlib
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class HouseholdApp:
    def __init__(self, root):
        self.root = root
        self.root.title("💰 나만의 스마트 가계부")
        self.root.geometry("1100x700")
        
        # 스타일 및 폰트 설정
        style = ttk.Style()
        try:
            style.theme_use('clam')
        except:
            pass
        
        # 기본 폰트 설정 (맑은 고딕)
        font_main = ('Malgun Gothic', 10)
        font_bold = ('Malgun Gothic', 10, 'bold')
        
        style.configure(".", font=font_main)
        style.configure("Treeview.Heading", font=font_bold)
        style.configure("Treeview", font=font_main, rowheight=30)
        
        # 데이터베이스 매니저 생성
        self.db = HouseholdDB()
        
        # 현재 선택된 연/월
        now = datetime.datetime.now()
        self.current_year = now.year
        self.current_month = now.month
        
        # 카테고리 목록
        self.income_categories = ["월급", "용돈", "상여금", "이자/배당", "기타수입"]
        self.expense_categories = ["식비", "교통/차량", "문화/생활", "주거/통신", "건강/의료", "쇼핑", "기타지출"]
        
        self._build_ui()
        self.refresh_data()
        
    def _build_ui(self):
        # 상단 타이틀 및 월 선택 프레임
        top_frame = ttk.Frame(self.root, padding=10)
        top_frame.pack(fill='x')
        
        btn_prev = ttk.Button(top_frame, text="◀ 이전달", command=self.prev_month, width=8)
        btn_prev.pack(side='left', padx=5)
        
        self.lbl_month = ttk.Label(top_frame, text=f"{self.current_year}년 {self.current_month}월", font=('Malgun Gothic', 18, 'bold'))
        self.lbl_month.pack(side='left', padx=20)
        
        btn_next = ttk.Button(top_frame, text="다음달 ▶", command=self.next_month, width=8)
        btn_next.pack(side='left', padx=5)
        
        btn_export = ttk.Button(top_frame, text="📊 엑셀 저장", command=self.export_excel)
        btn_export.pack(side='right', padx=5)
        
        # 메인 콘텐츠 영역 (좌측: 내역/입력, 우측: 통계)
        main_paned = ttk.PanedWindow(self.root, orient='horizontal')
        main_paned.pack(expand=True, fill='both', padx=10, pady=10)
        
        left_frame = ttk.Frame(main_paned)
        right_frame = ttk.Frame(main_paned)
        
        main_paned.add(left_frame, weight=3)
        main_paned.add(right_frame, weight=2)
        
        # ================= 우측: 대시보드 및 통계 =================
        dash_frame = ttk.LabelFrame(right_frame, text=" 📊 월간 요약 ", padding=10)
        dash_frame.pack(fill='x', pady=(0, 10))
        
        self.lbl_income = ttk.Label(dash_frame, text="수입: 0 원", foreground="blue", font=('Malgun Gothic', 14, 'bold'))
        self.lbl_income.pack(anchor='w', pady=2)
        
        self.lbl_expense = ttk.Label(dash_frame, text="지출: 0 원", foreground="red", font=('Malgun Gothic', 14, 'bold'))
        self.lbl_expense.pack(anchor='w', pady=2)
        
        self.lbl_balance = ttk.Label(dash_frame, text="잔액: 0 원", font=('Malgun Gothic', 15, 'bold'))
        self.lbl_balance.pack(anchor='w', pady=10)
        
        chart_frame = ttk.LabelFrame(right_frame, text=" 🍕 지출 카테고리 분석 ", padding=10)
        chart_frame.pack(expand=True, fill='both')
        
        self.figure = plt.Figure(figsize=(5, 4), dpi=100)
        # 한글 폰트 설정
        matplotlib.rcParams['font.family'] = 'Malgun Gothic'
        matplotlib.rcParams['axes.unicode_minus'] = False
        
        self.ax = self.figure.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.figure, master=chart_frame)
        self.canvas.get_tk_widget().pack(expand=True, fill='both')
        
        # ================= 좌측 상단: 입력 폼 =================
        input_frame = ttk.LabelFrame(left_frame, text=" 📝 내역 입력 ", padding=10)
        input_frame.pack(fill='x', pady=(0, 10))
        
        # 날짜
        ttk.Label(input_frame, text="날짜:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        try:
            from tkcalendar import DateEntry
            self.ent_date = DateEntry(input_frame, width=12, background='gray', 
                                      foreground='white', borderwidth=2, 
                                      date_pattern='yyyy-mm-dd')
        except ImportError:
            self.ent_date = ttk.Entry(input_frame, width=12)
            self.ent_date.insert(0, datetime.datetime.now().strftime("%Y-%m-%d"))
        self.ent_date.grid(row=0, column=1, padx=5, pady=5, sticky='w')
        
        # 유형 (수입/지출)
        ttk.Label(input_frame, text="유형:").grid(row=0, column=2, padx=5, pady=5, sticky='w')
        self.var_type = tk.StringVar(value="지출")
        cb_type = ttk.Combobox(input_frame, textvariable=self.var_type, values=["수입", "지출"], width=8, state='readonly')
        cb_type.grid(row=0, column=3, padx=5, pady=5, sticky='w')
        cb_type.bind('<<ComboboxSelected>>', self.update_category_list)
        
        # 카테고리
        ttk.Label(input_frame, text="분류:").grid(row=0, column=4, padx=5, pady=5, sticky='w')
        self.cb_category = ttk.Combobox(input_frame, values=self.expense_categories, width=12)
        self.cb_category.current(0)
        self.cb_category.grid(row=0, column=5, padx=5, pady=5, sticky='w')
        
        # 금액
        ttk.Label(input_frame, text="금액:").grid(row=1, column=0, padx=5, pady=5, sticky='w')
        self.ent_amount = ttk.Entry(input_frame, width=12)
        self.ent_amount.grid(row=1, column=1, padx=5, pady=5, sticky='w')
        self.ent_amount.bind('<KeyRelease>', self.format_amount)
        
        # 메모
        ttk.Label(input_frame, text="메모:").grid(row=1, column=2, padx=5, pady=5, sticky='w')
        self.ent_note = ttk.Entry(input_frame, width=30)
        self.ent_note.grid(row=1, column=3, columnspan=3, padx=5, pady=5, sticky='we')
        
        # 버튼
        btn_add = ttk.Button(input_frame, text="저장하기", command=self.add_record)
        btn_add.grid(row=1, column=6, padx=10, pady=5)
        
        # ================= 좌측 하단: 내역 목록 =================
        list_frame = ttk.LabelFrame(left_frame, text=" 📖 상세 내역 ", padding=5)
        list_frame.pack(expand=True, fill='both')
        
        cols = ('ID', '날짜', '유형', '카테고리', '금액', '메모')
        self.tv = ttk.Treeview(list_frame, columns=cols, show='headings')
        
        self.tv.heading('ID', text='ID')
        self.tv.column('ID', width=0, stretch=False) # 숨김
        self.tv.heading('날짜', text='날짜')
        self.tv.column('날짜', width=100, anchor='center')
        self.tv.heading('유형', text='유형')
        self.tv.column('유형', width=60, anchor='center')
        self.tv.heading('카테고리', text='카테고리')
        self.tv.column('카테고리', width=100, anchor='center')
        self.tv.heading('금액', text='금액(원)')
        self.tv.column('금액', width=120, anchor='e')
        self.tv.heading('메모', text='메모')
        self.tv.column('메모', width=200, anchor='w')
        
        # 스크롤바
        scroll = ttk.Scrollbar(list_frame, orient='vertical', command=self.tv.yview)
        self.tv.configure(yscroll=scroll.set)
        scroll.pack(side='right', fill='y')
        self.tv.pack(expand=True, fill='both')
        
        # 삭제 버튼 프레임
        bot_frame = ttk.Frame(list_frame)
        bot_frame.pack(fill='x', pady=(5, 0))
        btn_del = ttk.Button(bot_frame, text="🗑️ 선택 삭제", command=self.delete_record)
        btn_del.pack(side='right')
        
        # 색상 태그 설정
        self.tv.tag_configure('수입', foreground='blue')
        self.tv.tag_configure('지출', foreground='red')

    def update_category_list(self, event=None):
        t_type = self.var_type.get()
        default_cats = self.income_categories if t_type == "수입" else self.expense_categories
        
        # DB에서 사용자가 입력했던 커스텀 카테고리 불러와서 병합
        custom_cats = self.db.get_unique_categories(t_type)
        merged_cats = list(default_cats)
        for cat in custom_cats:
            if cat not in merged_cats:
                merged_cats.append(cat)
                
        self.cb_category.config(values=merged_cats)
        self.cb_category.current(0)

    def format_amount(self, event):
        """금액 입력 시 자동으로 콤마(,) 추가"""
        if event.keysym in ['Left', 'Right', 'Up', 'Down', 'BackSpace', 'Delete']:
            return
            
        content = self.ent_amount.get().replace(',', '')
        if not content: return
        
        if content.isdigit():
            formatted = f"{int(content):,}"
            self.ent_amount.delete(0, 'end')
            self.ent_amount.insert(0, formatted)
        else:
            # 숫자가 아닌 문자 제거
            clean_content = ''.join(filter(str.isdigit, content))
            if clean_content:
                formatted = f"{int(clean_content):,}"
                self.ent_amount.delete(0, 'end')
                self.ent_amount.insert(0, formatted)
            else:
                self.ent_amount.delete(0, 'end')

    def prev_month(self):
        if self.current_month == 1:
            self.current_month = 12
            self.current_year -= 1
        else:
            self.current_month -= 1
        self.lbl_month.config(text=f"{self.current_year}년 {self.current_month}월")
        self.refresh_data()

    def next_month(self):
        if self.current_month == 12:
            self.current_month = 1
            self.current_year += 1
        else:
            self.current_month += 1
        self.lbl_month.config(text=f"{self.current_year}년 {self.current_month}월")
        self.refresh_data()

    def refresh_data(self):
        # 1. 트리뷰 초기화
        for item in self.tv.get_children():
            self.tv.delete(item)
            
        # 2. 내역 불러오기
        records = self.db.get_transactions_by_month(self.current_year, self.current_month)
        for r in records:
            # r = (id, date, type, category, amount, note)
            fmt_amount = f"{r[4]:,}원"
            self.tv.insert('', 'end', values=(r[0], r[1], r[2], r[3], fmt_amount, r[5]), tags=(r[2],))
            
        # 3. 요약 데이터 업데이트
        summary = self.db.get_monthly_summary(self.current_year, self.current_month)
        self.lbl_income.config(text=f"수입: +{summary['income']:,} 원")
        self.lbl_expense.config(text=f"지출: -{summary['expense']:,} 원")
        self.lbl_balance.config(text=f"잔액: {summary['balance']:,} 원")
        
        # 4. 차트 업데이트
        self.ax.clear()
        exp_dict = summary['expense_by_category']
        if exp_dict:
            total_exp = sum(exp_dict.values())
            merged_dict = {}
            기타_합계 = 0
            
            # 3% 미만은 '기타'로 묶기
            for cat, amt in exp_dict.items():
                if (amt / total_exp) < 0.03:
                    기타_합계 += amt
                else:
                    merged_dict[cat] = amt
                    
            if 기타_합계 > 0:
                if '기타' in merged_dict:
                    merged_dict['기타'] += 기타_합계
                else:
                    merged_dict['기타'] = 기타_합계
                    
            labels = list(merged_dict.keys())
            sizes = list(merged_dict.values())
            self.ax.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=90, textprops={'fontsize': 9})
            self.ax.axis('equal')
        else:
            self.ax.text(0.5, 0.5, '지출 내역이 없습니다.', ha='center', va='center', fontsize=12, color='gray')
            self.ax.axis('off')
            
        self.canvas.draw()

    def add_record(self):
        date = self.ent_date.get().strip()
        t_type = self.var_type.get()
        category = self.cb_category.get()
        amount_str = self.ent_amount.get().strip().replace(',', '')
        note = self.ent_note.get().strip()
        
        if not date or not amount_str:
            messagebox.showwarning("입력 오류", "날짜와 금액을 모두 입력해주세요.")
            return
            
        try:
            amount = int(amount_str)
            datetime.datetime.strptime(date, "%Y-%m-%d")
        except ValueError:
            messagebox.showwarning("입력 오류", "금액은 숫자만, 날짜는 YYYY-MM-DD 형식으로 입력하세요.")
            return
            
        self.db.add_transaction(date, t_type, category, amount, note)
        
        # 입력창 초기화
        self.ent_amount.delete(0, 'end')
        self.ent_note.delete(0, 'end')
        
        self.refresh_data()
        
    def delete_record(self):
        selection = self.tv.selection()
        if not selection:
            messagebox.showwarning("선택 안 됨", "삭제할 내역을 선택해주세요.")
            return
            
        if messagebox.askyesno("삭제 확인", "선택한 내역을 정말 삭제하시겠습니까?"):
            for item in selection:
                t_id = self.tv.item(item, 'values')[0]
                self.db.delete_transaction(t_id)
            self.refresh_data()
            
    def export_excel(self):
        filepath = self.db.export_to_excel(f"가계부_내역_{self.current_year}년{self.current_month}월.xlsx")
        if filepath == "PERMISSION_ERROR":
            messagebox.showerror("저장 실패", "이미 해당 엑셀 파일이 열려있습니다.\n파일을 닫은 후 다시 시도해주세요.")
        elif filepath:
            messagebox.showinfo("저장 완료", f"엑셀 파일이 저장되었습니다.\n{filepath}")
        else:
            messagebox.showinfo("알림", "저장할 내역이 없습니다.")

if __name__ == "__main__":
    root = tk.Tk()
    app = HouseholdApp(root)
    root.mainloop()
