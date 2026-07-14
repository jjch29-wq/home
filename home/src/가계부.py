import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
import datetime
import threading
import base64
import json
from 가계부_DB import HouseholdDB
import matplotlib
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class RecurringManagerWindow(tk.Toplevel):
    """고정 지출/수입 관리를 위한 독립된 팝업창 클래스"""
    def __init__(self, parent, db, refresh_callback):
        super().__init__(parent)
        self.db = db
        self.refresh_callback = refresh_callback
        
        self.title("고정 지출/수입 관리")
        self.geometry("600x400")
        self.transient(parent)
        self.grab_set() # 팝업이 떠 있는 동안 부모 창 클릭 방지
        
        self._build_ui()
        self.load_list()
        
    def _build_ui(self):
        # --- 입력 폼 ---
        frm_input = ttk.Frame(self, padding=10)
        frm_input.pack(fill='x')
        
        ttk.Label(frm_input, text="유형:").grid(row=0, column=0, padx=5, pady=5)
        self.var_t = tk.StringVar(value="지출")
        ttk.Combobox(frm_input, textvariable=self.var_t, values=["수입", "지출"], width=5, state='readonly').grid(row=0, column=1)
        
        ttk.Label(frm_input, text="분류:").grid(row=0, column=2, padx=5, pady=5)
        self.ent_c = ttk.Entry(frm_input, width=10)
        self.ent_c.grid(row=0, column=3)
        
        ttk.Label(frm_input, text="금액:").grid(row=0, column=4, padx=5, pady=5)
        self.ent_a = ttk.Entry(frm_input, width=10)
        self.ent_a.grid(row=0, column=5)
        
        ttk.Label(frm_input, text="메모:").grid(row=1, column=0, padx=5, pady=5)
        self.ent_n = ttk.Entry(frm_input, width=30)
        self.ent_n.grid(row=1, column=1, columnspan=4, sticky='we')
        
        ttk.Button(frm_input, text="등록", command=self.save_recurring).grid(row=1, column=5, padx=5)
        
        # --- 목록 표 ---
        frm_list = ttk.Frame(self, padding=10)
        frm_list.pack(expand=True, fill='both')
        
        self.tv_r = ttk.Treeview(frm_list, columns=('ID', '유형', '카테고리', '금액', '메모'), show='headings')
        self.tv_r.heading('ID', text='ID'); self.tv_r.column('ID', width=0, stretch=False)
        self.tv_r.heading('유형', text='유형'); self.tv_r.column('유형', width=50, anchor='center')
        self.tv_r.heading('카테고리', text='카테고리'); self.tv_r.column('카테고리', width=80, anchor='center')
        self.tv_r.heading('금액', text='금액'); self.tv_r.column('금액', width=100, anchor='e')
        self.tv_r.heading('메모', text='메모'); self.tv_r.column('메모', width=150, anchor='w')
        self.tv_r.pack(expand=True, fill='both', side='left')
        
        ttk.Button(frm_list, text="선택 삭제", command=self.del_recurring).pack(side='bottom', pady=5)

    def load_list(self):
        for i in self.tv_r.get_children(): 
            self.tv_r.delete(i)
        for r in self.db.get_recurring_templates():
            self.tv_r.insert('', 'end', values=(r[0], r[1], r[2], f"{r[3]:,}원", r[4]))
            
    def save_recurring(self):
        try:
            amt = int(self.ent_a.get().replace(',', ''))
            self.db.add_recurring_template(self.var_t.get(), self.ent_c.get(), amt, self.ent_n.get())
            self.ent_c.delete(0, 'end')
            self.ent_a.delete(0, 'end')
            self.ent_n.delete(0, 'end')
            self.load_list()
        except Exception:
            messagebox.showerror("오류", "금액을 정확히 입력하세요.", parent=self)
            
    def del_recurring(self):
        sel = self.tv_r.selection()
        if sel:
            self.db.delete_recurring_template(self.tv_r.item(sel[0], 'values')[0])
            self.load_list()


class HouseholdApp:
    def __init__(self, root):
        self.root = root
        self.root.title("💰 나만의 스마트 가계부")
        self.root.geometry("1200x750")
        
        self._init_styles()
        
        self.db = HouseholdDB()
        
        now = datetime.datetime.now()
        self.current_year = now.year
        self.current_month = now.month
        
        self.income_categories = ["월급", "용돈", "상여금", "이자/배당", "기타수입"]
        self.expense_categories = ["식비", "교통/차량", "문화/생활", "주거/통신", "건강/의료", "쇼핑", "기타지출"]
        
        self._build_ui()
        self.refresh_data()
        
    def _init_styles(self):
        """테마 및 폰트 초기화"""
        style = ttk.Style()
        try: style.theme_use('clam')
        except: pass
        
        font_main = ('Malgun Gothic', 10)
        font_bold = ('Malgun Gothic', 10, 'bold')
        
        style.configure(".", font=font_main)
        style.configure("Treeview.Heading", font=font_bold)
        style.configure("Treeview", font=font_main, rowheight=30)
        style.configure("TProgressbar", thickness=20)
        
    def _build_ui(self):
        """전체 UI 레이아웃 조립"""
        top_frame = ttk.Frame(self.root, padding=10)
        top_frame.pack(fill='x')
        self._build_top_menu(top_frame)
        
        main_paned = ttk.PanedWindow(self.root, orient='horizontal')
        main_paned.pack(expand=True, fill='both', padx=10, pady=10)
        
        left_frame = ttk.Frame(main_paned)
        right_frame = ttk.Frame(main_paned)
        
        main_paned.add(left_frame, weight=3)
        main_paned.add(right_frame, weight=2)
        
        # 좌측 영역 (입력 폼 + 리스트)
        input_frame = ttk.LabelFrame(left_frame, text=" 📝 내역 입력 ", padding=10)
        input_frame.pack(fill='x', pady=(0, 10))
        self._build_input_form(input_frame)
        
        list_frame = ttk.LabelFrame(left_frame, text=" 📖 상세 내역 ", padding=5)
        list_frame.pack(expand=True, fill='both')
        self._build_list_view(list_frame)
        
        # 우측 영역 (탭 구조 통계)
        self.notebook = ttk.Notebook(right_frame)
        self.notebook.pack(expand=True, fill='both')
        
        tab_daily = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab_daily, text=" 📅 일간 요약 ")
        self._build_daily_tab(tab_daily)

        tab_month = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab_month, text=" 📊 월간 요약 ")
        self._build_dashboard_tab(tab_month)
        
        tab_year = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab_year, text=" 📈 연간 트렌드 ")
        self._build_trends_tab(tab_year)
        
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)

    # ----------------------------------------------------
    # UI 빌더 파트 (세부 화면 구성)
    # ----------------------------------------------------
    def _build_top_menu(self, parent):
        ttk.Button(parent, text="◀ 이전달", command=self.prev_month, width=8).pack(side='left', padx=5)
        self.lbl_month = ttk.Label(parent, text=f"{self.current_year}년 {self.current_month}월", font=('Malgun Gothic', 18, 'bold'))
        self.lbl_month.pack(side='left', padx=20)
        ttk.Button(parent, text="다음달 ▶", command=self.next_month, width=8).pack(side='left', padx=5)
        
        ttk.Button(parent, text="📊 엑셀 저장", command=self.export_excel).pack(side='right', padx=5)
        ttk.Button(parent, text="♻️ 고정 지출 관리", command=self.open_recurring_manager).pack(side='right', padx=5)
        ttk.Button(parent, text="🎯 예산 설정", command=self.set_budget).pack(side='right', padx=5)

    def _build_dashboard_tab(self, parent):
        # 예산 프레임
        budget_frame = ttk.LabelFrame(parent, text=" 이번 달 지출 예산 현황 ", padding=10)
        budget_frame.pack(fill='x', pady=(0, 10))
        self.lbl_budget_status = ttk.Label(budget_frame, text="목표 예산: 설정 안됨", font=('Malgun Gothic', 10, 'bold'))
        self.lbl_budget_status.pack(anchor='w', pady=2)
        self.progress_budget = ttk.Progressbar(budget_frame, orient='horizontal', mode='determinate', style="TProgressbar")
        self.progress_budget.pack(fill='x', pady=5)
        
        # 요약 프레임
        summary_frame = ttk.Frame(parent)
        summary_frame.pack(fill='x', pady=10)
        self.lbl_income = ttk.Label(summary_frame, text="수입: 0 원", foreground="blue", font=('Malgun Gothic', 14, 'bold'))
        self.lbl_income.pack(anchor='w', pady=2)
        self.lbl_expense = ttk.Label(summary_frame, text="지출: 0 원", foreground="red", font=('Malgun Gothic', 14, 'bold'))
        self.lbl_expense.pack(anchor='w', pady=2)
        self.lbl_balance = ttk.Label(summary_frame, text="잔액: 0 원", font=('Malgun Gothic', 15, 'bold'))
        self.lbl_balance.pack(anchor='w', pady=10)
        
        # 파이 차트 설정
        self.fig_pie = plt.Figure(figsize=(4, 3), dpi=100)
        matplotlib.rcParams['font.family'] = 'Malgun Gothic'
        matplotlib.rcParams['axes.unicode_minus'] = False
        self.ax_pie = self.fig_pie.add_subplot(111)
        self.canvas_pie = FigureCanvasTkAgg(self.fig_pie, master=parent)
        self.canvas_pie.get_tk_widget().pack(expand=True, fill='both')

    def _build_trends_tab(self, parent):
        self.fig_line = plt.Figure(figsize=(5, 4), dpi=100)
        self.ax_line = self.fig_line.add_subplot(111)
        self.canvas_line = FigureCanvasTkAgg(self.fig_line, master=parent)
        self.canvas_line.get_tk_widget().pack(expand=True, fill='both')

    def _build_daily_tab(self, parent):
        self.fig_daily = plt.Figure(figsize=(5, 4), dpi=100)
        self.ax_daily = self.fig_daily.add_subplot(111)
        self.canvas_daily = FigureCanvasTkAgg(self.fig_daily, master=parent)
        self.canvas_daily.get_tk_widget().pack(expand=True, fill='both')
        self.canvas_daily.mpl_connect('button_press_event', self.on_daily_chart_click)

    def _build_input_form(self, parent):
        ttk.Label(parent, text="날짜:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        try:
            from tkcalendar import DateEntry
            self.ent_date = DateEntry(parent, width=12, background='gray', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
        except ImportError:
            self.ent_date = ttk.Entry(parent, width=12)
            self.ent_date.insert(0, datetime.datetime.now().strftime("%Y-%m-%d"))
        self.ent_date.grid(row=0, column=1, padx=5, pady=5, sticky='w')
        
        ttk.Label(parent, text="유형:").grid(row=0, column=2, padx=5, pady=5, sticky='w')
        self.var_type = tk.StringVar(value="지출")
        cb_type = ttk.Combobox(parent, textvariable=self.var_type, values=["수입", "지출"], width=8, state='readonly')
        cb_type.grid(row=0, column=3, padx=5, pady=5, sticky='w')
        cb_type.bind('<<ComboboxSelected>>', self.update_category_list)
        
        ttk.Label(parent, text="분류:").grid(row=0, column=4, padx=5, pady=5, sticky='w')
        self.cb_category = ttk.Combobox(parent, values=self.expense_categories, width=12)
        self.cb_category.current(0)
        self.cb_category.grid(row=0, column=5, padx=5, pady=5, sticky='w')
        
        ttk.Label(parent, text="금액:").grid(row=1, column=0, padx=5, pady=5, sticky='w')
        self.ent_amount = ttk.Entry(parent, width=12)
        self.ent_amount.grid(row=1, column=1, padx=5, pady=5, sticky='w')
        self.ent_amount.bind('<KeyRelease>', self.format_amount)
        
        ttk.Label(parent, text="메모:").grid(row=1, column=2, padx=5, pady=5, sticky='w')
        self.ent_note = ttk.Entry(parent, width=30)
        self.ent_note.grid(row=1, column=3, columnspan=3, padx=5, pady=5, sticky='we')
        
        btn_frame = ttk.Frame(parent)
        btn_frame.grid(row=1, column=6, padx=10, pady=5, sticky='we')
        ttk.Button(btn_frame, text="저장하기", command=self.add_record).pack(side='left', padx=2)
        ttk.Button(btn_frame, text="📸 영수증 인식", command=self.open_receipt_ocr).pack(side='left', padx=2)

    def _build_list_view(self, parent):
        search_frame = ttk.Frame(parent)
        search_frame.pack(fill='x', pady=(0, 5))
        
        ttk.Label(search_frame, text="통합 검색:").pack(side='left', padx=5)
        self.ent_search = ttk.Entry(search_frame, width=20)
        self.ent_search.pack(side='left', padx=5)
        self.ent_search.bind('<Return>', lambda e: self.search_data())
        
        ttk.Button(search_frame, text="🔍 검색", command=self.search_data).pack(side='left', padx=2)
        ttk.Button(search_frame, text="초기화", command=self.refresh_data).pack(side='left', padx=2)
        ttk.Button(search_frame, text="➕ 이번 달 고정내역 일괄 등록", command=self.load_recurring_to_current_month).pack(side='right', padx=5)
        
        cols = ('ID', '날짜', '유형', '카테고리', '금액', '메모')
        self.tv = ttk.Treeview(parent, columns=cols, show='headings')
        
        self.tv.heading('ID', text='ID'); self.tv.column('ID', width=0, stretch=False) 
        self.tv.heading('날짜', text='날짜'); self.tv.column('날짜', width=100, anchor='center')
        self.tv.heading('유형', text='유형'); self.tv.column('유형', width=60, anchor='center')
        self.tv.heading('카테고리', text='카테고리'); self.tv.column('카테고리', width=100, anchor='center')
        self.tv.heading('금액', text='금액(원)'); self.tv.column('금액', width=120, anchor='center')
        self.tv.heading('메모', text='메모'); self.tv.column('메모', width=200, anchor='w')
        
        for col in cols:
            self.tv.heading(col, command=lambda c=col: self._treeview_sort_column(c, False))
        
        scroll = ttk.Scrollbar(parent, orient='vertical', command=self.tv.yview)
        self.tv.configure(yscroll=scroll.set)
        scroll.pack(side='right', fill='y')
        self.tv.pack(expand=True, fill='both')
        
        ttk.Button(parent, text="🗑️ 선택 삭제", command=self.delete_record).pack(side='right', pady=5)
        
        self.tv.tag_configure('수입', foreground='blue')
        self.tv.tag_configure('지출', foreground='red')
        self.tv.tag_configure('summary', background='#eaeaea', font=('Malgun Gothic', 9, 'bold'))
        self.tv.bind('<<TreeviewSelect>>', self.on_treeview_select)

    def _treeview_sort_column(self, col, reverse):
        children = self.tv.get_children('')
        items = []
        summary_items = []
        for k in children:
            tags = self.tv.item(k, "tags")
            if tags and 'summary' in tags:
                summary_items.append(k)
            else:
                items.append((self.tv.set(k, col), k))
        
        try:
            if col == '금액':
                items.sort(key=lambda t: int(t[0].replace(',', '').replace('원', '')), reverse=reverse)
            elif col == 'ID':
                items.sort(key=lambda t: int(t[0]), reverse=reverse)
            else:
                items.sort(reverse=reverse)
        except Exception:
            items.sort(reverse=reverse)

        for index, (val, k) in enumerate(items):
            self.tv.move(k, '', index)
            
        for k in summary_items:
            self.tv.move(k, '', 'end')

        self.tv.heading(col, command=lambda: self._treeview_sort_column(col, not reverse))

    # ----------------------------------------------------
    # 데이터 처리 및 이벤트 핸들러
    # ----------------------------------------------------
    def on_treeview_select(self, event):
        selected = self.tv.selection()
        if not selected:
            return
            
        item_id = selected[0]
        values = self.tv.item(item_id, 'values')
        
        # 요약(총합계) 줄이거나 빈 값이면 무시
        if not values or not str(values[0]).isdigit():
            return
            
        date_val, t_type, cat, amount_str, note = values[1], values[2], values[3], values[4], values[5]
        
        # 날짜
        if hasattr(self.ent_date, 'set_date'):
            try:
                dt = datetime.datetime.strptime(date_val, "%Y-%m-%d").date()
                self.ent_date.set_date(dt)
            except:
                pass
        else:
            self.ent_date.delete(0, 'end')
            self.ent_date.insert(0, date_val)
            
        # 유형
        self.var_type.set(t_type)
        self.update_category_list()
        
        # 카테고리
        self.cb_category.set(cat)
            
        # 금액
        clean_amt = amount_str.replace(',', '').replace('원', '').strip()
        self.ent_amount.delete(0, 'end')
        self.ent_amount.insert(0, f"{int(clean_amt):,}" if clean_amt.isdigit() else clean_amt)
        
        # 메모
        self.ent_note.delete(0, 'end')
        self.ent_note.insert(0, note)

    def update_category_list(self, event=None):
        t_type = self.var_type.get()
        default_cats = self.income_categories if t_type == "수입" else self.expense_categories
        
        custom_cats = self.db.get_unique_categories(t_type)
        merged_cats = list(default_cats)
        for cat in custom_cats:
            if cat not in merged_cats:
                merged_cats.append(cat)
                
        self.cb_category.config(values=merged_cats)
        self.cb_category.current(0)

    def format_amount(self, event):
        if event.keysym in ['Left', 'Right', 'Up', 'Down', 'BackSpace', 'Delete']: return
        content = self.ent_amount.get().replace(',', '')
        if not content: return
        
        clean_content = ''.join(filter(str.isdigit, content))
        if clean_content:
            self.ent_amount.delete(0, 'end')
            self.ent_amount.insert(0, f"{int(clean_content):,}")
        else:
            self.ent_amount.delete(0, 'end')

    def prev_month(self):
        if self.current_month == 1:
            self.current_month = 12
            self.current_year -= 1
        else:
            self.current_month -= 1
        self.lbl_month.config(text=f"{self.current_year}년 {self.current_month}월")
        self.ent_search.delete(0, 'end')
        self.refresh_data()

    def next_month(self):
        if self.current_month == 12:
            self.current_month = 1
            self.current_year += 1
        else:
            self.current_month += 1
        self.lbl_month.config(text=f"{self.current_year}년 {self.current_month}월")
        self.ent_search.delete(0, 'end')
        self.refresh_data()

    def _add_summary_row(self, records):
        if not records: return
        t_in = sum(r[4] for r in records if r[2] == "수입")
        t_out = sum(r[4] for r in records if r[2] == "지출")
        
        summary_val = f"{t_in - t_out:,}원"
        memo_str = f"(수입: {t_in:,} / 지출: {t_out:,})"
        self.tv.insert('', 'end', values=('', '[총 합계]', '-', '-', summary_val, memo_str), tags=('summary',))

    def search_data(self):
        keyword = self.ent_search.get().strip()
        if not keyword:
            self.refresh_data()
            return
            
        for item in self.tv.get_children(): self.tv.delete(item)
        records = self.db.search_transactions(keyword)
        for r in records:
            self.tv.insert('', 'end', values=(r[0], r[1], r[2], r[3], f"{r[4]:,}원", r[5]), tags=(r[2],))
            
        self._add_summary_row(records)
        self.lbl_month.config(text=f"검색결과: '{keyword}'")

    def refresh_data(self):
        self.lbl_month.config(text=f"{self.current_year}년 {self.current_month}월")
        self.ent_search.delete(0, 'end')
        
        # 1. 트리뷰 갱신
        for item in self.tv.get_children(): self.tv.delete(item)
        records = self.db.get_transactions_by_month(self.current_year, self.current_month)
        for r in records:
            self.tv.insert('', 'end', values=(r[0], r[1], r[2], r[3], f"{r[4]:,}원", r[5]), tags=(r[2],))
            
        self._add_summary_row(records)
            
        # 2. 요약 및 예산 갱신
        summary = self.db.get_monthly_summary(self.current_year, self.current_month)
        self.lbl_income.config(text=f"수입: +{summary['income']:,} 원")
        self.lbl_expense.config(text=f"지출: -{summary['expense']:,} 원")
        self.lbl_balance.config(text=f"잔액: {summary['balance']:,} 원")
        
        budget_str = self.db.get_setting("monthly_budget", "0")
        budget = int(budget_str) if budget_str.isdigit() else 0
            
        if budget > 0:
            percent = min((summary['expense'] / budget) * 100, 100)
            self.progress_budget['value'] = percent
            
            if percent >= 100: color = "red"
            elif percent >= 80: color = "orange"
            else: color = "green"
            
            self.lbl_budget_status.config(text=f"목표 예산: {budget:,} 원 | 현재 {percent:.1f}% 사용", foreground=color)
        else:
            self.progress_budget['value'] = 0
            self.lbl_budget_status.config(text="목표 예산: 미설정 (상단 [예산 설정] 버튼 클릭)", foreground="black")
            
        # 3. 차트 갱신
        self.update_pie_chart(summary['expense_by_category'])
        self.update_line_chart()
        self.update_daily_chart()

    def update_pie_chart(self, exp_dict):
        self.ax_pie.clear()
        if exp_dict:
            total_exp = sum(exp_dict.values())
            merged_dict, 기타_합계 = {}, 0
            for cat, amt in exp_dict.items():
                if (amt / total_exp) < 0.03: 기타_합계 += amt
                else: merged_dict[cat] = amt
            if 기타_합계 > 0: merged_dict['기타'] = merged_dict.get('기타', 0) + 기타_합계
                    
            self.ax_pie.pie(list(merged_dict.values()), labels=list(merged_dict.keys()), autopct='%1.1f%%', startangle=90, textprops={'fontsize': 9})
            self.ax_pie.axis('equal')
        else:
            self.ax_pie.text(0.5, 0.5, '지출 내역이 없습니다.', ha='center', va='center', fontsize=12, color='gray')
            self.ax_pie.axis('off')
        self.canvas_pie.draw()

    def update_line_chart(self):
        self.ax_line.clear()
        annual_data = self.db.get_annual_summary(self.current_year)
        months = [f"{i:02d}" for i in range(1, 13)]
        
        self.ax_line.plot(months, [annual_data['income'].get(m, 0) for m in months], marker='o', color='blue', label='수입')
        self.ax_line.plot(months, [annual_data['expense'].get(m, 0) for m in months], marker='x', color='red', label='지출')
        
        self.ax_line.set_title(f"{self.current_year}년 수입/지출 트렌드", fontsize=12)
        self.ax_line.legend()
        self.ax_line.grid(True, linestyle='--', alpha=0.6)
        self.ax_line.ticklabel_format(style='plain', axis='y')
        self.canvas_line.draw()

    def update_daily_chart(self):
        self.ax_daily.clear()
        daily_cat_data = self.db.get_daily_summary(self.current_year, self.current_month)
        
        if not daily_cat_data:
            self.ax_daily.text(0.5, 0.5, '지출 내역이 없습니다.', ha='center', va='center', fontsize=12, color='gray')
            self.ax_daily.axis('off')
            self.current_chart_days = []
        else:
            days = sorted(daily_cat_data.keys())
            self.current_chart_days = days
            
            # 모든 날짜에 쓰인 카테고리 종류 수집
            all_cats = set()
            for day_cats in daily_cat_data.values():
                all_cats.update(day_cats.keys())
            all_cats = sorted(list(all_cats))
            
            bottoms = [0] * len(days)
            
            for cat in all_cats:
                amounts = [daily_cat_data[d].get(cat, 0) for d in days]
                self.ax_daily.bar(days, amounts, bottom=bottoms, label=cat)
                bottoms = [b + a for b, a in zip(bottoms, amounts)]
            
            self.ax_daily.set_title(f"{self.current_month}월 일간 카테고리별 지출", fontsize=12)
            self.ax_daily.ticklabel_format(style='plain', axis='y')
            self.ax_daily.tick_params(axis='x', rotation=45)
            self.ax_daily.grid(True, axis='y', linestyle='--', alpha=0.6)
            
            # 막대 위에 총 지출액 표시
            for i, day in enumerate(days):
                total_val = bottoms[i]
                if total_val > 0:
                    self.ax_daily.text(i, total_val, f'{int(total_val):,}', va='bottom', ha='center', fontsize=8)
            
            # 범례 표시
            if all_cats:
                self.ax_daily.legend(bbox_to_anchor=(1.01, 1), loc='upper left', fontsize=8)
                self.fig_daily.tight_layout()
                
        self.canvas_daily.draw()

    def on_tab_changed(self, event):
        self.canvas_pie.draw()
        self.canvas_line.draw()
        self.canvas_daily.draw()

    def on_daily_chart_click(self, event):
        if event.inaxes != self.ax_daily or event.xdata is None:
            return
        
        import math
        idx = math.floor(event.xdata + 0.5)
        
        if hasattr(self, 'current_chart_days') and 0 <= idx < len(self.current_chart_days):
            clicked_day = self.current_chart_days[idx]
            keyword = f"{self.current_year}-{self.current_month:02d}-{clicked_day}"
            
            self.ent_search.delete(0, 'end')
            self.ent_search.insert(0, keyword)
            self.search_data()
            
            messagebox.showinfo("필터링 완료", f"그래프를 클릭하여 {keyword}의 내역을 검색했습니다.\n좌측 '상세 내역' 표를 확인해 주세요!")

    # ----------------------------------------------------
    # 버튼 액션 모음
    # ----------------------------------------------------
    def set_budget(self):
        current = self.db.get_setting("monthly_budget", "0")
        new_budget = simpledialog.askinteger("예산 설정", "이번 달 목표 지출 예산(원):", initialvalue=int(current), parent=self.root)
        if new_budget is not None:
            self.db.set_setting("monthly_budget", new_budget)
            self.refresh_data()
            messagebox.showinfo("설정 완료", f"목표 예산이 {new_budget:,}원으로 설정되었습니다.")

    def open_recurring_manager(self):
        RecurringManagerWindow(self.root, self.db, self.refresh_data)

    def load_recurring_to_current_month(self):
        templates = self.db.get_recurring_templates()
        if not templates:
            messagebox.showinfo("알림", "등록된 고정 내역이 없습니다.\n상단 [고정 지출 관리] 메뉴에서 먼저 등록해주세요.")
            return
            
        if messagebox.askyesno("일괄 등록", f"총 {len(templates)}건의 고정 내역을 이번 달({self.current_month}월) 장부에 등록하시겠습니까?"):
            target_date = datetime.datetime.now()
            if self.current_year != target_date.year or self.current_month != target_date.month:
                target_date_str = f"{self.current_year}-{self.current_month:02d}-01"
            else:
                target_date_str = target_date.strftime("%Y-%m-%d")
                
            for r in templates:
                self.db.add_transaction(target_date_str, r[1], r[2], r[3], f"[고정] {r[4]}")
            
            self.refresh_data()
            messagebox.showinfo("등록 완료", f"{len(templates)}건의 고정 내역이 장부에 추가되었습니다.")

    def add_record(self):
        date, t_type, category = self.ent_date.get().strip(), self.var_type.get(), self.cb_category.get()
        amt_str, note = self.ent_amount.get().strip().replace(',', ''), self.ent_note.get().strip()
        
        if not date or not amt_str:
            messagebox.showwarning("입력 오류", "날짜와 금액을 모두 입력해주세요."); return
        try:
            amount = int(amt_str)
            datetime.datetime.strptime(date, "%Y-%m-%d")
        except ValueError:
            messagebox.showwarning("입력 오류", "금액은 숫자만, 날짜는 YYYY-MM-DD 형식으로 입력하세요."); return
            
        self.db.add_transaction(date, t_type, category, amount, note)
        self.ent_amount.delete(0, 'end'); self.ent_note.delete(0, 'end')
        self.refresh_data()
        
    def delete_record(self):
        selection = self.tv.selection()
        if not selection:
            messagebox.showwarning("선택 안 됨", "삭제할 내역을 선택해주세요."); return
        if messagebox.askyesno("삭제 확인", "선택한 내역을 정말 삭제하시겠습니까?"):
            for item in selection:
                self.db.delete_transaction(self.tv.item(item, 'values')[0])
            self.refresh_data()
            
    def export_excel(self):
        filepath = self.db.export_to_excel(f"가계부_내역_{self.current_year}년{self.current_month}월.xlsx")
        if filepath == "PERMISSION_ERROR":
            messagebox.showerror("저장 실패", "이미 해당 엑셀 파일이 열려있습니다.\n파일을 닫은 후 다시 시도해주세요.")
        elif filepath:
            messagebox.showinfo("저장 완료", f"엑셀 파일이 저장되었습니다.\n{filepath}")
        else:
            messagebox.showinfo("알림", "저장할 내역이 없습니다.")

    def open_receipt_ocr(self):
        try:
            import requests
            import mimetypes
        except ImportError:
            messagebox.showerror("오류", "requests 라이브러리가 필요합니다.\n터미널에서 'pip install requests'를 실행해주세요.")
            return

        api_key = self.db.get_setting("gemini_api_key", "")
        if not api_key:
            api_key = simpledialog.askstring("API 키 설정", "Google Gemini API 키를 입력하세요 (AIza...):", show='*', parent=self.root)
            if not api_key:
                return
            self.db.set_setting("gemini_api_key", api_key)

        file_path = filedialog.askopenfilename(
            title="영수증 사진 선택",
            filetypes=[("Image files", "*.jpg *.jpeg *.png")],
            parent=self.root
        )
        if not file_path:
            return

        wait_win = tk.Toplevel(self.root)
        wait_win.title("분석 중")
        wait_win.geometry("300x120")
        wait_win.transient(self.root)
        wait_win.grab_set()
        ttk.Label(wait_win, text="AI가 영수증을 분석하고 있습니다...\n할인 내역까지 꼼꼼히 계산 중입니다! 🕒", justify='center').pack(expand=True)

        def process_image():
            try:
                with open(file_path, "rb") as f:
                    img_data = base64.b64encode(f.read()).decode('utf-8')
                    
                mime_type, _ = mimetypes.guess_type(file_path)
                if not mime_type:
                    mime_type = "image/jpeg"

                prompt = """
영수증 사진을 분석해서 가계부 내역(지출)으로 정리해주세요.
응답은 반드시 아래 JSON 배열 형식으로만 반환해주세요. (다른 설명 텍스트 없이 JSON만 반환)
[
  {
    "category": "식비",
    "amount": 4480,
    "note": "서울우유 후레쉬밀크"
  },
  {
    "category": "식비",
    "amount": 3180,
    "note": "농심 안성탕면"
  }
]
- 카테고리는 다음 중에서만 선택하세요: 식비, 교통/차량, 문화/생활, 주거/통신, 건강/의료, 쇼핑, 기타지출
- 절대 품목을 뭉뚱그려 합치지 마세요. 영수증에 적힌 **모든 개별 상품(내역)을 하나하나 각각의 항목으로 분리해서 전부 작성**해주세요.
- 행사 할인이나 쿠폰 할인이 적용된 상품이 있다면, 해당 상품의 원래 가격에서 할인된 금액을 뺀 '실제 결제 금액(할인가)'을 amount에 적어주세요.
- amount는 쉼표 없는 순수 숫자만 적어주세요.
"""
                url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash:generateContent?key={api_key}"
                payload = {
                    "contents": [{
                        "parts": [
                            {"text": prompt},
                            {"inline_data": {"mime_type": mime_type, "data": img_data}}
                        ]
                    }]
                }
                
                resp = requests.post(url, json=payload, timeout=60)
                
                if resp.status_code != 200:
                    if resp.status_code == 400 and "API key" in resp.text:
                        raise Exception("api_key")
                    elif resp.status_code == 503:
                        raise Exception("503_unavailable")
                    else:
                        raise Exception(f"API 에러: {resp.status_code} {resp.text}")

                data = resp.json()
                result_text = data['candidates'][0]['content']['parts'][0]['text'].strip()
                
                if result_text.startswith("```json"):
                    result_text = result_text.replace("```json", "").replace("```", "").strip()
                elif result_text.startswith("```"):
                    result_text = result_text.replace("```", "").strip()

                parsed_data = json.loads(result_text)
                self.root.after(0, lambda: self._on_ocr_success(wait_win, parsed_data))
            except Exception as e:
                err_msg = str(e)
                self.root.after(0, lambda msg=err_msg: self._on_ocr_fail(wait_win, msg))

        threading.Thread(target=process_image, daemon=True).start()

    def _on_ocr_success(self, wait_win, data):
        wait_win.destroy()
        today = datetime.datetime.now().strftime("%Y-%m-%d")
        count = 0
        for item in data:
            cat = item.get("category", "기타지출")
            if cat not in self.expense_categories:
                cat = "기타지출"
            self.db.add_transaction(today, "지출", cat, int(item.get("amount", 0)), item.get("note", "영수증 인식"))
            count += 1
            
        self.refresh_data()
        messagebox.showinfo("완료", f"{count}건의 내역이 영수증에서 자동 추출되어 추가되었습니다!")

    def _on_ocr_fail(self, wait_win, err_msg):
        wait_win.destroy()
        if err_msg == "api_key":
            messagebox.showerror("오류", "유효하지 않은 API 키입니다.\n설정을 초기화합니다. 다시 시도해주세요.")
            self.db.set_setting("gemini_api_key", "")
        elif err_msg == "503_unavailable":
            messagebox.showwarning("서버 혼잡", "현재 구글 AI 서버에 사용자가 몰려 처리가 지연되고 있습니다.\n잠시 후 다시 시도해주세요.")
        else:
            messagebox.showerror("분석 실패", f"영수증 분석 중 오류가 발생했습니다.\n{err_msg}")


if __name__ == "__main__":
    root = tk.Tk()
    app = HouseholdApp(root)
    root.mainloop()
