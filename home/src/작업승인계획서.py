import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
import os
import json
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime

class WorkApprovalApp:
    def __init__(self, root):
        self.root = root
        self.root.title("작업승인계획서 자동 생성기")
        self.root.geometry("650x950")
        self.root.resizable(True, True)
        
        style = ttk.Style()
        style.theme_use('clam')
        
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=5, pady=5)
        
        main_frame = ttk.Frame(self.notebook, padding=15)
        self.notebook.add(main_frame, text="작업승인계획서")
        
        # Add TBM Form Tab
        try:
            from tbm_tab import TBMFormTab
            self.tab_tbm = ttk.Frame(self.notebook)
            self.notebook.add(self.tab_tbm, text='TBM 회의록')
            self.tbm_manager = TBMFormTab(self.tab_tbm, main_app=self)
            self.tbm_manager.pack(fill='both', expand=True)
        except Exception as e:
            print(f"TBM 모듈 로드 실패: {e}")
            
        # Add Risk Assessment Tab
        try:
            import importlib.util
            import sys
            
            # 위험성 평가표 all.py 파일이 띄어쓰기가 있어 importlib 사용
            module_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "위험성 평가표 all.py")
            spec = importlib.util.spec_from_file_location("risk_module", module_path)
            risk_module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(risk_module)
            
            self.tab_risk = ttk.Frame(self.notebook)
            self.notebook.add(self.tab_risk, text='위험성 평가표')
            # The RiskAssessmentApp doesn't have a pack method for itself, it packs into root
            self.risk_manager = risk_module.RiskAssessmentApp(self.tab_risk)
        except Exception as e:
            print(f"위험성 평가표 모듈 로드 실패: {e}")
        
        # Title
        ttk.Label(main_frame, text="[서식 3] 작업승인계획서 생성기", font=('Malgun Gothic', 16, 'bold')).pack(pady=(0, 15))
        
        # Form Container
        form_frame = ttk.Frame(main_frame)
        form_frame.pack(fill='both', expand=True)
        
        row_idx = 0
        
        # 1. Date & Company
        ttk.Label(form_frame, text="일자:").grid(row=row_idx, column=0, sticky='e', padx=5, pady=5)
        
        date_frame = ttk.Frame(form_frame)
        date_frame.grid(row=row_idx, column=1, sticky='w', padx=5, pady=5)
        
        self.ent_date = ttk.Combobox(date_frame, width=30)
        self.ent_date.insert(0, datetime.now().strftime("%Y년 %m월 %d일 O요일"))
        self.ent_date.pack(side='left')
        self.ent_date.bind('<<ComboboxSelected>>', self.on_date_select)
        
        btn_calc_date = ttk.Button(date_frame, text="[이전 날짜에서 누계 불러오기]", command=self.load_previous_totals)
        btn_calc_date.pack(side='left', padx=10)
        
        self.btn_generate = ttk.Button(date_frame, text="엑셀 생성 (승인계획서)", command=self.generate_files)
        self.btn_generate.pack(side='right', padx=2)
        
        self.btn_unified = ttk.Button(date_frame, text="🔥 일일 안전서류 통합 엑셀 생성", command=self.generate_unified_excel)
        self.btn_unified.pack(side='right', padx=10)
        
        self.lbl_status = ttk.Label(date_frame, text="대기 중...", foreground="gray")
        self.lbl_status.pack(side='left', padx=5)
        
        row_idx += 1
        
        ttk.Label(form_frame, text="수급업체:").grid(row=row_idx, column=0, sticky='e', padx=5, pady=5)
        self.ent_company = ttk.Entry(form_frame, width=50)
        self.ent_company.insert(0, "서울검사(주)")
        self.ent_company.grid(row=row_idx, column=1, sticky='w', padx=5, pady=5)
        row_idx += 1
        
        # 2. 총 투입 현황
        lbl_sec1 = ttk.Label(form_frame, text="1. 총 투입 현황", font=('Malgun Gothic', 10, 'bold'))
        lbl_sec1.grid(row=row_idx, column=0, columnspan=2, sticky='w', pady=(10, 5))
        row_idx += 1
        
        ttk.Label(form_frame, text="총 작업 개소:").grid(row=row_idx, column=0, sticky='e', padx=5, pady=2)
        self.ent_locations = ttk.Entry(form_frame, width=50)
        self.ent_locations.insert(0, "00 개소")
        self.ent_locations.grid(row=row_idx, column=1, sticky='w', padx=5, pady=2)
        row_idx += 1
        
        ttk.Label(form_frame, text="총 인원 (계):").grid(row=row_idx, column=0, sticky='e', padx=5, pady=2)
        self.ent_personnel = ttk.Entry(form_frame, width=50)
        self.ent_personnel.insert(0, "00 명")
        self.ent_personnel.grid(row=row_idx, column=1, sticky='w', padx=5, pady=2)
        row_idx += 1
        
        ttk.Label(form_frame, text="총 장비 (계):").grid(row=row_idx, column=0, sticky='e', padx=5, pady=2)
        self.ent_equipment = ttk.Entry(form_frame, width=50)
        self.ent_equipment.insert(0, "00 대")
        self.ent_equipment.grid(row=row_idx, column=1, sticky='w', padx=5, pady=2)
        row_idx += 1
        
        ttk.Label(form_frame, text="RT / 크롤러 투입:").grid(row=row_idx, column=0, sticky='e', padx=5, pady=2)
        self.ent_rt = ttk.Entry(form_frame, width=50)
        self.ent_rt.insert(0, "RT조사기 0대, 크롤러 0대")
        self.ent_rt.grid(row=row_idx, column=1, sticky='w', padx=5, pady=2)
        row_idx += 1
        
        ttk.Label(form_frame, text="기타 장비 현황:").grid(row=row_idx, column=0, sticky='e', padx=5, pady=2)
        self.ent_etc = ttk.Entry(form_frame, width=50)
        self.ent_etc.insert(0, "UT 0대, PT 0세트, 발전기 0대")
        self.ent_etc.grid(row=row_idx, column=1, sticky='w', padx=5, pady=2)
        row_idx += 1
        
        # 3. 팀별 세부 작업
        sec2_frame = ttk.Frame(form_frame)
        sec2_frame.grid(row=row_idx, column=0, columnspan=2, sticky='w', pady=(10, 5))
        
        lbl_sec2 = ttk.Label(sec2_frame, text="2. 팀별 세부 작업 (금일 작업 내용 및 현황)", font=('Malgun Gothic', 10, 'bold'))
        lbl_sec2.pack(side='left', padx=(0, 20))
        
        self.team_a_active = tk.BooleanVar(value=True)
        self.team_b_active = tk.BooleanVar(value=True)
        self.team_a_rt = tk.BooleanVar(value=True)
        self.team_a_ut = tk.BooleanVar(value=False)
        self.team_a_pt = tk.BooleanVar(value=False)
        self.team_b_rt = tk.BooleanVar(value=False)
        self.team_b_ut = tk.BooleanVar(value=True)
        self.team_b_pt = tk.BooleanVar(value=True)
        
        ttk.Checkbutton(sec2_frame, text="A팀 작업 진행", variable=self.team_a_active, command=self.toggle_team_mode).pack(side='left', padx=5)
        ttk.Checkbutton(sec2_frame, text="B팀 작업 진행", variable=self.team_b_active, command=self.toggle_team_mode).pack(side='left', padx=5)
        row_idx += 1
        
        # A팀
        lbl_a = ttk.Label(form_frame, text="A팀 개소/내용:")
        lbl_a.grid(row=row_idx, column=0, sticky='ne', padx=5, pady=2)
        
        container_a = ttk.Frame(form_frame)
        container_a.grid(row=row_idx, column=1, sticky='w', padx=5, pady=0)
        
        frame_a = ttk.Frame(container_a)
        frame_a.pack(anchor='w', pady=(0, 2))
        ttk.Label(frame_a, text="작업개소:").pack(side='left')
        self.ent_team_a_loc = ttk.Entry(frame_a, width=8)
        self.ent_team_a_loc.insert(0, "00")
        self.ent_team_a_loc.pack(side='left', padx=2)
        ttk.Label(frame_a, text="개소").pack(side='left', padx=(0, 15))
        
        ttk.Checkbutton(frame_a, text="RT", variable=self.team_a_rt).pack(side='left', padx=2)
        ttk.Checkbutton(frame_a, text="UT", variable=self.team_a_ut).pack(side='left', padx=2)
        ttk.Checkbutton(frame_a, text="PT", variable=self.team_a_pt).pack(side='left', padx=2)
        
        self.txt_team_a = tk.Text(container_a, width=50, height=4, font=('Malgun Gothic', 9))
        self.txt_team_a.insert('1.0', "[구간: OO천 ~ OOO천]\n(내용) 30\" 주배관 맞대기 용접부 방사선투과검사(RT)\n※ 작업시간: (08:00~17:00)\n\n[투입 현황]\n인원: 00명\n장비: 조사기 1, 크롤러 1, 차폐막 2")
        self.txt_team_a.pack(anchor='w')
        row_idx += 1
        
        # B팀
        lbl_b = ttk.Label(form_frame, text="B팀 개소/내용:")
        lbl_b.grid(row=row_idx, column=0, sticky='ne', padx=5, pady=2)
        
        container_b = ttk.Frame(form_frame)
        container_b.grid(row=row_idx, column=1, sticky='w', padx=5, pady=0)
        
        frame_b = ttk.Frame(container_b)
        frame_b.pack(anchor='w', pady=(0, 2))
        ttk.Label(frame_b, text="작업개소:").pack(side='left')
        self.ent_team_b_loc = ttk.Entry(frame_b, width=8)
        self.ent_team_b_loc.insert(0, "00")
        self.ent_team_b_loc.pack(side='left', padx=2)
        ttk.Label(frame_b, text="개소").pack(side='left', padx=(0, 15))
        
        ttk.Checkbutton(frame_b, text="RT", variable=self.team_b_rt).pack(side='left', padx=2)
        ttk.Checkbutton(frame_b, text="UT", variable=self.team_b_ut).pack(side='left', padx=2)
        ttk.Checkbutton(frame_b, text="PT", variable=self.team_b_pt).pack(side='left', padx=2)
        
        self.txt_team_b = tk.Text(container_b, width=50, height=4, font=('Malgun Gothic', 9))
        self.txt_team_b.insert('1.0', "[구간: OO관리소 내부]\n(내용) Tie-in 필릿 용접부 초음파(UT) 및 침투(PT)\n※ 작업시간: (08:00~17:00)\n\n[투입 현황]\n인원: 00명\n장비: UT 1, PT 1")
        self.txt_team_b.pack(anchor='w')
        row_idx += 1

        # 4. 기타 진행 현황 (자동 계산)
        lbl_sec3 = ttk.Label(form_frame, text="3. 진행 현황 (목표량: RT 24,536매, UT 319.02M, PT 338.63M)", font=('Malgun Gothic', 10, 'bold'))
        lbl_sec3.grid(row=row_idx, column=0, columnspan=2, sticky='w', pady=(10, 5))
        row_idx += 1
        
        # RT Input
        frame_rt = ttk.Frame(form_frame)
        frame_rt.grid(row=row_idx, column=1, sticky='w')
        ttk.Label(form_frame, text="RT (매):").grid(row=row_idx, column=0, sticky='e', padx=5, pady=2)
        ttk.Label(frame_rt, text="전일누계").pack(side='left')
        self.ent_rt_prev = ttk.Entry(frame_rt, width=15)
        self.ent_rt_prev.insert(0, "0")
        self.ent_rt_prev.pack(side='left', padx=(5, 15))
        
        ttk.Label(frame_rt, text="금일계획").pack(side='left')
        self.ent_rt_today = ttk.Entry(frame_rt, width=15)
        self.ent_rt_today.insert(0, "0")
        self.ent_rt_today.pack(side='left', padx=5)
        row_idx += 1
        
        # UT Input
        frame_ut = ttk.Frame(form_frame)
        frame_ut.grid(row=row_idx, column=1, sticky='w')
        ttk.Label(form_frame, text="UT (M):").grid(row=row_idx, column=0, sticky='e', padx=5, pady=2)
        ttk.Label(frame_ut, text="전일누계").pack(side='left')
        self.ent_ut_prev = ttk.Entry(frame_ut, width=15)
        self.ent_ut_prev.insert(0, "0")
        self.ent_ut_prev.pack(side='left', padx=(5, 15))
        
        ttk.Label(frame_ut, text="금일계획").pack(side='left')
        self.ent_ut_today = ttk.Entry(frame_ut, width=15)
        self.ent_ut_today.insert(0, "0")
        self.ent_ut_today.pack(side='left', padx=5)
        row_idx += 1
        
        # PT Input
        frame_pt = ttk.Frame(form_frame)
        frame_pt.grid(row=row_idx, column=1, sticky='w')
        ttk.Label(form_frame, text="PT (M):").grid(row=row_idx, column=0, sticky='e', padx=5, pady=2)
        ttk.Label(frame_pt, text="전일누계").pack(side='left')
        self.ent_pt_prev = ttk.Entry(frame_pt, width=15)
        self.ent_pt_prev.insert(0, "0")
        self.ent_pt_prev.pack(side='left', padx=(5, 15))
        
        ttk.Label(frame_pt, text="금일계획").pack(side='left')
        self.ent_pt_today = ttk.Entry(frame_pt, width=15)
        self.ent_pt_today.insert(0, "0")
        self.ent_pt_today.pack(side='left', padx=5)
        row_idx += 1
        
        ttk.Label(form_frame, text="요청사항:").grid(row=row_idx, column=0, sticky='ne', padx=5, pady=2)
        self.txt_req = tk.Text(form_frame, width=50, height=3, font=('Malgun Gothic', 9))
        self.txt_req.insert('1.0', "- 익일 야간 RT 작업 승인 별도 요청\n- 크롤러 전원 지원 요망")
        self.txt_req.grid(row=row_idx, column=1, sticky='w', padx=5, pady=2)
        row_idx += 1
        
        # Generate Button moved to the top date_frame
        
        # Load saved config
        self.load_config()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def get_history_path(self):
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), 'work_approval_history.json')

    def load_history_data(self):
        path = self.get_history_path()
        if os.path.exists(path):
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                pass
        return {}

    def save_history_data(self, history):
        try:
            with open(self.get_history_path(), 'w', encoding='utf-8') as f:
                json.dump(history, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"Failed to save history: {e}")

    def on_date_select(self, event=None):
        date_str = self.ent_date.get()
        history = self.load_history_data()
        if date_str in history:
            self.populate_ui(history[date_str], date_str)

    def load_previous_totals(self):
        history = self.load_history_data()
        if not history:
            messagebox.showinfo("알림", "저장된 이전 기록이 없습니다.")
            return
            
        current_date_str = self.ent_date.get()
        dates = sorted(list(history.keys()))
        
        prev_date = None
        for d in reversed(dates):
            if d < current_date_str:
                prev_date = d
                break
                
        if prev_date is None:
            prev_date = dates[-1]
            
        prev_data = history[prev_date]
        
        try:
            rt_total = self.parse_float(prev_data.get('rt_prev', '0')) + self.parse_float(prev_data.get('rt_today', '0'))
            ut_total = self.parse_float(prev_data.get('ut_prev', '0')) + self.parse_float(prev_data.get('ut_today', '0'))
            pt_total = self.parse_float(prev_data.get('pt_prev', '0')) + self.parse_float(prev_data.get('pt_today', '0'))
            
            def fmt(v): return str(int(v)) if v.is_integer() else str(v)
            
            self.ent_rt_prev.delete(0, tk.END)
            self.ent_rt_prev.insert(0, fmt(rt_total))
            
            self.ent_ut_prev.delete(0, tk.END)
            self.ent_ut_prev.insert(0, fmt(ut_total))
            
            self.ent_pt_prev.delete(0, tk.END)
            self.ent_pt_prev.insert(0, fmt(pt_total))
            
            self.ent_rt_today.delete(0, tk.END)
            self.ent_rt_today.insert(0, "0")
            self.ent_ut_today.delete(0, tk.END)
            self.ent_ut_today.insert(0, "0")
            self.ent_pt_today.delete(0, tk.END)
            self.ent_pt_today.insert(0, "0")
            
            messagebox.showinfo("성공", f"[{prev_date}] 의 누계 데이터를 불러왔습니다.\n(오늘 계획은 0으로 초기화됨)")
        except Exception as e:
            messagebox.showerror("오류", f"누계 불러오기 실패: {e}")

    def save_config(self):
        history = self.load_history_data()
        current_date = self.ent_date.get()
        
        data = {
            'company': self.ent_company.get(),
            'locations': self.ent_locations.get(),
            'personnel': self.ent_personnel.get(),
            'equipment': self.ent_equipment.get(),
            'rt': self.ent_rt.get(),
            'etc': self.ent_etc.get(),
            'team_a': self.txt_team_a.get('1.0', tk.END).strip(),
            'team_b': self.txt_team_b.get('1.0', tk.END).strip(),
            'team_a_loc': self.ent_team_a_loc.get().strip(),
            'team_b_loc': self.ent_team_b_loc.get().strip(),
            'team_a_active': self.team_a_active.get(),
            'team_b_active': self.team_b_active.get(),
            'team_a_rt': self.team_a_rt.get(),
            'team_a_ut': self.team_a_ut.get(),
            'team_a_pt': self.team_a_pt.get(),
            'team_b_rt': self.team_b_rt.get(),
            'team_b_ut': self.team_b_ut.get(),
            'team_b_pt': self.team_b_pt.get(),
            'rt_prev': self.ent_rt_prev.get(),
            'rt_today': self.ent_rt_today.get(),
            'ut_prev': self.ent_ut_prev.get(),
            'ut_today': self.ent_ut_today.get(),
            'pt_prev': self.ent_pt_prev.get(),
            'pt_today': self.ent_pt_today.get(),
            'req': self.txt_req.get('1.0', tk.END).strip(),
        }
        
        history[current_date] = data
        
        # Cascading update for future dates
        dates = [k for k in history.keys() if not k.startswith('_')]
        sorted_dates = sorted(dates)
        if current_date in sorted_dates:
            idx = sorted_dates.index(current_date)
            
            def p(val): return self.parse_float(val)
            def fmt(v): return str(int(v)) if v.is_integer() else str(v)
            
            cur_rt_total = p(data['rt_prev']) + p(data['rt_today'])
            cur_ut_total = p(data['ut_prev']) + p(data['ut_today'])
            cur_pt_total = p(data['pt_prev']) + p(data['pt_today'])
            
            for i in range(idx + 1, len(sorted_dates)):
                next_date = sorted_dates[i]
                next_data = history[next_date]
                
                next_data['rt_prev'] = fmt(cur_rt_total)
                next_data['ut_prev'] = fmt(cur_ut_total)
                next_data['pt_prev'] = fmt(cur_pt_total)
                
                cur_rt_total = cur_rt_total + p(next_data.get('rt_today', '0'))
                cur_ut_total = cur_ut_total + p(next_data.get('ut_today', '0'))
                cur_pt_total = cur_pt_total + p(next_data.get('pt_today', '0'))
                
        self.save_history_data(history)
        
        dates = [k for k in history.keys() if not k.startswith('_')]
        self.ent_date['values'] = sorted(dates, reverse=True)

    def on_close(self):
        history = self.load_history_data()
        history['_window_geometry'] = self.root.geometry()
        self.save_history_data(history)
        self.root.destroy()

    def load_config(self):
        history = self.load_history_data()
        if not history:
            return

        if '_window_geometry' in history:
            try:
                geom = history['_window_geometry']
                if 'x' in geom:
                    w, rest = geom.split('x', 1)
                    if '+' in rest:
                        h_str, pos = rest.split('+', 1)
                        pos = '+' + pos
                    elif '-' in rest:
                        h_str, pos = rest.split('-', 1)
                        pos = '-' + pos
                    else:
                        h_str = rest
                        pos = ''
                    
                    h = int(h_str)
                    if h < 950:
                        h = 950
                    self.root.geometry(f"{w}x{h}{pos}")
                else:
                    self.root.geometry(geom)
            except:
                pass

        dates = [k for k in history.keys() if not k.startswith('_')]
        if not dates:
            return
            
        sorted_dates = sorted(dates, reverse=True)
        self.ent_date['values'] = sorted_dates

        most_recent_date = sorted_dates[0]
        self.ent_date.delete(0, tk.END)
        self.ent_date.insert(0, most_recent_date)
        
        self.populate_ui(history[most_recent_date], most_recent_date)

    def populate_ui(self, data, date_str=None):
        def set_ent(ent, val):
            if val is not None:
                ent.delete(0, tk.END)
                ent.insert(0, str(val))
                
        def set_txt(txt, val):
            if val is not None:
                txt.delete('1.0', tk.END)
                txt.insert('1.0', str(val))
        
        if date_str:
            set_ent(self.ent_date, date_str)
            
        set_ent(self.ent_company, data.get('company'))
        set_ent(self.ent_locations, data.get('locations'))
        set_ent(self.ent_personnel, data.get('personnel'))
        set_ent(self.ent_equipment, data.get('equipment'))
        set_ent(self.ent_rt, data.get('rt'))
        set_ent(self.ent_etc, data.get('etc'))
        
        set_txt(self.txt_team_a, data.get('team_a'))
        set_txt(self.txt_team_b, data.get('team_b'))
        
        if hasattr(self, 'ent_team_a_loc'):
            set_ent(self.ent_team_a_loc, data.get('team_a_loc', '00'))
        if hasattr(self, 'ent_team_b_loc'):
            set_ent(self.ent_team_b_loc, data.get('team_b_loc', '00'))
        
        if 'team_a_active' in data:
            self.team_a_active.set(data['team_a_active'])
        if 'team_b_active' in data:
            self.team_b_active.set(data['team_b_active'])
            
        if 'team_a_rt' in data: self.team_a_rt.set(data['team_a_rt'])
        if 'team_a_ut' in data: self.team_a_ut.set(data['team_a_ut'])
        if 'team_a_pt' in data: self.team_a_pt.set(data['team_a_pt'])
        
        if 'team_b_rt' in data: self.team_b_rt.set(data['team_b_rt'])
        if 'team_b_ut' in data: self.team_b_ut.set(data['team_b_ut'])
        if 'team_b_pt' in data: self.team_b_pt.set(data['team_b_pt'])
            
        self.toggle_team_mode()
        
        set_ent(self.ent_rt_prev, data.get('rt_prev'))
        set_ent(self.ent_rt_today, data.get('rt_today'))
        set_ent(self.ent_ut_prev, data.get('ut_prev'))
        set_ent(self.ent_ut_today, data.get('ut_today'))
        set_ent(self.ent_pt_prev, data.get('pt_prev'))
        set_ent(self.ent_pt_today, data.get('pt_today'))
        
        set_txt(self.txt_req, data.get('req'))

    def toggle_team_mode(self):
        if self.team_a_active.get():
            self.txt_team_a.config(state='normal', background='white')
        else:
            self.txt_team_a.config(state='disabled', background='#f0f0f0')
            
        if self.team_b_active.get():
            self.txt_team_b.config(state='normal', background='white')
        else:
            self.txt_team_b.config(state='disabled', background='#f0f0f0')

    def parse_float(self, val_str):
        try:
            return float(val_str.replace(',', '').strip())
        except ValueError:
            return 0.0

    def generate_files(self, silent_path=None):
        # Targets
        TARGET_RT = 24536
        TARGET_UT = 319.02
        TARGET_PT = 338.63
        
        rt_prev = self.parse_float(self.ent_rt_prev.get())
        rt_today = self.parse_float(self.ent_rt_today.get())
        ut_prev = self.parse_float(self.ent_ut_prev.get())
        ut_today = self.parse_float(self.ent_ut_today.get())
        pt_prev = self.parse_float(self.ent_pt_prev.get())
        pt_today = self.parse_float(self.ent_pt_today.get())
        
        rt_prog = ((rt_prev + rt_today) / TARGET_RT) * 100 if TARGET_RT > 0 else 0
        ut_prog = ((ut_prev + ut_today) / TARGET_UT) * 100 if TARGET_UT > 0 else 0
        pt_prog = ((pt_prev + pt_today) / TARGET_PT) * 100 if TARGET_PT > 0 else 0
        
        def fmt(val, unit, is_int=False):
            if is_int:
                return f"{int(val):,} {unit}"
            else:
                return f"{val:,.2f} {unit}"
                
        str_prev = f"RT: {fmt(rt_prev, '매', True)}\nUT: {fmt(ut_prev, 'M')}\nPT: {fmt(pt_prev, 'M')}"
        str_today = f"RT: {fmt(rt_today, '매', True)}\nUT: {fmt(ut_today, 'M')}\nPT: {fmt(pt_today, 'M')}"
        str_prog = f"RT: {rt_prog:.1f}% / {fmt(TARGET_RT, '매', True)}\nUT: {ut_prog:.1f}% / {fmt(TARGET_UT, 'M')}\nPT: {pt_prog:.1f}% / {fmt(TARGET_PT, 'M')}"

        params = {
            'date': self.ent_date.get().strip(),
            'company': self.ent_company.get().strip(),
            'locations': self.ent_locations.get().strip(),
            'personnel': self.ent_personnel.get().strip(),
            'equipment': self.ent_equipment.get().strip(),
            'rt': self.ent_rt.get().strip(),
            'etc': self.ent_etc.get().strip(),
            'team_a': self.txt_team_a.get('1.0', tk.END).strip(),
            'team_b': self.txt_team_b.get('1.0', tk.END).strip(),
            'team_a_loc': self.ent_team_a_loc.get().strip(),
            'team_b_loc': self.ent_team_b_loc.get().strip(),
            'team_a_active': self.team_a_active.get(),
            'team_b_active': self.team_b_active.get(),
            'team_a_rt': self.team_a_rt.get(),
            'team_a_ut': self.team_a_ut.get(),
            'team_a_pt': self.team_a_pt.get(),
            'team_b_rt': self.team_b_rt.get(),
            'team_b_ut': self.team_b_ut.get(),
            'team_b_pt': self.team_b_pt.get(),
            'prev': str_prev,
            'today': str_today,
            'prog': str_prog,
            'req': self.txt_req.get('1.0', tk.END).strip(),
        }
        
        if silent_path:
            output_path = silent_path
        else:
            initial_dir = os.path.dirname(os.path.abspath(__file__))
            output_dir = filedialog.askdirectory(title="저장할 폴더를 선택하세요", initialdir=initial_dir)
            if not output_dir:
                return
            output_path = os.path.join(output_dir, "작업승인계획서_NDT전용.xlsx")
        
        self.save_config()
        self.btn_generate.config(state='disabled')
        self.lbl_status.config(text="엑셀 파일 생성 중...", foreground="blue")
        self.root.update()
        
        try:
            self.create_excel(output_path, params)
            if not silent_path:
                messagebox.showinfo("생성 완료", f"작업승인계획서가 성공적으로 생성되었습니다!\n\n저장 위치:\n{output_path}")
                self.lbl_status.config(text="완료!", foreground="green")
            else:
                self.lbl_status.config(text="임시 파일 생성 완료!", foreground="green")
        except Exception as e:
            if not silent_path:
                messagebox.showerror("오류", f"엑셀 파일 생성 중 오류가 발생했습니다:\n{e}")
            self.lbl_status.config(text="오류 발생", foreground="red")
        finally:
            self.btn_generate.config(state='normal')

    def generate_unified_excel(self):
        initial_dir = os.path.dirname(os.path.abspath(__file__))
        output_dir = filedialog.askdirectory(title="일일 안전서류 통합 엑셀 저장 폴더를 선택하세요", initialdir=initial_dir)
        if not output_dir: return
        
        date_str = self.ent_date.get().replace("-", "")
        final_path = os.path.join(output_dir, f"일일_안전서류_통합_{date_str}.xlsx").replace("/", "\\")
        
        self.lbl_status.config(text="통합 엑셀 생성 중...", foreground="blue")
        self.btn_generate.config(state='disabled')
        self.btn_unified.config(state='disabled')
        self.root.update()
        
        temp_approval = os.path.join(output_dir, f"temp_approval_{date_str}.xlsx").replace("/", "\\")
        temp_tbm = os.path.join(output_dir, f"temp_tbm_{date_str}.xlsx").replace("/", "\\")
        
        excel = None
        try:
            # 0. 날짜 및 UI 동기화 (메인 탭 기준)
            main_date = self.ent_date.get().strip()
            
            if hasattr(self, 'tbm_manager'):
                self.tbm_manager.ent_date.delete(0, tk.END)
                self.tbm_manager.ent_date.insert(0, main_date)
                self.tbm_manager.update_work_and_hazards()
                
            if hasattr(self, 'risk_manager'):
                try:
                    import datetime
                    dt = datetime.datetime.strptime(main_date, "%Y-%m-%d")
                    risk_date = dt.strftime("%Y년 %m월 %d일")
                    self.risk_manager.ent_write_date.delete(0, tk.END)
                    self.risk_manager.ent_write_date.insert(0, risk_date)
                except:
                    pass

            # 1. 작업승인계획서 생성
            self.generate_files(silent_path=temp_approval)
            
            # 2. TBM 생성
            if hasattr(self, 'tbm_manager'):
                self.tbm_manager.export_excel(silent_path=temp_tbm)
                
            # 3. 위험성평가표 생성
            temp_risks = []
            if hasattr(self, 'risk_manager'):
                temp_risks = self.risk_manager.generate_files(silent_dir=output_dir, date_str=date_str) or []
                
            # 합치기
            import win32com.client as win32
            excel = win32.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            # 새 통합 워크북 생성
            wb_master = excel.Workbooks.Add()
            default_sheet = wb_master.Sheets(1)
            
            # 작업승인계획서 복사 (Before=default_sheet)
            if os.path.exists(temp_approval):
                wb1 = excel.Workbooks.Open(temp_approval)
                wb1.Sheets(1).Copy(wb_master.Sheets(1))
                wb1.Close(False)
                try: os.remove(temp_approval)
                except: pass
                
            # TBM 복사 (After=첫번째 시트)
            if os.path.exists(temp_tbm):
                wb2 = excel.Workbooks.Open(temp_tbm)
                wb2.Sheets(1).Copy(None, wb_master.Sheets(1))
                wb2.Close(False)
                try: os.remove(temp_tbm)
                except: pass
                
            # 위험성평가표 복사 (After=마지막 시트)
            for temp_risk in temp_risks:
                if os.path.exists(temp_risk):
                    last_sheet = wb_master.Sheets(wb_master.Sheets.Count)
                    wb3 = excel.Workbooks.Open(temp_risk)
                    wb3.Sheets(1).Copy(None, last_sheet)
                    wb3.Close(False)
                    try: os.remove(temp_risk)
                    except: pass
                    
            # 기본 Sheet1 삭제
            default_sheet.Delete()
            
            wb_master.SaveAs(final_path)
            wb_master.Close(False)
            excel.Quit()
            
            messagebox.showinfo("통합 완료", f"일일 안전서류 통합 엑셀이 성공적으로 생성되었습니다!\n\n저장 위치:\n{final_path}")
            self.lbl_status.config(text="통합 완료!", foreground="green")
            os.startfile(final_path)
            
        except Exception as e:
            messagebox.showerror("오류", f"통합 엑셀 생성 중 오류가 발생했습니다:\n{e}")
            self.lbl_status.config(text="통합 오류", foreground="red")
            try:
                if excel: excel.Quit()
            except: pass
        finally:
            self.btn_generate.config(state='normal')
            self.btn_unified.config(state='normal')

    def create_excel(self, output_path, params):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "작업승인계획서"

        bold_font = Font(bold=True)
        title_font = Font(bold=True, size=16)
        center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
        header_fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                             top=Side(style='thin'), bottom=Side(style='thin'))

        def set_border(ws, min_col, min_row, max_col, max_row):
            for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
                for cell in row:
                    cell.border = thin_border

        ws.column_dimensions['A'].width = 20
        ws.column_dimensions['B'].width = 50
        ws.column_dimensions['C'].width = 40
        ws.column_dimensions['D'].width = 40
        ws.column_dimensions['E'].width = 25

        ws.merge_cells('A1:E1')
        ws['A1'] = f"[서울안전건설사무소] {params['company']} 비파괴검사 주요 일일작업"
        ws['A1'].font = title_font
        ws['A1'].alignment = center_align
        ws.row_dimensions[1].height = 45

        ws.merge_cells('A3:C3')
        ws['A3'] = f"일자: {params['date']}"
        ws['A3'].font = bold_font

        ws['D2'] = "수급업체"
        ws['E2'] = "KOGAS"
        ws['D3'] = "(인)"
        ws['E3'] = "(인)"

        for r in range(2, 4):
            for c in range(4, 6):
                ws.cell(row=r, column=c).alignment = center_align
                ws.cell(row=r, column=c).border = thin_border
                ws.cell(row=r, column=c).font = bold_font
        ws['D2'].fill = header_fill
        ws['E2'].fill = header_fill

        ws.merge_cells('A5:E5')
        ws['A5'] = "1. 총 투입 현황"
        ws['A5'].font = bold_font

        headers_sec1 = ["총 작업 개소", "인원 (계)", "장비 (계)", "RT / 크롤러 투입", "UT / PT / 기타 장비"]
        for col, val in enumerate(headers_sec1, start=1):
            cell = ws.cell(row=6, column=col, value=val)
            cell.font = bold_font
            cell.alignment = center_align
            cell.fill = header_fill

        values_sec1 = [params['locations'], params['personnel'], params['equipment'], params['rt'], params['etc']]
        for col, val in enumerate(values_sec1, start=1):
            cell = ws.cell(row=7, column=col, value=val)
            cell.alignment = center_align

        set_border(ws, 1, 6, 5, 7)

        ws.merge_cells('A9:E9')
        ws['A9'] = "2. 팀별 세부 작업 및 안전관리 계획"
        ws['A9'].font = bold_font

        headers_sec2 = ["구 분", "금 일 작 업 (내용 및 시간)", "주요 위험 요소 (위험성 평가)", "안전관리 중점사항 (대책)", "시공자\n(관리감독자)"]
        for col, val in enumerate(headers_sec2, start=1):
            cell = ws.cell(row=10, column=col, value=val)
            cell.font = bold_font
            cell.alignment = center_align
            cell.fill = header_fill

        active_teams = []
        
        import hashlib
        
        rt_haz = [
            ("(방사선 피폭) 방사선 투과검사 중 피폭", "콜리메이터 사용, 통제구역 설정/감시자 배치"), 
            ("(추락) 지상 2m 이상 배관 위 검사", "고소작업 시 2인 1조 필수, 안전대 체결"),
            ("(질식) 배관 내부 진입 시 산소 결핍", "산소농도 측정 및 환기 실시, 밀폐공간 진입 통제"),
            ("(협착) 크롤러 등 장비 이동 중 끼임", "장비 이동 시 주변 확인, 작업 지휘자 배치"),
            ("(근골격계) 무거운 납 차폐체/장비 운반", "스트레칭 실시, 중량물 2인 이상 운반")
        ]
        ut_haz = [
            ("(추락) 고소 배관 용접부 UT 탐상", "안전대 체결, 비계 발판 상태 사전 점검"),
            ("(충돌) 좁은 공간 내 타 공정 장비 충돌", "안전감독관 사전 조율 후 작업 통제"),
            ("(근골격계) 부자연스러운 자세로 장시간 탐상", "주기적인 휴식 및 스트레칭 실시"),
            ("(전도) 현장 내 자재/공구에 걸려 넘어짐", "작업장 주변 정리정돈 철저, 조도 확보")
        ]
        pt_haz = [
            ("(화학물질) PT 용제 취급 시 흡입/피부접촉", "MSDS 비치 및 방독마스크, 장갑 착용"),
            ("(화재) 가연성 세척액 사용으로 인한 화재", "화기 구역 분리, 소화기 비치"),
            ("(밀폐공간) 환기 불량 구역 PT 검사 시 질식", "국소배기장치 가동, 작업 중 주기적 환기"),
            ("(근골격계) 바닥면 배관 쪼그려 앉아 검사", "적절한 휴식시간 부여, 스트레칭 유도")
        ]

        def build_team_hazards(date_str, is_rt, is_ut, is_pt):
            hash_val = int(hashlib.md5(date_str.encode('utf-8')).hexdigest(), 16)
            def pick_hazard(hz_list, offset=0):
                return hz_list[(hash_val + offset) % len(hz_list)]
                
            selected_methods = []
            if is_rt: selected_methods.append('RT')
            if is_ut: selected_methods.append('UT')
            if is_pt: selected_methods.append('PT')
            
            picked = []
            if len(selected_methods) == 0:
                picked = [pick_hazard(rt_haz, 0), pick_hazard(rt_haz, 1), pick_hazard(rt_haz, 2)]
            elif len(selected_methods) == 1:
                m1 = selected_methods[0]
                list1 = rt_haz if m1 == 'RT' else (ut_haz if m1 == 'UT' else pt_haz)
                picked = [pick_hazard(list1, 0), pick_hazard(list1, 1), pick_hazard(list1, 2)]
            elif len(selected_methods) == 2:
                m1, m2 = selected_methods[0], selected_methods[1]
                list1 = rt_haz if m1 == 'RT' else (ut_haz if m1 == 'UT' else pt_haz)
                list2 = rt_haz if m2 == 'RT' else (ut_haz if m2 == 'UT' else pt_haz)
                picked = [pick_hazard(list1, 0), pick_hazard(list2, 0), pick_hazard(list1, 1)]
            elif len(selected_methods) >= 3:
                picked = [pick_hazard(rt_haz, 0), pick_hazard(ut_haz, 0), pick_hazard(pt_haz, 0)]
                
            c_text = "\n".join([f"{idx+1}. {item[0]}" for idx, item in enumerate(picked)])
            d_text = "\n".join([f"{idx+1}. {item[1]}" for idx, item in enumerate(picked)])
            return c_text, d_text

        if params.get('team_a_active'):
            loc_a = params.get('team_a_loc', '00')
            c_a, d_a = build_team_hazards(params['date'], params.get('team_a_rt'), params.get('team_a_ut'), params.get('team_a_pt'))
            active_teams.append({
                'A': f"비파괴 A팀 (본관)\n\n(작업개소: {loc_a}개소)",
                'B': params['team_a'],
                'C': c_a,
                'D': d_a,
                'E': "(서명)"
            })
        if params.get('team_b_active'):
            loc_b = params.get('team_b_loc', '00')
            c_b, d_b = build_team_hazards(params['date'], params.get('team_b_rt'), params.get('team_b_ut'), params.get('team_b_pt'))
            active_teams.append({
                'A': f"비파괴 B팀 (관리소)\n\n(작업개소: {loc_b}개소)",
                'B': params['team_b'],
                'C': c_b,
                'D': d_b,
                'E': "(서명)"
            })
        if len(active_teams) == 1:
            active_teams[0]['A'] = active_teams[0]['A'].replace('B팀', 'A팀')

        for i in range(2):
            row_num = 11 + i
            if i < len(active_teams):
                team = active_teams[i]
                ws[f'A{row_num}'] = team['A']
                ws[f'B{row_num}'] = team['B']
                ws[f'C{row_num}'] = team['C']
                ws[f'D{row_num}'] = team['D']
                ws[f'E{row_num}'] = team['E']
                
                ws[f'A{row_num}'].alignment = center_align
                ws[f'B{row_num}'].alignment = left_align
                ws[f'C{row_num}'].alignment = left_align
                ws[f'D{row_num}'].alignment = left_align
                ws[f'E{row_num}'].alignment = center_align
            else:
                ws[f'A{row_num}'] = ""
                ws[f'B{row_num}'] = ""
                ws[f'C{row_num}'] = "내용 없음"
                ws[f'D{row_num}'] = ""
                ws[f'E{row_num}'] = ""
                
                ws[f'A{row_num}'].alignment = center_align
                ws[f'B{row_num}'].alignment = left_align
                ws[f'C{row_num}'].alignment = center_align
                ws[f'D{row_num}'].alignment = left_align
                ws[f'E{row_num}'].alignment = center_align

        ws.row_dimensions[11].height = 130
        ws.row_dimensions[12].height = 120
        ws.row_dimensions[10].height = 30

        set_border(ws, 1, 10, 5, 12)

        ws.merge_cells('A14:E14')
        ws['A14'] = "3. 기타 진행 현황 및 요청사항"
        ws['A14'].font = bold_font

        headers_sec3 = ["작업 구간", "전일 누계", "금일 계획", "전체 진행률", "기타 작업 현황 및 요청사항"]
        for col, val in enumerate(headers_sec3, start=1):
            cell = ws.cell(row=15, column=col, value=val)
            cell.font = bold_font
            cell.alignment = center_align
            cell.fill = header_fill

        values_sec3 = ["전체 공구", params['prev'], params['today'], params['prog'], params['req']]
        for col, val in enumerate(values_sec3, start=1):
            cell = ws.cell(row=16, column=col, value=val)
            if col == 5:
                cell.alignment = left_align
            else:
                cell.alignment = center_align

        # Set row height to 75 to fit 3 lines of text comfortably
        ws.row_dimensions[16].height = 75
        set_border(ws, 1, 15, 5, 16)

        ws.merge_cells('A18:E18')
        ws['A18'] = "※ 참고: 비파괴검사는 '위험작업'에 속하므로 본 계획서와 함께 [위험성평가표]를 첨부하여 작업 1일 전(D-1)까지 승인을 득해야 합니다."
        ws['A18'].font = Font(bold=True, color="FF0000")

        # 가로 방향 인쇄(Landscape) 및 한 페이지에 모두 맞춤 설정
        ws.set_printer_settings(paper_size=9, orientation='landscape')
        ws.print_options.horizontalCentered = True
        ws.print_options.verticalCentered = True
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1

        wb.save(output_path)

if __name__ == "__main__":
    root = tk.Tk()
    app = WorkApprovalApp(root)
    root.mainloop()
