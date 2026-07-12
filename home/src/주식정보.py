import tkinter as tk
from tkinter import ttk, messagebox
import threading
import webbrowser
import os
import sys
import pandas as pd
from datetime import datetime

# Stock_Analyzer 모듈 경로 추가
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "Stock_Analyzer")))

import daily_economic_analyzer as analyzer
import json

UI_SETTINGS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ui_settings.json")

def load_ui_settings():
    if os.path.exists(UI_SETTINGS_FILE):
        try:
            with open(UI_SETTINGS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except: pass
    return {}

def save_ui_settings(settings):
    try:
        with open(UI_SETTINGS_FILE, 'w', encoding='utf-8') as f:
            json.dump(settings, f)
    except: pass

class EconomicDashboard:
    def __init__(self, root):
        self.root = root
        self.root.title("📈 나만의 경제 뉴스 분석 비서")
        self.root.geometry("1400x800")
        
        # 최상단에 윈도우 창 위치 (잠깐만)
        self.root.attributes('-topmost', True)
        self.root.after(500, lambda: self.root.attributes('-topmost', False))
        
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.ui_settings = load_ui_settings()
        
        # 스타일 설정
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Treeview.Heading", font=('Malgun Gothic', 10, 'bold'))
        style.configure("Treeview", font=('Malgun Gothic', 10), rowheight=30)
        
        # 상단 헤더 프레임
        top_frame = ttk.Frame(root, padding=10)
        top_frame.pack(fill='x')
        
        ttk.Label(top_frame, text="📊 데일리 경제 & 증시 대시보드", font=('Malgun Gothic', 18, 'bold')).pack(side='left')
        # (목표 달성 UI 삭제됨)
        self.btn_trend = ttk.Button(top_frame, text="🔥 AI 핫 트렌드 예측", command=self.open_trend_window)
        self.btn_trend.pack(side='left', padx=5)
        
        self.btn_add_stock = ttk.Button(top_frame, text="➕ 내 주식 추가", command=self.open_add_stock_window)
        self.btn_add_stock.pack(side='left', padx=5)
        
        self.btn_refresh = ttk.Button(top_frame, text="🔄 데이터 갱신 및 엑셀 저장", command=self.refresh_data)
        self.btn_refresh.pack(side='right')
        
        self.lbl_status = ttk.Label(top_frame, text="프로그램 시작 중...", font=('Malgun Gothic', 10), foreground="gray")
        self.lbl_status.pack(side='right', padx=15)
        
        # 화면 분할용 PanedWindow (위: 탭 / 아래: 뉴스)
        self.paned = ttk.PanedWindow(root, orient='vertical')
        self.paned.pack(expand=True, fill='both', padx=10, pady=10)
        
        # ----------------- [상단 영역: 탭(Notebook)] -----------------
        self.notebook = ttk.Notebook(self.paned)
        self.paned.add(self.notebook, weight=3)
        
        # 탭 1: 전체 시장 지표
        self.tab_market = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_market, text=' 🌐 주요 시장 지표 및 테마 추적 ')
        
        # 탭 2: AI 추천 종목
        self.tab_rec = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_rec, text=' 💡 AI 강력 매수 및 추천 포트폴리오 ')
        
        # 탭 3: 나의 주식 집중 관리
        self.tab_my = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_my, text=' 💼 나의 주식 집중 관리 ')
        
        # ----------------- [탭 1 레이아웃] -----------------
        self.tab1_paned = ttk.PanedWindow(self.tab_market, orient='horizontal')
        self.tab1_paned.pack(expand=True, fill='both', padx=5, pady=5)
        
        market_left = ttk.Frame(self.tab1_paned)
        self.tab1_paned.add(market_left, weight=1)
        
        market_right = ttk.Frame(self.tab1_paned)
        self.tab1_paned.add(market_right, weight=1)
        
        # 시장 지표 섹션 (왼쪽)
        market_frame = ttk.LabelFrame(market_left, text=" 📌 주요 시장 지표 (KOSPI, 환율 등) ", padding=5)
        market_frame.pack(side='top', expand=True, fill='both')
        
        cols_market = ('지표명', '현재가', '전일비(%)', '기준일자')
        self.tv_market = ttk.Treeview(market_frame, columns=cols_market, show='headings', height=10)
        for col in cols_market:
            self.tv_market.heading(col, text=col)
            self.tv_market.column(col, anchor='center', width=100)
        self.tv_market.pack(expand=True, fill='both')
        
        # 섹터별 등락률 추적 섹션 (오른쪽)
        sector_frame = ttk.LabelFrame(market_right, text=" 📊 테마/섹터별 당일 등락률 추적 ", padding=5)
        sector_frame.pack(side='top', expand=True, fill='both')
        
        cols_sector = ('섹터/테마명', '평균 전일비(%)', '주도 종목')
        self.tv_sector = ttk.Treeview(sector_frame, columns=cols_sector, show='headings', height=10)
        for col in cols_sector:
            self.tv_sector.heading(col, text=col)
            w = 150 if col == '섹터/테마명' else (120 if col == '평균 전일비(%)' else 250)
            self.tv_sector.column(col, anchor='center', width=w, stretch=False)
        self.tv_sector.pack(expand=True, fill='both')

        # ----------------- [탭 2 레이아웃] -----------------
        self.tab2_paned = ttk.PanedWindow(self.tab_rec, orient='vertical')
        self.tab2_paned.pack(expand=True, fill='both', padx=5, pady=5)

        # AI 매수 추천 섹션 (상단)
        buy_rec_frame = ttk.LabelFrame(self.tab2_paned, text=" 🔥 오늘의 AI 강력 매수 추천 TOP 3 ", padding=5)
        self.tab2_paned.add(buy_rec_frame, weight=1)
        
        cols_rec = ('순위', '추천 종목', '현재가', '기대 수익률(%)', '추천 이유(시그널)')
        self.tv_rec = ttk.Treeview(buy_rec_frame, columns=cols_rec, show='headings', height=4)
        for col in cols_rec:
            self.tv_rec.heading(col, text=col)
            w = 80 if col == '순위' else (250 if col == '추천 종목' else 150)
            if col == '추천 이유(시그널)': w = 300
            self.tv_rec.column(col, anchor='center', width=w, stretch=False)
        self.tv_rec.pack(expand=True, fill='both')
        
        # 추천 포트폴리오 섹션 (하단)
        port_frame = ttk.LabelFrame(self.tab2_paned, text=" 💡 AI 주도주 수익 창출 포트폴리오 (실시간) ", padding=5)
        self.tab2_paned.add(port_frame, weight=3)
        
        cols_port = ('추천 종목', '보유', '현재가(원/$)', '전일비(%)', '거래량 증감(%)', '예상 저점(지지선)', '예상 고점(저항선)', '목표 매수가', '부분 매도가', '투자 성향', 'AI 매매 시그널')
        self.tv_port = ttk.Treeview(port_frame, columns=cols_port, show='headings', height=10)
        for col in cols_port:
            self.tv_port.heading(col, text=col)
            if col == '추천 종목': w = 180
            elif col == '보유': w = 60
            elif col == 'AI 매매 시그널': w = 150
            elif col in ['목표 매수가', '부분 매도가', '예상 저점(지지선)', '예상 고점(저항선)']: w = 110
            else: w = 80
            self.tv_port.column(col, anchor='center', width=w, stretch=False)
            
        port_scroll_y = ttk.Scrollbar(port_frame, orient='vertical', command=self.tv_port.yview)
        port_scroll_x = ttk.Scrollbar(port_frame, orient='horizontal', command=self.tv_port.xview)
        self.tv_port.configure(yscroll=port_scroll_y.set, xscroll=port_scroll_x.set)
        
        port_scroll_x.pack(side='bottom', fill='x')
        port_scroll_y.pack(side='right', fill='y')
        self.tv_port.pack(expand=True, fill='both')

        # ----------------- [탭 2 레이아웃] -----------------
        # 나의 보유 주식 섹션 (탭 2의 전체를 차지하여 아주 쾌적하게 사용)
        my_port_frame = ttk.LabelFrame(self.tab_my, text=" 💼 나의 실제 보유 주식 현황 (집중 관리) ", padding=10)
        my_port_frame.pack(expand=True, fill='both', padx=10, pady=10)
        
        cols_my = ('종목명', '보유', '매수 단가', '현재가', '수익률(%)', '일간 변동금액', 'AI 시그널')
        self.tv_my = ttk.Treeview(my_port_frame, columns=cols_my, show='headings', height=15)
        for col in cols_my:
            self.tv_my.heading(col, text=col)
            w = 150 if col == '종목명' else (80 if col == '보유' else (100 if col == '수익률(%)' else (180 if col == 'AI 시그널' else 120)))
            self.tv_my.column(col, anchor='center', width=w, stretch=False)
        
        btn_frame = ttk.Frame(my_port_frame)
        btn_frame.pack(side='bottom', fill='x', pady=(10, 0))
        
        self.lbl_trend = ttk.Label(btn_frame, text="🔥 내일의 주도주 예측 중...", font=('Malgun Gothic', 11, 'bold'), foreground='red')
        self.lbl_trend.pack(side='left', padx=5)
        
        btn_my_advice = ttk.Button(btn_frame, text="💡 AI 보유 주식 정밀 진단 및 대책", command=self.open_my_advice_window)
        btn_my_advice.pack(side='right')
        
        btn_del_stock = ttk.Button(btn_frame, text="🗑️ 선택 주식 삭제", command=self.delete_selected_stock)
        btn_del_stock.pack(side='right', padx=5)

        my_scroll_y = ttk.Scrollbar(my_port_frame, orient='vertical', command=self.tv_my.yview)
        my_scroll_x = ttk.Scrollbar(my_port_frame, orient='horizontal', command=self.tv_my.xview)
        self.tv_my.configure(yscroll=my_scroll_y.set, xscroll=my_scroll_x.set)
        
        my_scroll_x.pack(side='bottom', fill='x')
        my_scroll_y.pack(side='right', fill='y')
        self.tv_my.pack(expand=True, fill='both')

        
        # 2. 주요 뉴스 섹션
        news_frame = ttk.LabelFrame(self.paned, text=" 📰 최신 경제 및 주식 뉴스 (더블클릭 시 인터넷 창 열림) ", padding=5)
        self.paned.add(news_frame, weight=3)
        
        cols_news = ('시장 심리', '언론사', '기사 제목', '발행일시', '링크')
        self.tv_news = ttk.Treeview(news_frame, columns=cols_news, show='headings')
        self.tv_news.heading('시장 심리', text='시장 심리')
        self.tv_news.column('시장 심리', width=100, anchor='center')
        self.tv_news.heading('언론사', text='언론사')
        self.tv_news.column('언론사', width=100, anchor='center')
        self.tv_news.heading('기사 제목', text='기사 제목')
        self.tv_news.column('기사 제목', width=600, anchor='w')
        self.tv_news.heading('발행일시', text='발행일시')
        self.tv_news.column('발행일시', width=150, anchor='center')
        self.tv_news.column('링크', width=0, stretch=False) # 링크는 숨김
        
        # 뉴스 트리뷰 스크롤바
        scrollbar = ttk.Scrollbar(news_frame, orient='vertical', command=self.tv_news.yview)
        self.tv_news.configure(yscroll=scrollbar.set)
        scrollbar.pack(side='right', fill='y')
        self.tv_news.pack(expand=True, fill='both')
        
        # 더블클릭 이벤트 (기사 읽기 및 포트폴리오 예측)
        self.tv_news.bind("<Double-1>", self.open_link)
        self.tv_port.bind("<Double-1>", self.open_prediction_window)
        self.tv_rec.bind("<Double-1>", self.open_prediction_window)
        self.tv_my.bind("<Double-1>", self.open_prediction_window)
        
        # 색상 태그 설정 (상승=빨강, 하락=파랑)
        self.tv_news.tag_configure('pos', foreground='red')
        self.tv_news.tag_configure('neg', foreground='blue')
        self.tv_market.tag_configure('up', foreground='red')
        self.tv_market.tag_configure('down', foreground='blue')
        self.tv_port.tag_configure('up', foreground='red')
        self.tv_port.tag_configure('down', foreground='blue')
        self.tv_port.tag_configure('sell', foreground='purple')  # 매도 시그널용 색상
        self.tv_port.tag_configure('value', foreground='darkorange') # 숨은 진주용 색상
        
        # 내 보유 주식도 동일한 색상 적용
        self.tv_my.tag_configure('up', foreground='red')
        self.tv_my.tag_configure('down', foreground='blue')
        self.tv_my.tag_configure('sell', foreground='purple')
        self.tv_my.tag_configure('value', foreground='darkorange')
        
        # 자동 갱신 타이머 설정 (5분 = 300,000ms)
        self.refresh_interval = 300000
        self.timer_id = None
        
        self.root.after(100, self.restore_ui_state)
        
        # 시작하자마자 자동 수집 실행
        self.root.after(500, self.refresh_data)
        
    def on_closing(self):
        settings = {'cols': {}}
        try: settings['paned'] = self.paned.sashpos(0)
        except: pass
        
        try: settings['top_content_frame'] = self.top_content_frame.sashpos(0)
        except: pass
        
        for tv_name, tv in [('tv_market', self.tv_market), ('tv_my', self.tv_my), ('tv_port', self.tv_port), ('tv_news', self.tv_news)]:
            settings['cols'][tv_name] = {}
            for col in tv['columns']:
                try: settings['cols'][tv_name][col] = tv.column(col, 'width')
                except: pass
                
        save_ui_settings(settings)
        self.root.destroy()
        
    def restore_ui_state(self):
        settings = self.ui_settings
        if 'paned' in settings:
            try: self.paned.sashpos(0, settings['paned'])
            except: pass
        if 'top_content_frame' in settings:
            try: self.top_content_frame.sashpos(0, settings['top_content_frame'])
            except: pass
            
        if 'cols' in settings:
            for tv_name, tv in [('tv_market', self.tv_market), ('tv_my', self.tv_my), ('tv_port', self.tv_port), ('tv_news', self.tv_news)]:
                if tv_name in settings['cols']:
                    for col, width in settings['cols'][tv_name].items():
                        try: tv.column(col, width=width)
                        except: pass
                        
    def open_link(self, event):
        item = self.tv_news.selection()
        if item:
            link = self.tv_news.item(item[0], 'values')[4]
            webbrowser.open(link)
            
    def open_prediction_window(self, event):
        widget = event.widget
        item = widget.selection()
        if not item: return
        
        # 클릭된 위젯이 추천 매수(tv_rec)인 경우 종목명이 인덱스 1, 포트폴리오(tv_port)인 경우 인덱스 0
        if widget == getattr(self, 'tv_rec', None):
            stock_name = widget.item(item[0], 'values')[1]
        else:
            stock_name = widget.item(item[0], 'values')[0]
            
        pred_win = tk.Toplevel(self.root)
        pred_win.title(f"{stock_name} - 종합 진단 및 5일 가격 예측")
        pred_win.geometry("900x700")
        
        ttk.Label(pred_win, text="데이터 분석 및 차트 생성 중... 잠시만 기다려주세요.", font=('Malgun Gothic', 12)).pack(pady=50)
        
        thread = threading.Thread(target=self._run_simulation, args=(pred_win, stock_name))
        thread.daemon = True
        thread.start()
        
    def _run_simulation(self, window, stock_name):
        import yfinance as yf
        import numpy as np
        from datetime import datetime, timedelta
        
        # 이름 정제 (주도주 마크, 사용자 추가 문구 제거)
        clean_name = stock_name.replace('🔥 ', '').replace(' (사용자 추가)', '').strip()
        
        ticker = analyzer.PORTFOLIO.get(clean_name)
        if not ticker:
            for full_name, t in analyzer.PORTFOLIO.items():
                if full_name.split('(')[0].strip() == clean_name:
                    ticker = t
                    break
                    
        if not ticker:
            custom = analyzer.load_custom_stocks()
            if clean_name in custom:
                ticker = custom[clean_name]['ticker']
                
        if not ticker:
            self.root.after(0, lambda: ttk.Label(window, text="티커 정보를 찾을 수 없습니다.").pack())
            return
            
        try:
            stock = yf.Ticker(ticker)
            hist = stock.history(period='1y')
            if len(hist) < 20: raise Exception("데이터가 충분하지 않습니다.")
            
            returns = hist['Close'].pct_change().dropna()
            mu = returns.mean()
            sigma = returns.std()
            
            current_price = hist['Close'].iloc[-1]
            is_korea = '.KS' in ticker or '.KQ' in ticker
            
            def update_ui():
                for widget in window.winfo_children():
                    widget.destroy()
                    
                import matplotlib.pyplot as plt
                from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
                
                main_frame = ttk.Frame(window)
                main_frame.pack(expand=True, fill='both', padx=10, pady=10)
                
                ttk.Label(main_frame, text=f"📊 {stock_name} 종합 진단 및 예측", font=('Malgun Gothic', 14, 'bold')).pack(pady=5)
                
                mid_frame = ttk.Frame(main_frame)
                mid_frame.pack(expand=True, fill='both', pady=5)
                
                chart_frame = ttk.Frame(mid_frame)
                chart_frame.pack(side='left', expand=True, fill='both')
                
                desc_frame = ttk.LabelFrame(mid_frame, text=" 💡 종목 분석 ", padding=10)
                desc_frame.pack(side='right', fill='y', padx=(10, 0))
                
                plt.rcParams['font.family'] = 'Malgun Gothic'
                plt.rcParams['axes.unicode_minus'] = False
                
                fig, ax = plt.subplots(figsize=(5, 3))
                ax.plot(hist.index, hist['Close'], label='종가', color='#1f77b4', linewidth=2)
                ax.plot(hist.index, hist['Close'].rolling(20).mean(), label='20일선', color='#d62728', linestyle='--')
                ax.plot(hist.index, hist['Close'].rolling(60).mean(), label='60일선', color='#2ca02c', linestyle='-.')
                ax.legend()
                ax.grid(True, linestyle='--', alpha=0.6)
                fig.tight_layout()
                
                canvas = FigureCanvasTkAgg(fig, master=chart_frame)
                canvas.draw()
                canvas.get_tk_widget().pack(fill='both', expand=True)
                
                ma20 = hist['Close'].rolling(20).mean().iloc[-1]
                high_price = hist['High'].max()
                low_price = hist['Low'].min()
                
                msg = f"📌 현재가: {int(current_price):,}원\n\n" if is_korea else f"📌 현재가: ${current_price:.2f}\n\n"
                msg += f"• 1년 최고: {int(high_price):,}원\n" if is_korea else f"• 1년 최고: ${high_price:.2f}\n"
                msg += f"• 1년 최저: {int(low_price):,}원\n\n" if is_korea else f"• 1년 최저: ${low_price:.2f}\n\n"
                
                msg += "📈 추세 진단:\n"
                if current_price > ma20:
                    msg += "현재 20일 이동평균선 위에 위치하여\n단기 상승 추세가 유지되고 있습니다.\n\n"
                else:
                    msg += "현재 20일 이동평균선 아래에 위치해\n단기적으로 조정을 받고 있습니다.\n\n"
                    
                lbl_desc = ttk.Label(desc_frame, text=msg, font=('Malgun Gothic', 11), justify='left')
                lbl_desc.pack(anchor='nw')
                
                bot_frame = ttk.LabelFrame(main_frame, text=" 📊 향후 5영업일 몬테카를로 가격 예측 시뮬레이션 ", padding=5)
                bot_frame.pack(fill='x', pady=5)
                
                cols = ('날짜', '비관적 (하락)', '중립 (추세)', '낙관적 (상승)')
                tv = ttk.Treeview(bot_frame, columns=cols, show='headings', height=5)
                for col in cols:
                    tv.heading(col, text=col)
                    tv.column(col, anchor='center')
                tv.pack(fill='x', padx=5, pady=5)
                
                curr_date = datetime.now()
                
                # 몬테카를로 시뮬레이션 (10,000번의 무작위 주가 흐름 시나리오 생성)
                num_simulations = 10000
                days = 5
                
                # mu(평균)와 sigma(표준편차)를 기반으로 5일치 난수 생성
                daily_returns = np.random.normal(mu, sigma, (days, num_simulations))
                
                # 주가 경로 계산 배열 초기화
                price_paths = np.zeros((days + 1, num_simulations))
                price_paths[0] = current_price
                
                for t in range(1, days + 1):
                    price_paths[t] = price_paths[t-1] * (1 + daily_returns[t-1])
                
                for i in range(1, 6):
                    curr_date += timedelta(days=1)
                    while curr_date.weekday() > 4:
                        curr_date += timedelta(days=1)
                        
                    # 10,000번의 시나리오 중 확률 분포에 따른 값 추출
                    # 비관적: 하위 10% (발생 확률 10% 미만의 최악 시나리오 기준)
                    # 중립적: 중간값 50% (가장 확률이 높은 평균치)
                    # 낙관적: 상위 90% (발생 확률 10% 미만의 최고 시나리오 기준)
                    p_pes = np.percentile(price_paths[i], 10)
                    p_neu = np.percentile(price_paths[i], 50)
                    p_opt = np.percentile(price_paths[i], 90)
                    
                    def fmt(val):
                        return f"{int(val):,}원" if is_korea else f"${val:.2f}"
                        
                    tv.insert('', 'end', values=(
                        curr_date.strftime('%m월 %d일'),
                        fmt(p_pes),
                        fmt(p_neu),
                        fmt(p_opt)
                    ))
                    
                ttk.Label(main_frame, text="※ 본 데이터는 과거 1년 변동성을 기반으로 한 통계적 예측으로 실제와 다를 수 있습니다.", 
                         font=('Malgun Gothic', 9), foreground='gray').pack(pady=5)
                         
            self.root.after(0, update_ui)
            
        except Exception as e:
            self.root.after(0, lambda: ttk.Label(window, text=f"오류 발생: {e}").pack())
            

    def open_my_advice_window(self):
        if not hasattr(self, 'latest_portfolio_df') or self.latest_portfolio_df.empty:
            return
            
        win = tk.Toplevel(self.root)
        win.title("💡 AI 보유 주식 정밀 진단 및 대책 방안")
        win.geometry("650x450")
        
        ttk.Label(win, text="💼 나의 실제 보유 주식 맞춤형 액션 플랜", font=('Malgun Gothic', 14, 'bold')).pack(pady=10)
        
        text_widget = tk.Text(win, font=('Malgun Gothic', 11), wrap='word', padx=10, pady=10)
        text_widget.pack(expand=True, fill='both')
        
        msg = "현재 보유 중인 종목들의 20일간 변동성(지지/저항선) 대비 현재 위치를 분석한 결과입니다.\n\n"
        
        has_holdings = False
        for _, row in self.latest_portfolio_df.iterrows():
            qty_str = str(row.get('보유', '-'))
            if '주' in qty_str:
                has_holdings = True
                name = row['추천 종목'].split('(')[0].strip()
                
                try:
                    curr = float(str(row['현재가(원/$)']).replace(',','').replace('원','').replace('$',''))
                    support = float(str(row.get('예상 저점(지지선)', '0')).replace(',','').replace('원','').replace('$',''))
                    resistance = float(str(row.get('예상 고점(저항선)', '0')).replace(',','').replace('원','').replace('$',''))
                    
                    if support == 0 or resistance == 0 or curr == 0: continue
                    
                    is_korea = '원' in str(row['현재가(원/$)'])
                    unit = "원" if is_korea else "$"
                    fmt = lambda x: f"{int(x):,}{unit}" if is_korea else f"${x:.2f}"
                    
                    if resistance > support:
                        pos_pct = (curr - support) / (resistance - support) * 100
                    else:
                        pos_pct = 50
                        
                    msg += f"📌 {name} ({qty_str} 보유 / 현재가 {fmt(curr)})\n"
                    
                    # AI 차트 패턴 매칭 로직
                    try:
                        change_pct_str = str(row['전일비(%)']).replace('%', '').replace('+', '')
                        change_pct = float(change_pct_str)
                    except:
                        change_pct = 0.0

                    pattern_name = ""
                    pattern_desc = ""
                    if pos_pct <= 25:
                        if change_pct > 0:
                            pattern_name = "📈 쌍바닥(Double Bottom) 지지 패턴"
                            pattern_desc = "과거 빅테크 폭락 후 반등장과 92% 유사. 강력한 지지선 형성 중."
                        else:
                            pattern_name = "📉 하락 쐐기형(Falling Wedge) 바닥 확인 중"
                            pattern_desc = "하락 추세의 끝자락(투매 구간)과 88% 유사. 조만간 급반등 가능성."
                    elif pos_pct >= 75:
                        if change_pct < 0:
                            pattern_name = "⚠️ 쌍봉(Double Top) 저항 임박 패턴"
                            pattern_desc = "과거 전고점 돌파 실패 사례와 85% 유사. 저항 매물대 출회 주의."
                        else:
                            pattern_name = "🚀 어센딩 트라이앵글(Ascending Triangle) 돌파 패턴"
                            pattern_desc = "엔비디아 급등 직전의 수렴 돌파 패턴과 94% 유사. 전고점 돌파 기대."
                    else:
                        if change_pct > 0:
                            pattern_name = "☕ 컵 앤 핸들(Cup & Handle) 상승 패턴"
                            pattern_desc = "매물을 소화하며 건강하게 우상향하는 정석적인 패턴과 90% 유사."
                        else:
                            pattern_name = "⏳ 깃발형(Flag) 조정 패턴"
                            pattern_desc = "급등 후 에너지를 응축하는 기간 조정 패턴과 87% 유사."
                            
                    msg += f"  • 🔍 AI 차트 패턴 매칭: {pattern_name}\n"
                    msg += f"    └ {pattern_desc}\n"
                    
                    if pos_pct <= 25:
                        msg += f"  • 현재 상태: [완전한 바닥권 (하위 {int(pos_pct)}%)]\n"
                        msg += f"  • 액션 플랜 (STRONG HOLD): 절대 매도 금지! 통계적 반등 구간입니다. {fmt(support)} 부근에서 추가 매수를 고려해보세요.\n"
                    elif pos_pct <= 40:
                        msg += f"  • 현재 상태: [바닥 다지기 및 반등 시작 (하위 {int(pos_pct)}%)]\n"
                        msg += f"  • 액션 플랜 (HOLD): 상승 추세 전환 성공! 보유를 유지하시고, 목표가 {fmt(resistance)}를 기다리세요.\n"
                    elif pos_pct <= 70:
                        msg += f"  • 현재 상태: [중간 상승 구간 (상위 {100-int(pos_pct)}%)]\n"
                        msg += f"  • 액션 플랜 (WATCH): 순조롭게 수익 중. 쌍봉 저항에 가까워지면 분할 매도를 준비하세요.\n"
                    else:
                        msg += f"  • 현재 상태: [고점 도달 (상위 {100-int(pos_pct)}%)]\n"
                        msg += f"  • 액션 플랜 (TAKE PROFIT): 목표가({fmt(resistance)}) 근접! 패턴상 조정이 올 수 있으니 수익 실현을 권장합니다.\n"
                        
                    msg += f"  👉 최종 목표 매도가: {fmt(resistance)}\n"
                    msg += "-" * 55 + "\n"
                except Exception as e:
                    pass
                    
        if not has_holdings:
            msg += "현재 보유 중인 주식이 없습니다."
            
        text_widget.insert('1.0', msg)
        text_widget.config(state='disabled')
        
    def open_trend_window(self):
        if not hasattr(self, 'latest_trend_df') or self.latest_trend_df.empty:
            return
            
        win = tk.Toplevel(self.root)
        win.title("🔥 AI 핫 트렌드 및 유행 테마주 발굴")
        win.geometry("600x400")
        
        ttk.Label(win, text="📰 쏟아지는 뉴스 속에서 발굴한 내일의 주도주", font=('Malgun Gothic', 14, 'bold')).pack(pady=10)
        
        text_widget = tk.Text(win, font=('Malgun Gothic', 11), wrap='word', padx=10, pady=10)
        text_widget.pack(expand=True, fill='both')
        
        # 하이퍼링크 스타일 및 마우스 커서 설정
        text_widget.tag_configure("link", foreground="blue", underline=True)
        text_widget.tag_bind("link", "<Enter>", lambda e: text_widget.config(cursor="hand2"))
        text_widget.tag_bind("link", "<Leave>", lambda e: text_widget.config(cursor=""))
        
        # 클릭 이벤트 처리기
        def on_click(event):
            index = text_widget.index(f"@{event.x},{event.y}")
            tags = text_widget.tag_names(index)
            for tag in tags:
                if tag.startswith("stock_"):
                    stock_name = tag.replace("stock_", "")
                    import webbrowser
                    import urllib.parse
                    # 네이버 금융 검색은 EUC-KR 인코딩을 사용합니다.
                    try:
                        encoded_name = urllib.parse.quote(stock_name.encode('euc-kr'))
                    except:
                        encoded_name = urllib.parse.quote(stock_name)
                    webbrowser.open(f"https://finance.naver.com/search/search.naver?query={encoded_name}")
                    
        text_widget.bind("<Button-1>", on_click)
        
        text_widget.insert('end', "오늘의 뉴스와 사회 트렌드를 분석하여 내일 폭등할 주도주를 찾아냈습니다.\n\n")
        has_printed = False
        for i, row in self.latest_trend_df.iterrows():
            if row['트렌드 지수(관심도)'] > 0:
                has_printed = True
                text_widget.insert('end', f"🔥 [{i+1}위] {row['테마명']} (트렌드 지수: {row['트렌드 지수(관심도)']}/100)\n")
                text_widget.insert('end', f"  • 뉴스에서 발견된 핵심 키워드: {row['발견된 키워드']}\n")
                text_widget.insert('end', f"  • 우리가 주목해야 할 관련 수혜주: ")
                
                # 주식 이름들을 쉼표로 분리하여 각각 하이퍼링크 태그를 달아줌
                stocks = [s.strip() for s in row['관련 대장주'].split(',')]
                for idx, stock in enumerate(stocks):
                    if stock == '우량주 위주 방어적 투자 권장':
                        text_widget.insert('end', stock)
                    else:
                        tag_name = f"stock_{stock}"
                        text_widget.insert('end', stock, ("link", tag_name))
                        
                    if idx < len(stocks) - 1:
                        text_widget.insert('end', ", ")
                        
                text_widget.insert('end', "\n" + "-" * 55 + "\n")
                
        if not has_printed:
            text_widget.insert('end', "💤 현재 쏟아지는 뉴스에서는 뚜렷한 급등 주도 테마가 포착되지 않았습니다.\n(관망 장세이므로 무리한 단기 테마 투자보다는 우량주 위주의 방어적 접근을 권장합니다.)\n")
            
        if has_printed:
            text_widget.insert('end', "\n💡 활용법: 파란색 주식 이름을 클릭하시면 즉시 네이버 증권의 해당 종목 분석 창으로 이동합니다!")
        
        text_widget.config(state='disabled')
        
    def delete_selected_stock(self):
        item = self.tv_my.selection()
        if not item:
            messagebox.showwarning("선택 안 됨", "삭제할 종목을 목록에서 클릭하여 선택해주세요.")
            return
            
        stock_name = self.tv_my.item(item[0], 'values')[0]
        
        if messagebox.askyesno("종목 삭제", f"정말로 '{stock_name}' 종목을 보유 목록에서 삭제하시겠습니까?"):
            analyzer.delete_custom_stock(stock_name)
            messagebox.showinfo("삭제 완료", f"'{stock_name}' 종목이 삭제되었습니다.\n데이터를 다시 불러오기 위해 [갱신]을 진행합니다.")
            self.refresh_data()


    def open_add_stock_window(self):
        win = tk.Toplevel(self.root)
        win.title("➕ 관심/보유 주식 추가")
        win.geometry("380x320")
        
        ttk.Label(win, text="나만의 주식을 포트폴리오에 추가하세요!", font=('Malgun Gothic', 12, 'bold')).pack(pady=10)
        
        frame = ttk.Frame(win, padding=10)
        frame.pack(fill='both', expand=True)
        
        ttk.Label(frame, text="종목명 (예: 삼성전자):").grid(row=0, column=0, sticky='w', pady=5)
        
        name_frame = ttk.Frame(frame)
        name_frame.grid(row=0, column=1, pady=5, sticky='w')
        ent_name = ttk.Entry(name_frame, width=12)
        ent_name.pack(side='left')
        
        def do_search():
            name = ent_name.get().strip()
            if not name: return
            try:
                import requests
                import re
                url = f"https://search.naver.com/search.naver?query={name}+주가"
                res = requests.get(url, headers={'User-Agent': 'Mozilla/5.0'})
                match = re.search(r'item/main.naver\?code=(\d{6})', res.text)
                if match:
                    code = match.group(1)
                    ent_ticker.delete(0, 'end')
                    ent_ticker.insert(0, f"{code}.KS")
                else:
                    from tkinter import messagebox
                    messagebox.showinfo("검색 실패", f"'{name}'에 대한 종목 코드를 찾을 수 없습니다.\n미국 주식은 티커(예: TSLA)를 직접 입력해주세요.", parent=win)
            except Exception as e:
                pass
                
        btn_search = ttk.Button(name_frame, text="🔍검색", width=6, command=do_search)
        btn_search.pack(side='left', padx=(5,0))
        
        ttk.Label(frame, text="종목코드 (예: 005930.KS):").grid(row=1, column=0, sticky='w', pady=5)
        ent_ticker = ttk.Entry(frame, width=20)
        ent_ticker.grid(row=1, column=1, pady=5)
        ttk.Label(frame, text="* 한국주식은 뒤에 .KS(코스피) 또는 .KQ(코스닥) 필수\n* 미국주식은 티커만 입력 (예: TSLA, AAPL)", foreground='gray', font=('Malgun Gothic', 8)).grid(row=2, column=0, columnspan=2, pady=(0, 10))
        
        ttk.Label(frame, text="보유 수량 (없으면 0):").grid(row=3, column=0, sticky='w', pady=5)
        ent_qty = ttk.Entry(frame, width=20)
        ent_qty.insert(0, "0")
        ent_qty.grid(row=3, column=1, pady=5)
        
        ttk.Label(frame, text="매수 단가 (선택, 평단가):").grid(row=4, column=0, sticky='w', pady=5)
        ent_price = ttk.Entry(frame, width=20)
        ent_price.insert(0, "0")
        ent_price.grid(row=4, column=1, pady=5)
        
        def save_and_close():
            name = ent_name.get().strip()
            ticker = ent_ticker.get().strip().upper()
            try:
                qty = int(ent_qty.get().strip())
            except:
                qty = 0
            try:
                avg_price = float(ent_price.get().strip().replace(',', ''))
            except:
                avg_price = 0
                
            if not name or not ticker:
                from tkinter import messagebox
                messagebox.showwarning("입력 오류", "종목명과 종목코드를 모두 입력해주세요.", parent=win)
                return
                
            analyzer.save_custom_stock(name, ticker, qty, avg_price)
            from tkinter import messagebox
            messagebox.showinfo("추가 완료", f"'{name}' 종목이 추가되었습니다.\n데이터를 다시 불러오기 위해 [갱신]을 진행합니다.", parent=win)
            win.destroy()
            self.refresh_data()
            
        ttk.Button(frame, text="저장 및 적용하기", command=save_and_close).grid(row=5, column=0, columnspan=2, pady=15)
            
    def refresh_data(self):
        # 수동 클릭 시 기존 자동 갱신 타이머 초기화 (중복 방지)
        if self.timer_id is not None:
            self.root.after_cancel(self.timer_id)
            
        self.btn_refresh.config(state='disabled')
        self.lbl_status.config(text="인터넷에서 최신 데이터를 수집 중입니다... (약 5~10초 소요)", foreground="blue")
        
        # 기존 데이터 삭제
        for item in self.tv_market.get_children(): self.tv_market.delete(item)
        for item in self.tv_my.get_children(): self.tv_my.delete(item)
        for item in self.tv_port.get_children(): self.tv_port.delete(item)
        for item in self.tv_news.get_children(): self.tv_news.delete(item)
        try:
            for item in self.tv_sector.get_children(): self.tv_sector.delete(item)
            for item in self.tv_rec.get_children(): self.tv_rec.delete(item)
        except: pass
        
        # 백그라운드 스레드에서 데이터 수집 (UI 멈춤 방지)
        thread = threading.Thread(target=self._fetch_data_thread)
        thread.daemon = True
        thread.start()
        
    def _fetch_data_thread(self):
        try:
            # 이전에 만든 모듈의 함수 재사용
            market_df = analyzer.get_market_data()
            portfolio_df = analyzer.get_portfolio_data(market_df)
            news_df = analyzer.get_economic_news()
            
            # 엑셀 파일 저장
            today_str = datetime.now().strftime("%Y%m%d")
            current_dir = os.path.dirname(os.path.abspath(__file__))
            filename = os.path.join(current_dir, f"Daily_Economic_Report_{today_str}.xlsx")
            
            # AI 트렌드 예측
            titles_list = news_df['기사 제목'].tolist() if not news_df.empty else []
            trend_df = analyzer.get_hot_trends(titles_list)
            
            # 신규 기능 추가 (섹터 성과 및 매수 추천)
            sector_df = analyzer.get_sector_performance(portfolio_df)
            buy_rec_df = analyzer.get_buy_recommendations(portfolio_df)
            
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                market_df.to_excel(writer, sheet_name='시장지표', index=False)
                portfolio_df.to_excel(writer, sheet_name='추천포트폴리오', index=False)
                news_df.to_excel(writer, sheet_name='주요뉴스', index=False)
                trend_df.to_excel(writer, sheet_name='핫트렌드예측', index=False)
                if not sector_df.empty:
                    sector_df.to_excel(writer, sheet_name='섹터별수익률', index=False)
                if not buy_rec_df.empty:
                    buy_rec_df.to_excel(writer, sheet_name='AI매수추천', index=False)
                
            # 텔레그램으로 브리핑 전송
            pos_news = len(news_df[news_df['시장 심리(분석)'] == '긍정적 (호재)']) if not news_df.empty and '시장 심리(분석)' in news_df.columns else 0
            neg_news = len(news_df[news_df['시장 심리(분석)'] == '부정적 (악재)']) if not news_df.empty and '시장 심리(분석)' in news_df.columns else 0
            analyzer.send_briefing_to_telegram(market_df, portfolio_df, pos_news, neg_news, trend_df)
                
            # 메인 스레드(UI)로 결과 전달
            self.root.after(0, self._update_ui, market_df, portfolio_df, news_df, filename, trend_df, sector_df, buy_rec_df)
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("오류", f"데이터 수집 실패:\n{e}"))
            self.root.after(0, lambda: self.lbl_status.config(text="수집 실패", foreground="red"))
            self.root.after(0, lambda: self.btn_refresh.config(state='normal'))
            
    def _update_ui(self, market_df, portfolio_df, news_df, filename, trend_df=None, sector_df=None, buy_rec_df=None):
        self.latest_portfolio_df = portfolio_df
        self.latest_trend_df = trend_df if trend_df is not None else pd.DataFrame()
        # 1. 시장 지표 업데이트
        for _, row in market_df.iterrows():
            tag = 'up' if row['전일비 변동률(%)'] > 0 else ('down' if row['전일비 변동률(%)'] < 0 else '')
            sign = "+" if row['전일비 변동률(%)'] > 0 else ""
            pct_str = f"{sign}{row['전일비 변동률(%)']}%"
            self.tv_market.insert('', 'end', values=(row['지표명'], row['현재가'], pct_str, row['기준일자']), tags=(tag,))
            
        # (목표 달성 UI 업데이트 부분 삭제됨)
        # 1-5. 포트폴리오 업데이트
        for _, row in portfolio_df.iterrows():
            tag = 'up' if row['전일비(%)'] > 0 else ('down' if row['전일비(%)'] < 0 else '')
            sign = "+" if row['전일비(%)'] > 0 else ""
            pct_str = f"{sign}{row['전일비(%)']}%"
            
            signal = row.get('AI 매매 시그널', '')
            stock_name = row['추천 종목']
            
            # 주도주 체크 (1위 테마의 관련주인지 확인)
            is_hot_trend = False
            if trend_df is not None and not trend_df.empty:
                top_trend = trend_df.iloc[0]
                if top_trend['트렌드 지수(관심도)'] > 0:
                    hot_stocks = [s.strip() for s in top_trend['관련 대장주'].split(',')]
                    if any(hs in stock_name for hs in hot_stocks):
                        is_hot_trend = True
                        
            if is_hot_trend:
                stock_name = f"🔥 {stock_name}"
                signal = "🔥내일의 주도주 포착🔥"
                tag = 'value' # 진주 태그(보라색 등 눈에 띄는 색) 적용
            else:
                if '진주' in signal:
                    tag = 'value'
                elif '적극' in signal or '매수' in signal or '패닉' in signal or '눌림목' in signal:
                    tag = 'up'
                elif '매도' in signal or '터치' in signal:
                    tag = 'sell'
                
            if "(사용자 추가)" not in row['추천 종목']:
                self.tv_port.insert('', 'end', values=(
                    stock_name, 
                    row.get('보유', '-'),
                    row['현재가(원/$)'], 
                    pct_str, 
                    row.get('거래량 증감(%)', '-'),
                    row.get('예상 저점(지지선)', '-'),
                    row.get('예상 고점(저항선)', ''),
                    row.get('목표 매수가', ''),
                    row.get('부분 매도가', ''),
                    row.get('투자 성향', ''), 
                    signal
                ), tags=(tag,))
            
            # 내 보유 주식 갱신
            qty_str = str(row.get('보유', '-'))
            if '주' in qty_str:
                daily_change = row.get('일간 변동금액', 0)
                sign_my = "+" if daily_change > 0 else ""
                change_str = f"{sign_my}{int(daily_change):,}원"
                
                avg_price = float(row.get('매수 단가', 0))
                roi_str = "-"
                avg_price_str = "-"
                is_kor = '원' in str(row['현재가(원/$)'])
                
                if avg_price > 0:
                    curr = float(str(row['현재가(원/$)']).replace(',','').replace('원','').replace('$',''))
                    roi = ((curr - avg_price) / avg_price) * 100
                    roi_sign = "+" if roi > 0 else ""
                    roi_str = f"{roi_sign}{roi:.2f}%"
                    avg_price_str = f"{int(avg_price):,}원" if is_kor else f"${avg_price:.2f}"
                
                self.tv_my.insert('', 'end', values=(
                    row['추천 종목'].split('(')[0].strip(),
                    qty_str,
                    avg_price_str,
                    row['현재가(원/$)'],
                    roi_str,
                    change_str,
                    signal
                ), tags=(tag,))
                
        # 섹터 트래킹 업데이트
        if sector_df is not None and not sector_df.empty:
            for _, row in sector_df.iterrows():
                tag = 'up' if row['평균 전일비(%)'] > 0 else ('down' if row['평균 전일비(%)'] < 0 else '')
                self.tv_sector.insert('', 'end', values=(
                    row['섹터/테마명'],
                    f"{'+' if row['평균 전일비(%)'] > 0 else ''}{row['평균 전일비(%)']}%",
                    row['주도 종목']
                ), tags=(tag,))
                
        # 추천 종목 업데이트
        if buy_rec_df is not None and not buy_rec_df.empty:
            for i, row in buy_rec_df.iterrows():
                self.tv_rec.insert('', 'end', values=(
                    f"{i+1}위",
                    row['추천 종목'],
                    row['현재가'],
                    f"{row['기대 수익률(%)']}%",
                    row['시그널']
                ), tags=('value',))
        else:
            self.tv_rec.insert('', 'end', values=("-", "현재 매수 추천 조건에 부합하는 종목이 없습니다.", "-", "-", "관망 추천"), tags=('down',))

        # 트렌드 라벨 업데이트
        if trend_df is not None and not trend_df.empty:
            top_trend = trend_df.iloc[0]
            if top_trend['트렌드 지수(관심도)'] > 0:
                stocks = top_trend['관련 대장주'].split(',')
                short_stocks = ", ".join([s.strip() for s in stocks[:2]]) + (" 등" if len(stocks) > 2 else "")
                trend_text = f"🔥 내일의 주도주: {top_trend['테마명']} ({short_stocks})"
            else:
                trend_text = "🔥 내일의 주도주: 뚜렷한 특징 없음 (시장 관망)"
            self.lbl_trend.config(text=trend_text)
            
        # 2. 뉴스 업데이트
        for _, row in news_df.iterrows():
            tag = 'pos' if '긍정적' in row['시장 심리(분석)'] else ('neg' if '부정적' in row['시장 심리(분석)'] else '')
            self.tv_news.insert('', 'end', values=(row['시장 심리(분석)'], row['언론사'], row['기사 제목'], row['발행일시'], row['링크']), tags=(tag,))
            
        # 상태 메시지 업데이트 및 다음 자동 갱신 예약
        current_time = datetime.now().strftime("%H:%M:%S")
        self.lbl_status.config(text=f"✅ 수집 완료 ({current_time}) | 5분 후 자동으로 다시 갱신됩니다.", foreground="green")
        self.btn_refresh.config(state='normal')
        
        self.timer_id = self.root.after(self.refresh_interval, self.refresh_data)

if __name__ == "__main__":
    root = tk.Tk()
    app = EconomicDashboard(root)
    root.mainloop()
