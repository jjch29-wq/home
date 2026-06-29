import tkinter as tk
from tkinter import ttk, messagebox
import threading
import webbrowser
import os
import pandas as pd
from datetime import datetime
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
        
        self.lbl_goal = ttk.Label(top_frame, text="🎯 하루 1만 원 수익 목표: 데이터 수집 중...", font=('Malgun Gothic', 12, 'bold'), foreground="purple")
        self.lbl_goal.pack(side='left', padx=15)
        
        self.btn_swing = ttk.Button(top_frame, text="🏆 1만 원 스윙 추천", command=self.open_swing_window)
        self.btn_swing.pack(side='left', padx=5)
        
        self.btn_trend = ttk.Button(top_frame, text="🔥 AI 핫 트렌드 예측", command=self.open_trend_window)
        self.btn_trend.pack(side='left', padx=5)
        
        self.btn_add_stock = ttk.Button(top_frame, text="➕ 내 주식 추가", command=self.open_add_stock_window)
        self.btn_add_stock.pack(side='left', padx=5)
        
        self.btn_refresh = ttk.Button(top_frame, text="🔄 데이터 갱신 및 엑셀 저장", command=self.refresh_data)
        self.btn_refresh.pack(side='right')
        
        self.lbl_status = ttk.Label(top_frame, text="프로그램 시작 중...", font=('Malgun Gothic', 10), foreground="gray")
        self.lbl_status.pack(side='right', padx=15)
        
        # 화면 분할용 PanedWindow
        self.paned = ttk.PanedWindow(root, orient='vertical')
        self.paned.pack(expand=True, fill='both', padx=10, pady=10)
        
        # 상단 영역 분할 (왼쪽: 시장 지표 및 내 보유주식, 오른쪽: 추천 포트폴리오)
        self.top_content_frame = ttk.PanedWindow(self.paned, orient='horizontal')
        self.paned.add(self.top_content_frame, weight=1)
        
        # 1. 왼쪽 컨테이너 (시장 지표 + 보유 주식)
        left_frame = ttk.Frame(self.top_content_frame)
        self.top_content_frame.add(left_frame, weight=1)
        
        # 1-1. 시장 지표 섹션 (왼쪽 위)
        market_frame = ttk.LabelFrame(left_frame, text=" 📌 주요 시장 지표 (KOSPI, 환율 등) ", padding=5)
        market_frame.pack(side='top', expand=False, fill='x')
        
        cols_market = ('지표명', '현재가', '전일비(%)', '기준일자')
        self.tv_market = ttk.Treeview(market_frame, columns=cols_market, show='headings', height=4)
        for col in cols_market:
            self.tv_market.heading(col, text=col)
            self.tv_market.column(col, anchor='center', width=100)
        self.tv_market.pack(expand=True, fill='both')
        
        # 1-2. 나의 보유 주식 섹션 (왼쪽 아래)
        my_port_frame = ttk.LabelFrame(left_frame, text=" 💼 나의 실제 보유 주식 현황 ", padding=5)
        my_port_frame.pack(side='bottom', expand=True, fill='both', pady=(10, 0))
        
        cols_my = ('종목명', '보유', '현재가', '일간 변동금액', '전일비(%)')
        self.tv_my = ttk.Treeview(my_port_frame, columns=cols_my, show='headings', height=6)
        for col in cols_my:
            self.tv_my.heading(col, text=col)
            w = 140 if col == '종목명' else (100 if col == '일간 변동금액' else 80)
            self.tv_my.column(col, anchor='center', width=w, stretch=False)
        
        # 버튼 및 주도주 표시 프레임
        btn_frame = ttk.Frame(my_port_frame)
        btn_frame.pack(side='bottom', fill='x', pady=(5, 0))
        
        self.lbl_trend = ttk.Label(btn_frame, text="🔥 내일의 주도주 예측 중...", font=('Malgun Gothic', 10, 'bold'), foreground='red')
        self.lbl_trend.pack(side='left', padx=5)
        
        btn_my_advice = ttk.Button(btn_frame, text="💡 AI 보유 주식 정밀 진단 및 대책", command=self.open_my_advice_window)
        btn_my_advice.pack(side='right')

        # 스크롤바 추가
        my_scroll_y = ttk.Scrollbar(my_port_frame, orient='vertical', command=self.tv_my.yview)
        my_scroll_x = ttk.Scrollbar(my_port_frame, orient='horizontal', command=self.tv_my.xview)
        self.tv_my.configure(yscroll=my_scroll_y.set, xscroll=my_scroll_x.set)
        
        my_scroll_x.pack(side='bottom', fill='x')
        my_scroll_y.pack(side='right', fill='y')
        self.tv_my.pack(expand=True, fill='both')
        
        # 1-5. 추천 포트폴리오 섹션 (오른쪽)
        port_frame = ttk.LabelFrame(self.top_content_frame, text=" 💡 100만 원 초보자 맞춤 포트폴리오 (실시간) ", padding=5)
        self.top_content_frame.add(port_frame, weight=3)
        
        cols_port = ('추천 종목', '보유', '현재가(원/$)', '전일비(%)', '예상 저점(지지선)', '예상 고점(저항선)', '목표 매수가', '부분 매도가', '투자 성향', 'AI 매매 시그널')
        self.tv_port = ttk.Treeview(port_frame, columns=cols_port, show='headings', height=15)
        for col in cols_port:
            self.tv_port.heading(col, text=col)
            if col == '추천 종목': w = 180
            elif col == '보유': w = 60
            elif col == 'AI 매매 시그널': w = 150
            elif col in ['목표 매수가', '부분 매도가', '예상 저점(지지선)', '예상 고점(저항선)']: w = 110
            else: w = 80
            self.tv_port.column(col, anchor='center', width=w, stretch=False)
            
        # 포트폴리오 스크롤바 추가
        port_scroll_y = ttk.Scrollbar(port_frame, orient='vertical', command=self.tv_port.yview)
        port_scroll_x = ttk.Scrollbar(port_frame, orient='horizontal', command=self.tv_port.xview)
        self.tv_port.configure(yscroll=port_scroll_y.set, xscroll=port_scroll_x.set)
        
        port_scroll_x.pack(side='bottom', fill='x')
        port_scroll_y.pack(side='right', fill='y')
        self.tv_port.pack(expand=True, fill='both')
        
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
        
        # 색상 태그 설정 (상승=빨강, 하락=파랑)
        self.tv_news.tag_configure('pos', foreground='red')
        self.tv_news.tag_configure('neg', foreground='blue')
        self.tv_market.tag_configure('up', foreground='red')
        self.tv_market.tag_configure('down', foreground='blue')
        self.tv_port.tag_configure('up', foreground='red')
        self.tv_port.tag_configure('down', foreground='blue')
        self.tv_port.tag_configure('sell', foreground='purple')  # 매도 시그널용 색상
        self.tv_port.tag_configure('value', foreground='darkorange') # 숨은 진주용 색상
        
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
        item = self.tv_port.selection()
        if not item: return
        stock_name = self.tv_port.item(item[0], 'values')[0]
        
        pred_win = tk.Toplevel(self.root)
        pred_win.title(f"{stock_name} - 향후 5일 가격 예측 시뮬레이션")
        pred_win.geometry("600x350")
        
        ttk.Label(pred_win, text="데이터 분석 중... 잠시만 기다려주세요.", font=('Malgun Gothic', 12)).pack(pady=50)
        
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
                    
                ttk.Label(window, text=f"📊 {stock_name} 5영업일 시뮬레이션", font=('Malgun Gothic', 14, 'bold')).pack(pady=10)
                
                cols = ('날짜', '비관적 (하락)', '중립 (추세)', '낙관적 (상승)')
                tv = ttk.Treeview(window, columns=cols, show='headings', height=5)
                for col in cols:
                    tv.heading(col, text=col)
                    tv.column(col, anchor='center', width=130)
                tv.pack(fill='both', expand=True, padx=10, pady=10)
                
                p_neu, p_pes, p_opt = current_price, current_price, current_price
                curr_date = datetime.now()
                
                for i in range(1, 6):
                    curr_date += timedelta(days=1)
                    while curr_date.weekday() > 4:
                        curr_date += timedelta(days=1)
                        
                    p_neu = p_neu * (1 + mu)
                    p_pes = p_pes * (1 + mu - sigma)
                    p_opt = p_opt * (1 + mu + sigma)
                    
                    def fmt(val):
                        return f"{int(val):,}원" if is_korea else f"${val:.2f}"
                        
                    tv.insert('', 'end', values=(
                        curr_date.strftime('%m월 %d일'),
                        fmt(p_pes),
                        fmt(p_neu),
                        fmt(p_opt)
                    ))
                    
                ttk.Label(window, text="※ 본 데이터는 과거 1년 변동성을 기반으로 한 통계적 예측으로 실제와 다를 수 있습니다.", 
                         font=('Malgun Gothic', 9), foreground='gray').pack(pady=5)
                         
            self.root.after(0, update_ui)
            
        except Exception as e:
            self.root.after(0, lambda: ttk.Label(window, text=f"오류 발생: {e}").pack())
            
    def open_swing_window(self):
        if not hasattr(self, 'latest_portfolio_df') or self.latest_portfolio_df.empty:
            return
            
        win = tk.Toplevel(self.root)
        win.title("🏆 1만 원 수익 달성 스윙 추천 종목")
        win.geometry("600x400")
        
        ttk.Label(win, text="📊 현재 가장 저평가된 스윙 매매 추천 종목 TOP 3", font=('Malgun Gothic', 14, 'bold')).pack(pady=10)
        
        results = []
        for _, row in self.latest_portfolio_df.iterrows():
            name = row['추천 종목'].split('(')[0].strip()
            
            try:
                curr = float(str(row['현재가(원/$)']).replace(',','').replace('원','').replace('$',''))
                support = float(str(row.get('예상 저점(지지선)', '0')).replace(',','').replace('원','').replace('$',''))
                resistance = float(str(row.get('예상 고점(저항선)', '0')).replace(',','').replace('원','').replace('$',''))
                
                if support == 0 or resistance == 0 or curr == 0: continue
                
                # 보수적인 매도가 적용 (-5%)
                resistance = resistance * 0.95
                
                profit_per_share = resistance - curr
                if profit_per_share > 0:
                    dist = (curr - support) / support
                    shares_needed = int(10000 / profit_per_share) + 1
                    capital_needed = shares_needed * curr
                    results.append({
                        'name': name, 'curr': curr, 'support': support, 'resistance': resistance,
                        'dist': dist, 'profit': profit_per_share, 'shares': shares_needed, 'capital': capital_needed,
                        'is_korea': '원' in str(row['현재가(원/$)'])
                    })
            except Exception as e:
                pass
                
        results.sort(key=lambda x: x['dist'])
        
        text_widget = tk.Text(win, font=('Malgun Gothic', 11), wrap='word', padx=10, pady=10)
        text_widget.pack(expand=True, fill='both')
        
        msg = "목표: 1주당 매도 수익을 극대화하여 1만 원 벌기\n\n"
        for i, r in enumerate(results[:3], 1):
            unit = "원" if r['is_korea'] else "$"
            fmt = lambda x: f"{int(x):,}{unit}" if r['is_korea'] else f"${x:.2f}"
            msg += f"[{i}순위] {r['name']} (현재가 {fmt(r['curr'])})\n"
            msg += f" • 추천 매수: {fmt(r['support'])} 부근 예약 매수\n"
            msg += f" • 추천 매도: {fmt(r['resistance'])} 부근 전량 매도\n"
            msg += f" • 1만원 달성법: {r['shares']}주 매수 (필요 자본금 약 {fmt(r['capital'])})\n"
            msg += "-" * 50 + "\n"
            
        msg += "\n💡 팁: 현재가에 바로 사지 마시고, 증권사 앱에서 추천 매수가(지지선)에 예약 주문을 걸어두세요!"
        text_widget.insert('1.0', msg)
        text_widget.config(state='disabled')
            
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
            qty_str = str(row.get('보유 수량', '-'))
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
                    
                    if pos_pct <= 20:
                        msg += f"  • 현재 상태: [완전한 바닥권 (하위 {int(pos_pct)}%)]\n"
                        msg += f"  • 대책 방안 (STRONG HOLD / BUY MORE): 절대 매도 금지! 지금 팔면 최하점에서 파는 격입니다. 오히려 {fmt(support)} 부근에서 추가 매수를 고려해보세요.\n"
                    elif pos_pct <= 40:
                        msg += f"  • 현재 상태: [바닥 다지기 및 반등 시작 (하위 {int(pos_pct)}%)]\n"
                        msg += f"  • 대책 방안 (HOLD): 상승 추세로 전환될 가능성이 큽니다. 보유를 유지하시고, 목표가 {fmt(resistance)}를 기다리세요.\n"
                    elif pos_pct <= 70:
                        msg += f"  • 현재 상태: [중간 상승 구간 (상위 {100-int(pos_pct)}%)]\n"
                        msg += f"  • 대책 방안 (WATCH): 순조롭게 수익이 커지고 있습니다. 조금 더 지켜보다가 저항선에 가까워지면 매도를 준비하세요.\n"
                    else:
                        msg += f"  • 현재 상태: [고점 도달 (상위 {100-int(pos_pct)}%)]\n"
                        msg += f"  • 대책 방안 (SELL / TAKE PROFIT): 목표가({fmt(resistance)})에 근접했습니다! 욕심을 줄이고 수익을 실현(매도)할 타이밍입니다.\n"
                        
                    msg += f"  👉 목표 매도가: {fmt(resistance)}\n"
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
        
        msg = "오늘의 경제 뉴스를 분석하여 가장 뜨거운 반응을 얻고 있는 테마와 관련 주식을 찾아냈습니다.\n\n"
        
        for i, row in self.latest_trend_df.iterrows():
            if row['트렌드 지수(관심도)'] > 0:
                msg += f"🔥 [{i+1}위] {row['테마명']} (트렌드 지수: {row['트렌드 지수(관심도)']}/100)\n"
                msg += f"  • 뉴스에서 발견된 핵심 키워드: {row['발견된 키워드']}\n"
                msg += f"  • 우리가 주목해야 할 관련 수혜주: {row['관련 대장주']}\n"
                msg += "-" * 55 + "\n"
                
        msg += "\n💡 활용법: 지금 바로 증권사 앱을 열어 [관련 수혜주]를 검색해 보세요. 아직 오르지 않았다면 큰 기회일 수 있습니다!"
        
        text_widget.insert('1.0', msg)
        text_widget.config(state='disabled')
        
    def open_add_stock_window(self):
        win = tk.Toplevel(self.root)
        win.title("➕ 관심/보유 주식 추가")
        win.geometry("380x280")
        
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
        
        def save_and_close():
            name = ent_name.get().strip()
            ticker = ent_ticker.get().strip().upper()
            try:
                qty = int(ent_qty.get().strip())
            except:
                qty = 0
                
            if not name or not ticker:
                from tkinter import messagebox
                messagebox.showwarning("입력 오류", "종목명과 종목코드를 모두 입력해주세요.", parent=win)
                return
                
            analyzer.save_custom_stock(name, ticker, qty)
            from tkinter import messagebox
            messagebox.showinfo("추가 완료", f"'{name}' 종목이 추가되었습니다.\n데이터를 다시 불러오기 위해 [갱신]을 진행합니다.", parent=win)
            win.destroy()
            self.refresh_data()
            
        ttk.Button(frame, text="저장 및 적용하기", command=save_and_close).grid(row=4, column=0, columnspan=2, pady=15)
            
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
        
        # 백그라운드 스레드에서 데이터 수집 (UI 멈춤 방지)
        thread = threading.Thread(target=self._fetch_data_thread)
        thread.daemon = True
        thread.start()
        
    def _fetch_data_thread(self):
        try:
            # 이전에 만든 모듈의 함수 재사용
            market_df = analyzer.get_market_data()
            portfolio_df = analyzer.get_portfolio_data()
            news_df = analyzer.get_economic_news()
            
            # 엑셀 파일 저장
            today_str = datetime.now().strftime("%Y%m%d")
            current_dir = os.path.dirname(os.path.abspath(__file__))
            filename = os.path.join(current_dir, f"Daily_Economic_Report_{today_str}.xlsx")
            
            # AI 트렌드 예측
            titles_list = news_df['기사 제목'].tolist() if not news_df.empty else []
            trend_df = analyzer.get_hot_trends(titles_list)
            
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                market_df.to_excel(writer, sheet_name='시장지표', index=False)
                portfolio_df.to_excel(writer, sheet_name='추천포트폴리오', index=False)
                news_df.to_excel(writer, sheet_name='주요뉴스', index=False)
                trend_df.to_excel(writer, sheet_name='핫트렌드예측', index=False)
                
            # 텔레그램으로 브리핑 전송
            pos_news = len(news_df[news_df['시장 심리(분석)'] == '긍정적 (호재)']) if not news_df.empty and '시장 심리(분석)' in news_df.columns else 0
            neg_news = len(news_df[news_df['시장 심리(분석)'] == '부정적 (악재)']) if not news_df.empty and '시장 심리(분석)' in news_df.columns else 0
            analyzer.send_briefing_to_telegram(market_df, portfolio_df, pos_news, neg_news, trend_df)
                
            # 메인 스레드(UI)로 결과 전달
            self.root.after(0, self._update_ui, market_df, portfolio_df, news_df, filename, trend_df)
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("오류", f"데이터 수집 실패:\n{e}"))
            self.root.after(0, lambda: self.lbl_status.config(text="수집 실패", foreground="red"))
            self.root.after(0, lambda: self.btn_refresh.config(state='normal'))
            
    def _update_ui(self, market_df, portfolio_df, news_df, filename, trend_df=None):
        self.latest_portfolio_df = portfolio_df
        self.latest_trend_df = trend_df if trend_df is not None else pd.DataFrame()
        # 1. 시장 지표 업데이트
        for _, row in market_df.iterrows():
            tag = 'up' if row['전일비 변동률(%)'] > 0 else ('down' if row['전일비 변동률(%)'] < 0 else '')
            sign = "+" if row['전일비 변동률(%)'] > 0 else ""
            pct_str = f"{sign}{row['전일비 변동률(%)']}%"
            self.tv_market.insert('', 'end', values=(row['지표명'], row['현재가'], pct_str, row['기준일자']), tags=(tag,))
            
        # 목표 달성률 UI 업데이트
        total_daily_profit = int(portfolio_df['일간 변동금액'].sum())
        if total_daily_profit >= 10000:
            goal_text = f"🎯 하루 1만 원 수익 목표: 달성! 🥳 (+{total_daily_profit:,}원)"
            goal_color = "green"
        elif total_daily_profit > 0:
            goal_text = f"🎯 하루 1만 원 수익 목표: {int((total_daily_profit/10000)*100)}% 진행 중 (+{total_daily_profit:,}원)"
            goal_color = "blue"
        else:
            goal_text = f"🎯 하루 1만 원 수익 목표: 기회 엿보기 ({total_daily_profit:,}원)"
            goal_color = "red"
        self.lbl_goal.config(text=goal_text, foreground=goal_color)

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
                elif '적극' in signal or '좋은' in signal:
                    tag = 'up'
                elif '매도' in signal:
                    tag = 'sell'
                
            self.tv_port.insert('', 'end', values=(
                stock_name, 
                row.get('보유', '-'),
                row['현재가(원/$)'], 
                pct_str, 
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
                
                self.tv_my.insert('', 'end', values=(
                    row['추천 종목'].split('(')[0].strip(),
                    qty_str,
                    row['현재가(원/$)'],
                    change_str,
                    pct_str
                ), tags=(tag,))
                
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
