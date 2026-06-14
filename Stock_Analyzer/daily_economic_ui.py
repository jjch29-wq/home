import tkinter as tk
from tkinter import ttk, messagebox
import threading
import webbrowser
import os
import pandas as pd
from datetime import datetime
import daily_economic_analyzer as analyzer

class EconomicDashboard:
    def __init__(self, root):
        self.root = root
        self.root.title("📈 나만의 경제 뉴스 분석 비서")
        self.root.geometry("1100x750")
        
        # 최상단에 윈도우 창 위치 (잠깐만)
        self.root.attributes('-topmost', True)
        self.root.after(500, lambda: self.root.attributes('-topmost', False))
        
        # 스타일 설정
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Treeview.Heading", font=('Malgun Gothic', 10, 'bold'))
        style.configure("Treeview", font=('Malgun Gothic', 10), rowheight=30)
        
        # 상단 헤더 프레임
        top_frame = ttk.Frame(root, padding=10)
        top_frame.pack(fill='x')
        
        ttk.Label(top_frame, text="📊 데일리 경제 & 증시 대시보드", font=('Malgun Gothic', 18, 'bold')).pack(side='left')
        
        self.btn_refresh = ttk.Button(top_frame, text="🔄 데이터 갱신 및 엑셀 저장", command=self.refresh_data)
        self.btn_refresh.pack(side='right')
        
        self.lbl_status = ttk.Label(top_frame, text="프로그램 시작 중...", font=('Malgun Gothic', 10), foreground="gray")
        self.lbl_status.pack(side='right', padx=15)
        
        # 화면 분할용 PanedWindow
        paned = ttk.PanedWindow(root, orient='vertical')
        paned.pack(expand=True, fill='both', padx=10, pady=10)
        
        # 상단 영역 분할 (왼쪽: 시장 지표, 오른쪽: 추천 포트폴리오)
        top_content_frame = ttk.Frame(paned)
        paned.add(top_content_frame, weight=1)
        
        # 1. 시장 지표 섹션 (왼쪽)
        market_frame = ttk.LabelFrame(top_content_frame, text=" 📌 주요 시장 지표 (KOSPI, 환율 등) ", padding=5)
        market_frame.pack(side='left', expand=True, fill='both', padx=(0, 5))
        
        cols_market = ('지표명', '현재가', '전일비(%)', '기준일자')
        self.tv_market = ttk.Treeview(market_frame, columns=cols_market, show='headings', height=4)
        for col in cols_market:
            self.tv_market.heading(col, text=col)
            self.tv_market.column(col, anchor='center', width=100)
        self.tv_market.pack(expand=True, fill='both')
        
        # 1-5. 추천 포트폴리오 섹션 (오른쪽)
        port_frame = ttk.LabelFrame(top_content_frame, text=" 💡 100만 원 초보자 맞춤 포트폴리오 (실시간) ", padding=5)
        port_frame.pack(side='right', expand=True, fill='both', padx=(5, 0))
        
        cols_port = ('추천 종목', '현재가(원/$)', '전일비(%)', '투자 성향', '매수 타이밍')
        self.tv_port = ttk.Treeview(port_frame, columns=cols_port, show='headings', height=15)
        for col in cols_port:
            self.tv_port.heading(col, text=col)
            if col == '추천 종목': w = 220
            elif col == '매수 타이밍': w = 150
            else: w = 90
            self.tv_port.column(col, anchor='center', width=w)
            
        # 포트폴리오 스크롤바 추가
        port_scroll = ttk.Scrollbar(port_frame, orient='vertical', command=self.tv_port.yview)
        self.tv_port.configure(yscroll=port_scroll.set)
        port_scroll.pack(side='right', fill='y')
        self.tv_port.pack(expand=True, fill='both')
        
        # 2. 주요 뉴스 섹션
        news_frame = ttk.LabelFrame(paned, text=" 📰 최신 경제 및 주식 뉴스 (더블클릭 시 인터넷 창 열림) ", padding=5)
        paned.add(news_frame, weight=3)
        
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
        
        # 더블클릭 이벤트 (기사 읽기)
        self.tv_news.bind("<Double-1>", self.open_link)
        
        # 색상 태그 설정 (상승=빨강, 하락=파랑)
        self.tv_news.tag_configure('pos', foreground='red')
        self.tv_news.tag_configure('neg', foreground='blue')
        self.tv_market.tag_configure('up', foreground='red')
        self.tv_market.tag_configure('down', foreground='blue')
        self.tv_port.tag_configure('up', foreground='red')
        self.tv_port.tag_configure('down', foreground='blue')
        
        # 자동 갱신 타이머 설정 (5분 = 300,000ms)
        self.refresh_interval = 300000
        self.timer_id = None
        
        # 시작하자마자 자동 수집 실행
        self.root.after(500, self.refresh_data)
        
    def open_link(self, event):
        item = self.tv_news.selection()
        if item:
            link = self.tv_news.item(item[0], 'values')[4]
            webbrowser.open(link)
            
    def refresh_data(self):
        # 수동 클릭 시 기존 자동 갱신 타이머 초기화 (중복 방지)
        if self.timer_id is not None:
            self.root.after_cancel(self.timer_id)
            
        self.btn_refresh.config(state='disabled')
        self.lbl_status.config(text="인터넷에서 최신 데이터를 수집 중입니다... (약 5~10초 소요)", foreground="blue")
        
        # 기존 데이터 삭제
        for item in self.tv_market.get_children(): self.tv_market.delete(item)
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
            
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                market_df.to_excel(writer, sheet_name='시장지표', index=False)
                portfolio_df.to_excel(writer, sheet_name='추천포트폴리오', index=False)
                news_df.to_excel(writer, sheet_name='주요뉴스', index=False)
                
            # 메인 스레드(UI)로 결과 전달
            self.root.after(0, self._update_ui, market_df, portfolio_df, news_df, filename)
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("오류", f"데이터 수집 실패:\n{e}"))
            self.root.after(0, lambda: self.lbl_status.config(text="수집 실패", foreground="red"))
            self.root.after(0, lambda: self.btn_refresh.config(state='normal'))
            
    def _update_ui(self, market_df, portfolio_df, news_df, filename):
        # 1. 시장 지표 업데이트
        for _, row in market_df.iterrows():
            tag = 'up' if row['전일비 변동률(%)'] > 0 else ('down' if row['전일비 변동률(%)'] < 0 else '')
            sign = "+" if row['전일비 변동률(%)'] > 0 else ""
            pct_str = f"{sign}{row['전일비 변동률(%)']}%"
            self.tv_market.insert('', 'end', values=(row['지표명'], row['현재가'], pct_str, row['기준일자']), tags=(tag,))
            
        # 1-5. 포트폴리오 업데이트
        for _, row in portfolio_df.iterrows():
            tag = 'up' if row['전일비(%)'] > 0 else ('down' if row['전일비(%)'] < 0 else '')
            sign = "+" if row['전일비(%)'] > 0 else ""
            pct_str = f"{sign}{row['전일비(%)']}%"
            
            signal = row.get('매수 타이밍', '')
            if '적극' in signal or '좋은' in signal:
                tag = 'up'
                
            self.tv_port.insert('', 'end', values=(row['추천 종목'], row['현재가(원/$)'], pct_str, row.get('투자 성향', ''), signal), tags=(tag,))
            
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
