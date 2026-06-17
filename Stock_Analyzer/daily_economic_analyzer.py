import yfinance as yf
import feedparser
import pandas as pd
from datetime import datetime
import os
import re

def get_market_data():
    """주요 글로벌 지수 및 환율 데이터를 수집합니다."""
    print("[1/3] 주요 시장 지표 데이터를 수집 중입니다...")
    tickers = {
        'KOSPI': '^KS11',
        'NASDAQ': '^IXIC',
        'S&P 500': '^GSPC',
        'USD/KRW 환율': 'KRW=X'
    }
    
    data_list = []
    for name, ticker in tickers.items():
        try:
            stock = yf.Ticker(ticker)
            # 최근 2일 데이터를 가져와서 전일 대비 등락 확인
            hist = stock.history(period="5d")
            if len(hist) >= 2:
                current_price = hist['Close'].iloc[-1]
                prev_price = hist['Close'].iloc[-2]
                change_pct = ((current_price - prev_price) / prev_price) * 100
                
                data_list.append({
                    '지표명': name,
                    '현재가': round(current_price, 2),
                    '전일비 변동률(%)': round(change_pct, 2),
                    '기준일자': hist.index[-1].strftime('%Y-%m-%d')
                })
        except Exception as e:
            print(f"{name} 데이터 수집 실패: {e}")
            
    return pd.DataFrame(data_list)

def get_portfolio_data():
    """100만 원 초보자 맞춤형 포트폴리오의 실시간 데이터를 수집합니다."""
    print("초보자 추천 포트폴리오 데이터를 수집 중입니다...")
    portfolio = {
        'TIGER 미국S&P500 (미국전체/안전)': '360750.KS',
        'TIGER 미국나스닥100 (미국기술주)': '133690.KS',
        'TIGER 미국배당다우존스 (매월배당)': '458730.KS',
        'TIGER 미국테크TOP10 (빅테크 집중)': '381170.KS',
        'KODEX 미국반도체MV (AI/반도체)': '390390.KS',
        'TIGER 인도니프티50 (신흥국/고성장)': '453870.KS',
        '삼성전자 (국내 시총 1위)': '005930.KS',
        'SK하이닉스 (HBM/반도체 주도)': '000660.KS',
        '현대차 (국내 자동차 1위/수출)': '005380.KS',
        '기아 (자동차/고배당)': '000270.KS',
        'NAVER (국내 IT/플랫폼 1위)': '035420.KS',
        '삼성바이오로직스 (바이오 1위)': '207940.KS',
        'KB금융 (국내 은행/고배당)': '105560.KS',
        'PHO (미국 글로벌 수자원 ETF)': 'PHO',
        'MOO (글로벌 농업/식량 ETF)': 'MOO',
        '대동 (국내 1위 농기계/스마트팜)': '000490.KS',
        '시노펙스 (국내 수처리/첨단필터)': '025320.KQ',
        'URA (글로벌 우라늄/원자력 ETF)': 'URA',
        'REMX (글로벌 희토류/전략금속 ETF)': 'REMX',
        'LIT (글로벌 리튬/배터리 ETF)': 'LIT',
        'COPX (글로벌 구리 광산 ETF)': 'COPX',
        'LG에너지솔루션 (국내 1위 배터리)': '373220.KS',
        '삼성SDI (프리미엄/전고체 배터리)': '006400.KS',
        '에코프로비엠 (국내 양극재 1위)': '247540.KQ'
    }
    
    data_list = []
    for name, ticker in portfolio.items():
        try:
            stock = yf.Ticker(ticker)
            hist = stock.history(period="3mo") # 이평선 및 RSI 계산을 위해 3개월치 데이터 가져오기
            if len(hist) >= 20:
                current_price = hist['Close'].iloc[-1]
                prev_price = hist['Close'].iloc[-2]
                change_pct = ((current_price - prev_price) / prev_price) * 100
                
                # 티커에 따라 달러/원화 포맷팅 구분
                is_korea = '.KS' in ticker or '.KQ' in ticker
                price_str = f"{int(current_price):,}원" if is_korea else f"${current_price:.2f}"
                
                # 20일 이동평균선 (20MA) 계산
                ma20 = hist['Close'].rolling(window=20).mean().iloc[-1]
                
                # 14일 RSI 계산
                delta = hist['Close'].diff()
                gain = (delta.where(delta > 0, 0)).rolling(window=14).mean()
                loss = (-delta.where(delta < 0, 0)).rolling(window=14).mean()
                rs = gain / loss
                rsi = 100 - (100 / (1 + rs))
                current_rsi = rsi.iloc[-1]
                
                # 최근 20일 최고가 대비 하락률 계산 (트레일링 스탑 논리)
                high_20d = hist['Close'].rolling(window=20).max().iloc[-1]
                drawdown_from_high = ((current_price - high_20d) / high_20d) * 100
                
                # PBR 및 배당수익률 가져오기 (가치주 필터링용)
                info = stock.info
                pbr = info.get('priceToBook')
                div_yield = info.get('dividendYield')
                
                pbr_value = pbr if pbr is not None else 999
                div_yield_value = div_yield if div_yield is not None else 0
                
                is_value_stock = (pbr_value < 1.0) and (div_yield_value >= 0.04) and (current_rsi <= 40)
                
                # AI 매수/매도 타이밍 신호 생성 (모든 경우의 수 반영)
                if is_value_stock:
                    trade_signal = "👑 숨은 진주 (초저평가 매수)"
                elif pd.isna(current_rsi):
                    trade_signal = "데이터 부족"
                elif current_rsi >= 75 and current_price < ma20:
                    trade_signal = "🚨 강력 매도 (과열 후 이탈)"
                elif drawdown_from_high <= -5.0 and current_price < ma20:
                    trade_signal = "✂️ 부분 매도 (고점대비 5%하락)"
                elif current_rsi >= 70:
                    trade_signal = "⚠️ 단기 고점 (매도 준비/관망)"
                elif current_rsi <= 35:
                    trade_signal = "🔥 적극 매수 (RSI 바닥)"
                elif current_price < ma20:
                    trade_signal = "👍 좋은 기회 (20일선 아래)"
                else:
                    trade_signal = "✅ 보유 및 분할 매수"
                
                # 투자 성향 분류
                if 'S&P500' in name: style = "코어(필수)"
                elif '나스닥' in name or 'TOP10' in name or '반도체' in name or '하이닉스' in name or 'NAVER' in name or '바이오' in name: style = "성장(공격)"
                elif '배당' in name or 'KB' in name or '기아' in name or '현대차' in name: style = "배당/가치(안전)"
                elif '인도' in name: style = "신흥국(알파)"
                elif 'PHO' in name or 'MOO' in name or '대동' in name or '시노펙스' in name: style = "기후/식량(방어)"
                elif 'URA' in name or 'REMX' in name or 'LIT' in name or 'COPX' in name or '에너지솔루션' in name or 'SDI' in name or '에코프로' in name: style = "미래자원/배터리"
                else: style = "국내 우량주"
                
                data_list.append({
                    '추천 종목': name,
                    '현재가(원/$)': price_str,
                    '전일비(%)': round(change_pct, 2),
                    '투자 성향': style,
                    'AI 매매 시그널': trade_signal
                })
        except Exception as e:
            pass
            
    return pd.DataFrame(data_list)

def analyze_sentiment(text):
    """뉴스 제목을 기반으로 간단한 긍정/부정 심리를 분석합니다."""
    positive_words = ['상승', '돌파', '호조', '기대', '최고', '급등', '안정', '회복', '수혜', '흑자']
    negative_words = ['하락', '폭락', '우려', '침체', '최저', '위기', '불안', '쇼크', '적자', '둔화']
    
    pos_count = sum(1 for word in positive_words if word in text)
    neg_count = sum(1 for word in negative_words if word in text)
    
    if pos_count > neg_count:
        return '긍정적 (호재)'
    elif neg_count > pos_count:
        return '부정적 (악재)'
    else:
        return '중립'

def get_economic_news():
    """구글 뉴스 RSS를 통해 최신 경제/주식 뉴스를 수집합니다."""
    print("[2/3] 최신 경제 및 주식 뉴스를 수집 중입니다...")
    # 구글 뉴스 RSS (한국어, 경제/주식/금리 관련 검색)
    rss_url = "https://news.google.com/rss/search?q=경제+OR+증시+OR+금리+OR+주식&hl=ko&gl=KR&ceid=KR:ko"
    feed = feedparser.parse(rss_url)
    
    news_list = []
    # 상위 20개 기사만 추출
    for entry in feed.entries[:20]:
        # 제목에서 언론사명 분리 (보통 "제목 - 언론사" 형태)
        title_parts = entry.title.rsplit(' - ', 1)
        title = title_parts[0]
        publisher = title_parts[1] if len(title_parts) > 1 else '알 수 없음'
        
        sentiment = analyze_sentiment(title)
        
        news_list.append({
            '발행일시': entry.published,
            '언론사': publisher,
            '기사 제목': title,
            '시장 심리(분석)': sentiment,
            '링크': entry.link
        })
        
    return pd.DataFrame(news_list)

def generate_daily_report():
    print("=== 나만의 경제 뉴스 분석 비서 시작 ===")
    
    market_df = get_market_data()
    portfolio_df = get_portfolio_data()
    news_df = get_economic_news()
    
    print("[3/3] 엑셀 리포트를 생성 중입니다...")
    today_str = datetime.now().strftime("%Y%m%d")
    
    # 바탕화면에 저장할 경우 경로 설정 (또는 현재 폴더)
    current_dir = os.path.dirname(os.path.abspath(__file__))
    filename = os.path.join(current_dir, f"Daily_Economic_Report_{today_str}.xlsx")
    
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        market_df.to_excel(writer, sheet_name='시장지표', index=False)
        portfolio_df.to_excel(writer, sheet_name='추천포트폴리오', index=False)
        news_df.to_excel(writer, sheet_name='주요뉴스_및_심리분석', index=False)
        
        # 열 너비 자동 조절 로직
        for sheetname in writer.sheets:
            worksheet = writer.sheets[sheetname]
            for col in worksheet.columns:
                max_length = 0
                column = col[0].column_letter # Get the column name
                for cell in col:
                    try: # Necessary to avoid error on empty cells
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = (max_length + 2)
                # 최대 너비 제한 (기사 제목 등이 너무 길어지는 것 방지)
                if adjusted_width > 80:
                    adjusted_width = 80
                worksheet.column_dimensions[column].width = adjusted_width

    print(f"\n[완료] 데일리 리포트가 성공적으로 생성되었습니다.")
    print(f"[위치] 파일 저장 경로: {filename}")
    
    # 간단 브리핑 출력
    print("\n[오늘의 주요 시장 요약]")
    for _, row in market_df.iterrows():
        sign = "+" if row['전일비 변동률(%)'] > 0 else ""
        print(f" - {row['지표명']}: {row['현재가']} ({sign}{row['전일비 변동률(%)']}%)")

    # 시장 심리 요약
    pos_news = len(news_df[news_df['시장 심리(분석)'] == '긍정적 (호재)'])
    neg_news = len(news_df[news_df['시장 심리(분석)'] == '부정적 (악재)'])
    print(f"\n[오늘의 뉴스 분위기]")
    print(f"수집된 주요 기사 20개 중 긍정 기사 {pos_news}개, 부정 기사 {neg_news}개 입니다.")

if __name__ == "__main__":
    generate_daily_report()
