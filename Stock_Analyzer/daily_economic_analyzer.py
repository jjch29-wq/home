import yfinance as yf
import feedparser
import pandas as pd
from datetime import datetime
import os
import re
import requests

def get_market_data():
    """주요 글로벌 지수 및 환율 데이터를 수집합니다."""
    print("[1/3] 주요 시장 지표 데이터를 수집 중입니다...")
    tickers = {
        'KOSPI': '^KS11',
        'NASDAQ': '^IXIC',
        'S&P 500': '^GSPC',
        'USD/KRW 환율': 'KRW=X',
        'WTI 원유(유가)': 'CL=F'
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

def get_hot_trends(titles_list):
    """뉴스 제목들을 분석하여 현재 유행하는 테마와 관련주를 예측합니다."""
    themes = {
        "K-푸드/식품 (수출/소비)": {
            "keywords": ["식품", "K푸드", "수출", "라면", "삼양", "불닭", "냉동김밥", "과자", "빙과", "먹거리", "외식"],
            "stocks": ["삼양식품", "CJ제일제당", "대상", "빙그레", "농심"]
        },
        "AI/반도체 (기술)": {
            "keywords": ["AI", "반도체", "엔비디아", "HBM", "인공지능", "데이터센터", "TSMC", "오픈AI"],
            "stocks": ["SK하이닉스", "한미반도체", "이수페타시스", "삼성전자"]
        },
        "K-뷰티/화장품 (트렌드)": {
            "keywords": ["화장품", "K뷰티", "올리브영", "아마존", "스킨케어", "미용", "수출대박", "인디브랜드"],
            "stocks": ["실리콘투", "아모레퍼시픽", "클리오", "코스맥스", "브이티"]
        },
        "원전/전력기기 (인프라)": {
            "keywords": ["원전", "전력", "변압기", "송전", "체코", "SMR", "에너지", "그리드"],
            "stocks": ["LS일렉트릭", "HD현대일렉트릭", "두산에너빌리티", "효성중공업"]
        },
        "엔터/콘텐츠 (문화소비)": {
            "keywords": ["드라마", "넷플릭스", "K팝", "아이돌", "웹툰", "콘텐츠", "영화", "오징어게임", "BTS", "구독"],
            "stocks": ["스튜디오드래곤", "하이브", "JYP Ent.", "CJ ENM", "네이버"]
        },
        "바이오/제약 (건강)": {
            "keywords": ["바이오", "비만", "FDA", "신약", "임상", "제약", "GLP-1", "항암"],
            "stocks": ["유한양행", "한미약품", "삼성바이오로직스", "알테오젠"]
        },
        "e커머스/온라인쇼핑 (소비트렌드)": {
            "keywords": ["쿠팡", "쇼핑", "e커머스", "온라인", "직구", "알리", "테무", "배송", "택배", "결제액", "소비진작"],
            "stocks": ["네이버", "카카오", "CJ대한통운", "신세계", "이마트"]
        },
        "여행/항공 (보복소비)": {
            "keywords": ["여행", "항공", "관광", "출국", "공항", "여객", "휴가", "보복소비", "환율", "호텔"],
            "stocks": ["대한항공", "진에어", "제주항공", "하나투어", "모두투어", "호텔신라"]
        },
        "가전/디바이스 (IT소비)": {
            "keywords": ["스마트폰", "가전", "TV", "갤럭시", "아이폰", "판매량", "웨어러블", "노트북", "교체주기"],
            "stocks": ["삼성전자", "LG전자", "LG디스플레이", "삼성전기"]
        },
        "결제/핀테크 (소비지표)": {
            "keywords": ["결제", "페이", "카드", "토스", "소비", "지출", "핀테크", "포인트", "할인"],
            "stocks": ["카카오페이", "네이버페이", "KG이니시스", "NHN KCP"]
        },
        "정유/에너지 (원자재)": {
            "keywords": ["정유", "유가", "기름", "휘발유", "원유", "WTI", "에너지", "석유"],
            "stocks": ["S-Oil", "SK이노베이션", "GS"]
        }
    }
    
    theme_scores = {theme: 0 for theme in themes.keys()}
    found_keywords = {theme: set() for theme in themes.keys()}
    
    # 계절(월)에 따른 테마 가중치 부여
    current_month = datetime.now().month
    seasonal_themes = []
    if current_month in [3, 4, 5]: seasonal_themes = ["여행/항공 (보복소비)", "K-뷰티/화장품 (트렌드)"]
    elif current_month in [6, 7, 8]: seasonal_themes = ["여행/항공 (보복소비)", "원전/전력기기 (인프라)", "가전/디바이스 (IT소비)"]
    elif current_month in [9, 10, 11]: seasonal_themes = ["엔터/콘텐츠 (문화소비)", "e커머스/온라인쇼핑 (소비트렌드)"]
    elif current_month in [12, 1, 2]: seasonal_themes = ["바이오/제약 (건강)", "e커머스/온라인쇼핑 (소비트렌드)"]
    
    for theme in themes.keys():
        if theme in seasonal_themes:
            theme_scores[theme] += 15 # 계절적 성수기 프리미엄
            found_keywords[theme].add("🌱성수기/계절수혜")
            
    # 유가 변동률 확인
    try:
        oil = yf.Ticker("CL=F").history(period="5d")
        if len(oil) >= 2:
            oil_pct = ((oil['Close'].iloc[-1] - oil['Close'].iloc[-2]) / oil['Close'].iloc[-2]) * 100
            oil_price = oil['Close'].iloc[-1]
            if oil_price > 80 or oil_pct >= 1.5:
                theme_scores["정유/에너지 (원자재)"] += 20
                found_keywords["정유/에너지 (원자재)"].add("🔥유가상승수혜")
                theme_scores["여행/항공 (보복소비)"] -= 15
                found_keywords["여행/항공 (보복소비)"].add("⚠️유가부담")
            elif oil_price < 70 or oil_pct <= -1.5:
                theme_scores["여행/항공 (보복소비)"] += 20
                found_keywords["여행/항공 (보복소비)"].add("✈️유가하락수혜")
    except: pass
    
    # 판매, 소비행위, 실적 상승을 의미하는 키워드 (가중치 부여용)
    demand_keywords = ["판매", "수요", "품절", "매출", "수출", "실적", "대박", "품귀", "오픈런", "돌풍", "주문", "흑자", "품귀현상", "결제액", "소비", "지출", "구매력", "영업이익"]
    
    for title in titles_list:
        title_upper = title.upper()
        has_demand = any(dk in title_upper for dk in demand_keywords)
        
        for theme, data in themes.items():
            for kw in data['keywords']:
                if kw.upper() in title_upper:
                    # 소비자 구매/실적 관련 단어와 함께 언급되면 가중치 3배 부여!
                    score_to_add = 30 if has_demand else 10
                    theme_scores[theme] += score_to_add
                    found_keywords[theme].add(kw)
                    if has_demand:
                        found_keywords[theme].add("🛍️소비/매출폭발")
                
    sorted_themes = sorted(theme_scores.items(), key=lambda x: x[1], reverse=True)
    
    results = []
    for theme, score in sorted_themes:
        if score > 0:
            results.append({
                '테마명': theme,
                '발견된 키워드': ", ".join(list(found_keywords[theme])),
                '관련 대장주': ", ".join(themes[theme]['stocks']),
                '트렌드 지수(관심도)': min(score + 30, 99) # 점수에 기본값 30을 더해 가시성 확보
            })
            
    if not results:
        results.append({
            '테마명': '관망 장세 (특징 테마 없음)',
            '발견된 키워드': '-',
            '관련 대장주': '우량주 위주 방어적 투자 권장',
            '트렌드 지수(관심도)': 0
        })
        
    return pd.DataFrame(results[:3])

PORTFOLIO = {
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
    '에코프로비엠 (국내 양극재 1위)': '247540.KQ',
    'KODEX 미국달러선물 (환율/안전)': '261240.KS',
    'TIGER 미국채10년선물 (안전자산/금리인하)': '305080.KS',
    '맥쿼리인프라 (국내 인프라/고배당)': '088980.KS',
    '기업은행 (국내 국책은행/고배당)': '024110.KS',
    'KT&G (국내 담배/경기방어/고배당)': '033780.KS',
    'SK텔레콤 (국내 통신/고배당)': '017670.KS',
    
    # --- 수익 극대화 고성장/주도주 (추가) ---
    '한미반도체 (AI/반도체 장비)': '042700.KS',
    '삼양식품 (K-푸드/수출 대장)': '003230.KS',
    '알테오젠 (바이오/신약 기술)': '196170.KQ',
    'HD현대일렉트릭 (전력기기/슈퍼사이클)': '267260.KS',
    '실리콘투 (K-뷰티/수출 대장)': '257720.KQ',
    'LS일렉트릭 (전력/스마트그리드)': '010120.KS'
}

# --- 실제 보유 주식 ---
OWNED_STOCKS = {
    '000270.KS': 1, # 기아
    '105560.KS': 1, # KB금융
    'MOO': 1        # MOO
}

# --- 텔레그램 봇 설정 ---
TELEGRAM_TOKEN = "7830526088:AAEqyb5l6utOJAre-nngJ-529XHG8K-sdSQ"
TELEGRAM_CHAT_ID = "8391233271" # 여기에 숫자 챗 ID를 넣으세요 (예: "123456789")

def send_telegram_message(text):
    """텔레그램으로 텍스트 메시지를 전송합니다."""
    if not TELEGRAM_TOKEN or not TELEGRAM_CHAT_ID:
        return
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    payload = {
        'chat_id': TELEGRAM_CHAT_ID,
        'text': text
    }
    try:
        requests.post(url, json=payload, timeout=5)
    except Exception as e:
        print(f"텔레그램 전송 실패: {e}")

def send_telegram_photo(photo_path, caption=None):
    """텔레그램으로 사진을 전송합니다."""
    if not TELEGRAM_TOKEN or not TELEGRAM_CHAT_ID:
        return
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendPhoto"
    try:
        with open(photo_path, 'rb') as photo:
            files = {'photo': photo}
            data = {'chat_id': TELEGRAM_CHAT_ID}
            if caption:
                data['caption'] = caption
            requests.post(url, data=data, files=files, timeout=10)
    except Exception as e:
        print(f"텔레그램 사진 전송 실패: {e}")

def generate_holdings_graph(portfolio_df):
    """보유 주식의 최근 3개월 수익률 추이 그래프를 생성합니다."""
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    
    # 한글 폰트 설정
    plt.rcParams['font.family'] = 'Malgun Gothic'
    plt.rcParams['axes.unicode_minus'] = False
    
    holdings = []
    for _, row in portfolio_df.iterrows():
        qty = str(row.get('보유', '-'))
        if '주' in qty:
            name = row['추천 종목'].split('(')[0].strip()
            clean_name = name.replace('🔥 ', '').replace(' (사용자 추가)', '').strip()
            
            ticker = PORTFOLIO.get(clean_name)
            if not ticker:
                for full_name, t in PORTFOLIO.items():
                    if full_name.split('(')[0].strip() == clean_name:
                        ticker = t
                        break
            if not ticker:
                custom = load_custom_stocks()
                if clean_name in custom:
                    ticker = custom[clean_name]['ticker']
                    
            if ticker:
                holdings.append((name, ticker))
                
    if not holdings:
        return None
        
    plt.figure(figsize=(10, 6))
    has_data = False
    
    for name, ticker in holdings:
        try:
            hist = yf.Ticker(ticker).history(period='3mo')
            if not hist.empty and len(hist) > 0:
                base_price = hist['Close'].iloc[0]
                normalized = (hist['Close'] / base_price - 1) * 100
                plt.plot(hist.index, normalized, label=name)
                has_data = True
        except Exception as e:
            pass
            
    if not has_data:
        plt.close()
        return None
        
    plt.title("내 보유 주식 최근 3개월 수익률 추이 (%)", fontsize=15, fontweight='bold')
    plt.xlabel("날짜", fontsize=12)
    plt.ylabel("수익률 (%)", fontsize=12)
    plt.grid(True, linestyle='--', alpha=0.7)
    plt.axhline(y=0, color='r', linestyle='-', alpha=0.3)
    
    # x축 레이블이 겹치지 않게 회전
    plt.xticks(rotation=45)
    
    plt.legend(loc='upper left', bbox_to_anchor=(1, 1))
    plt.tight_layout()
    
    import os
    from datetime import datetime
    current_dir = os.path.dirname(os.path.abspath(__file__))
    filename = os.path.join(current_dir, f"holdings_graph_{datetime.now().strftime('%Y%m%d%H%M%S')}.png")
    plt.savefig(filename, dpi=100)
    plt.close()
    
    return filename

def send_briefing_to_telegram(market_df, portfolio_df, pos_news, neg_news, trend_df=None):
    """데이터프레임을 요약하여 텔레그램으로 전송합니다."""
    if not TELEGRAM_TOKEN or not TELEGRAM_CHAT_ID:
        return
        
    msg = "📊 [오늘의 나만의 경제 비서 브리핑]\n\n"
    
    total_profit = int(portfolio_df['일간 변동금액'].sum())
    msg += f"🎯 [오늘의 수익 목표 (5천 원)]\n"
    if total_profit >= 5000:
        msg += f"✅ 달성 완료! (오늘 수익: +{total_profit:,}원)\n\n"
    elif total_profit > 0:
        msg += f"🏃 진행 중 (오늘 수익: +{total_profit:,}원)\n\n"
    else:
        msg += f"📉 내일을 기약해요 (오늘 손익: {total_profit:,}원)\n\n"
    
    msg += "📈 [주요 시장 지표]\n"
    for _, row in market_df.iterrows():
        sign = "+" if row['전일비 변동률(%)'] > 0 else ""
        msg += f"• {row['지표명']}: {row['현재가']} ({sign}{row['전일비 변동률(%)']}%)\n"
        
    msg += "\n🎯 [5천 원 스윙 추천 TOP 3]\n"
    swing_results = []
    for _, row in portfolio_df.iterrows():
        name = row['추천 종목'].split('(')[0].strip()
        try:
            curr = float(str(row['현재가(원/$)']).replace(',','').replace('원','').replace('$',''))
            support = float(str(row.get('예상 저점(지지선)', '0')).replace(',','').replace('원','').replace('$',''))
            resistance = float(str(row.get('예상 고점(저항선)', '0')).replace(',','').replace('원','').replace('$',''))
            if support == 0 or resistance == 0 or curr == 0: continue
            
            resistance = resistance * 0.95
            profit = resistance - curr
            if profit > 0:
                dist = (curr - support) / support
                shares = int(5000 / profit) + 1
                capital = shares * curr
                is_korea = '원' in str(row['현재가(원/$)'])
                swing_results.append({'name': name, 'support': support, 'resistance': resistance, 'dist': dist, 'shares': shares, 'capital': capital, 'is_korea': is_korea})
        except: pass
        
    swing_results.sort(key=lambda x: x['dist'])
    for i, r in enumerate(swing_results[:3], 1):
        unit = "원" if r['is_korea'] else "$"
        fmt = lambda x: f"{int(x):,}{unit}" if r['is_korea'] else f"${x:.2f}"
        msg += f"{i}. {r['name']} ({r['shares']}주 / {fmt(r['capital'])})\n"
        msg += f"  👉 {fmt(r['support'])} 매수 ➡️ {fmt(r['resistance'])} 매도\n"
        
    if 'trend_df' in globals() or True: # trend_df는 ui에서 넘겨받지 않고 그냥 내부 변수로 사용불가하니 인자로 추가해야함.
        pass # UI 단에서 넘기도록 수정해야 함
    msg_holdings = "💼 [나의 보유 주식 상세 진단 및 시뮬레이션]\n"
    has_holdings = False
    for _, row in portfolio_df.iterrows():
        qty = str(row.get('보유', '-'))
        if '주' in qty:
            has_holdings = True
            name = row['추천 종목'].split('(')[0].strip()
            try:
                curr = float(str(row['현재가(원/$)']).replace(',','').replace('원','').replace('$',''))
                support = float(str(row.get('예상 저점(지지선)', '0')).replace(',','').replace('원','').replace('$',''))
                resistance = float(str(row.get('예상 고점(저항선)', '0')).replace(',','').replace('원','').replace('$',''))
                if support == 0 or resistance == 0 or curr == 0: continue
                
                pos_pct = (curr - support) / (resistance - support) * 100 if resistance > support else 50
                
                avg_price = float(row.get('매수 단가', 0))
                is_kor = '원' in str(row['현재가(원/$)'])
                fmt2 = lambda x: f"{int(x):,}원" if is_kor else f"${x:.2f}"
                
                if avg_price > 0:
                    roi = ((curr - avg_price) / avg_price) * 100
                    sign = "+" if roi > 0 else ""
                    msg_holdings += f"• {name} ({qty}) - 현재가: {fmt2(curr)} (매수단가: {fmt2(avg_price)} / 수익률: {sign}{roi:.2f}%)\n"
                else:
                    msg_holdings += f"• {name} ({qty}) - 현재가: {fmt2(curr)}\n"
                if pos_pct <= 20:
                    msg_holdings += f"  • 현재 상태: [완전한 바닥권 (하위 {int(pos_pct)}%)]\n"
                    msg_holdings += f"  • 대책 방안: 절대 매도 금지! 지금 팔면 최하점입니다. {fmt2(support)} 부근 추가 매수 고려\n"
                elif pos_pct <= 40:
                    msg_holdings += f"  • 현재 상태: [바닥 다지기 및 반등 시작 (하위 {int(pos_pct)}%)]\n"
                    msg_holdings += f"  • 대책 방안: 상승 추세 전환 가능성. 보유 유지하며 목표가 {fmt2(resistance)} 대기\n"
                elif pos_pct <= 70:
                    msg_holdings += f"  • 현재 상태: [중간 상승 구간 (상위 {100-int(pos_pct)}%)]\n"
                    msg_holdings += f"  • 대책 방안: 수익 확대 중. 저항선({fmt2(resistance)}) 근접 시 매도 준비\n"
                else:
                    msg_holdings += f"  • 현재 상태: [고점 도달 (상위 {100-int(pos_pct)}%)]\n"
                    msg_holdings += f"  • 대책 방안: 목표가({fmt2(resistance)}) 근접! 욕심을 줄이고 수익 실현(매도) 타이밍\n"
                    
                # 5일 예측 계산
                clean_name = row['추천 종목'].replace('🔥 ', '').replace(' (사용자 추가)', '').split('(')[0].strip()
                ticker = PORTFOLIO.get(clean_name)
                
                # 티커 부분 일치 검색
                if not ticker:
                    for full_name, t in PORTFOLIO.items():
                        if full_name.split('(')[0].strip() == clean_name:
                            ticker = t
                            break
                            
                if not ticker:
                    custom = load_custom_stocks()
                    if clean_name in custom:
                        ticker = custom[clean_name]['ticker']
                        
                if ticker:
                    hist = yf.Ticker(ticker).history(period='1y')
                    if len(hist) >= 20:
                        returns = hist['Close'].pct_change().dropna()
                        mu = returns.mean()
                        sigma = returns.std()
                        
                        import numpy as np
                        from datetime import datetime, timedelta
                        
                        num_sims = 10000
                        days = 5
                        daily_returns = np.random.normal(mu, sigma, (days, num_sims))
                        price_paths = np.zeros((days + 1, num_sims))
                        price_paths[0] = curr
                        
                        for t in range(1, days + 1):
                            price_paths[t] = price_paths[t-1] * (1 + daily_returns[t-1])
                            
                        is_kor = '원' in str(row['현재가(원/$)'])
                        fmt2 = lambda x: f"{int(x):,}" if is_kor else f"{x:.2f}"
                        unit = "원" if is_kor else "$"
                        
                        msg_holdings += f"  📈 [향후 5일 시뮬레이션(단위:{unit})]\n"
                        curr_date = datetime.now()
                        for i in range(1, 6):
                            curr_date += timedelta(days=1)
                            while curr_date.weekday() > 4:
                                curr_date += timedelta(days=1)
                                
                            p_pes = np.percentile(price_paths[i], 10)
                            p_neu = np.percentile(price_paths[i], 50)
                            p_opt = np.percentile(price_paths[i], 90)
                            
                            d_str = curr_date.strftime('%m/%d')
                            msg_holdings += f"    • {d_str} ➡️ 하락 {fmt2(p_pes)} / 중립 {fmt2(p_neu)} / 상승 {fmt2(p_opt)}\n"
            except: pass
            
    if not has_holdings:
        msg_holdings += "• 현재 보유 중인 주식이 없습니다.\n"
        
    if trend_df is not None and not trend_df.empty:
        msg += "\n🔥 [AI 유행 예측 및 수혜주 발굴]\n"
        for i, row in trend_df.iterrows():
            if row['트렌드 지수(관심도)'] > 0:
                msg += f"[{i+1}위] {row['테마명']} (지수: {row['트렌드 지수(관심도)']})\n"
                msg += f"  👉 힌트 키워드: {row['발견된 키워드']}\n"
                msg += f"  👉 관련 수혜주: {row['관련 대장주']}\n"
        
    msg += "\n🔔 [AI 핵심 매매 시그널]\n"
    signal_count = 0
    for _, row in portfolio_df.iterrows():
        signal = row['AI 매매 시그널']
        # '보유' 제외하고 적극적인 매수/매도 시그널만 필터링
        if '진주' in signal or '적극 매수' in signal or '좋은 기회' in signal or '매도' in signal:
            short_name = row['추천 종목'].split('(')[0].strip()
            msg += f"• {short_name}: {signal}\n"
            signal_count += 1
            
    if signal_count == 0:
        msg += "• 오늘은 주목할 만한 매수/매도 시그널이 없습니다.\n"
        
    msg += f"\n📰 [오늘의 시장 심리]\n긍정 뉴스 {pos_news}개 / 부정 뉴스 {neg_news}개\n"
    
    send_telegram_message(msg)
    
    # 보유 주식 그래프 생성 및 전송
    try:
        graph_path = generate_holdings_graph(portfolio_df)
        if graph_path:
            caption = "📈 [나의 보유 주식 최근 3개월 수익률 추이]"
            send_telegram_photo(graph_path, caption)
            import os
            try:
                os.remove(graph_path)
            except:
                pass
    except Exception as e:
        print(f"그래프 전송 오류: {e}")
        
    # 상세 설명(진단 내용)을 그래프 바로 밑에 전송
    if msg_holdings:
        send_telegram_message(msg_holdings)

import json

CUSTOM_PORT_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "custom_portfolio.json")

def load_custom_stocks():
    if os.path.exists(CUSTOM_PORT_FILE):
        try:
            with open(CUSTOM_PORT_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            pass
    return {}

def save_custom_stock(name, ticker, qty=0, avg_price=0):
    custom = load_custom_stocks()
    custom[name] = {"ticker": ticker, "qty": qty, "avg_price": avg_price}
    with open(CUSTOM_PORT_FILE, 'w', encoding='utf-8') as f:
        json.dump(custom, f, ensure_ascii=False, indent=2)

def delete_custom_stock(stock_name):
    custom = load_custom_stocks()
    
    # 기본 포트폴리오 종목인지 확인
    is_default = False
    default_ticker = None
    for full_name, ticker in PORTFOLIO.items():
        clean_name = full_name.split('(')[0].strip()
        if clean_name == stock_name:
            is_default = True
            default_ticker = ticker
            break
            
    if is_default:
        # 기본 종목이면 수량을 0으로 덮어씀
        custom[stock_name] = {"ticker": default_ticker, "qty": 0}
    else:
        # 순수 사용자 추가 종목이면 완전 삭제 (키가 다를 수도 있으므로 이름 부분 매치)
        keys_to_delete = []
        for k in custom.keys():
            if stock_name in k:
                keys_to_delete.append(k)
        for k in keys_to_delete:
            del custom[k]
            
    with open(CUSTOM_PORT_FILE, 'w', encoding='utf-8') as f:
        json.dump(custom, f, ensure_ascii=False, indent=2)

def get_naver_realtime_price(ticker):
    """네이버 증권에서 실시간 현재가를 스크래핑합니다 (bs4 의존성 제거)."""
    try:
        import requests
        import re
        code = ticker.split('.')[0]
        url = f"https://finance.naver.com/item/main.naver?code={code}"
        headers = {'User-Agent': 'Mozilla/5.0'}
        # 타임아웃을 짧게 주어 UI 멈춤 현상 최소화
        res = requests.get(url, headers=headers, timeout=2.0)
        
        # 정규식으로 <p class="no_today"> 안의 맹인용 텍스트(현재가) 추출
        match = re.search(r'<p class="no_today"[^>]*>.*?<span class="blind">([\d,]+)</span>', res.text, re.DOTALL)
        if match:
            return float(match.group(1).replace(',', ''))
    except Exception as e:
        pass
    return None

def get_portfolio_data(market_df=None):
    """AI 주도주 수익 창출 포트폴리오의 실시간 데이터를 수집합니다."""
    print("수익 창출 포트폴리오 데이터를 수집 중입니다...")
    
    current_portfolio = PORTFOLIO.copy()
    current_owned = OWNED_STOCKS.copy()
    current_avg_price = {}
    
    # KOSPI 및 NASDAQ 전일비 변동률(%) 파싱
    kospi_change = 0
    nasdaq_change = 0
    if market_df is not None and not market_df.empty:
        try:
            k_row = market_df[market_df['지표명'] == 'KOSPI']
            if not k_row.empty:
                kospi_change = k_row['전일비 변동률(%)'].iloc[0]
                
            n_row = market_df[market_df['지표명'] == 'NASDAQ']
            if not n_row.empty:
                nasdaq_change = n_row['전일비 변동률(%)'].iloc[0]
        except Exception:
            pass
    
    custom_stocks = load_custom_stocks()
    for name, info in custom_stocks.items():
        # 기존 포트폴리오에 동일한 종목코드(티커)가 있는지 확인
        existing_name = None
        for p_name, p_ticker in PORTFOLIO.items():
            if p_ticker.split('.')[0] == info['ticker'].split('.')[0]:
                existing_name = p_name
                break
                
        if existing_name:
            # 이미 있는 종목이면 수량 및 매수 단가 업데이트
            if info.get('qty', 0) >= 0:
                current_owned[existing_name] = info['qty']
                current_owned[PORTFOLIO[existing_name]] = info['qty'] 
                if info.get('avg_price', 0) > 0:
                    current_avg_price[existing_name] = info['avg_price']
                    current_avg_price[PORTFOLIO[existing_name]] = info['avg_price']
        else:
            # 새로운 종목이면 (사용자 추가) 꼬리표 달고 새로 등록
            display_name = f"{name} (사용자 추가)"
            current_portfolio[display_name] = info['ticker']
            if info.get('qty', 0) > 0:
                current_owned[display_name] = info['qty']
            if info.get('avg_price', 0) > 0:
                current_avg_price[display_name] = info['avg_price']
            
    try:
        vix = yf.Ticker("^VIX").history(period="1d")['Close'].iloc[-1]
        if vix >= 25: market_sentiment = "extreme_fear"
        elif vix >= 20: market_sentiment = "fear"
        elif vix <= 15: market_sentiment = "extreme_greed"
        else: market_sentiment = "normal"
    except:
        market_sentiment = "normal"
        
    data_list = []
    for name, ticker in current_portfolio.items():
        try:
            stock = yf.Ticker(ticker)
            hist = stock.history(period="3mo") # 이평선 및 RSI 계산을 위해 3개월치 데이터 가져오기
            
            # 사용자 추가 시 코스피(.KS)/코스닥(.KQ) 혼동에 대한 자동 대비 (데이터 미존재 시 상호 변환 재시도)
            if len(hist) < 20 and (ticker.endswith('.KS') or ticker.endswith('.KQ')):
                fallback_ticker = ticker.replace('.KS', '.KQ') if ticker.endswith('.KS') else ticker.replace('.KQ', '.KS')
                stock = yf.Ticker(fallback_ticker)
                hist = stock.history(period="3mo")
                ticker = fallback_ticker # 이후 변수들에서 정상 동작하도록 업데이트
                
            if len(hist) >= 20:
                current_price = hist['Close'].iloc[-1]
                prev_price = hist['Close'].iloc[-2]
                
                is_korea = '.KS' in ticker or '.KQ' in ticker
                
                # 한국 주식일 경우 네이버 실시간 현재가 우선 적용
                if is_korea:
                    realtime_price = get_naver_realtime_price(ticker)
                    if realtime_price is not None:
                        current_price = realtime_price
                        
                change_pct = ((current_price - prev_price) / prev_price) * 100
                
                # 티커에 따라 달러/원화 포맷팅 구분
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
                
                # 최근 20일 최고/최저가 계산 (예상 오르내림 폭)
                high_20d = hist['Close'].rolling(window=20).max().iloc[-1]
                low_20d = hist['Close'].rolling(window=20).min().iloc[-1]
                drawdown_from_high = ((current_price - high_20d) / high_20d) * 100
                rise_from_low = ((current_price - low_20d) / low_20d) * 100 if low_20d > 0 else 0
                
                # PBR 및 배당수익률 가져오기 (가치주 필터링용)
                info = stock.info
                pbr = info.get('priceToBook')
                div_yield = info.get('dividendYield')
                
                pbr_value = pbr if pbr is not None else 999
                div_yield_value = div_yield if div_yield is not None else 0
                
                is_value_stock = (pbr_value < 1.0) and (div_yield_value >= 0.04) and (current_rsi <= 40)
                
                # 거래량 증감 계산 (20일 평균 대비)
                try:
                    current_vol = hist['Volume'].iloc[-1]
                    avg_vol_20d = hist['Volume'].rolling(window=20).mean().iloc[-2]
                    vol_surge = (current_vol / avg_vol_20d) * 100 if avg_vol_20d > 0 else 0
                    vol_surge_str = f"{int(vol_surge)}%" if vol_surge > 0 else "-"
                except:
                    vol_surge = 0
                    vol_surge_str = "-"
                
                # 시장 폭락/급등 여부 확인 (지수 1.5% 이상 변동 시)
                market_crash = False
                market_boom = False
                if is_korea:
                    if kospi_change <= -1.5: market_crash = True
                    elif kospi_change >= 1.5: market_boom = True
                else:
                    if nasdaq_change <= -1.5: market_crash = True
                    elif nasdaq_change >= 1.5: market_boom = True

                # AI 매수/매도 타이밍 신호 생성 (스윙/평균회귀 전략 중심)
                if vol_surge >= 200 and change_pct > 0:
                    trade_signal = "🚀 거래량 급증 (수급 폭발)"
                elif is_value_stock:
                    trade_signal = "👑 숨은 진주 (초저평가 매수)"
                elif pd.isna(current_rsi):
                    trade_signal = "데이터 부족"
                elif drawdown_from_high <= -15.0 and current_rsi <= 35:
                    trade_signal = "🔥 패닉 셀링 (과대 낙폭/적극 매수)"
                elif drawdown_from_high <= -8.0 and current_price < ma20:
                    trade_signal = "👍 눌림목 매수 (스윙 타점 진입)"
                elif current_price >= high_20d * 0.98:
                    trade_signal = "🚨 전고점 터치 (전량 매도 권장)"
                elif rise_from_low >= 10.0 and current_price > ma20 and current_rsi >= 70:
                    trade_signal = "⚠️ 과열 구간 (분할 매도 준비)"
                elif current_rsi >= 75:
                    trade_signal = "🚨 단기 과매수 (이익 실현)"
                elif current_rsi <= 35:
                    trade_signal = "🔥 적극 매수 (RSI 바닥)"
                elif current_price < ma20:
                    if market_crash:
                        trade_signal = "🛡️ 보수적 접근 (시장 하락, 관망)"
                    elif market_boom:
                        trade_signal = "🚀 시장 상승 편승 (매수 기회)"
                    else:
                        trade_signal = "⏳ 지지선 대기 (조정 중)"
                else:
                    if market_crash:
                        trade_signal = "🛡️ 관망 (시장 변동성 확대)"
                    elif market_boom:
                        trade_signal = "✅ 보유 (시장 상승세)"
                    else:
                        trade_signal = "✅ 보유 및 추세 관망"
                
                # 투자 성향 분류
                if 'S&P500' in name: style = "코어(필수)"
                elif '나스닥' in name or 'TOP10' in name or '반도체' in name or '하이닉스' in name or 'NAVER' in name or '바이오' in name: style = "성장(공격)"
                elif '배당' in name or 'KB' in name or '기아' in name or '현대차' in name: style = "배당/가치(안전)"
                elif '인도' in name: style = "신흥국(알파)"
                elif 'PHO' in name or 'MOO' in name or '대동' in name or '시노펙스' in name: style = "기후/식량(방어)"
                elif 'URA' in name or 'REMX' in name or 'LIT' in name or 'COPX' in name or '에너지솔루션' in name or 'SDI' in name or '에코프로' in name: style = "미래자원/배터리"
                elif '달러' in name or '골드' in name or '미국채' in name: style = "헷징/안전자산"
                else: style = "국내 우량주"
                
                support_price = f"{int(low_20d):,}원" if is_korea else f"${low_20d:.2f}"
                resistance_price = f"{int(high_20d):,}원" if is_korea else f"${high_20d:.2f}"
                target_buy = f"{int(ma20):,}원" if is_korea else f"${ma20:.2f}"
                target_sell = f"{int(high_20d * 0.95):,}원" if is_korea else f"${high_20d * 0.95:.2f}"
                
                qty = current_owned.get(name, OWNED_STOCKS.get(ticker, 0))
                qty_str = f"{qty}주" if qty > 0 else "-"
                avg_price = current_avg_price.get(name, 0)
                
                change_amt = current_price - prev_price
                if is_korea:
                    daily_change_krw = change_amt * qty
                else:
                    daily_change_krw = change_amt * qty * 1380 # 환율 대략 1380원 적용
                
                data_list.append({
                    '추천 종목': name,
                    '보유': qty_str,
                    '매수 단가': avg_price,
                    '일간 변동금액': daily_change_krw,
                    '현재가(원/$)': price_str,
                    '전일비(%)': round(change_pct, 2),
                    '거래량 증감(%)': vol_surge_str,
                    '예상 저점(지지선)': support_price,
                    '예상 고점(저항선)': resistance_price,
                    '목표 매수가': target_buy,
                    '부분 매도가': target_sell,
                    '투자 성향': style,
                    'AI 매매 시그널': trade_signal
                })
        except Exception as e:
            pass
            
    return pd.DataFrame(data_list)

def get_sector_performance(portfolio_df):
    """투자 성향(섹터/테마)별 평균 등락률과 주도 종목을 계산합니다."""
    if portfolio_df.empty:
        return pd.DataFrame()
        
    results = []
    grouped = portfolio_df.groupby('투자 성향')
    
    for sector, group in grouped:
        avg_change = group['전일비(%)'].mean()
        
        # 해당 섹터에서 등락률이 가장 높은 종목 찾기
        top_stock_idx = group['전일비(%)'].idxmax()
        top_stock = group.loc[top_stock_idx, '추천 종목'].split('(')[0].strip()
        top_stock_change = group.loc[top_stock_idx, '전일비(%)']
        
        top_stock_str = f"{top_stock} ({'+' if top_stock_change > 0 else ''}{top_stock_change}%)"
        
        results.append({
            '섹터/테마명': sector,
            '평균 전일비(%)': round(avg_change, 2),
            '주도 종목': top_stock_str
        })
        
    res_df = pd.DataFrame(results)
    res_df = res_df.sort_values('평균 전일비(%)', ascending=False).reset_index(drop=True)
    return res_df

def get_buy_recommendations(portfolio_df):
    """매수 추천 조건에 부합하는 상위 3개 종목을 추출합니다."""
    if portfolio_df.empty:
        return pd.DataFrame()
        
    # 우선순위: 1. 패닉 셀링, 2. 눌림목 매수, 3. 거래량 급증, 4. 숨은 진주, 5. RSI 바닥, 6. 지지선 대기
    priority = {
        '🔥 패닉 셀링 (과대 낙폭/적극 매수)': 1,
        '👍 눌림목 매수 (스윙 타점 진입)': 2,
        '🚀 거래량 급증 (수급 폭발)': 3,
        '👑 숨은 진주 (초저평가 매수)': 4,
        '🔥 적극 매수 (RSI 바닥)': 5,
        '⏳ 지지선 대기 (조정 중)': 6
    }
    
    seen_stocks = set()
    recs = []
    for _, row in portfolio_df.iterrows():
        signal = row.get('AI 매매 시그널', '')
        if signal in priority:
            # 수익률(%) = (예상 고점 - 현재가) / 현재가 * 100
            try:
                curr = float(str(row['현재가(원/$)']).replace(',','').replace('원','').replace('$',''))
                res = float(str(row['예상 고점(저항선)']).replace(',','').replace('원','').replace('$',''))
                expected_profit_pct = ((res * 0.95) - curr) / curr * 100 if curr > 0 else 0
            except:
                expected_profit_pct = 0
                
            if expected_profit_pct > 0:
                stock_name = row['추천 종목'].split('(')[0].strip()
                if stock_name not in seen_stocks:
                    seen_stocks.add(stock_name)
                    recs.append({
                        '추천 종목': stock_name,
                        '현재가': row['현재가(원/$)'],
                        '예상 저점': row['예상 저점(지지선)'],
                        '기대 수익률(%)': round(expected_profit_pct, 1),
                        '시그널': signal,
                        '우선순위': priority[signal]
                    })
                
    if not recs:
        return pd.DataFrame()
        
    rec_df = pd.DataFrame(recs)
    # 1. 시그널 우선순위 높은순, 2. 기대 수익률 높은순 정렬
    rec_df = rec_df.sort_values(['우선순위', '기대 수익률(%)'], ascending=[True, False]).head(3)
    return rec_df

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
    # 구글 뉴스 RSS (경제, 주식 뿐만 아니라 문화, 트렌드, 사회 분야 종합 검색)
    rss_url = "https://news.google.com/rss/search?q=경제+OR+증시+OR+주식+OR+문화+OR+트렌드+OR+테크&hl=ko&gl=KR&ceid=KR:ko"
    feed = feedparser.parse(rss_url)
    
    news_list = []
    # 상위 30개 기사 추출 (더 폭넓은 트렌드 반영을 위해 개수 증가)
    for entry in feed.entries[:30]:
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
    if not news_df.empty and '시장 심리(분석)' in news_df.columns:
        pos_news = len(news_df[news_df['시장 심리(분석)'] == '긍정적 (호재)'])
        neg_news = len(news_df[news_df['시장 심리(분석)'] == '부정적 (악재)'])
        print(f"\n[오늘의 뉴스 분위기]")
        print(f"수집된 주요 기사 {len(news_df)}개 중 긍정 기사 {pos_news}개, 부정 기사 {neg_news}개 입니다.")
    else:
        print("\n[오늘의 뉴스 분위기]\n뉴스 데이터를 불러오지 못했습니다.")

if __name__ == "__main__":
    generate_daily_report()
