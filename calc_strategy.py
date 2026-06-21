import yfinance as yf
import pandas as pd

PORTFOLIO = {
    '삼성전자': '005930.KS',
    '기아': '000270.KS',
    'KB금융': '105560.KS',
    'NAVER': '035420.KS',
    '현대차': '005380.KS',
    '삼성SDI': '006400.KS',
    '에코프로비엠': '247540.KQ',
    'SK하이닉스': '000660.KS'
}

results = []
for name, ticker in PORTFOLIO.items():
    try:
        stock = yf.Ticker(ticker)
        hist = stock.history(period='3mo')
        if len(hist) < 20: continue
        
        current_price = hist['Close'].iloc[-1]
        recent_20 = hist.tail(20)
        low_20d = recent_20['Low'].min()
        high_20d = recent_20['High'].max()
        
        support = low_20d
        resistance = high_20d * 0.95 
        
        distance_to_support = (current_price - support) / support
        profit_per_share = resistance - current_price
        
        if profit_per_share > 0:
            shares_needed = int(10000 / profit_per_share) + 1
            capital_needed = shares_needed * current_price
            
            results.append({
                'name': name,
                'current': current_price,
                'support': support,
                'resistance': resistance,
                'dist': distance_to_support,
                'profit_per_share': profit_per_share,
                'shares_needed': shares_needed,
                'capital_needed': capital_needed
            })
    except Exception:
        pass

results.sort(key=lambda x: x['dist'])

print('=== 추천 스윙 트레이딩 전략 ===')
for r in results[:3]:
    print(f"{r['name']}: 현재가 {int(r['current']):,}원 (지지선 {int(r['support']):,}원 근접)")
    print(f"  - 매수 시점: 현재가 ~ {int(r['support']):,}원 사이 분할 매수")
    print(f"  - 매도 시점: {int(r['resistance']):,}원 도달 시 전량 매도 (1주당 {int(r['profit_per_share']):,}원 기대)")
    print(f"  - 1만원 목표 달성법: {r['shares_needed']}주 매수 (약 {int(r['capital_needed']):,}원 필요)")
