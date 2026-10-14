"""작업일보는 선택 행 수와 무관하게 같은 날짜·현장의 전체 기록을 사용한다."""

import pandas as pd


def select_daily_report_records(frame, date, site, loaded_index=None):
    site_col = 'Site' if 'Site' in frame.columns else '현장'
    date_col = 'Date' if 'Date' in frame.columns else '날짜'
    if site_col not in frame.columns or date_col not in frame.columns:
        return frame.iloc[0:0].copy()
    # 불러온 행의 현장은 조회 기준, 입력창 공사명은 출력 표기다.
    if loaded_index in frame.index:
        source = frame.loc[loaded_index]
        source_date = pd.to_datetime(source[date_col], errors='coerce')
        if pd.notna(source_date) and source_date.date() == pd.to_datetime(date).date():
            site = source[site_col]
    dates = pd.to_datetime(frame[date_col], errors='coerce').dt.date
    sites = frame[site_col].fillna('').astype(str).str.strip().str.upper()
    return frame[(dates == pd.to_datetime(date).date()) & (sites == str(site).strip().upper())].copy()
