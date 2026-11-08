"""Read billing records from the same daily ledger used by the central app."""
from pathlib import Path
import re

import pandas as pd


def load_billing_daily_records(app, start='', end='', site='전체'):
    current = getattr(app, 'daily_usage_df', None)
    if isinstance(current, pd.DataFrame):
        frame = current.copy()
    else:
        selected = getattr(app, 'db_path', None)
        path = selected or Path(__file__).resolve().parents[2] / 'data' / 'Material_Inventory.xlsx'
        frame = pd.read_excel(path, sheet_name='DailyUsage')
    aliases = {'날짜': 'Date', '현장': 'Site', '수량': 'Usage'}
    frame.columns = frame.columns.astype(str).str.strip()
    for old, new in aliases.items():
        if old in frame and new not in frame:
            frame = frame.rename(columns={old: new})
    if '검사방법' not in frame:
        raise ValueError('일일 장부에 검사방법 열이 없습니다.')
    methods = frame['검사방법'].fillna('').astype(str).str.strip().str.upper()
    frame = frame.loc[methods.isin(['RT', 'UT', 'PT', 'PAUT', 'MT'])].copy()
    frame['검사방법'] = methods.loc[frame.index]
    if frame.empty:
        return frame
    # Excel columns containing only blanks may be inferred as float64.
    # Film specifications are strings, including when the first value is filled.
    if 'Material' in frame:
        frame['Material'] = frame['Material'].astype(object)
    else:
        frame['Material'] = pd.Series(index=frame.index, dtype=object)
    materials = getattr(app, 'materials_df', None)
    if 'MaterialID' in frame and not isinstance(materials, pd.DataFrame):
        selected = getattr(app, 'db_path', None)
        path = selected or Path(__file__).resolve().parents[2] / 'data' / 'Material_Inventory.xlsx'
        materials = pd.read_excel(path, sheet_name='Materials')
    names = {}
    if isinstance(materials, pd.DataFrame) and {'MaterialID', '품목명'}.issubset(materials.columns):
        def material_id(value):
            return str(value).strip().removesuffix('.0')
        names = {material_id(row['MaterialID']): row['품목명'] for _, row in materials.iterrows()}
    for index, row in frame.loc[frame['검사방법'] == 'RT'].iterrows():
        item_name = names.get(str(row.get('MaterialID', '')).strip().removesuffix('.0'), '')
        for candidate in (row.get('Material', ''), row.get('품목명', ''), item_name):
            match = re.search(r'[x×*]\s*(6|12|17)\s*(?:["″]|$|\[)', str(candidate), re.I)
            if match:
                frame.loc[index, 'Material'] = f'3 1/3 x {match.group(1)}"'
                break
        else:
            if pd.notna(row.get('MaterialID')):
                raise ValueError('RT 필름 규격을 확인할 수 없습니다. 장부의 자재 연결을 확인해 주세요.')
    # Ledger rows mix date-only strings, timestamps and Excel datetime values.
    dates = frame['Date'].map(lambda value: pd.to_datetime(value, errors='coerce'))
    if dates.isna().any():
        raise ValueError('일일 장부의 검사 기록에 올바르지 않은 날짜가 있습니다. 날짜를 확인해 주세요.')
    frame['Date'] = dates
    if start:
        frame = frame.loc[frame['Date'] >= pd.to_datetime(start).normalize()]
    if end:
        frame = frame.loc[frame['Date'] < pd.to_datetime(end).normalize() + pd.Timedelta(days=1)]
    if site and site != '전체':
        frame = frame.loc[frame['Site'].fillna('').astype(str).str.strip() == site.strip()]
    return frame.sort_values('Date', kind='stable')
