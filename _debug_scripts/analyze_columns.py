"""
원본 파일의 실제 pandas 헤더 인식 상세 분석
"""
import pandas as pd, os

FOLDER = r'C:\Users\-\OneDrive\바탕 화면\JHC RT'

for fname in ['SIT-K1-JHC-PIP-RT-0001.xlsx', 'SIT-K1-JHC-PIP-RT-0022.xlsx']:
    path = os.path.join(FOLDER, fname)
    print(f'\n=== {fname} ===')
    
    # 모든 시트 확인
    xls = pd.ExcelFile(path)
    print(f'시트 목록: {xls.sheet_names}')
    
    for sheet in xls.sheet_names:
        raw = pd.read_excel(path, sheet_name=sheet, header=None)
        print(f'\n시트: {sheet} | Shape: {raw.shape}')
        
        # 처음 5행 전체 출력
        for r in range(min(5, len(raw))):
            row_data = [(i, str(v)[:30]) for i, v in enumerate(raw.iloc[r]) if pd.notna(v)]
            print(f'  Row {r}: {row_data}')
        
        # 데이터 행 샘플 (10번째 행)
        if len(raw) > 10:
            row_data = [(i, str(v)[:20]) for i, v in enumerate(raw.iloc[10]) if pd.notna(v)]
            print(f'  Row 10 (data): {row_data}')
        
        # header=2로 읽을 때 컬럼 확인
        df2 = pd.read_excel(path, sheet_name=sheet, header=2)
        print(f'\n  header=2 컬럼: {list(df2.columns[:15])}')
        
        # Joint 컬럼이 있으면 첫 5개 값 출력
        if 'Joint' in df2.columns:
            print(f'  Joint 값 (first 5): {df2["Joint"].head(5).tolist()}')
        if 'Defect Rev' in df2.columns:
            print(f'  Defect Rev 값: {df2["Defect Rev"].dropna().tolist()}')
