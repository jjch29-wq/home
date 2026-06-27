"""
RT-0001, RT-0022 집중 검증 - 헤더 선택 및 joint/defect 수량 확인
"""
import importlib.util, os, sys, types

# tkinter mock
for mod_name in ['tkinter', 'tkinter.ttk', 'tkinter.filedialog', 'tkinter.messagebox']:
    m = types.ModuleType(mod_name)
    sys.modules[mod_name] = m

import tkinter as tk
tk.Tk = lambda: None
tk.BooleanVar = lambda **kw: None
tk.StringVar = lambda **kw: None

spec = importlib.util.spec_from_file_location('merger', r'c:\Users\-\OneDrive\바탕 화면\home\요청서 합치기.py')
merger_mod = importlib.util.module_from_spec(spec)
try:
    spec.loader.exec_module(merger_mod)
except Exception as e:
    pass

import pandas as pd

FOLDER = r'C:\Users\-\OneDrive\바탕 화면\JHC RT'

# 직접 헤더 탐지 함수만 테스트
import re

def normalize(text):
    if pd.isna(text): return ""
    t = str(text).lower()
    t = re.sub(r'[^a-z0-9가-힣]', '', t)
    return t.strip()

header_kws = ['no', 'joint', 'dwg', 'size', 'thk', 'result', 'date']

NORM_SYNONYMS = {
    "no": [normalize(s) for s in ["no", "no.", "seq", "순번", "연번", "일련번호", "번호", "num", "idx"]],
    "dwg": [normalize(s) for s in ["dwg", "dwgno", "dwg.no", "도면", "도면번호", "도면명", "drawing", "iso"]],
    "joint": [normalize(s) for s in ["joint", "jointno", "joint.no", "조인트", "jnt", "jntno"]],
    "thk": [normalize(s) for s in ["thk", "thickness", "두께", "thick"]],
    "result": [normalize(s) for s in ["result", "결과", "판정"]],
    "date": [normalize(s) for s in ["date", "날짜", "검사일"]],
}

for fname in ['SIT-K1-JHC-PIP-RT-0001.xlsx', 'SIT-K1-JHC-PIP-RT-0022.xlsx']:
    path = os.path.join(FOLDER, fname)
    raw_df = pd.read_excel(path, sheet_name=0, header=None)
    
    best_row = 0
    max_score = 0
    for idx, row in raw_df.iterrows():
        if idx > 60: break
        row_content = "".join([str(v) for v in row.values if pd.notna(v)])
        norm_content = normalize(row_content)
        score = sum(1 for kw in header_kws
                    if normalize(kw) in norm_content or
                    any(syn in norm_content for syn in NORM_SYNONYMS.get(normalize(kw), [])))
        if score > max_score:
            max_score = score
            best_row = idx
    
    # look-ahead
    final_header_row = best_row
    best_row_nonnull = sum(1 for v in raw_df.iloc[best_row] if pd.notna(v) and str(v).strip() not in ['', 'nan'])
    for look_ahead in range(1, 4):
        check_idx = best_row + look_ahead
        if check_idx >= len(raw_df): break
        row_content = "".join([str(v) for v in raw_df.iloc[check_idx].values if pd.notna(v)])
        norm_content = normalize(row_content)
        row_score = sum(1 for kw in header_kws
                        if normalize(kw) in norm_content or
                        any(syn in norm_content for syn in NORM_SYNONYMS.get(normalize(kw), [])))
        row_nonnull = sum(1 for v in raw_df.iloc[check_idx] if pd.notna(v) and str(v).strip() not in ['', 'nan'])
        if row_score >= max_score and row_nonnull > best_row_nonnull:
            final_header_row = check_idx
            best_row_nonnull = row_nonnull
    
    print(f"\n{fname}: best_row={best_row}(score={max_score}), final_header_row={final_header_row}(nonnull={best_row_nonnull})")
    
    df = pd.read_excel(path, sheet_name=0, header=final_header_row)
    df.columns = [str(c).strip() if not str(c).startswith('Unnamed') else f"_col{i}" for i, c in enumerate(df.columns)]
    df = df.loc[:, ~df.columns.duplicated()]
    
    print(f"  Columns: {list(df.columns[:12])}")
    if 'Joint' in df.columns:
        joints = df['Joint'].dropna().astype(str).str.strip()
        joints = joints[~joints.isin(['', 'Joint']) & ~joints.str.contains('total|sub|grand|소계|합계', case=False, na=False)]
        print(f"  Joint count (data rows): {len(joints)}")
        print(f"  Joint values: {joints.tolist()}")
    if 'Defect Rev' in df.columns:
        defect_sum = pd.to_numeric(df['Defect Rev'], errors='coerce').sum()
        print(f"  Defect Rev sum: {defect_sum}")
