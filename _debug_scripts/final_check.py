"""
최종 검증 - Defect Rev 컬럼만 확인
"""
import os, glob, pandas as pd

FOLDER = r'C:\Users\-\OneDrive\바탕 화면\JHC RT'
KEYWORDS = 'No, Joint, Dwg, THK, Result, Date, Report No, Defect Rev'

import re

def normalize(text):
    if pd.isna(text): return ""
    t = str(text).lower()
    t = re.sub(r'[^a-z0-9가-힣]', '', t)
    return t.strip()

SYNONYMS = {
    "no": ["no", "no.", "seq", "순번", "연번", "일련번호", "번호", "num", "idx", "호"],
    "dwg": ["dwg", "dwgno", "dwg.no", "도면", "도면번호", "도면명", "drawing", "drawingno", "drawing.no", "iso"],
    "joint": ["joint", "jointno", "joint.no", "조인트", "jnt", "jntno", "point", "포인트"],
    "size": ["size", "규격", "구경", "사이즈", "dia", "nps"],
    "thk": ["thk", "thickness", "두께", "t", "thk.", "thick"],
    "result": ["result", "결과", "판정", "결과판정", "decision"],
    "date": ["date", "날짜", "검사일", "검사일자", "일자"],
    "reportno": ["reportno", "report.no", "report_no", "성적서번호", "성적서", "보고서번호"],
    "defect": ["defect", "reject", "불합격", "결함", "defectqty", "rejectqty", "defect rev", "defectrev", "defect_rev"],
}
NORM_SYNONYMS = {k: [normalize(s) for s in v] for k, v in SYNONYMS.items()}

# 파일별 Defect Rev 컬럼 인식 테스트
test_files = [
    'SIT-K1-JHC-PIP-RT-0022.xlsx',
    'SIT-K1-JHC-PIP-RT-0022R1.xlsx',
    'SIT-K1-JHC-PIP-RT-0050.xlsx',
    'SIT-K1-JHC-PIP-RT-0001.xlsx',
]

raw_kws = [k.strip() for k in KEYWORDS.split(',')]
norm_keywords = [normalize(k) for k in raw_kws]

for fname in test_files:
    path = os.path.join(FOLDER, fname)
    df_raw = pd.read_excel(path, header=None)

    # 최적 헤더 행 찾기 (row 0)
    for best_row in range(5):
        row_str = ' '.join(str(v) for v in df_raw.iloc[best_row].values if pd.notna(v))
        if 'Joint' in row_str:
            break

    df = pd.read_excel(path, header=best_row)
    df.columns = [str(c).strip() if pd.notna(c) else f"Unnamed_{i}" for i, c in enumerate(df.columns)]

    # 컬럼 매핑
    col_map = {}
    used = set()
    for kw_raw, norm_kw in zip(raw_kws, norm_keywords):
        for col in df.columns:
            norm_col = normalize(col)
            if norm_col in used:
                continue
            std_key = None
            for k, syns in NORM_SYNONYMS.items():
                if norm_kw in syns or norm_kw == k:
                    std_key = k
                    break
            match = (norm_col == norm_kw or
                     (std_key and norm_col in NORM_SYNONYMS.get(std_key, [])) or
                     (len(norm_kw) >= 4 and norm_kw in norm_col) or
                     (len(norm_col) >= 4 and norm_col in norm_kw))
            if match:
                col_map[kw_raw] = col
                used.add(norm_col)
                break

    defect_mapped = col_map.get("Defect Rev", "NOT FOUND")
    defect_sum = 0
    if defect_mapped != "NOT FOUND" and defect_mapped in df.columns:
        defect_sum = pd.to_numeric(df[defect_mapped], errors='coerce').sum()

    print(f"{fname}: 'Defect Rev' -> '{defect_mapped}' (합계={defect_sum})")
