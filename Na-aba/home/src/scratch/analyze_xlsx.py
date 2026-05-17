import pandas as pd
import openpyxl
import os
import re
import glob

folder = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba"
files = glob.glob(os.path.join(folder, "*.xlsm"))
print(f"Found {len(files)} files in folder.")

SYNONYMS = {
    "no": ["no", "no.", "seq", "순번", "연번", "일련번호", "번호", "num", "idx", "호"],
    "dwg": ["dwg", "dwgno", "dwg.no", "도면", "도면번호", "도면명", "drawing", "drawingno", "drawing.no", "iso"],
    "joint": ["joint", "jointno", "joint.no", "조인트", "jnt", "jntno", "point", "포인트"],
    "size": ["size", "규격", "구경", "사이즈", "dia", "nps", "pipe size", "파이프 규격"],
    "thk": ["thk", "thickness", "두께", "t", "thk.", "thick"],
    "result": ["result", "결과", "판정", "결과판정", "decision", "판정결과"],
    "date": ["date", "날짜", "검사일", "검사일자", "일자", "dateofexam", "요청일", "요청일자"],
    "reportno": ["reportno", "report.no", "report_no", "성적서번호", "성적서", "보고서번호", "보고서"],
    "identificationno": ["identificationno", "idno", "id_no", "관리번호", "식별번호", "id"]
}

def normalize(text):
    if pd.isna(text): return ""
    t = str(text).lower()
    t = re.sub(r'[^a-z0-9가-힣]', '', t)
    return t.strip()

NORM_SYNONYMS = {k: [normalize(s) for s in v] for k, v in SYNONYMS.items()}
header_kws = ['no', 'joint', 'dwg', 'size', 'thk', 'result', 'date']

for f in files[:4]:
    print(f"\n======================================")
    print(f"File: {os.path.basename(f)}")
    xls = pd.ExcelFile(f)
    print(f"Sheets: {xls.sheet_names}")
    
    for sheet_name in xls.sheet_names:
        raw_df = pd.read_excel(f, sheet_name=sheet_name, header=None)
        
        # Header detection
        best_row = 0
        max_score = 0
        for idx, row in raw_df.iterrows():
            if idx > 60: break
            row_content = "".join([str(v) for v in row.values if pd.notna(v)])
            norm_content = normalize(row_content)
            
            score = 0
            for kw in header_kws:
                if normalize(kw) in norm_content or any(syn in norm_content for syn in NORM_SYNONYMS.get(normalize(kw), [])):
                    score += 1
                    
            if score > max_score:
                max_score = score
                best_row = idx
                
        if max_score >= 3:
            df = pd.read_excel(f, sheet_name=sheet_name, skiprows=best_row)
            df.columns = [str(c).strip() for c in df.columns]
            df = df.loc[:, ~df.columns.duplicated()]
            
            # Match columns
            final_cols = []
            norm_col_map = {col: normalize(col) for col in df.columns}
            keyword_to_col = {}
            
            raw_kws = ["No", "Joint", "Dwg", "Size", "THK", "Result", "Date"]
            for raw_kw in raw_kws:
                kw = normalize(raw_kw)
                match = next((orig for orig, norm in norm_col_map.items() if norm == kw and orig not in final_cols), None)
                if match:
                    final_cols.append(match)
                    keyword_to_col[kw] = match
                    continue
                
                syns = NORM_SYNONYMS.get(kw, [])
                match = next((orig for orig, norm in norm_col_map.items() if norm in syns and orig not in final_cols), None)
                if match:
                    final_cols.append(match)
                    keyword_to_col[kw] = match
                    continue
                    
                if len(kw) >= 3:
                    match = next((orig for orig, norm in norm_col_map.items() if (kw in norm or any(syn in norm for syn in syns)) and orig not in final_cols), None)
                    if match:
                        final_cols.append(match)
                        keyword_to_col[kw] = match
                        continue
            
            # Check mandatory joint column
            if "joint" not in keyword_to_col:
                print(f"  [SKIP] Sheet '{sheet_name}' (Score: {max_score}, missing Joint column)")
                continue
                
            df = df[final_cols]
            
            # ffill
            dwg_col = keyword_to_col.get("dwg")
            joint_col = keyword_to_col.get("joint")
            if dwg_col: df[dwg_col] = df[dwg_col].ffill()
            if joint_col: df[joint_col] = df[joint_col].ffill()
            
            # Filter rows
            no_col = keyword_to_col.get("no")
            filter_col = no_col if no_col else joint_col
            
            if filter_col:
                df = df.dropna(subset=[filter_col])
                df = df[df[filter_col].astype(str).str.strip() != ""]
                
                if filter_col == no_col:
                    df = df[df[no_col].astype(str).str.contains(r'\d', regex=True, na=False)]
                    df = df[~df[no_col].astype(str).str.contains(r'[a-zA-Z가-힣]', regex=True, na=False)]
                else:
                    exclude_terms = r'(?i)page|페이지|total|합계|sub-total|소계|grand|seoul|inspection|testing|report|project|client|customer|date'
                    df = df[~df[filter_col].astype(str).str.contains(exclude_terms, regex=True, na=False)]
                    
            print(f"  [PROCESS] Sheet '{sheet_name}' -> Extracted {len(df)} rows. Mapped columns: {list(keyword_to_col.keys())}")
        else:
            print(f"  [SKIP] Sheet '{sheet_name}' (Score: {max_score})")
