import os
import re
import pandas as pd

def normalize(text):
    if pd.isna(text): return ""
    t = str(text).lower()
    t = re.sub(r'[^a-z0-9가-힣]', '', t)
    return t.strip()

def get_standard_key(kw, norm_synonyms):
    if kw in norm_synonyms:
        return kw
    for std_key, syns in norm_synonyms.items():
        if kw in syns:
            return std_key
    return None

def run_simulation():
    selected_folder = "Na-aba"
    excel_files = [f for f in os.listdir(selected_folder) 
                   if (f.endswith('.xlsx') or f.endswith('.xlsm')) and not f.startswith('~$') and "Smart_Merged" not in f]
    
    # User's Korean keywords list
    raw_kws = ["순번", "조인트", "도면", "두께", "수량"]
    norm_keywords = [normalize(k) for k in raw_kws]
    
    SYNONYMS = {
        "no": ["no", "no.", "seq", "순번", "연번", "일련번호", "번호", "num", "idx", "호"],
        "dwg": ["dwg", "dwgno", "dwg.no", "도면", "도면번호", "도면명", "drawing", "drawingno", "drawing.no", "iso"],
        "joint": ["joint", "jointno", "joint.no", "조인트", "jnt", "jntno", "point", "포인트"],
        "size": ["size", "규격", "구경", "사이즈", "dia", "nps", "pipe size", "파이프 규격"],
        "thk": ["thk", "thickness", "두께", "t", "thk.", "thick"],
        "result": ["result", "결과", "판정", "결과판정", "decision", "판정결과"],
        "date": ["date", "날짜", "검사일", "검사일자", "일자", "dateofexam", "요청일", "요청일자"],
        "reportno": ["reportno", "report.no", "report_no", "성적서번호", "성적서", "보고서번호", "보고서"],
        "identificationno": ["identificationno", "idno", "id_no", "관리번호", "식별번호", "id"],
        "film": ["film", "filmno", "film.no", "필름", "필름번호", "매수", "수량", "qty", "quantity", "filmqty", "filmquantity", "nooffilm", "numberoffilm", "jointqty", "jointquantity", "joint수량", "조인트수량"]
    }
    
    NORM_SYNONYMS = {k: [normalize(s) for s in v] for k, v in SYNONYMS.items()}
    header_kws = ['no', 'joint', 'dwg', 'size', 'thk', 'result', 'date']
    
    all_data = []
    
    for file in excel_files[:5]:  # Process first 5 files
        print(f"\nAnalyzing file: {file}")
        file_path = os.path.join(selected_folder, file)
        try:
            xls = pd.ExcelFile(file_path)
            for sheet_name in xls.sheet_names:
                raw_df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
                
                # 1. Global metadata (Report No)
                meta_info = {}
                for r_idx, row in raw_df.iterrows():
                    if r_idx > 50: break
                    row_vals = row.values
                    for c_idx, val in enumerate(row_vals):
                        if pd.isna(val): continue
                        s_val = str(val)
                        s_upper = s_val.upper()
                        if ("REPORT" in s_upper and "NO" in s_upper) or ("성적서" in s_val and "번호" in s_val):
                            extracted_val = ""
                            if ":" in s_val:
                                extracted_val = s_val.split(":", 1)[1].strip()
                            else:
                                tmp = re.sub(r'(?i)(report\s*no\.?|성적서\s*번호)', '', s_val).strip()
                                if tmp: extracted_val = tmp
                            
                            if not extracted_val or extracted_val.upper() == "NAN":
                                for offset in range(1, 4):
                                    if c_idx + offset < len(row_vals):
                                        v = str(row_vals[c_idx + offset]).strip()
                                        if v and v.upper() != "NAN" and len(v) < 30:
                                            extracted_val = v
                                            break
                            if extracted_val and extracted_val.upper() != "NAN":
                                clean_val = re.split(r'(?i)\n|/|\s{2,}|date|일자|page|페이지|\(|\[|<', extracted_val)[0].strip()
                                if clean_val:
                                    meta_info["Report No"] = clean_val
                                    break
                    if "Report No" in meta_info: break
                
                if "Report No" not in meta_info:
                    meta_info["Report No"] = os.path.splitext(file)[0]
                
                # 2. Find header row
                best_row = 0
                max_score = 0
                for idx, row in raw_df.iterrows():
                    if idx > 60: break
                    row_content = "".join([str(v) for v in row.values if pd.notna(v)])
                    norm_content = normalize(row_content)
                    
                    score = 0
                    for kw in header_kws:
                        norm_kw = normalize(kw)
                        if norm_kw in norm_content or any(syn in norm_content for syn in NORM_SYNONYMS.get(norm_kw, [])):
                            score += 1
                    if score > max_score:
                        max_score = score
                        best_row = idx
                
                if max_score >= 3:
                    df = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=best_row)
                    df.columns = [str(c).strip() for c in df.columns]
                    df = df.loc[:, ~df.columns.duplicated()]
                    
                    final_cols = []
                    norm_col_map = {col: normalize(col) for col in df.columns}
                    keyword_to_col = {}
                    
                    for raw_kw in raw_kws:
                        kw = normalize(raw_kw)
                        std_key = get_standard_key(kw, NORM_SYNONYMS)
                        
                        # 1. Exact match
                        match = next((orig for orig, norm in norm_col_map.items() if norm == kw and orig not in final_cols), None)
                        if match:
                            final_cols.append(match)
                            keyword_to_col[kw] = match
                            continue
                        
                        # 2. Synonym match
                        if std_key:
                            syns = NORM_SYNONYMS.get(std_key, [])
                            match = next((orig for orig, norm in norm_col_map.items() if norm in syns and orig not in final_cols), None)
                            if match:
                                final_cols.append(match)
                                keyword_to_col[kw] = match
                                continue
                                
                        # 3. Partial match
                        if len(kw) >= 3:
                            syns = NORM_SYNONYMS.get(std_key, []) if std_key else []
                            match = next((orig for orig, norm in norm_col_map.items() if (kw in norm or any(syn in norm for syn in syns if len(syn) >= 3)) and orig not in final_cols), None)
                            if match:
                                final_cols.append(match)
                                keyword_to_col[kw] = match
                                continue
                    
                    # Check if joint column is matched
                    joint_col_name = next((c for c in keyword_to_col.keys() if c == "joint" or any(c == syn for syn in NORM_SYNONYMS.get("joint", []))), None)
                    if not joint_col_name:
                        print(f"  [Warn] Skipped sheet: '{sheet_name}' (No Joint column matched). Detected columns: {list(df.columns)}")
                        continue
                    
                    df = df[final_cols]
                    rename_map = {keyword_to_col[kw]: raw_kw for raw_kw in raw_kws if (kw := normalize(raw_kw)) in keyword_to_col}
                    df.rename(columns=rename_map, inplace=True)
                    
                    # Forward fill using resolved columns
                    dwg_col_real = next((c for c in df.columns if normalize(c) == "dwg" or any(normalize(c) == syn for syn in NORM_SYNONYMS.get("dwg", []))), None)
                    joint_col_real = next((c for c in df.columns if normalize(c) == "joint" or any(normalize(c) == syn for syn in NORM_SYNONYMS.get("joint", []))), None)
                    
                    if dwg_col_real: df[dwg_col_real] = df[dwg_col_real].ffill()
                    if joint_col_real: df[joint_col_real] = df[joint_col_real].ffill()
                    
                    no_col_real = next((c for c in df.columns if normalize(c) == "no" or any(normalize(c) == syn for syn in NORM_SYNONYMS.get("no", []))), None)
                    filter_col = no_col_real if no_col_real else joint_col_real
                    
                    if filter_col:
                        df = df.dropna(subset=[filter_col])
                        df = df[df[filter_col].astype(str).str.strip() != ""]
                        if filter_col == no_col_real:
                            df = df[df[filter_col].astype(str).str.contains(r'\d', regex=True, na=False)]
                            df = df[~df[filter_col].astype(str).str.contains(r'[a-zA-Z가-힣]', regex=True, na=False)]
                        else:
                            exclude_terms = r'(?i)page|페이지|total|합계|sub-total|소계|grand|seoul|inspection|testing|report|project|client|customer|date'
                            df = df[~df[filter_col].astype(str).str.contains(exclude_terms, regex=True, na=False)]
                    
                    for m_key, m_val in meta_info.items():
                        df[m_key] = m_val
                    
                    print(f"  [OK] Extracted from '{sheet_name}': {len(df)} rows. Columns: {list(df.columns)}")
                    all_data.append(df)
        except Exception as e:
            print(f"  [Error] file: {file}, error: {e}")
            
    if all_data:
        combined_df = pd.concat(all_data, ignore_index=True, sort=False)
        combined_df = combined_df.loc[:, ~combined_df.columns.duplicated()]
        combined_df.drop_duplicates(inplace=True)
        
        print(f"\nCombined DataFrame shape: {combined_df.shape}")
        print(f"Columns: {list(combined_df.columns)}")
        
        # Test the improved total logic
        film_col = next((c for c in combined_df.columns if normalize(c) == "film" or any(normalize(c) == syn for syn in NORM_SYNONYMS.get("film", []))), None)
        joint_col = next((c for c in combined_df.columns if normalize(c) == "joint" or any(normalize(c) == syn for syn in NORM_SYNONYMS.get("joint", []))), None)
        
        print(f"Detected film_col: {film_col}")
        print(f"Detected joint_col: {joint_col}")
        
        if len(combined_df) > 0:
            no_col_name = next((c for c in combined_df.columns if normalize(c) == "no" or any(normalize(c) == syn for syn in NORM_SYNONYMS.get("no", []))), None)
            report_col = next((c for c in combined_df.columns if normalize(c) == "reportno" or any(normalize(c) == syn for syn in NORM_SYNONYMS.get("reportno", []))), None)
            label_col = no_col_name if no_col_name else next((c for c in combined_df.columns if c != joint_col and c != film_col and c != report_col), combined_df.columns[0])
            
            print(f"no_col_name: {no_col_name}")
            print(f"report_col: {report_col}")
            print(f"label_col: {label_col}")
            
            if film_col:
                combined_df['_temp_numeric_film'] = pd.to_numeric(combined_df[film_col].astype(str).str.extract(r'(\d+\.?\d*)')[0], errors='coerce')
            
            new_dfs = []
            grand_total = 0
            grand_joint_total = 0
            
            if report_col:
                print("\nGrouping by report_col...")
                for rep_no, group in combined_df.groupby(report_col, sort=False):
                    sub_total = group['_temp_numeric_film'].sum() if film_col else 0
                    grand_total += sub_total
                    
                    sub_joint = 0
                    if joint_col:
                        valid_joints = group[joint_col].dropna().astype(str).str.strip()
                        valid_joints = valid_joints[valid_joints != ""]
                        sub_joint = valid_joints.nunique()
                    grand_joint_total += sub_joint
                    print(f"Report: {rep_no} | Rows: {len(group)} | sub_joint count: {sub_joint} | sub_film: {sub_total}")
            else:
                grand_total = combined_df['_temp_numeric_film'].sum() if film_col else 0
                if joint_col:
                    valid_joints = combined_df[joint_col].dropna().astype(str).str.strip()
                    valid_joints = valid_joints[valid_joints != ""]
                    grand_joint_total = valid_joints.nunique()
                print(f"No report_col. Grand total joint: {grand_joint_total} | Grand total film: {grand_total}")

run_simulation()
