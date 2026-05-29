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

def check_all_files():
    selected_folder = "Na-aba"
    excel_files = [f for f in os.listdir(selected_folder) 
                   if (f.endswith('.xlsx') or f.endswith('.xlsm')) and not f.startswith('~$') and "Smart_Merged" not in f]
    
    raw_kws = ["순번", "조인트", "도면", "두께", "수량"]
    
    SYNONYMS = {
        "no": ["no", "no.", "seq", "순번", "연번", "일련번호", "번호", "num", "idx", "호"],
        "dwg": ["dwg", "dwgno", "dwg.no", "도면", "도면번호", "도면명", "drawing", "drawingno", "drawing.no", "iso"],
        "joint": ["joint", "jointno", "joint.no", "조인트", "jnt", "jntno", "point", "포인트"],
        "size": ["size", "규격", "구경", "사이즈", "dia", "nps", "pipe size", "파이프 규격"],
        "thk": ["thk", "thickness", "두께", "t", "thk.", "thick"],
        "film": ["film", "filmno", "film.no", "필름", "필름번호", "매수", "수량", "qty", "quantity", "filmqty", "filmquantity", "nooffilm", "numberoffilm", "jointqty", "jointquantity", "joint수량", "조인트수량"]
    }
    
    NORM_SYNONYMS = {k: [normalize(s) for s in v] for k, v in SYNONYMS.items()}
    header_kws = ['no', 'joint', 'dwg', 'size', 'thk']
    
    for file in excel_files:
        file_path = os.path.join(selected_folder, file)
        try:
            xls = pd.ExcelFile(file_path)
            for sheet_name in xls.sheet_names:
                raw_df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
                
                # Report No extraction
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
                                clean_val = re.split(r'(?i)\s+[가-힣]{2,}|\s+(?:rev|sheet|insp|note|remark)', clean_val)[0].strip()
                                if ":" in clean_val:
                                    tmp = clean_val.split(":")[0].strip()
                                    clean_val = re.sub(r'\s+[A-Za-z가-힣]+$', '', tmp).strip()
                                clean_val = re.split(r'(?i)\s+(?:IP|LF|CR|PO|UC|BT|NF|INCOMPLETE|LACK|DEFECT|결함)\b', clean_val)[0].strip()
                                clean_val = re.sub(r'^[:\-]+|[:\-]+$', '', clean_val).strip()
                                if clean_val:
                                    meta_info["Report No"] = clean_val
                                    break
                    if "Report No" in meta_info: break
                
                if "Report No" not in meta_info:
                    meta_info["Report No"] = os.path.splitext(file)[0]
                
                # Find header row
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
                        match = next((orig for orig, norm in norm_col_map.items() if norm == kw and orig not in final_cols), None)
                        if match:
                            final_cols.append(match)
                            keyword_to_col[kw] = match
                            continue
                        if std_key:
                            syns = NORM_SYNONYMS.get(std_key, [])
                            match = next((orig for orig, norm in norm_col_map.items() if norm in syns and orig not in final_cols), None)
                            if match:
                                final_cols.append(match)
                                keyword_to_col[kw] = match
                                continue
                        if len(kw) >= 3:
                            syns = NORM_SYNONYMS.get(std_key, []) if std_key else []
                            match = next((orig for orig, norm in norm_col_map.items() if (kw in norm or any(syn in norm for syn in syns if len(syn) >= 3)) and orig not in final_cols), None)
                            if match:
                                final_cols.append(match)
                                keyword_to_col[kw] = match
                                continue
                    
                    # Check joint
                    joint_col_name = next((c for c in keyword_to_col.keys() if c == "joint" or any(c == syn for syn in NORM_SYNONYMS.get("joint", []))), None)
                    if not joint_col_name:
                        continue
                    
                    df = df[final_cols]
                    rename_map = {keyword_to_col[kw]: raw_kw for raw_kw in raw_kws if (kw := normalize(raw_kw)) in keyword_to_col}
                    df.rename(columns=rename_map, inplace=True)
                    
                    # Forward fill
                    joint_col_real = next((c for c in df.columns if normalize(c) == "joint" or any(normalize(c) == syn for syn in NORM_SYNONYMS.get("joint", []))), None)
                    if joint_col_real:
                        df[joint_col_real] = df[joint_col_real].ffill()
                        
                    # Filter empty/signature rows
                    if joint_col_real:
                        df_filtered = df.dropna(subset=[joint_col_real])
                        df_filtered = df_filtered[df_filtered[joint_col_real].astype(str).str.strip() != ""]
                        exclude_terms = r'(?i)page|페이지|total|합계|sub-total|소계|grand|seoul|inspection|testing|report|project|client|customer|date'
                        df_filtered = df_filtered[~df_filtered[joint_col_real].astype(str).str.contains(exclude_terms, regex=True, na=False)]
                        
                        unique_joints = df_filtered[joint_col_real].dropna().astype(str).str.strip().unique()
                        print(f"File: {file} | Sheet: {sheet_name}")
                        print(f"  Extracted Report No: {meta_info['Report No']}")
                        print(f"  Total Rows: {len(df_filtered)}")
                        print(f"  Unique Joint Count: {len(unique_joints)}")
                        print(f"  Joints List: {list(unique_joints)}")
        except Exception as e:
            print(f"Error file {file}: {e}")

check_all_files()
