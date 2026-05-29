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

def simulate():
    print("=== Running Robust Merger Simulation ===")
    selected_folder = "Na-aba"
    excel_files = [f for f in os.listdir(selected_folder) 
                   if (f.endswith('.xlsx') or f.endswith('.xlsm')) and not f.startswith('~$') and "Smart_Merged" not in f and "Simulation_Merged" not in f]
    
    # Selected keywords (corresponds to default UI keywords)
    raw_kws = ["No", "Joint", "Dwg", "Size", "THK", "Result", "Date", "Report No", "Identification No", "Defect", "Rev"]
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
        "film": ["film", "filmno", "film.no", "필름", "필름번호", "매수", "수량", "qty", "quantity", "filmqty", "filmquantity", "nooffilm", "numberoffilm", "jointqty", "jointquantity", "joint수량", "조인트수량"],
        "defect": ["defect", "reject", "불합격", "결함", "defectqty", "rejectqty", "불합격수량", "결함수량"],
        "rev": ["rev", "revision", "repair", "classification", "개정", "보수", "수리", "revno", "revisionno", "rev.no", "orca", "o/r/c/a", "구분"]
    }
    
    NORM_SYNONYMS = {k: [normalize(s) for s in v] for k, v in SYNONYMS.items()}
    header_kws = ['no', 'joint', 'dwg', 'size', 'thk', 'result', 'date']
    
    all_data = []
    
    for file in excel_files:
        print(f"Analyzing: {file}")
        file_path = os.path.join(selected_folder, file)
        try:
            xls = pd.ExcelFile(file_path)
            for sheet_name in xls.sheet_names:
                raw_df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
                
                # 1. Report No extraction
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
                                if not re.search(r'\d', s_val):
                                    extracted_val = ""
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
                                    if re.search(r'\d', clean_val) or clean_val.upper() in ['N/A', '-', 'TBD', 'NA']:
                                        meta_info["Report No"] = clean_val
                                        print(f"  -> Extracted Report No: {clean_val}")
                                        break
                                    else:
                                        extracted_val = ""
                    if "Report No" in meta_info: break
                
                if "Report No" not in meta_info:
                    base_name = os.path.splitext(file)[0]
                    meta_info["Report No"] = base_name
                    print(f"  -> Extracted Report No (filename fallback): {base_name}")
                
                # 2. Header identification
                best_row = 0
                max_score = 0
                for idx, row in raw_df.iterrows():
                    if idx > 60: break
                    row_content = "".join([str(v) for v in row.values if pd.notna(v)])
                    norm_content = normalize(row_content)
                    score = sum(1 for kw in header_kws if normalize(kw) in norm_content or any(syn in norm_content for syn in NORM_SYNONYMS.get(normalize(kw), [])))
                    if score > max_score:
                        max_score = score
                        best_row = idx
                
                if max_score >= 3:
                    df = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=best_row)
                    df.columns = [str(c).strip() for c in df.columns]
                    df = df.loc[:, ~df.columns.duplicated()]
                    
                    # Column matching
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
                    
                    joint_matched = any(get_standard_key(k, NORM_SYNONYMS) == "joint" for k in keyword_to_col.keys())
                    if not joint_matched:
                        print(f"  Sheet '{sheet_name}': Skipped (No Joint Column)")
                        continue
                    
                    df = df[final_cols]
                    rename_map = {keyword_to_col[kw]: raw_kw for raw_kw in raw_kws if (kw := normalize(raw_kw)) in keyword_to_col}
                    df.rename(columns=rename_map, inplace=True)
                    
                    # Find Dwg, Joint, No columns dynamically
                    dwg_col_real = next((c for c in df.columns if normalize(c) == "dwg" or any(normalize(c) == syn for syn in NORM_SYNONYMS.get("dwg", []))), None)
                    joint_col_real = next((c for c in df.columns if normalize(c) == "joint" or any(normalize(c) == syn for syn in NORM_SYNONYMS.get("joint", []))), None)
                    no_col_real = next((c for c in df.columns if normalize(c) == "no" or any(normalize(c) == syn for syn in NORM_SYNONYMS.get("no", []))), None)
                    
                    if dwg_col_real:
                        df[dwg_col_real] = df[dwg_col_real].ffill()
                    if joint_col_real:
                        df[joint_col_real] = df[joint_col_real].ffill()
                        
                    # Row filtering
                    filter_col = no_col_real if no_col_real else joint_col_real
                    if filter_col:
                        df = df.dropna(subset=[filter_col])
                        df = df[df[filter_col].astype(str).str.strip() != ""]
                        if filter_col == no_col_real:
                            df = df[df[no_col_real].astype(str).str.contains(r'\d', regex=True, na=False)]
                            exclude_terms = r'(?i)page|페이지|total|합계|sub-total|소계|grand|seoul|inspection|testing|report|project|client|customer|date'
                            df = df[~df[no_col_real].astype(str).str.contains(exclude_terms, regex=True, na=False)]
                        else:
                            exclude_terms = r'(?i)page|페이지|total|합계|sub-total|소계|grand|seoul|inspection|testing|report|project|client|customer|date'
                            df = df[~df[filter_col].astype(str).str.contains(exclude_terms, regex=True, na=False)]
                    
                    for m_key, m_val in meta_info.items():
                        df[m_key] = m_val
                    
                    print(f"  Sheet '{sheet_name}': Success, {len(df)} rows extracted.")
                    all_data.append(df)
        except Exception as e:
            print(f"  Error reading {file}: {e}")
            
    if not all_data:
        print("No data matched.")
        return
        
    combined_df = pd.concat(all_data, ignore_index=True, sort=False)
    combined_df = combined_df.loc[:, ~combined_df.columns.duplicated()]
    combined_df.drop_duplicates(inplace=True)
    
    # Summing
    film_col = next((c for c in combined_df.columns if normalize(c) == "film" or any(normalize(c) == syn for syn in NORM_SYNONYMS.get("film", []))), None)
    joint_col = next((c for c in combined_df.columns if normalize(c) == "joint" or any(normalize(c) == syn for syn in NORM_SYNONYMS.get("joint", []))), None)
    defect_col = next((c for c in combined_df.columns if normalize(c) == "defect" or any(normalize(c) == syn for syn in NORM_SYNONYMS.get("defect", []))), None)
    rev_col = next((c for c in combined_df.columns if normalize(c) == "rev" or any(normalize(c) == syn for syn in NORM_SYNONYMS.get("rev", []))), None)
    
    print(f"\nStandardized Columns detected for totals:")
    print(f"  Joint column: {joint_col}")
    print(f"  Film column: {film_col}")
    print(f"  Defect column: {defect_col}")
    print(f"  Rev column: {rev_col}")
    
    if len(combined_df) > 0:
        no_col_name = next((c for c in combined_df.columns if normalize(c) == "no"), None)
        report_col = next((c for c in combined_df.columns if normalize(c) == "reportno" or any(syn in normalize(c) for syn in NORM_SYNONYMS.get("reportno", []))), None)
        label_col = no_col_name if no_col_name else next((c for c in combined_df.columns if c != joint_col and c != film_col and c != report_col), combined_df.columns[0])
        
        if film_col:
            combined_df['_temp_numeric_film'] = pd.to_numeric(combined_df[film_col].astype(str).str.extract(r'(\d+\.?\d*)')[0], errors='coerce')
        if defect_col:
            combined_df['_temp_numeric_defect'] = pd.to_numeric(combined_df[defect_col].astype(str).str.extract(r'(\d+\.?\d*)')[0], errors='coerce')
        if rev_col:
            combined_df['_temp_numeric_rev'] = pd.to_numeric(combined_df[rev_col].astype(str).str.extract(r'(\d+\.?\d*)')[0], errors='coerce')
            
        new_dfs = []
        grand_total = 0
        grand_joint_total = 0
        grand_defect_total = 0
        grand_rev_total = 0
        
        if report_col:
            for rep_no, group in combined_df.groupby(report_col, sort=False):
                new_dfs.append(group)
                
                sub_total = group['_temp_numeric_film'].sum() if film_col else 0
                grand_total += sub_total
                
                sub_joint = 0
                if joint_col:
                    valid_joints = group[joint_col].dropna().astype(str).str.strip()
                    valid_joints = valid_joints[valid_joints != ""]
                    sub_joint = valid_joints.nunique()
                grand_joint_total += sub_joint
                
                # Defect
                sub_defect = 0
                if defect_col:
                    num_sum = group['_temp_numeric_defect'].sum()
                    if num_sum > 0:
                        sub_defect = num_sum
                    else:
                        valid_vals = group[defect_col].dropna().astype(str).str.strip().str.lower()
                        sub_defect = valid_vals[valid_vals.str.contains('reject|fail|ng|불합격|결함', na=False)].count()
                grand_defect_total += sub_defect
                
                # Rev
                sub_rev = 0
                if rev_col:
                    num_sum = group['_temp_numeric_rev'].sum()
                    if num_sum > 0:
                        sub_rev = num_sum
                    else:
                        valid_vals = group[rev_col].dropna().astype(str).str.strip().str.lower()
                        sub_rev = valid_vals[valid_vals.str.contains(r'^r|repair|보수|수리|개정', na=False)].count()
                grand_rev_total += sub_rev
                
                sub_row = {col: "" for col in combined_df.columns}
                if label_col == report_col:
                    sub_row[report_col] = f"{rep_no} (소계)"
                else:
                    sub_row[label_col] = "Sub-Total"
                    sub_row[report_col] = rep_no
                if film_col and film_col != label_col:
                    sub_row[film_col] = int(sub_total) if sub_total % 1 == 0 else sub_total
                if joint_col and joint_col != label_col:
                    sub_row[joint_col] = int(sub_joint)
                if defect_col and defect_col != label_col:
                    sub_row[defect_col] = int(sub_defect) if sub_defect % 1 == 0 else sub_defect
                if rev_col and rev_col != label_col:
                    sub_row[rev_col] = int(sub_rev) if sub_rev % 1 == 0 else sub_rev
                
                print(f"  Subtotal for Report '{rep_no}': Joint Count = {sub_joint}, Defect Count = {sub_defect}, Rev Count = {sub_rev}")
                new_dfs.append(pd.DataFrame([sub_row]))
                
            combined_df = pd.concat(new_dfs, ignore_index=True)
            
        # Grand Total row
        total_row = {col: "" for col in combined_df.columns}
        if label_col == report_col:
            total_row[report_col] = "총합계 (Grand Total)"
        else:
            total_row[label_col] = "Grand Total"
        if film_col and film_col != label_col:
            total_row[film_col] = int(grand_total) if grand_total % 1 == 0 else grand_total
        if joint_col and joint_col != label_col:
            total_row[joint_col] = int(grand_joint_total)
        if defect_col and defect_col != label_col:
            total_row[defect_col] = int(grand_defect_total) if grand_defect_total % 1 == 0 else grand_defect_total
        if rev_col and rev_col != label_col:
            total_row[rev_col] = int(grand_rev_total) if grand_rev_total % 1 == 0 else grand_rev_total
            
        print(f"  Grand Total: Joint Count = {grand_joint_total}, Defect Count = {grand_defect_total}, Rev Count = {grand_rev_total}")
        
        combined_df = pd.concat([combined_df, pd.DataFrame([total_row])], ignore_index=True)
        
        if film_col and '_temp_numeric_film' in combined_df.columns:
            combined_df.drop(columns=['_temp_numeric_film'], inplace=True)
        if defect_col and '_temp_numeric_defect' in combined_df.columns:
            combined_df.drop(columns=['_temp_numeric_defect'], inplace=True)
        if rev_col and '_temp_numeric_rev' in combined_df.columns:
            combined_df.drop(columns=['_temp_numeric_rev'], inplace=True)
            
    out_path = os.path.join(selected_folder, "Simulation_Merged.xlsx")
    combined_df.to_excel(out_path, index=False)
    print(f"\nSuccessfully wrote merged simulation to {out_path}")

simulate()
