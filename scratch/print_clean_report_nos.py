import os
import re
import pandas as pd

def normalize(text):
    if pd.isna(text): return ""
    t = str(text).lower()
    t = re.sub(r'[^a-z0-9가-힣]', '', t)
    return t.strip()

def check_all_files():
    folders = [".", "Na-aba", "Na-aba/보고서 자료"]
    raw_kws = ["순번", "조인트", "도면", "두께", "수량"]
    
    SYNONYMS = {
        "no": ["no", "no.", "seq", "순번", "연번", "일련번호", "번호", "num", "idx", "호"],
        "dwg": ["dwg", "dwgno", "dwg.no", "도면", "도면번호", "도면명", "drawing", "drawingno", "drawing.no", "iso"],
        "joint": ["joint", "jointno", "joint.no", "조인트", "jnt", "jntno", "point", "포인트"],
        "size": ["size", "규격", "구경", "사이즈", "dia", "nps", "pipe size", "파이프 규격"],
        "thk": ["thk", "thickness", "두께", "t", "thk.", "thick"]
    }
    NORM_SYNONYMS = {k: [normalize(s) for s in v] for k, v in SYNONYMS.items()}
    
    for folder in folders:
        if not os.path.exists(folder): continue
        for file in os.listdir(folder):
            if not (file.endswith('.xlsx') or file.endswith('.xlsm') or file.endswith('.xls')): continue
            if file.startswith('~$') or "Smart_Merged" in file: continue
            
            file_path = os.path.join(folder, file)
            try:
                xls = pd.ExcelFile(file_path)
                for sheet_name in xls.sheet_names:
                    raw_df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
                    
                    # Report No extraction
                    meta_info = {}
                    for r_idx, row in raw_df.iterrows():
                        if r_idx > 40: break
                        for c_idx, val in enumerate(row.values):
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
                                        if c_idx + offset < len(row.values):
                                            v = str(row.values[c_idx + offset]).strip()
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
                    
                    # If this sheet contains joints
                    best_row = 0
                    max_score = 0
                    for idx, row in raw_df.iterrows():
                        if idx > 40: break
                        row_content = "".join([str(v) for v in row.values if pd.notna(v)])
                        norm_content = normalize(row_content)
                        score = sum(1 for kw in ['no', 'joint', 'dwg', 'size', 'thk'] if normalize(kw) in norm_content or any(syn in norm_content for syn in NORM_SYNONYMS.get(normalize(kw), [])))
                        if score > max_score:
                            max_score = score
                            best_row = idx
                    
                    if max_score >= 3:
                        df = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=best_row)
                        df.columns = [str(c).strip() for c in df.columns]
                        df = df.loc[:, ~df.columns.duplicated()]
                        
                        joint_col = next((c for c in df.columns if normalize(c) == "joint" or any(normalize(c) == syn for syn in NORM_SYNONYMS.get("joint", []))), None)
                        if joint_col:
                            df[joint_col] = df[joint_col].ffill()
                            df_filtered = df.dropna(subset=[joint_col])
                            df_filtered = df_filtered[df_filtered[joint_col].astype(str).str.strip() != ""]
                            exclude_terms = r'(?i)page|페이지|total|합계|sub-total|소계|grand|seoul|inspection|testing|report|project|client|customer|date'
                            df_filtered = df_filtered[~df_filtered[joint_col].astype(str).str.contains(exclude_terms, regex=True, na=False)]
                            unique_joints = df_filtered[joint_col].dropna().astype(str).str.strip().unique()
                            
                            # Print in safe format
                            print(f"Folder: {folder} | File: {file} | Sheet: {sheet_name} | Extracted Report No: {repr(meta_info['Report No'])} | Unique Count: {len(unique_joints)}")
                            if "0001" in meta_info['Report No'] or "SIT" in meta_info['Report No']:
                                print(f"  --> MATCH! Joints: {list(unique_joints)}")
            except Exception as e:
                pass

check_all_files()
