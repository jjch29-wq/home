import os
import re
import pandas as pd
from datetime import datetime

# Mimic the app class
class MockApp:
    def __init__(self):
        self.selected_folder = "Na-aba"
        self.excel_files = [f for f in os.listdir(self.selected_folder) 
                            if (f.endswith('.xlsx') or f.endswith('.xlsm')) and not f.startswith('~$') and "Smart_Merged" not in f]
        self.keyword_var = MockVar("Joint, Dwg, THK")
        self.only_totals_var = MockVar(False)

    def add_log(self, msg):
        print("LOG:", msg)

class MockVar:
    def __init__(self, val):
        self.val = val
    def get(self):
        return self.val

# Instantiate and run the logic
app = MockApp()

def normalize(text):
    if pd.isna(text): return ""
    t = str(text).lower()
    t = re.sub(r'[^a-z0-9가-힣]', '', t)
    return t.strip()

# Paste merge_logic from 요청서 합치기.py
all_data = []
raw_kws = [k.strip() for k in app.keyword_var.get().split(',') if k.strip()]
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
    "identificationno": ["identificationno", "idno", "id_no", "관리번호", "식별번호", "id"]
}
NORM_SYNONYMS = {k: [normalize(s) for s in v] for k, v in SYNONYMS.items()}
header_kws = ['no', 'joint', 'dwg', 'size', 'thk', 'result', 'date']

for file in app.excel_files[:5]:
    file_path = os.path.join(app.selected_folder, file)
    try:
        xls = pd.ExcelFile(file_path)
        for sheet_name in xls.sheet_names:
            raw_df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
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
                base_name = os.path.splitext(file)[0]
                meta_info["Report No"] = base_name

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

                if "joint" not in keyword_to_col:
                    continue
                    
                df = df[final_cols]
                rename_map = {keyword_to_col[kw]: raw_kw for raw_kw in raw_kws if (kw := normalize(raw_kw)) in keyword_to_col}
                df.rename(columns=rename_map, inplace=True)
                
                if "Dwg" in df.columns:
                    df["Dwg"] = df["Dwg"].ffill()
                if "Joint" in df.columns:
                    df["Joint"] = df["Joint"].ffill()
                    
                filter_col = "No" if "No" in df.columns else ("Joint" if "Joint" in df.columns else None)
                if filter_col:
                    df = df.dropna(subset=[filter_col])
                    df = df[df[filter_col].astype(str).str.strip() != ""]
                    if filter_col == "No":
                        df = df[df["No"].astype(str).str.contains(r'\d', regex=True, na=False)]
                        df = df[~df["No"].astype(str).str.contains(r'[a-zA-Z가-힣]', regex=True, na=False)]
                    else:
                        exclude_terms = r'(?i)page|페이지|total|합계|sub-total|소계|grand|seoul|inspection|testing|report|project|client|customer|date'
                        df = df[~df[filter_col].astype(str).str.contains(exclude_terms, regex=True, na=False)]
                
                for m_key, m_val in meta_info.items():
                    df[m_key] = m_val
                    
                all_data.append(df)
    except Exception as e:
        print(f"Error {file}: {e}")

if all_data:
    combined_df = pd.concat(all_data, ignore_index=True, sort=False)
    combined_df = combined_df.loc[:, ~combined_df.columns.duplicated()]
    combined_df.drop_duplicates(inplace=True)
    
    # Let's see what combined_df has here
    print("Initial combined_df columns:", list(combined_df.columns))
    print("Initial combined_df shape:", combined_df.shape)
    
    # We execute total logic
    film_col = next((c for c in combined_df.columns if "film" in normalize(c) and "size" not in normalize(c) and "loc" not in normalize(c)), None)
    joint_col = next((c for c in combined_df.columns if normalize(c) == "joint"), None)
    
    if len(combined_df) > 0:
        no_col_name = next((c for c in combined_df.columns if normalize(c) == "no"), None)
        report_col = next((c for c in combined_df.columns if normalize(c) == "reportno" or any(syn in normalize(c) for syn in NORM_SYNONYMS.get("reportno", []))), None)
        label_col = no_col_name if no_col_name else next((c for c in combined_df.columns if c != joint_col and c != film_col and c != report_col), combined_df.columns[0])
        
        print("Variables in total logic:")
        print("film_col:", film_col)
        print("joint_col:", joint_col)
        print("no_col_name:", no_col_name)
        print("report_col:", report_col)
        print("label_col:", label_col)
        
        if film_col:
            combined_df['_temp_numeric_film'] = pd.to_numeric(combined_df[film_col].astype(str).str.extract(r'(\d+\.?\d*)')[0], errors='coerce')
        
        new_dfs = []
        grand_total = 0
        grand_joint_total = 0
        
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
                new_dfs.append(pd.DataFrame([sub_row]))
            
            combined_df = pd.concat(new_dfs, ignore_index=True)
        else:
            grand_total = combined_df['_temp_numeric_film'].sum() if film_col else 0
            if joint_col:
                valid_joints = combined_df[joint_col].dropna().astype(str).str.strip()
                valid_joints = valid_joints[valid_joints != ""]
                grand_joint_total = valid_joints.nunique()
            else:
                grand_joint_total = 0

        total_row = {col: "" for col in combined_df.columns}
        if label_col == report_col:
            total_row[report_col] = "총합계 (Grand Total)"
        else:
            total_row[label_col] = "Grand Total"
        if film_col and film_col != label_col:
            total_row[film_col] = int(grand_total) if grand_total % 1 == 0 else grand_total
        if joint_col and joint_col != label_col:
            total_row[joint_col] = int(grand_joint_total)
        combined_df = pd.concat([combined_df, pd.DataFrame([total_row])], ignore_index=True)
        
        if film_col and '_temp_numeric_film' in combined_df.columns:
            combined_df.drop(columns=['_temp_numeric_film'], inplace=True)
            
        if app.only_totals_var.get():
            mask = combined_df.astype(str).apply(lambda row: row.str.contains("Sub-Total|Grand Total|소계|총합계", case=False).any(), axis=1)
            combined_df = combined_df[mask]

    print("\nFinal combined_df columns:", list(combined_df.columns))
    print("Final combined_df shape:", combined_df.shape)
    print("Last 10 rows:")
    print(combined_df.tail(10).to_string())
