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

def test_merger_logic():
    print("=== Testing Merger Logic ===")
    
    # 1. Synonyms Dictionary definition
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
    
    # 2. Mock sheet columns and data
    # Scenario: User requests columns: 순번, 조인트, 도면, 두께, 수량, Defect, Rev
    raw_kws = ["순번", "조인트", "도면", "두께", "수량", "Defect", "Rev"]
    
    # Mock Sheet 1 (Standard English/Korean headers)
    # The sheet has 'Drawing No.', 'Joint No.', 'THK', 'No. of Film', 'Test Location', 'Reject', 'Classification'
    df_cols = ['Drawing No.', 'Joint No.', 'THK', 'No. of Film', 'Test Location', 'Reject', 'Classification']
    df_data = [
        ['D-100', '1', '9.5', '3', 'Shop', '0', 'O'],
        ['D-100', '2', '9.5', '3', 'Shop', '1', 'R1'], # repair row
        ['D-100', '3', '9.5', '3', 'Shop', '0', 'O'],
    ]
    df = pd.DataFrame(df_data, columns=df_cols)
    
    # Column matching
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
                
        # 3. Partial match (with len >= 3 safeguard for synonyms)
        if len(kw) >= 3:
            syns = NORM_SYNONYMS.get(std_key, []) if std_key else []
            match = next((orig for orig, norm in norm_col_map.items() if (kw in norm or any(syn in norm for syn in syns if len(syn) >= 3)) and orig not in final_cols), None)
            if match:
                final_cols.append(match)
                keyword_to_col[kw] = match
                continue

    print("Keyword to column mapping:")
    for k, v in keyword_to_col.items():
        print(f"  {k} -> {v}")
        
    # Check that THK mapped to THK and NOT to Test Location
    assert keyword_to_col.get(normalize("두께")) == "THK", f"Failed: '두께' mapped to {keyword_to_col.get(normalize('두께'))}"
    assert "Test Location" not in final_cols, "Failed: 'Test Location' was incorrectly matched."
    print("Match Safeguard Check Passed: '두께' mapped to 'THK' and not 'Test Location'.")
    
    # Rename columns to standardized names
    df = df[final_cols]
    rename_map = {keyword_to_col[kw]: raw_kw for raw_kw in raw_kws if (kw := normalize(raw_kw)) in keyword_to_col}
    df.rename(columns=rename_map, inplace=True)
    
    print("Standardized DataFrame columns:", list(df.columns))
    
    # Forward fill
    joint_col_real = next((c for c in df.columns if normalize(c) == "joint" or any(normalize(c) == syn for syn in NORM_SYNONYMS.get("joint", []))), None)
    if joint_col_real:
        df[joint_col_real] = df[joint_col_real].ffill()
        
    print("DataFrame content after mapping and ffill:")
    print(df.to_string())
    
    # Test Totals logic
    combined_df = df
    film_col = next((c for c in combined_df.columns if normalize(c) == "film" or any(normalize(c) == syn for syn in NORM_SYNONYMS.get("film", []))), None)
    joint_col = next((c for c in combined_df.columns if normalize(c) == "joint" or any(normalize(c) == syn for syn in NORM_SYNONYMS.get("joint", []))), None)
    defect_col = next((c for c in combined_df.columns if normalize(c) == "defect" or any(normalize(c) == syn for syn in NORM_SYNONYMS.get("defect", []))), None)
    rev_col = next((c for c in combined_df.columns if normalize(c) == "rev" or any(normalize(c) == syn for syn in NORM_SYNONYMS.get("rev", []))), None)
    
    print(f"Detected film_col: {film_col}")
    print(f"Detected joint_col: {joint_col}")
    print(f"Detected defect_col: {defect_col}")
    print(f"Detected rev_col: {rev_col}")
    
    # Calculate Sub-Totals/Grand Totals
    # Unique joints
    grand_joint_total = 0
    if joint_col:
        valid_joints = combined_df[joint_col].dropna().astype(str).str.strip()
        valid_joints = valid_joints[valid_joints != ""]
        grand_joint_total = valid_joints.nunique()
        
    # Film sums
    grand_film_total = 0
    if film_col:
        temp_numeric = pd.to_numeric(combined_df[film_col].astype(str).str.extract(r'(\d+\.?\d*)')[0], errors='coerce')
        grand_film_total = temp_numeric.sum()
        
    # Defect sums
    grand_defect_total = 0
    if defect_col:
        temp_numeric = pd.to_numeric(combined_df[defect_col].astype(str).str.extract(r'(\d+\.?\d*)')[0], errors='coerce')
        num_sum = temp_numeric.sum()
        if num_sum > 0:
            grand_defect_total = num_sum
        else:
            valid_vals = combined_df[defect_col].dropna().astype(str).str.strip().str.lower()
            grand_defect_total = valid_vals[valid_vals.str.contains('reject|fail|ng|불합격|결함', na=False)].count()
            
    # Rev sums
    grand_rev_total = 0
    if rev_col:
        temp_numeric = pd.to_numeric(combined_df[rev_col].astype(str).str.extract(r'(\d+\.?\d*)')[0], errors='coerce')
        num_sum = temp_numeric.sum()
        if num_sum > 0:
            grand_rev_total = num_sum
        else:
            valid_vals = combined_df[rev_col].dropna().astype(str).str.strip().str.lower()
            grand_rev_total = valid_vals[valid_vals.str.contains(r'^r|repair|보수|수리|개정', na=False)].count()
            
    print(f"Grand Joint Total (expected 3): {grand_joint_total}")
    print(f"Grand Film Total (expected 9): {grand_film_total}")
    print(f"Grand Defect Total (expected 1): {grand_defect_total}")
    print(f"Grand Rev Total (expected 1): {grand_rev_total}")
    
    assert grand_joint_total == 3
    assert grand_film_total == 9
    assert grand_defect_total == 1
    assert grand_rev_total == 1
    print("Totals calculation checks passed!")

if __name__ == "__main__":
    test_merger_logic()
