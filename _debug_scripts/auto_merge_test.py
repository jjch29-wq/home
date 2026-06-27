"""
간단 자동 병합 테스트 (merge_logic 직접 복사 방식)
"""
import os, sys, glob, re, threading
import pandas as pd
from datetime import datetime

FOLDER = r'C:\Users\-\OneDrive\바탕 화면\JHC RT'
KEYWORDS = 'No, Joint, Dwg, Size, THK, Result, Date, Report No, Identification No, Defect Rev, Rev'
ONLY_TOTALS = True

# merge_logic을 실행하기 위해 ExcelMergerApp에서 필요한 메서드만 추출
# --- normalize / get_standard_key 복사 ---

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

# ---------- 메인 병합 로직 시작 ----------
excel_files = [
    f for f in os.listdir(FOLDER)
    if f.endswith('.xlsx') and not f.startswith('~$') and 'Smart_Merged' not in f
]
print(f"총 파일 수: {len(excel_files)}")

raw_kws = [k.strip() for k in KEYWORDS.split(',') if k.strip()]
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
    "defect": ["defect", "reject", "불합격", "결함", "defectqty", "rejectqty", "불합격수량", "결함수량", "defectrev", "defect rev", "defect_rev"],
    "rev": ["rev", "revision", "repair", "classification", "개정", "보수", "수리", "revno", "revisionno", "rev.no", "orca", "o/r/c/a", "구분"]
}
NORM_SYNONYMS = {k: [normalize(s) for s in v] for k, v in SYNONYMS.items()}
header_kws = ['no', 'joint', 'dwg', 'size', 'thk', 'result', 'date']

all_data = []
log_msgs = []

def add_log(msg):
    log_msgs.append(msg)

# 파일별 추출
for file in sorted(excel_files):
    file_path = os.path.join(FOLDER, file)
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
                                    break
                if "Report No" in meta_info: break
            if "Report No" not in meta_info:
                meta_info["Report No"] = os.path.splitext(file)[0]

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

            if max_score < 2:
                continue

            df = pd.read_excel(file_path, sheet_name=sheet_name, header=best_row)
            df.columns = [str(c).strip() if pd.notna(c) else f"Unnamed_{i}" for i, c in enumerate(df.columns)]

            # 컬럼 매핑
            col_map = {}
            used_raw_cols = set()
            for kw_raw, norm_kw in zip(raw_kws, norm_keywords):
                std_key = get_standard_key(norm_kw, NORM_SYNONYMS)
                for col in df.columns:
                    norm_col = normalize(col)
                    if norm_col in used_raw_cols:
                        continue
                    min_len = 3
                    if (norm_col == norm_kw or
                        (std_key and norm_col in NORM_SYNONYMS.get(std_key, [])) or
                        (len(norm_kw) >= min_len and norm_kw in norm_col) or
                        (len(norm_col) >= min_len and norm_col in norm_kw)):
                        col_map[kw_raw] = col
                        used_raw_cols.add(norm_col)
                        break

            if len(col_map) < 2:
                continue

            selected_cols = list(dict.fromkeys(col_map.values()))
            df = df[selected_cols].copy()
            df = df.rename(columns={v: k for k, v in col_map.items() if v in df.columns})

            joint_col = col_map.get("Joint") if "Joint" in col_map else None
            report_col = col_map.get("Report No") if "Report No" in col_map else None
            film_col = col_map.get("Film No") if "Film No" in col_map else None
            defect_col = col_map.get("Defect Rev") if "Defect Rev" in col_map else None
            rev_col = col_map.get("Rev") if "Rev" in col_map else None

            if joint_col and joint_col in df.columns:
                joint_col = "Joint"
            if defect_col and defect_col in df.columns:
                defect_col = "Defect Rev"
            if rev_col and rev_col in df.columns:
                rev_col = "Rev"

            filter_col = "Report No" if "Report No" in df.columns else ("Joint" if "Joint" in df.columns else None)
            if filter_col:
                df = df.dropna(subset=[filter_col])
                df = df[df[filter_col].astype(str).str.strip() != ""]
                if filter_col == "Joint":
                    exclude_terms = r'(?i)page|페이지|total|합계|sub-total|소계|grand|seoul|inspection|testing|report|project|client|customer|date|joint|조인트'
                    df = df[~df[filter_col].astype(str).str.contains(exclude_terms, regex=True, na=False)]
                else:
                    exclude_terms = r'(?i)page|페이지|total|합계|sub-total|소계|grand|seoul|inspection|testing|report|project|client|customer|date'
                    df = df[~df[filter_col].astype(str).str.contains(exclude_terms, regex=True, na=False)]

            for m_key, m_val in meta_info.items():
                df[m_key] = m_val

            all_data.append(df)
    except Exception as e:
        print(f"오류 [{file}]: {e}")

if not all_data:
    print("추출된 데이터가 없습니다!")
    sys.exit(1)

print(f"\n총 {len(all_data)}개 시트에서 데이터 추출 완료")

# --- 병합 및 소계 계산 ---
combined_df = pd.concat(all_data, ignore_index=True)
print(f"병합 행수: {len(combined_df)}")

# 컬럼 확인
joint_col = "Joint" if "Joint" in combined_df.columns else None
defect_col = "Defect Rev" if "Defect Rev" in combined_df.columns else None
rev_col = "Rev" if "Rev" in combined_df.columns else None
report_col = "Report No" if "Report No" in combined_df.columns else None
film_col = None  # film 제외

print(f"joint_col={joint_col}, defect_col={defect_col}, rev_col={rev_col}")

# temp 숫자 컬럼
if defect_col:
    combined_df['_temp_numeric_defect'] = pd.to_numeric(combined_df[defect_col], errors='coerce').fillna(0)
if rev_col:
    combined_df['_temp_numeric_rev'] = pd.to_numeric(combined_df[rev_col], errors='coerce').fillna(0)

# report_col 기준 label_col 결정
if report_col:
    label_col = joint_col if joint_col else report_col
else:
    label_col = joint_col

new_dfs = []
grand_total = 0
grand_joint_total = 0
grand_defect_total = 0
grand_rev_total = 0

if report_col:
    for rep_no, group in combined_df.groupby(report_col, sort=False):
        new_dfs.append(group)
        sub_joint = 0
        if joint_col:
            valid_joints = group[joint_col].dropna().astype(str).str.strip()
            valid_joints = valid_joints[valid_joints != ""]
            sub_joint = valid_joints.count()
        grand_joint_total += sub_joint

        sub_defect = 0
        if defect_col:
            num_sum = group['_temp_numeric_defect'].sum()
            if num_sum > 0:
                sub_defect = num_sum
        grand_defect_total += sub_defect

        sub_rev = 0
        if rev_col:
            num_sum = group['_temp_numeric_rev'].sum()
            if num_sum > 0:
                sub_rev = num_sum
        grand_rev_total += sub_rev

        sub_row = {col: "" for col in combined_df.columns}
        sub_row[label_col] = "Sub-Total"
        sub_row[report_col] = rep_no
        if joint_col and joint_col != label_col:
            sub_row[joint_col] = int(sub_joint)
        if defect_col:
            sub_row[defect_col] = int(sub_defect) if sub_defect >= 1 else ""
        if rev_col:
            sub_row[rev_col] = int(sub_rev) if sub_rev >= 1 else ""
        new_dfs.append(pd.DataFrame([sub_row]))

    combined_df = pd.concat(new_dfs, ignore_index=True)

# Grand Total
total_row = {col: "" for col in combined_df.columns}
total_row[label_col] = "Grand Total"
if joint_col and joint_col != label_col:
    total_row[joint_col] = int(grand_joint_total)
if defect_col:
    total_row[defect_col] = int(grand_defect_total) if grand_defect_total >= 1 else ""
if rev_col:
    total_row[rev_col] = int(grand_rev_total) if grand_rev_total >= 1 else ""
combined_df = pd.concat([combined_df, pd.DataFrame([total_row])], ignore_index=True)

# 임시 컬럼 제거
for tmp_col in ['_temp_numeric_defect', '_temp_numeric_rev']:
    if tmp_col in combined_df.columns:
        combined_df.drop(columns=[tmp_col], inplace=True)

# Defect/Rev 0 값 빈칸 처리
for col in [defect_col, rev_col]:
    if col and col in combined_df.columns:
        def clean_zero(val):
            if pd.isna(val) or str(val).strip() == "":
                return ""
            try:
                num = pd.to_numeric(val)
                return int(num) if num >= 1 else ""
            except:
                return val
        combined_df[col] = combined_df[col].apply(clean_zero)

# 합계만 추출
if ONLY_TOTALS:
    mask = combined_df.astype(str).apply(
        lambda row: row.str.contains("Sub-Total|Grand Total|소계|총합계", case=False).any(), axis=1
    )
    combined_df = combined_df[mask]

# 저장
out_name = f"Final_Smart_Merged_v2.8_{datetime.now().strftime('%H%M%S')}.xlsx"
out_path = os.path.join(FOLDER, out_name)
combined_df.to_excel(out_path, index=False)
print(f"\n✅ 저장 완료: {out_name}")
print(f"컬럼: {list(combined_df.columns)}")
print(f"총 행수: {len(combined_df)}")

# 검증
r0001 = combined_df[combined_df['Report No'].astype(str).str.contains('SIT-K1-JHC-PIP-RT-0001', na=False)]
print(f"\n[RT-0001 소계]: {r0001[['Joint','Report No']].to_string()}")

if defect_col in combined_df.columns:
    defect_non_blank = combined_df[(combined_df[defect_col] != '') & (combined_df[defect_col].notna())]
    print(f"\n[Defect Rev 값 있는 행]: {len(defect_non_blank)}개")
    print(defect_non_blank[['Joint','Report No', defect_col]].head(10).to_string())

if rev_col in combined_df.columns:
    rev_non_blank = combined_df[(combined_df[rev_col] != '') & (combined_df[rev_col].notna())]
    print(f"\n[Rev 값 있는 행]: {len(rev_non_blank)}개")
    print(rev_non_blank[['Joint','Report No', rev_col]].head(10).to_string())
