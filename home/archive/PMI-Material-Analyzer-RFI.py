import pandas as pd
import os
import glob
import sys
import warnings
import tkinter as tk
from tkinter import filedialog

# 경고 메시지 무시
warnings.simplefilter("ignore")

# ========================================================
# [설정] 작업 폴더 자동 인식
# ========================================================
if getattr(sys, 'frozen', False):
    folder_path = os.path.dirname(sys.executable)
elif '__file__' in globals():
    folder_path = os.path.dirname(os.path.abspath(__file__))
else:
    folder_path = os.getcwd()

print(f"📂 현재 작업 폴더: {folder_path}")

# ========================================================
# [추가 기능] 원하는 순번 입력 받기
# ========================================================
print("\n" + "="*50)
print("Option: 특정 순번(NO)만 골라서 뽑고 싶으면 입력하세요.")
print("예시: 1 3 5  또는  1, 2, 10 (구분자는 공백이나 콤마)")
print("👉 전체를 다 하려면 그냥 [Enter] 키를 누르세요.")
target_input = input("입력 > ").strip()

target_no_list = []
if target_input:
    # 콤마를 공백으로 바꾸고 잘라서 리스트로 만듦
    target_no_list = [x.strip() for x in target_input.replace(',', ' ').split() if x.strip()]
    print(f"✅ 선택된 순번: {target_no_list}")
else:
    print("✅ 전체 데이터를 분석합니다.")
print("="*50 + "\n")

# ========================================================
# [함수] 데이터 처리 및 판정 로직
# ========================================================
def to_float(val):
    """퍼센트, ND, 공란 등을 숫자로 변환"""
    if pd.isna(val): return 0.0
    s = str(val).upper().replace("%", "").strip()
    if "<" in s or "ND" in s or s == "": return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0

def find_col(df, keywords, exclude_keywords=None):
    if exclude_keywords is None: exclude_keywords = []
    for col in df.columns:
        c_str = str(col).upper().strip()
        if any(ex in c_str for ex in exclude_keywords): continue
        if any(k in c_str for k in keywords): return col
    return None

def check_material_grade_with_reason(row):
    """
    재질 판정 및 [탈락 사유] 반환
    """
    cr = row['val_Cr']
    ni = row['val_Ni']
    mo = row['val_Mo']
    mn = row.get('val_Mn', 0.0) 
    
    margin = 0.1 # 10% 여유
    
    # ----------------------------------------------------
    # [1] SUS 316 판정 (Cr:16~18 / Ni:10~14 / Mo:2~3)
    # ----------------------------------------------------
    limit_316_cr = (16.0 * (1-margin), 18.0 * (1+margin)) # 14.4 ~ 19.8
    limit_316_ni = (10.0 * (1-margin), 14.0 * (1+margin)) # 9.0 ~ 15.4
    limit_316_mo = (2.0 * (1-margin), 3.0 * (1+margin))   # 1.8 ~ 3.3

    is_316 = (
        (limit_316_cr[0] <= cr <= limit_316_cr[1]) and
        (limit_316_ni[0] <= ni <= limit_316_ni[1]) and
        (limit_316_mo[0] <= mo <= limit_316_mo[1])
    )
    if is_316:
        return "SUS 316", "적합"
    
    # ----------------------------------------------------
    # [2] Duplex (이상) 판정
    # Cr:22~23 / Mo:3~3.5 / Ni:4.5~6.5 / Mn:2.0이하
    # ----------------------------------------------------
    limit_dup_cr = (22.0 * (1-margin), 23.0 * (1+margin)) # 19.8 ~ 25.3
    limit_dup_mo = (3.0 * (1-margin), 3.5 * (1+margin))   # 2.7 ~ 3.85
    limit_dup_ni = (4.5 * (1-margin), 6.5 * (1+margin))   # 4.05 ~ 7.15
    limit_dup_mn_max = 2.0 * (1+margin)                   # 2.2 이하

    is_duplex = (
        (limit_dup_cr[0] <= cr <= limit_dup_cr[1]) and
        (limit_dup_ni[0] <= ni <= limit_dup_ni[1]) and
        (limit_dup_mo[0] <= mo <= limit_dup_mo[1]) and
        (mn <= limit_dup_mn_max)
    )
    if is_duplex:
        return "Duplex", "적합"

    # ----------------------------------------------------
    # [3] SUS 310 (310S) 판정
    # Cr:24~26 / Ni:19~22 (고온용 고Cr, 고Ni)
    # ----------------------------------------------------
    limit_310_cr = (24.0 * (1-margin), 26.0 * (1+margin)) # 21.6 ~ 28.6
    limit_310_ni = (19.0 * (1-margin), 22.0 * (1+margin)) # 17.1 ~ 24.2
    
    is_310 = (
        (limit_310_cr[0] <= cr <= limit_310_cr[1]) and
        (limit_310_ni[0] <= ni <= limit_310_ni[1])
    )
    if is_310:
        return "SUS 310", "적합"

    # ----------------------------------------------------
    # [4] SUS 304 판정
    # Cr:18↑ / Ni:8↑ / Mo:0.5↓ / Mn:2.0↓
    # ----------------------------------------------------
    std_304_cr_min = 18.0
    std_304_ni_min = 8.0
    std_304_mo_max = 0.5 
    std_304_mn_max = 2.0

    limit_304_cr_min = std_304_cr_min * (1 - margin) # 16.2 이상
    limit_304_ni_min = std_304_ni_min * (1 - margin) # 7.2 이상
    limit_304_mo_max = std_304_mo_max * (1 + margin) # 0.55 이하
    limit_304_mn_max = std_304_mn_max * (1 + margin) # 2.2 이하

    is_304 = (
        (cr >= limit_304_cr_min) and 
        (ni >= limit_304_ni_min) and 
        (mo <= limit_304_mo_max) and
        (mn <= limit_304_mn_max)
    )
    if is_304:
        return "SUS 304", "적합"

    # ----------------------------------------------------
    # [5] 탈락 사유 분석 (Others)
    # ----------------------------------------------------
    reasons = []
    
    if cr == 0 and ni == 0:
        return "Others", "값 인식 실패(0.0)"

    if cr < limit_304_cr_min: reasons.append(f"Cr낮음({cr}<{limit_304_cr_min:.1f})")
    if ni < limit_304_ni_min: 
        reasons.append(f"Ni낮음(304기준 {ni}<{limit_304_ni_min:.1f})")
    if mo > limit_304_mo_max: reasons.append(f"Mo초과({mo}>{limit_304_mo_max:.2f})")
    if mn > limit_304_mn_max: reasons.append(f"Mn초과({mn}>{limit_304_mn_max:.1f})")

    if not reasons:
        return "Others", "성분 불일치"
        
    return "Others", ", ".join(reasons)

# ========================================================
# [메인] 실행 로직 - 파일 직접 선택
# ========================================================
root = tk.Tk()
root.withdraw()  # 메인 창 숨기기
root.attributes('-topmost', True) # 창을 맨 위로

print("📂 분석할 RFI 파일을 선택해주세요...")
target_file = filedialog.askopenfilename(
    title="분석할 RFI 파일을 선택하세요",
    initialdir=folder_path,
    filetypes=[("Excel Files", "*.xls*"), ("All Files", "*.*")]
)

if not target_file:
    print("❌ 파일이 선택되지 않았습니다. 종료합니다.")
    sys.exit()

# 선택된 파일이 속한 폴더를 작업 폴더로 업데이트 (결과 저장을 위해)
folder_path = os.path.dirname(target_file)
print(f"📂 분석 대상: {os.path.basename(target_file)}")

result_list_304 = []
result_list_316 = []
result_list_duplex = []
result_list_310 = []
result_list_others = []

try:
    xls = pd.ExcelFile(target_file)
    
    for sheet_name in xls.sheet_names:
        try:
            temp_df = pd.read_excel(target_file, sheet_name=sheet_name, header=None, nrows=20)
        except: continue
        
        header_idx = None
        for i, row in temp_df.iterrows():
            r_str = " ".join([str(x).upper() for x in row if pd.notna(x)])
            if "CR" in r_str and "NI" in r_str and "MO" in r_str:
                header_idx = i
                break
        
        if header_idx is None: continue
        
        df = pd.read_excel(target_file, sheet_name=sheet_name, header=header_idx)
        
        col_no = find_col(df, ["NO.", "NO", "SEQ", "NUM", "#", "RFI", "PMI", "ITEM"], exclude_keywords=["NOTE", "NOT"])
        col_cr = find_col(df, ["CR", "CHROMIUM", "크롬"], exclude_keywords=["UNIT"])
        col_ni = find_col(df, ["NI", "NICKEL", "니켈"], exclude_keywords=["UNIT"])
        col_mo = find_col(df, ["MO", "MOLYBDENUM", "몰리브덴"], exclude_keywords=["UNIT"])
        col_mn = find_col(df, ["MN", "MANGANESE", "망간"], exclude_keywords=["UNIT"])
        
        if not (col_cr and col_ni and col_mo): 
            print(f"   ⚠️ 시트 '{sheet_name}' 필수 컬럼 미확인으로 스킵")
            continue
        
        # 순번 컬럼 처리
        if col_no: 
            # 비교를 위해 문자열로 변환하고 앞뒤 공백 제거
            df['val_No'] = df[col_no].astype(str).str.strip()
        else: 
            df['val_No'] = ""

        # [필터링 로직] 사용자가 순번을 입력했다면 여기서 거릅니다.
        if target_no_list:
            # 엑셀의 순번이 사용자가 입력한 리스트에 있는 경우만 남김
            original_len = len(df)
            df = df[df['val_No'].isin(target_no_list)]
            filtered_len = len(df)
            if filtered_len == 0:
                # 해당 시트에 찾는 번호가 없으면 다음 시트로 넘어감
                continue
            
        df['val_Cr'] = df[col_cr].apply(to_float)
        df['val_Ni'] = df[col_ni].apply(to_float)
        df['val_Mo'] = df[col_mo].apply(to_float)
        if col_mn: df['val_Mn'] = df[col_mn].apply(to_float)
        else: df['val_Mn'] = 0.0

        df[['Detected_Grade', 'Reason']] = df.apply(
            lambda row: pd.Series(check_material_grade_with_reason(row)), axis=1
        )
        
        def make_result_df(sub_df, include_reason=False):
            out = pd.DataFrame()
            out['No.'] = sub_df['val_No']
            out['Ni (%)'] = sub_df['val_Ni'].round(1)
            out['Cr (%)'] = sub_df['val_Cr'].round(1)
            out['Mn (%)'] = sub_df['val_Mn'].round(1)
            out['Mo (%)'] = sub_df['val_Mo'].round(1)
            if include_reason:
                out['탈락 사유'] = sub_df['Reason']
            out['비고'] = f"출처: {sheet_name}"
            return out

        # [304 저장]
        df_304 = df[df['Detected_Grade'] == "SUS 304"].copy()
        if not df_304.empty:
            print(f"   👉 [시트: {sheet_name}] SUS 304: {len(df_304)}개")
            result_list_304.append(make_result_df(df_304))

        # [316 저장]
        df_316 = df[df['Detected_Grade'] == "SUS 316"].copy()
        if not df_316.empty:
            print(f"   👉 [시트: {sheet_name}] SUS 316: {len(df_316)}개")
            result_list_316.append(make_result_df(df_316))

        # [Duplex 저장]
        df_duplex = df[df['Detected_Grade'] == "Duplex"].copy()
        if not df_duplex.empty:
            print(f"   👉 [시트: {sheet_name}] Duplex: {len(df_duplex)}개")
            result_list_duplex.append(make_result_df(df_duplex))

        # [SUS 310 저장]
        df_310 = df[df['Detected_Grade'] == "SUS 310"].copy()
        if not df_310.empty:
            print(f"   👉 [시트: {sheet_name}] SUS 310: {len(df_310)}개")
            result_list_310.append(make_result_df(df_310))
            
        # [Others 저장]
        df_others = df[df['Detected_Grade'] == "Others"].copy()
        # 필터링 모드일 때는 값이 없어도 사용자가 요청한 번호라면 결과에 보여줌 (확인용)
        if target_no_list:
            pass # 그대로 둠
        else:
            df_others = df_others[ (df_others['val_Cr'] > 0) | (df_others['val_No'] != "") ]
        
        if not df_others.empty:
            print(f"   👉 [시트: {sheet_name}] 미분류: {len(df_others)}개")
            result_list_others.append(make_result_df(df_others, include_reason=True))

    # ====================================================
    # [저장]
    # ====================================================
    print("\n" + "="*50)
    if any([result_list_304, result_list_316, result_list_duplex, result_list_310, result_list_others]):
        save_name = "PMI_Material_Verified_Report.xlsx"
        save_path = os.path.join(folder_path, save_name)
        
        # [추가] 입력한 순서대로 정렬하는 유틸 함수
        def sort_by_input_order(df_to_sort):
            if target_no_list and not df_to_sort.empty:
                # 보조 정렬 컬럼 생성
                df_to_sort['sort_key'] = df_to_sort['No.'].astype(str).str.strip().apply(
                    lambda x: target_no_list.index(x) if x in target_no_list else 999999
                )
                df_to_sort = df_to_sort.sort_values('sort_key').drop(columns=['sort_key'])
            return df_to_sort

        while True:
            try:
                with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                    if result_list_304:
                        df_final = pd.concat(result_list_304, ignore_index=True)
                        sort_by_input_order(df_final).to_excel(writer, sheet_name='SUS 304', index=False)
                    if result_list_316:
                        df_final = pd.concat(result_list_316, ignore_index=True)
                        sort_by_input_order(df_final).to_excel(writer, sheet_name='SUS 316', index=False)
                    if result_list_duplex:
                        df_final = pd.concat(result_list_duplex, ignore_index=True)
                        sort_by_input_order(df_final).to_excel(writer, sheet_name='Duplex', index=False)
                    if result_list_310:
                        df_final = pd.concat(result_list_310, ignore_index=True)
                        sort_by_input_order(df_final).to_excel(writer, sheet_name='SUS 310', index=False)
                    if result_list_others:
                        df_final = pd.concat(result_list_others, ignore_index=True)
                        sort_by_input_order(df_final).to_excel(writer, sheet_name='Others (미분류)', index=False)
                
                print(f"💾 저장 완료 (입력 순서 반영됨): {save_path}")
                break

            except PermissionError:
                print(f"\n🚫 [오류] 엑셀 파일이 열려 있습니다. 닫고 엔터를 누르세요.")
                input()
                
    else:
        print("⚠️ 해당 데이터(순번)를 찾을 수 없습니다.")

    input("\n엔터 키를 누르면 종료합니다...")

except Exception as e:
    import traceback
    traceback.print_exc()
    input("오류 발생")