import pandas as pd
import openpyxl
import os

db_path = r'c:\\Users\\jjch2\\Desktop\\보고서\\Project PROVIDENCE\\Request\\PMI\\Na-aba\\home\\data\\Material_Inventory.xlsx'
if not os.path.exists(db_path):
    print("DB not found")
    exit(1)

# Read all sheets to preserve them
excel_data = pd.read_excel(db_path, sheet_name=None)
df_mat = excel_data['Materials']

# Ensure MaterialID is checked robustly
def is_432(val):
    try:
        return str(val).strip().replace('.0', '') == '432'
    except:
        return False

if not df_mat['MaterialID'].apply(is_432).any():
    print("432 not found. Adding it...")
    # Create new row with dict to ensure types
    new_row = {
        'MaterialID': 432,
        '회사코드': '',
        '관리품번': '',
        '품목명': 'MT약품',
        'SN': '',
        '창고': '현장',
        '모델명': '',
        '규격': '자동등록',
        '품목군코드': '',
        '공급업체': '',
        '제조사': '',
        '제조국': '',
        '가격': 0,
        '원가': 0,
        '관리단위': 'EA',
        '수량': 0,
        '재고하한': 0,
        '상태': '사용가능',
        'Active': 1
    }
    df_mat = pd.concat([df_mat, pd.DataFrame([new_row])], ignore_index=True)
    excel_data['Materials'] = df_mat
    
    # Save back
    with pd.ExcelWriter(db_path, engine='openpyxl') as writer:
        for sheet_name, df in excel_data.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    print("Database saved with MaterialID 432.")
else:
    print("MaterialID 432 already exists.")
    # Ensure it's active
    mask = df_mat['MaterialID'].apply(is_432)
    df_mat.loc[mask, 'Active'] = 1
    excel_data['Materials'] = df_mat
    with pd.ExcelWriter(db_path, engine='openpyxl') as writer:
        for sheet_name, df in excel_data.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    print("Ensured 432 is Active and saved.")
