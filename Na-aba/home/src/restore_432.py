import pandas as pd
import os

db_path = r'c:\\Users\\jjch2\\Desktop\\보고서\\Project PROVIDENCE\\Request\\PMI\\Na-aba\\home\\data\\Material_Inventory.xlsx'
if os.path.exists(db_path):
    with pd.ExcelWriter(db_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_mat = pd.read_excel(db_path, sheet_name='Materials')
        
        # Ensure MaterialID is treated as a clean string for checking
        def clean_id(val):
            return str(val).strip().replace('.0', '')
            
        current_ids = [clean_id(x) for x in df_mat['MaterialID'].tolist()]
        
        if '432' not in current_ids:
            print("Restoring MaterialID 432...")
            new_row = pd.DataFrame([{
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
            }])
            df_mat = pd.concat([df_mat, new_row], ignore_index=True)
        else:
            print("MaterialID 432 already exists, ensuring it is Active.")
            df_mat.loc[df_mat['MaterialID'].astype(str).str.strip().str.replace('.0', '', regex=False) == '432', 'Active'] = 1
            
        df_mat.to_excel(writer, sheet_name='Materials', index=False)
        print("Database updated.")
else:
    print("Database not found.")
