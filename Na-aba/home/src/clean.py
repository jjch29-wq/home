import pandas as pd
import os

db_path = r'c:\\Users\\jjch2\\Desktop\\보고서\\Project PROVIDENCE\\Request\\PMI\\Na-aba\\home\\data\\Material_Inventory.xlsx'
if os.path.exists(db_path):
    with pd.ExcelWriter(db_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        # Update Materials
        df_mat = pd.read_excel(db_path, sheet_name='Materials')
        
        errant_mat_ids = df_mat[(df_mat['품목명'] == 'MT약품') & (df_mat['모델명'] == 'MT약품')]['MaterialID'].tolist()
        df_mat = df_mat[~df_mat['MaterialID'].isin(errant_mat_ids)]
        
        for idx, row in df_mat.iterrows():
            if row['품목명'] == 'MT약품' and row['모델명'] in ['백색페인트', '흑색자분', '형광자분']:
                df_mat.at[idx, '모델명'] = 'MT ' + str(row['모델명'])
            elif row['품목명'] == 'PT약품' and row['모델명'] in ['세척제', '침투제', '현상제', '형광침투제']:
                df_mat.at[idx, '모델명'] = 'PT ' + str(row['모델명'])
                
        df_mat.to_excel(writer, sheet_name='Materials', index=False)
        
        # Update Transactions
        df_trans = pd.read_excel(db_path, sheet_name='Transactions')
        df_trans = df_trans[~df_trans['MaterialID'].isin(errant_mat_ids)]
        df_trans.to_excel(writer, sheet_name='Transactions', index=False)
        
    print('Database cleaned successfully!')
