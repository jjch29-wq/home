import pandas as pd
import sys

file_path = r'C:\Users\-\OneDrive\바탕 화면\산출내역서(가산~가평 천연가스 공급시설 건설공사 비파괴검사 기술용역)_감독경유.xlsx'

try:
    df = pd.read_excel(file_path, sheet_name='6-1. RT', header=None)
    with open(r'c:\Users\-\OneDrive\바탕 화면\home_new\home\src\rt_data.txt', 'w', encoding='utf-8') as f:
        for index, row in df.iterrows():
            row_vals = [str(x) if not pd.isna(x) else '' for x in row.values]
            f.write('\t'.join(row_vals) + '\n')
    print('Data extracted successfully')
except Exception as e:
    print('Error:', e)
