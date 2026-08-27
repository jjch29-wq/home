
import pandas as pd
import sys
import codecs
sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer)
file_path = r'C:\Users\-\OneDrive\바탕 화면\산출내역서(가산~가평 천연가스 공급시설 건설공사 비파괴검사 기술용역)_감독경유.xlsx'
xls = pd.ExcelFile(file_path)
df = pd.read_excel(xls, sheet_name=5)
for idx, row in df.iterrows():
    print(f'Row {idx}: ' + ' | '.join([str(x) for x in row.values]))

