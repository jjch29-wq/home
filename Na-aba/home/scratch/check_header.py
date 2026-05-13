import pandas as pd
import sys

try:
    path = r'c:\Users\-\OneDrive\바탕 화면\home\Na-aba\home\data\가스공사 의뢰서.xlsx'
    xls = pd.ExcelFile(path)
    df = pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=None, nrows=20)
    with open('header_check.txt', 'w', encoding='utf-8') as f:
        f.write(df.to_string())
    print("SUCCESS")
except Exception as e:
    with open('header_check.txt', 'w', encoding='utf-8') as f:
        f.write(str(e))
    print("ERROR")
