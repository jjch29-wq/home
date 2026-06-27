import pandas as pd
import sys

sys.stdout.reconfigure(encoding='utf-8')
filepath = r"C:\Users\jjch2\Desktop\산출내역서(가산~가평 천연가스 공급시설 건설공사 비파괴검사 기술용역)_감독경유.xlsx"

try:
    df_summary = pd.read_excel(filepath, sheet_name='1. 용역비총괄표', header=None)
    print("\n--- Summary Sheet ---")
    print(df_summary.dropna(how='all').head(30).to_string(index=False, na_rep=''))
        
except Exception as e:
    print(f"Error: {e}")
