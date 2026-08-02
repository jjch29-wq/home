import pandas as pd
import sys

sys.stdout.reconfigure(encoding='utf-8')
filepath = r"C:\Users\jjch2\Desktop\산출내역서(가산~가평 천연가스 공급시설 건설공사 비파괴검사 기술용역)_감독경유.xlsx"

try:
    df = pd.read_excel(filepath, sheet_name='2. 산출명세서', header=None)
    
    # Let's search for the titles and print the rows around them
    # "가. 방사선 투과검사(RT) - 열배관"
    
    for idx, row in df.iterrows():
        text = str(row.values).replace('\n', '')
        if "방사선" in text or "초음파" in text or "침투" in text:
            print(f"Row {idx}: {row[0]} | {row[1]} | {row[2]}")
            
    # Also print any row that has a large number > 10,000,000 in the last column
    print("\n--- Large Amounts ---")
    for idx, row in df.iterrows():
        try:
            val = row.dropna().iloc[-1]
            if isinstance(val, (int, float)) and val > 10000000:
                print(f"Row {idx}: {row[0]} => {val}")
        except:
            pass

except Exception as e:
    print(f"Error: {e}")
