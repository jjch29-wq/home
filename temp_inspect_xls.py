import pandas as pd
import sys

try:
    file_path = r'c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\resources\Template_DailyWorkReport.xls'
    df = pd.read_excel(file_path, header=None)
    
    print("CELL MAPPING EXTRACTED:")
    for r_idx, row in df.iterrows():
        for c_idx, val in enumerate(row):
            if pd.notnull(val):
                print(f"Row {r_idx}, Col {c_idx} (Cell {chr(65+c_idx)}{r_idx+1}): {val}")
except Exception as e:
    print(f"Error: {e}")
