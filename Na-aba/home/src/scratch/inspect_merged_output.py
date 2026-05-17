import pandas as pd
import glob
import os

folder = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba"
merged_files = glob.glob(os.path.join(folder, "Final_Smart_Merged_v2.8_*.xlsx"))

if merged_files:
    latest_file = max(merged_files, key=os.path.getctime)
    print(f"Reading latest merged file: {os.path.basename(latest_file)}")
    df = pd.read_excel(latest_file)
    print(f"Merged DataFrame shape: {df.shape}")
    print("\nColumns inside the merged output:")
    print(list(df.columns))
    print("\nFirst 20 rows of the merged output:")
    print(df.head(20).to_string())
else:
    print("No merged file found!")
