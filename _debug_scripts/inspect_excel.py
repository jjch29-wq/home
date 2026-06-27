import pandas as pd
import glob
import os
import re

# Simulate the merge logic to see what sub_joint gives
SYNONYMS = {
    "joint": ["joint", "조인트", "jnt", "point", "포인트"],
    "reportno": ["reportno", "report_no", "성적서번호", "성적서", "보고서번호"]
}
NORM_SYNONYMS = {k: [re.sub(r'[^a-z0-9가-힣]', '', str(s).lower()).strip() for s in v] for k, v in SYNONYMS.items()}

def get_norm(c):
    return re.sub(r'[^a-z0-9가-힣]', '', str(c).lower()).strip()

all_data = []
for f in glob.glob("RT*.xls*"):
    if "Smart_Merged" in f or f.startswith("~"): continue
    try:
        df = pd.read_excel(f, sheet_name=0, header=25)
        df['Report No'] = f
        all_data.append(df)
    except:
        pass

if all_data:
    combined = pd.concat(all_data, ignore_index=True)
    joint_col = next((c for c in combined.columns if get_norm(c) == "joint" or any(get_norm(c) == s for s in NORM_SYNONYMS["joint"])), None)
    
    if joint_col:
        print(f"Joint col found: {joint_col}")
        for rep, group in combined.groupby('Report No'):
            valid = group[joint_col].dropna().astype(str).str.strip()
            valid = valid[valid != ""]
            print(f"Report: {rep}")
            print(f"Count: {valid.count()}, NuNique: {valid.nunique()}")
            if len(valid) > 0:
                print(f"First: {valid.iloc[0]}, Last: {valid.iloc[-1]}")
            print("-" * 30)
