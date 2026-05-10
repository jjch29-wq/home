import glob
import os

folder = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\data"
pattern = os.path.join(folder, "RT KS*.xlsx")
matches = glob.glob(pattern)

if matches:
    print(f"FOUND: {matches[0]}")
else:
    print("NOT FOUND")
