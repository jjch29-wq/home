import zipfile
import os
import glob

# Find latest report
folder = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\data"
pattern = os.path.join(folder, "RT_Report_*.xlsx")
matches = glob.glob(pattern)
if not matches:
    print("NO REPORTS FOUND")
    exit(1)
latest_report = max(matches, key=os.path.getmtime)

try:
    with zipfile.ZipFile(latest_report, 'r') as z:
        print(f"Inspecting report: {latest_report}")
        for f in z.namelist():
            if 'drawings' in f or 'media' in f or '_rels' in f:
                print(f"  {f}")
except Exception as e:
    print(f"ERROR: {e}")
