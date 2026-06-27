import zipfile
import os
import glob

# Find latest report
folder = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\data"
pattern = os.path.join(folder, "RT_Report_*.xlsx")
matches = glob.glob(pattern)
latest_report = max(matches, key=os.path.getmtime)

try:
    with zipfile.ZipFile(latest_report, 'r') as z:
        content = z.read('xl/drawings/drawing1.xml').decode('utf-8', errors='ignore')
        print(f"Content of drawing1.xml in report (last 1000 chars):\n{content[-1000:]}")
except Exception as e:
    print(f"ERROR: {e}")
