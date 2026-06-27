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
        content = z.read('xl/worksheets/sheet1.xml').decode('utf-8', errors='ignore')
        if "<drawing" in content:
            start = content.find("<drawing")
            end = content.find("/>", start) + 2
            print(f"Drawing tag in sheet1.xml: {content[start:end]}")
        else:
            print("No drawing tag in sheet1.xml")
            
    with zipfile.ZipFile(latest_report, 'r') as z:
        content = z.read('xl/worksheets/_rels/sheet1.xml.rels').decode('utf-8', errors='ignore')
        print(f"\nContent of sheet1.xml.rels:\n{content}")
        
except Exception as e:
    print(f"ERROR: {e}")
