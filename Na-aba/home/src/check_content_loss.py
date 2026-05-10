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
        print(f"Inspecting content of {latest_report}")
        for sheet_file in ['xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml']:
            if sheet_file in z.namelist():
                content = z.read(sheet_file).decode('utf-8', errors='ignore')
                print(f"\n--- {sheet_file} (first 500 chars) ---\n{content[:500]}")
                # Check for some cell values
                if '<v>' in content:
                    print("Found some values (<v> tag present)")
                else:
                    print("NO values found (<v> tag absent)")
except Exception as e:
    print(f"ERROR: {e}")
