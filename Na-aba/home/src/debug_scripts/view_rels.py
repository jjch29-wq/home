import zipfile

report_path = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\data\RT_Report_20260510_203428.xlsx"

try:
    with zipfile.ZipFile(report_path, 'r') as z:
        content = z.read('xl/drawings/_rels/drawing1.xml.rels').decode('utf-8', errors='ignore')
        print(content)
except Exception as e:
    print(f"ERROR: {e}")
