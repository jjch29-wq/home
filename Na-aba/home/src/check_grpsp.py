import zipfile

report_path = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\data\RT_Report_20260510_205846.xlsx"

try:
    with zipfile.ZipFile(report_path, 'r') as z:
        content = z.read('xl/drawings/drawing1.xml').decode('utf-8', errors='ignore')
        if "<xdr:grpSp" in content:
            print("FOUND grpSp in report")
        else:
            print("NOT FOUND grpSp in report")
except Exception as e:
    print(f"ERROR: {e}")
