import zipfile

report_path = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\data\RT_Report_20260510_203428.xlsx"

try:
    with zipfile.ZipFile(report_path, 'r') as z:
        content = z.read('xl/drawings/drawing1.xml').decode('utf-8', errors='ignore')
        print("drawing1.xml content summary:")
        print(f"AbsoluteAnchor count: {content.count('<xdr:absoluteAnchor')}")
        print(f"OneCellAnchor count: {content.count('<xdr:oneCellAnchor')}")
        print(f"TwoCellAnchor count: {content.count('<xdr:twoCellAnchor')}")
except Exception as e:
    print(f"ERROR: {e}")
