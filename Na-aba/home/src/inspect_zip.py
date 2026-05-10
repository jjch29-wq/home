import zipfile
import os

report_path = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\data\RT_Report_20260510_203428.xlsx"

try:
    with zipfile.ZipFile(report_path, 'r') as z:
        print("Files in ZIP:")
        for f in z.namelist():
            if "drawing" in f:
                print(f"  {f}")
                content = z.read(f).decode('utf-8', errors='ignore')
                print(f"  Length: {len(content)}")
                # Check for anchors
                if "<xdr:absoluteAnchor" in content: print("    Found absoluteAnchor")
                if "<xdr:oneCellAnchor" in content: print("    Found oneCellAnchor")
                if "<xdr:twoCellAnchor" in content: print("    Found twoCellAnchor")
except Exception as e:
    print(f"ERROR: {e}")
