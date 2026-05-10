import openpyxl
import os
import glob
import zipfile

# Find template
folder = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\data"
pattern = os.path.join(folder, "RT KS*.xlsx")
matches = glob.glob(pattern)
if not matches:
    print("TEMPLATE NOT FOUND")
    exit(1)
template_path = matches[0]

test_output = "test_no_images.xlsx"

try:
    print(f"Testing save with NO images added: {template_path}")
    # Load WITHOUT images (if possible) - no, openpyxl always loads them if present
    wb = openpyxl.load_workbook(template_path)
    # Don't add any images
    wb.save(test_output)
    
    with zipfile.ZipFile(test_output, 'r') as z:
        content = z.read('xl/drawings/drawing1.xml').decode('utf-8', errors='ignore')
        if "<xdr:grpSp" in content:
            print("Preserved grpSp")
        else:
            print("LOST grpSp")
except Exception as e:
    print(f"ERROR: {e}")
finally:
    if os.path.exists(test_output): os.remove(test_output)
