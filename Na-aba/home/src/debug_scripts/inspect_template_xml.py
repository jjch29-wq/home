import zipfile
import os
import glob

# Find template
folder = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\data"
pattern = os.path.join(folder, "RT KS*.xlsx")
matches = glob.glob(pattern)
if not matches:
    print("TEMPLATE NOT FOUND")
    exit(1)
template_path = matches[0]

try:
    with zipfile.ZipFile(template_path, 'r') as z:
        content = z.read('xl/drawings/drawing1.xml').decode('utf-8', errors='ignore')
        print(content)
except Exception as e:
    print(f"ERROR: {e}")
