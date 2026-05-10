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
        print("Template files related to drawings:")
        for f in z.namelist():
            if 'drawings' in f or 'media' in f or '_rels' in f:
                print(f"  {f}")
except Exception as e:
    print(f"ERROR: {e}")
