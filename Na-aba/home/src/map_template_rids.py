import zipfile
import os
import glob

# Find template
folder = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\data"
pattern = os.path.join(folder, "RT KS*.xlsx")
matches = glob.glob(pattern)
template_path = matches[0]

try:
    with zipfile.ZipFile(template_path, 'r') as z:
        for f in z.namelist():
            if f.startswith('xl/worksheets/sheet') and f.endswith('.xml'):
                content = z.read(f).decode('utf-8', errors='ignore')
                import re
                rid_match = re.search(r'<drawing [^>]*r:id="(rId\d+)"', content)
                rid = rid_match.group(1) if rid_match else "None"
                print(f"{f}: Drawing ID = {rid}")
except Exception as e:
    print(f"ERROR: {e}")
