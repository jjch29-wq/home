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
        content = z.read('xl/worksheets/sheet1.xml').decode('utf-8', errors='ignore')
        if "<drawing" in content:
            start = content.find("<drawing")
            end = content.find("/>", start) + 2
            print(f"Template Drawing tag: {content[start:end]}")
            
        content = z.read('xl/worksheets/_rels/sheet1.xml.rels').decode('utf-8', errors='ignore')
        print(f"\nTemplate sheet1.xml.rels:\n{content}")
        
except Exception as e:
    print(f"ERROR: {e}")
