import zipfile
import os

# Find template
template_path = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\template\RT_Report_Template.xlsx"

try:
    with zipfile.ZipFile(template_path, 'r') as z:
        if 'xl/drawings/drawing1.xml' in z.namelist():
            content = z.read('xl/drawings/drawing1.xml').decode('utf-8', errors='ignore')
            print(f"Top 2000 chars of template drawing1.xml:\n{content[:2000]}")
            import re
            tags = re.findall(r'<xdr:([a-zA-Z0-9]+)', content)
            print(f"\nTags found: {set(tags)}")
except Exception as e:
    print(f"ERROR: {e}")
