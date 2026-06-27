import os
import re

file_path = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\src\main.py"
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Instead of hardcoding ["Ni", "Cr", "Mo"], use a check that applies to any non-string key
# For PMI, the standard string keys are: No, Joint, Loc, Grade, Date, Dwg, _status, ST, V, selected, order_index
# Any other key in PMI is an element (Ni, Cr, Mo, Mn, C, Si, P, S, etc.)

replacement_str = 'key not in ["No", "Joint", "Loc", "Grade", "Date", "Dwg", "_status", "ST", "V", "selected", "order_index"]'

# We have lines like: if mode == "PMI" and key in ["Ni", "Cr", "Mo"]:
# We'll replace `key in ["Ni", "Cr", "Mo"]` with our replacement string.

content, count = re.subn(r'key\s+in\s+\["Ni",\s*"Cr",\s*"Mo"\]', replacement_str, content)

print(f"Replaced {count} instances.")

# Wait! There's also `key in ["Ni", "Cr", "Mo"]` check without mode == "PMI" in `copy_cell` maybe? No, it's usually `mode == "PMI" and ...`.
# Wait! In finish_edit:
# if mode == "PMI" and key in ["Ni", "Cr", "Mo"]:
# In _run_pmi_process, it explicitly loops over [('Ni', 8), ('Cr', 9), ('Mo', 10)]. This should remain hardcoded because the report only has those columns.

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
