import sys

# Read the report methods
with open(r'C:\Users\jjch2\.gemini\antigravity\brain\05c89134-70f9-4a38-b7d1-36123c292ba3\report_methods.py', 'r', encoding='utf-8') as f:
    methods_code = f.read()

# Read the MaterialManager.py file
with open(r'c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\MaterialManager.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

# Find the insertion point (after line 952: self.update_report_view())
insertion_line = 953  # Insert after line 953 (0-indexed: 952)

# Add proper indentation to methods code
indented_methods = '\n    '.join(methods_code.split('\n'))
methods_to_insert = '\n    ' + indented_methods + '\n'

# Insert the methods
new_lines = lines[:insertion_line] + [methods_to_insert] + lines[insertion_line:]

# Write back to file
with open(r'c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\MaterialManager.py', 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Successfully inserted update_report_view and export_report_to_excel methods!")
