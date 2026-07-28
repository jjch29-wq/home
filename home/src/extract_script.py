import ast
import os

source_file = 'home/src/자재작업일보기성서류Ver.2.py'

with open(source_file, 'r', encoding='utf-8') as f:
    lines = f.readlines()
    
tree = ast.parse(''.join(lines))
m = next(n for n in tree.body if isinstance(n, ast.ClassDef) and n.name == 'MaterialManager')

def get_method_lines(name):
    func = next((n for n in m.body if isinstance(n, ast.FunctionDef) and n.name == name), None)
    if not func: return None
    return (func.lineno - 1, func.end_lineno)

def extract(name):
    bounds = get_method_lines(name)
    if not bounds: return None
    start, end = bounds
    block_lines = lines[start:end]
    new_block = []
    for line in block_lines:
        if line.strip().startswith(f'def {name}(self'):
            new_line = line.replace(f'def {name}(self', f'def {name}_impl(self')
            if new_line.startswith('    '):
                new_line = new_line[4:]
            new_block.append(new_line)
        else:
            new_block.append(line[4:] if line.startswith('    ') else line)
    return ''.join(new_block) + '\n\n', start, end

ld = extract('load_data')
sd = extract('save_data')

dl_header = 'import pandas as pd\nimport json\nimport os\nimport traceback\nimport tkinter as tk\nfrom tkinter import messagebox\nfrom datetime import datetime\n\n'
with open('home/src/services/data_loader.py', 'w', encoding='utf-8') as f:
    f.write(dl_header + ld[0] + sd[0])

ed = extract('export_daily_work_report')
em = extract('export_materials')
se = extract('save_df_to_excel_autofit')

ex_header = 'import pandas as pd\nimport os\nimport traceback\nimport tkinter as tk\nfrom tkinter import messagebox, filedialog\nfrom datetime import datetime\nimport openpyxl\nfrom openpyxl.styles import Alignment, Border, Side, PatternFill, Font\nfrom openpyxl.utils import get_column_letter\nfrom utils.helpers import normalize_id\nimport json\nimport sys\nimport subprocess\n\n'
with open('home/src/services/excel_exporter.py', 'w', encoding='utf-8') as f:
    f.write(ex_header + ed[0] + em[0] + se[0])

import py_compile
try:
    py_compile.compile('home/src/services/data_loader.py', doraise=True)
    py_compile.compile('home/src/services/excel_exporter.py', doraise=True)
    print('Both compiled successfully.')
except Exception as e:
    print('Compile error:', e)

funcs = [
    ('load_data', ld[1], ld[2]),
    ('save_data', sd[1], sd[2]),
    ('export_daily_work_report', ed[1], ed[2]),
    ('export_materials', em[1], em[2]),
    ('save_df_to_excel_autofit', se[1], se[2])
]
funcs.sort(key=lambda x: x[1], reverse=True)

for name, start, end in funcs:
    stub = [
        f'    def {name}(self, *args, **kwargs):\n',
        f'        if "export" in "{name}" or "excel" in "{name}":\n',
        f'            from services.excel_exporter import {name}_impl\n',
        f'        else:\n',
        f'            from services.data_loader import {name}_impl\n',
        f'        return {name}_impl(self, *args, **kwargs)\n'
    ]
    lines = lines[:start] + stub + lines[end:]

with open(source_file, 'w', encoding='utf-8') as f:
    f.writelines(lines)
print('Source file updated successfully.')
