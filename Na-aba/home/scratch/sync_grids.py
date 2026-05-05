import re

file_path = 'src/Material-Master-Manager-V13.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Synchronize NDT and RTK grid columns
content = content.replace('grid_frame = ttk.Frame(company_frame)', 
    'grid_frame = ttk.Frame(company_frame)\n        for c in range(6): grid_frame.columnconfigure(c, weight=1, uniform="ndt_rtk")')

content = content.replace('self.rtk_entries = {}', 
    'for c in range(6): rtk_grid.columnconfigure(c, weight=1, uniform="ndt_rtk")\n        self.rtk_entries = {}')

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
print('NDT and RTK grid columns synchronized.')
