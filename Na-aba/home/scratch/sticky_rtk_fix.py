import re

file_path = 'src/Material-Master-Manager-V13.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Update setup_daily_usage_tab to include a sticky_bottom_panel
content = content.replace(
    "self.entry_canvas.pack(side='left', fill='both', expand=True)",
    "self.entry_canvas.pack(side='left', fill='both', expand=True)\n        \n        # [NEW] Sticky Bottom Panel for NDT/RTK (Pinned to bottom of top pane)\n        self.sticky_bottom_panel = ttk.Frame(entry_frame)\n        self.sticky_bottom_panel.pack(side='bottom', fill='x', padx=2, pady=2)"
)

# 2. Move NDT and RTK creation to the sticky_bottom_panel
# First, remove the old grid placements in master_form_panel
content = content.replace(
    'self.ndt_frame = ttk.LabelFrame(self.master_form_panel, text="NDT 자재 소모량 (회사별)")',
    'self.ndt_frame = ttk.LabelFrame(self.sticky_bottom_panel, text="NDT 자재 소모량 (회사별)")'
)
content = content.replace(
    "self.ndt_frame.grid(row=1, column=0, padx=5, pady=2, sticky='ew')",
    "self.ndt_frame.pack(side='top', fill='x', padx=5, pady=2)"
)

content = content.replace(
    'rtk_grid = ttk.LabelFrame(self.master_form_panel, text="RTK 분류")',
    'rtk_grid = ttk.LabelFrame(self.sticky_bottom_panel, text="RTK 분류")'
)
content = content.replace(
    "rtk_grid.grid(row=2, column=0, padx=5, pady=2, sticky='ew')",
    "rtk_grid.pack(side='top', fill='x', padx=5, pady=2)"
)

# 3. Update Workers frame rowspan
content = content.replace(
    "workers_box_frame.grid(row=0, column=1, rowspan=3, sticky='nsew', padx=5, pady=5)",
    "workers_box_frame.grid(row=0, column=1, rowspan=1, sticky='nsew', padx=5, pady=5)"
)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
print('NDT and RTK moved to sticky bottom panel successfully.')
