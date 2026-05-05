import re

file_path = 'src/Material-Master-Manager-V13.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Remove sticky_bottom_panel creation and packing
content = content.replace(
    "\n        \n        # [NEW] Sticky Bottom Panel for NDT/RTK (Pinned to bottom of top pane)\n        self.sticky_bottom_panel = ttk.Frame(entry_frame)\n        self.sticky_bottom_panel.pack(side='bottom', fill='x', padx=2, pady=2)",
    ""
)

# 2. Move NDT and RTK back to master_form_panel
content = content.replace(
    'self.ndt_frame = ttk.LabelFrame(self.sticky_bottom_panel, text="NDT 자재 소모량 (회사별)")',
    'self.ndt_frame = ttk.LabelFrame(self.master_form_panel, text="NDT 자재 소모량 (회사별)")'
)
content = content.replace(
    "self.ndt_frame.pack(side='top', fill='x', padx=5, pady=2)",
    "self.ndt_frame.grid(row=1, column=0, padx=5, pady=2, sticky='ew')"
)

content = content.replace(
    'rtk_grid = ttk.LabelFrame(self.sticky_bottom_panel, text="RTK 분류")',
    'rtk_grid = ttk.LabelFrame(self.master_form_panel, text="RTK 분류")'
)
content = content.replace(
    "rtk_grid.pack(side='top', fill='x', padx=5, pady=2)",
    "rtk_grid.grid(row=2, column=0, padx=5, pady=2, sticky='ew')"
)

# 3. Restore Workers frame rowspan
content = content.replace(
    "workers_box_frame.grid(row=0, column=1, rowspan=1, sticky='nsew', padx=5, pady=5)",
    "workers_box_frame.grid(row=0, column=1, rowspan=3, sticky='nsew', padx=5, pady=5)"
)

# 4. Revert toggle_sash_lock content_h calculation
content = content.replace(
    'content_h = (bbox[1] + bbox[3]) + (self.sticky_bottom_panel.winfo_height() if hasattr(self, "sticky_bottom_panel") else 0) + 40',
    'content_h = bbox[1] + bbox[3] + 40 # Add some padding'
)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
print('Reverted sticky bottom panel changes successfully.')
