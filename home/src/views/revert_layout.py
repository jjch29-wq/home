import sys

file_path = 'g:/내 드라이브/07_Antigravity/PMI_한국지역난방/home/src/views/daily_usage_view.py'

with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Revert ndt_frame
content = content.replace(
    'self.ndt_frame = ttk.LabelFrame(form_content, text="NDT 자재 소모량 (회사별)")',
    'self.ndt_frame = ttk.LabelFrame(self.master_form_panel, text="NDT 자재 소모량 (회사별)")'
)
content = content.replace(
    "self.ndt_frame.grid(row=10, column=0, columnspan=4, sticky='ew', pady=(10, 2))",
    "self.ndt_frame.grid(row=1, column=1, padx=5, pady=2, sticky='new')"
)

# 2. Revert rtk_grid
content = content.replace(
    'self.rtk_grid = ttk.LabelFrame(form_content, text="RTK 분류")',
    'self.rtk_grid = ttk.LabelFrame(self.master_form_panel, text="RTK 분류")'
)
content = content.replace(
    "self.rtk_grid.grid(row=10, column=0, columnspan=4, sticky='ew', pady=(10, 2))",
    "self.rtk_grid.grid(row=1, column=1, padx=5, pady=2, sticky='new')"
)

# 3. Revert empty_guide_frame
content = content.replace(
    'self.empty_guide_frame = ttk.LabelFrame(form_content, text="PAUT / UT 검사 안내")',
    'self.empty_guide_frame = ttk.LabelFrame(self.master_form_panel, text="PAUT / UT 검사 안내")'
)
content = content.replace(
    "self.empty_guide_frame.grid(row=10, column=0, columnspan=4, sticky='ew', pady=(10, 2))",
    "self.empty_guide_frame.grid(row=1, column=1, padx=5, pady=2, sticky='new')"
)

# 4. Revert fixed_vehicle_frame
content = content.replace(
    'self.fixed_vehicle_frame = ttk.LabelFrame(self.master_form_panel, text="차량점검 (상시 패널)")\n    self.fixed_vehicle_frame.grid(row=1, column=1, padx=5, pady=2, sticky="new")',
    'self.fixed_vehicle_frame = ttk.LabelFrame(self.bottom_dashboard, text="차량점검 (상시 패널)")\n    self.bottom_dashboard.add(self.fixed_vehicle_frame, weight=9)'
)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
print('Reverted UI layout.')
