import re

file_path = 'src/Material-Master-Manager-V13.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Make NDT internal frames fill the width to align with the frame size
content = content.replace("company_frame.pack(anchor='w'", "company_frame.pack(fill='x'")
content = content.replace("header_frame.pack(anchor='w'", "header_frame.pack(fill='x'")
content = content.replace("grid_frame.pack(anchor='w'", "grid_frame.pack(fill='x'")

# Make NDT grid entries stretch within their uniform columns
# Looking for the NDT entry grid line in add_ndt_company_section
content = content.replace(
    "e.grid(row=r, column=c+1, padx=1, pady=1, sticky='w')",
    "e.grid(row=r, column=c+1, padx=1, pady=1, sticky='ew')"
)

# Make RTK grid entries stretch within their uniform columns
# Looking for the RTK entry grid line in setup_daily_usage_tab
content = content.replace(
    "e.grid(row=r, column=col+1, padx=1, pady=1, sticky='w')",
    "e.grid(row=r, column=col+1, padx=1, pady=1, sticky='ew')"
)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
print('Content alignment updated to fill the frames.')
