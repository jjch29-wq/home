import os

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\ndt_billing_tab.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Update content_frame
old_content_frame = """        content_frame = ttk.Frame(billing_container)
        content_frame.pack(fill=tk.BOTH, expand=True)"""
new_content_frame = """        content_frame = ttk.Frame(billing_container)
        content_frame.pack(fill=tk.BOTH, expand=True)
        content_frame.columnconfigure(0, weight=2)  # Left frame takes 2/3
        content_frame.columnconfigure(1, weight=1)  # Right frame takes 1/3
        content_frame.rowconfigure(0, weight=1)"""
if old_content_frame in content:
    content = content.replace(old_content_frame, new_content_frame)
else:
    print("Failed to find content_frame layout")

# 2. Update contract_frame.pack -> grid
old_contract_pack = """contract_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))"""
new_contract_grid = """contract_frame.grid(row=0, column=0, sticky='nsew', padx=(0, 10))"""
if old_contract_pack in content:
    content = content.replace(old_contract_pack, new_contract_grid)
else:
    print("Failed to find contract_frame.pack")

# 3. Update exp_frame.pack -> grid
old_exp_pack = """exp_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)"""
new_exp_grid = """exp_frame.grid(row=0, column=1, sticky='nsew')"""
if old_exp_pack in content:
    content = content.replace(old_exp_pack, new_exp_grid)
else:
    print("Failed to find exp_frame.pack")

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)

print("Billing layout successfully updated to 2:1 ratio.")
