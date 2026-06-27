import os

path = r'Material-Master-Manager-V13.py'
with open(path, 'r', encoding='utf-8', errors='ignore') as f:
    lines = f.readlines()

# Search for the broken line containing 'command=_save' around 11889
fixed = False
for i in range(11850, 11950): # Wider range to be safe
    if i < len(lines) and 'command=_save' in lines[i] and 'width=10' in lines[i]:
        print(f"Fixing line {i+1}: {repr(lines[i])}")
        lines[i] = '        ttk.Button(btn_frame, text="저장", command=_save, width=10).pack(side=\'left\', padx=8)\n'
        fixed = True
        break

if fixed:
    with open(path, 'w', encoding='utf-8', errors='ignore') as f:
        f.writelines(lines)
    print("SUCCESS: Line fixed")
else:
    print("FAILED: Could not find the target line to fix")
