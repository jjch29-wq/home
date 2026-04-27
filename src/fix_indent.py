import os

path = r'Material-Master-Manager-V13.py'
with open(path, 'r', encoding='utf-8', errors='ignore') as f:
    lines = f.readlines()

# Find the bad line around 12477 and fix its indentation
# The line has 24 spaces before 'default_filename' but should have 12 spaces
fixed = 0
for i in range(12460, 12500):
    if i < len(lines) and 'default_filename' in lines[i] and lines[i].startswith('                        '):
        print(f"Fixing indentation at line {i+1}: {repr(lines[i][:50])}")
        lines[i] = lines[i].replace('                        default_filename', '            default_filename', 1)
        fixed += 1
        break

if fixed:
    with open(path, 'w', encoding='utf-8', errors='ignore') as f:
        f.writelines(lines)
    print("SUCCESS: Indentation fixed")
else:
    print("FAILED: Could not find target line")
