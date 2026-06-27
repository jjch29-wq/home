import os

path = r'Material-Master-Manager-V13.py'
with open(path, 'r', encoding='utf-8', errors='ignore') as f:
    content = f.read()

# The mess starts with "# [NEW] NDT..." and ends with "pass  if row.get"
bad_block_start = "                        # [NEW] NDT"
bad_block_end = "except: pass  if row.get('Active', 1) == 0:"

if bad_block_start not in content:
    print("FAILED: bad_block_start not found")
    exit(1)
if bad_block_end not in content:
    print("FAILED: bad_block_end not found")
    exit(1)

parts = content.split(bad_block_start, 1)
prefix = parts[0]
rest = parts[1].split(bad_block_end, 1)
suffix = rest[1]

# Restore original lines
restored_middle = """                        # 휴면 계정 자동 활성화
                        if row.get('Active', 1) == 0:"""

new_content = prefix + restored_middle + suffix

with open(path, 'w', encoding='utf-8', errors='ignore') as f:
    f.write(new_content)
print("SUCCESS: File restored to original state at 12051+")
