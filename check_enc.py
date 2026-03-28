
import sys

target_file = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\MaterialManager-12.py"

with open(target_file, 'rb') as f:
    content = f.read()

lines = content.splitlines()

# Check line 18 (MARKER_PATTERN)
print(f"Line 18 raw: {lines[17]}")
try:
    print(f"Line 18 decode utf-8: {lines[17].decode('utf-8')}")
except Exception as e:
    print(f"Line 18 decode utf-8 error: {e}")

# Check line 5758 (default_worktimes)
print(f"Line 5758 raw: {lines[5757]}")
try:
    print(f"Line 5758 decode utf-8: {lines[5757].decode('utf-8')}")
except Exception as e:
    print(f"Line 5758 decode utf-8 error: {e}")
