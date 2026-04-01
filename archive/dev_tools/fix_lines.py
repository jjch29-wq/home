
import os

target_file = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\MaterialManager-12.py"

with open(target_file, 'r', encoding='utf-8') as f:
    lines = f.readlines()

# Line numbers are 1-indexed in the tool, so subtract 1 for 0-indexed list access.
# Line 7003: Replace with marker_pattern = MARKER_PATTERN
lines[7002] = "        marker_pattern = MARKER_PATTERN\n"

# Line 7076: Replace with marker_pattern = MARKER_PATTERN
lines[7075] = "        marker_pattern = MARKER_PATTERN\n"

# Line 8761: Replace with val = str(ot_value).strip().replace(' ', '').replace('익일', '')
lines[8760] = "            val = str(ot_value).strip().replace(' ', '').replace('익일', '')\n"

with open(target_file, 'w', encoding='utf-8') as f:
    f.writelines(lines)

print("Successfully updated lines 7003, 7076, and 8761.")
