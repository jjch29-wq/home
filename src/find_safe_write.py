import sys

with open(r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\src\daily_work_report_manager.py", "r", encoding="utf-8") as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "safe_write" in line and "(" in line:
        print(f"Line {i+1}: {line.strip()}")
        # Print next 2 lines if they exist
        if i+1 < len(lines):
            print(f"  +1: {lines[i+1].strip()}")
        if i+2 < len(lines):
            print(f"  +2: {lines[i+2].strip()}")
