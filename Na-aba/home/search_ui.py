import sys

file_path = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\src\main.py"
try:
    with open(file_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
        for i, line in enumerate(lines):
            if '_create_preview_ui' in line or '전체선택' in line or '전체 선택' in line or 'action_buttons' in line:
                print(f"{i+1}: {line.strip()}")
except Exception as e:
    print(e)
