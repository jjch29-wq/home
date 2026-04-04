import os
import re

file_path = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\src\main.py"
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Pattern for the action bar buttons inside _setup_*_ui
# Looking for lines like:
# ttk.Button(action_bar, text=" 📝 데이터 추출 ", style="Action.TButton", command=self.extract_only).pack(side='left', fill='x', expand=True, padx=(0, 5))
# ttk.Button(action_bar, text=" ✨ 생성 시작 ", style="Action.TButton", command=self.run_process).pack(side='left', fill='x', expand=True)

import re

# Match a block containing the exact two buttons or similar (they all use extract_only and run_process)
# Let's replace any instance of side-by-side buttons with vertically stacked ones.

pattern = re.compile(
    r'(ttk\.Button\([^)]*command=self\.extract_only\)\.pack\()side=\'left\', fill=\'x\', expand=True, padx=\(0, 5\)(\))(\n\s*)(ttk\.Button\([^)]*command=self\.run_process\)\.pack\()side=\'left\', fill=\'x\', expand=True(\))'
)

def replacer(match):
    # Swap them so run_process is on top
    extract_btn_code = match.group(1).replace('style="Action.TButton", ', '') + "fill='x'" + match.group(2)
    run_btn_code = match.group(4) + "fill='x', pady=(0, 5)" + match.group(5)
    return run_btn_code + match.group(3) + extract_btn_code

new_content, count = pattern.subn(replacer, content)

if count > 0:
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(new_content)
    print(f"Replaced {count} instances.")
else:
    print("No matches found.")
