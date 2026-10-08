import os
import glob
import re

src_dir = r"c:\Users\jjch2\Desktop\PMI\home\src\site_apps\central\src"

files_to_check = glob.glob(os.path.join(src_dir, "**/*.py"), recursive=True)

for file in files_to_check:
    with open(file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    original_content = content
    
    # 1. Replace get_history_path()
    # and its variants
    content = re.sub(
        r"os\.path\.join\(os\.path\.dirname\(os\.path\.abspath\(__file__\)\),\s*'daily_work_history\.json'\)",
        r"get_history_path()",
        content
    )
    
    # 2. Replace hardcoded absolute paths in paut_writer.py and ndt_section_writer.py
    content = re.sub(
        r"r'c:\\Users\\jjch2\\Desktop\\PMI\\home\\src\\daily_work_history\.json'",
        r"get_history_path()",
        content
    )
    
    if content != original_content:
        # Add import if data_sync isn't there? Actually __import__('data_sync') doesn't require import statement.
        # But wait, __import__('data_sync') works if it's in sys.path.
        # Since these are run from various directories, let's inject a safe import or just use relative.
        # Actually, if we just use `from site_apps.central.src.data_sync import get_history_path`, it might fail depending on sys.path.
        # Let's dynamically add the sys path if needed, or simply let the user run the script and we'll fix it if it errors.
        
        with open(file, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"Updated {file}")

print("Update completed.")
