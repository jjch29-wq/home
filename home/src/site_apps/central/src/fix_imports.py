import os
import glob

src_dir = r"c:\Users\jjch2\Desktop\PMI\home\src\site_apps\central\src"
files_to_check = glob.glob(os.path.join(src_dir, "**/*.py"), recursive=True)

for file in files_to_check:
    with open(file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    original = content
    
    # Replace get_history_path()
    content = content.replace("get_history_path()", "get_history_path()")
    
    if content != original:
        # Check if import is already there
        if "from site_apps.central.src.data_sync import" not in content:
            # add import after the other imports
            if "import os" in content:
                content = content.replace("import os", "import os\nfrom site_apps.central.src.data_sync import get_history_path, get_process_photos_dir")
            else:
                content = "from site_apps.central.src.data_sync import get_history_path, get_process_photos_dir\n" + content
                
        with open(file, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"Fixed {file}")
