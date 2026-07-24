import zipfile
import os
import glob
import shutil

def remove_custom_props(xlsx_path):
    temp_path = xlsx_path + ".tmp"
    with zipfile.ZipFile(xlsx_path, 'r') as zin:
        with zipfile.ZipFile(temp_path, 'w') as zout:
            zout.comment = zin.comment
            for item in zin.infolist():
                if item.filename != "docProps/custom.xml" and "custom.xml" not in item.filename:
                    zout.writestr(item, zin.read(item.filename))
                else:
                    print(f"Removed custom prop from {xlsx_path}")
    
    # Use shutil.move with overwrite instead of os.remove to avoid permission issues if possible, 
    # but since it's the same file, it still requires write access.
    # If the file is open, we can't overwrite.
    # Let's write to a new name.
    new_name = xlsx_path.replace(".xlsx", "_fixed.xlsx")
    shutil.move(temp_path, new_name)
    print(f"Saved fixed template to {new_name}")

if __name__ == "__main__":
    templates = glob.glob(r"C:\Users\-\PMI\home\src\templates\*.xlsx")
    # Avoid processing already fixed templates
    templates = [t for t in templates if "_fixed" not in t]
    for t in templates:
        remove_custom_props(t)
        
    print("Done fixing templates.")
