import zipfile
import os
import glob

def remove_custom_props(xlsx_path):
    temp_path = xlsx_path + ".tmp"
    with zipfile.ZipFile(xlsx_path, 'r') as zin:
        with zipfile.ZipFile(temp_path, 'w') as zout:
            zout.comment = zin.comment # preserve the comment
            for item in zin.infolist():
                if item.filename != "docProps/custom.xml" and "custom.xml" not in item.filename:
                    zout.writestr(item, zin.read(item.filename))
                else:
                    print(f"Removed custom prop from {xlsx_path}")
    os.remove(xlsx_path)
    os.rename(temp_path, xlsx_path)

if __name__ == "__main__":
    templates = glob.glob(r"C:\Users\-\PMI\home\src\templates\*.xlsx")
    for t in templates:
        remove_custom_props(t)
    print("Done fixing templates.")
