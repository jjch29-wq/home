import zipfile
import os
import glob
import shutil

def remove_custom_props(xlsx_path, new_name):
    temp_path = new_name + ".tmp"
    try:
        with zipfile.ZipFile(xlsx_path, 'r') as zin:
            with zipfile.ZipFile(temp_path, 'w') as zout:
                if zin.comment:
                    zout.comment = zin.comment
                for item in zin.infolist():
                    if item.filename != "docProps/custom.xml" and "custom.xml" not in item.filename:
                        zout.writestr(item, zin.read(item.filename))
                    else:
                        print(f"Removed custom prop from {xlsx_path}")
        if os.path.exists(new_name):
            os.remove(new_name)
        os.rename(temp_path, new_name)
        print(f"Saved fixed template to {new_name}")
    except Exception as e:
        print("Error processing", xlsx_path, e)

if __name__ == "__main__":
    t1 = r"C:\Users\-\PMI\home\src\templates\인원_장비투입_동탄양식.xlsx"
    t2 = r"C:\Users\-\PMI\home\src\templates\인원_장비투입_기본양식.xlsx"
    
    out1 = r"C:\Users\-\PMI\home\src\templates\양식_동탄.xlsx"
    out2 = r"C:\Users\-\PMI\home\src\templates\양식_기본.xlsx"
    
    remove_custom_props(t1, out1)
    remove_custom_props(t2, out2)
    print("Done")
