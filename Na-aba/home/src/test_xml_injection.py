import zipfile
import os
import shutil
import openpyxl
import glob

def inject_data_to_template(template_path, data_xlsx_path, output_path):
    shutil.copy2(template_path, output_path)
    with zipfile.ZipFile(data_xlsx_path, 'r') as z_data:
        ws_files = [f for f in z_data.namelist() if f.startswith('xl/worksheets/sheet') and f.endswith('.xml')]
        shared_strings = None
        if 'xl/sharedStrings.xml' in z_data.namelist():
            shared_strings = z_data.read('xl/sharedStrings.xml')
        workbook_xml = z_data.read('xl/workbook.xml')
        
        temp_out = output_path + ".tmp"
        with zipfile.ZipFile(output_path, 'r') as z_in:
            with zipfile.ZipFile(temp_out, 'w') as z_out:
                for item in z_in.infolist():
                    if item.filename in ws_files:
                        z_out.writestr(item, z_data.read(item.filename))
                    elif item.filename == 'xl/sharedStrings.xml' and shared_strings:
                        z_out.writestr(item, shared_strings)
                    elif item.filename == 'xl/workbook.xml':
                        z_out.writestr(item, workbook_xml)
                    else:
                        z_out.writestr(item, z_in.read(item.filename))
        os.remove(output_path)
        os.rename(temp_out, output_path)

if __name__ == "__main__":
    folder = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\data"
    template = glob.glob(os.path.join(folder, "RT KS*.xlsx"))[0]
    wb = openpyxl.load_workbook(template)
    ws = wb.worksheets[0]
    ws['B5'] = "INJECTED PROJECT NAME"
    wb.save("temp_data.xlsx")
    inject_data_to_template(template, "temp_data.xlsx", "patched_report.xlsx")
    with zipfile.ZipFile("patched_report.xlsx", 'r') as z:
        content = z.read('xl/drawings/drawing1.xml').decode('utf-8', errors='ignore')
        if "<xdr:grpSp" in content:
            print("SUCCESS: grpSp preserved!")
        else:
            print("FAILURE: grpSp lost")
    if os.path.exists("temp_data.xlsx"): os.remove("temp_data.xlsx")
    if os.path.exists("patched_report.xlsx"): os.remove("patched_report.xlsx")
