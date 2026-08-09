import win32com.client
import os
import shutil

src = None
for f in os.listdir(r'C:\Users\jjch2\Desktop'):
    if '월용역' in f:
        src = os.path.join(r'C:\Users\jjch2\Desktop', f)
        break

if not src:
    print("File not found on Desktop!")
else:
    save_path = r'c:\Users\jjch2\Desktop\PMI\test_paut_output_3.xlsx'
    if os.path.exists(save_path): os.remove(save_path)
    shutil.copy2(src, save_path)
    excel = win32com.client.Dispatch('Excel.Application')
    excel.Visible = False
    wb = excel.Workbooks.Open(os.path.abspath(save_path))
    print([s.Name for s in wb.Sheets])
    wb.Close(False)
    excel.Quit()
