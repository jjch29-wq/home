import win32com.client as win32
import os

hwp_path = os.path.abspath(r'C:\Users\-\OneDrive\바탕 화면\4.4.1 위험성평가표(RT).hwp')
out_path = r'c:\Users\-\OneDrive\바탕 화면\home\Na-aba\home\hwp_rt_out.txt'

try:
    hwp = win32.Dispatch('HWPFrame.HwpObject')
    hwp.RegisterModule('FilePathCheckDLL', 'SecurityModule')
    hwp.Open(hwp_path)
    hwp.InitScan()
    text = ""
    while True:
        ret, t = hwp.GetText()
        text += t
        if ret <= 1:
            break
    hwp.ReleaseScan()
    hwp.Quit()
    
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(text)
    print("HWP extracted successfully")
except Exception as e:
    print(f"Error extracting HWP: {e}")
