import glob, os, openpyxl, re
files = glob.glob(r'c:\Users\jjch2\Desktop\**\RT_Report*.xlsx', recursive=True)
files = [f for f in files if not os.path.basename(f).startswith('~$')]
if files:
    latest = max(files, key=os.path.getmtime)
    print('Generated File:', latest)
    import zipfile
    try:
        wb = openpyxl.load_workbook(latest)
        print('H1 Value:', repr(wb.worksheets[0]['H1'].value))
    except Exception as e:
        print('Error loading:', e)
        with zipfile.ZipFile(latest, 'r') as z:
            sheet_xml = z.read('xl/worksheets/sheet1.xml').decode('utf-8')
            m2 = re.search(r'<c r="H1"[^>]*>.*?<t>(.*?)</t>.*?</c>', sheet_xml)
            if m2: print('H1 inline:', m2.group(1))
