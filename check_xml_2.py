import zipfile, re
with zipfile.ZipFile(r'c:\Users\jjch2\Desktop\PMI\test_merge_save_2.xlsx', 'r') as z:
    sheet_xml = z.read('xl/worksheets/sheet1.xml').decode('utf-8')
    m = re.search(r'<mergeCells[^>]*>(.*?)</mergeCells>', sheet_xml)
    if m:
        for ms in re.findall(r'<mergeCell ref="(.*?)"/>', m.group(1)):
            if 'H' in ms or 'O' in ms:
                print(ms)
