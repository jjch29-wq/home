import glob, os, zipfile, re

files = glob.glob(r'c:\Users\jjch2\Desktop\**\RT_Report*.xlsx', recursive=True)
files = [f for f in files if not os.path.basename(f).startswith('~$')]
if files:
    latest = max(files, key=os.path.getmtime)
    print('Generated File:', latest)
    try:
        with zipfile.ZipFile(latest, 'r') as z:
            sheet_xml = z.read('xl/worksheets/sheet1.xml').decode('utf-8')
            m = re.search(r'<mergeCells[^>]*>(.*?)</mergeCells>', sheet_xml)
            if m:
                merges = m.group(1)
                for merge_str in re.findall(r'<mergeCell ref="(.*?)"/>', merges):
                    if 'H' in merge_str or 'O' in merge_str:
                        print(merge_str)
            else:
                print('No mergeCells tag found')
    except Exception as e:
        print('Error:', e)
