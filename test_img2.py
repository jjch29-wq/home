import zipfile, re

file_path = r'C:\Users\jjch2\Desktop\PT_Report_20260916_171353.xlsx'
with zipfile.ZipFile(file_path, 'r') as z:
    for name in z.namelist():
        if name.startswith('xl/worksheets/sheet') and name.endswith('.xml'):
            xml = z.read(name).decode('utf-8')
            rels_name = name.replace('worksheets/', 'worksheets/_rels/') + '.rels'
            
            has_drawing_tag = bool(re.search(r'<drawing r:id=', xml))
            
            try:
                rels = z.read(rels_name).decode('utf-8')
                has_drawing_rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing' in rels
            except:
                has_drawing_rel = False
                
            if has_drawing_rel and not has_drawing_tag:
                print(f'{name}: HAS RELATION BUT NO DRAWING TAG IN XML! CORRUPTION FOUND!')
                