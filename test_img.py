import zipfile, re

file_path = r'C:\Users\jjch2\Desktop\PT_Report_20260916_171353.xlsx'
with zipfile.ZipFile(file_path, 'r') as z:
    for name in z.namelist():
        if name.startswith('xl/drawings/drawing') and name.endswith('.xml'):
            xml = z.read(name).decode('utf-8')
            rels_name = name.replace('drawings/', 'drawings/_rels/') + '.rels'
            
            try:
                rels = z.read(rels_name).decode('utf-8')
            except:
                print(f'{rels_name} is missing!')
                continue
                
            blip_ids = re.findall(r'<a:blip.*?r:embed=\x22(.*?)\x22', xml)
            for b_id in blip_ids:
                if f'Id="{b_id}"' not in rels:
                    print(f'{name}: Missing relation for {b_id} in {rels_name}')
                    
print('Image relationship checks done.')