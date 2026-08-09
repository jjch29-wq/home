import codecs
import re

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\한국지역난방 중앙지사.py'

with codecs.open(file_path, 'r', 'utf-8') as f:
    text = f.read()

old_block = re.search(r'def aggregate_dynamic_ndt\(site_df, method_filter\):.*?return results', text, flags=re.DOTALL)
if old_block:
    new_block = '''def aggregate_dynamic_ndt(site_df, method_filter):
                    """그룹화된 실적 데이터 생성 (PAUT, MT 등)"""
                    grouped = {}
                    for _, row in site_df.iterrows():
                        method = str(row.get('검사방법', '')).strip().upper()
                        if method != method_filter:
                            continue
                            
                        company = str(row.get('업체명', '')).strip()
                        if not company:
                            company = str(row.get('시공사', '')).strip()
                            
                        section = str(row.get('Section', '')).strip()
                        line_no = str(row.get('Line No.', '')).strip()
                        if not line_no and '도면번호' in row:
                            line_no = str(row.get('도면번호', '')).strip()
                            
                        inch = str(row.get('관경(Inch)', '')).strip()
                        
                        try:
                            joints = int(float(str(row.get('조인트수', 0)).replace(',', '') or 0))
                        except:
                            joints = 0
                            
                        spec = str(row.get('규격', '')).strip()
                        if not spec:
                            spec = str(row.get('작업형태', '')).strip()
                            
                        unit = str(row.get('단위', 'm')).strip()
                        if not unit or unit == 'nan':
                            unit = 'm'
                        
                        insp_type = str(row.get('검사구분', 'ORI')).strip().upper()
                        
                        try:
                            qty = float(str(row.get('검사량', row.get('Usage', 0))).replace(',', '') or 0)
                        except:
                            qty = 0.0
                            
                        key = (company, section, line_no, inch, spec, unit)
                        
                        if key not in grouped:
                            grouped[key] = {'joints': 0, 'ORI': 0.0, 'RE': 0.0, 'TOTAL': 0.0, 'remarks': set()}
                            
                        grouped[key]['joints'] += joints
                        
                        if insp_type == 'REP' or insp_type == 'RE':
                            grouped[key]['RE'] += qty
                        else:
                            grouped[key]['ORI'] += qty
                            
                        grouped[key]['TOTAL'] += qty
                        
                        note = str(row.get('비고', '')).strip()
                        if note and note != 'nan':
                            grouped[key]['remarks'].add(note)
                            
                    results = []
                    for idx, (k, v) in enumerate(grouped.items(), 1):
                        results.append({
                            'seq': idx,
                            'company': k[0],
                            'section': k[1],
                            'line_no': k[2],
                            'inch': k[3],
                            'joints': v['joints'],
                            'spec': k[4],
                            'unit': k[5],
                            'ori': v['ORI'],
                            're': v['RE'],
                            'total': v['TOTAL'],
                            'remark': ', '.join(v['remarks'])
                        })
                    return results'''
    
    text = text[:old_block.start()] + new_block + text[old_block.end():]
    
    with codecs.open(file_path, 'w', 'utf-8') as f:
        f.write(text)
    print('Replaced successfully.')
else:
    print('Could not find the function block.')
