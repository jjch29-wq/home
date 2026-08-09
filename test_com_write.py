import win32com.client
import os
import json
import shutil

template = r'c:\Users\jjch2\Desktop\PMI\home\data\기성서류_기본양식.xlsx'
save_path = r'c:\Users\jjch2\Desktop\PMI\test_paut_save.xlsx'

if os.path.exists(save_path):
    os.remove(save_path)
shutil.copy2(template, save_path)

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
wb = excel.Workbooks.Open(os.path.abspath(save_path))

try:
    with open(r'c:\Users\jjch2\Desktop\PMI\home\src\daily_work_history.json', 'r', encoding='utf-8') as f:
        history = json.load(f)

    paut_records = []
    target_month_str = "2026-08"

    for date_key, log_data in history.items():
        if date_key.startswith(target_month_str):
            ndt_results = log_data.get('ndt_results', [])
            for r in ndt_results:
                if str(r.get('검사방법', '')).strip().upper() == 'PAUT':
                    r['_date'] = date_key
                    paut_records.append(r)

    paut_records.sort(key=lambda x: x['_date'])
    
    groups = {}
    for r in paut_records:
        key = (str(r.get('업체', '')), str(r.get('구간', '')), str(r.get('라인번호', '')), str(r.get('관경', '')), str(r.get('Joint No.', '')))
        paut_val = str(r.get('PAUT', '0')).strip()
        try: val = float(paut_val)
        except: val = 0.0
        if val == 0: continue
        shift = str(r.get('규격', '주간')).strip()
        
        if key not in groups:
            groups[key] = {'ORI': 0.0, 'RE': 0.0, 'shift': shift}
        if groups[key]['ORI'] == 0.0:
            groups[key]['ORI'] += val
        else:
            groups[key]['RE'] += val

    paut_sheet_name = '3. 비파괴검사 현황 (열배관)'
    ws_paut = wb.Sheets(paut_sheet_name)
    
    print(f"Groups count: {len(groups)}")
    if len(groups) > 1:
        for _ in range(len(groups) - 1):
            ws_paut.Rows(406).Insert(Shift=-4121, CopyOrigin=0)
            
    current_row = 405
    for i, (key, data) in enumerate(groups.items()):
        업체, 구간, 라인번호, 관경, joint = key
        ws_paut.Cells(current_row, 2).Value = 업체
        ws_paut.Cells(current_row, 3).Value = i + 1
        ws_paut.Cells(current_row, 4).Value = 구간
        ws_paut.Cells(current_row, 5).Value = 라인번호
        ws_paut.Cells(current_row, 6).Value = 관경
        ws_paut.Cells(current_row, 7).Value = joint
        ws_paut.Cells(current_row, 8).Value = data['shift']
        ws_paut.Cells(current_row, 9).Value = 'M'
        
        ori = data['ORI']
        re_val = data['RE']
        tot = ori + re_val
        
        if ori > 0: ws_paut.Cells(current_row, 10).Value = round(ori, 4)
        if re_val > 0: ws_paut.Cells(current_row, 11).Value = round(re_val, 4)
        if tot > 0: ws_paut.Cells(current_row, 12).Value = round(tot, 4)
        current_row += 1
        
    print("Writing finished without errors.")
except Exception as e:
    import traceback
    print("Error:")
    traceback.print_exc()
finally:
    wb.Save()
    wb.Close(False)
    excel.Quit()
