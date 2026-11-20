import openpyxl
import json
import shutil
import os
import datetime

def migrate():
    excel_path = 'home/src/site_apps/central/data/Material_Inventory.xlsx'
    backup_dir = 'home/src/site_apps/central/data/backups'
    os.makedirs(backup_dir, exist_ok=True)
    
    timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_path = os.path.join(backup_dir, f'Material_Inventory_pre_contractor_migration_{timestamp}.xlsx')
    shutil.copy2(excel_path, backup_path)
    print(f"1. 백업 생성 완료: {backup_path}")
    
    # Load billing records and daily work history
    with open('home/src/site_apps/central/src/config.json', 'r', encoding='utf-8') as f:
        cfg = json.load(f)
    billing_records = cfg.get('BILLING_RECORDS', [])
    
    with open('home/src/site_apps/central/src/daily_work_history.json', 'r', encoding='utf-8') as f:
        work_history = json.load(f)
        
    wb = openpyxl.load_workbook(excel_path)
    ws = wb['DailyUsage']
    
    print("\n2. DailyUsage 행별 시공사/발주처 마이그레이션 시작...")
    modified_count = 0
    
    for row_idx in range(2, ws.max_row + 1):
        date_val = ws.cell(row=row_idx, column=1).value
        if not date_val:
            continue
            
        date_str = str(date_val)[:10]
        method = str(ws.cell(row=row_idx, column=21).value or '').strip()
        try:
            qty = float(ws.cell(row=row_idx, column=16).value or 0)
        except (ValueError, TypeError):
            qty = 0.0
            
        old_client = ws.cell(row=row_idx, column=65).value # Col 64 (0-based) = Col 65 (1-based, 회사코드)
        old_contractor = ws.cell(row=row_idx, column=66).value # Col 65 (0-based) = Col 66 (1-based, 업체명)
        
        # Determine target contractor
        matched_contractor = None
        
        # 1) Match from BILLING_RECORDS
        for b in billing_records:
            if b.get('date') == date_str:
                b_method = str(b.get('ndt_type', '')).strip()
                b_qty = float(b.get('qty', 0))
                if b_method == method and abs(b_qty - qty) < 0.01:
                    matched_contractor = b.get('company')
                    break
                    
        # 2) Match from work_history
        if not matched_contractor and date_str in work_history:
            res_list = work_history[date_str].get('ndt_results', [])
            comps = [str(x.get('업체', '')).strip() for x in res_list if str(x.get('업체', '')).strip()]
            if len(set(comps)) == 1:
                matched_contractor = comps[0]
            elif comps:
                # multiple, try method match
                for x in res_list:
                    if str(x.get('검사방법', '')).strip() == method:
                        matched_contractor = str(x.get('업체', '')).strip()
                        break
                        
        # 3) Fallback date rules
        if not matched_contractor:
            if date_str in ['2026-10-24', '2026-11-04', '2026-11-05', '2026-11-06', '2026-09-29', '2026-10-03', '2026-10-12', '2026-10-16', '2026-10-20']:
                matched_contractor = '다림건설(주)'
            elif date_str in ['2026-08-07', '2026-08-08', '2026-08-22', '2026-09-03', '2026-09-05', '2026-09-08', '2026-09-17', '2026-09-18', '2026-10-06', '2026-10-07', '2026-10-21', '2026-10-23', '2026-11-19']:
                matched_contractor = 'GS네오텍'
            elif date_str == '2026-10-19':
                if method == 'RT' or (method == 'PAUT' and abs(qty - 1.6802) < 0.01):
                    matched_contractor = 'GS네오텍'
                else:
                    matched_contractor = '다림건설(주)'
            else:
                matched_contractor = old_contractor or 'GS네오텍'
                
        new_client = '한국지역난방공사'
        new_contractor = matched_contractor
        
        ws.cell(row=row_idx, column=65, value=new_client)
        ws.cell(row=row_idx, column=66, value=new_contractor)
        
        print(f"Row {row_idx:2d} | {date_str} {method} ({qty:7.4f}) | [발주처] {old_client} -> {new_client} | [시공사] {old_contractor} -> {new_contractor}")
        modified_count += 1
        
    wb.save(excel_path)
    print(f"\n3. 마이그레이션 완료! 총 {modified_count}개 행이 갱신되었습니다.")

if __name__ == '__main__':
    migrate()

