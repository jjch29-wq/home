import urllib.request
import json
import time

def sync_to_google_sheet():
    get_url = 'https://script.google.com/macros/s/AKfycbzJWmpM4bAhFebNsY-seBDMo93WBsEIg1-AqVzQYfir3mquQd5aeSC1Omv_ZScRpW9dQQ/exec?api=recent&limit=30'
    post_url = 'https://script.google.com/macros/s/AKfycbzJWmpM4bAhFebNsY-seBDMo93WBsEIg1-AqVzQYfir3mquQd5aeSC1Omv_ZScRpW9dQQ/exec'
    
    req = urllib.request.Request(get_url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req) as resp:
        res = json.loads(resp.read().decode('utf-8'))
        
    records = res.get('records', [])
    print(f"구글 시트에서 {len(records)}건의 레코드를 가져왔습니다.")
    
    contractor_map = {
        '2026-11-06': '다림건설(주)',
        '2026-11-05': '다림건설(주)',
        '2026-11-04': '다림건설(주)',
        '2026-10-24': '다림건설(주)',
        '2026-10-23': 'GS네오텍',
        '2026-10-21': 'GS네오텍',
        '2026-10-20': '다림건설(주)',
    }
    
    updated_count = 0
    for r in records:
        d = r.get('date')
        comp = r.get('company')
        method = r.get('method')
        qty = float(r.get('qty') or 0)
        
        target_comp = None
        if d in contractor_map:
            target_comp = contractor_map[d]
        elif d == '2026-10-19':
            if method == 'RT' or (method == 'PAUT' and abs(qty - 1.6802) < 0.01):
                target_comp = 'GS네오텍'
            else:
                target_comp = '다림건설(주)'
                
        if not target_comp or comp == target_comp:
            continue
            
        print(f"업데이트 대상: {d} {method} ({qty}) | {comp} -> {target_comp} (entry_time: {r.get('entry_time')})")
        
        # Build update payload matching updateRecordCore
        # Ensure workers is a list of dicts
        raw_workers = r.get('workers') or []
        worker_list = []
        if isinstance(raw_workers, str):
            for w_part in raw_workers.split(','):
                w_part = w_part.strip()
                if w_part:
                    name = w_part.split('(')[0].strip()
                    worker_list.append({'user': name, 'shift': '주간', 'time': '09:00~18:00', 'ot': '0'})
        elif isinstance(raw_workers, list):
            worker_list = raw_workers
            
        payload = {
            'action': 'updateRecord',
            'entry_time': r.get('entry_time'),
            'date': d,
            'company': target_comp,
            'company_code': '한국지역난방공사',
            'loc_type': r.get('loc_type') or '열수송관',
            'method': method,
            'pipe': r.get('pipe') or '',
            'item_name': r.get('item_name') or 'Pipe',
            'material': r.get('material') or '',
            'applied_code': r.get('applied_code') or '지역난방시방서',
            'report_no': r.get('report_no') or '',
            'qty': qty,
            'unit': r.get('unit') or ('매' if method == 'RT' else 'M'),
            'unit_price': float(r.get('unit_price') or 0),
            'travel_cost': float(r.get('travel_cost') or 0),
            'total_fee': float(r.get('total_fee') or 0),
            'workers': worker_list
        }
        
        post_data = json.dumps(payload).encode('utf-8')
        post_req = urllib.request.Request(
            post_url,
            data=post_data,
            headers={'Content-Type': 'application/json', 'User-Agent': 'Mozilla/5.0'}
        )
        try:
            with urllib.request.urlopen(post_req, timeout=15) as p_resp:
                res_json = json.loads(p_resp.read().decode('utf-8'))
                print(f"  -> 결과: {res_json.get('status')} ({res_json.get('message') or res_json.get('detail')})")
                if res_json.get('status') == 'success':
                    updated_count += 1
        except Exception as e:
            print(f"  -> 오류: {e}")
            
        time.sleep(0.5)
        
    print(f"\n총 {updated_count}건의 구글 시트 행이 시공사명으로 업데이트되었습니다.")

if __name__ == '__main__':
    sync_to_google_sheet()

