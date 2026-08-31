import json
import re
from collections import defaultdict
import datetime

# Mock parsing logic to test it locally
history_data = {
    "2026-08-31": {
        "ndt_results": [
            {"공사구분": "주배관", "용접사": "김동현", "구간정보": "1/P, 4/LF", "결과": "불합격", "검사방법": "RT"},
            {"공사구분": "주배관", "용접사": "김동현", "구간정보": "N/1, O, O", "결과": "합격", "검사방법": "RT"},
            {"공사구분": "관리소", "용접사": "박지성", "구간정보": "2/IP", "결과": "불합격", "검사방법": "RT"}
        ]
    }
}

DEFECT_MAP = {
    'CRACK': 0, 'CR': 0,
    'IP': 1,
    'LF': 2,
    'S': 3, 'SLAG': 3,
    'P': 4, 'POROSITY': 4,
    'UC': 5,
    'RUC': 6,
    'BT': 7,
    'TI': 8,
    'RC': 9,
    'EP': 10,
    'SD': 11,
    'C/P': 12, 'CP': 12,
    'OVER GR': 13, 'OGR': 13,
    'GR': 14,
    'CL': 15,
    'C': 16
}

def parse_defects(sec_info):
    defects = []
    if not sec_info:
        return defects
    parts = str(sec_info).split(',')
    for part in parts:
        part = part.strip()
        if not part: continue
        if '/' in part:
            code = part.split('/')[-1].strip().upper()
            if code in DEFECT_MAP:
                defects.append(code)
    return defects

monthly_defects = defaultdict(lambda: defaultdict(lambda: [0]*17)) # year -> month -> counts
welder_stats = defaultdict(lambda: defaultdict(lambda: {'total': 0, 'fail': 0, 'defects': [0]*17})) # gongsa -> welder -> stats

for date_str, daily_data in history_data.items():
    if not re.match(r'^\d{4}-\d{2}-\d{2}$', str(date_str)): continue
    year = date_str[:4]
    month = int(date_str[5:7])
    
    for row in daily_data.get('ndt_results', []):
        method = str(row.get('검사방법', '')).strip().upper()
        if not method: continue
        
        gongsa = row.get('공사구분', '주배관')
        welder = str(row.get('용접사', '')).strip()
        if not welder: continue
        
        result = str(row.get('결과', '')).strip()
        is_fail = (result == '불합격' or result == '재촬영')
        
        # Welder stats
        welder_stats[gongsa][welder]['total'] += 1
        if is_fail:
            welder_stats[gongsa][welder]['fail'] += 1
            
        # Parse defects
        sec_info = row.get('구간정보', '')
        defects = parse_defects(sec_info)
        
        for d in defects:
            idx = DEFECT_MAP[d]
            monthly_defects[year][month][idx] += 1
            welder_stats[gongsa][welder]['defects'][idx] += 1

print(monthly_defects)
print(welder_stats)
