import os
import re

path = r'Material-Master-Manager-V13.py'
with open(path, 'r', encoding='utf-8', errors='ignore') as f:
    content = f.read()

# Find the inspector block and replace it with full OT collection logic
# The current block looks like:
#   inspectors = []
#   if not site_records.empty:
#       for _, row in site_records.iterrows():
#           ...
#   data['inspector'] = ", ".join(inspectors)

# We find the anchor: data['inspector'] = 
anchor_start = "            inspectors = []\n            if not site_records.empty:\n"
anchor_end = 'data[\'inspector\'] = ", ".join(inspectors)'

if anchor_start not in content:
    print("FAILED: anchor_start not found")
    exit(1)
if anchor_end not in content:
    print("FAILED: anchor_end not found")
    exit(1)

# Split on anchor_start, then on anchor_end
parts = content.split(anchor_start, 1)
prefix = parts[0]
rest = parts[1].split(anchor_end, 1)
rest_after = rest[1]

# New OT block that replaces the old inspectors block
new_ot_block = '''            # [NEW] 작업자 / OT 정보 DB에서 집계 (현장 탭 기록 기준)
            def _clean_name(n):
                if not n or str(n).lower() in ('nan', 'none', ''): return ''
                text = str(n).strip()
                titles = ['부장', '차장', '과장', '대리', '주임', '기사', '선임', '수석', '책임',
                          '팀장', '이사', '본부장', '실장', '소장', '직장', '반장', '팀원', '계장']
                for t in titles:
                    import re as _re
                    text = _re.sub(r'[\\s/(\\[]*' + t + r'[\\s)\\]]*$', '', text)
                    text = _re.sub(r'^[\\s/(\\[]*' + t + r'[\\s)\\]]*', '', text)
                return text.strip()

            inspectors = []
            ot_groups = {}  # key: (work_time, ot_amount) -> {names:[], company:''}

            if not site_records.empty:
                company_val = data.get('company', '')
                for _, row in site_records.iterrows():
                    for i in range(1, 11):
                        u_key   = 'User'     if i == 1 else f'User{i}'
                        wt_key  = 'WorkTime' if i == 1 else f'WorkTime{i}'
                        ot_key  = 'OT'       if i == 1 else f'OT{i}'

                        name = str(row.get(u_key, '')).strip()
                        if not name or name == 'nan': continue

                        if name not in inspectors:
                            inspectors.append(name)

                        wt  = str(row.get(wt_key, '')).strip()
                        ot_raw = str(row.get(ot_key,  '')).strip()
                        oa  = ''.join(c for c in ot_raw if c.isdigit())

                        if wt == 'nan': wt = ''
                        if not wt and not oa: continue

                        key = (wt, oa)
                        if key not in ot_groups:
                            ot_groups[key] = {'names': [], 'company': company_val}
                        if name not in ot_groups[key]['names']:
                            ot_groups[key]['names'].append(name)

            # Inspector display (titles stripped when 3+)
            if len(inspectors) >= 3:
                disp_insp = [_clean_name(n) for n in inspectors]
            else:
                disp_insp = inspectors
            data['inspector'] = ', '.join(disp_insp)

            # Build ot_status list
            data['ot_status'] = []
            for (wt, oa), grp in ot_groups.items():
                names = grp['names']
                if len(names) >= 3:
                    name_disp = ', '.join([_clean_name(n) for n in names])
                else:
                    name_disp = ', '.join(names)
                wt_disp = f'{wt} (동일)' if len(names) > 1 else wt
                data['ot_status'].append({
                    'names':      name_disp,
                    'ot_hours':   wt_disp,
                    'ot_amount':  f'{int(oa):,}' if oa else '',
                    'company':    grp['company'],
                    'method':     method,
                })

            '''

# Rebuild content
content = prefix + new_ot_block + rest_after

with open(path, 'w', encoding='utf-8', errors='ignore') as f:
    f.write(content)
print("SUCCESS: OT data collection from DB inserted")
