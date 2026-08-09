import codecs
import re

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\한국지역난방 중앙지사.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    text = f.read()

# 1. Update the key in groups loop to include '용접사'
target1 = r"key = \(str\(r\.get\('업체', ''\)\), str\(r\.get\('구간', ''\)\), str\(r\.get\('라인번호', ''\)\), str\(r\.get\('관경', ''\)\), str\(r\.get\('Joint No\.', ''\)\)\)"
replacement1 = r"key = (str(r.get('업체', '')), str(r.get('구간', '')), str(r.get('라인번호', '')), str(r.get('관경', '')), str(r.get('Joint No.', '')), str(r.get('용접사', '')))"
if re.search(target1, text): print('Found target1')
text = re.sub(target1, replacement1, text)

# 2. Update the columns
target2 = r'''                                업체, 구간, 라인번호, 관경, joint = key
                                ws_paut\.cell\(row=current_row, column=2\)\.value = 업체
                                ws_paut\.cell\(row=current_row, column=3\)\.value = i \+ 1
                                ws_paut\.cell\(row=current_row, column=4\)\.value = 구간
                                ws_paut\.cell\(row=current_row, column=5\)\.value = 라인번호
                                ws_paut\.cell\(row=current_row, column=6\)\.value = 관경
                                ws_paut\.cell\(row=current_row, column=7\)\.value = joint
                                ws_paut\.cell\(row=current_row, column=8\)\.value = data\['shift'\]
                                ws_paut\.cell\(row=current_row, column=9\)\.value = 'M'
                                
                                ori = data\['ORI'\]
                                re_val = data\['RE'\]
                                tot = ori \+ re_val
                                
                                if ori > 0: ws_paut\.cell\(row=current_row, column=10\)\.value = round\(ori, 4\)
                                if re_val > 0: ws_paut\.cell\(row=current_row, column=11\)\.value = round\(re_val, 4\)
                                if tot > 0: ws_paut\.cell\(row=current_row, column=12\)\.value = round\(tot, 4\)'''

replacement2 = r'''                                업체, 구간, 라인번호, 관경, joint, 용접사 = key
                                ws_paut.cell(row=current_row, column=2).value = 업체
                                ws_paut.cell(row=current_row, column=3).value = i + 1
                                ws_paut.cell(row=current_row, column=4).value = 구간
                                ws_paut.cell(row=current_row, column=6).value = 라인번호
                                ws_paut.cell(row=current_row, column=10).value = 관경
                                ws_paut.cell(row=current_row, column=12).value = joint
                                ws_paut.cell(row=current_row, column=14).value = 용접사
                                ws_paut.cell(row=current_row, column=16).value = '합격'
                                
                                ori = data['ORI']
                                re_val = data['RE']
                                tot = ori + re_val
                                
                                if ori > 0: ws_paut.cell(row=current_row, column=17).value = round(ori, 4)
                                if re_val > 0: ws_paut.cell(row=current_row, column=18).value = round(re_val, 4)
                                if tot > 0: ws_paut.cell(row=current_row, column=20).value = round(tot, 4)'''

if re.search(target2, text): print('Found target2')
text = re.sub(target2, replacement2, text)

with codecs.open(file_path, 'w', 'utf-8') as f:
    f.write(text)
print('Patched columns and merged cell fixes!')
