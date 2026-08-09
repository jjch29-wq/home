import codecs
import re

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\한국지역난방 중앙지사.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    text = f.read()

target = r'''                if main_agg:
                    sheet_name = wb\.sheetnames\[0\]
                    if sheet_name in sheet_names:
                        write_ndt_sheet_openpyxl\(wb\[sheet_name\], main_agg\)
                        log\(f"✅ '\{sheet_name\}' 시트 기입 완료"\)
                    else:
                        log\(f"⚠️ '\{sheet_name\}' 시트를 찾을 수 없습니다\."\)
                        
                if mgmt_agg:
                    sheet_name = wb\.sheetnames\[0\]
                    if sheet_name in sheet_names:
                        write_ndt_sheet_openpyxl\(wb\[sheet_name\], mgmt_agg\)
                        log\(f"✅ '\{sheet_name\}' 시트 기입 완료"\)
                    else:
                        log\(f"⚠️ '\{sheet_name\}' 시트를 찾을 수 없습니다\."\)'''

replacement = r'''                if main_agg:
                    sheet_name = '3. 비파괴검사 현황 (열배관)'
                    if sheet_name in sheet_names:
                        write_ndt_sheet_openpyxl(wb[sheet_name], main_agg)
                        log(f"✅ '{sheet_name}' 시트 기입 완료")
                    else:
                        log(f"⚠️ '{sheet_name}' 시트를 찾을 수 없습니다.")
                        
                if mgmt_agg:
                    sheet_name = '3. 비파괴검사 현황 (관리소)'
                    if sheet_name in sheet_names:
                        write_ndt_sheet_openpyxl(wb[sheet_name], mgmt_agg)
                        log(f"✅ '{sheet_name}' 시트 기입 완료")
                    else:
                        log(f"⚠️ '{sheet_name}' 시트를 찾을 수 없습니다.")'''

if re.search(target, text):
    text = re.sub(target, replacement, text)
    with codecs.open(file_path, 'w', 'utf-8') as f:
        f.write(text)
    print('Reverted main_agg and mgmt_agg sheet names successfully!')
else:
    print('Failed to match target!')
