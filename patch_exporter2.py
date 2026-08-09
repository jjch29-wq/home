import codecs
import re

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\한국지역난방 중앙지사.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    text = f.read()

target = r'(wb\.save\(save_path\)\r?\n\s+wb\.close\(\)\r?\n\s+log\(f"\\n.*? 저장 완료: \{save_path\}"\)\r?\n\s+messagebox\.showinfo\("완료", f".*?\\n\{save_path\}"\)\r?\n\s+import os\r?\n\s+os\.startfile\(os\.path\.dirname\(save_path\)\)\r?\n\s+)(except Exception as e:)'

replacement = r'''\1
                # --- 4.5 태그 변환 (MonthlyReportExporter) ---
                try:
                    from monthly_report_exporter import MonthlyReportExporter
                    log("▶ 월간 진도보고서 태그 변환을 시작합니다...")
                    exporter = MonthlyReportExporter(history, target_month_str, save_path, doc_num)
                    exporter.export()
                    log("✅ 태그 변환 완료")
                except Exception as ex:
                    log(f"⚠️ 태그 변환 중 오류 (무시됨): {ex}")
                    
            \2'''

if re.search(target, text):
    text = re.sub(target, replacement, text)
    with codecs.open(file_path, 'w', 'utf-8') as f:
        f.write(text)
    print('Patched MonthlyReportExporter!')
else:
    print('Failed to match target!')
