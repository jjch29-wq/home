import codecs
import re

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\한국지역난방 중앙지사.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    text = f.read()

# 1. Add save path prompt at the beginning of do_export
pattern_start = r'(def do_export\(\):\s*year = year_var\.get\(\)\s*month = month_var\.get\(\)\s*doc_num = doc_num_var\.get\(\)\.strip\(\)\s*filepath = file_path_var\.get\(\)\s*if not filepath:\s*messagebox\.showwarning\("경고", "템플릿 파일을 선택하세요\."\)\s*return)'

replacement_start = r'''\1
                
                save_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    initialfile=f"월간진도보고서_{year}년_{month:02d}월.xlsx",
                    filetypes=[("Excel files", "*.xlsx")],
                    title="월간진도보고서 통합 저장"
                )
                if not save_path:
                    return'''

text = re.sub(pattern_start, replacement_start, text, count=1)

# 2. Add MonthlyReportExporter logic right before openpyxl.load_workbook
pattern_open = r'(# --- 5\. 엑셀 기입 ---\s*import openpyxl\s*wb = openpyxl\.load_workbook\()filepath(\))'

replacement_open = r'''
                # --- 4.5 태그 변환 (MonthlyReportExporter) ---
                try:
                    from monthly_report_exporter import MonthlyReportExporter
                    import json, os
                    history = {}
                    history_path = os.path.join(self.data_dir if hasattr(self, 'data_dir') else 'data', "daily_work_history.json")
                    if os.path.exists(history_path):
                        with open(history_path, 'r', encoding='utf-8') as f:
                            history = json.load(f)
                    
                    target_month_str = f"{year}-{month:02d}"
                    exporter = MonthlyReportExporter(history, target_month_str, filepath, doc_num)
                    exporter.generate(save_path)
                except Exception as e:
                    messagebox.showerror("오류", f"태그 변환 중 오류: {e}")
                    import traceback
                    log(traceback.format_exc())
                    return

                \1save_path\2'''

text = re.sub(pattern_open, replacement_open, text, count=1)

# 3. Add wb.save(save_path) at the very end of the try block
pattern_end = r'(ws\.cell\(row=start_row, column=20, value=round\(tot_total, 4\) if tot_total else \'-\'\)\s*)(except Exception as e:)'

replacement_end = r'''\1
                wb.save(save_path)
                messagebox.showinfo("완료", "월간 진도보고서 (태그 및 1.1표) 작성이 완료되었습니다!")
                os.startfile(os.path.dirname(save_path))
            \2'''

text = re.sub(pattern_end, replacement_end, text, count=1)

with codecs.open(file_path, 'w', 'utf-8') as f:
    f.write(text)
