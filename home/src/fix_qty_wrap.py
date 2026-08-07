import os

with open('daily_work_log_exporter.py', 'r', encoding='utf-8') as f:
    code = f.read()

old_code = """        headers_qty = ['방법', '규격', '예상량', '전일 누계', '금일 작업', '총 누계', '공정률(%)', '불량', '불량률(%)', '비고']
        for i, header in enumerate(headers_qty):
            col_letter = get_column_letter(i+1)
            set_cell(f'{col_letter}8', header, font=self.font_bold, fill=self.fill_header)"""

new_code = """        headers_qty = ['방법', '규격', '예상량', '전일 누계', '금일 작업', '총 누계', '공정률(%)', '불량', '불량률(%)', '비고']
        for i, header in enumerate(headers_qty):
            col_letter = get_column_letter(i+1)
            set_cell(f'{col_letter}8', header, font=self.font_bold, fill=self.fill_header, align=self.align_nowrap)"""

code = code.replace(old_code, new_code)

with open('daily_work_log_exporter.py', 'w', encoding='utf-8') as f:
    f.write(code)
print("Updated qty headers alignment successfully")
