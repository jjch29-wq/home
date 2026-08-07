import os

with open('ndt_summary_exporter.py', 'r', encoding='utf-8') as f:
    code = f.read()

# Fix _write_summary_sheet
old_summary_header = """        for col_idx, header_text in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_idx, value=header_text)
            cell.font = self.font_bold
            cell.alignment = self.align_center
            cell.fill = self.fill_header
            cell.border = self.border_thin"""

new_summary_header = """        # Write Title
        ws.merge_cells('A1:F1')
        title_cell = ws.cell(row=1, column=1, value="용 역 명 : 2026년 중앙지사 열수송관 비파괴검사용역")
        title_cell.font = Font(name='맑은 고딕', size=14, bold=True)
        title_cell.alignment = Alignment(horizontal='left', vertical='center')
        ws.row_dimensions[1].height = 25

        for col_idx, header_text in enumerate(headers, start=1):
            cell = ws.cell(row=2, column=col_idx, value=header_text)
            cell.font = self.font_bold
            cell.alignment = self.align_center
            cell.fill = self.fill_header
            cell.border = self.border_thin"""

code = code.replace(old_summary_header, new_summary_header)

old_summary_row = "row_idx = 2\n        for key in sorted_keys:"
new_summary_row = "row_idx = 3\n        for key in sorted_keys:"
code = code.replace(old_summary_row, new_summary_row)


# Fix _write_sheet
old_sheet_header = """        # Write Headers
        for col_idx, header_text in enumerate(all_headers, start=1):
            cell = ws.cell(row=1, column=col_idx, value=header_text)
            cell.font = self.font_bold
            cell.alignment = self.align_center
            cell.fill = self.fill_header
            cell.border = self.border_thin"""

new_sheet_header = """        # Write Title
        num_cols = len(all_headers)
        col_letter = get_column_letter(num_cols)
        ws.merge_cells(f'A1:{col_letter}1')
        title_cell = ws.cell(row=1, column=1, value="용 역 명 : 2026년 중앙지사 열수송관 비파괴검사용역")
        title_cell.font = Font(name='맑은 고딕', size=14, bold=True)
        title_cell.alignment = Alignment(horizontal='left', vertical='center')
        ws.row_dimensions[1].height = 25

        # Write Headers
        for col_idx, header_text in enumerate(all_headers, start=1):
            cell = ws.cell(row=2, column=col_idx, value=header_text)
            cell.font = self.font_bold
            cell.alignment = self.align_center
            cell.fill = self.fill_header
            cell.border = self.border_thin"""

code = code.replace(old_sheet_header, new_sheet_header)

old_sheet_row = "for row_idx, row_data in enumerate(rows, start=2):"
new_sheet_row = "for row_idx, row_data in enumerate(rows, start=3):"
code = code.replace(old_sheet_row, new_sheet_row)


with open('ndt_summary_exporter.py', 'w', encoding='utf-8') as f:
    f.write(code)
print("Updated NDT summary exporter with title successfully")
