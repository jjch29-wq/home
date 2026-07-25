import openpyxl

wb = openpyxl.load_workbook('home/src/templates/양식_기본.xlsx')
ws = [sheet for sheet in wb.worksheets if '조직도' in sheet.title][0]

with open('dump.txt', 'w', encoding='utf-8') as f:
    f.write(f'Sheet: {ws.title}\n')
    for row in ws.iter_rows(min_row=1, max_row=30):
        if any(c.value for c in row):
            vals = [str(c.value) if c.value else '' for c in row]
            f.write(f'{row[0].row}: ' + ' | '.join(vals) + '\n')
