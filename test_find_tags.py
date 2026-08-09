import openpyxl
import sys
import codecs

sys.stdout = codecs.getwriter('utf-8')(sys.stdout.detach())

wb = openpyxl.load_workbook(r'C:/Users/jjch2/Desktop/템플릿_최종완성본_V74.xlsx', data_only=True)
ws = wb.worksheets[0]

print("=== 태그 검색 시작 ===")
tags_found = []
for row in ws.iter_rows():
    for cell in row:
        if cell.value and isinstance(cell.value, str):
            v = cell.value.strip()
            if '[[' in v and ']]' in v:
                print(f"태그 발견: {v} (Row {cell.row}, Col {cell.column})")
                tags_found.append({'tag': v, 'row': cell.row, 'col': cell.column})

if not tags_found:
    print("태그를 찾지 못했습니다. 엑셀 파일 저장이 완료되었는지 확인해 주세요.")
