import openpyxl
import sys
import codecs

sys.stdout = codecs.getwriter('utf-8')(sys.stdout.detach())

# Check the generated file as well
files_to_check = [
    r'C:/Users/jjch2/Desktop/템플릿_최종완성본_V74.xlsx',
    r'C:/Users/jjch2/Desktop/월간진도보고서_2026년_08월.xlsx'
]

for file_path in files_to_check:
    print(f"\n=== 파일 확인: {file_path} ===")
    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        ws = wb.worksheets[0]
        
        tags_found = []
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    v = cell.value.strip()
                    if '[[' in v and ']]' in v:
                        print(f"  👉 태그 발견: {v} (Row {cell.row}, Col {cell.column})")
                        tags_found.append(v)
        if not tags_found:
            print("  ⚠️ 이 파일에서 태그를 찾지 못했습니다.")
    except Exception as e:
        print(f"  ❌ 파일 열기 실패: {e}")
