import re

with open(r'c:\Users\jjch2\Desktop\PMI\home\src\_archive\Archived-Main-App-20260405-RT-Fix.py', 'r', encoding='utf-8') as f:
    text = f.read()

# 1. Update combobox values
text = text.replace('values=["1", "2", "3"]', 'values=["1", "2", "3", "4"]')

# 2. Change GRID_COLS from 6 to 12
text = text.replace('GRID_COLS = 6', 'GRID_COLS = 12')

# 3. Update CELL_WIDTH_PX scaling
text = re.sub(
    r'if num_cols == 1: CELL_WIDTH_PX = unit_per_px \* 6\n\s*elif num_cols == 2: CELL_WIDTH_PX = unit_per_px \* 3\n\s*else: CELL_WIDTH_PX = unit_per_px \* 2',
    'if num_cols == 1: CELL_WIDTH_PX = unit_per_px * 12\n            elif num_cols == 2: CELL_WIDTH_PX = unit_per_px * 6\n            elif num_cols == 3: CELL_WIDTH_PX = unit_per_px * 4\n            else: CELL_WIDTH_PX = unit_per_px * 3',
    text
)

# 4. Update photo_col_spans
text = text.replace(
'''            if num_cols == 1:
                photo_col_spans = [(0, GRID_COLS - 1)]
                CELL_WIDTH_PX = unit_per_px * 6
            elif num_cols == 2:
                photo_col_spans = [(0, 2), (3, 5)]
                CELL_WIDTH_PX = unit_per_px * 3
            else: # 3 Columns
                photo_col_spans = [(0, 1), (2, 3), (4, 5)]
                CELL_WIDTH_PX = unit_per_px * 2''',
'''            if num_cols == 1:
                photo_col_spans = [(0, GRID_COLS - 1)]
                CELL_WIDTH_PX = unit_per_px * 12
            elif num_cols == 2:
                photo_col_spans = [(0, 5), (6, 11)]
                CELL_WIDTH_PX = unit_per_px * 6
            elif num_cols == 3:
                photo_col_spans = [(0, 3), (4, 7), (8, 11)]
                CELL_WIDTH_PX = unit_per_px * 4
            else: # 4 Columns
                photo_col_spans = [(0, 2), (3, 5), (6, 8), (9, 11)]
                CELL_WIDTH_PX = unit_per_px * 3'''
)

# 5. Update header merge_ranges
text = text.replace(
'''            worksheet.merge_range(1, 0, 3, 2, "", company_format)''',
'''            worksheet.merge_range(1, 0, 3, 5, "", company_format)'''
)

text = text.replace('worksheet.merge_range(1, 3, 1, 5,', 'worksheet.merge_range(1, 6, 1, 11,')
text = text.replace('worksheet.merge_range(2, 3, 2, 5,', 'worksheet.merge_range(2, 6, 2, 11,')
text = text.replace('worksheet.merge_range(3, 3, 3, 5,', 'worksheet.merge_range(3, 6, 3, 11,')

# 6. Update photos_per_page
text = text.replace('photos_per_page = 4 if num_cols == 1 else (8 if num_cols == 2 else 12)', 'photos_per_page = 4 if num_cols == 1 else (8 if num_cols == 2 else (12 if num_cols == 3 else 16))')

with open(r'c:\Users\jjch2\Desktop\PMI\home\src\_archive\Archived-Main-App-20260405-RT-Fix.py', 'w', encoding='utf-8') as f:
    f.write(text)
print('Done!')
