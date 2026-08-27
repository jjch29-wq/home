import os
import re

file_path = r"c:\Users\jjch2\Desktop\PMI\home\src\services\monthly_report_manager.py"

with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

old_qty_mapping = """        qty_mapping = {
            'PAUT_300A이상': qty_start_row,
            'PAUT_300A이상-야간': qty_start_row + 1,
            'PAUT_250A': qty_start_row + 2,
            'PAUT_200A': qty_start_row + 3,
            'PAUT_200A-야간': qty_start_row + 4,
            'PAUT_소계': qty_start_row + 5,
            
            'RT_150A~100A': qty_start_row + 6,
            'RT_150A~100A-야간': qty_start_row + 7,
            'RT_80A이하': qty_start_row + 8,
            'RT_80A이하-야간': qty_start_row + 9,
            'RT_소계': qty_start_row + 10,
            
            'MT_전체(주간)': qty_start_row + 11,
            'MT_전체(야간)': qty_start_row + 12,
            
            'PT_전체(주간)': qty_start_row + 13,
            'PT_전체(야간)': qty_start_row + 14,
        }"""
        
new_qty_mapping = """        qty_mapping = {
            'RT_B필름: 3⅓"x17"': qty_start_row,
            'RT_A필름: 3⅓"x12"': qty_start_row + 1,
            'RT_A/2필름: 3⅓"x6"': qty_start_row + 2,
            'RT_소계': qty_start_row + 3,
            'UT_초음파탐상': qty_start_row + 4,
            'PT_침투탐상': qty_start_row + 5,
            'MT_자분탐상': qty_start_row + 6,
        }"""

if old_qty_mapping in content:
    content = content.replace(old_qty_mapping, new_qty_mapping)
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(content)
    print("Updated monthly_report_manager successfully")
else:
    print("Could not find old_qty_mapping")
