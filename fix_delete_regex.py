import os
import re

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\자재작업일보기성서류Ver.1.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Fix str.contains in delete_recent_entry
old_contains = """(self.transactions_df['Note'].str.contains(f"{site} 현장 사용", na=False))"""
new_contains = """(self.transactions_df['Note'].str.contains(f"{site} 현장 사용", na=False, regex=False))"""

if old_contains in content:
    content = content.replace(old_contains, new_contains)
    print("Fixed str.contains regex issue.")
else:
    print("Could not find the str.contains pattern.")

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
