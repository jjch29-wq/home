with open(r'c:\Users\jjch2\Desktop\PMI\home\src\비파괴검사보고서.py', 'r', encoding='utf-8') as f:
    lines = f.read().split('\n')

for i, line in enumerate(lines):
    if 'self.log(f"[ERROR] Failed to merge H1:O4: {e}")' in line:
        insert_code = """
                with open(r'c:\\Users\\jjch2\\Desktop\\PMI\\merge_error.txt', 'a', encoding='utf-8') as f_err:
                    f_err.write(f'Merge error: {e}\\n')
"""
        lines.insert(i+1, insert_code.strip('\n'))
    if 'self.log(f"[ERROR] unmerge fail: {e}")' in line:
        insert_code = """
                    with open(r'c:\\Users\\jjch2\\Desktop\\PMI\\merge_error.txt', 'a', encoding='utf-8') as f_err:
                        f_err.write(f'Unmerge error: {e}\\n')
"""
        lines.insert(i+1, insert_code.strip('\n'))

with open(r'c:\Users\jjch2\Desktop\PMI\home\src\비파괴검사보고서.py', 'w', encoding='utf-8') as f:
    f.write('\n'.join(lines))
print('SUCCESS')
