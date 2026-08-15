with open(r'c:\Users\jjch2\Desktop\PMI\home\src\비파괴검사보고서.py', 'r', encoding='utf-8') as f:
    lines = f.read().split('\n')

if lines[8780] == '            self._write_gapji_metadata(ws0)':
    lines[8780] = '            self._write_gapji_metadata(ws0, mode="PT")'

if lines[8931] == '            self._write_gapji_metadata(ws0)':
    lines[8931] = '            self._write_gapji_metadata(ws0, mode=mode)'

with open(r'c:\Users\jjch2\Desktop\PMI\home\src\비파괴검사보고서.py', 'w', encoding='utf-8') as f:
    f.write('\n'.join(lines))
print('SUCCESS')
