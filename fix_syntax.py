with open(r'c:\Users\jjch2\Desktop\PMI\home\src\비파괴검사보고서_recovered.py', 'r', encoding='utf-8') as f:
    lines = f.read().split('\n')

for i in range(len(lines)):
    if 'self.status_log.insert(tk.END, f"[{datetime.datetime.now().strftime(' in lines[i]:
        if i+1 < len(lines) and lines[i+1].strip() == '")':
            lines[i] = lines[i] + '\\n")'
            lines[i+1] = ''

with open(r'c:\Users\jjch2\Desktop\PMI\home\src\비파괴검사보고서_recovered.py', 'w', encoding='utf-8') as f:
    f.write('\n'.join(lines))
print('Fixed line 386')
