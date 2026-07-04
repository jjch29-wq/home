import re
with open(r'c:\Users\jjch2\Desktop\PMI\home\src\_archive\Archived-Main-App-20260405-RT-Fix.py', 'r', encoding='utf-8') as f:
    text = f.read()
text = re.sub(r' *self\.log\(f"\[PhotoLog\].*?req_h_pt.*?\n', '\n', text)
with open(r'c:\Users\jjch2\Desktop\PMI\home\src\_archive\Archived-Main-App-20260405-RT-Fix.py', 'w', encoding='utf-8') as f:
    f.write(text)
