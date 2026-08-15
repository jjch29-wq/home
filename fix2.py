with open(r'c:\Users\jjch2\Desktop\PMI\home\src\비파괴검사보고서.py', 'r', encoding='utf-8') as f:
    text = f.read()

new_text = text.replace(r'<sheetPr\\b([^>]*)/>', r'<sheetPr\b([^>]*)/>').replace(r'<sheetPr\\1>', r'<sheetPr\1>')

if new_text != text:
    with open(r'c:\Users\jjch2\Desktop\PMI\home\src\비파괴검사보고서.py', 'w', encoding='utf-8') as f:
        f.write(new_text)
    print('SUCCESS 2')
else:
    print('NO MATCH 2')
