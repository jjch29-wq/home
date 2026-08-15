import re

with open(r'c:\Users\jjch2\Desktop\PMI\home\src\비파괴검사보고서.py', 'r', encoding='utf-8') as f:
    text = f.read()

def repl(m):
    return m.group(1) + '''
                                # [FIX] Inject fitToPage into sheetPr if fitToWidth or fitToHeight is used
                                if 'fitToWidth=\"1\"' in str(rep_ps) or 'fitToHeight=\"1\"' in str(rep_ps):
                                    if '<pageSetUpPr ' not in tmpl_sheet:
                                        if '</sheetPr>' in tmpl_sheet:
                                            tmpl_sheet = tmpl_sheet.replace('</sheetPr>', '<pageSetUpPr fitToPage=\"1\"/></sheetPr>', 1)
                                        elif re.search(r'<sheetPr\\\\b([^>]*)/>', tmpl_sheet):
                                            tmpl_sheet = re.sub(r'<sheetPr\\\\b([^>]*)/>', r'<sheetPr\\\\1><pageSetUpPr fitToPage=\"1\"/></sheetPr>', tmpl_sheet, count=1)
''' + m.group(2)

pattern = r'(if rep_ps:\s*if re\.search\(r\'<pageSetup\\b\[\^>\]\*/\>\', tmpl_sheet\):\s*tmpl_sheet = re\.sub\(r\'<pageSetup\\b\[\^>\]\*/\>\', rep_ps, tmpl_sheet, count=1\)\s*else:\s*#[^\n]*\n\s*tmpl_sheet = tmpl_sheet\.replace\(\'</sheetData>\', \'</sheetData>\' \+ rep_ps, 1\))(\s*elif rep_scale:)'

new_text = re.sub(pattern, repl, text)

if new_text != text:
    with open(r'c:\Users\jjch2\Desktop\PMI\home\src\비파괴검사보고서.py', 'w', encoding='utf-8') as f:
        f.write(new_text)
    print('SUCCESS')
else:
    print('NO MATCH')
