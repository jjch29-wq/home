import docx

file_path = r'C:\Users\-\OneDrive\바탕 화면\PAUT_SIS-H-264 (2024ed)KI rev.0(최종)_코멘트반영.docx'
doc = docx.Document(file_path)

with open('search_result.txt', 'w', encoding='utf-8') as f:
    for i, para in enumerate(doc.paragraphs):
        text = para.text
        if '4.6' in text or 'Demonstration' in text or 'demonstration' in text:
            f.write(f'[{i}] {text}\n')
