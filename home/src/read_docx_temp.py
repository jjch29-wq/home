import docx

doc = docx.Document(r'C:\Users\-\OneDrive\바탕 화면\PAUT_SIS-H-264 (2024ed)KI rev.0(최종)수정.docx')
for i in range(290, 305):
    try:
        print(f"[{i}] {doc.paragraphs[i].text}")
    except Exception as e:
        pass
