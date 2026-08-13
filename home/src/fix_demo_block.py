import docx

original_file = r'C:\Users\-\OneDrive\바탕 화면\PAUT_SIS-H-264 (2024ed)KI rev.0(최종)_코멘트반영.docx'
doc = docx.Document(original_file)

old_eng1 = "The demonstration block's weld joint geometry shall be representative of the production joint's\ndetails."
old_eng2 = "The demonstration block's weld joint geometry shall be representative of the production joint's details."

new_eng = "The block shall be representative of the component to be inspected. The weld joint geometry shall be representative of the production joint's details. The weld caps conditions being the same as those encountered for production weld testing."

old_kor1 = "(검증 시험편의 용접접합부 형상은 생산 접합부의 세부사항을 대표해야 한다.)"
new_kor = "(검증 시험편은 검사 대상 부품을 대표해야 한다. 용접 접합부의 기하학적 형상은 생산 접합부의 세부 사항을 대표해야 한다. 용접 덧살(캡) 상태는 생산 용접 검사 시 마주치는 것과 동일해야 한다.)"

for para in doc.paragraphs:
    if "3.  Weld Joint Configuration." in para.text:
        text = para.text
        if old_eng1 in text:
            text = text.replace(old_eng1, new_eng)
        elif old_eng2 in text:
            text = text.replace(old_eng2, new_eng)
            
        if old_kor1 in text:
            text = text.replace(old_kor1, new_kor)
            
        para.text = text
        break

doc.save(r'C:\Users\-\OneDrive\바탕 화면\PAUT_SIS-H-264 (2024ed)KI rev.0(최종)_코멘트최종반영.docx')
print("Fixed and saved.")
