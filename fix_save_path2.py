import codecs
import re

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\한국지역난방 중앙지사.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    text = f.read()

# Replace save block
text = re.sub(
    r'(# 저장\r?\n\s*wb\.save\()filepath(\)\r?\n\s*wb\.close\(\)\r?\n\s*log\(f"\\n✅ 저장 완료: \{)filepath(\}"\)\r?\n\s*messagebox\.showinfo\("완료", f"월간 진도보고서 비파괴검사 현황이 업데이트되었습니다!\\n\{)filepath(\}"\))',
    r'\1save_path\2save_path\3save_path\4\n                import os\n                os.startfile(os.path.dirname(save_path))',
    text
)

with codecs.open(file_path, 'w', 'utf-8') as f:
    f.write(text)
