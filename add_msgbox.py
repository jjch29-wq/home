import codecs

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\한국지역난방 중앙지사.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if 'log(f"▶ \'{paut_sheet_name}\' 시트 기입 완료 (총 {len(groups)}건)")' in line:
        indent = line[:line.find('log')]
        lines.insert(i+1, indent + "messagebox.showinfo('PAUT 디버그', f'PAUT 데이터 {len(groups)}건 기입 완료!\\npaut_records: {len(paut_records)}건')\n")
        break
    if 'log(f"⚠️ \'{paut_sheet_name}\' 작성 중 오류: {e}")' in line:
        indent = line[:line.find('log')]
        lines.insert(i+1, indent + "messagebox.showerror('PAUT 에러', f'에러 발생: {e}')\n")

with codecs.open(file_path, 'w', 'utf-8') as f:
    f.writelines(lines)
