import codecs

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\한국지역난방 중앙지사.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    lines = f.readlines()

for i in range(len(lines)):
    if 'wb.save(filepath)' in lines[i] and i > 12000 and i < 12500:
        lines[i] = lines[i].replace('wb.save(filepath)', 'wb.save(save_path)')
        if 'wb.close()' in lines[i+1]:
            pass
        if 'log(' in lines[i+2] and 'filepath' in lines[i+2]:
            lines[i+2] = lines[i+2].replace('filepath', 'save_path')
        if 'messagebox.showinfo(' in lines[i+3] and 'filepath' in lines[i+3]:
            lines[i+3] = lines[i+3].replace('filepath', 'save_path')
            lines[i+3] = lines[i+3].replace('월간 진도보고서 비파괴검사 현황이 업데이트되었습니다!', '월간 진도보고서 (태그 및 1.1표) 작성이 완료되었습니다!')
            # insert os.startfile
            lines.insert(i+4, '                import os\n                os.startfile(os.path.dirname(save_path))\n')
        break

for i in range(len(lines)):
    if 'if not main_sites and not mgmt_sites:' in lines[i]:
        # replace the next two lines
        if 'messagebox.showwarning' in lines[i+1]:
            lines[i+1] = "                    if '현장명' in self.daily_usage_df.columns:\n"
            lines[i+2] = "                        main_sites = list(self.daily_usage_df['현장명'].dropna().unique())\n"
            lines.insert(i+3, "                    elif 'Site' in self.daily_usage_df.columns:\n")
            lines.insert(i+4, "                        main_sites = list(self.daily_usage_df['Site'].dropna().unique())\n")
        break

with codecs.open(file_path, 'w', 'utf-8') as f:
    f.writelines(lines)
