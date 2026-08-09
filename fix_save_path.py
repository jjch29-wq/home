import codecs

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\한국지역난방 중앙지사.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    text = f.read()

target1 = """                if not main_sites and not mgmt_sites:
                    messagebox.showwarning("입력 오류", "열배관 또는 관리소에 해당하는 현장명을 최소 1개 입력해주세요.")
                    return"""

replacement1 = """                if not main_sites and not mgmt_sites:
                    # 빈칸일 경우 등록된 모든 현장을 '열배관'으로 일괄 처리
                    if '현장명' in self.daily_usage_df.columns:
                        main_sites = list(self.daily_usage_df['현장명'].dropna().unique())
                    elif 'Site' in self.daily_usage_df.columns:
                        main_sites = list(self.daily_usage_df['Site'].dropna().unique())"""

text = text.replace(target1, replacement1)


target2 = """                # 저장
                wb.save(filepath)
                wb.close()
                log(f"\\n✅ 저장 완료: {filepath}")
                messagebox.showinfo("완료", f"월간 진도보고서 비파괴검사 현황이 업데이트되었습니다!\\n{filepath}")"""

replacement2 = """                # 저장
                wb.save(save_path)
                wb.close()
                log(f"\\n✅ 저장 완료: {save_path}")
                messagebox.showinfo("완료", f"월간 진도보고서 (태그 및 1.1표 종합) 작성이 완료되었습니다!\\n{save_path}")
                os.startfile(os.path.dirname(save_path))"""

text = text.replace(target2, replacement2)

with codecs.open(file_path, 'w', 'utf-8') as f:
    f.write(text)
