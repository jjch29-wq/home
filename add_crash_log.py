import codecs

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\한국지역난방 중앙지사.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    text = f.read()

import re
old_catch = '''            except Exception as e:
                log(f"❌ 오류 발생: {e}")'''

new_catch = '''            except Exception as e:
                import traceback
                with open(r'c:\\Users\\jjch2\\Desktop\\PMI\\crash_log.txt', 'w', encoding='utf-8') as crash_f:
                    crash_f.write(traceback.format_exc())
                log(f"❌ 오류 발생: {e}")'''

text = text.replace(old_catch, new_catch)

with codecs.open(file_path, 'w', 'utf-8') as f:
    f.write(text)
