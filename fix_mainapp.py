import codecs
import re

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\한국지역난방 중앙지사.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    text = f.read()

text = re.sub(
    r'(self\.tab_daily_work_log = DailyWorkLogTab\(self\.notebook\)\s+)(self\.notebook\.add)',
    r'\1self.tab_daily_work_log.main_app = self\n        \2',
    text
)

with codecs.open(file_path, 'w', 'utf-8') as f:
    f.write(text)
