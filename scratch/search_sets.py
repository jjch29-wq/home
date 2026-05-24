with open(r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\src\Archived-Main-App-20260405-RT-Fix.py", "r", encoding="utf-8") as f:
    lines = f.readlines()

in_load_settings = False
print("Lines in load_settings calling .set():")
for i, line in enumerate(lines):
    line_num = i + 1
    if "def load_settings" in line:
        in_load_settings = True
    if in_load_settings and "def " in line and "def load_settings" not in line:
        in_load_settings = False
    if in_load_settings and ".set(" in line:
        print(f"{line_num}: {line.strip()}")
