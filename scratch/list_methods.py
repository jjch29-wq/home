with open(r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\src\Archived-Main-App-20260405-RT-Fix.py", "r", encoding="utf-8") as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if line.strip().startswith("def "):
        print(f"{i+1}: {line.strip()}")
