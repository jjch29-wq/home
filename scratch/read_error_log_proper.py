import sys

encodings = ['utf-16', 'utf-16-le', 'utf-16-be', 'utf-8', 'cp949', 'euc-kr']
for enc in encodings:
    try:
        with open(r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\error_log.txt", "rb") as f:
            data = f.read()
        text = data.decode(enc)
        print(f"--- Encoding {enc} worked! Last 15 lines: ---")
        lines = text.strip().splitlines()
        for line in lines[-15:]:
            print(line)
        break
    except Exception as e:
        print(f"Encoding {enc} failed: {e}")
