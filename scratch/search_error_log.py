with open(r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\error_log.txt", "rb") as f:
    content = f.read()

text = content.decode('utf-16', errors='replace')
lines = text.splitlines()

print("Tracebacks found in error_log.txt:")
for i, line in enumerate(lines):
    if "Traceback" in line or "Error" in line or "Exception" in line or "Fix" in line:
        # Print around this line
        start = max(0, i - 2)
        end = min(len(lines), i + 8)
        print(f"--- Match at line {i} ---")
        for j in range(start, end):
            ascii_line = lines[j].encode('ascii', errors='replace').decode('ascii')
            print(f"{j}: {ascii_line}")
        print()
