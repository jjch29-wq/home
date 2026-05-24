with open(r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\error_log.txt", "rb") as f:
    content = f.read()

# Let's decode as utf-16 with replace errors, and print using encode('utf-8') or sys.stdout.buffer
import sys
try:
    text = content.decode('utf-16', errors='replace')
    # Print last 20 lines
    lines = text.splitlines()
    print("--- Last 40 lines of error log ---")
    for line in lines[-40:]:
        print(line.encode('ascii', errors='replace').decode('ascii'))
except Exception as e:
    print("Error:", e)
