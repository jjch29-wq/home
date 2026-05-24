with open(r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\error_log.txt", "rb") as f:
    data = f.read()

try:
    text = data.decode('utf-16')
    print("--- UTF-16 SUCCESS (printed via repr) ---")
    lines = text.strip().splitlines()
    for line in lines[-30:]:
        # Encode back to sys.stdout encoding or print representation
        print(line.encode('cp949', errors='replace').decode('cp949'))
except Exception as e:
    print("Failed to print:", e)
