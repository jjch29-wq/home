import traceback

try:
    with open(r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\error_log.txt", "r", encoding="utf-16") as f:
        content = f.read()
    print("UTF-16 encoding worked. Last 1000 chars:")
    print(content[-1000:])
except Exception as e:
    print("UTF-16 failed, trying UTF-8...")
    try:
        with open(r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\error_log.txt", "r", encoding="utf-8", errors="ignore") as f:
            content = f.read()
        print("UTF-8 encoding worked. Last 1000 chars:")
        print(content[-1000:])
    except Exception as e2:
        print("Both failed:", e2)
