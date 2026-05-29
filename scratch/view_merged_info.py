import os

for encoding in ['utf-8', 'utf-16', 'utf-16le', 'cp949', 'euc-kr']:
    try:
        with open('merged_info.txt', 'r', encoding=encoding) as f:
            content = f.read()
            print(f"--- Encoding: {encoding} ---")
            print(content[:1000])
            break
    except Exception as e:
        print(f"Failed with {encoding}: {e}")
