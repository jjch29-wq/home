encodings = ['utf-8', 'cp949', 'euc-kr', 'utf-16']
for enc in encodings:
    try:
        with open(r'Na-aba\Material-Master-Manager-V14.py', 'r', encoding=enc) as f:
            content = f.read()
            if '검사비' in content:
                print(f"Found '검사비' with encoding: {enc}")
            else:
                print(f"Did not find '검사비' with encoding: {enc}")
    except Exception as e:
        print(f"Failed with encoding {enc}: {e}")
