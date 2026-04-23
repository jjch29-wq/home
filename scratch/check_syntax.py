import sys
import traceback

filename = r'home/src/Material-Master-Manager-V13.py'
try:
    with open(filename, 'r', encoding='utf-8') as f:
        source = f.read()
    compile(source, filename, 'exec')
    print("No syntax errors found.")
except SyntaxError as e:
    print(f"Syntax Error: {e}")
    print(f"Line: {e.lineno}")
    print(f"Offset: {e.offset}")
    print(f"Text: {e.text}")
except Exception as e:
    print(f"An error occurred: {e}")
    traceback.print_exc()
