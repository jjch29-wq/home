import re

def find_usages():
    filepath = "c:/Users/jjch2/Desktop/보고서/Project PROVIDENCE/Request/PMI/Na-aba/home/src/Archived-Main-App-20260405-RT-Fix.py"
    with open(filepath, "r", encoding="utf-8") as f:
        lines = f.readlines()
        
    print("=== Usages of Dwg_Sub ===")
    for i, line in enumerate(lines):
        if "Dwg_Sub" in line or "dwg_sub" in line:
            print(f"Line {i+1}: {line.strip()}")
            
    print("\n=== Usages of 'Dwg' and 'Dwg_Sub' in writing ===")
    for i, line in enumerate(lines):
        if any(k in line for k in ["['Dwg']", "['Dwg_Sub']", "get('Dwg')", "get('Dwg_Sub')"]):
            if i > 6800: # Only show write-out parts
                print(f"Line {i+1}: {line.strip()}")

if __name__ == "__main__":
    find_usages()
