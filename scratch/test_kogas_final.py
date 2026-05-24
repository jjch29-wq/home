import re

# Mock the defect mapping
defect_map = {
    "D1": "C", "D2": "IP", "D3": "LF", "D4": "S", "D5": "P",
    "D6": "UC", "D7": "RUC", "D8": "BT", "D9": "TI", "D10": "CP",
    "D11": "RC", "D12": "Mis", "D13": "EP", "D14": "SD", "D15": "Oth"
}

def get_item_val(it):
    grade = str(it.get('Deg', '')).strip()
    if grade.lower() == 'nan':
        grade = ''
    checked = []
    for d_key, abbrev in defect_map.items():
        val = str(it.get(d_key, '')).strip()
        if val in ["√", "1", "v", "V", "o", "O", "●", "?"]:
            checked.append(abbrev)
    if checked:
        return f"{grade if grade else '4'}/{','.join(checked)}"
    else:
        if any(x in str(it.get('Result', '')).upper() for x in ["ACC", "OK"]):
            return "1"
        return grade if grade else "4"

# Test Cases
test_cases = [
    # 1. Accept only, no defects
    {"Deg": "", "Result": "ACC", "expected": "1"},
    # 2. Accept only, Grade is 1, no defects
    {"Deg": "1", "Result": "ACC", "expected": "1"},
    # 3. Grade 1, Porosity checked as '1'
    {"Deg": "1", "Result": "ACC", "D5": "1", "expected": "1/P"},
    # 4. Grade empty, Porosity checked as '√' -> defaults to 4/P
    {"Deg": "", "Result": "REJ", "D5": "√", "expected": "4/P"},
    # 5. Reject, no grade, no defects -> defaults to 4
    {"Deg": "", "Result": "REJ", "expected": "4"},
    # 6. Reject, grade 4, no defects -> 4
    {"Deg": "4", "Result": "REJ", "expected": "4"},
    # 7. Multiple defects (Porosity and Slag), Grade 2
    {"Deg": "2", "Result": "ACC", "D4": "V", "D5": "1", "expected": "2/S,P"},
]

print("=== Running get_item_val Verification ===")
all_passed = True
for idx, tc in enumerate(test_cases, 1):
    item = {k: v for k, v in tc.items() if k != "expected"}
    res = get_item_val(item)
    expected = tc["expected"]
    if res == expected:
        print(f"Test {idx} PASSED: {item} -> {res}")
    else:
        print(f"Test {idx} FAILED: {item} -> Got: {res}, Expected: {expected}")
        all_passed = False

if all_passed:
    print("\nSUCCESS: All get_item_val formatting test cases passed!")
else:
    print("\nFAILURE: Some test cases failed.")
