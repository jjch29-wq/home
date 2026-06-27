
import re

def to_float(val):
    try: return float(str(val).replace(',', '').strip())
    except: return 0.0

def get_value(item, key):
    val = item.get(key, "")
    if key == "selected":
        return (0, 0 if val is True else 1)
    if val is None: return (2, "")
    if key in ["Ni", "Cr", "Mo", "Mn"]: return (0, to_float(val))
    s_val = str(val).strip()
    if not s_val: return (2, "")
    if re.match(r'^\d+(\.\d+)?$', s_val):
        try: return (0, float(s_val))
        except: pass
    natural_key = re.sub(r'(\d+)', lambda m: m.group(1).zfill(20), s_val.lower())
    return (1, natural_key)

def test_final_sort():
    data = [
        {'Dwg': 'ISO-B', 'No': '1', 'Joint': 'J1', 'selected': True},
        {'Dwg': 'ISO-A', 'No': '2', 'Joint': 'J1', 'selected': False}
    ]
    
    # Simulate clicking "No"
    print("--- Sorting by 'No' (Primary) ---")
    data_no = sorted(data, key=lambda x: (get_value(x, "No"), get_value(x, "Dwg")))
    for d in data_no:
        print(f"  No: {d['No']}, ISO: {d['Dwg']}")

    # Simulate clicking "Dwg"
    print("\n--- Sorting by 'Dwg' (Primary) ---")
    data_dwg = sorted(data, key=lambda x: (get_value(x, "Dwg"), get_value(x, "No")))
    for d in data_dwg:
        print(f"  ISO: {d['Dwg']}, No: {d['No']}")

    # Simulate clicking "V" (Selection)
    print("\n--- Sorting by 'V' (Selection) ---")
    data_v = sorted(data, key=lambda x: (get_value(x, "selected"), get_value(x, "Dwg")))
    for d in data_v:
        print(f"  Selected: {d['selected']}, ISO: {d['Dwg']}")

if __name__ == "__main__":
    test_final_sort()
