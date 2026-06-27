
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

def test_hierarchy_sort():
    data = [
        {'Dwg': 'ISO-B', 'Joint': '01', 'No': '1', 'order_index': 0},
        {'Dwg': 'ISO-A', 'Joint': '10', 'No': '2', 'order_index': 1},
        {'Dwg': 'ISO-A', 'Joint': '02', 'No': '3', 'order_index': 2},
    ]
    
    # 1. Simulate clicking "Joint No" or "ISO/DWG"
    print("--- Sorting by 'Joint No' (Should prioritize ISO first) ---")
    # data_key is "Joint" -> logic: (ISO, Joint)
    sort_key = lambda x: (get_value(x, "Dwg"), get_value(x, "Joint"), x.get('order_index', 0))
    data_hier = sorted(data, key=sort_key)
    for d in data_hier:
        print(f"  ISO: {d['Dwg']}, Joint: {d['Joint']}, No: {d['No']}")

    # 2. Verify that ISO-A J02 comes before ISO-A J10
    if data_hier[0]['Dwg'] == 'ISO-A' and data_hier[0]['Joint'] == '02':
        print("\nSUCCESS: Natural sort and ISO grouping confirmed.")
    
    # 3. Simulate clicking "No" (Global sort)
    print("\n--- Sorting by 'No' (Global) ---")
    sort_key_no = lambda x: (get_value(x, "No"), get_value(x, "Dwg"), get_value(x, "Joint"), x.get('order_index', 0))
    data_no = sorted(data, key=sort_key_no)
    for d in data_no:
        print(f"  No: {d['No']}, ISO: {d['Dwg']}, Joint: {d['Joint']}")

if __name__ == "__main__":
    test_hierarchy_sort()
