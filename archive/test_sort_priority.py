
import re

def to_float(val):
    try: return float(str(val).replace(',', '').strip())
    except: return 0.0

def get_value(val, key):
    if val is None: return (2, "")
    if key in ["Ni", "Cr", "Mo", "Mn"]: return (0, to_float(val))
    s_val = str(val).strip()
    if not s_val: return (2, "")
    if re.match(r'^\d+(\.\d+)?$', s_val):
        try: return (0, float(s_val))
        except: pass
    natural_key = re.sub(r'(\d+)', lambda m: m.group(1).zfill(20), s_val.lower())
    return (1, natural_key)

def test_no_sort_priority():
    data = [
        {'Dwg': 'ISO-B', 'No': '1', 'Joint': 'J1'},
        {'Dwg': 'ISO-A', 'No': '2', 'Joint': 'J1'}
    ]
    
    # Current logic for "No" sort (Dwg is primary)
    print("Sorting by 'No' with current logic (Dwg primary):")
    # sort_key = lambda x: (get_value(x['Dwg'], "Dwg"), get_value(x['No'], "No"), ...)
    sort_key_current = lambda x: (get_value(x['Dwg'], "Dwg"), get_value(x['No'], "No"))
    
    data_current = sorted(data, key=sort_key_current)
    for d in data_current:
        print(f"  ISO: {d['Dwg']}, No: {d['No']}")
    
    print("\nSorting by 'No' with refined logic (No primary):")
    sort_key_refined = lambda x: (get_value(x['No'], "No"), get_value(x['Dwg'], "Dwg"))
    data_refined = sorted(data, key=sort_key_refined)
    for d in data_refined:
        print(f"  ISO: {d['Dwg']}, No: {d['No']}")

if __name__ == "__main__":
    test_no_sort_priority()
