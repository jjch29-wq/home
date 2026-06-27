
import re

def to_float(val):
    try:
        return float(str(val).replace(',', '').strip())
    except: return 0.0

def get_value(val, key):
    if val is None: return (2, "")
    
    if key in ["Ni", "Cr", "Mo", "Mn"]:
        return (0, to_float(val))
    
    s_val = str(val).strip()
    if not s_val: return (2, "")
    
    if re.match(r'^\d+(\.\d+)?$', s_val):
        try: 
            f_val = float(s_val)
            return (0, f_val)
        except: pass
    
    # Natural sort key
    natural_key = re.sub(r'(\d+)', lambda m: m.group(1).zfill(20), s_val.lower())
    return (1, natural_key)

def test_robust_sort():
    # Mixed data: pure numbers (varying formats), alphanumeric (varying numbers)
    raw_data = ['1', '1.0', '1  ', '2', '10', 'ISO-1', 'ISO-10', 'ISO-2', 'B-1']
    
    print("Testing robust sort with data:", raw_data)
    
    try:
        sorted_data = sorted(raw_data, key=lambda x: get_value(x, "Dwg"))
        print("\nSorted successfully!")
        for item in sorted_data:
            kv = get_value(item, "Dwg")
            print(f"  Item: '{item:6}' -> Key: {kv}")
            
    except TypeError as e:
        print("\nFAILED: TypeError during sort:", e)
    except Exception as e:
        print("\nFAILED: Unexpected error:", e)

if __name__ == "__main__":
    test_robust_sort()
