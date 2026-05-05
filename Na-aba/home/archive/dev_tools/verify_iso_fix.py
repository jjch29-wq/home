
import re
import pandas as pd

def to_float(val):
    try:
        if pd.isna(val): return 0.0
        return float(str(val).replace(',', '').strip())
    except:
        return 0.0

def get_value(item, key):
    val = item.get(key, "")
    if val is None: return ""
    
    if key in ["Ni", "Cr", "Mo", "Mn"]:
        return to_float(val)
        
    if key in ["No", "Joint", "Dwg"]:
        try: 
            if isinstance(val, (int, float)): 
                if float(val) == int(float(val)): return float(int(float(val)))
                return float(val)
            
            s_val = str(val).strip()
            if not s_val: return ""
            if re.match(r'^\d+(\.\d+)?$', s_val):
                f_val = float(s_val)
                if f_val == int(f_val): return float(int(f_val))
                return f_val
            if key in ["No", "Joint"]:
                cleaned = re.sub(r'[^0-9.]', '', s_val)
                if cleaned and not cleaned.endswith('.'):
                    return float(cleaned)
            return s_val.lower()
        except: pass
        return str(val).strip().lower()
    return str(val).strip().lower()

def normalize_iso(val):
    s_val = str(val).strip()
    try:
        if re.match(r'^\d+(\.\d+)?$', s_val):
            f_val = float(s_val)
            if f_val == int(f_val): return str(int(f_val))
            return str(f_val)
    except: pass
    return s_val.lower()

def test_fix():
    data = [
        {'Dwg': '1', 'Joint': 'A', 'order_index': 0},
        {'Dwg': '1.0', 'Joint': 'B', 'order_index': 1},
        {'Dwg': '1 ', 'Joint': 'C', 'order_index': 2},
        {'Dwg': '2', 'Joint': 'A', 'order_index': 3},
    ]
    
    print("Original data:")
    for d in data: print(f"  ISO: '{d['Dwg']}', Joint: {d['Joint']}")
    
    # Sort
    sort_key = lambda x: (get_value(x, "Dwg"), get_value(x, "Joint"), x.get('order_index', 0))
    data.sort(key=sort_key)
    
    print("\nSorted data (NEW logic):")
    for d in data: print(f"  ISO: '{d['Dwg']}', Joint: {d['Joint']} -> Sort Key: {get_value(d, 'Dwg')}")

    # Visual grouping check
    print("\nVisual Grouping (NEW logic):")
    last_iso = None
    group_count = 0
    for d in data:
        curr_iso = d.get('Dwg', '')
        norm_iso = normalize_iso(curr_iso)
        
        if last_iso is not None and norm_iso != normalize_iso(last_iso):
            group_count += 1
            print(f"--- Group Changed to {group_count} ---")
        print(f"  ISO: '{d['Dwg']}' (Normalized: {norm_iso})")
        last_iso = curr_iso

if __name__ == "__main__":
    test_fix()
