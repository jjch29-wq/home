
import re

def get_value(val, key):
    # Simplified version of the code in the main script
    if val is None: return ""
    
    try: 
        s_val = str(val).strip()
        if not s_val: return ""
        
        # 완전한 숫자 형태인 경우 (예: "1.0", "01")
        if re.match(r'^\d+(\.\d+)?$', s_val):
            f_val = float(s_val)
            if f_val == int(f_val): return float(int(f_val))
            return f_val
            
        return s_val.lower()
    except: pass
    return str(val).strip().lower()

def test_mixed_sort():
    data = ['1', '1.0', 'A', '10', '2']
    try:
        data.sort(key=lambda x: get_value(x, "Dwg"))
        print("Sorted successfully:", data)
    except TypeError as e:
        print("Caught expected TypeError:", e)

if __name__ == "__main__":
    test_mixed_sort()
