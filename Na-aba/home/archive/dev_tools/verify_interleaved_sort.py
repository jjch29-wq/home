
import re

def get_natural_key(val):
    if val is None: return [2, ""]
    
    # Simulating pure numeric detection
    try:
        s_val = str(val).strip()
        if re.match(r'^-?\d+(\.\d+)?$', s_val):
            return [0, float(s_val)]
    except: pass

    s_val = str(val).strip().lower()
    if not s_val: return [2, ""]
    
    def convert(text):
        return int(text) if text.isdigit() else text
    
    return [1] + [convert(c) for c in re.split(r'(\d+)', s_val) if c]

def test_interleaved_sort():
    # Test cases that often cause "jumbled" results
    data = ['1', '1A', '1B', '2', '10', '2A', 'ISO-1', 'ISO-10', 'ISO-2', 'A1', 'A10', 'A2']
    
    print("Testing Interleaved Natural Sort with case-sensitive and mixed data:")
    sorted_data = sorted(data, key=get_natural_key)
    
    for item in sorted_data:
        print(f"  Item: {item:6} -> Key: {get_natural_key(item)}")

    # Specific check for 1, 1A, 2
    idx1 = sorted_data.index('1')
    idx1A = sorted_data.index('1A')
    idx2 = sorted_data.index('2')
    
    if idx1 < idx1A < idx2:
        print("\nSUCCESS: '1' < '1A' < '2' confirmed!")
    else:
        print("\nFAILURE: Ordering of 1, 1A, 2 is incorrect.")

    # Check for ISO-2 < ISO-10
    idx_iso2 = sorted_data.index('ISO-2')
    idx_iso10 = sorted_data.index('ISO-10')
    if idx_iso2 < idx_iso10:
        print("SUCCESS: 'ISO-2' < 'ISO-10' confirmed!")
    else:
        print("FAILURE: ISO natural sort failed.")

if __name__ == "__main__":
    test_interleaved_sort()
