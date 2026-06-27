
import re

def get_natural_key(val):
    if val is None: return [[2, ""]]
    
    try:
        s_val = str(val).strip()
        if re.match(r'^-?\d+(\.\d+)?$', s_val):
            return [[0, float(s_val)]]
    except: pass

    s_val = str(val).strip().lower()
    if not s_val: return [[2, ""]]
    
    def segment_to_tuple(text):
        if text.isdigit():
            return (0, int(text))
        return (1, text)
    
    return [[1]] + [segment_to_tuple(c) for c in re.split(r'(\d+)', s_val) if c]

def test_type_safe_sort():
    data = ['1', '1A', '1B', '2', '10', '2A', 'ISO-1', 'ISO-10', 'ISO-2', 'A1', 'A10', 'A2', '', None]
    
    print("Testing Type-Safe Interleaved Natural Sort:")
    try:
        sorted_data = sorted(data, key=get_natural_key)
        for item in sorted_data:
            print(f"  Item: {str(item):6} -> Key: {get_natural_key(item)}")
        
        print("\nSUCCESS: Sorted without TypeErrors!")
        
        # Verify specific orderings
        assert sorted_data.index('1') < sorted_data.index('1A') < sorted_data.index('2')
        assert sorted_data.index('ISO-2') < sorted_data.index('ISO-10')
        assert sorted_data.index('A2') < sorted_data.index('A10')
        print("All assertions passed!")

    except TypeError as e:
        print("\nFAILURE: Still got TypeError:", e)
    except Exception as e:
        print("\nFAILURE: Unexpected error:", e)

if __name__ == "__main__":
    test_type_safe_sort()
