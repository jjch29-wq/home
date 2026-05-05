
import re

def get_natural_key(val):
    if val is None: return [(2, "")]
    s_val = str(val).strip().lower()
    if not s_val: return [(2, "")]
    
    def segment_to_tuple(text):
        if text.isdigit():
            return (0, int(text))
        return (1, text)
    
    return [segment_to_tuple(c) for c in re.split(r'(\d+)', s_val) if c]

def test_unified_sort():
    data = ['1', '1A', '1B', '2', '10', '2A', 'ISO-1', 'ISO-10', 'ISO-2', 'A1', 'A10', 'A2', '', None]
    
    print("Testing Unified Interleaved Natural Sort:")
    try:
        sorted_data = sorted(data, key=get_natural_key)
        for item in sorted_data:
            print(f"  Item: {str(item):6} -> Key: {get_natural_key(item)}")
        
        print("\nChecking Orderings...")
        
        # 1 < 1A < 1B < 2 < 2A < 10
        assert sorted_data.index('1') < sorted_data.index('1A') < sorted_data.index('1B') < sorted_data.index('2') < sorted_data.index('2A') < sorted_data.index('10')
        print("SUCCESS: '1' < '1A' < '1B' < '2' < '2A' < '10' confirmed!")
        
        # ISO-1 < ISO-2 < ISO-10
        assert sorted_data.index('ISO-1') < sorted_data.index('ISO-2') < sorted_data.index('ISO-10')
        print("SUCCESS: 'ISO-1' < 'ISO-2' < 'ISO-10' confirmed!")
        
        # A1 < A2 < A10
        assert sorted_data.index('A1') < sorted_data.index('A2') < sorted_data.index('A10')
        print("SUCCESS: 'A1' < 'A2' < 'A10' confirmed!")

    except AssertionError as e:
        print("\nFAILURE: Assertion failed. Order is not as expected.")
    except Exception as e:
        print("\nFAILURE: Unexpected error:", e)

if __name__ == "__main__":
    test_unified_sort()
