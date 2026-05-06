
def test_merge_display_logic():
    data_list = [
        {'No': '1', 'Joint': 'J1', 'Dwg': 'ISO1'},                 # Natural 1
        {'No': '2', 'Joint': 'J1', 'Dwg': 'ISO1'},                 # Natural 2 (Duplicate)
        {'No': '3', 'Joint': 'J2', 'Dwg': 'ISO2'},                 # Natural 3
        {'No': '4', 'Joint': 'J2', 'Dwg': 'ISO2', 'is_merged_iso': True, 'is_merged_joint': True}, # User Merged
    ]
    
    results = []
    
    for idx, item in enumerate(data_list):
        curr_iso = item.get('Dwg', '')
        curr_joint = item.get('Joint', '')
        
        display_iso = curr_iso
        display_joint = curr_joint
        
        # Logic from JJCHSITPMI-V2-Unified.py
        if item.get('is_merged_iso'):
            display_iso = ""
        if item.get('is_merged_joint'):
            display_joint = ""
            
        results.append((display_iso, display_joint))
        print(f"Row {idx+1}: ISO='{display_iso}', JOINT='{display_joint}'")

    # Row 1 & 2 are natural duplicates, should both show
    assert results[0] == ('ISO1', 'J1')
    assert results[1] == ('ISO1', 'J1')
    
    # Row 3 is natural, should show
    assert results[2] == ('ISO2', 'J2')
    
    # Row 4 is user merged, should be empty
    assert results[3] == ('', '')
    
    print("Merge display logic test passed!")

if __name__ == "__main__":
    test_merge_display_logic()
