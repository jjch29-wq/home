
import re

def normalize_iso(val):
    s_val = str(val).strip()
    try:
        if re.match(r'^\d+(\.\d+)?$', s_val):
            f_val = float(s_val)
            if f_val == int(f_val): return str(int(f_val))
            return str(f_val)
    except: pass
    return s_val.lower()

def simulate_populate(data_list):
    last_iso = None
    results = []
    
    for idx, item in enumerate(data_list):
        curr_iso = item.get('Dwg', '')
        norm_iso = normalize_iso(curr_iso)
        
        is_new_iso = (last_iso is None or normalize_iso(last_iso) != norm_iso)
        is_block_start = (idx % 3 == 0)
        
        display_iso = curr_iso if (is_new_iso or is_block_start) else ""
        
        # Proposed Fix: Only hide if it's NOT a new ISO group
        if item.get('is_merged_iso') and not is_new_iso: 
            display_iso = ""
        
        results.append(display_iso)
        last_iso = curr_iso
    return results

# Case 1: Manual merge and then sort
data = [
    {'Dwg': 'ISO-1', 'order_index': 0}, # Original base
    {'Dwg': 'ISO-1', 'order_index': 1, 'is_merged_iso': True},
    {'Dwg': 'ISO-1', 'order_index': 2, 'is_merged_iso': True},
]

print("Original display:", simulate_populate(data))

# Sort reverse
data_sorted = sorted(data, key=lambda x: x['order_index'], reverse=True)
print("Sorted display (Reverse):", simulate_populate(data_sorted))

# Case 3: Data Corruption Test
# If we merge using a suppressed item as the "anchor"
data3 = [
    {'Dwg': 'ISO-5', 'id': 1},
    {'Dwg': 'ISO-5', 'id': 2}, # This one would be "" in view
]

# Simulate corruption if logic was bad:
# item_view_val = "" if norm_iso == last_iso else "ISO-5"
# In the original bad code: first_iso = tree.set(selected[0], "#4") -> would be "" if selected[0] was id 2
def simulate_merge_iso(selected_items):
    # Old bad logic would take suppressed value from view
    # Correct logic takes from extracted_data
    # We'll just print if it's empty
    first_iso = selected_items[0]['Dwg']
    return first_iso

print("Merge result (Correct):", simulate_merge_iso(data3))

# Case 4: Misaligned Joint Test
# Current bad logic hides Joint-2 if it starts at index 1 and ISO is the same.
data4 = [
    {'Dwg': 'ISO-1', 'Joint': 'J-1'},
    {'Dwg': 'ISO-1', 'Joint': 'J-2'}, # This would be HIDDEN in current bad logic
]

def simulate_populate_v3(data_list):
    last_iso = None
    last_joint = None
    results = []
    
    for idx, item in enumerate(data_list):
        curr_iso = item.get('Dwg', '')
        curr_joint = item.get('Joint', '')
        norm_iso = normalize_iso(curr_iso)
        
        is_new_iso = (last_iso is None or normalize_iso(last_iso) != norm_iso)
        is_new_joint = (last_joint is None or last_joint != curr_joint)
        is_block_start = (idx % 3 == 0)
        
        # Fixed logic: show if ISO changed OR Joint changed OR block start
        is_show = is_new_iso or is_new_joint or is_block_start
        
        display_iso = curr_iso if is_show else ""
        display_joint = curr_joint if is_show else ""
        
        results.append((display_iso, display_joint))
        last_iso = curr_iso
        last_joint = curr_joint
    return results

print("Misaligned Joint (Fixed Logic):", simulate_populate_v3(data4))
