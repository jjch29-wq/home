import pandas as pd

def test_logic():
    # Setup mock data for two groups (reports)
    # Report 1 has 0 defect and 0 rev
    # Report 2 has 2 defects and 1 rev
    data = {
        "Report No": ["REP-0001", "REP-0001", "REP-0002", "REP-0002"],
        "Joint": ["01", "02", "03", "04"],
        "Defect": [0, 0, 1, 1],
        "Rev": ["O", "O", "O", "R1"]
    }
    df = pd.DataFrame(data)
    
    # Process
    film_col = None
    joint_col = "Joint"
    defect_col = "Defect"
    rev_col = "Rev"
    report_col = "Report No"
    label_col = "Joint"
    
    df['_temp_numeric_defect'] = pd.to_numeric(df[defect_col].astype(str).str.extract(r'(\d+\.?\d*)')[0], errors='coerce')
    df['_temp_numeric_rev'] = pd.to_numeric(df[rev_col].astype(str).str.extract(r'(\d+\.?\d*)')[0], errors='coerce')
    
    new_dfs = []
    grand_defect_total = 0
    grand_rev_total = 0
    grand_joint_total = 0
    
    for rep_no, group in df.groupby(report_col, sort=False):
        new_dfs.append(group)
        
        # Joint Count (using total count instead of nunique)
        sub_joint = group[joint_col].dropna().astype(str).str.strip().count()
        grand_joint_total += sub_joint
        
        # Defect Count
        sub_defect = 0
        num_sum = group['_temp_numeric_defect'].sum()
        if num_sum > 0:
            sub_defect = num_sum
        else:
            valid_vals = group[defect_col].dropna().astype(str).str.strip().str.lower()
            sub_defect = valid_vals[valid_vals.str.contains('reject|fail|ng|불합격|결함', na=False)].count()
        grand_defect_total += sub_defect
        
        # Rev Count
        sub_rev = 0
        num_sum = group['_temp_numeric_rev'].sum()
        if num_sum > 0:
            sub_rev = num_sum
        else:
            valid_vals = group[rev_col].dropna().astype(str).str.strip().str.lower()
            sub_rev = valid_vals[valid_vals.str.contains(r'^r|repair|보수|수리|개정', na=False)].count()
        grand_rev_total += sub_rev
        
        # Subtotal row
        sub_row = {col: "" for col in df.columns}
        sub_row[label_col] = "Sub-Total"
        sub_row[report_col] = rep_no
        sub_row[joint_col] = int(sub_joint)
        
        # Display only if >= 1
        if sub_defect > 0:
            sub_row[defect_col] = int(sub_defect)
        else:
            sub_row[defect_col] = ""
            
        if sub_rev > 0:
            sub_row[rev_col] = int(sub_rev)
        else:
            sub_row[rev_col] = ""
            
        new_dfs.append(pd.DataFrame([sub_row]))
        
    combined_df = pd.concat(new_dfs, ignore_index=True)
    
    # Grand Total row
    total_row = {col: "" for col in combined_df.columns}
    total_row[label_col] = "Grand Total"
    total_row[joint_col] = int(grand_joint_total)
    
    if grand_defect_total > 0:
        total_row[defect_col] = int(grand_defect_total)
    else:
        total_row[defect_col] = ""
        
    if grand_rev_total > 0:
        total_row[rev_col] = int(grand_rev_total)
    else:
        total_row[rev_col] = ""
        
    combined_df = pd.concat([combined_df, pd.DataFrame([total_row])], ignore_index=True)
    
    # Clean temp cols
    combined_df.drop(columns=['_temp_numeric_defect', '_temp_numeric_rev'], inplace=True)
    
    print(combined_df.to_string())

test_logic()
