
import pandas as pd
import numpy as np

def test_fillna_behavior():
    print("Testing fillna('') on float column...")
    df = pd.DataFrame({'A': [1.0, 2.0, np.nan], 'B': [1, 2, 3]})
    print(f"Original dtypes:\n{df.dtypes}")
    
    try:
        df['A'] = df['A'].fillna('')
        print("fillna('') successful.")
        print(f"New dtypes:\n{df.dtypes}")
        print(df)
    except Exception as e:
        print(f"fillna('') failed: {e}")

    # Simulating the concatenation
    print("\nTesting concatenation with float column...")
    new_row = pd.DataFrame({'A': [4.0], 'B': [4]})
    try:
        df = pd.concat([df, new_row], ignore_index=True)
        print("Concatenation successful.")
        print(f"Result dtypes:\n{df.dtypes}")
        print(df)
    except Exception as e:
        print(f"Concatenation failed: {e}")

def test_material_manager_logic():
    print("\nSimulating MaterialManager load_data logic...")
    # Simulate reading excel with NaNs in Material ID (float)
    df = pd.DataFrame({
        'Material ID': [1.0, 2.0, np.nan],
        'Site': ['A', 'B', np.nan]
    })
    print(f"Loaded dtypes:\n{df.dtypes}")
    
    # MaterialManager logic
    string_columns = ['Site', 'Material ID']
    for col in string_columns:
        if col in df.columns:
            print(f"Filling {col} with ''")
            try:
                df[col] = df[col].fillna('')
            except Exception as e:
                print(f"Error filling {col}: {e}")
                
    print(f"After fillna dtypes:\n{df.dtypes}")
    
    # Simulate add_daily_usage_entry
    print("Simulating add_daily_usage_entry...")
    new_entry = {
         'Material ID': 4.0, # Float
         'Site': 'D'
    }
    try:
        df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)
        print("Concatenation successful.")
        print(df)
    except Exception as e:
        print(f"Concatenation failed: {e}")

if __name__ == "__main__":
    test_fillna_behavior()
    test_material_manager_logic()
