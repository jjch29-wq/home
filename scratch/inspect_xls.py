import os
import glob
import pandas as pd

def inspect_files():
    # Match all xls and xlsx files recursively
    files = []
    for ext in ["*.xls", "*.xlsx"]:
        files.extend(glob.glob(ext))
        files.extend(glob.glob("data/**/" + ext, recursive=True))
        files.extend(glob.glob("**/ " + ext, recursive=True))
    
    # Also find files by os.walk to be absolutely sure
    for root, dirs, filenames in os.walk("."):
        for filename in filenames:
            if filename.endswith(".xls") or filename.endswith(".xlsx"):
                files.append(os.path.join(root, filename))
                
    files = list(set([os.path.abspath(f) for f in files]))
    files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    
    print(f"Found {len(files)} unique Excel files.")
    for f in files:
        basename = os.path.basename(f)
        if "Material_Inventory" in basename:
            continue
        print(f"\n--- FILE: {basename} ---")
        try:
            with pd.ExcelFile(f) as xls:
                for sheet in xls.sheet_names[:1]:
                    print(f"Sheet: {sheet}")
                    df = pd.read_excel(xls, sheet_name=sheet, header=None, nrows=30)
                    print("--- Rows 8 to 30 ---")
                    for r_idx in range(7, min(len(df), 30)):
                        row_vals = []
                        for c_idx in range(min(len(df.columns), 15)):
                            row_vals.append(f"{c_idx}:{df.iloc[r_idx, c_idx]}")
                        print(f"Row {r_idx + 1}: " + " | ".join(row_vals))
        except Exception as e:
            print(f"Error: {e}")

if __name__ == "__main__":
    inspect_files()
