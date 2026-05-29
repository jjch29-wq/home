import os

def list_na_aba():
    print("Listing files in Na-aba:")
    for root, dirs, files in os.walk('Na-aba'):
        # Just show directories and count of files to avoid huge logs
        excel_files = [f for f in files if f.endswith('.xlsx') or f.endswith('.xls') or f.endswith('.xlsm')]
        if excel_files:
            print(f"Directory: {root} | Excel files count: {len(excel_files)}")
            for f in excel_files[:5]:
                print(f"  - {f}")
            if len(excel_files) > 5:
                print(f"  - ... and {len(excel_files) - 5} more")

list_na_aba()
