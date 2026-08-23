import pandas as pd
try:
    df = pd.read_excel('home/data/Material_Inventory.xlsx', sheet_name='Budget')
    print('Columns:', df.columns.tolist())
    
    # Try to find site column
    site_col = next((c for c in df.columns if 'site' in str(c).lower() or '현장' in str(c)), None)
    if site_col:
        matches = df[df[site_col].astype(str).str.contains('중앙지사', na=False)]
        if not matches.empty:
            for idx, row in matches.iterrows():
                # try to find revenue column
                rev_col = next((c for c in df.columns if 'revenue' in str(c).lower() or '금액' in str(c) or '매출' in str(c)), None)
                revenue = row.get(rev_col, 0) if rev_col else 0
                if pd.isna(revenue): revenue = 0
                print(f"Site: {row[site_col]}, Contract Amount (Revenue): {revenue:,.0f}원")
        else:
            print('No matching site found. All sites:')
            print(df[site_col].tolist())
    else:
        print('Could not find Site column')
except Exception as e:
    print('Error:', e)
