import pandas as pd
import re

# Simulate daily_usage_df
data = {
    'Date': ['2026-02-25'],
    'Site': ['가스공사 (의왕)'],
}
df = pd.DataFrame(data)

filter_site = '가스공사 (의왕)'

# Method 1: Partial match with str.contains (Tab 6 logic)
try:
    match_partial = df[df['Site'].astype(str).str.contains(filter_site, na=False, case=False)]
    print(f"Partial match (Tab 6) count: {len(match_partial)}")
    if len(match_partial) > 0:
        print(f"Matched Site: {match_partial['Site'].iloc[0]}")
    else:
        print("Tab 6 failed to match.")
except Exception as e:
    print(f"Tab 6 error: {e}")

# Method 2: Exact match (Tab 8 logic)
match_exact = df[df['Site'] == filter_site]
print(f"Exact match (Tab 8) count: {len(match_exact)}")
if len(match_exact) > 0:
    print(f"Matched Site: {match_exact['Site'].iloc[0]}")
else:
    print("Tab 8 failed to match.")

# Method 3: Partial match with regex=False
match_literal = df[df['Site'].astype(str).str.contains(filter_site, na=False, case=False, regex=False)]
print(f"Literal partial match count: {len(match_literal)}")
if len(match_literal) > 0:
    print(f"Matched Site: {match_literal['Site'].iloc[0]}")
