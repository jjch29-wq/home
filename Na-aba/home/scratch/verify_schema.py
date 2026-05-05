import pandas as pd
import os
import sys

# Mocking the MaterialManager environment
class MockManager:
    def _sync_dataframe_schema(self, df, sheet_name):
        if df is None: return None
        schemas = {
            'DailyUsage': ['Date', 'Site', '차량번호', '주행거리', '차량점검', '차량비고']
        }
        required_cols = schemas.get(sheet_name, [])
        for col in required_cols:
            if col not in df.columns:
                df[col] = ""
        return df

# 1. Create a dummy DataFrame missing columns
df = pd.DataFrame({'Date': ['2026-04-21'], 'Site': ['Test Site']})
print("Original Columns:", df.columns.tolist())

# 2. Sync schema
manager = MockManager()
synced_df = manager._sync_dataframe_schema(df, 'DailyUsage')
print("Synced Columns:", synced_df.columns.tolist())

# 3. Check if vehicle columns exist
for col in ['차량번호', '주행거리', '차량점검', '차량비고']:
    if col in synced_df.columns:
        print(f"PASS: Column '{col}' successfully created.")
    else:
        print(f"FAIL: Column '{col}' missing.")
