import pandas as pd

filepath = r"F:\내 드라이브\Office\착공\Project PROVIDENCE 중 비파괴검사\준공\박스라벨.xls"
df = pd.read_excel(filepath, sheet_name='2021 (2)', header=None)

print("--- 2021 (2) Sheet Data ---")
for r in range(min(10, len(df))):
    row_data = []
    for c in range(len(df.columns)):
        val = df.iloc[r, c]
        if pd.notna(val):
            row_data.append(f"({r},{c}): {val}")
    print(f"Row {r}: " + " | ".join(row_data))
