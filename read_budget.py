import sqlite3
import pandas as pd

try:
    conn = sqlite3.connect("home/data/app_db.sqlite")
    df = pd.read_sql_query("SELECT name FROM sqlite_master WHERE type='table';", conn)
    print("Tables:", df["name"].tolist())
    
    for table in df["name"]:
        if "budget" in table.lower():
            b_df = pd.read_sql_query(f"SELECT * FROM {table}", conn)
            print(f"\nTable: {table}")
            print(b_df.to_string())
except Exception as e:
    print("Error:", e)
