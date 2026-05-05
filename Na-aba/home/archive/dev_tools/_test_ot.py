import sqlite3
import pandas as pd

db_path = r'c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\pmi_database.db'
conn = sqlite3.connect(db_path)
cur = conn.cursor()
cur.execute("SELECT name FROM sqlite_master WHERE type='table';")
tables = cur.fetchall()

print("Tables:", tables)

for t in tables:
    if 'daily' in t[0].lower() or 'usage' in t[0].lower() or 'performance' in t[0].lower() or 'sales' in t[0].lower():
        print(f"\nTable: {t[0]}")
        try:
            df = pd.read_sql_query(f'SELECT * FROM "{t[0]}" LIMIT 10', conn)
            cols = [c for c in df.columns if 'WorkTime' in c or 'OT' in c]
            if cols:
                print("Cols:", cols)
                for _, row in df.iterrows():
                    print({c: row.get(c) for c in cols})
        except Exception as e:
            print(e)
