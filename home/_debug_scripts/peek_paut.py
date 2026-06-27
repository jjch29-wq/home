import sqlite3
import pandas as pd

conn = sqlite3.connect('pmi_database.db')
try:
    df = pd.read_sql_query("SELECT * FROM Materials WHERE 품목명 LIKE '%PAUT%'", conn)
    print("--- Materials ---")
    print(df)
    
    df_usage = pd.read_sql_query("SELECT * FROM DailyUsage WHERE MaterialID IN (SELECT MaterialID FROM Materials WHERE 품목명 LIKE '%PAUT%')", conn)
    print("\n--- DailyUsage ---")
    print(df_usage)
except Exception as e:
    print(f"Error: {e}")
finally:
    conn.close()
