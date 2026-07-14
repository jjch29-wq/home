import sqlite3
import os
import requests
import json

def main():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    db_path = os.path.join(current_dir, "household.db")
    
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute("SELECT value FROM settings WHERE key='gemini_api_key'")
    row = cursor.fetchone()
    
    if not row or not row[0]:
        print("No API key found in DB.")
        return
        
    api_key = row[0]
    
    url = f"https://generativelanguage.googleapis.com/v1beta/models?key={api_key}"
    resp = requests.get(url)
    if resp.status_code == 200:
        models = resp.json().get("models", [])
        print("Available models:")
        for m in models:
            print(m["name"])
    else:
        print("Error:", resp.text)

if __name__ == "__main__":
    main()
