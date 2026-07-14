import sys
import requests
import json

def get_models(api_key):
    url = f"https://generativelanguage.googleapis.com/v1beta/models?key={api_key}"
    resp = requests.get(url)
    if resp.status_code == 200:
        models = resp.json().get("models", [])
        for m in models:
            print(m["name"])
    else:
        print("Error:", resp.text)

if __name__ == "__main__":
    if len(sys.argv) > 1:
        get_models(sys.argv[1])
