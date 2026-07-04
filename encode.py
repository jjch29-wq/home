import base64
path = r'c:\Users\jjch2\Desktop\PMI\Assets\SITCO.png'
with open(path, 'rb') as f:
    b64 = base64.b64encode(f.read()).decode('utf-8')
    with open(r'c:\Users\jjch2\Desktop\PMI\b64_logo.py', 'w', encoding='utf-8') as out:
        out.write(f'DEFAULT_SITCO_LOGO_B64 = "{b64}"\n')
