import olefile
import zlib
import os

hwp_path = r'C:\Users\-\OneDrive\바탕 화면\test_hwp_out.hwp'
out_path = r'c:\Users\-\OneDrive\바탕 화면\home\Na-aba\home\scratch\hwp_output.txt'

try:
    f = olefile.OleFileIO(hwp_path)
    dirs = f.listdir()
    
    text = ""
    # Try PrvText first
    if ['PrvText'] in dirs:
        stream = f.openstream('PrvText')
        data = stream.read()
        text = data.decode('utf-16le', errors='ignore')
    else:
        # Fallback to BodyText
        for d in dirs:
            if d[0] == 'BodyText':
                stream = f.openstream(d)
                data = stream.read()
                try:
                    decompressed = zlib.decompress(data, -15)
                    text += decompressed.decode('utf-16le', errors='ignore')
                except:
                    pass
                    
    with open(out_path, 'w', encoding='utf-8') as outfile:
        outfile.write(text)
    print("Successfully extracted HWP text.")
        
except Exception as e:
    print(f"Error: {e}")
