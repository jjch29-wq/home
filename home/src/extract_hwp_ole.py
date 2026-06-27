import olefile
import zlib
import struct
import os

hwp_path = os.path.abspath(r'C:\Users\-\OneDrive\바탕 화면\4.4.1 위험성평가표(RT).hwp')
out_path = r'c:\Users\-\OneDrive\바탕 화면\home\Na-aba\home\hwp_rt_out.txt'

try:
    f = olefile.OleFileIO(hwp_path)
    dirs = f.listdir()
    
    # Check if PrvText exists
    if ['PrvText'] in dirs:
        stream = f.openstream('PrvText')
        data = stream.read()
        # Decode PrvText (UTF-16LE)
        text = data.decode('utf-16le', errors='ignore')
        with open(out_path, 'w', encoding='utf-8') as outfile:
            outfile.write(text)
        print("Successfully extracted PrvText.")
    else:
        print("PrvText stream not found. Trying to extract BodyText...")
        text = ""
        for d in dirs:
            if d[0] == 'BodyText':
                stream = f.openstream(d)
                data = stream.read()
                try:
                    # Try to decompress
                    decompressed = zlib.decompress(data, -15)
                    text += decompressed.decode('utf-16le', errors='ignore')
                except:
                    pass
        with open(out_path, 'w', encoding='utf-8') as outfile:
            outfile.write(text)
        print("Extracted from BodyText.")
        
except Exception as e:
    print(f"Error: {e}")
