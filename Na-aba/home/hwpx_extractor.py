import zipfile
import xml.etree.ElementTree as ET

filepath = r"C:\Users\-\OneDrive\바탕 화면\3. 착수 전 안전보건회의 자료(수급업체 제공용).hwpx"
output_path = r"c:\Users\-\OneDrive\바탕 화면\home\Na-aba\home\hwpx_output.txt"

text_content = []

try:
    with zipfile.ZipFile(filepath, 'r') as z:
        for filename in z.namelist():
            if filename.startswith('Contents/section') and filename.endswith('.xml'):
                with z.open(filename) as f:
                    tree = ET.parse(f)
                    root = tree.getroot()
                    
                    for elem in root.iter():
                        if elem.tag.endswith('}p'): # Paragraph
                            # Gather text from children
                            p_text = ""
                            for child in elem.iter():
                                if child.tag.endswith('}t') and child.text:
                                    p_text += child.text
                            if p_text:
                                text_content.append(p_text)

    with open(output_path, "w", encoding="utf-8") as out:
        out.write("\n".join(text_content))
    print("Success")
except Exception as e:
    print(f"Error: {e}")
