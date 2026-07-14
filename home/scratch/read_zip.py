import zipfile, xml.etree.ElementTree as ET
try:
    with zipfile.ZipFile('scratch/KS_B_0845_temp.docx') as z:
        xml_content = z.read('word/document.xml')
    root = ET.fromstring(xml_content)
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    texts = []
    for node in root.iterfind('.//w:t', ns):
        if node.text: texts.append(node.text)
    print(''.join(texts)[:3000])
except Exception as e:
    print('Error:', e)