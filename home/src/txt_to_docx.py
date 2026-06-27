from docx import Document

input_txt = r"c:\Users\-\OneDrive\바탕 화면\home\Na-aba\home\hwpx_output.txt"
output_docx = r"C:\Users\-\OneDrive\바탕 화면\3. 착수 전 안전보건회의 자료(수급업체 제공용).docx"

def convert_txt_to_docx(txt_file, docx_file):
    doc = Document()
    
    with open(txt_file, 'r', encoding='utf-8') as f:
        for line in f:
            # Add each line as a paragraph
            doc.add_paragraph(line.strip())
            
    doc.save(docx_file)
    print(f"Successfully saved to {docx_file}")

if __name__ == "__main__":
    convert_txt_to_docx(input_txt, output_docx)
