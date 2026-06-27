import fitz
import sys

def read_pdf():
    path = r'C:\Users\-\OneDrive\바탕 화면\2026년 7월 안전보건협의체 참석요청.pdf'
    try:
        doc = fitz.open(path)
        with open('output.txt', 'w', encoding='utf-8') as f:
            for page in doc:
                f.write(page.get_text() + '\n')
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    read_pdf()
