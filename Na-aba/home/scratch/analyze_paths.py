import fitz
import os

def analyze_pdf_paths(pdf_path):
    doc = fitz.open(pdf_path)
    page = doc.load_page(0)
    paths = page.get_drawings()
    
    print(f"Total drawing paths: {len(paths)}")
    
    # Check first few paths to see structure
    for i, path in enumerate(paths[:10]):
        print(f"Path {i}: {path['type']} - {len(path['items'])} items")
        for item in path['items']:
            print(f"  Item: {item[0]}") # 'l' for line, 'c' for curve, 're' for rect, 'qu' for quad

if __name__ == "__main__":
    pdf_file = r"C:\Users\-\OneDrive\문서\카카오톡 받은 파일\A1-C001-Rev 1.pdf"
    analyze_pdf_paths(pdf_file)
