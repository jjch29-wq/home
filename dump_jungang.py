import fitz
import pandas as pd
import traceback

pdf_path = r"C:\Users\jjch2\Desktop\2. 과업내용서(2026년 중앙지사 열수송관 비파괴검사용역 단가계약).pdf"
xlsx_path = r"C:\Users\jjch2\Desktop\산출내역서(2026년 중앙지사 열수송관 비파괴검사용역 단가계약).xlsx"

pdf_out = r"C:\Users\jjch2\Desktop\PMI\jungang_pdf.txt"
xlsx_out = r"C:\Users\jjch2\Desktop\PMI\jungang_xlsx.txt"

# PDF
try:
    doc = fitz.open(pdf_path)
    text = ""
    for page in doc:
        text += page.get_text()
    with open(pdf_out, "w", encoding="utf-8") as f:
        f.write(text)
except Exception as e:
    with open(pdf_out, "w", encoding="utf-8") as f:
        f.write(traceback.format_exc())

# XLSX
try:
    xls = pd.ExcelFile(xlsx_path)
    with open(xlsx_out, "w", encoding="utf-8") as f:
        f.write(f"Sheets: {xls.sheet_names}\n\n")
        for sheet_name in xls.sheet_names:
            f.write(f"\n{'='*20} Sheet: {sheet_name} {'='*20}\n")
            df = pd.read_excel(xls, sheet_name=sheet_name)
            for r_idx, row in df.iterrows():
                row_vals = [str(x).strip() for x in row if pd.notnull(x) and str(x).strip() != '']
                if row_vals:
                    f.write(f"Row {r_idx}: " + " | ".join(row_vals) + "\n")
except Exception as e:
    with open(xlsx_out, "w", encoding="utf-8") as f:
        f.write(traceback.format_exc())
