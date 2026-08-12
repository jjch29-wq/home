from pathlib import Path
import sys

import fitz
from docx import Document
from docx.oxml.ns import qn


sys.stdout.reconfigure(encoding="utf-8")

pdf_path = Path(r"C:\Users\-\Downloads\RLNG-000-MT-SP-6313_1.pdf")
pdf = fitz.open(pdf_path)
for page_number in (25, 26):
    text = pdf[page_number - 1].get_text("text") or ""
    print(f"PDF PAGE {page_number}\n{text}\n")

docx_path = Path(r"C:\Users\-\PMI\PAUT_SIS-H-264_작업본.docx")
document = Document(docx_path)
for index, paragraph in enumerate(document.paragraphs):
    text = " ".join(paragraph.text.split())
    drawings = paragraph._p.xpath(".//w:drawing")
    pictures = paragraph._p.xpath(".//w:pict")
    if 305 <= index <= 335 or drawings or pictures:
        if 305 <= index <= 335 or (drawings or pictures):
            print(
                f"DOCX P{index}: drawings={len(drawings)} pict={len(pictures)} "
                f"text={text[:700]!r}"
            )
            for blip in paragraph._p.xpath(".//a:blip"):
                relationship_id = blip.get(qn("r:embed"))
                if relationship_id:
                    part = document.part.related_parts[relationship_id]
                    print(
                        f"  IMAGE {part.partname}: {len(part.blob)} bytes, "
                        f"content_type={part.content_type}"
                    )
