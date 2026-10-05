from pathlib import Path

import fitz
from PIL import Image, ImageDraw


PDF = Path(r"I:\주진철\도서\자격증\비파괴검사 기술사\기술사 2022-1.pdf")
OUT = Path("tmp/pdfs/2022-1-samples")
OUT.mkdir(parents=True, exist_ok=True)

doc = fitz.open(PDF)
sample_pages = sorted(set(list(range(12)) + list(range(19, doc.page_count, 20)) + [doc.page_count - 1]))
thumbs = []
for page_index in sample_pages:
    page = doc[page_index]
    pix = page.get_pixmap(matrix=fitz.Matrix(0.7, 0.7), alpha=False)
    image = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    image.thumbnail((520, 720))
    canvas = Image.new("RGB", (540, 760), "white")
    canvas.paste(image, ((540 - image.width) // 2, 30))
    ImageDraw.Draw(canvas).text((10, 8), f"PDF page {page_index + 1}", fill="black")
    thumbs.append(canvas)

cols = 4
rows = (len(thumbs) + cols - 1) // cols
sheet = Image.new("RGB", (cols * 540, rows * 760), "#cccccc")
for i, thumb in enumerate(thumbs):
    sheet.paste(thumb, ((i % cols) * 540, (i // cols) * 760))
sheet.save(OUT / "contact-sheet.jpg", quality=88)
print(OUT / "contact-sheet.jpg")
