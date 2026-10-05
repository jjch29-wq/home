from pathlib import Path

import fitz
from PIL import Image, ImageDraw


ROOT = Path(r"I:\주진철\도서\자격증\비파괴검사 기술사")
OUT = Path("tmp/pdfs/ndt-inventory")
OUT.mkdir(parents=True, exist_ok=True)

for pdf in sorted(ROOT.glob("*.pdf")):
    doc = fitz.open(pdf)
    pages = sorted(set([0, 1, 2, 9, 19, 39, 79, 119, 159, doc.page_count - 1]))
    pages = [p for p in pages if 0 <= p < doc.page_count]
    thumbs = []
    for index in pages:
        pix = doc[index].get_pixmap(matrix=fitz.Matrix(0.65, 0.65), alpha=False)
        image = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        image.thumbnail((480, 680))
        canvas = Image.new("RGB", (500, 720), "white")
        canvas.paste(image, ((500 - image.width) // 2, 28))
        ImageDraw.Draw(canvas).text((10, 7), f"PDF {index + 1}", fill="black")
        thumbs.append(canvas)
    cols = 5
    rows = (len(thumbs) + cols - 1) // cols
    sheet = Image.new("RGB", (cols * 500, rows * 720), "#cccccc")
    for i, thumb in enumerate(thumbs):
        sheet.paste(thumb, ((i % cols) * 500, (i // cols) * 720))
    target = OUT / f"{pdf.stem}.jpg"
    sheet.save(target, quality=90)
    print(f"{pdf.name}\tpages={doc.page_count}\t{target}")
