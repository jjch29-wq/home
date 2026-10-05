from pathlib import Path

import easyocr
import fitz
import numpy as np


pdf = Path(r"I:\주진철\도서\자격증\비파괴검사 기술사\기술사 2022-1.pdf")
out = Path("tmp/pdfs/2022-1-toc-retry.txt")
pages = [12, 13, 17, 18, 21]
doc = fitz.open(pdf)
reader = easyocr.Reader(["ko", "en"], gpu=False, model_storage_directory=str(Path.home() / ".EasyOCR" / "model"), download_enabled=False)
with out.open("w", encoding="utf-8") as stream:
    for page_no in pages:
        pix = doc[page_no - 1].get_pixmap(matrix=fitz.Matrix(1.8, 1.8), alpha=False)
        image = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, 3)
        lines = reader.readtext(image, detail=0, paragraph=False, width_ths=0.7)
        stream.write(f"\n===== PDF PAGE {page_no} =====\n")
        stream.write("\n".join(lines))
        print(f"page {page_no}: {len(lines)} lines", flush=True)
print(out)
