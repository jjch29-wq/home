import argparse
from pathlib import Path

import easyocr
import fitz
import numpy as np


parser = argparse.ArgumentParser()
parser.add_argument("--start", type=int, default=11, help="1-based first page")
parser.add_argument("--end", type=int, default=167, help="1-based last page")
args = parser.parse_args()

pdf = Path(r"I:\주진철\도서\자격증\비파괴검사 기술사\기술사 2022-1.pdf")
out = Path("tmp/pdfs/2022-1-ocr.txt")
doc = fitz.open(pdf)
reader = easyocr.Reader(["ko", "en"], gpu=False, model_storage_directory=str(Path.home() / ".EasyOCR" / "model"), download_enabled=False)

with out.open("a", encoding="utf-8") as stream:
    for page_no in range(args.start, min(args.end, doc.page_count) + 1):
        page = doc[page_no - 1]
        pix = page.get_pixmap(matrix=fitz.Matrix(1.35, 1.35), alpha=False)
        image = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, 3)
        lines = reader.readtext(image, detail=0, paragraph=False, width_ths=0.8)
        stream.write(f"\n\n===== PDF PAGE {page_no} =====\n")
        stream.write("\n".join(lines))
        stream.flush()
        print(f"page {page_no}: {len(lines)} lines", flush=True)

print(out)
