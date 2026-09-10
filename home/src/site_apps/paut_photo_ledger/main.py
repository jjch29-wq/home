"""독립형 PAUT 사진대장 실행기."""
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from site_apps.paut_photo_ledger.src.app import PautPhotoLedgerApp


def main():
    import tkinter as tk

    root = tk.Tk()
    PautPhotoLedgerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
