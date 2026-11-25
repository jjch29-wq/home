# -*- coding: utf-8 -*-
import sys
import os
import glob
import math
import traceback
import re
import warnings
import json
import tempfile
import pandas as pd
import openpyxl
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import threading
import xlsxwriter
import datetime
from PIL import Image as PILImage, ImageChops, ImageOps
import io
import base64
import time
import copy
from site_apps.central.src.reports.photo_log_manager import PhotoLogMixin
from site_apps.central.src.reports.report_sheet_styler import ReportSheetStylerMixin
from site_apps.central.src.reports.ndt_report_runners import NDTReportRunnersMixin
from site_apps.central.src.reports.report_preview_view import ReportPreviewViewMixin
from site_apps.central.src.reports.report_table_controller import ReportTableControllerMixin




try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)
except:
    pass

# --- Versioning ---
APP_VERSION = "v260404.03"
from openpyxl.cell.cell import MergedCell
from openpyxl.worksheet.pagebreak import Break
from openpyxl.drawing.image import Image as XLImage
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.utils.cell import coordinate_from_string, coordinate_to_tuple, column_index_from_string
# Ignore library warnings
warnings.simplefilter("ignore")

# --- Constants & Paths ---
NAN_PATTERN = re.compile(r'^nan(\.0+)?$|^none$|^null$|^0\.0+|-0\.0+$', re.IGNORECASE)
DOT_ZERO_PATTERN = re.compile(r'\.0$')


def find_exact_column(df, keywords):
    """Return a column only when its header exactly matches a keyword.

    PT source files often contain headers such as ``Spool No`` and
    ``Request No`` even when they do not have an independent sequence-number
    column.  A substring match for ``NO`` therefore maps unrelated identifiers
    into the preview's No field.
    """
    normalized_keywords = {
        re.sub(r'\s+', '', str(keyword)).upper() for keyword in keywords
    }
    for column in df.columns:
        normalized_column = re.sub(r'\s+', '', str(column)).upper()
        for keyword in normalized_keywords:
            # Multi-row Excel headers can repeat the same label when the
            # header rows are combined (for example,
            # "Request Date Request Date").  Treat only repetitions of the
            # complete keyword as exact; unrelated names such as Spool No
            # must still not match NO.
            if normalized_column == keyword:
                return column
            if keyword and len(normalized_column) > len(keyword):
                repeat_count, remainder = divmod(len(normalized_column), len(keyword))
                if remainder == 0 and repeat_count > 1 and normalized_column == keyword * repeat_count:
                    return column
    return None


def format_thickness_range(items):
    """Format the minimum and maximum numeric PT thickness values."""
    thicknesses = []
    for item in items:
        raw_value = item.get("Thk.", "") if isinstance(item, dict) else ""
        if isinstance(raw_value, (int, float)) and not pd.isna(raw_value):
            thicknesses.append(float(raw_value))
            continue
        match = re.fullmatch(
            r'\s*(\d+(?:\.\d+)?)\s*(?:mm|t)?\s*',
            str(raw_value),
            re.IGNORECASE,
        )
        if match:
            thicknesses.append(float(match.group(1)))

    if not thicknesses:
        return ""

    def _display(value):
        return f"{value:.10f}".rstrip("0").rstrip(".")

    return f"{_display(min(thicknesses))} ~ {_display(max(thicknesses))}mm"


def first_report_date(items):
    """Return the first valid extracted report date as a date value."""
    for item in items:
        raw_value = item.get("Date", "") if isinstance(item, dict) else ""
        if isinstance(raw_value, pd.Timestamp):
            return raw_value.to_pydatetime()
        if isinstance(raw_value, datetime.datetime):
            return raw_value
        if isinstance(raw_value, datetime.date):
            return datetime.datetime.combine(raw_value, datetime.time.min)

        text = str(raw_value).strip()
        for date_format in ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d", "%Y%m%d"):
            try:
                return datetime.datetime.strptime(text[:10], date_format)
            except ValueError:
                continue
    return None


def remove_unused_report_sheets(workbook, used_sheets):
    """Remove template sheets that were not used by the generated report."""
    used_ids = {id(sheet) for sheet in used_sheets}
    for sheet in list(workbook.worksheets):
        if id(sheet) not in used_ids:
            workbook.remove(sheet)

# [PT] SCH -> 두께(mm) 변환 테이블
SCH_TO_THK = {
    "1/2": {"5S": 1.65, "10S": 2.11, "40": 2.77, "80": 3.73, "160": 4.75, "XXS": 7.47},
    "3/4": {"5S": 1.65, "10S": 2.11, "40": 2.87, "80": 3.91, "160": 5.56, "XXS": 7.82},
    "1": {"5S": 1.65, "10S": 2.77, "40": 3.38, "80": 4.55, "160": 6.35, "XXS": 9.09},
    "1-1/4": {"5S": 1.65, "10S": 2.77, "40": 3.56, "80": 4.85, "160": 6.35, "XXS": 9.70},
    "1-1/2": {"5S": 1.65, "10S": 2.77, "40": 3.68, "80": 5.08, "160": 7.14, "XXS": 10.16},
    "2": {"5S": 1.65, "10S": 2.77, "40": 3.91, "80": 5.54, "160": 8.74, "XXS": 11.07},
    "2-1/2": {"5S": 2.11, "10S": 3.05, "40": 5.16, "80": 7.01, "160": 9.53, "XXS": 14.02},
    "3": {"5S": 2.11, "10S": 3.05, "40": 5.49, "80": 7.62, "160": 11.13, "XXS": 15.24},
    "4": {"5S": 2.11, "10S": 3.05, "40": 6.02, "80": 8.56, "160": 13.49, "XXS": 17.12},
    "5": {"5S": 2.77, "10S": 3.40, "40": 6.55, "80": 9.53, "120": 12.70, "160": 15.88},
    "6": {"5S": 2.77, "10S": 3.40, "40": 7.11, "80": 10.97, "120": 14.27, "160": 18.26, "XXS": 21.95},
    "8": {"5S": 2.77, "10S": 3.76, "20": 6.35, "30": 7.04, "40": 8.18, "60": 10.31, "80": 12.70, "100": 15.09, "120": 18.26, "140": 20.62, "160": 23.01, "XXS": 22.23},
    "10": {"5S": 3.40, "10S": 4.19, "20": 6.35, "30": 7.80, "40": 9.27, "60": 12.70, "80": 15.09, "100": 18.26, "120": 21.44, "140": 25.40, "160": 28.58},
    "12": {"5S": 3.96, "10S": 4.57, "20": 6.35, "30": 8.38, "40": 10.31, "60": 14.27, "80": 17.48, "100": 21.44, "120": 25.40, "140": 28.58, "160": 33.32},
    "14": {"5S": 3.96, "10S": 4.78, "10": 6.35, "20": 7.92, "30": 9.53, "40": 11.13, "60": 15.09, "80": 19.05, "100": 23.83, "120": 27.79, "140": 31.75, "160": 35.71},
    "16": {"5S": 4.19, "10S": 4.78, "10": 6.35, "20": 7.92, "30": 9.53, "40": 12.70, "60": 16.66, "80": 21.44, "100": 26.19, "120": 30.96, "140": 36.53, "160": 40.49},
    "18": {"5S": 4.19, "10S": 4.78, "10": 6.35, "20": 7.92, "30": 11.13, "40": 14.27, "60": 19.05, "80": 23.83, "100": 29.36, "120": 34.93, "140": 39.67, "160": 45.24},
    "20": {"5S": 4.78, "10S": 5.54, "10": 6.35, "20": 9.53, "30": 12.70, "40": 15.09, "60": 20.62, "80": 26.19, "100": 32.54, "120": 38.10, "140": 44.45, "160": 50.01},
    "24": {"5S": 5.54, "10S": 6.35, "10": 6.35, "20": 9.53, "30": 14.27, "40": 17.48, "60": 24.61, "80": 30.96, "100": 38.89, "120": 46.02, "140": 52.37, "160": 59.54},
}

def convert_sch_to_thk(size_val, thk_val):
    """SCH 값을 두께(mm)로 변환"""
    if pd.isna(thk_val) or str(thk_val).strip() == "": return ""
    thk_str = str(thk_val).strip().upper()
    try:
        val = float(thk_str.replace("MM", "").replace("T", "").strip())
        if 0 < val < 100: return f"{val:.2f}"
    except: pass
    sch_match = re.search(r'(?:SCH[.\s]?|S/)?(\d+S?|XXS|XS)', thk_str, re.IGNORECASE)
    if not sch_match: return thk_str
    sch = sch_match.group(1).upper()
    if sch.endswith('S') and sch not in ['5S', '10S', 'XXS', 'XS']: sch = sch[:-1]
    if pd.isna(size_val) or str(size_val).strip() == "": return thk_str
    size_str = str(size_val).strip().replace('"', '').replace("'", "")
    size_str = re.sub(r'\s+', '-', size_str)
    if size_str in SCH_TO_THK and sch in SCH_TO_THK[size_str]:
        return f"{SCH_TO_THK[size_str][sch]:.2f}"
    try:
        size_int = str(int(float(size_str)))
        if size_int in SCH_TO_THK and sch in SCH_TO_THK[size_int]:
            return f"{SCH_TO_THK[size_int][sch]:.2f}"
    except: pass
    return thk_str

# 공통 스타일 (테두리 등)
thin_side = Side(style='thin')
medium_side = Side(style='medium')
thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

def find_project_home(start_path):
    current = os.path.abspath(start_path)
    while True:
        if os.path.isdir(os.path.join(current, "config")) and os.path.isdir(os.path.join(current, "resources")):
            return current
        parent = os.path.dirname(current)
        if parent == current:
            return os.path.abspath(start_path)
        current = parent


if getattr(sys, 'frozen', False):
    # PyInstaller creates a temp folder and stores path in _MEIPASS
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
    SCRIPT_HOME = base_path
    BASE_DIR = base_path
    CONFIG_DIR = base_path
    RESOURCE_DIR = os.path.join(base_path, "resources")
else:
    SCRIPT_HOME = os.path.dirname(os.path.abspath(__file__))
    BASE_DIR = find_project_home(SCRIPT_HOME)
    CONFIG_DIR = os.path.join(BASE_DIR, "config")
    RESOURCE_DIR = os.path.join(BASE_DIR, "resources")

SETTINGS_FILE = os.path.join(CONFIG_DIR, "logo_settings_unified.json")
SRC_ROOT = str(Path(__file__).resolve().parents[4]) if 'Path' in globals() else os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", ".."))



# Embedded Logo Fallback
DEFAULT_SITCO_LOGO_B64 = "iVBORw0KGgoAAAANSUhEUgAAAE4AAAAxCAYAAAB03p8JAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAAEnQAABJ0Ad5mH3gAAB6QSURBVGhD7Xt3tFbFuf5WY4leRYogAtIRpIqIoNguGIwJGiMiV4I3ptgrosbEXDFGkxgjIJrYSwIIiALSpQkcyuEcTu+99/adr+06z28978x3RCL+1s0y/91hnXM2e8+eeeeZd96+LUhTAALzF4APwAPgKsBXsAG0AYgoBXgeEHGBuA/f8xGCj8BVCFwgxFHigGMDrRzNlZeAsEI0ABxOwXETPzEAcSBQARy4CDh2HGgH348AKgQ4LngJP4DylVw6JNMH3EDBU0CUcyXG5pxx3ldCg8MHfIFrYuNz/nTNHyCsXDQjjiAIZByOyT5KKQTERW52/ZFmfS1wClBeANd24fGvFyDqB4i6DoLAQeA5gO/D9V2E4AloQVwhzFEcDVwLr31ARX0oWwmNgVmc8vUzOBoFDz5suFCeB98BOkAg4/ARhht4uh+pVnoT4yTT02MQyCgUfC4yUAi46oiCG/XlGTeFDBBwz7k2Ay55QBFQ24MLDzHlwPUd+F7Al7r6OAigBHw9H0kR4AQq/mKvL6ETIiKBi7gKEHgBgogL5XqIBlG0Ic79ga+4NB8eiXX1xhJT19YcyklcpbmiE0oDKcToBbM/aXSg4JIlAp/MKUwK20EoCKMZnmZ8vq+UvE+GIXp+PEATAkQDFzHfRkx58G2PyCMgeOzsc1y9cZxTlhkANhSiHEt5gG8jcGOIqbh+x9asRdpkU3gRB3xPMymHOAY4A5gBzlcBOoMYwrDhydbaQDyOGGKoQRghWT45w4dDlrM9zaUkxuOxVbJbMSFQCRgqRvA013ARJELAlZPoyyJ5TCPkik4PnYGLVijE9IZrLpE1KOFiP+6jg6ICnhz1dhWF7doIhMt9+A530RdupBjpNFKIP9EgQER5CHPb/Ajg8IrAUfSQyxLAaYlFor8CnADFXxzNsLPcI6cFMeGEBJiwo/C9doTQiZiy5Wgpn4xOmWdDRWJy5Hw4cJQDFXgCath3ST/QTgQ8YX0VKH36+CMcS6HlweYlUQq5cAJPbvO0yjEzAEZIrM05ebQSh4fPXDh+FPEgBhVwo215kRvVocUpPG4kl+IRIBtRihpZIc9mDJ4bgeM5XbKMo/N94UCeenPfEpmZUAYGOLlHYaVsIBKDl1sJFNUBMZ41Tk+CeHRdwHURUAtwOygfZARNIhULG9k9wu0j24mgM5MkdiQhXoWlzDOug8dEhP2Xwl/LNAIXlU1AOADy6+GTxjAhMSD4Yfg+AdRcojlFwYePGJmBwtTzgbYQ/KwiIK8SCEWAwIarbA2W4Tpey1IMxwtwxPtYYehyeEpvDm6HUb10OVInzUfRrIVoWr0LkawStO9KglNZC0/5UB45y4bf1gZV0YCgvgZOcyX82loElS0ImkLoDHzUc18jNvzaOvh1TQgq2+HVheA0dsCva4Hb2AinqQl+bTuCqhC8qib4He1aIRIPbr2IgABh2RgH6HTRsnIP8m55HNmX/xT1b36KeHUDQgdT0J6XhRAoWpT0pnwUeSc8p/mps6AU9Y+/iqJxc1F8/X1ofHsD/MYWePLPaGLyhtIWAzW1VrnHAadZkWeHnAXUvfcJPug/BbtHzMKWITOxvN+VOHLNndg98WbUbkuSnfC4c8pH+rp12PvLJ1Gw4HcoXvhb1Dz8HNJ+/gzSXnwbsYgtE9uZ+Sh75FnkPbII+Q//EbvuX4Sdjz2PgsdeQP4TL2Dfgt/i0L3PouT+P6PsF8+h+vUPEbfjoPwWs8VX6FSeyFmuqmHlVrw7/D/x8cU3YM/g7+O9wVcj+Y4FOHLlPBT+9QORhR2ieLQ5w1NCjmJzSqqxddbPsOH0S5A5aBb29J+O5T0vRfZTLwEh7pRhXso2X5tXMeEsAxwllPChnDRqQG3weNkVSB46Czuv/xnie1PQtGordkyZhz3WRdhljUDZm59oAU8BpICdf1+BD8ffgG1njkPGGRcj7Ttjsb3vDOy6+4/wI/pMNr21BgdPH4ktp4/ChnMnY+1187Fy4g+x8/Qx2HfmGHw68UZsnvYT7DvrSqRYI5A26UeINDTqBcRoB/qIK8LhI55fhlUTZuGD6XPRsPULhNZ/gXUz5mGTNQaZ1kjUPfxH6n1tIsnxJ4PYtEGADgcpP3sGH1nDUTDnKXifHUBk5TbsnHoHVn53DBqWfWwkTYAgps8qZaTYsV8HHOkTuycAMhctw74BM+AeLJKO7O4k5SNj6M3YZw1C1VOvC8dFKQujLuxIDKFP9yO73wwUWX1RdOYEdL6wBnZ+O3yqdwClv1qCo9b52HnaSGTfuhDxzArULvkH0q2hyLcGovGVfyCeUo7yWQuQYQ1EXu/L0ZmbrzdfhHxAHStHI+fpZVhx/pVoTEqT56S5IykVB0bdjHxrMBpufRw+zQs+E4vEp7Umsrhl+Q6sPG00MmbcDZS2AF5clFXsUD62DJyO/UNuBDIrRWx5lBVRbTtGEkJPA0dBajhOkwDV1IlVV9+Kww/8Tm46cNAhkgJofHU1NlkDUfT9x7VhycGi5Dpt9dfMfw4Z1rkomzYXaIhqTSgWuY+SO3+DHdY52H/NbAS51TJePC0Dxd0moPassfAPHZV7HQfSsbP7BKSfNBRt2z6XDdW2MmFw4BXXYUfv65A37zcGMyoJfQTL/roCu63BSLvsdngR8omWU1QIUV60diL56jvxyXcuQmjrfmGIdsTEc5D3n39H1tf03Bvyf5FrdGJkeeT1Y4EzpoGeGmjdkYzlw65C2ZY9+mXXQbtP08RDUNmAozPvxeHBN8GraNIvxBQ8mgWUi8+9hUNWD+TMvlcWS3tJLL7qRqSPuwU7L5yMhk+3CehxeLCrSlExcBrKuo9FLDNLNF/ICeHQg09g/0kD0fSXd0RGc1mKx8yLofz5t5F05jRE1x0SjqIHQ7nEo+k1tyLj+w9hfe8rEU7X3EoCwvBEG9et2Y71Jw9H2s0PAPEYOmCjFjZaFE2QAHZODfZeOB0HRs1AuLRU1ATXxzFk844HLqF22fJfeBPbRt0At6lNa4yIAztqw49SqgF2Ui72DJyJug1fGOCAsM+jAFS98h52Wucge/4jIhdo0IqI+iId+86ehPqX3tVq3fZB2Dsry1E04AoU9hiPSF6hANRGAJpqkDZuJgpm3i1OLtUB2d+prsTeodcj9+r7gLCHODW7o6BcH7ZHwweIHinAZwOuRdXSNV3AUbt6sTCSZtyJ7dYQdKz6XB41w0az8uA6NHE84c66Xy3DJqsn0ha9BM93u0who0OlaeDEmdW2Dlvxo3/C0bE/AtojCHwPKk7iAgQ2da4PxF3snf0Qkl99R/rT3As5PNBA5WvLcdA6G4XzHxHQOoytVvHuxzg44Dr4ueUyl2/7aAYQLqtAwflXIq/nZQhnFsnmhV2Rtqh95s84OvQmBM2dAijdpNDOA9h6yigULPizNvloOzoKdkAPxtE2ZaBw8KdPI/We3+vJ6ZRwQzILsLf75TjS63J4RzPE8CUnxvxA3DclLgvQ+enn2Gv1wp7JN8Fu6xBcxESlrDQYaeDk7Grjkq3y/heQNfQH8GrbhYtcX/uZvk/jUTNrwYpPkPbJeu2SxAiqHrJ26UqUWj1QfduDerFaSePgi0twZM6jot3IiTx1BMkuq0Ne3+uQ3n0Swqll0ld1cAt8RFKO4vBl8xDJL9eAciNe/hBbrSGoXv2ZbJQfcY0vrP1iRUeZp2bVRhz9g5ZTidWW/309DlvDkHvpLXDba7VstBnZUYjbLpyorT2T9HTknj0Kqd0uQTQtXwIWIdGOX45lKRG42tbR4hEouXsR0s+ZBjeXpx/ia0pIRZxlX/xKv74D8eZ2ceZtiUhohKqWrUCx1R0Ncx6TMRM7VLRlB6q3JMn/O3hUyRikpbQOWedfjYweExFNq9CdaTVwK2MR1L23GdGKBhmLLlXtL5/GHmsYGvcekk2lUy/TG0FON5AjO/XNCBeWa4VnrP3qp15GmtUPVbc9DiVmDR1fE6DwGKzwBIN4fR2yh12DTGsQQis3yobRDv0KcFTPXLIWfLpVPPR7ZFhjEdmUYgSztqJJoJgEDKyRUtfE6WSxmrqy15cj3eqG8jsWdIHGIyGxNi8R3vnSd3HK65FzwVXI6j4WkYxS3Z+dfC5LiS1upAmU14HimXNx+LTRaD+SrY+qXr7MZfD5ylVCbtP4Lb3pXhyxzkPloy93PT/uRbl0olGkXTEXmdYFqHlhqXQRptI8Js2iv5YALiH4qn67FEnWQJQtfE06erTVfMqQABEE2mEnRSZ4KbLRuCIE7qh1DkrmPdZFj3gjpml3lEaMvkfgci+4CtndxyCSWSz3aHjKRhnPVwjklK31SB8/A7m9psDOLpXxKfdO1ISjzXOvtQVFl/8Yqdb5KPvd28d3/UpTQYDSObQl+yDn3qdka8SVPx44I/e6gCt9fQU+twYgZchMxPal65u+DV/CMArtjNgaphEHmAMYw7Bi2dcDR2ISjPa/AU5oMgZmqKQY+/pcgqJhMxFU0PslMF9uyvEtwZFs4ZJSpA/m8RuIyteMtv2G1vjQ88iweuPQrLvgByZaYjaQrQs4Dm+YBnWb92DPWRORYg1B7vfvgn0ky3RnvC0QLksEKmUsWaR+ufrV5Uj7toEzBIcPp+DAacORM/5mBPWtcjsx74kaw99sLek5ONxjEnKsEWj4x+bju/1Ta3r6JWRbfXH4ytvgdzJBcIzA1sCJ7S+LItPwOlZYhh0jZyDP6o8cayiyrrkd5cvXAW3UXdpda2GoSNhaI0gLigutW0IZdw6Kv0XgjBeIyKZdOGoNQtplt8BvapfFnJjfdEsA13ggDcmnj0XOKSPRtn738d3+qTU+txgFVn9kjZsFv7Ze1iJDmUVZWgjrXdXREdppMSTfdg+yrPNQbw0V33TzoCtQ8thLCIzJQC+QikLCa5In0OGb+sXLkfEtA0ePkH9DKzYiy+qPlCm3wGvp0HLHzPG1jQ+NCGnadgCpJ41EzhmjEdpx4Pie/9Qa/vwmiq3+KBz6PbjF5TKUzGd2yiKBXcB5DGtrgdL64adIP2MYKqxByD95FFJOHY1kaxSyrr0LrXsOSh/aQRLgc78ErnHxcmR+y8BRhHDsznfWItvqjdSpt8JrNcB9E3KyWm2jNq/dgQzrIuSeOR5thv4TNfavXvy+BCtK+k+Hm6sVUYL72SzekWv2JrNJzskHmtuRP+d+HLYGoMQajcKTR6PMGo29Vj/smzILsYwSGSBOl8RTYoBSUTDC8a8BNxaRTD0mcwRfB1xo2QrkWj1xlMC1dcgzMV1O1OiDu2LmomX1NmRZI5DXbSLaD6d2yfWvaxy3YvF7KLL6oKDXNMTStM/L9SU26p+Ao0+n7WeFaFouCmbegyPWRThijUDpaWNQbg3GIasvCuc/DTSHJMPkezqOTwIb/h3AGUXU/vIHyLd6ImPKrfBbmaYhB5xo+QY42o7kuDXbdfiq5yREMvO+ETjOV730A5RYFyDn3CkIGeX4FeAMjRo9CZ0EksCg+cGmUktRfvdz2NxrAvZb/VB/8nBUWkOQes6liG4/pF8zkRWOWffKP/5FGXdi4GjuyHF7+X3kWT2RN+nHEuI2nU+MAB+ZvEfTmm1ItYaisPuliB7N0aAf3/+YVreEwPVD3vnXIJZVIPcS5hqbANd1dg1wzJXSufcTFDe2o/ztlUiffDPSrAtRaA1CujUYJc8uk8kNs0qrfYWByW8XuIRcqX3lfeRYvVAy+gfwqo2mozV+AkFHjeqLtQ40frZbTk7+WWMQ3pvyjRzH1vCnt1BsXYCsQdPhFGqFaMiQJsDJwTwGOKbfGAlmqCjhhvEte38qsm/8KQ6eNAQZ1oU4PPMu+FFtuycGrVn87wOuetnfkWX1QWn/6+AWVWrgmKn6BuA8ZuLIQTuTcOCUkcg7dRQ6N+sA5te/pVvToiUotPoi5aLr4VXW6JvHIGeRKApe4WhPJ5Fpk1F0tEvSNtC7Knk6BS+rCBlT/wtp1gAcvvhGRIr0oIkxqyUU/q8YwCcGTi64+PfWIMfqj9JukxHNMEFKcxS/rnEWSXTz3aQU7PvuaJScNBbhT/bI7Mdy0PGtZeFLKLD64cDEGxE0MgCmj1VCpApwXdlqAqd00JChYkaKyE+UEyx/cMVZA1o+2oHk00Zg17kT0bw3U+79W4EzwqV57SbkW0NQ+p3xaN13RN/8htXT3ZLQN2nOysPhPpNRYl2M0PLt/1/gGn7xLPKtfkiafjuCCOMjmo7EOxajDxKo4/gm4CduDNMIDn1SJVqTwcxIEOgoQdhB1tRbsMMagtZ1h2UgwxSoWkqXqxtKj4mOcLwu4BL/N7ASuPzjgeMzA5yICjN4y9ZdYvmXWhehadNOM7jZHjO28BcvjHgm/WyxqhpkjfoesqwL0fi3TxOv6F/GskhsNGktm/0osq0+SJ59r+RLpH0FOBNQlLAwaz8SlTGsHPI9dEptCNPhAbwwM9x6yXWPLRJftvOdHTJQQhZWLf4IKdZ5KJtrgOO4ZhFsvCccbm54ZXUoPf9KZPe6BOGccn3PaHqOKa6cUYFth5KR0m00yq0LUbvsQz0Aq4j411dS9iV5VE7IBAurKgxwfjiCnKvuQJrVB5UvvquBYsLUYfmYj2jAWBxtPgUvHkfW9T+RsFL5g39IYCt5XWa+2CxWNMhJ8E0hihPXVS8efQhbwkixwIeKB4xpS4SESY2OV99EtjUS0WWbZKBE3Kt88Soctvqg+vbHNVg0qk2lFhuZh8Zy3Nzwy2pQff40HO15KdqztPYyxU/Sl8DZ1O4K6CgtQdKQy1Fs9UDB/c/KPVZeuNxYn2JGoRkB4o6LoCaG3A170dzKVLJOe+bd+jgKrIGoW8Ccqy4rY2qTFkRYMX8h6Xm4ra1IufR7SLYGoOkPKzWdHCOgjat5zpJ6NoLo6NhaS2MTGtdsB1paJH5P+eZJIoTFgVQeOnQYWrUKKaeMQ/itrTJQApii11Yi2eqNhrlP6OgJw9om7pk4SpQYLCRic8srUdL3GuzrPQmtOcblYqjeyEPKO1ZL8d1oaxOSp/4Q+da5yJk+H25LRDZMyszkoCjEWfKlPHSkFWPbT55CW01D18Irn16GImsYGm57BEpFtDLm2pjdYm6FtTEcr7QM2YOmIemki9G8SiekSA85LhHGsrhZAhwTLnTSK6qw8YY70bIjSa8sposnuHBaTqY4AG2bN2LTmWPRtmqnLErvFVD4BmVcH9TOeUwix+QchsolSkzGMTE8qUgSGVeBzAuuwBd9L0VbAjjKVA4n2oHKKTAAxZA5525kW+fhyMCrESqpNJxg7GCuzuGBA0pXbcCBmx6Ciuilko761VuRag1C2aS5cKMsidTynbV7PmvTOjVw0eQMZHx3LHadNxXRDJMHYYRGjmoCOBLEX5JCA0KtrVg77Tak3fsbCTd7nou4r0sJOr+sAED96rVYed4ktO1ntog7olVf2ZsrkGb1QvnsB2QnSZxE2dlJp6pko5iNY3OrqpDbbwpSe09A2PiEBCqRJGdnEkzZxVcK/udlpFoDkXraBNSt3qLHIKcJDSSQafc4Dt73K6Q9/Rd5zgwY1xbNL0Va72uR2/MaeHmlwgR1CNBOLS/5Cr248PJt2Gf1w/4Z8+GHKKxMOQ1PXwI42TFe0cuSMFGA7IV/wvpel6Bi5Xrp1AkXnczGd9kHQPoTL2Dj+B/Ca9ABxYBahsdh6QdIs3qi/PaHJYHNkIHYiexEBOM6P5Io0HMrylHR9woU9hyPOBPIUpLK0rFEtZA2h3R5KxDel4ajva+V+pBDU2cjnKvdIdulWNeraXh/HT67Yhba0jL0M0WOFe2HukeWYA+DmUtWyL7Us76EC5M65ECsh/zbnsLuU4ai6e8bNI0cxACna+kMcNq2ovmhnfva1RuxzRqAIyNvhJtTLp5EhFVJUSZtFdAaRfK4W5Az79fGDCA8Griy376KFKsnKuc8bHSUrsUVcWDqfxk9NsUJiGRmorLbRJT0mIBwTq7Mz9pIHh8hWAgLRHDb5Go7QM7cX+OoNRxHrAHI/+8nELTYXXZCkFWNXSNuwJ5593UJEG0OaZvGSS/CtvOm4ciU+bIOaknfY00Ja46BeHIRtp19GbIm/BhoaEWMKlIsDSPoElpVCz3jdlH9wUco/SjKekxFLuPz978grC+7yUy3D7R/vBcHTpuE9jc+k0FUoMuuuAOl83+NFKs7qv7zTl2oZ0ATk8fsXqI6kq1xw+coPHk4Ss8Yg6Ztu3R/criRp1q7BvBZ+elp3V27eicOnjoedfSZzxiHogf/CHUgG/b2Q8ie+QA+sS5G3dqEtmctMCO0DJYxK+8h957fY6s1Fp2r92oiPFNYqxRKFi7GNmsYWn//PpSiORaX+KBARg4zroMGLlHbaouTAru+BqWjbpSY+4Ezx6FlyQogytqQAKo+gqTp92DP4BsRlNWxLhy+G0e8pgahtZ+j9MIbUGkNQFHPqWh7dQ3c9CopS/UcH56vM2q+q+AVNiG2MxPZcxYgyxqMAmsYin/0KOxtqfBKa40g1qYLgw6KWs+TOgK4je3IvGq++JJV1nCknnEJMibNRua42dhqDUfStXchaNXxOobJ5NjTmA9sWaedlI0v/mMq8q7+BVDQaGwCG9HPD2F398lIG3YTVF6VKIIYudz482ITJoATE4nltKLNeLg8+LEIMq+dj31njURyv6nYOOgqZDzzF9Ss2oCkX/4Gq3pMRu2ftIxoZgLO9ZD0xvt4Z/z3kGJNQPnJl+DASWPx8dDp2PXLZ6EaYl3uHI1Usl7BX9dhzeibsKb7JCSdMRYHzxiDPWddhnVjb0L2e+vMETXZewJHK1hyG5oLQys24+i545BjXShxtv3WUOyxBmHX2VPR9pHmXFEIvCDaHt0vD22kOh6gceEb2PidUUj98YOIHjqAqq0bsWXiD7Dj1IvRsWyNmdvVnwt4+lMEKSw0vrElWpJFQEZgx7iqeAz5C19E8f+8iMjeJGy/7Wd4q9c4bBl+Pd7teyn2/nQhUNUpO8DBKfnTV6/H+l88gbKfL0LjPc+j/L7nsfvnT2LvoqVQIRG7aFcuQobtSz7Zje13LED6A4tQ9vDzyL//GeTc/Qw2/PejyNxs3CnjAjIPIsDZGkiW6KMzitpnXsGn543E5/8xAvtPvRhfdJuEsmdeE0OdfflRipT+My0s3gv9cGNQVLXhyMNP44O+E7Buyg+wdvINWHnuWBQ/+SIQDgmntnBmSXjoCDcrGgQoBVjysYX5+IGl8XSxEHMQTsmEV0+7ykNnaSnKnv0bam95EjXPv4VocZkUhbB42mV1ue/BbQ3BrmpCUN0EVVMPVdMAt7oB8bpWqf5mFXp7oIW+FwTwmjvgVjZC1TYjqKqDW10Nr6YO0bo62FHjVIvI1TXJ4i+S41hKJsF9D0FjE8pXr0HRg0+i8efPoO31dfBapFhBuIBBCjo8BJE6jRzPinfXo+TzYLfWovpvK1A47wkU3vE4ml77CF5LA+JwUAsXbYTYWBIETjguARyrygQ4T6f8wiSQFrSUwXcCvpTIAC1RoKAOaGANkoeAuQYhRJfsa2NH/zHVEbqJhc3SfltK6/2AH5bwh7rOqEJp3EHyhE4zykcZCQOYwFHAS9iLWjmAIxFq1l7FoWqrgeJqIKKNX9f3ZVrJ+ZBlHS2KRO+TGV0XUe3XANE4kFMNFFYDNs+cg7iKodGkPwVxqXYPxKyRI8CjKuYd120+73FsD37YQdRl2saB70cFpMTS+Ca1KJW32EZSlMx72oCMB0qOUlgEqY9AykTpfZB6lvnrL3I4uoSpjDvF4kXPicFWNH04njFfWLsnBoUu8iawMc9BPBqRKAaLgEzwRA4h3TOaLTRfJI7IccgHiZoVU6XfSbst4Cq0pcBxWxBGOAgjoNPf5RhobhBjXtwHAxzrw0iQDMxdigVwnECIJ3fQzklUyzPqwEiCChzY8rWMQkAXwGV9GmtKAvk8gt6LMKyjEFcOIiomRq145PxyhR9+0MjlCqLUlto7oMfACA2ddCFQDgMrpPgePWSanz5UzEYQcqDiClFa86xvozx0FHyHEtGFreKwXV3bKzvDiiTSJ/MQfFoDOmjLYK3j8XMrftCiv8oRY59rkKk5a6C1cwI4OaoEjnTyV1yfZZF9xo/VGsWk4uSHXKOVScBdZWGf0mEdkcYURAngAkYetIzjkWOFI4GTbynoh8VMRRJ33fHki0QuKHH05ZMmA5z+Bobv8IMHT3+WRE43tihlGU0WrjYWhGEHtjbuDXDsS5CZMpQTwg8TGXXh51NGi7NWTj6OS4DG79hoCiUI0sfOfFmT+L8WMyKfzFGW7tos/mpLeETyjo766DHYWXYhIfO4YNpS2sPXJQlmdvNu4mXxUc2jxF+xJMwEHEuigaJqv9pXLrtEJrnDM3N++UymOW7OBGMlXu3i0MRPIkyTaObSfHb5f+1/2/4fWo4AjuPpKacAAAAASUVORK5CYII="

class PMIReportApp(PhotoLogMixin, ReportSheetStylerMixin, NDTReportRunnersMixin, ReportPreviewViewMixin, ReportTableControllerMixin):
    def __init__(self, root, central_app=None, is_embedded=False):
        self._is_ready = False
        self.root = root
        self.central_app = central_app
        self.is_embedded = is_embedded
        if not is_embedded and hasattr(self.root, "title"):
            try:
                self.root.title(f"SITCO 통합 성적서 자동 생성기 ({APP_VERSION})")
                self.root.geometry("950x850")
                self.root.configure(background="#f9fafb")
            except Exception:
                pass
        
        # 1. Initialize Configuration first (needed for state variables)
        self.config = {
            # UI State (Persistence)
            'PMI_SASH_RATIO': 0.5, 'RT_SASH_RATIO': 0.5, 'PT_SASH_RATIO': 0.5, 'PAUT_SASH_RATIO': 0.5,
            'PMI_COL_WIDTHS': {}, 'RT_COL_WIDTHS': {}, 'PT_COL_WIDTHS': {}, 'PAUT_COL_WIDTHS': {},
            # 갑지 (Cover)
            'SEOUL_COVER_ANCHOR': "E5", 'SEOUL_COVER_W': 200.0, 'SEOUL_COVER_H': 18.0, 'SEOUL_COVER_X': 30.0, 'SEOUL_COVER_Y': 15.0,
            'SITCO_COVER_ANCHOR': "A6", 'SITCO_COVER_W': 80.0, 'SITCO_COVER_H': 40.0, 'SITCO_COVER_X': 10.0, 'SITCO_COVER_Y': 10.0,
            'FOOTER_COVER_ANCHOR': "P50", 'FOOTER_COVER_W': 80.0, 'FOOTER_COVER_H': 20.0, 'FOOTER_COVER_X': 0.0, 'FOOTER_COVER_Y': 10.0,
            'FOOTER_PT_COVER_ANCHOR': "A50", 'FOOTER_PT_COVER_W': 100.0, 'FOOTER_PT_COVER_H': 25.0, 'FOOTER_PT_COVER_X': 5.0, 'FOOTER_PT_COVER_Y': 10.0,
            # 을지 (Data)
            'SEOUL_DATA_ANCHOR': "F5", 'SEOUL_DATA_W': 200.0, 'SEOUL_DATA_H': 18.0, 'SEOUL_DATA_X': 30.0, 'SEOUL_DATA_Y': 15.0,
            'SITCO_DATA_ANCHOR': "A6", 'SITCO_DATA_W': 100.0, 'SITCO_DATA_H': 50.0, 'SITCO_DATA_X': 20.0, 'SITCO_DATA_Y': 5.0,
            'FOOTER_DATA_ANCHOR': "M46", 'FOOTER_DATA_W': 100.0, 'FOOTER_DATA_H': 15.0, 'FOOTER_DATA_X': 25.0, 'FOOTER_DATA_Y': 10.0,
            'FOOTER_PT_DATA_ANCHOR': "A46", 'FOOTER_PT_DATA_W': 100.0, 'FOOTER_PT_DATA_H': 25.0, 'FOOTER_PT_DATA_X': 3.0, 'FOOTER_PT_DATA_Y': 7.0,
            'MARGIN_DATA_TOP': 0.2, 'MARGIN_DATA_BOTTOM': 0.0, 'MARGIN_DATA_LEFT': 0.5, 'MARGIN_DATA_RIGHT': 0.3, 'PRINT_SCALE_DATA': 95,
            'START_ROW': 19, 'DATA_END_ROW': 45, 'PRINT_END_ROW': 47,
            # [NEW] 선택적 행/열 조절용 설정
            'CUSTOM_ROWS_COVER': '23-38', 'CUSTOM_ROW_HEIGHT_COVER': 21.0,
            'CUSTOM_COLS_COVER': '', 'CUSTOM_COL_WIDTH_COVER': 10.0,
            'CUSTOM_ROWS_DATA': '', 'CUSTOM_ROW_HEIGHT_DATA': 20.55,
            'CUSTOM_COLS_DATA': '', 'CUSTOM_COL_WIDTH_DATA': 10.0,
            
            # PAUT Specific (B31.1)
            'PAUT_START_ROW': 11, 'PAUT_DATA_END_ROW': 40, 'PAUT_PRINT_END_ROW': 45,
            
            # RT Specific (Radiographic Testing)
            'RT_START_ROW': 11, 'RT_DATA_END_ROW': 45, 'RT_PRINT_END_ROW': 47,
            
            # Photo Log Sash
            'PHOTO_SASH_RATIO': 0.45
        }

        # 2. State Variables
        self.setting_vars = {} # [CRITICAL FIX] Container for all dynamic UI variables
        self.logo_folder_path = tk.StringVar(value="")
        self.target_file_path = tk.StringVar(value=self.config.get('PMI_TARGET_PATH', ""))
        pmi_tmpl = self.config.get('PMI_TEMPLATE_PATH', "")
        if pmi_tmpl and not os.path.exists(pmi_tmpl):
            for candidate in ("home/data/SIT-PMI-001.xlsx", "home/data/SIT-PMI-K3 롯데 템플릿보고서.xlsm"):
                full_c = os.path.abspath(candidate)
                if os.path.exists(full_c):
                    pmi_tmpl = full_c
                    break
        self.template_file_path = tk.StringVar(value=pmi_tmpl)
        self.sequence_filter = tk.StringVar()
        self.rt_extract_keyword = tk.StringVar()  # [NEW] 추출 키워드 필터
        self.extraction_mode = tk.StringVar(value="전체")
        self.auto_verify = tk.BooleanVar(value=True)
        self.pmi_layout_mode = tk.StringVar(value=self.config.get('PMI_LAYOUT_MODE', 'STANDARD'))
        self.pmi_pane_ratio = self.config.get('PMI_SASH_RATIO', 0.5)
        self.current_mode = "PMI" # [NEW] Track active report mode
        
        self.show_selected_only = tk.BooleanVar(value=False)
        self.extracted_data = [] 
        self.pmi_sort_col = "" 
        self.pmi_sort_rev = False 
        self.rt_sort_col = ""
        self.rt_sort_rev = False
        self.pt_sort_col = ""
        self.mt_sort_col = ""
        self.pt_sort_rev = False
        self.mt_sort_rev = False
        self.paut_sort_col = ""
        self.paut_sort_rev = False
        
        # [NEW] PMI Preview Search & Filter
        self.pmi_search_loc = tk.StringVar()
        self.pmi_show_deficiency_only = tk.BooleanVar(value=False)
        
        # [NEW] 원소 함량 필터링용 상태 변수 (PMI 전용)
        self.element_filters = [] # list of dict: {'key': StringVar, 'min': StringVar, 'max': StringVar}
        # 기본 필터 추가 (Cr, Ni, Mo)
        for k in ["Cr", "Ni", "Mo"]:
            self.element_filters.append({
                'key': tk.StringVar(value=k), 
                'min': tk.StringVar(value=''), 
                'max': tk.StringVar(value='')
            })
            
        # --- PAUT State Variables ---
        self.paut_target_file_path = tk.StringVar(value=self.config.get('PAUT_TARGET_PATH', ""))
        
        default_paut_path = r"C:\Users\jjch2\Desktop\지역난방 PAUT.xlsx"
        saved_paut_path = self.config.get('PAUT_TEMPLATE_PATH', "")
        if not saved_paut_path and os.path.exists(default_paut_path):
            saved_paut_path = default_paut_path
            
        self.paut_template_file_path = tk.StringVar(value=saved_paut_path)
        self.paut_manual_vars = {
            't': tk.StringVar(), 'h': tk.StringVar(), 'l': tk.StringVar(), 'd': tk.StringVar(),
            'nature': tk.StringVar(value=""), 'loc': tk.StringVar(value="-"),
            'db': tk.StringVar(value="6"), 'peak': tk.StringVar(value="80"), 
            'z1': tk.StringVar(), 'z2': tk.StringVar(), 
            'L1': tk.StringVar(), 'L2': tk.StringVar(),
            'D_Upper': tk.StringVar(), 'D_Lower': tk.StringVar(),
            'target_fsh': tk.StringVar(value="40.0%")
        }
        self.paut_eval_mode = tk.StringVar(value="ECA") # [NEW] Standard / ECA
        self.paut_manual_vars['peak'].trace_add("write", lambda *args: self._update_paut_target_fsh())
        self.paut_manual_vars['db'].trace_add("write", lambda *args: self._update_paut_target_fsh())
        self.paut_extracted_data = []
        
        self.rt_target_file_path = tk.StringVar(value=self.config.get('RT_TARGET_PATH', ""))
        self.rt_template_file_path = tk.StringVar(value=self.config.get('RT_TEMPLATE_PATH', ""))
        self.kogas_target_file_path = tk.StringVar(value=self.config.get('KOGAS_TARGET_PATH', ""))
        self.kogas_template_file_path = tk.StringVar(value=self.config.get('KOGAS_TEMPLATE_PATH', ""))
        self.rt_kogas_mode = tk.BooleanVar(value=False) 
        self.rt_kogas_mode.trace_add("write", lambda *a: self._on_rt_kogas_mode_change())
        self.extracted_data = []
        self.rt_extracted_data = []
        self.kogas_extracted_data = []
        
        # --- Column Keys Initialization ---
        self.column_keys = ["selected", "No", "Date", "Report No", "Dwg", "Joint", "Loc", "Ni", "Cr", "Mo", "Result"]
        self.rt_column_keys = ["selected", "No", "Date", "Report No", "Sec", "Dwg", "Joint", "Loc", "T", "Mat", "Size", "Deg", "Acc", "Rej", "D1", "D2", "D3", "D4", "D5", "D6", "D7", "D8", "D9", "D10", "D11", "D12", "D13", "D14", "D15", "Result", "Welder", "Remarks"]
        self.kogas_column_keys = list(self.rt_column_keys)
        self.pt_column_keys = ["selected", "No", "Date", "Report No", "Dwg", "Joint", "Loc", "T", "Mat", "Deg", "Result", "Welder", "Remarks"]
        self.mt_column_keys = ["selected", "No", "Date", "Report No", "Dwg", "Joint", "Loc", "T", "Mat", "Deg", "Result", "Welder", "Remarks"]
        self.paut_column_keys = ["selected", "No", "Date", "Report No", "ISO", "Joint", "Size", "Loc", "T", "Acc", "Rej", "Mat", "Grade", "Nature", "Type", "a/l", "a/t", "Evaluation", "Remarks"]
        
        # --- PT State Variables ---
        self.pt_target_file_path = tk.StringVar(value=self.config.get('PT_TARGET_PATH', ""))
        self.pt_template_file_path = tk.StringVar(value=self.config.get('PT_TEMPLATE_PATH', ""))
        self.pt_extracted_data = []
        self.item_idx_map = []      # [NEW] Missing for PMI
        self.rt_item_idx_map = []   # [NEW] Missing for RT
        self.pt_item_idx_map = []
        
        self.mt_target_file_path = tk.StringVar(value=self.config.get('MT_TARGET_PATH', ""))
        self.mt_template_file_path = tk.StringVar(value=self.config.get('MT_TEMPLATE_PATH', ""))
        self.mt_extracted_data = []
        self.mt_item_idx_map = []
        
        self.paut_item_idx_map = [] 
        self.kogas_item_idx_map = []
        
        # [REFINED] Column Keys Mapping (Must match Treeview column count and order)
        self.column_keys = ["_status", "selected", "No", "Date", "Report No", "Dwg", "Joint", "Loc", "Ni", "Cr", "Mo", "Grade"]
        self.rt_column_keys = ["selected", "No", "Date", "Report No", "Sec", "Dwg", "Joint", "Loc", "T", "Mat", "Size", "Acc", "Rej", "Deg", "D1", "D2", "D3", "D4", "D5", "D6", "D7", "D8", "D9", "D10", "D11", "D12", "D13", "D14", "D15", "Welder", "Remarks"]
        self.kogas_column_keys = list(self.rt_column_keys)
        self.pt_column_keys = ["selected", "No", "Date", "Report No", "Dwg", "Joint", "NPS", "Thk.", "Material", "TestItem", "Result", "Welder", "Remarks"]
        self.mt_column_keys = ["selected", "No", "Date", "Report No", "Dwg", "Joint", "Material", "TestItem", "Result", "Welder", "Remarks"]
        self.paut_column_keys = ["selected", "No", "Date", "Report No", "Line No.", "Joint No.", "Size", "Th'k(mm)", "Acc", "Rej", "Start", "End", "Length(mm)", "Upper", "Lower", "Height(mm)", "TCG MAX(%)", "Type of Flaw", "a/l", "a/t", "Welder", "Tested Length", "Evaluation", "Remarks"]
        
        self.date_listbox = None
    
        # [NEW] Handle Application Closing for Final State Capture
        if not self.is_embedded and hasattr(self.root, "protocol"):
            try: self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
            except Exception: pass
        self.rt_date_listbox = None
        self.mt_date_listbox = None
        self.paut_date_listbox = None
        self.paut_line_listbox = None
        self.paut_size_listbox = None
        self.kogas_date_listbox = None

        # [NEW] Add traces for Template-Linked Config Auto-Load
        self.rt_template_file_path.trace_add("write", lambda *args: self.load_template_specific_config(self.rt_template_file_path.get(), "RT"))
        self.kogas_template_file_path.trace_add("write", lambda *args: self.load_template_specific_config(self.kogas_template_file_path.get(), "KOGAS"))
        self.pt_template_file_path.trace_add("write", lambda *args: self.load_template_specific_config(self.pt_template_file_path.get(), "PT"))
        self.paut_template_file_path.trace_add("write", lambda *args: self.load_template_specific_config(self.paut_template_file_path.get(), "PAUT"))
        self.template_file_path.trace_add("write", lambda *args: self.load_template_specific_config(self.template_file_path.get(), "PMI"))

        # Keys are now verified to match Treeview column indices 1-N.

        # 3. UI Initialization
        # [NEW] File Info Strings for Preview Headers
        self.file_info_vars = {
            'PMI': tk.StringVar(value="📄 파일을 선택해주세요."),
            'RT': tk.StringVar(value="📄 파일을 선택해주세요."),
            'KOGAS': tk.StringVar(value="📄 파일을 선택해주세요."),
            'PT': tk.StringVar(value="📄 파일을 선택해주세요."),
            'MT': tk.StringVar(value="📄 파일을 선택해주세요."),
            'PAUT': tk.StringVar(value="📄 파일을 선택해주세요."),
            'PHOTO': tk.StringVar(value="📸 사진 리스트를 구성해주세요.")
        }

        # --- Gapji (Cover) Metadata Variables ---
        self.gapji_project = tk.StringVar(value=self.config.get(
            'GAPJI_PROJECT', "2026년 중앙지사 열수송관 비파괴검사 단가계약"
        ))
        self.gapji_customer = tk.StringVar(value=self.config.get(
            'GAPJI_CUSTOMER', "한국지역난방공사 중앙지사"
        ))
        self.gapji_item = tk.StringVar(value=self.config.get('GAPJI_ITEM', "PIPE"))
        self.gapji_material = tk.StringVar(value=self.config.get('GAPJI_MATERIAL', ""))
        self.paut_sec = tk.StringVar(value=self.config.get('PAUT_SEC', ''))
        self.paut_contractor = tk.StringVar(value=self.config.get('PAUT_CONTRACTOR', 'GS네오텍'))
        self.mt_surface_temp = tk.StringVar(value=self.config.get('MT_SURFACE_TEMP', ''))
        self.gapji_report_no = tk.StringVar(value=self.config.get('GAPJI_REPORT_NO', ""))
        self.gapji_exam_date = tk.StringVar(value=self.config.get('GAPJI_EXAM_DATE', datetime.datetime.now().strftime("%Y-%m-%d")))
        self.pmi_chem_elements = [
            tk.StringVar(value=self.config.get(f'PMI_CHEM_ELEMENT_{idx}', default))
            for idx, default in enumerate(("Ni", "Cr", "Mo", "-", "-"), start=1)
        ]
        self.pmi_chem_specs = [
            tk.StringVar(value=self.config.get(f'PMI_CHEM_SPEC_{idx}', ""))
            for idx in (1, 2, 3)
        ]
        self.pmi_spec1_mat = tk.StringVar(value=self.config.get('PMI_SPEC1_MAT', ''))
        self.pmi_spec2_mat = tk.StringVar(value=self.config.get('PMI_SPEC2_MAT', ''))
        self.pmi_spec3_mat = tk.StringVar(value=self.config.get('PMI_SPEC3_MAT', ''))
        self.rt_isotope = tk.StringVar(value=self.config.get('RT_ISOTOPE', 'Ir-192'))
        self.rt_activity = tk.StringVar(value=self.config.get('RT_ACTIVITY', '17'))
        self.rt_activity_unit = tk.StringVar(value=self.config.get('RT_ACTIVITY_UNIT', 'Ci'))
        self.rt_film_brand = tk.StringVar(value=self.config.get('RT_FILM_BRAND', 'FUJI#80'))
        self.rt_film_type = tk.StringVar(value=self.config.get('RT_FILM_TYPE', 'I'))
        self.rt_film_size = tk.StringVar(value=self.config.get('RT_FILM_SIZE', '3⅓ X 12'))
        self.rt_exposure_time = tk.StringVar(value=self.config.get('RT_EXPOSURE_TIME', '2 MIN 40 SEC'))
        
        # [NEW] Trace Gapji metadata for real-time preview updates
        for var in [self.gapji_project, self.gapji_customer, self.gapji_item, 
                    self.gapji_material, self.gapji_report_no, self.gapji_exam_date,
                    self.rt_isotope, self.rt_activity, self.rt_activity_unit,
                    self.rt_film_brand, self.rt_film_type, self.rt_film_size, self.rt_exposure_time]:
            var.trace_add("write", lambda *a: self._update_gapji_preview(getattr(self, 'current_mode', 'PMI')))

        # --- Photo Log State Variables ---
        self.photo_header_map = {
            "PAUT": "REPORT OF PHASED ARRAY UT EXAMINATION (위 상 배 열 초 음 파 탐 상 검 사 보 고 서)",
            "RT": "REPORT OF RADIOGRAPHIC EXAMINATION (방 사 선 투 과 검 사 보 고 서)",
            "PT": "REPORT OF LIQUID PENETRANT EXAMINATION (침 투 탐 상 검 사 보 고 서)",
            "MT": "REPORT OF MAGNETIC PARTICLE EXAMINATION (자 분 탐 상 검 사 보 고 서)",
            "PMI": "REPORT OF POSITIVE MATERIAL IDENTIFICATION (재 질 성 분 분 석 검 사 보 고 서)",
            "UT": "REPORT OF ULTRASONIC EXAMINATION (초 음 파 탐 상 검 사 보 고 서)",
            "NDT": "REPORT OF NON-DESTRUCTIVE EXAMINATION (비 파 괴 검 사 보 고 서)",
            "기타 (직접 입력)": ""
        }
        self.photo_orderer = tk.StringVar(value="서울에너지공사")
        self.photo_inspect_date = tk.StringVar(value=datetime.datetime.now().strftime("%Y-%m-%d"))
        self.photo_report_no = tk.StringVar(value="SIT/GI-SE-PAUT-TNTFJPWJ001")
        self.photo_inspect_type = tk.StringVar(value="PAUT")
        self.photo_report_title = tk.StringVar(value=self.photo_header_map["PAUT"])
        self.photo_cols_per_row = tk.StringVar(value="2")
        self.photo_rows_per_page = tk.StringVar(value="4")
        self.photo_keep_aspect = tk.BooleanVar(value=True)
        self.photo_output_name = tk.StringVar(value="NDT_Photo_Log_Final.xlsx")
        _def_logo = os.path.join(RESOURCE_DIR, "logo.png")
        self.photo_logo_path = tk.StringVar(value=_def_logo if os.path.exists(_def_logo) else "")
        self.photo_logo_width_var = tk.StringVar(value="80")
        self.photo_logo_height_var = tk.StringVar(value="") # [NEW] Custom Logo Height
        self.photo_logo_x_var = tk.StringVar(value="2")
        self.photo_logo_y_var = tk.StringVar(value="0")
        self.photo_cell_width_var = tk.StringVar(value="53.0")
        self.photo_cell_height_var = tk.StringVar(value="178.0")
        self.photo_margin_top_var = tk.StringVar(value="0.5")
        self.photo_margin_bottom_var = tk.StringVar(value="0.5")
        self.photo_margin_left_var = tk.StringVar(value="0.4")
        self.photo_margin_right_var = tk.StringVar(value="0.4")
        self.photo_print_scale_var = tk.StringVar(value="100")
        self.photo_desc_height_var = tk.StringVar(value="20.0")
        self.photo_align_var = tk.StringVar(value="중앙 정렬")
        self.photo_fit_width_var = tk.BooleanVar(value=False)
        self.photo_auto_fit_page_var = tk.BooleanVar(value=True) # [NEW] Auto-fit to A4 page
        self.photo_auto_rotate_var = tk.BooleanVar(value=False)
        self.photo_width_pct_var = tk.StringVar(value="100.0")
        self.photo_height_pct_var = tk.StringVar(value="140.0")
        self.photo_width_pixel_adj_var = tk.StringVar(value="0")
        self.photo_height_pixel_adj_var = tk.StringVar(value="10") # [NEW] Height padding inside cell
        self.photo_shift_x_var = tk.StringVar(value="0")
        self.photo_shift_y_var = tk.StringVar(value="0")
        self.photo_dpi_var = tk.StringVar(value="96") # [NEW] User Configurable DPI
        self.photo_selected_files = []
        self.photo_item_metadata = {}

        # 검사방법별 기본정보를 저장하고 사진대장 정보에도 즉시 반영한다.
        for var in (
            self.gapji_project, self.gapji_customer, self.gapji_item,
            self.gapji_material, self.gapji_report_no, self.gapji_exam_date,
        ):
            var.trace_add("write", lambda *_args: self._on_report_info_changed())
        
        self.load_settings()
        self._initialize_report_info_by_mode()
        self._load_report_info_for_mode("PMI")
        self._sync_photo_report_info()
        self.photo_report_no.trace_add('write', lambda *_args: self._sync_photo_output_name())
        self._sync_photo_output_name()
        
        # [MIGRATION] Add missing keys from loaded config
        if "Size" not in self.paut_column_keys:
            try:
                idx = self.paut_column_keys.index("Joint No.")
                self.paut_column_keys.insert(idx + 1, "Size")
            except:
                self.paut_column_keys.append("Size")
        if "Acc" not in self.paut_column_keys:
            try:
                idx = self.paut_column_keys.index("Th'k(mm)")
                self.paut_column_keys.insert(idx + 1, "Acc")
                self.paut_column_keys.insert(idx + 2, "Rej")
            except:
                self.paut_column_keys.extend(["Acc", "Rej"])
                
        self.last_photo_save_dir = self.config.get('PHOTO_LOG_SETTINGS', {}).get('last_save_dir', "")

        self.create_widgets()
        
        # [NEW] Sync initial file info from loaded config
        self._sync_all_file_infos()
        
        # [NEW] Auto-save on exit
        if not self.is_embedded and hasattr(self.root, "protocol"):
            try: self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
            except Exception: pass
        
        self._is_ready = True
        self.log("[INFO] 통합 버전을 시작했습니다.")
        
    def _safe_update_scrollregion(self):
        """Safely update canvas scrollregion, handling None bbox."""
        try:
            bbox = self.canvas.bbox("all")
            if bbox:
                self.canvas.configure(scrollregion=bbox)
        except Exception:
            pass

    def log(self, message):
        if hasattr(self, 'status_log'):
            self.status_log.config(state='normal')
            self.status_log.insert(tk.END, f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {message}\n")
            self.status_log.see(tk.END)
            self.status_log.config(state='disabled')
        else:
            print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {message}")

    def fix_material_name(self, t):
        """재질명을 표준 포맷으로 변환"""
        t_str = str(t).upper()
        if t_str == 'NAN': return ""
        t_str = t_str.replace('A312-TP304', 'S/S').replace('A312-304L', 'S/S').replace('A312-305L', 'S/S').replace('A53-B', 'C/S').replace('A106-B', 'C/S')
        return t_str.replace('C2','C/S').replace('C4','C/S').replace('CS','C/S').replace('S99','S/S').replace('SS','S/S')

    def force_two_digit(self, val):
        """숫자 값을 2자리 문자열로 변환 (예: 1 -> 01)"""
        try:
            s = str(val).strip()
            f = float(s)
            if f.is_integer(): return f"{int(f):02d}"
            return s
        except: return str(val).strip()

    def load_settings(self):
        # 1. 내장 설정 파일 (실행파일 내부에 번들링된 기본값)
        if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
            bundle_settings = os.path.join(sys._MEIPASS, "logo_settings_unified.json")
            if os.path.exists(bundle_settings):
                try:
                    with open(bundle_settings, 'r', encoding='utf-8') as f:
                        bundle_data = json.load(f)
                        self.config.update(bundle_data)
                    print("SUCCESS: 내장된 기본 설정을 로드했습니다.")
                except Exception as e:
                    print(f"WARNING: 내장 설정 불러오기 실패: {e}")

        # 2. 외부 설정 파일 (사용자 저장값 - 내장값보다 우선순위 높음)
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                    saved_data = json.load(f)
                    
                    # [REFINED] Case-insensitive Sanitization for all modules
                    sanit_map = {
                        'PMI_NAME_DWG': (['ISO/DWG', 'ISO/DWG.'], 'Drawing No.'),
                        'PMI_NAME_JOINT': (['JOINT NO', 'JOINT NO.'], 'Joint No.'),
                        'PMI_NAME_LOC': (['TEST LOCATION', 'LOC'], 'Location'),
                        'PMI_NAME_RES': (['GRADE', 'RESULT'], 'Result'),
                        'RT_NAME_DWG': (['ISO/DWG', 'ISO/DWG.', 'DWG'], 'Drawing No.'),
                        'RT_NAME_JOINT': (['JOINT NO', 'FILM IDENT', 'JOINT'], 'Film Ident. No.'),
                        'RT_NAME_LOC': (['LOC', 'LOC.'], 'Film Location'),
                        'RT_NAME_WELDER': (['WELDER'], 'Welder No'),
                        'RT_NAME_DEG': (['물성', '물성(DEG)', 'DEG'], 'Deg')
                    }
                    for k, (olds, new) in sanit_map.items():
                        if k in saved_data:
                            val = str(saved_data[k]).upper().strip()
                            if any(val == o.upper() for o in olds):
                                saved_data[k] = new

                    # [NEW] Sanitize old "Element-" prefixes and Mn exclusion
                    for k in ['PMI_NAME_NI', 'PMI_NAME_CR', 'PMI_NAME_MO']:
                        if k in saved_data and isinstance(saved_data[k], str) and (saved_data[k].startswith("Element-") or "ELEMENT" in saved_data[k].upper()):
                            saved_data[k] = saved_data[k].replace("Element-", "").replace("ELEMENT-", "").replace("Element", "").strip()
                    
                    if 'column_keys' in saved_data and isinstance(saved_data['column_keys'], list):
                        while 'Mn' in saved_data['column_keys']: saved_data['column_keys'].remove('Mn')

                    # [NEW] Defect Names Sanitization (D1-D15 to circled labels)
                    defect_standards = {
                        1: "① Crack", 2: "② IP", 3: "③ LF", 4: "④ Slag", 5: "⑤ Por",
                        6: "⑥ U/C", 7: "⑦ RUC", 8: "⑧ BT", 9: "⑨ TI", 10: "⑩ CP",
                        11: "⑪ RC", 12: "⑫ Mis", 13: "⑬ EP", 14: "⑭ SD", 15: "⑮ Oth"
                    }
                    for d_i, std_name in defect_standards.items():
                        d_k = f'RT_NAME_D{d_i}'
                        if d_k in saved_data:
                            curr_val = str(saved_data[d_k]).upper().strip()
                            if curr_val == f"D{d_i}" or not curr_val or curr_val == "NAN":
                                saved_data[d_k] = std_name
                    
                    # [NEW] Restore Photo Log Variables
                    photo_vars = {
                        'orderer': self.photo_orderer, 'inspect_date': self.photo_inspect_date,
                        'inspect_type': self.photo_inspect_type, 'report_title': self.photo_report_title,
                        'report_no': self.photo_report_no,
                        'cols_per_row': self.photo_cols_per_row, 'rows_per_page': self.photo_rows_per_page, 'keep_aspect': self.photo_keep_aspect,
                        'output_name': self.photo_output_name, 'logo_path': self.photo_logo_path,
                        'logo_width': self.photo_logo_width_var, 'logo_height': self.photo_logo_height_var,
                        'logo_x': self.photo_logo_x_var, 'logo_y': self.photo_logo_y_var,
                        'cell_width': self.photo_cell_width_var, 'cell_height': self.photo_cell_height_var,
                        'm_top': self.photo_margin_top_var, 'm_bottom': self.photo_margin_bottom_var,
                        'm_left': self.photo_margin_left_var, 'm_right': self.photo_margin_right_var,
                        'print_scale': self.photo_print_scale_var, 'desc_height': self.photo_desc_height_var,
                        'photo_align': self.photo_align_var, 'auto_fit_page': self.photo_auto_fit_page_var, 'fit_width': self.photo_fit_width_var,
                        'auto_rotate': self.photo_auto_rotate_var, 'width_pct': self.photo_width_pct_var,
                        'height_pct': self.photo_height_pct_var,
                        'pixel_adj': self.photo_width_pixel_adj_var, 'shift_x': self.photo_shift_x_var,
                        'shift_y': self.photo_shift_y_var, 'dpi': self.photo_dpi_var
                    }
                    if 'PHOTO_LOG_SETTINGS' in saved_data:
                        plist = saved_data['PHOTO_LOG_SETTINGS']
                        for pk, pvar in photo_vars.items():
                            if pk in plist:
                                if isinstance(pvar, tk.BooleanVar): pvar.set(bool(plist[pk]))
                                else: pvar.set(str(plist[pk]))
                        # Log success for key fine-tune variables
                        self.log(f"[Restore] 사진대장 설정 복구: 비율({self.photo_width_pct_var.get()}%), 추가({self.photo_width_pixel_adj_var.get()}px), 이동({self.photo_shift_x_var.get()}px)")
                        if 'selected_files' in plist:
                            self.photo_selected_files = plist['selected_files']
                            if hasattr(self, 'photo_listbox'):
                                self.photo_listbox.delete(0, tk.END)
                                for f_path in self.photo_selected_files:
                                    self.photo_listbox.insert(tk.END, f_path)
                        if isinstance(plist.get('item_metadata'), dict):
                            self.photo_item_metadata = plist['item_metadata']

                    # [NEW] Restore custom column layouts if they exist
                    if 'column_keys' in saved_data and isinstance(saved_data['column_keys'], list):
                        self.column_keys = list(saved_data['column_keys'])
                        if "Report No" not in self.column_keys:
                            try: self.column_keys.insert(self.column_keys.index("Date") + 1, "Report No")
                            except: self.column_keys.insert(3, "Report No")
                    if 'rt_column_keys' in saved_data and isinstance(saved_data['rt_column_keys'], list):
                        self.rt_column_keys = list(saved_data['rt_column_keys'])
                        if "Report No" not in self.rt_column_keys:
                            try: self.rt_column_keys.insert(self.rt_column_keys.index("Date") + 1, "Report No")
                            except: self.rt_column_keys.insert(3, "Report No")
                        if "Sec" not in self.rt_column_keys:
                            insert_at = self.rt_column_keys.index("Date") + 2 if "Date" in self.rt_column_keys else 4
                            self.rt_column_keys.insert(insert_at, "Sec")
                    if 'kogas_column_keys' in saved_data and isinstance(saved_data['kogas_column_keys'], list):
                        self.kogas_column_keys = list(saved_data['kogas_column_keys'])
                        if "Report No" not in self.kogas_column_keys:
                            try: self.kogas_column_keys.insert(self.kogas_column_keys.index("Date") + 1, "Report No")
                            except: self.kogas_column_keys.insert(3, "Report No")
                    if 'pt_column_keys' in saved_data and isinstance(saved_data['pt_column_keys'], list):
                        self.pt_column_keys = list(saved_data['pt_column_keys'])
                        if "Report No" not in self.pt_column_keys:
                            try: self.pt_column_keys.insert(self.pt_column_keys.index("Date") + 1, "Report No")
                            except: self.pt_column_keys.insert(3, "Report No")
                        # [마이그레이션] 구버전 config에 Thk.가 없으면 Joint 뒤에 삽입
                        if "Thk." not in self.pt_column_keys:
                            try:
                                insert_at = self.pt_column_keys.index("Joint") + 1
                            except ValueError:
                                insert_at = 5
                            self.pt_column_keys.insert(insert_at, "Thk.")
                    
                    if "No" in self.pt_column_keys:
                        self.pt_column_keys.remove("No")
                    self.pt_column_keys.insert(1, "No")

                    if "NPS" in self.pt_column_keys:
                        self.pt_column_keys.remove("NPS")
                    try:
                        self.pt_column_keys.insert(self.pt_column_keys.index("Joint") + 1, "NPS")
                    except ValueError:
                        self.pt_column_keys.insert(6, "NPS")
                    if 'mt_column_keys' in saved_data:
                        self.mt_column_keys = list(saved_data['mt_column_keys'])
                        if "Report No" not in self.mt_column_keys:
                            try: self.mt_column_keys.insert(self.mt_column_keys.index("Date") + 1, "Report No")
                            except: self.mt_column_keys.insert(3, "Report No")
                    if 'paut_column_keys' in saved_data and isinstance(saved_data['paut_column_keys'], list):
                        loaded_paut_keys = list(saved_data['paut_column_keys'])
                        # [FIX] Ensure new columns are added even if loading from old config
                        default_paut_keys = ["selected", "No", "Date", "Report No", "Line No.", "Joint No.", "Size", "Th'k(mm)", "Acc", "Rej", "Start", "End", "Length(mm)", "Upper", "Lower", "Height(mm)", "TCG MAX(%)", "Type of Flaw", "a/l", "a/t", "Welder", "Tested Length", "Evaluation", "Remarks"]
                        for dk in default_paut_keys:
                            if dk not in loaded_paut_keys:
                                insert_idx = default_paut_keys.index(dk)
                                if insert_idx <= len(loaded_paut_keys):
                                    loaded_paut_keys.insert(insert_idx, dk)
                                else:
                                    loaded_paut_keys.append(dk)
                        self.paut_column_keys = loaded_paut_keys
                    self.config.update(saved_data)
                    
                    # Restore Path StringVars in UI
                    pmi_tp = self.config.get('PMI_TEMPLATE_PATH', "")
                    if pmi_tp and not os.path.exists(pmi_tp):
                        for candidate in ("home/data/PMI-템플릿보고서.xlsx", "home/data/SIT-PMI-001.xlsx", "home/data/SIT-PMI-K3 롯데 템플릿보고서.xlsm"):
                            full_c = os.path.abspath(candidate)
                            if os.path.exists(full_c):
                                pmi_tp = full_c
                                break
                    self.target_file_path.set(self.config.get('PMI_TARGET_PATH', ""))
                    self.template_file_path.set(pmi_tp)
                    self.rt_target_file_path.set(self.config.get('RT_TARGET_PATH', ""))
                    self.rt_template_file_path.set(self.config.get('RT_TEMPLATE_PATH', ""))
                    self.kogas_target_file_path.set(self.config.get('KOGAS_TARGET_PATH', ""))
                    self.kogas_template_file_path.set(self.config.get('KOGAS_TEMPLATE_PATH', ""))
                    self.pt_target_file_path.set(self.config.get('PT_TARGET_PATH', ""))
                    self.pt_template_file_path.set(self.config.get('PT_TEMPLATE_PATH', ""))
                    self.mt_target_file_path.set(self.config.get('MT_TARGET_PATH', ''))
                    self.mt_template_file_path.set(self.config.get('MT_TEMPLATE_PATH', ''))
                    self.mt_surface_temp.set(str(self.config.get('MT_SURFACE_TEMP', '')))
                    self.paut_target_file_path.set(self.config.get('PAUT_TARGET_PATH', ""))
                    self.paut_template_file_path.set(self.config.get('PAUT_TEMPLATE_PATH', ""))
                self.log("[SUCCESS] 사용자 저장 설정을 모두 불러왔습니다.")
                
                # Restore Window Geometry/State
                if 'WINDOW_GEOMETRY' in saved_data and not getattr(self, 'is_embedded', False) and hasattr(self.root, 'geometry'):
                    try: self.root.geometry(saved_data['WINDOW_GEOMETRY'])
                    except: pass
                if 'WINDOW_STATE' in saved_data and not getattr(self, 'is_embedded', False) and hasattr(self.root, 'state'):
                    try: self.root.state(saved_data['WINDOW_STATE'])
                    except: pass
                if 'ACTIVE_TAB' in saved_data:
                    self._initial_active_tab = saved_data['ACTIVE_TAB']
            except Exception as e:
                print(f"WARNING: 사용자 설정 불러오기 실패: {e}")

        # PAUT 마이그레이션 (UT_* -> PAUT_*)
        for k in list(self.config.keys()):
            if k.startswith("UT_"):
                new_k = k.replace("UT_", "PAUT_")
                if new_k not in self.config:
                    self.config[new_k] = self.config[k]
        
        # [NEW] Load Photo Log specific config
        self.load_photo_log_config()

    def capture_ui_state(self):
        """UI 요소들의 현재 상태(컬럼 너비, 분할선 위치 등)를 config에 반영"""
        try:
            d = self.__dict__
            trees = [
                ('PMI', d.get('preview_tree')), 
                ('RT', d.get('rt_preview_tree')), 
                ('PT', d.get('pt_preview_tree')), 
                ('MT', d.get('mt_preview_tree')),
                ('PAUT', d.get('paut_preview_tree')),
                ('KOGAS', d.get('kogas_preview_tree'))
            ]
            for mode, tree in trees:
                if tree:
                    try:
                        if tree.winfo_ismapped() and tree.winfo_width() > 100:
                            self.config[f"{mode}_COL_WIDTHS"] = {col: tree.column(col, "width") for col in tree["columns"]}
                    except: pass

            # 2. 패널 분할선(Sash) 비율 캡처
            paned_windows = [
                ('PMI', d.get('pmi_paned')),
                ('RT', d.get('rt_paned')),
                ('PT', d.get('pt_paned')),
                ('PAUT', d.get('paut_paned'))
            ]
            for mode, pw in paned_windows:
                if pw:
                    try:
                        total_w = pw.winfo_width()
                        if total_w > 100:
                            self.config[f'{mode}_SASH_RATIO'] = pw.sash_coord(0)[0] / total_w
                    except: pass

            # 3. 윈도우 및 활성 탭 상태 캡처
            try:
                if not getattr(self, 'is_embedded', False) and hasattr(self.root, 'geometry'):
                    self.config['WINDOW_GEOMETRY'] = self.root.geometry()
                if not getattr(self, 'is_embedded', False) and hasattr(self.root, 'state'):
                    self.config['WINDOW_STATE'] = self.root.state()
                if 'mode_notebook' in d:
                    self.config['ACTIVE_TAB'] = self.mode_notebook.index("current")
                
                # 4. 파일 경로 관성 저장
                self.config['PMI_TARGET_PATH'] = self.target_file_path.get()
                self.config['PMI_TEMPLATE_PATH'] = self.template_file_path.get()
                self.config['RT_TARGET_PATH'] = self.rt_target_file_path.get()
                self.config['RT_TEMPLATE_PATH'] = self.rt_template_file_path.get()
                self.config['PT_TARGET_PATH'] = self.pt_target_file_path.get()
                self.config['PT_TEMPLATE_PATH'] = self.pt_template_file_path.get()
                self.config['MT_TARGET_PATH'] = self.mt_target_file_path.get()
                self.config['MT_TEMPLATE_PATH'] = self.mt_template_file_path.get()
                self.config['PAUT_TARGET_PATH'] = self.paut_target_file_path.get()
                self.config['PAUT_TEMPLATE_PATH'] = self.paut_template_file_path.get()
                self.config['KOGAS_TARGET_PATH'] = self.kogas_target_file_path.get()
                self.config['KOGAS_TEMPLATE_PATH'] = self.kogas_template_file_path.get()
                if 'pmi_layout_mode' in d:
                    self.config['PMI_LAYOUT_MODE'] = self.pmi_layout_mode.get()
                
                if 'paut_equip_vars' in d:
                    for k, var in self.paut_equip_vars.items():
                        self.config[f'PAUT_EQUIP_{k.upper()}'] = var.get()
                
                if 'paut_photo_pages_var' in d:
                    self.config['PAUT_PHOTO_PAGES'] = str(self.paut_photo_pages_var.get())
                # 5. Photo Log Settings Capture
                prev_photo_cfg = self.config.get('PHOTO_LOG_SETTINGS', {}) if isinstance(self.config.get('PHOTO_LOG_SETTINGS'), dict) else {}
                self.config['PHOTO_LOG_SETTINGS'] = {
                    'orderer': self.photo_orderer.get(), 'inspect_date': self.photo_inspect_date.get(),
                    'inspect_type': self.photo_inspect_type.get(), 'report_title': self.photo_report_title.get(),
                    'report_no': self.photo_report_no.get(),
                    'start_page': d['photo_start_page'].get() if 'photo_start_page' in d else prev_photo_cfg.get('start_page', str(self.config.get('PHOTO_START_PAGE', ''))),
                    'total_pages': d['photo_total_pages'].get() if 'photo_total_pages' in d else prev_photo_cfg.get('total_pages', str(self.config.get('PHOTO_TOTAL_PAGES', ''))),
                    'cols_per_row': self.photo_cols_per_row.get(), 'rows_per_page': self.photo_rows_per_page.get(), 'keep_aspect': self.photo_keep_aspect.get(),
                    'output_name': self.photo_output_name.get(), 'logo_path': self.photo_logo_path.get(),
                    'logo_width': self.photo_logo_width_var.get(), 'logo_height': self.photo_logo_height_var.get(),
                    'logo_x': self.photo_logo_x_var.get(), 'logo_y': self.photo_logo_y_var.get(),
                    'cell_width': self.photo_cell_width_var.get(), 'cell_height': self.photo_cell_height_var.get(),
                    'm_top': self.photo_margin_top_var.get(), 'm_bottom': self.photo_margin_bottom_var.get(),
                    'm_left': self.photo_margin_left_var.get(), 'm_right': self.photo_margin_right_var.get(),
                    'print_scale': self.photo_print_scale_var.get(), 'desc_height': self.photo_desc_height_var.get(),
                    'photo_align': self.photo_align_var.get(), 'auto_fit_page': self.photo_auto_fit_page_var.get(), 'fit_width': self.photo_fit_width_var.get(),
                    'auto_rotate': self.photo_auto_rotate_var.get(), 'width_pct': self.photo_width_pct_var.get(),
                    'height_pct': self.photo_height_pct_var.get(),
                    'pixel_adj': self.photo_width_pixel_adj_var.get(), 'shift_x': self.photo_shift_x_var.get(),
                    'shift_y': self.photo_shift_y_var.get(), 'dpi': self.photo_dpi_var.get(),
                    'selected_files': self.photo_selected_files,
                    'item_metadata': self.photo_item_metadata,
                    'last_save_dir': d.get('last_photo_save_dir', "")
                }
                
                # [NEW] Save custom column list configurations
                self.config['column_keys'] = list(self.column_keys)
                self.config['rt_column_keys'] = list(self.rt_column_keys)
                self.config['kogas_column_keys'] = list(self.kogas_column_keys)
                self.config['pt_column_keys'] = list(self.pt_column_keys)
                self.config['mt_column_keys'] = list(self.mt_column_keys)
                self.config['paut_column_keys'] = list(self.paut_column_keys)
            except Exception as e:
                self.log(f"[ERROR] UI 상태 캡처 실패: {e}")
        except Exception as e:
            self.log(f"[ERROR] 전체 상태 캡처 실패: {e}")

    def save_photo_log_config(self):
        """Save Photo Log specific layout settings to a separate JSON."""
        config_path = os.path.join(os.getcwd(), "photolog_config.json")
        self.capture_ui_state()
        plist = self.config.get('PHOTO_LOG_SETTINGS', {})
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(plist, f, indent=4, ensure_ascii=False)
            self.log("💾 사진대장 전용 설정이 저장되었습니다.")
            messagebox.showinfo("저장 완료", "사진대장 레이아웃 설정이 저장되었습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"저장 실패: {e}")

    def load_photo_log_config(self):
        """Load Photo Log specific layout settings."""
        config_path = os.path.join(os.getcwd(), "photolog_config.json")
        if not os.path.exists(config_path): return
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                plist = json.load(f)
            
            # Update UI vars (using same logic as in load_settings)
            photo_vars = {
                'orderer': self.photo_orderer, 'inspect_date': self.photo_inspect_date,
                'inspect_type': self.photo_inspect_type, 'report_title': self.photo_report_title,
                'report_no': self.photo_report_no,
                'start_page': self.__dict__.get('photo_start_page'),
                'total_pages': self.__dict__.get('photo_total_pages'),
                'cols_per_row': self.photo_cols_per_row, 'rows_per_page': self.photo_rows_per_page, 'keep_aspect': self.photo_keep_aspect,
                'output_name': self.photo_output_name, 'logo_path': self.photo_logo_path,
                'logo_width': self.photo_logo_width_var, 'logo_height': self.photo_logo_height_var,
                'logo_x': self.photo_logo_x_var, 'logo_y': self.photo_logo_y_var,
                'cell_width': self.photo_cell_width_var, 'cell_height': self.photo_cell_height_var,
                'm_top': self.photo_margin_top_var, 'm_bottom': self.photo_margin_bottom_var,
                'm_left': self.photo_margin_left_var, 'm_right': self.photo_margin_right_var,
                'print_scale': self.photo_print_scale_var, 'desc_height': self.photo_desc_height_var,
                'photo_align': self.photo_align_var, 'auto_fit_page': self.photo_auto_fit_page_var, 'fit_width': self.photo_fit_width_var,
                'auto_rotate': self.photo_auto_rotate_var, 'width_pct': self.photo_width_pct_var,
                'height_pct': self.photo_height_pct_var,
                'pixel_adj': self.photo_width_pixel_adj_var, 'shift_x': self.photo_shift_x_var,
                'shift_y': self.photo_shift_y_var, 'dpi': self.photo_dpi_var
            }
            for pk, pvar in photo_vars.items():
                if pk in plist and pvar is not None:
                    if isinstance(pvar, tk.BooleanVar): pvar.set(bool(plist[pk]))
                    else: pvar.set(str(plist[pk]))
            
            if 'selected_files' in plist:
                self.photo_selected_files = plist['selected_files']
                if 'photo_listbox' in self.__dict__:
                    self.photo_listbox.delete(0, tk.END)
                    for f_path in self.photo_selected_files:
                        self.photo_listbox.insert(tk.END, f_path)
            if isinstance(plist.get('item_metadata'), dict):
                self.photo_item_metadata = plist['item_metadata']
            
            self.log("📂 사진대장 전용 설정을 불러왔습니다.")
        except Exception as e:
            print(f"Error loading photolog config: {e}")

    def load_template_specific_config(self, template_path, mode="RT"):
        """템플릿 파일과 연동된 JSON 설정 파일을 로드하여 UI에 반영"""
        if not template_path or not os.path.exists(template_path): return
        
        # [NEW] Template-name based heuristics (e.g. K1/K3/롯데/플랜트 배관용)
        if mode == "PMI":
            if any(k in os.path.basename(template_path).upper() for k in ["LOTTE", "PROVIDENCE", "롯데", "모재", "3단"]):
                if hasattr(self, 'pmi_layout_mode'):
                    self.pmi_layout_mode.set("PLANT_TRIPLE")
            else:
                if hasattr(self, 'pmi_layout_mode'):
                    self.pmi_layout_mode.set("STANDARD")

        config_path = os.path.splitext(template_path)[0] + ".json"
        if not os.path.exists(config_path): return
        
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                t_config = json.load(f)
            
            self.log(f"📂 [Auto-Load] '{os.path.basename(template_path)}' 전용 설정을 불러왔습니다.")
            
            # Update self.config and UI variables
            for key, value in t_config.items():
                self.config[key] = value
                if key in self.setting_vars:
                    self.setting_vars[key].set(str(value))
            
            # Mode-specific UI update (Force refresh logic if needed)
            if mode in ["RT", "KOGAS"]:
                self._update_rt_preview_columns(mode)
            elif mode == "PT":
                try: self._update_pt_preview_columns()
                except: pass
            elif mode == "PAUT":
                try: self._update_paut_target_fsh()
                except: pass
            elif mode == "PMI":
                try:
                    if "PMI_LAYOUT_MODE" in t_config and hasattr(self, 'pmi_layout_mode'):
                        self.pmi_layout_mode.set(t_config["PMI_LAYOUT_MODE"])
                    elif any(k in os.path.basename(template_path).upper() for k in ["LOTTE", "PROVIDENCE", "롯데", "모재", "3단"]):
                        if hasattr(self, 'pmi_layout_mode'):
                            self.pmi_layout_mode.set("PLANT_TRIPLE")
                    else:
                        if hasattr(self, 'pmi_layout_mode'):
                            self.pmi_layout_mode.set("STANDARD")
                    self._update_pmi_filter_ui()
                except: pass

        except Exception as e:
            self.log(f"[ERROR] 템플릿 전용 설정 로드 실패: {e}")

    def save_template_specific_config(self, mode="RT"):
        """현재 설정을 선택된 템플릿 전용 JSON 파일로 저장"""
        self.save_settings()
        if mode == "RT": template_path = self.rt_template_file_path.get()
        elif mode == "KOGAS": template_path = self.kogas_template_file_path.get()
        elif mode == "PT": template_path = self.pt_template_file_path.get()
        elif mode == "MT": template_path = self.mt_template_file_path.get()
        elif mode == "PAUT": template_path = self.paut_template_file_path.get()
        else: template_path = self.template_file_path.get()
        
        if not template_path or not os.path.exists(template_path):
            messagebox.showwarning("경고", "먼저 템플릿 파일을 선택해주세요.")
            return
            
        config_path = os.path.splitext(template_path)[0] + ".json"
        self.capture_ui_state() # Ensure current UI is in self.config
        
        # Determine relevant keys to save for this template
        # 1. Logos & Print settings (Template-dependent)
        keys_to_save = []
        for ctx in ["COVER", "DATA"]:
            # Correct Prefix mapping: SITCO_RT_COVER, etc.
            # Actually, mode is like "RT", ctx is "COVER" -> prefix is "RT_COVER"
            # In UI: self._create_setting_grid(t_cover, "RT_COVER")
            # Inside grid: keys are SITCO_RT_COVER_X etc.
            for prefix in ["SITCO", "SEOUL", "FOOTER", "FOOTER_PT"]:
                for suffix in ["_PATH", "_ANCHOR", "_W", "_H", "_X", "_Y"]:
                    keys_to_save.append(f"{prefix}_{mode}_{ctx}{suffix}")
            for suffix in ["_TOP", "_BOTTOM", "_LEFT", "_RIGHT"]:
                keys_to_save.append(f"MARGIN_{mode}_{ctx}{suffix}")
            keys_to_save.append(f"PRINT_SCALE_{mode}_{ctx}")
            keys_to_save.append(f"PRINT_AREA_{mode}_{ctx}")
            
        # 2. Mode-specific boundaries and column mappings
        if mode == "RT":
            keys_to_save += ["RT_START_ROW", "RT_END_ROW"]
            keys_to_save += ["RT_FILM_BRAND", "RT_FILM_TYPE", "RT_FILM_SIZE",
                             "RT_ISOTOPE", "RT_ACTIVITY", "RT_ACTIVITY_UNIT"]
            keys_to_save += [k for k in self.config.keys() if k.startswith("RT_COL_") or k.startswith("RT_NAME_")]
        elif mode == "KOGAS":
            keys_to_save += ["KOGAS_START_ROW", "KOGAS_DATA_END_ROW", "RT_KOGAS_D_START_COL"]
            keys_to_save += [k for k in self.config.keys() if k.startswith("KOGAS_")]
        elif mode == "PT":
            keys_to_save += ["PT_START_ROW", "PT_END_ROW"]
            keys_to_save += [k for k in self.config.keys() if k.startswith("PT_COL_") or k.startswith("PT_NAME_")]
        elif mode == "MT":
            keys_to_save += [k for k in self.config if k.startswith("MT_")]
        elif mode == "PAUT":
            keys_to_save += ["PAUT_START_ROW", "PAUT_END_ROW"]
            keys_to_save += [k for k in self.config.keys() if k.startswith("PAUT_COL_") or k.startswith("PAUT_NAME_")]
        elif mode == "PMI":
            keys_to_save += ["PMI_START_ROW", "PMI_DATA_END_ROW", "PMI_PRINT_END_ROW", "PMI_LAYOUT_MODE"]
            keys_to_save += [k for k in self.config.keys() if k.startswith("PMI_COL_") or k.startswith("PMI_NAME_")]

        # Build template-specific dict
        t_data = {k: self.config[k] for k in keys_to_save if k in self.config}
        
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(t_data, f, indent=4, ensure_ascii=False)
            self.log(f"💾 [SUCCESS] '{os.path.basename(template_path)}' 전용 설정이 저장되었습니다.")
            messagebox.showinfo("저장 완료", f"템플릿 전용 설정이 저장되었습니다.\n{os.path.basename(config_path)}")
        except Exception as e:
            messagebox.showerror("저장 실패", f"설정 저장 중 오류 발생: {e}")


    def save_template_specific_config_silent(self, mode="RT"):
        """현재 설정을 선택된 템플릿 전용 JSON 파일로 조용히(팝업 없이) 저장"""
        if mode == "RT": template_path = self.rt_template_file_path.get()
        elif mode == "KOGAS": template_path = self.kogas_template_file_path.get()
        elif mode == "PT": template_path = self.pt_template_file_path.get()
        elif mode == "MT": template_path = self.mt_template_file_path.get()
        elif mode == "PAUT": template_path = self.paut_template_file_path.get()
        else: template_path = self.template_file_path.get()
        
        if not template_path or not os.path.exists(template_path):
            return
            
        config_path = os.path.splitext(template_path)[0] + ".json"
        
        # Determine relevant keys to save for this template
        keys_to_save = []
        for ctx in ["COVER", "DATA"]:
            for prefix in ["SITCO", "SEOUL", "FOOTER", "FOOTER_PT"]:
                for suffix in ["_PATH", "_ANCHOR", "_W", "_H", "_X", "_Y"]:
                    keys_to_save.append(f"{prefix}_{mode}_{ctx}{suffix}")
            for suffix in ["_TOP", "_BOTTOM", "_LEFT", "_RIGHT"]:
                keys_to_save.append(f"MARGIN_{mode}_{ctx}{suffix}")
            keys_to_save.append(f"PRINT_SCALE_{mode}_{ctx}")
            keys_to_save.append(f"PRINT_AREA_{mode}_{ctx}")
            
        # 2. Mode-specific boundaries and column mappings
        if mode == "RT":
            keys_to_save += ["RT_START_ROW", "RT_END_ROW"]
            keys_to_save += ["RT_FILM_BRAND", "RT_FILM_TYPE", "RT_FILM_SIZE",
                             "RT_ISOTOPE", "RT_ACTIVITY", "RT_ACTIVITY_UNIT"]
            keys_to_save += [k for k in self.config.keys() if k.startswith("RT_COL_") or k.startswith("RT_NAME_")]
        elif mode == "KOGAS":
            keys_to_save += ["KOGAS_START_ROW", "KOGAS_DATA_END_ROW", "RT_KOGAS_D_START_COL"]
            keys_to_save += [k for k in self.config.keys() if k.startswith("KOGAS_")]
        elif mode == "PT":
            keys_to_save += ["PT_START_ROW", "PT_END_ROW"]
            keys_to_save += [k for k in self.config.keys() if k.startswith("PT_COL_") or k.startswith("PT_NAME_")]
        elif mode == "MT":
            keys_to_save += [k for k in self.config if k.startswith("MT_")]
        elif mode == "PAUT":
            keys_to_save += ["PAUT_START_ROW", "PAUT_END_ROW"]
            keys_to_save += [k for k in self.config.keys() if k.startswith("PAUT_COL_") or k.startswith("PAUT_NAME_")]
        elif mode == "PMI":
            keys_to_save += ["PMI_START_ROW", "PMI_DATA_END_ROW", "PMI_PRINT_END_ROW", "PMI_LAYOUT_MODE"]
            keys_to_save += [k for k in self.config.keys() if k.startswith("PMI_COL_") or k.startswith("PMI_NAME_")]

        # Build template-specific dict
        t_data = {k: self.config[k] for k in keys_to_save if k in self.config}
        
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(t_data, f, indent=4, ensure_ascii=False)
        except Exception as e:
            print(f"[ERROR] 템플릿 전용 설정 자동 저장 실패: {e}")


    def save_settings(self, silent=True):
        """현재 설정을 파일(JSON)에 저장"""
        if not getattr(self, '_is_ready', True):
            return True, "초기화 중 저장 생략"
        self._store_report_info_for_mode(getattr(self, 'current_mode', ''))
        self.capture_ui_state()
        try:
            if hasattr(self, 'setting_vars'):
                for key, var in self.setting_vars.items():
                    val = var.get()
                    try:
                        # [FIX] Handle various key types correctly for persistence
                        if "AREA" in key or "_PATH" in key or "_NAME" in key or ("_COL_" in key and "WIDTH" not in key):
                            self.config[key] = str(val).strip()
                        elif any(key.endswith(x) for x in ['_X', '_Y', '_W', '_H']) or any(x in key for x in ['MARGIN', 'HEIGHT', 'WIDTH']):
                            try: self.config[key] = float(val) if str(val).strip() else 0.0
                            except: self.config[key] = 0.0
                        elif 'SCALE' in key or key.endswith(('_ROW', '_COL', '_IDX')):
                            # Column indexes and row numbers should be handled as strings if they are like 'A', 
                            # but for safety in logic we often use col_to_num. 
                            # Here we store them as strings to preserve 'A', 'B' etc.
                            self.config[key] = str(val).strip()
                        else:
                            self.config[key] = str(val)
                    except Exception as e:
                        self.config[key] = str(val)

            # [NEW] 필터 설정 저장 (PMI)
            if hasattr(self, 'element_filters'):
                filter_data = []
                for f_item in self.element_filters:
                    filter_data.append({
                        'key': f_item['key'].get(),
                        'min': f_item['min'].get(),
                        'max': f_item['max'].get()
                    })
                self.config["PMI_FILTERS"] = filter_data

            # [NEW] 탭별 전용 행 설정 강제 저장
            self.config["PMI_START_ROW"] = self.config.get("START_ROW", 19)
            self.config["PMI_DATA_END_ROW"] = self.config.get("DATA_END_ROW", 47)
            self.config["PMI_PRINT_END_ROW"] = self.config.get("PRINT_END_ROW", 47)

            # [NEW] Window State Capture
            if not getattr(self, 'is_embedded', False) and hasattr(self.root, 'geometry'):
                self.config['WINDOW_GEOMETRY'] = self.root.geometry()
            if not getattr(self, 'is_embedded', False) and hasattr(self.root, 'state'):
                self.config['WINDOW_STATE'] = self.root.state()
            try: self.config['ACTIVE_TAB'] = self.mode_notebook.index("current")
            except: pass
            # [FIX] Ensure the config directory exists (Prevents Errno 2)
            os.makedirs(CONFIG_DIR, exist_ok=True)

            with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4, default=str)

            # [NEW] 모든 모드의 템플릿 전용 설정 파일도 함께 동기화(자동 저장)하여 값 덮어쓰기 방지
            for m in ["PMI", "RT", "KOGAS", "PT", "MT", "PAUT"]:
                self.save_template_specific_config_silent(m)

            if not silent:
                self.log("[SUCCESS] 모든 설정이 파일에 안전하게 저장되었습니다.")
            return True, "성공"
            return True, "성공"
        except Exception as e:
            import traceback
            err_msg = traceback.format_exc()
            self.log(f"[ERROR] 설정 저장 실패 상세:\n{err_msg}")
            return False, str(e)

    def manual_save_settings(self):
        """사용자가 수동으로 버튼을 눌러 저장할 때 알림 표시"""
        success, msg = self.save_settings(silent=False)
        if success:
            messagebox.showinfo("성공", "현재 설정이 안전하게 저장되었습니다.")
        else:
            messagebox.showerror("오류", f"설정 저장 중 문제가 발생했습니다.\n\n원인: {msg}")


    def load_daily_usage_from_central(self):
        """중앙지사 앱의 일보 데이터(daily_usage_df)로부터 데이터를 추출하여 현재 보고서에 채움"""
        import pandas as pd
        
        target_date = None
        if hasattr(self, 'central_app') and self.central_app:
            if hasattr(self.central_app, 'selected_date') and self.central_app.selected_date:
                target_date = str(self.central_app.selected_date).strip()
            elif hasattr(self.central_app, 'date_entry'):
                try:
                    target_date = str(self.central_app.date_entry.get()).strip()
                except Exception:
                    pass

        if not target_date:
            target_date = datetime.datetime.now().strftime("%Y-%m-%d")

        input_date = simpledialog.askstring(
            "중앙지사 일보 데이터 불러오기",
            "불러올 작업 일자를 입력하세요 (YYYY-MM-DD):",
            initialvalue=target_date,
            parent=self.root
        )
        if not input_date:
            return
        target_date = input_date.strip()

        df = None
        if hasattr(self, 'central_app') and self.central_app and hasattr(self.central_app, 'daily_usage_df'):
            df = self.central_app.daily_usage_df
        else:
            # Fallback: Material_Inventory.xlsx 직접 로드
            try:
                inv_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "data", "Material_Inventory.xlsx")
                if os.path.exists(inv_path):
                    df = pd.read_excel(inv_path, sheet_name="DailyUsage")
            except Exception as e:
                messagebox.showerror("오류", f"재고 데이터 파일을 읽을 수 없습니다: {e}", parent=self.root)
                return

        if df is None or df.empty:
            messagebox.showwarning("알림", "불러올 일보 데이터가 없습니다.", parent=self.root)
            return

        date_col = 'Date' if 'Date' in df.columns else df.columns[0]
        # 날짜 필터링 (다양한 날짜 포맷 허용)
        df_filtered = df[df[date_col].astype(str).str.contains(target_date, na=False)]
        if df_filtered.empty:
            # YYYYMMDD 형태도 확인
            compact_date = target_date.replace('-', '').replace('.', '')
            df_filtered = df[df[date_col].astype(str).str.contains(compact_date, na=False)]

        if df_filtered.empty:
            messagebox.showinfo("알림", f"{target_date} 일자의 일보 데이터가 없습니다.", parent=self.root)
            return

        # 기본 정보 설정
        first_row = df_filtered.iloc[0]
        self.gapji_exam_date.set(target_date)
        cust = str(first_row.get('업체명') or first_row.get('Site') or '한국지역난방공사 중앙지사')
        if cust and cust != 'nan':
            self.gapji_customer.set(cust)
        proj = str(first_row.get('구분') or first_row.get('Site') or '열수송관 개체공사')
        if proj and proj != 'nan':
            self.gapji_project.set(proj)
        item = str(first_row.get('검사품명') or 'Pipe')
        if item and item != 'nan':
            self.gapji_item.set(item)

        # 현재 활성화된 모드에 데이터 추가
        current_mode = getattr(self, 'current_mode', 'PMI')
        count = 0

        for idx, row in df_filtered.iterrows():
            m_method = str(row.get('검사방법') or '').strip().upper()
            welder = str(row.get('검사자') or row.get('User') or '').strip()
            size = str(row.get('관경(Inch)') or '').strip()
            joints = row.get('조인트수') or 1
            try:
                num_joints = int(joints)
            except:
                num_joints = 1

            for j_idx in range(num_joints):
                joint_name = f"W-{j_idx+1:02d}" if num_joints > 1 else "W-01"
                count += 1
                item_dict = {
                    "selected": True,
                    "No": str(count),
                    "Date": target_date,
                    "Report No": str(row.get('성적서번호') or ''),
                    "Dwg": str(row.get('구분') or '-'),
                    "Joint": joint_name,
                    "Loc": str(row.get('Site') or '중앙지사'),
                    "Size": size,
                    "T": "8.0",
                    "Mat": str(row.get('Material') or 'SPW'),
                    "Result": "ACC",
                    "Welder": welder,
                    "Remarks": str(row.get('Remark') or '')
                }

                if current_mode == 'RT':
                    self.rt_extracted_data.append(item_dict)
                elif current_mode == 'MT':
                    self.mt_extracted_data.append(item_dict)
                elif current_mode == 'PT':
                    self.pt_extracted_data.append(item_dict)
                elif current_mode == 'PAUT':
                    item_dict.update({
                        "Line No.": str(row.get('구분') or '-'),
                        "Joint No.": joint_name,
                        "Th'k(mm)": "8.0",
                        "Acc": "O",
                        "Rej": "-",
                        "Evaluation": "ACC"
                    })
                    self.paut_extracted_data.append(item_dict)
                else:
                    self.extracted_data.append(item_dict)

        # 뷰 새로고침
        if hasattr(self, '_update_preview_table'):
            try:
                self._update_preview_table(current_mode)
            except Exception:
                pass
        
        self.log(f"[INFO] {target_date} 일보 데이터 {count}건 불러오기 완료 ({current_mode} 모드)")
        messagebox.showinfo("완료", f"{target_date} 일자의 일보 데이터 {count}건을 성공적으로 불러왔습니다!\n현재 모드: {current_mode}", parent=self.root)

    def on_closing(self):
        """애플리케이션 종료 시 최종 상태 저장"""
        try:
            self.capture_ui_state()
            self.save_settings()
        except: pass
        if not getattr(self, "is_embedded", False):
            try:
                self.root.destroy()
            except: pass

    def _apply_sash_ratio(self, mode):
        """Helper to apply saved sash ratio with a slight delay to ensure UI mapping."""
        pw = getattr(self, f"{mode.lower()}_paned", None)
        ratio = self.config.get(f"{mode}_SASH_RATIO")
        if pw and ratio is not None:
            def apply():
                total_w = pw.winfo_width()
                if total_w > 100:
                    pw.sash_place(0, int(total_w * ratio), 0)
                else:
                    self.root.after(100, apply)
            self.root.after(100, apply)

    def evaluate_paut_flaw(self, t, h, l, depth, flaw_nature, mode="ECA"):
        """
        PAUT Acceptance Criteria Logic
        - Standard (ASME B31.1): Fatal flaws (Crack, LOF, IP) are rejected immediately.
        - ECA (Fracture Mechanics): Only Crack is fatal. LOF/IP are evaluated by size.
        """
        try:
            t_val = float(t)
            h_val = float(h)
            l_val = float(l)
            d_val = float(depth)
        except (ValueError, TypeError):
            return "Error (Invalid Dimension)", "Unknown"

        # 0. Automatic Location Determination
        s_top = d_val
        s_bottom = t_val - (d_val + h_val)
        s = min(s_top, s_bottom)
        s_limit = 0.4 * (h_val / 2)
        
        loc = "Surface" if s <= s_limit else "Subsurface"

        # 1. Immediate Rejection (Crack)
        nature_str = str(flaw_nature).strip().lower()
        if 'crack' in nature_str or '균열' in nature_str:
            return "Reject (Crack)", loc
            
        # [NEW] Mode-based Rejection for LOF (LF) and IP
        if mode == "Standard":
            unacceptable_types = ['lof', 'lack of fusion', 'ip', 'incomplete penetration', 'lf']
            if any(x in nature_str for x in unacceptable_types):
                return f"Reject ({flaw_nature})", loc
        
        if t_val <= 0:
            return "Error (Zero/Negative Th'k)", "-"
            
        if l_val <= 0 and h_val <= 0:
            return "Accept", "-"
            
        if l_val <= 0 or h_val <= 0:
            return "Error (Zero/Negative Value)", loc

        # 1.1 Special Rules for 6mm <= t < 13mm
        if 6 <= t_val < 13:
            if l_val > 6.4: return f"Reject (L: {l_val} > 6.4mm)", loc
            if t_val < 10: h_surf_max, h_sub_max = 0.95, 0.96
            elif t_val < 12: h_surf_max, h_sub_max = 1.04, 1.04
            else: h_surf_max, h_sub_max = 1.13, 1.14
                
            limit = h_surf_max if loc == "Surface" else h_sub_max
            if h_val > limit: return f"Reject ({loc} h: {h_val} > {limit}mm)", loc
            return f"Accept (h: {h_val} <= {limit}mm)", loc

        # 1.2 Special Rules for 13mm <= t < 25.4mm
        if 13 <= t_val < 25.4:
            if l_val > 6.4: return f"Reject (L: {l_val} > 6.4mm)", loc
            actual_h_t = h_val / t_val
            allowed_h_t = 0.087 if loc == "Surface" else 0.143 
            if actual_h_t > allowed_h_t: return f"Reject ({loc} h/t: {actual_h_t:.3f} > {allowed_h_t:.3f})", loc
            return f"Accept (h/t: {actual_h_t:.3f} <= {allowed_h_t:.3f})", loc
                
        # 2. Aspect Ratio (a/l) logic for t >= 25.4mm
        a_val = h_val if loc == "Surface" else h_val / 2
        aspect_ratio_a_l = a_val / l_val
        
        master_table = [
            (0.00, 0.031, 0.034), (0.05, 0.033, 0.038), (0.10, 0.036, 0.043),
            (0.15, 0.041, 0.054), (0.20, 0.047, 0.066), (0.25, 0.055, 0.078),
            (0.30, 0.064, 0.090), (0.35, 0.074, 0.103), (0.40, 0.083, 0.116),
            (0.45, 0.085, 0.129), (0.50, 0.087, 0.143)
        ]
        
        allowed_a_t = 0
        for ar_limit, surf_a_t, sub_a_t in master_table:
            if aspect_ratio_a_l <= ar_limit:
                allowed_a_t = surf_a_t if loc == 'Surface' else sub_a_t
                break
        if allowed_a_t == 0: allowed_a_t = master_table[-1][1] if loc == 'Surface' else master_table[-1][2]

        actual_a_t = a_val / t_val
        if actual_a_t <= allowed_a_t: return f"Accept (a/t: {actual_a_t:.3f} <= {allowed_a_t:.3f})", loc
        return f"Reject ({loc} a/t: {actual_a_t:.3f} > {allowed_a_t:.3f})", loc

    _LAZY_SUBTAB_ATTRS = {
        'preview_tree': 0, 'pmi_paned': 0, 'pmi_tab_notebook': 0, 'filter_container': 0,
        'rt_preview_tree': 1, 'kogas_preview_tree': 1, 'rt_paned': 1, 'rt_tab_nb': 1, 'rt_preview_nb': 1, 'rt_sub_mode': 1,
        'pt_preview_tree': 2, 'pt_paned': 2, 'pt_tab_notebook': 2,
        'mt_preview_tree': 3, 'mt_paned': 3, 'mt_tab_notebook': 3,
        'paut_preview_tree': 4, 'paut_paned': 4, 'paut_tab_notebook': 4, 'paut_equip_vars': 4, 'paut_photo_pages_var': 4,
        'photo_listbox': 5, 'photo_paned': 5, 'photo_preview_canvas': 5, 'photo_start_page': 5, 'photo_total_pages': 5,
    }

    def __getattr__(self, name):
        tab_idx = self._LAZY_SUBTAB_ATTRS.get(name)
        if tab_idx is not None and '_subtab_builders' in self.__dict__:
            if tab_idx not in self.__dict__.get('_built_subtabs', set()):
                self._ensure_subtab_built(tab_idx)
                if name in self.__dict__:
                    return self.__dict__[name]
        raise AttributeError(f"'{type(self).__name__}' object has no attribute '{name}'")

    def _ensure_subtab_built(self, tab_idx):
        if not hasattr(self, '_subtab_builders') or tab_idx in self._built_subtabs:
            return
        builder = self._subtab_builders.get(tab_idx)
        if not builder:
            return
        self._built_subtabs.add(tab_idx)
        prev_ready = getattr(self, '_is_ready', False)
        self._is_ready = False
        try:
            builder()
        finally:
            self._is_ready = prev_ready

    def _on_main_tab_changed(self, event):
        try:
            # 탭 전환 시 메모리 내 모드별 리포트 정보만 즉시 동기화 (불필요한 8개 JSON 디스크 쓰기 생략)
            self._store_report_info_for_mode(getattr(self, 'current_mode', ''))
        except: pass
        try:
            selected_tab = self.mode_notebook.select()
            tab_idx = self.mode_notebook.index(selected_tab)
            self._ensure_subtab_built(tab_idx)
            tab_text = self.mode_notebook.tab(selected_tab, "text")
            if "PMI" in tab_text: self.current_mode = "PMI"
            elif "RT" in tab_text:
                # RT 탭 진입 시 현재 설정된 서브 모드(표준/가스공사) 확인
                self.current_mode = self.rt_sub_mode.get()
            elif "PT" in tab_text: self.current_mode = "PT"
            elif "MT" in tab_text: self.current_mode = "MT"
            elif "PAUT" in tab_text: self.current_mode = "PAUT"
            elif "사진" in tab_text:
                self.current_mode = "PHOTO"
                self._sync_photo_report_info()

            if self.current_mode in self._report_info_modes():
                self._load_report_info_for_mode(self.current_mode)
            
            # [NEW] 모드 변경 시 해당 모드의 설정을 다시 로드
            if self.current_mode in ["RT", "KOGAS"]:
                t_path = self.rt_template_file_path.get() if self.current_mode == "RT" else self.kogas_template_file_path.get()
                self.load_template_specific_config(t_path, self.current_mode)
        except: pass

    @staticmethod
    def _report_info_modes():
        return ("PMI", "RT", "KOGAS", "PT", "MT", "PAUT")

    def _initialize_report_info_by_mode(self):
        """기존 공통정보를 마이그레이션하여 검사방법별 정보로 준비한다."""
        common_values = {
            'project': str(self.config.get('GAPJI_PROJECT', '')).strip(),
            'customer': str(self.config.get('GAPJI_CUSTOMER', '')).strip(),
            'item': str(self.config.get('GAPJI_ITEM', '')).strip(),
            'material': str(self.config.get('GAPJI_MATERIAL', '')).strip(),
            'report_no': str(self.config.get('GAPJI_REPORT_NO', '')).strip(),
            'exam_date': str(self.config.get('GAPJI_EXAM_DATE', '')).strip(),
        }
        self.report_info_by_mode = {}
        for mode in self._report_info_modes():
            self.report_info_by_mode[mode] = {
                key: str(self.config.get(
                    f'{mode}_{key.upper()}', common_value
                )).strip()
                for key, common_value in common_values.items()
            }
        if self.report_info_by_mode['MT']['item'].upper() == 'HOT TEPPING':
            self.report_info_by_mode['MT']['item'] = 'HOT TAPPING'
            self.config['MT_ITEM'] = 'HOT TAPPING'
        mt_material = {'PIPE': 'SPPS380', 'HOT TAPPING': 'SPPS250'}.get(self.report_info_by_mode['MT']['item'].upper())
        if mt_material:
            self.report_info_by_mode['MT']['material'] = mt_material
            self.config['MT_MATERIAL'] = mt_material
        self._report_info_loading = False

    def _store_report_info_for_mode(self, mode):
        if (
            mode not in self._report_info_modes()
            or not hasattr(self, 'report_info_by_mode')
            or getattr(self, '_report_info_loading', False)
        ):
            return
        info = {
            'project': self.gapji_project.get().strip(),
            'customer': self.gapji_customer.get().strip(),
            'item': self.gapji_item.get().strip(),
            'material': self.gapji_material.get().strip(),
            'report_no': self.gapji_report_no.get().strip(),
            'exam_date': self.gapji_exam_date.get().strip(),
        }
        self.report_info_by_mode[mode] = info
        for key, value in info.items():
            self.config[f'{mode}_{key.upper()}'] = value
        # 현재 탭의 성적서 생성 함수가 사용하는 공통 키도 함께 맞춘다.
        for key, value in info.items():
            self.config[f'GAPJI_{key.upper()}'] = value

    def _load_report_info_for_mode(self, mode):
        if mode not in self._report_info_modes() or not hasattr(self, 'report_info_by_mode'):
            return
        info = self.report_info_by_mode[mode]
        self._report_info_loading = True
        try:
            self.gapji_project.set(info['project'])
            self.gapji_customer.set(info['customer'])
            self.gapji_item.set(info['item'])
            self.gapji_material.set(info['material'])
            self.gapji_report_no.set(info['report_no'])
            self.gapji_exam_date.set(info['exam_date'])
            for key, value in info.items():
                self.config[f'GAPJI_{key.upper()}'] = value
        finally:
            self._report_info_loading = False

    def _on_report_info_changed(self):
        if getattr(self, '_report_info_loading', False):
            return
        if getattr(self, 'current_mode', '') == 'MT':
            material = {'PIPE': 'SPPS380', 'HOT TAPPING': 'SPPS250'}.get(self.gapji_item.get().strip().upper())
            if material and self.gapji_material.get() != material:
                self.gapji_material.set(material)
        self._store_report_info_for_mode(getattr(self, 'current_mode', ''))
        self._sync_photo_report_info()

    def _apply_mt_item_selection(self):
        item_name = self.gapji_item.get().strip()
        material = {'PIPE': 'SPPS380', 'HOT TAPPING': 'SPPS250'}.get(item_name)
        if not material:
            return
        self.gapji_material.set(material)
        for item in self.mt_extracted_data:
            if (item.get('selected', True) and item.get('date_filtered', True)
                    and item.get('drawing_filtered', True)):
                item['TestItem'] = item_name
                item['Material'] = material
        self.save_settings()
        self.populate_preview(self.mt_extracted_data, switch_tab=False, mode='MT')

    def _sync_photo_report_info(self):
        """사진대장에서 선택한 검사방법의 저장정보를 불러온다."""
        if not hasattr(self, 'photo_orderer'):
            return
        self.photo_orderer.set("한국지역난방공사 중앙지사")
        mode = self.photo_inspect_type.get().strip().upper()
        info = getattr(self, 'report_info_by_mode', {}).get(mode)
        if info:
            report_no = info['report_no'].strip()
            self.photo_report_no.set(f"Photo {report_no}" if report_no else "")
            self.photo_inspect_date.set(info['exam_date'])

    def _sync_photo_output_name(self):
        number = self.photo_report_no.get().strip()
        if not number:
            self.photo_output_name.set('Photo.xlsx')
            return
        if not re.match(r'^Photo(?:\s|$)', number, re.I):
            number = 'Photo ' + number
        number = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '-', number).rstrip(' .')
        self.photo_output_name.set(number[:180].rstrip(' .') + '.xlsx')

    def create_widgets(self):
        style = ttk.Style()
        if not getattr(self, 'is_embedded', False):
            style.theme_use('clam')
            style.configure("TFrame", background="#f9fafb")
            style.configure("TLabel", background="#f9fafb", font=("Malgun Gothic", 10))
            style.configure("TLabelframe", background="#f9fafb", font=("Malgun Gothic", 10, "bold"))
            style.configure("TLabelframe.Label", background="#f9fafb", font=("Malgun Gothic", 10, "bold"))
            style.configure("Action.TButton", font=("Malgun Gothic", 11, "bold"), padding=10)
            style.map("TEntry", 
                      selectbackground=[('focus', '#3b82f6'), ('!focus', '#3b82f6')],
                      selectforeground=[('focus', 'white'), ('!focus', 'white')])
            style.configure("Treeview", rowheight=28, font=("Malgun Gothic", 10))
            style.map("Treeview", background=[('selected', '#3b82f6')], foreground=[('selected', 'white')])
            style.configure("Treeview.Heading", font=("Malgun Gothic", 10, "bold"))
        
        # [NEW] 탭 스타일 설정
        style.configure("Main.TNotebook.Tab", font=("Malgun Gothic", 11, "bold"), padding=[10, 3])

        # --- Main Scrollable Container ---
        self.canvas = tk.Canvas(self.root, background="#f9fafb", highlightthickness=0, yscrollincrement=40)
        self.scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, background="#f9fafb")

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self._safe_update_scrollregion()
        )

        self.canvas_frame_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        def _on_canvas_configure(event):
            # [ROOT CAUSE FIX] Always force width, handle None bbox gracefully
            self.canvas.itemconfigure(self.canvas_frame_window, width=event.width)
            self._safe_update_scrollregion()

        self.canvas.bind("<Configure>", _on_canvas_configure)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        self.root.option_add('*selectBackground', '#3b82f6')
        self.root.option_add('*selectForeground', 'white')

        self._create_entry_context_menu()
        self._create_tree_context_menu()

        # Mouse Wheel Support
        def _on_mousewheel(event):
            try:
                focused = self.root.focus_get()
                curr = focused
                while curr:
                    try:
                        c_name = curr.winfo_class()
                        if c_name in ('Treeview', 'Text', 'Listbox'):
                            curr.yview_scroll(int(-1*(event.delta/120)*4.0), "units")
                            return "break"
                    except: pass
                    curr = curr.master
            except Exception: pass
            self.canvas.yview_scroll(int(-1*(event.delta/120)*2.0), "units")
            return "break"
        self.root.bind_all("<MouseWheel>", _on_mousewheel, add="+")

        def _on_root_click(event):
            try:
                if "treeview" in str(event.widget).lower(): return
                w_class = event.widget.winfo_class()
                if w_class in ('Frame', 'TFrame', 'Label', 'TLabel', 'Canvas', 'Labelframe', 'TLabelframe'):
                    self.root.focus_set()
            except: pass
        self.root.bind_all("<Button-1>", _on_root_click, add="+")

        # [ROOT CAUSE FIX] Place Notebook DIRECTLY in root, NOT in Canvas scrollable_frame.
        # Canvas windows do not reliably propagate width to children.
        # Since PMI uses splitscreen (no scrolling needed), direct packing guarantees full width.
        
        # Hide Canvas & Scrollbar (no longer needed for layout)
        self.canvas.pack_forget()
        self.scrollbar.pack_forget()
        
        # ===== PACK ORDER: bottom-most items first =====
        
        # [UX] ESC Key Binding: Unfocus only (preserving content)
        self.root.bind_class("TEntry", "<Escape>", self._on_entry_esc)
        self.root.bind_class("Entry", "<Escape>", self._on_entry_esc)
        
        # 1. Status bar (very bottom)
        self.status_bar = tk.Frame(self.root, background="#e2e8f0", height=22)
        self.status_bar.pack(fill='x', side='bottom')
        self.status_bar.pack_propagate(False)
        tk.Label(self.status_bar, text="준비됨", font=("Malgun Gothic", 8), background="#e2e8f0", foreground="#64748b").pack(side='left', padx=10)
        tk.Label(self.status_bar, text=f"Build Version: v{APP_VERSION}", font=("Arial", 8), background="#e2e8f0", foreground="#94a3b8").pack(side='right', padx=10)

        # 2. Progress bar + Log (above status bar, compact)
        bottom_frame = tk.Frame(self.root, background="#f9fafb")
        bottom_frame.pack(fill='x', side='bottom', padx=5)

        self.progress = ttk.Progressbar(bottom_frame, orient="horizontal", mode="determinate")
        self.progress.pack(fill='x', pady=(5, 2))

        log_frame = tk.Frame(bottom_frame, background="#f9fafb")
        log_frame.pack(fill='x')

        self.status_log = tk.Text(log_frame, height=4, font=("Consolas", 8), state='disabled', background="#1e1e2e", foreground="#10b981", padx=5, pady=3)
        vsb = ttk.Scrollbar(log_frame, orient="vertical", command=self.status_log.yview)
        self.status_log.configure(yscrollcommand=vsb.set)
        self.status_log.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')
        

        
        # 3. THEN pack Notebook (fills all remaining space)
        self.mode_notebook = ttk.Notebook(self.root, style="Main.TNotebook")
        self.mode_notebook.pack(fill='both', expand=True, padx=5, pady=(5, 0))
        
        # [NEW] Track mode change
        self.mode_notebook.bind("<<NotebookTabChanged>>", self._on_main_tab_changed)
        
        self.pmi_mode_frame = tk.Frame(self.mode_notebook, background="#f9fafb")
        self.rt_mode_frame = tk.Frame(self.mode_notebook, background="#f9fafb")
        self.pt_mode_frame = tk.Frame(self.mode_notebook, background="#f9fafb")
        self.paut_mode_frame = tk.Frame(self.mode_notebook, background="#f9fafb")
        self.photo_mode_frame = tk.Frame(self.mode_notebook, background="#f9fafb")

        self.mode_notebook.add(self.pmi_mode_frame, text=" 🔬 PMI (OES) ")
        self.mode_notebook.add(self.rt_mode_frame, text=" 🔬 RT (Standard) ")
        

        
        self.pt_mode_frame = tk.Frame(self.mode_notebook, background="#f9fafb")
        self.mode_notebook.add(self.pt_mode_frame, text=" 🔬 Penetrant (PT) ")
        self.mt_mode_frame = tk.Frame(self.mode_notebook, background="#f9fafb")
        self.mode_notebook.add(self.mt_mode_frame, text=" 🔬 Magnetic (MT) ")
        self.mode_notebook.add(self.paut_mode_frame, text=" 🔬 Phased Array (PAUT) ")
        self.mode_notebook.add(self.photo_mode_frame, text=" 📸 사진대장 (Photo Log) ")

        # Setup sub-tabs lazily: only build the initially active tab upfront
        self._built_subtabs = set()
        self._subtab_builders = {
            0: lambda: self._setup_pmi_ui(self.pmi_mode_frame),
            1: lambda: self._setup_rt_ui(self.rt_mode_frame),
            2: lambda: self._setup_pt_ui(self.pt_mode_frame),
            3: lambda: self._setup_mt_ui(self.mt_mode_frame),
            4: lambda: self._setup_paut_ui(self.paut_mode_frame),
            5: lambda: self._setup_photo_log_ui(self.photo_mode_frame),
        }
        init_tab = getattr(self, '_initial_active_tab', 0)
        if not isinstance(init_tab, int) or init_tab not in self._subtab_builders:
            init_tab = 0
        self._ensure_subtab_built(init_tab)
        init_mode_map = {0: "PMI", 1: "RT", 2: "PT", 3: "MT", 4: "PAUT", 5: "PHOTO"}
        self.current_mode = init_mode_map.get(init_tab, "PMI")
        if self.current_mode in self._report_info_modes():
            self._load_report_info_for_mode(self.current_mode)
        elif self.current_mode == "PHOTO":
            self._sync_photo_report_info()
        if init_tab != 0:
            try:
                self.mode_notebook.select(init_tab)
            except Exception:
                pass

    def _create_scrollable_sidebar(self, parent):
        """Creates a scrollable canvas/scrollbar container for sidebars."""
        canvas = tk.Canvas(parent, background="#f9fafb", highlightthickness=0, borderwidth=0)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, background="#f9fafb", padx=10, pady=0, highlightthickness=0, borderwidth=0)

        def _update_scrollregion(e):
            canvas.configure(scrollregion=canvas.bbox("all"))
            # Auto-hide scrollbar if content fits
            if canvas.bbox("all")[3] <= canvas.winfo_height():
                scrollbar.pack_forget()
            else:
                scrollbar.pack(side="right", fill="y")
        
        scrollable_frame.bind("<Configure>", _update_scrollregion)

        canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        def _on_canvas_configure(e):
            canvas.itemconfig(canvas_window, width=e.width)
        canvas.bind("<Configure>", _on_canvas_configure)
        
        canvas.configure(yscrollcommand=scrollbar.set)
        
        def _on_mousewheel(event):
            # Only allow scrolling if content actually exceeds the visible area
            if canvas.bbox("all")[3] > canvas.winfo_height():
                canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            return "break"
        
        canvas.bind('<Enter>', lambda e: canvas.bind_all("<MouseWheel>", _on_mousewheel))
        canvas.bind('<Leave>', lambda e: canvas.unbind_all("<MouseWheel>"))

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        return scrollable_frame

    def _setup_pmi_ui(self, parent):
        # [FORCE] Ensure parent (pmi_mode_frame) allows expansion
        container = tk.Frame(parent, background="#f9fafb")
        container.pack(fill='both', expand=True)

        # --- Dual Pane Layout (PanedWindow) ---
        self.pmi_paned = tk.PanedWindow(container, orient='horizontal', background="#d1d5db", 
                            sashwidth=6, sashpad=0, sashrelief='raised', borderwidth=0)
        self.pmi_paned.pack(fill='both', expand=True)

        # [LEFT] Filter/Settings Sidebar
        left_container = tk.Frame(self.pmi_paned, background="#f9fafb", highlightthickness=0, borderwidth=0)
        self.pmi_paned.add(left_container, width=425)

        # Scrollable area (Full size)
        left_pane = self._create_scrollable_sidebar(left_container)

        # FIXED FLOATING Header (Lifts over scrollable area)
        header_frame = tk.Frame(left_container, background="#f9fafb", highlightthickness=0, borderwidth=0)
        header_frame.place(x=0, y=0, relwidth=1, height=40)
        
        inner_header = tk.Frame(header_frame, background="#f9fafb", padx=20)
        inner_header.pack(fill='both', expand=True, pady=(5, 0))
        
        tk.Label(inner_header, text="🔬 PMI 관리", font=("Malgun Gothic", 15, "bold"), 
                 background="#f9fafb", foreground="#1e3a8a").pack(side='left')
        tk.Label(inner_header, text=f"v{APP_VERSION}", font=("Arial", 8), 
                 background="#f1f5f9", foreground="#64748b", padx=5, pady=0).pack(side='left', padx=10)
        
        # Bottom separator
        tk.Frame(header_frame, height=1, background="#e5e7eb").pack(side='bottom', fill='x')
        header_frame.lift() # Ensure it's on top

        # Spacer in scrollable area so first group starts below floating header
        tk.Frame(left_pane, height=40, background="#f9fafb").pack(fill='x')

        # 1. File Selection Group
        file_container = ttk.LabelFrame(left_pane, text=" 데이터 및 양식 (Data) ", padding=10)
        file_container.pack(fill='x', pady=(0, 10))

        def _add_very_compact_row(parent_f, label, var, row, is_dir=False, types=None):
            parent_f.columnconfigure(1, weight=1)
            ttk.Label(parent_f, text=label, font=("Malgun Gothic", 8)).grid(row=row, column=0, sticky='e', padx=2, pady=2)
            # Use width=1 for entry to allow proactive shrinking
            ttk.Entry(parent_f, textvariable=var, font=("Arial", 9), width=1, exportselection=False).grid(row=row, column=1, padx=2, pady=2, sticky='ew')
            cmd = (lambda : self._browse_dir(var)) if is_dir else (lambda : self._browse_file(var, types))
            ttk.Button(parent_f, text="...", width=3, command=cmd).grid(row=row, column=2, padx=2, pady=2)

        _add_very_compact_row(file_container, "데이터:", self.target_file_path, 0, types=[("Excel Source", "*.xls;*.xlsx;*.xlsm")])
        _add_very_compact_row(file_container, "양식:", self.template_file_path, 1, types=[("Excel Template", "*.xlsx;*.xlsm")])
        
        # [NEW] Template Config Save Button
        btn_frame = tk.Frame(file_container, background="#f0f0f0")
        btn_frame.grid(row=2, column=1, sticky='ew', pady=2)
        ttk.Button(btn_frame, text="💾 현재 설정을 이 양식 전용으로 저장", command=lambda: self.save_template_specific_config("PMI")).pack(side='left', padx=2)

        # [NEW] PMI 양식 모드 선택 (일반 1행 vs 모재+용접부+모재 3단)
        layout_mode_frame = tk.LabelFrame(file_container, text=" 양식 모드 (측정 방식) ", background="#f0f0f0", padx=4, pady=2)
        layout_mode_frame.grid(row=3, column=0, columnspan=3, sticky='ew', padx=2, pady=3)
        ttk.Radiobutton(layout_mode_frame, text="일반 단일행 (1포인트 = 1행)", variable=self.pmi_layout_mode, value="STANDARD").pack(side='left', padx=6)
        ttk.Radiobutton(layout_mode_frame, text="플랜트 배관용 (모재+용접부+모재 3단 / 1p당 9포인트)", variable=self.pmi_layout_mode, value="PLANT_TRIPLE").pack(side='left', padx=6)

        # 2. Configuration Notebook (Cover, Data, Rows - NO Preview tab here)
        config_frame = ttk.LabelFrame(left_pane, text=" 설정 (Config) ", padding=2)
        config_frame.pack(fill='both', expand=True, pady=(0, 10))

        self.pmi_tab_notebook = ttk.Notebook(config_frame)
        self.pmi_tab_notebook.pack(fill='both', expand=True)

        tab_cover = ttk.Frame(self.pmi_tab_notebook, padding=5)
        tab_data = ttk.Frame(self.pmi_tab_notebook, padding=5)
        tab_rows = ttk.Frame(self.pmi_tab_notebook, padding=5)
        tab_cols = ttk.Frame(self.pmi_tab_notebook, padding=5)
        tab_filter = ttk.Frame(self.pmi_tab_notebook, padding=10)  # 내부 호환용 위젯 유지 (노트북 탭 미노출)
        
        # [CRITICAL] Allow children to expand within tabs
        for t in [tab_cover, tab_data, tab_rows, tab_cols]: t.columnconfigure(0, weight=1)

        self.pmi_tab_notebook.add(tab_cover, text="갑지")
        self.pmi_tab_notebook.add(tab_data, text="을지")
        self.pmi_tab_notebook.add(tab_rows, text="행 설정")
        self.pmi_tab_notebook.add(tab_cols, text="컬럼 설정")
        # [CLEANUP] 사용 빈도가 낮고 혼선을 주는 레거시 '필터/옵션' 탭은 화면에서 숨김 (재질 자동 판정은 기본 True 유지)
        # self.pmi_tab_notebook.add(tab_filter, text="필터/옵션")
        if hasattr(self, 'auto_verify'):
            self.auto_verify.set(True)

        # [FIX] Do NOT reset setting_vars here to prevent data loss when switching tabs
        # [ALIGNED] Mode-specific context for logo grid
        self._create_gapji_meta_ui(tab_cover, use_pack=False, mode="PMI")
        self.pmi_tab_notebook.bind("<<NotebookTabChanged>>", self._update_gapji_preview_current)
        # PMI logos are placed automatically from the saved configuration/template.
        # Keep the settings and insertion engine, but do not expose manual logo
        # path/position controls in the PMI tabs.
        self._create_margin_settings(tab_cover, "PMI_COVER", use_pack=False)
        self._create_margin_settings(tab_data, "PMI_DATA", use_pack=False)
        self._create_row_settings(tab_rows, mode="PMI")
        
        pmi_items = [
            ("No:", "PMI_COL_NO", 13, "PMI_NAME_NO", "No", "No"),
            ("Drawing No:", "PMI_COL_DWG", 1, "PMI_NAME_DWG", "Drawing No.", "Dwg"),
            ("Joint No:", "PMI_COL_JOINT", 6, "PMI_NAME_JOINT", "Joint No.", "Joint"),
            ("Location:", "PMI_COL_LOC", 7, "PMI_NAME_LOC", "Location", "Loc"),
            ("Ni:", "PMI_COL_NI", 8, "PMI_NAME_NI", "Ni", "Ni"),
            ("Cr:", "PMI_COL_CR", 9, "PMI_NAME_CR", "Cr", "Cr"),
            ("Mo:", "PMI_COL_MO", 10, "PMI_NAME_MO", "Mo", "Mo"),
            ("판정 결과:", "PMI_COL_RES", 13, "PMI_NAME_RES", "Result", "Result")
        ]
        self._create_column_mapping_ui(tab_cols, "PMI", pmi_items)
        
        # Spacer for tabs
        tk.Frame(tab_rows, background="#f9fafb").pack(fill='x')

        # 3. Quick Options Section (Moved to Settings Notebook)
        # Filter Tag Area (Self-updating)
        filter_box = tk.Frame(tab_filter, background="#f9fafb")
        filter_box.pack(fill='x', pady=2)
        ttk.Label(filter_box, text="🔍 성분 필터:", background="#f9fafb", font=("Malgun Gothic", 9, "bold")).pack(side='left', padx=(0, 5))
        
        def add_filter():
            self.element_filters.append({'key': tk.StringVar(value=''), 'min': tk.StringVar(value=''), 'max': tk.StringVar(value='')})
            self._update_pmi_filter_ui()
        ttk.Button(filter_box, text="+", width=2, command=add_filter).pack(side='right')

        self.filter_container = tk.Frame(filter_box, background="#f9fafb")
        self.filter_container.pack(side='left', fill='x', expand=True)

        # Sequence Filter & Auto Check
        tk.Frame(tab_filter, height=1, background="#e5e7eb").pack(fill='x', pady=10)
        tk.Checkbutton(tab_filter, text="✅ 재질 자동 판정", variable=self.auto_verify, background="#f9fafb", font=("Malgun Gothic", 8)).pack(anchor='w', pady=(0, 5))
        seq_row = tk.Frame(tab_filter, background="#f9fafb")
        seq_row.pack(fill='x', pady=2)
        ttk.Label(seq_row, text="📊 특정순번:", background="#f9fafb", font=("Malgun Gothic", 8)).pack(side='left')
        ttk.Entry(seq_row, textvariable=self.sequence_filter, width=15, font=("Arial", 9), exportselection=False).pack(side='left', padx=5, fill='x', expand=True)

        self._update_pmi_filter_ui()

        # [RIGHT] Multi-Preview Pane (Data & Gapji)
        right_container = tk.Frame(self.pmi_paned, background="#f3f4f6")
        self.pmi_paned.add(right_container, stretch="always")

        # Top Action Bar (생성 시작, 데이터 추출, 일보 NDT 불러오기) - 모던 플랫 툴바 스타일
        action_bar = tk.Frame(right_container, background="#ffffff", padx=8, pady=6)
        action_bar.pack(fill='x')

        btn_run = tk.Button(
            action_bar, text="✨ 생성 시작", command=self.run_process,
            font=("Malgun Gothic", 10, "bold"), bg="#2563eb", fg="#ffffff",
            activebackground="#1d4ed8", activeforeground="#ffffff",
            relief="flat", cursor="hand2", padx=12, pady=5, bd=0
        )
        btn_run.pack(side='left', fill='x', expand=True, padx=(0, 6))
        btn_run.bind("<Enter>", lambda e: btn_run.configure(bg="#1d4ed8"))
        btn_run.bind("<Leave>", lambda e: btn_run.configure(bg="#2563eb"))

        btn_extract = tk.Button(
            action_bar, text="📝 데이터 추출", command=self.extract_only,
            font=("Malgun Gothic", 9, "bold"), bg="#f8fafc", fg="#334155",
            activebackground="#e2e8f0", activeforeground="#1e293b",
            relief="solid", bd=1, cursor="hand2", padx=10, pady=5
        )
        btn_extract.pack(side='left', fill='x', expand=True, padx=(0, 6))
        btn_extract.bind("<Enter>", lambda e: btn_extract.configure(bg="#e2e8f0"))
        btn_extract.bind("<Leave>", lambda e: btn_extract.configure(bg="#f8fafc"))

        btn_daily = tk.Button(
            action_bar, text="📋 일보 NDT 불러오기", command=self.load_daily_work_history,
            font=("Malgun Gothic", 9, "bold"), bg="#f8fafc", fg="#334155",
            activebackground="#e2e8f0", activeforeground="#1e293b",
            relief="solid", bd=1, cursor="hand2", padx=10, pady=5
        )
        btn_daily.pack(side='left', fill='x', expand=True)
        btn_daily.bind("<Enter>", lambda e: btn_daily.configure(bg="#e2e8f0"))
        btn_daily.bind("<Leave>", lambda e: btn_daily.configure(bg="#f8fafc"))

        # 툴바와 데이터 화면 사이 구분선
        tk.Frame(right_container, height=1, background="#e2e8f0").pack(fill='x', pady=(0, 4))
        
        # [DATA ONLY] 갑지 탭 제거 -> 100% 데이터 미리보기 전용 화면으로 쾌적화
        tab_data = tk.Frame(right_container, background="#ffffff")
        tab_data.pack(fill='both', expand=True)
        
        self._create_preview_ui(tab_data)
        self._apply_sash_ratio("PMI")
        
        self.template_file_path.trace_add("write", lambda *a: self.load_template_specific_config("PMI"))

    def _update_pmi_ratio(self):
        """Saves current sash position and persists it."""
        try:
            total_w = self.pmi_paned.winfo_width()
            if total_w > 100:
                current_sash = self.pmi_paned.sash_coord(0)[0]
                self.pmi_pane_ratio = current_sash / total_w
                self.config['PMI_SASH_RATIO'] = self.pmi_pane_ratio
                self.save_settings() # Persist immediately
        except: pass

    def _on_pmi_paned_configure(self, event):
        """Maintains the proportional split when the window resizes."""
        try:
            total_w = self.pmi_paned.winfo_width()
            if total_w > 100:
                # During window resize, RE-PLACE sash based on SAVED ratio
                if event and event.widget == self.pmi_paned:
                    new_pos = int(total_w * self.pmi_pane_ratio)
                    self.pmi_paned.sash_place(0, new_pos, 0)
        except: pass

    def _update_pmi_filter_ui(self):
        """Redraws the element filter tags in a modern, elastic style."""
        for widget in self.filter_container.winfo_children():
            widget.destroy()
        
        def remove_filter(idx):
            self.element_filters.pop(idx)
            self._update_pmi_filter_ui()

        # [HYPER-REACTIVE] Use grid for even horizontal distribution (Elastic, Stacked)
        cols_per_row = 1
        for i, f_item in enumerate(self.element_filters):
            self.filter_container.columnconfigure(0, weight=1)
            
            # Create an "Elastic Tag" 
            tag = tk.Frame(self.filter_container, background="#eff6ff", highlightthickness=1, 
                            highlightbackground="#bfdbfe", padx=2, pady=1)
            tag.grid(row=i, column=0, sticky='ew', padx=2, pady=2)
            
            # Inner grid for tag components
            tag.columnconfigure(0, weight=1, minsize=0) # Key
            tag.columnconfigure(2, weight=1, minsize=0) # Min
            tag.columnconfigure(4, weight=1, minsize=0) # Max
            
            # Key
            e_key = tk.Entry(tag, textvariable=f_item['key'], width=1, font=("Arial", 8), 
                             relief='flat', background="#eff6ff", justify='center')
            e_key.grid(row=0, column=0, sticky='ew')
            tk.Label(tag, text=":", background="#eff6ff", font=("Arial", 8)).grid(row=0, column=1)
            
            # Min
            e_min = tk.Entry(tag, textvariable=f_item['min'], width=1, font=("Arial", 8), 
                             relief='flat', background="#eff6ff", justify='center')
            e_min.grid(row=0, column=2, sticky='ew')
            tk.Label(tag, text="~", background="#eff6ff", font=("Arial", 8)).grid(row=0, column=3)
            
            # Max
            e_max = tk.Entry(tag, textvariable=f_item['max'], width=1, font=("Arial", 8), 
                             relief='flat', background="#eff6ff", justify='center')
            e_max.grid(row=0, column=4, sticky='ew')
            
            # Delete button (Small)
            btn_del = tk.Button(tag, text="×", command=lambda idx=i: remove_filter(idx), 
                                relief='flat', background="#eff6ff", foreground="#3b82f6", font=("Arial", 9, "bold"),
                                activebackground="#dbeafe", cursor="hand2", padx=0, pady=0)
            btn_del.grid(row=0, column=5, padx=(2, 0))

    def _setup_rt_ui(self, parent):
        """[RE-INTEGRATED] RT 탭 통합 구성 (내부에 서브 미리보기 탭 배치)"""
        container = tk.Frame(parent, background="#f9fafb")
        container.pack(fill='both', expand=True)

        # --- Dual Pane Layout (PanedWindow) ---
        self.rt_paned = tk.PanedWindow(container, orient='horizontal', background="#d1d5db", 
                            sashwidth=6, sashpad=0, sashrelief='raised', borderwidth=0)
        self.rt_paned.pack(fill='both', expand=True)

        # [LEFT] Settings Sidebar
        left_container = tk.Frame(self.rt_paned, background="#f9fafb", highlightthickness=0, borderwidth=0)
        self.rt_paned.add(left_container, width=425)
        left_pane = self._create_scrollable_sidebar(left_container)

        # FIXED Header
        header_frame = tk.Frame(left_container, background="#f9fafb")
        header_frame.place(x=0, y=0, relwidth=1, height=40)
        tk.Label(header_frame, text="🔬 RT 성적서 통합 관리", font=("Malgun Gothic", 15, "bold"), 
                 background="#f9fafb", foreground="#1e3a8a").pack(side='left', padx=20, pady=5)
        tk.Frame(header_frame, height=1, background="#e5e7eb").pack(side='bottom', fill='x')
        tk.Frame(left_pane, height=40, background="#f9fafb").pack(fill='x')

        # [NEW] 모드 선택 (표준 vs 가스공사)
        mode_select_f = tk.Frame(left_pane, background="#f9fafb")
        mode_select_f.pack(fill='x', pady=(0, 10))
        self.rt_sub_mode = tk.StringVar(value="RT")
        
        style = ttk.Style()
        style.configure("Mode.TRadiobutton", font=("Malgun Gothic", 9, "bold"))
        
        ttk.Radiobutton(mode_select_f, text=" 🔬 표준 RT ", variable=self.rt_sub_mode, value="RT", 
                        style="Mode.TRadiobutton", command=self._on_rt_sub_mode_change).pack(side='left', padx=5)
        ttk.Radiobutton(mode_select_f, text=" 🇰🇷 가스공사 ", variable=self.rt_sub_mode, value="KOGAS", 
                        style="Mode.TRadiobutton", command=self._on_rt_sub_mode_change).pack(side='left', padx=5)

        # 1. File Selection (Adaptive)
        self.rt_file_frame = ttk.LabelFrame(left_pane, text=" 데이터 및 양식 (Data) ", padding=10)
        self.rt_file_frame.pack(fill='x', pady=(0, 10))
        
        # We need to show different path entries based on mode
        self.rt_path_container = tk.Frame(self.rt_file_frame)
        self.rt_path_container.pack(fill='x')
        
        def _add_file_row(p, lbl, var, r, is_dir=False, types=None):
            p.columnconfigure(1, weight=1)
            ttk.Label(p, text=lbl, font=("Malgun Gothic", 8)).grid(row=r, column=0, sticky='e', padx=2, pady=2)
            ttk.Entry(p, textvariable=var, font=("Arial", 9), width=1, exportselection=False).grid(row=r, column=1, padx=2, pady=2, sticky='ew')
            cmd = (lambda: self._browse_dir(var)) if is_dir else (lambda: self._browse_file(var, types))
            ttk.Button(p, text="...", width=3, command=cmd).grid(row=r, column=2, padx=2, pady=2)

        # Container for actual path rows (will be cleared/re-added on mode change)
        self._update_rt_path_ui()

        # 2. Tabs for Configuration
        rt_config_frame = ttk.LabelFrame(left_pane, text=" 세부 설정 (Config) ", padding=2)
        rt_config_frame.pack(fill='both', expand=True, pady=(0, 10))

        self.rt_tab_nb = ttk.Notebook(rt_config_frame)
        self.rt_tab_nb.pack(fill='both', expand=True)

        self.rt_tabs = {}
        for t_name in ["갑지", "을지", "로고", "행 설정", "컬럼 설정"]:
            f = ttk.Frame(self.rt_tab_nb, padding=5)
            self.rt_tab_nb.add(f, text=t_name)
            self.rt_tabs[t_name] = f
            if t_name not in ["행 설정", "컬럼 설정"]: f.columnconfigure(0, weight=1)

        self.rt_tab_nb.bind("<<NotebookTabChanged>>", self._update_gapji_preview_current)
        self._refresh_rt_config_tabs() # Initial UI build

        # 3. Actions
        action_f = tk.Frame(left_pane, background="#ffffff", highlightthickness=1, highlightbackground="#d1d5db", padx=10, pady=5)
        action_f.pack(fill='x', pady=(0, 10))
        
        # Row 1: Sequence Filter
        r1 = tk.Frame(action_f, background="#ffffff")
        r1.pack(fill='x', pady=2)
        ttk.Label(r1, text="📊 특정순번 필터:", background="#ffffff").pack(side='left')
        ttk.Entry(r1, textvariable=self.sequence_filter, width=15).pack(side='left', padx=5, fill='x', expand=True)
        
        # Row 2: Keyword Filter
        r2 = tk.Frame(action_f, background="#ffffff")
        r2.pack(fill='x', pady=2)
        ttk.Label(r2, text="🔍 추출 키워드:", background="#ffffff").pack(side='left')
        ttk.Entry(r2, textvariable=self.rt_extract_keyword, width=15).pack(side='left', padx=5, fill='x', expand=True)

        # Actions moved to right Top Action Bar

        # [RIGHT] Multi-Preview (Sub-tabs)
        right_f = tk.Frame(self.rt_paned, background="#f3f4f6")
        self.rt_paned.add(right_f, stretch="always")
        
        self._create_top_action_bar(right_f, "RT")
        self.rt_preview_nb = ttk.Notebook(right_f)
        self.rt_preview_nb.pack(fill='both', expand=True)
        
        t_std = tk.Frame(self.rt_preview_nb, background="#ffffff")
        t_kogas = tk.Frame(self.rt_preview_nb, background="#ffffff")
        t_gapji = tk.Frame(self.rt_preview_nb, background="#ffffff")
        
        self.rt_preview_nb.add(t_std, text=" 🔬 표준 미리보기 ")
        self.rt_preview_nb.add(t_kogas, text=" 🇰🇷 가스공사 미리보기 ")
        self.rt_preview_nb.add(t_gapji, text=" 📄 갑지 미리보기 ")
        
        # Link Preview Tab to Mode Switch
        def _on_rt_preview_tab_changed(event):
            tab_text = self.rt_preview_nb.tab(self.rt_preview_nb.select(), "text")
            new_mode = None
            if "가스공사" in tab_text: new_mode = "KOGAS"
            elif "표준" in tab_text: new_mode = "RT"
            if new_mode and new_mode != self.rt_sub_mode.get():
                self.rt_sub_mode.set(new_mode)
                self._on_rt_sub_mode_change(from_tab=True)
        self.rt_preview_nb.bind("<<NotebookTabChanged>>", _on_rt_preview_tab_changed)

        self._create_rt_preview_ui(t_std, mode="RT")
        self._create_rt_preview_ui(t_kogas, mode="KOGAS")
        self._create_gapji_preview_ui(t_gapji, "RT") # Gapji uses current sub-mode context internally
        
        self._apply_sash_ratio("RT")
        self.rt_pane_ratio = self.config.get('RT_SASH_RATIO', 0.5)
        self.rt_paned.bind("<Configure>", self._on_rt_paned_configure)
        self.rt_paned.bind("<ButtonRelease-1>", lambda e: self.root.after(10, self._update_rt_ratio))
        self.root.after(500, lambda: self._on_rt_paned_configure(None))

    def _update_rt_path_ui(self):
        """모드에 따라 파일 선택 경로 입력 UI 갱신"""
        for w in self.rt_path_container.winfo_children(): w.destroy()
        mode = self.rt_sub_mode.get()
        v_target = self.rt_target_file_path if mode == "RT" else self.kogas_target_file_path
        v_template = self.rt_template_file_path if mode == "RT" else self.kogas_template_file_path

        def _add_file_row(p, lbl, var, r, is_dir=False, types=None):
            p.columnconfigure(1, weight=1)
            ttk.Label(p, text=lbl, font=("Malgun Gothic", 8)).grid(row=r, column=0, sticky='e', padx=2, pady=2)
            ttk.Entry(p, textvariable=var, font=("Arial", 9), width=1, exportselection=False).grid(row=r, column=1, padx=2, pady=2, sticky='ew')
            cmd = (lambda: self._browse_dir(var)) if is_dir else (lambda: self._browse_file(var, types))
            ttk.Button(p, text="...", width=3, command=cmd).grid(row=r, column=2, padx=2, pady=2)

        _add_file_row(self.rt_path_container, "기본 로고 폴더:", self.logo_folder_path, 0, is_dir=True)
        _add_file_row(self.rt_path_container, f"{mode} 데이터:", v_target, 1, types=[("Excel Source", "*.xls;*.xlsx;*.xlsm")])
        _add_file_row(self.rt_path_container, f"{mode} 양식:", v_template, 2, types=[("Excel Template", "*.xlsx;*.xlsm")])
        
        btn_f = tk.Frame(self.rt_path_container, background="#f0f0f0")
        btn_f.grid(row=3, column=1, sticky='ew', pady=2)
        ttk.Button(btn_f, text="💾 전용 설정 저장", command=lambda: self.save_template_specific_config(mode)).pack(side='left', padx=2)

    def _refresh_rt_config_tabs(self):
        """모드에 따라 행/컬럼 설정 탭 내용 갱신"""
        mode = self.rt_sub_mode.get()
        t_cover = self.rt_tabs["갑지"]
        t_data = self.rt_tabs["을지"]
        t_logo = self.rt_tabs["로고"]
        t_rows = self.rt_tabs["행 설정"]
        t_cols = self.rt_tabs["컬럼 설정"]
        
        for t in [t_rows, t_cols, t_logo, t_cover, t_data]:
            for w in t.winfo_children(): w.destroy()
        
        # Meta & Margins (Cover/Data)
        self._create_gapji_meta_ui(t_cover, use_pack=False, mode="RT")
        ctx_cover = f"{mode}_COVER"
        ctx_data = f"{mode}_DATA"
        self._create_margin_settings(t_cover, ctx_cover, use_pack=False)
        self._create_margin_settings(t_data, ctx_data, use_pack=False)
        
        # Logos
        tk.Label(t_logo, text="🚩 [갑지] 로고", font=("Malgun Gothic", 9, "bold"), background="#f9fafb").grid(row=0, column=0, sticky='w')
        r_next = self._create_setting_grid(t_logo, ctx_cover)
        tk.Label(t_logo, text="🚩 [을지] 로고", font=("Malgun Gothic", 9, "bold"), background="#f9fafb").grid(row=r_next+1, column=0, sticky='w', pady=(10, 0))
        self._create_setting_grid(t_logo, ctx_data)
        
        # Row Settings
        self._create_row_settings(t_rows, mode=mode)
        
        # Column Mapping (Ordered by current preview keys)
        all_possible = {
            "No": ("No:", f"{mode}_COL_NO", 1, f"{mode}_NAME_NO", "No", "No"),
            "Date": ("Date:", f"{mode}_COL_DATE", 2, f"{mode}_NAME_DATE", "Date", "Date"),
            "Sec": ("SEC:", f"{mode}_COL_SEC", 0, f"{mode}_NAME_SEC", "SEC", "Sec"),
            "Dwg": ("Dwg No.:", f"{mode}_COL_DWG", 3, f"{mode}_NAME_DWG", "Drawing No.", "Dwg"),
            "Joint": ("Film Ident. No.:", f"{mode}_COL_JOINT", 4, f"{mode}_NAME_JOINT", "Film Ident. No.", "Joint"),
            "Loc": ("Film Location:", f"{mode}_COL_LOC", 5, f"{mode}_NAME_LOC", "Film Location", "Loc"),
            "T": ("T:", f"{mode}_COL_THK", 6, f"{mode}_NAME_THK", "T", "T"),
            "Mat": ("Mat:", f"{mode}_COL_MAT", 7, f"{mode}_NAME_MAT", "Mat", "Mat"),
            "Size": ("구경(Size):", f"{mode}_COL_SIZE", 0, f"{mode}_NAME_SIZE", "Size", "Size"),
            "Acc": ("Acc:", f"{mode}_COL_ACC", 8, f"{mode}_NAME_ACC", "Acc", "Acc"),
            "Rej": ("Rej:", f"{mode}_COL_REJ", 9, f"{mode}_NAME_REJ", "Rej", "Rej"),
            "Deg": ("Deg:", f"{mode}_COL_DEG", 10, f"{mode}_NAME_DEG", "Deg", "Deg"),
            "D1": ("① Crack:", f"{mode}_COL_D1", 13, f"{mode}_NAME_D1", "① Crack", "D1"),
            "D2": ("② IP:", f"{mode}_COL_D2", 14, f"{mode}_NAME_D2", "② IP", "D2"),
            "D3": ("③ LF:", f"{mode}_COL_D3", 15, f"{mode}_NAME_D3", "③ LF", "D3"),
            "D4": ("④ Slag:", f"{mode}_COL_D4", 16, f"{mode}_NAME_D4", "④ Slag", "D4"),
            "D5": ("⑤ Por:", f"{mode}_COL_D5", 17, f"{mode}_NAME_D5", "⑤ Por", "D5"),
            "⑥ U/C": ("⑥ U/C:", f"{mode}_COL_D6", 18, f"{mode}_NAME_D6", "⑥ U/C", "D6"),
            "D6": ("⑥ U/C:", f"{mode}_COL_D6", 18, f"{mode}_NAME_D6", "⑥ U/C", "D6"),
            "D7": ("⑦ RUC:", f"{mode}_COL_D7", 19, f"{mode}_NAME_D7", "⑦ RUC", "D7"),
            "D8": ("⑧ BT:", f"{mode}_COL_D8", 20, f"{mode}_NAME_D8", "⑧ BT", "D8"),
            "D9": ("⑨ TI:", f"{mode}_COL_D9", 21, f"{mode}_NAME_D9", "⑨ TI", "D9"),
            "D10": ("⑩ CP:", f"{mode}_COL_D10", 22, f"{mode}_NAME_D10", "⑩ CP", "D10"),
            "D11": ("⑪ RC:", f"{mode}_COL_D11", 23, f"{mode}_NAME_D11", "⑪ RC", "D11"),
            "D12": ("⑫ Mis:", f"{mode}_COL_D12", 24, f"{mode}_NAME_D12", "⑫ Mis", "D12"),
            "D13": ("⑬ EP:", f"{mode}_COL_D13", 25, f"{mode}_NAME_D13", "⑬ EP", "D13"),
            "D14": ("⑭ SD:", f"{mode}_COL_D14", 26, f"{mode}_NAME_D14", "⑭ SD", "D14"),
            "D15": ("⑮ Oth:", f"{mode}_COL_D15", 27, f"{mode}_NAME_D15", "⑮ Oth", "D15"),
            "Result": ("판정(Result):", f"{mode}_COL_RES", 28, f"{mode}_NAME_RES", "Result", "Result"),
            "Welder": ("용접사:", f"{mode}_COL_WELDER", 29, f"{mode}_NAME_WELDER", "Welder No", "Welder"),
            "Remarks": ("비고:", f"{mode}_COL_REM", 30, f"{mode}_NAME_REM", "Remarks", "Remarks")
        }
        
        keys_attr = "rt_column_keys" if mode == "RT" else "kogas_column_keys"
        current_keys = getattr(self, keys_attr)
        
        # 1. Start with items that are currently visible/ordered
        rt_items = []
        for k in current_keys:
            if k in all_possible:
                rt_items.append(all_possible[k])
        
        # 2. Add remaining items at the end (so they remain configurable even if hidden)
        for k, v in all_possible.items():
            if v not in rt_items:
                rt_items.append(v)

        if mode == "KOGAS":
            self._create_kogas_separated_column_mapping_ui(t_cols)
        else:
            self._create_column_mapping_ui(t_cols, mode, rt_items)

    def _on_rt_sub_mode_change(self, from_tab=False):
        """RT 내부 서브 모드(표준/가스공사) 변경 시 UI 및 상태 동기화"""
        try:
            # [NEW] 서브 모드 전환 전 기존 설정 내용을 안전하게 자동 저장
            self.save_settings()
        except: pass
        mode = self.rt_sub_mode.get()
        self.current_mode = mode
        self._load_report_info_for_mode(mode)
        self._sub_mode_switching = True  # guard: prevent _on_rt_kogas_mode_change from overwriting paths
        self.rt_kogas_mode.set(mode == "KOGAS")
        self._sub_mode_switching = False
        self.log(f"🔄 RT 서브 모드 전환: {mode}")
        
        # 1. 오른쪽 미리보기 탭 동기화 (라디오 버튼으로 바꾼 경우)
        if not from_tab:
            target_idx = 0 if mode == "RT" else 1
            if self.rt_preview_nb.index("current") != target_idx:
                self.rt_preview_nb.select(target_idx)
        
        # 2. 왼쪽 사이드바 UI 갱신 (파일 경로 및 세부 설정 탭)
        self._update_rt_path_ui()
        self._refresh_rt_config_tabs()
        
        # 3. 템플릿 로드 트리거 [FIX] Pass correct path and avoid forced tab switch
        t_path = self.rt_template_file_path.get() if mode == "RT" else self.kogas_template_file_path.get()
        self.load_template_specific_config(t_path, mode)
        
        # 4. 갑지 미리보기 갱신
        self._update_gapji_preview(mode)

    def _update_rt_ratio(self):
        try:
            total_w = self.rt_paned.winfo_width()
            if total_w > 100:
                current_sash = self.rt_paned.sash_coord(0)[0]
                self.rt_pane_ratio = current_sash / total_w
                self.config['RT_SASH_RATIO'] = self.rt_pane_ratio
                self.save_settings()
        except: pass

    def _on_rt_paned_configure(self, event):
        try:
            total_w = self.rt_paned.winfo_width()
            if total_w > 100:
                if event and event.widget == self.rt_paned:
                    new_pos = int(total_w * self.rt_pane_ratio)
                    self.rt_paned.sash_place(0, new_pos, 0)
        except: pass

    def _on_rt_kogas_mode_change(self):
        """가스공사 모드 변경 시 템플릿 경로 자동 설정 및 미리보기 컬럼 업데이트"""
        # 서브 모드 전환 중이면 경로 자동 설정을 건너뜀 (중복 실행 방지)
        if getattr(self, '_sub_mode_switching', False):
            self._update_rt_preview_columns("KOGAS" if self.rt_kogas_mode.get() else "RT")
            return
        if self.rt_kogas_mode.get():
            # 가스공사 템플릿 경로 탐색
            data_dir = os.path.join(BASE_DIR, "Na-aba", "home", "data")
            if not os.path.exists(data_dir):
                data_dir = os.path.join(os.getcwd(), "Na-aba", "home", "data")
            
            kogas_template = os.path.join(data_dir, "가스공사 의뢰서.xlsx")
            if os.path.exists(kogas_template):
                self.kogas_template_file_path.set(kogas_template)
                self.log("💡 가스공사 전용 템플릿이 자동으로 선택되었습니다.")
            else:
                self.log("⚠️ 가스공사 의뢰서.xlsx 파일을 찾을 수 없습니다. (data 폴더 확인 필요)")
        else:
            # 일반 KS 양식으로 복구 시도 (있는 경우)
            data_dir = os.path.join(BASE_DIR, "Na-aba", "home", "data")
            if not os.path.exists(data_dir):
                data_dir = os.path.join(os.getcwd(), "Na-aba", "home", "data")
                
            ks_template = os.path.join(data_dir, "RT KS양식.xlsx")
            if os.path.exists(ks_template):
                self.rt_template_file_path.set(ks_template)
                self.log("💡 일반 KS 양식 템플릿으로 복구되었습니다.")
        
        # [NEW] 미리보기 컬럼 구성 업데이트
        self._update_rt_preview_columns("KOGAS" if self.rt_kogas_mode.get() else "RT")

    def _update_rt_preview_columns(self, mode="RT"):
        """모드 여부에 따라 RT 미리보기 컬럼을 동적으로 변경 (사용자 설정 키 준수)"""
        tree = getattr(self, "rt_preview_tree", None) if mode == "RT" else getattr(self, "kogas_preview_tree", None)
        if not tree: return
        
        # [DYNAMIC] 고정된 리스트 대신 사용자가 설정한 키(rt_column_keys 등)를 사용
        keys_attr = "rt_column_keys" if mode == "RT" else "kogas_column_keys"
        current_cols = list(getattr(self, keys_attr))
        if 'Report No' not in current_cols:
            current_cols.insert(current_cols.index('Date') + 1 if 'Date' in current_cols else 2, 'Report No')
        setattr(self, keys_attr, current_cols)
            
        tree["columns"] = tuple(current_cols)
        tree["displaycolumns"] = tuple(current_cols)
        
        # [NEW] 기본 표시 이름 맵핑
        default_names = {
            "selected": "V", "No": "No", "Date": "Date", "Report No": "Report No", "Sec": "SEC", "Dwg": "Drawing No.",
            "Joint": "Film Ident. No.", "Loc": "Film Location", "T": "T", "Mat": "Mat",
            "Size": "구경",
            "Acc": "Acc", "Rej": "Rej", "Deg": "Deg", "Welder": "Welder No", "Remarks": "Remarks",
            "Dwg_Sub": "Dwg(Sub)", "Welder_Sub": "Welder(Sub)", "Mat_Sub": "Mat(Sub)",
            "D1": "① Crack", "D2": "② IP", "D3": "③ LF", "D4": "④ Slag", "D5": "⑤ Por",
            "D6": "⑥ U/C", "D7": "⑦ RUC", "D8": "⑧ BT", "D9": "⑨ TI", "D10": "⑩ CP",
            "D11": "⑪ RC", "D12": "⑫ Mis", "D13": "⑬ EP", "D14": "⑭ SD", "D15": "⑮ Oth"
        }
        
        saved_widths = self.config.get(f"{mode}_COL_WIDTHS", {})
        default_widths = {
            "selected": 40, "No": 50, "Date": 90, "Sec": 90, "Dwg": 300, "Joint": 120, "Loc": 100, 
            "Report No": 220,
            "Acc": 40, "Rej": 40, "Deg": 40, "Welder": 100, "Remarks": 120,
            "Dwg_Sub": 200, "Welder_Sub": 100, "Mat_Sub": 100
        }
        
        # [NEW] 내부 컬럼 ID와 설정 Key(rt_items 정의) 간의 맵핑
        key_map = {
            "T": "THK", "Remarks": "REM", "Result": "RES"
        }
        
        for col in current_cols:
            suffix = key_map.get(col, col.upper())
            if mode == "KOGAS":
                name_key = f"KOGAS_R_NAME_{suffix}"
            else:
                name_key = f"{mode}_NAME_{suffix}"
            if col == "selected": name_key = None
            
            display_text = None
            if name_key:
                # 1. 현재 UI 입력값 확인 (최우선)
                if name_key in self.setting_vars:
                    display_text = self.setting_vars[name_key].get()
                
                # 2. 설정 파일(config) 확인
                if not display_text:
                    display_text = self.config.get(name_key)
                
                # 3. 보조 키 형식 확인 (소문자 등 대비)
                if not display_text:
                    if mode == "KOGAS":
                        alt_key = f"KOGAS_R_NAME_{col}"
                    else:
                        alt_key = f"{mode}_NAME_{col}"
                    if alt_key in self.setting_vars:
                        display_text = self.setting_vars[alt_key].get()
                    else:
                        display_text = self.config.get(alt_key)

            fallback_name = default_names.get(col, col)
            if not display_text:
                display_text = fallback_name
                
            tree.heading(col, text=display_text, anchor='center', command=lambda _c=col: self.sort_by_column(_c, mode=mode))
            w = saved_widths.get(col, default_widths.get(col, 80))
            tree.column(col, width=w, anchor='center', stretch=False)
        
        # 데이터가 이미 있다면 리프레시
        data = self.rt_extracted_data if mode == "RT" else self.kogas_extracted_data
        if data:
            self.populate_preview(data, mode=mode, switch_tab=False)
        elif mode in getattr(self, '_rt_check_menus', {}):
            self._refresh_rt_filter_headers(mode)

    def _setup_pt_ui(self, parent):
        # [FORCE] Ensure parent (pt_mode_frame) allows expansion
        container = tk.Frame(parent, background="#f9fafb")
        container.pack(fill='both', expand=True)

        # --- Dual Pane Layout (PanedWindow) ---
        self.pt_paned = tk.PanedWindow(container, orient='horizontal', background="#d1d5db", 
                            sashwidth=6, sashpad=0, sashrelief='raised', borderwidth=0)
        self.pt_paned.pack(fill='both', expand=True)

        # [LEFT] Settings Sidebar
        left_container = tk.Frame(self.pt_paned, background="#f9fafb", highlightthickness=0, borderwidth=0)
        self.pt_paned.add(left_container, width=425)
        
        # Scrollable area (Full size)
        left_pane = self._create_scrollable_sidebar(left_container)

        # FIXED FLOATING Header
        header_frame = tk.Frame(left_container, background="#f9fafb", highlightthickness=0, borderwidth=0)
        header_frame.place(x=0, y=0, relwidth=1, height=40)
        
        inner_header = tk.Frame(header_frame, background="#f9fafb", padx=20)
        inner_header.pack(fill='both', expand=True, pady=(5, 0))
        
        tk.Label(inner_header, text="🔬 PT 성적서 관리", font=("Malgun Gothic", 15, "bold"), 
                 background="#f9fafb", foreground="#1e3a8a").pack(side='left')
        
        tk.Frame(header_frame, height=1, background="#e5e7eb").pack(side='bottom', fill='x')
        header_frame.lift()

        # Spacer in scrollable area
        tk.Frame(left_pane, height=40, background="#f9fafb").pack(fill='x')

        # 1. File Selection Group
        file_frame = ttk.LabelFrame(left_pane, text=" 데이터 및 양식 (Data) ", padding=10)
        file_frame.pack(fill='x', pady=(0, 10))

        def _add_file_row(parent_f, label, var, row, is_dir=False, types=None):
            parent_f.columnconfigure(1, weight=1)
            ttk.Label(parent_f, text=label, font=("Malgun Gothic", 8)).grid(row=row, column=0, sticky='e', padx=2, pady=2)
            ttk.Entry(parent_f, textvariable=var, font=("Arial", 9), width=1, exportselection=False).grid(row=row, column=1, padx=2, pady=2, sticky='ew')
            cmd = (lambda: self._browse_dir(var)) if is_dir else (lambda: self._browse_file(var, types))
            ttk.Button(parent_f, text="...", width=3, command=cmd).grid(row=row, column=2, padx=2, pady=2)

        _add_file_row(file_frame, "기본 로고 폴더:", self.logo_folder_path, 0, is_dir=True)
        _add_file_row(file_frame, "PT 데이터:", self.pt_target_file_path, 1, types=[("Excel Source", "*.xls;*.xlsx;*.xlsm")])
        _add_file_row(file_frame, "PT 양식:", self.pt_template_file_path, 2, types=[("Excel Template", "*.xlsx;*.xlsm")])
        
        # [NEW] Template Config Save Button
        btn_frame = tk.Frame(file_frame, background="#f0f0f0")
        btn_frame.grid(row=3, column=1, sticky='ew', pady=2)
        ttk.Button(btn_frame, text="💾 현재 설정을 이 양식 전용으로 저장", command=lambda: self.save_template_specific_config("PT")).pack(side='left', padx=2)

        # 2. Configuration Tabs
        pt_config_frame = ttk.LabelFrame(left_pane, text=" 리포트 세부 설정 ", padding=2)
        pt_config_frame.pack(fill='both', expand=True, pady=(0, 10))

        self.pt_tab_notebook = ttk.Notebook(pt_config_frame)
        self.pt_tab_notebook.pack(fill='both', expand=True)

        pt_tab_cover = ttk.Frame(self.pt_tab_notebook, padding=5)
        pt_tab_data = ttk.Frame(self.pt_tab_notebook, padding=5)
        pt_tab_rows = ttk.Frame(self.pt_tab_notebook, padding=5)
        pt_tab_cols = ttk.Frame(self.pt_tab_notebook, padding=5)
        
        # [CRITICAL] Allow children to expand
        for t in [pt_tab_cover, pt_tab_data, pt_tab_rows, pt_tab_cols]: t.columnconfigure(0, weight=1)

        self.pt_tab_notebook.add(pt_tab_cover, text="갑지")
        self.pt_tab_notebook.add(pt_tab_data, text="을지")
        self.pt_tab_notebook.add(pt_tab_rows, text="행 설정")
        self.pt_tab_notebook.add(pt_tab_cols, text="컬럼 설정")

        # PT templates already contain their required logos and footer images.
        # Keep only metadata and print/layout controls in the PT tabs; exposing
        # separate image-path/position fields here was redundant and confusing.
        self._create_gapji_meta_ui(pt_tab_cover, use_pack=False, mode="PT")
        self.pt_tab_notebook.bind("<<NotebookTabChanged>>", self._update_gapji_preview_current)
        self._create_margin_settings(pt_tab_cover, "PT_COVER", use_pack=False)
        self._create_margin_settings(pt_tab_data, "PT_DATA", use_pack=False)
        self._create_row_settings(pt_tab_rows, mode="PT")
        
        pt_items = [
            ("순번(No):", "PT_COL_NO", 1, "PT_NAME_NO", "No", "No"),
            ("ISO/Dwg:", "PT_COL_DWG", 2, "PT_NAME_DWG", "Dwg", "Dwg"),
            ("Joint No:", "PT_COL_JOINT", 5, "PT_NAME_JOINT", "Joint", "Joint"),
            ("NPS:", "PT_COL_NPS", 6, "PT_NAME_NPS", "NPS", "NPS"),
            ("두께(Th'k):", "PT_COL_THK", 7, "PT_NAME_THK", "Thk.", "Thk."),
            ("재질(Material):", "PT_COL_MAT", 8, "PT_NAME_MAT", "Material", "Material"),
            ("용접사(Welder):", "PT_COL_WELDER", 9, "PT_NAME_WELDER", "Welder", "Welder"),
            ("검사타입(Type):", "PT_COL_TYPE", 10, "PT_NAME_TYPE", "WType", "WType"),
            ("판정 결과:", "PT_COL_RES", 11, "PT_NAME_RES", "Result", "Result")
        ]
        self._create_column_mapping_ui(pt_tab_cols, "PT", pt_items)
        

        # 3. Action Section
        action_frame = tk.Frame(left_pane, background="#ffffff", highlightthickness=1, highlightbackground="#d1d5db", padx=10, pady=5)
        action_frame.pack(fill='x', pady=(0, 10))

        filter_row = tk.Frame(action_frame, background="#ffffff")
        filter_row.pack(fill='x', pady=2)
        ttk.Label(filter_row, text="📊 특정순번:", background="#ffffff", font=("Malgun Gothic", 8)).pack(side='left')
        ttk.Entry(filter_row, textvariable=self.sequence_filter, width=15, font=("Arial", 9), exportselection=False).pack(side='left', padx=5, fill='x', expand=True)

        # Actions moved to right Top Action Bar

        # [RIGHT] Multi-Preview Pane (Data & Gapji)
        right_container = tk.Frame(self.pt_paned, background="#f3f4f6")
        self.pt_paned.add(right_container, stretch="always")
        
        self._create_top_action_bar(right_container, "PT")
        if not hasattr(self, 'preview_notebooks'): self.preview_notebooks = {}
        nb = ttk.Notebook(right_container)
        nb.pack(fill='both', expand=True)
        self.preview_notebooks["PT"] = nb
        
        tab_data = tk.Frame(nb, background="#ffffff")
        tab_gapji = tk.Frame(nb, background="#ffffff")
        nb.add(tab_data, text=" 🔬 데이터 미리보기 ")
        nb.add(tab_gapji, text=" 📄 갑지 미리보기 ")
        
        self._create_pt_preview_ui(tab_data)
        self._create_gapji_preview_ui(tab_gapji, "PT")
        self._apply_sash_ratio("PT")
        
        self.pt_template_file_path.trace_add("write", lambda *a: self.load_template_specific_config("PT"))

        # Adaptive Resizing Bindings
        self.pt_pane_ratio = self.config.get('PT_SASH_RATIO', 0.5)
        self.pt_paned.bind("<Configure>", self._on_pt_paned_configure)
        self.pt_paned.bind("<ButtonRelease-1>", lambda e: self.root.after(10, self._update_pt_ratio))
        self.root.after(500, lambda: self._on_pt_paned_configure(None))


    def _setup_mt_ui(self, parent):
        # [FORCE] Ensure parent (pt_mode_frame) allows expansion
        container = tk.Frame(parent, background="#f9fafb")
        container.pack(fill='both', expand=True)

        # --- Dual Pane Layout (PanedWindow) ---
        self.mt_paned = tk.PanedWindow(container, orient='horizontal', background="#d1d5db", 
                            sashwidth=6, sashpad=0, sashrelief='raised', borderwidth=0)
        self.mt_paned.pack(fill='both', expand=True)

        # [LEFT] Settings Sidebar
        left_container = tk.Frame(self.mt_paned, background="#f9fafb", highlightthickness=0, borderwidth=0)
        self.mt_paned.add(left_container, width=425)
        
        # Scrollable area (Full size)
        left_pane = self._create_scrollable_sidebar(left_container)

        # FIXED FLOATING Header
        header_frame = tk.Frame(left_container, background="#f9fafb", highlightthickness=0, borderwidth=0)
        header_frame.place(x=0, y=0, relwidth=1, height=40)
        
        inner_header = tk.Frame(header_frame, background="#f9fafb", padx=20)
        inner_header.pack(fill='both', expand=True, pady=(5, 0))
        
        tk.Label(inner_header, text="🔬 MT 성적서 관리", font=("Malgun Gothic", 15, "bold"), 
                 background="#f9fafb", foreground="#1e3a8a").pack(side='left')
        
        tk.Frame(header_frame, height=1, background="#e5e7eb").pack(side='bottom', fill='x')
        header_frame.lift()

        # Spacer in scrollable area
        tk.Frame(left_pane, height=40, background="#f9fafb").pack(fill='x')

        # 1. File Selection Group
        file_frame = ttk.LabelFrame(left_pane, text=" 데이터 및 양식 (Data) ", padding=10)
        file_frame.pack(fill='x', pady=(0, 10))

        def _add_file_row(parent_f, label, var, row, is_dir=False, types=None):
            parent_f.columnconfigure(1, weight=1)
            ttk.Label(parent_f, text=label, font=("Malgun Gothic", 8)).grid(row=row, column=0, sticky='e', padx=2, pady=2)
            ttk.Entry(parent_f, textvariable=var, font=("Arial", 9), width=1, exportselection=False).grid(row=row, column=1, padx=2, pady=2, sticky='ew')
            cmd = (lambda: self._browse_dir(var)) if is_dir else (lambda: self._browse_file(var, types))
            ttk.Button(parent_f, text="...", width=3, command=cmd).grid(row=row, column=2, padx=2, pady=2)

        _add_file_row(file_frame, "기본 로고 폴더:", self.logo_folder_path, 0, is_dir=True)
        _add_file_row(file_frame, "MT 데이터:", self.mt_target_file_path, 1, types=[("Excel Source", "*.xls;*.xlsx;*.xlsm")])
        _add_file_row(file_frame, "MT 양식:", self.mt_template_file_path, 2, types=[("Excel Template", "*.xlsx;*.xlsm")])
        
        # [NEW] Template Config Save Button
        btn_frame = tk.Frame(file_frame, background="#f0f0f0")
        btn_frame.grid(row=3, column=1, sticky='ew', pady=2)
        ttk.Button(btn_frame, text="💾 현재 설정을 이 양식 전용으로 저장", command=lambda: self.save_template_specific_config("MT")).pack(side='left', padx=2)

        # 2. Configuration Tabs
        pt_config_frame = ttk.LabelFrame(left_pane, text=" 리포트 세부 설정 ", padding=2)
        pt_config_frame.pack(fill='both', expand=True, pady=(0, 10))

        self.mt_tab_notebook = ttk.Notebook(pt_config_frame)
        self.mt_tab_notebook.pack(fill='both', expand=True)

        pt_tab_cover = ttk.Frame(self.mt_tab_notebook, padding=5)
        pt_tab_data = ttk.Frame(self.mt_tab_notebook, padding=5)
        pt_tab_rows = ttk.Frame(self.mt_tab_notebook, padding=5)
        pt_tab_cols = ttk.Frame(self.mt_tab_notebook, padding=5)
        
        # [CRITICAL] Allow children to expand
        for t in [pt_tab_cover, pt_tab_data, pt_tab_rows, pt_tab_cols]: t.columnconfigure(0, weight=1)

        self.mt_tab_notebook.add(pt_tab_cover, text="갑지")
        self.mt_tab_notebook.add(pt_tab_data, text="을지")
        self.mt_tab_notebook.add(pt_tab_rows, text="행 설정")
        self.mt_tab_notebook.add(pt_tab_cols, text="컬럼 설정")

        # [ALIGNED] Mode-specific context for logo grid
        self._create_gapji_meta_ui(pt_tab_cover, use_pack=False, mode="MT")
        self.mt_tab_notebook.bind("<<NotebookTabChanged>>", self._update_gapji_preview_current)
        next_row_mt_cover = self._create_setting_grid(pt_tab_cover, "MT_COVER")
        next_row_mt_data = self._create_setting_grid(pt_tab_data, "MT_DATA")
        self._create_margin_settings(pt_tab_cover, "MT_COVER", use_pack=False)
        self._create_margin_settings(pt_tab_data, "MT_DATA", use_pack=False)
        self._create_row_settings(pt_tab_rows, mode="MT")
        
        pt_items = [
            ("순번(No):", "MT_COL_NO", 1, "MT_NAME_NO", "No", "No"),
            ("ISO/Dwg:", "MT_COL_DWG", 2, "MT_NAME_DWG", "Dwg", "Dwg"),
            ("Joint No:", "MT_COL_JOINT", 5, "MT_NAME_JOINT", "Joint", "Joint"),
            ("NPS:", "MT_COL_NPS", 6, "MT_NAME_NPS", "NPS", "NPS"),
            ("두께(Th'k):", "MT_COL_THK", 7, "MT_NAME_THK", "Thk.", "Thk."),
            ("재질(Material):", "MT_COL_MAT", 8, "MT_NAME_MAT", "Material", "Material"),
            ("용접사(Welder):", "MT_COL_WELDER", 9, "MT_NAME_WELDER", "Welder", "Welder"),
            ("검사타입(Type):", "MT_COL_TYPE", 10, "MT_NAME_TYPE", "WType", "WType"),
            ("판정 결과:", "MT_COL_RES", 11, "MT_NAME_RES", "Result", "Result")
        ]
        self._create_column_mapping_ui(pt_tab_cols, "MT", pt_items)
        

        # 3. Action Section
        action_frame = tk.Frame(left_pane, background="#ffffff", highlightthickness=1, highlightbackground="#d1d5db", padx=10, pady=5)
        action_frame.pack(fill='x', pady=(0, 10))

        filter_row = tk.Frame(action_frame, background="#ffffff")
        filter_row.pack(fill='x', pady=2)
        ttk.Label(filter_row, text="📊 특정순번:", background="#ffffff", font=("Malgun Gothic", 8)).pack(side='left')
        ttk.Entry(filter_row, textvariable=self.sequence_filter, width=15, font=("Arial", 9), exportselection=False).pack(side='left', padx=5, fill='x', expand=True)

        # Actions moved to right Top Action Bar

        # [RIGHT] Multi-Preview Pane (Data & Gapji)
        right_container = tk.Frame(self.mt_paned, background="#f3f4f6")
        self.mt_paned.add(right_container, stretch="always")
        
        self._create_top_action_bar(right_container, "MT")
        if not hasattr(self, 'preview_notebooks'): self.preview_notebooks = {}
        nb = ttk.Notebook(right_container)
        nb.pack(fill='both', expand=True)
        self.preview_notebooks["MT"] = nb
        
        tab_data = tk.Frame(nb, background="#ffffff")
        tab_gapji = tk.Frame(nb, background="#ffffff")
        nb.add(tab_data, text=" 🔬 데이터 미리보기 ")
        nb.add(tab_gapji, text=" 📄 갑지 미리보기 ")
        
        self._create_mt_preview_ui(tab_data)
        self._create_gapji_preview_ui(tab_gapji, "MT")
        self._apply_sash_ratio("MT")
        
        self.mt_template_file_path.trace_add("write", lambda *a: self.load_template_specific_config(self.mt_template_file_path.get(), "MT"))

        # Adaptive Resizing Bindings
        self.mt_pane_ratio = self.config.get('MT_SASH_RATIO', 0.5)
        self.mt_paned.bind("<Configure>", self._on_mt_paned_configure)
        self.mt_paned.bind("<ButtonRelease-1>", lambda e: self.root.after(10, self._update_mt_ratio))
        self.root.after(500, lambda: self._on_mt_paned_configure(None))

    def _update_pt_ratio(self):
        try:
            total_w = self.pt_paned.winfo_width()
            if total_w > 100:
                current_sash = self.pt_paned.sash_coord(0)[0]
                self.pt_pane_ratio = current_sash / total_w
                self.config['PT_SASH_RATIO'] = self.pt_pane_ratio
                self.save_settings()
        except: pass


    def _update_mt_ratio(self):
        try:
            total_w = self.mt_paned.winfo_width()
            if total_w > 100:
                current_sash = self.mt_paned.sash_coord(0)[0]
                self.mt_pane_ratio = current_sash / total_w
                self.config['MT_SASH_RATIO'] = self.mt_pane_ratio
                self.save_settings()
        except: pass

    def _on_pt_paned_configure(self, event):
        try:
            total_w = self.pt_paned.winfo_width()
            if total_w > 100:
                if event and event.widget == self.pt_paned:
                    new_pos = int(total_w * self.pt_pane_ratio)
                    self.pt_paned.sash_place(0, new_pos, 0)
        except: pass


    def _on_mt_paned_configure(self, event):
        try:
            total_w = self.mt_paned.winfo_width()
            if total_w > 100:
                if event and event.widget == self.mt_paned:
                    new_pos = int(total_w * self.mt_pane_ratio)
                    self.mt_paned.sash_place(0, new_pos, 0)
        except: pass

    def _setup_paut_ui(self, parent):
        # [FORCE] Ensure parent (paut_mode_frame) allows expansion
        container = tk.Frame(parent, background="#f9fafb")
        container.pack(fill='both', expand=True)

        # --- Dual Pane Layout (PanedWindow) ---
        self.paut_paned = tk.PanedWindow(container, orient='horizontal', background="#d1d5db", 
                            sashwidth=6, sashpad=0, sashrelief='raised', borderwidth=0)
        self.paut_paned.pack(fill='both', expand=True)

        # [LEFT] Settings & Manual Eval Pane
        left_container = tk.Frame(self.paut_paned, background="#f9fafb", highlightthickness=0, borderwidth=0)
        self.paut_paned.add(left_container, width=425)
        
        # Scrollable area (Full size)
        left_pane = self._create_scrollable_sidebar(left_container)

        # FIXED FLOATING Header
        header_frame = tk.Frame(left_container, background="#f9fafb", highlightthickness=0, borderwidth=0)
        header_frame.place(x=0, y=0, relwidth=1, height=40)
        
        inner_header = tk.Frame(header_frame, background="#f9fafb", padx=20)
        inner_header.pack(fill='both', expand=True, pady=(5, 0))
        
        tk.Label(inner_header, text="🔬 PAUT (ASME B31.1) 관리", font=("Malgun Gothic", 15, "bold"), 
                 background="#f9fafb", foreground="#1e3a8a").pack(side='left')
        
        tk.Frame(header_frame, height=1, background="#e5e7eb").pack(side='bottom', fill='x')
        header_frame.lift()

        # Spacer in scrollable area
        tk.Frame(left_pane, height=40, background="#f9fafb").pack(fill='x')

        # 1. File Selection Group
        file_frame = ttk.LabelFrame(left_pane, text=" 데이터 및 양식 (Data) ", padding=10)
        file_frame.pack(fill='x', pady=(0, 10))

        def _add_file_row(parent_f, label, var, row, is_dir=False, types=None):
            parent_f.columnconfigure(1, weight=1)
            ttk.Label(parent_f, text=label, font=("Malgun Gothic", 8)).grid(row=row, column=0, sticky='e', padx=2, pady=2)
            ttk.Entry(parent_f, textvariable=var, font=("Arial", 9), width=1, exportselection=False).grid(row=row, column=1, padx=2, pady=2, sticky='ew')
            cmd = (lambda: self._browse_dir(var)) if is_dir else (lambda: self._browse_file(var, types))
            ttk.Button(parent_f, text="...", width=3, command=cmd).grid(row=row, column=2, padx=2, pady=2)

        _add_file_row(file_frame, "기본 로고 폴더:", self.logo_folder_path, 0, is_dir=True)
        _add_file_row(file_frame, "PAUT 데이터:", self.paut_target_file_path, 1, types=[("Excel Source", "*.xls;*.xlsx;*.xlsm")])
        _add_file_row(file_frame, "PAUT 양식:", self.paut_template_file_path, 2, types=[("Excel Template", "*.xlsx;*.xlsm")])

        # --- 2. Configuration Notebook (갑, 을, 행 설정, 개별 판정) ---
        config_frame = ttk.LabelFrame(left_pane, text=" 설정 및 판정 (Config & Eval) ", padding=2)
        config_frame.pack(fill='both', expand=True, pady=(0, 10))

        self.paut_tab_notebook = ttk.Notebook(config_frame)
        self.paut_tab_notebook.pack(fill='both', expand=True)

        tab_cover = ttk.Frame(self.paut_tab_notebook, padding=5)
        tab_data = ttk.Frame(self.paut_tab_notebook, padding=5)
        tab_rows = ttk.Frame(self.paut_tab_notebook, padding=5)
        tab_cols = ttk.Frame(self.paut_tab_notebook, padding=5)
        tab_eval = ttk.Frame(self.paut_tab_notebook, padding=5)

        # Allow expansion
        for t in [tab_cover, tab_data, tab_rows, tab_cols, tab_eval]: t.columnconfigure(0, weight=1)

        self.paut_tab_notebook.add(tab_eval, text="개별 판정") # Default/Main tab
        self.paut_tab_notebook.add(tab_cover, text="갑지")
        self.paut_tab_notebook.add(tab_data, text="을지")
        self.paut_tab_notebook.add(tab_rows, text="행 설정")
        self.paut_tab_notebook.add(tab_cols, text="컬럼 설정")
        
        # [NEW] PAUT Equipment Settings
        tab_equip_outer = ttk.Frame(self.paut_tab_notebook, padding=0)
        self.paut_tab_notebook.add(tab_equip_outer, text="장비/설정")
        
        equip_canvas = tk.Canvas(tab_equip_outer, background="#f9fafb", highlightthickness=0)
        equip_scrollbar = ttk.Scrollbar(tab_equip_outer, orient="vertical", command=equip_canvas.yview)
        tab_equip = ttk.Frame(equip_canvas, padding=5)
        
        tab_equip.bind("<Configure>", lambda e: equip_canvas.configure(scrollregion=equip_canvas.bbox("all")))
        
        # We don't force canvas width to avoid squeezing the Entry fields to 0 width.
        equip_canvas.create_window((0, 0), window=tab_equip, anchor="nw")
        
        # Use existing mousewheel binding if available
        if hasattr(self, '_bind_mousewheel_recursive'):
            self._bind_mousewheel_recursive(tab_equip)
            
        equip_canvas.configure(yscrollcommand=equip_scrollbar.set)
        equip_canvas.pack(side="left", fill="both", expand=True)
        equip_scrollbar.pack(side="right", fill="y")
        
        for col_idx in [1, 3]: # 2 columns layout
            tab_equip.columnconfigure(col_idx, weight=1)
        
        self.paut_equip_vars = {}
        self.paut_equip_map = {
            'equip': ('검사장비', 'A12', 'OmniScan MX2'),
            'module': ('검사모듈', 'H12', '32-128PR'),
            'probe': ('탐촉자', 'O12', '5L64-A2 / 5L16A10'),
            'wedge': ('웻지', 'V12', 'SA2-N55S-IHC / SA10-N55S'),
            'program': ('판독프로그램', 'AC12', 'Tomoview, Omni PC'),
            'method': ('검사기법', 'A15', 'PA/Sectoria'),
            'wave': ('파형', 'H15', '횡파'),
            'velocity': ('음속', 'O15', '3246.8m/s'),
            'filter': ('Band-Pass Filter', 'V15', 'Band-Pass 5MHz'),
            'medium': ('접촉 매질', 'AC15', '물 & 글리세린'),
            'voltage': ('Voltage', 'A17', '80V'),
            'gain': ('Gain', 'H17', '11 + 6dB / 21.4 + 6dB'),
            'mode': ('Mode', 'O17', 'PE(Pulse-Echo)'),
            'offset': ('Index Offset', 'V17', '±17 mm'),
            'ref_block': ('Ref Block', 'AC17', 'Test Level A Block(Φ2.5)'),

            'tcg1_p1': ('TCG1 Pos 1', 'H19', '0'),
            'tcg1_g1': ('TCG1 Gain 1', 'O19', '4.2dB'),
            'tcg1_p2': ('TCG1 Pos 2', 'H20', '8.96mm'),
            'tcg1_g2': ('TCG1 Gain 2', 'O20', '4.5dB'),
            'tcg1_p3': ('TCG1 Pos 3', 'H21', '17.13mm'),
            'tcg1_g3': ('TCG1 Gain 3', 'O21', '7.0dB'),
            'tcg1_p4': ('TCG1 Pos 4', 'H22', '24.79mm'),
            'tcg1_g4': ('TCG1 Gain 4', 'O22', '12.7dB'),

            'tcg2_p1': ('TCG2 Pos 1', 'AA19', '0'),
            'tcg2_g1': ('TCG2 Gain 1', 'AG19', '3.2dB'),
            'tcg2_p2': ('TCG2 Pos 2', 'AA20', '7.61mm'),
            'tcg2_g2': ('TCG2 Gain 2', 'AG20', '2.7dB'),
            'tcg2_p3': ('TCG2 Pos 3', 'AA21', '15.35mm'),
            'tcg2_g3': ('TCG2 Gain 3', 'AG21', '5.3dB'),
            'tcg2_p4': ('TCG2 Pos 4', 'AA22', '23.17mm'),
            'tcg2_g4': ('TCG2 Gain 4', 'AG22', '8.1dB'),

            'set1_qty': ('Gr.1 Used Elem', 'A25', '16'),
            'set1_first': ('Gr.1 First Elem', 'H25', '1'),
            'set1_last': ('Gr.1 Last Elem', 'O25', '16'),
            'set1_scan': ('Gr.1 Scan Res', 'V25', '1'),
            'set1_focal': ('Gr.1 Focal', 'AC25', '10'),
            'set1_start': ('Gr.1 Start Ang', 'A27', '45°'),
            'set1_stop': ('Gr.1 Stop Ang', 'H27', '75°'),
            'set1_ang_res': ('Gr.1 Ang Res', 'O27', '1°'),
            'set1_law': ('Gr.1 Law Config', 'V27', 'Sectorial'),
            
            'set2_qty': ('Gr.2 Used Elem', 'A30', '16'),
            'set2_first': ('Gr.2 First Elem', 'H30', '49'),
            'set2_last': ('Gr.2 Last Elem', 'O30', '64'),
            'set2_scan': ('Gr.2 Scan Res', 'V30', '1'),
            'set2_focal': ('Gr.2 Focal', 'AC30', '10'),
            'set2_start': ('Gr.2 Start Ang', 'A32', '45°'),
            'set2_stop': ('Gr.2 Stop Ang', 'H32', '75°'),
            'set2_ang_res': ('Gr.2 Ang Res', 'O32', '1°'),
            'set2_law': ('Gr.2 Law Config', 'V32', 'Sectorial')
        }
        
        equip_r, equip_c = 0, 0
        paut_standard_revision = "CENTRAL_PAUT_20260910"
        apply_new_standard = (
            self.config.get("PAUT_EQUIP_STANDARD_REV") != paut_standard_revision
        )
        for k, (label, cell, default_val) in self.paut_equip_map.items():
            config_val = self.config.get(f'PAUT_EQUIP_{k.upper()}')
            if apply_new_standard or config_val is None or str(config_val).strip() == '':
                config_val = default_val
                self.config[f'PAUT_EQUIP_{k.upper()}'] = default_val
            var = tk.StringVar(value=config_val)
            self.paut_equip_vars[k] = var
            ttk.Label(tab_equip, text=label).grid(row=equip_r, column=equip_c*2, sticky='e', padx=5, pady=5)
            ent = ttk.Entry(tab_equip, textvariable=var, width=16)
            ent.grid(row=equip_r, column=equip_c*2+1, sticky='ew', padx=5, pady=5)
            ent.bind("<FocusOut>", lambda e: self.save_settings())
            ent.bind("<Return>", lambda e: [self.root.focus_set(), self.save_settings()])
            
            equip_c += 1
            if equip_c > 1: # 2 columns layout
                equip_c = 0
                equip_r += 1
        self.config["PAUT_EQUIP_STANDARD_REV"] = paut_standard_revision

        # [NEW] Populate Settings
        self._create_gapji_meta_ui(tab_cover, use_pack=False, mode="PAUT")
        self.paut_tab_notebook.bind("<<NotebookTabChanged>>", self._update_gapji_preview_current)
        next_row_cover = self._create_setting_grid(tab_cover, "PAUT_COVER")
        next_row_data = self._create_setting_grid(tab_data, "PAUT_DATA")
        self._create_margin_settings(tab_cover, "PAUT_COVER", use_pack=False)
        self._create_margin_settings(tab_data, "PAUT_DATA", use_pack=False)
        self._create_row_settings(tab_rows, mode="PAUT")
        
        paut_items = [
            ("순번(No):", "PAUT_COL_NO", 1, "PAUT_COL_NO_NAME", "No", "No"),
            ("촬영일자(Date):", "PAUT_COL_DATE", 0, "PAUT_NAME_DATE", "Date", "Date"),
            ("Line No.:", "PAUT_COL_LINE", 2, "PAUT_NAME_LINE", "Line No.", "Line No."),
            ("Joint No.:", "PAUT_COL_JOINT", 3, "PAUT_NAME_JOINT", "Joint No.", "Joint No."),
            ("크기(Size):", "PAUT_COL_SIZE", 0, "PAUT_NAME_SIZE", "Size", "Size"),
            ("두께(Th'k):", "PAUT_COL_THK", 4, "PAUT_NAME_THK", "Th'k(mm)", "Th'k(mm)"),
            ("합격(Acc):", "PAUT_COL_ACC", 0, "PAUT_NAME_ACC", "Acc", "Acc"),
            ("불합격(Rej):", "PAUT_COL_REJ", 0, "PAUT_NAME_REJ", "Rej", "Rej"),
            ("결함높이(H):", "PAUT_COL_H", 5, "PAUT_NAME_H", "Height(mm)", "Height(mm)"),
            ("TCG MAX(%):", "PAUT_COL_TCG", 0, "PAUT_NAME_TCG", "TCG MAX(%)", "TCG MAX(%)"),
            ("결함길이(L):", "PAUT_COL_L", 6, "PAUT_NAME_L", "Length(mm)", "Length(mm)"),
            ("결함깊이(D):", "PAUT_COL_UP", 7, "PAUT_NAME_UP", "Upper", "Upper"),
            ("하부깊이(Low):", "PAUT_COL_LOW", 0, "PAUT_NAME_LOW", "Lower", "Lower"),
            ("시작위치:", "PAUT_COL_START", 0, "PAUT_NAME_START", "Start", "Start"),
            ("종료위치:", "PAUT_COL_END", 0, "PAUT_NAME_END", "End", "End"),
            ("a/l 비:", "PAUT_COL_AL", 0, "PAUT_NAME_AL", "a/l", "a/l"),
            ("a/t 비:", "PAUT_COL_AT", 0, "PAUT_NAME_AT", "a/t", "a/t"),
            ("결함종류:", "PAUT_COL_NAT", 8, "PAUT_NAME_NAT", "Type of Flaw", "Type of Flaw"),
            ("용접사(Welder):", "PAUT_COL_WELDER", 0, "PAUT_NAME_WELDER", "Welder", "Welder"),
            ("검사길이:", "PAUT_COL_TESTED_LEN", 0, "PAUT_NAME_TESTED_LEN", "Tested Length", "Tested Length"),
            ("판정(Eval):", "PAUT_COL_EVAL", 9, "PAUT_NAME_EVAL", "Evaluation", "Evaluation"),
            ("비고(Rem):", "PAUT_COL_REM", 10, "PAUT_NAME_REM", "Remarks", "Remarks")
        ]
        # [NEW] PAUT Column Settings Notebook for different templates
        self.paut_col_nb = ttk.Notebook(tab_cols)
        self.paut_col_nb.pack(fill='both', expand=True, padx=2, pady=2)
        
        for template_name in ["지역난방", "ASME", "KS", "AWS"]:
            frame = ttk.Frame(self.paut_col_nb)
            self.paut_col_nb.add(frame, text=template_name)
            
            # Localize the keys
            localized_items = []
            for item in paut_items:
                col_key = item[1].replace("PAUT_", f"PAUT_{template_name}_")
                name_key = item[3].replace("PAUT_", f"PAUT_{template_name}_")
                default_col = (23 if template_name == "지역난방" else 0) if item[1] == "PAUT_COL_TCG" else item[2]
                localized_items.append((item[0], col_key, default_col, name_key, item[4], item[5]))
                
            self._create_column_mapping_ui(frame, f"PAUT_{template_name}", localized_items)

        # --- 3. Manual Evaluation Content (Inside tab_eval) ---
        manual_container = tk.Frame(tab_eval, background="#f9fafb")
        manual_container.pack(fill='both', expand=True)
        
        # [NEW] Evaluation Mode Selector
        mode_frame = tk.Frame(manual_container, background="#f9fafb")
        mode_frame.pack(fill='x', pady=(0, 5))
        ttk.Label(mode_frame, text="✅ 판정 기준:", font=("Malgun Gothic", 8, "bold")).pack(side='left', padx=2)
        tk.Radiobutton(mode_frame, text="일반(B31.1)", variable=self.paut_eval_mode, value="Standard", background="#f9fafb").pack(side='left', padx=5)
        tk.Radiobutton(mode_frame, text="ECA(파과역학)", variable=self.paut_eval_mode, value="ECA", background="#f9fafb").pack(side='left', padx=5)

        # Sizing Helper Row
        calc_frame = tk.Frame(manual_container, background="#f1f5f9", padx=10, pady=8, highlightbackground="#cbd5e1", highlightthickness=1)
        calc_frame.pack(fill='x', pady=(0, 10))
        
        # Row 0: H-Calc Settings
        h_set_row = tk.Frame(calc_frame, background="#f1f5f9")
        h_set_row.pack(fill='x')
        ttk.Label(h_set_row, text="📏 H-Setting:", font=("Malgun Gothic", 8, "bold"), background="#f1f5f9").pack(side='left', padx=2)
        tk.Entry(h_set_row, textvariable=self.paut_manual_vars['peak'], width=5).pack(side='left', padx=2)
        ttk.Label(h_set_row, text="% /", background="#f1f5f9").pack(side='left')
        tk.Entry(h_set_row, textvariable=self.paut_manual_vars['db'], width=4).pack(side='left', padx=2)
        ttk.Label(h_set_row, text="dB ->", background="#f1f5f9").pack(side='left', padx=2)
        ttk.Label(h_set_row, textvariable=self.paut_manual_vars['target_fsh'], background="#f1f5f9", foreground="#ef4444", font=("Arial", 8, "bold")).pack(side='left', padx=2)
        
        # Row 1: H-Calc Data
        h_data_row = tk.Frame(calc_frame, background="#f1f5f9")
        h_data_row.pack(fill='x', pady=2)
        ttk.Label(h_data_row, text="📏 H-Data:", font=("Malgun Gothic", 8, "bold"), background="#f1f5f9").pack(side='left', padx=2)
        tk.Entry(h_data_row, textvariable=self.paut_manual_vars['z1'], width=6).pack(side='left', padx=2)
        ttk.Label(h_data_row, text="~", background="#f1f5f9").pack(side='left')
        tk.Entry(h_data_row, textvariable=self.paut_manual_vars['z2'], width=6).pack(side='left', padx=2)
        ttk.Button(h_data_row, text="입력", width=6, command=self._calculate_paut_h).pack(side='right', padx=5)

        # Row 2: L-Calc Data
        l_data_row = tk.Frame(calc_frame, background="#f1f5f9")
        l_data_row.pack(fill='x')
        ttk.Label(l_data_row, text="📏 L-Data:", font=("Malgun Gothic", 8, "bold"), background="#f1f5f9").pack(side='left', padx=2)
        tk.Entry(l_data_row, textvariable=self.paut_manual_vars['L1'], width=6).pack(side='left', padx=2)
        ttk.Label(l_data_row, text="~", background="#f1f5f9").pack(side='left')
        tk.Entry(l_data_row, textvariable=self.paut_manual_vars['L2'], width=6).pack(side='left', padx=2)
        ttk.Button(l_data_row, text="입력", width=6, command=self._calculate_paut_l).pack(side='right', padx=5)

        # Input Grid
        input_grid = tk.Frame(manual_container, background="#f9fafb")
        input_grid.pack(fill='x', pady=5)
        
        m_inputs = [("두께(T):", "t"), ("높이(H):", "h"), ("길이(L):", "l"), ("깊이(d):", "d"), 
                    ("시작(S):", "L1"), ("종료(E):", "L2"), ("상단(Up):", "D_Upper"), ("하단(Low):", "D_Lower")]
        for i, (lbl, key) in enumerate(m_inputs):
            grid_row = i // 2
            grid_col = (i % 2) * 2
            ttk.Label(input_grid, text=lbl, font=("Malgun Gothic", 8)).grid(row=grid_row, column=grid_col, sticky='e', padx=2, pady=1)
            ent = tk.Entry(input_grid, textvariable=self.paut_manual_vars[key], width=10)
            ent.grid(row=grid_row, column=grid_col+1, padx=2, pady=1, sticky='w')
            ent.bind("<Return>", lambda e: self.root.focus_set())

        # Nature & Evaluation Actions
        nature_row = tk.Frame(manual_container, background="#f9fafb")
        nature_row.pack(fill='x', pady=5)
        ttk.Label(nature_row, text="종류:").pack(side='left', padx=2)
        ttk.Combobox(nature_row, textvariable=self.paut_manual_vars['nature'], values=["Crack", "LOF", "IP", "Slag", "Porosity", "Others"], width=9).pack(side='left', padx=2)
        ttk.Label(nature_row, text="위치:").pack(side='left', padx=2)
        ttk.Label(nature_row, textvariable=self.paut_manual_vars['loc'], foreground="#3b82f6", font=("Arial", 9, "bold")).pack(side='left', padx=2)
        
        btn_row = tk.Frame(manual_container, background="#f9fafb")
        btn_row.pack(fill='x', pady=2)
        ttk.Button(btn_row, text="판정", width=8, command=self._run_manual_paut_eval).pack(side='left', expand=True, fill='x', padx=2)
        ttk.Button(btn_row, text="저장", width=8, command=self._save_paut_manual_to_tree).pack(side='left', expand=True, fill='x', padx=2)

        self.paut_res_label = tk.Label(manual_container, text="결과 대기", font=("Malgun Gothic", 10, "bold"), background="#f3f4f6")
        self.paut_res_label.pack(fill='x', pady=(5, 0))

        # Bottom Actions
        # [NEW] Photo Log Pages Input
        photo_pages_frame = tk.Frame(left_pane, background="#f9fafb")
        photo_pages_frame.pack(fill='x', pady=(5, 2))
        ttk.Label(photo_pages_frame, text="첨부할 사진대장 장수:", font=("Malgun Gothic", 9)).pack(side='left', padx=2)
        self.paut_photo_pages_var = tk.StringVar(value=str(self.config.get("PAUT_PHOTO_PAGES", "0")))
        ttk.Entry(photo_pages_frame, textvariable=self.paut_photo_pages_var, width=5).pack(side='left', padx=2)
        self.paut_photo_pages_var.trace_add(
            "write", lambda *_: self._sync_paut_photo_page_numbers()
        )
        
        # Core Actions moved to right Top Action Bar
        paut_photo_btn_row = tk.Frame(left_pane, background="#f9fafb")
        paut_photo_btn_row.pack(fill='x', pady=(0, 2))
        ttk.Button(paut_photo_btn_row, text=" 📸 사진대장 목록에서 PAUT 추출 (자동파싱) ", command=self._extract_paut_from_photos).pack(side='left', fill='x', expand=True)

        paut_session_row = tk.Frame(left_pane, background="#f9fafb")
        paut_session_row.pack(fill='x', pady=(0, 10))
        ttk.Button(paut_session_row, text=" 💾 세션 저장 ", command=self._export_paut_session).pack(side='left', fill='x', expand=True, padx=(0, 5))
        ttk.Button(paut_session_row, text=" 📂 세션 불러오기 ", command=self._import_paut_session).pack(side='left', fill='x', expand=True)

        # [RIGHT] Multi-Preview Pane (Data & Gapji)
        right_container = tk.Frame(self.paut_paned, background="#f3f4f6")
        self.paut_paned.add(right_container, stretch="always")
        
        self._create_top_action_bar(right_container, "PAUT")
        if not hasattr(self, 'preview_notebooks'): self.preview_notebooks = {}
        nb = ttk.Notebook(right_container)
        nb.pack(fill='both', expand=True)
        self.preview_notebooks["PAUT"] = nb
        
        tab_data = tk.Frame(nb, background="#ffffff")
        tab_gapji = tk.Frame(nb, background="#ffffff")
        nb.add(tab_data, text=" 🔬 데이터 미리보기 ")
        nb.add(tab_gapji, text=" 📄 갑지 미리보기 ")
        
        self._create_paut_preview_ui(tab_data)
        self._create_gapji_preview_ui(tab_gapji, "PAUT")
        self._apply_sash_ratio("PAUT")
        
    def _update_paut_ratio(self):
        try:
            total_w = self.paut_paned.winfo_width()
            if total_w > 100:
                current_sash = self.paut_paned.sash_coord(0)[0]
                self.paut_pane_ratio = current_sash / total_w
                self.config['PAUT_SASH_RATIO'] = self.paut_pane_ratio
                self.save_settings()
        except: pass

    def _on_paut_paned_configure(self, event):
        try:
            total_w = self.paut_paned.winfo_width()
            if total_w > 100:
                if event and event.widget == self.paut_paned:
                    new_pos = int(total_w * self.paut_pane_ratio)
                    self.paut_paned.sash_place(0, new_pos, 0)
        except: pass

    def _update_paut_auto_loc(self):
        try:
            t = float(self.paut_manual_vars['t'].get() or 0)
            d = float(self.paut_manual_vars['d'].get() or 0)
            h = float(self.paut_manual_vars['h'].get() or 0)
            if t > 0:
                is_surface = (d <= 0.1 * t) or (d + h >= 0.9 * t)
                loc = "Surface" if is_surface else "Subsurface"
                self.paut_manual_vars['loc'].set(loc)
        except: pass

    def _update_paut_target_fsh(self):
        """[NEW] Peak 나 dB 변경 시 목표 % FSH 실시간 업데이트"""
        try:
            peak = float(self.paut_manual_vars['peak'].get() or 0)
            db = float(self.paut_manual_vars['db'].get() or 0)
            if peak > 0:
                drop_factor = pow(10, -db / 20)
                target = peak * drop_factor
                self.paut_manual_vars['target_fsh'].set(f"{target:.1f}%")
            else:
                self.paut_manual_vars['target_fsh'].set("0.0%")
        except:
            self.paut_manual_vars['target_fsh'].set("-")

    def _calculate_paut_h(self):
        """[NEW] dB 강하법을 이용한 h 값 계산 및 자동 입력"""
        try:
            peak = float(self.paut_manual_vars['peak'].get() or 80)
            db = float(self.paut_manual_vars['db'].get() or 6)
            z1 = float(self.paut_manual_vars['z1'].get() or 0)
            z2 = float(self.paut_manual_vars['z2'].get() or 0)
            
            # Calculate Target FSH: Peak * 10^(-dB/20)
            drop_factor = pow(10, -db / 20)
            target_fsh = peak * drop_factor
            self.paut_manual_vars['target_fsh'].set(f"{target_fsh:.1f}%")

            h_calc = abs(z1 - z2)
            self.paut_manual_vars['h'].set(f"{h_calc:.2f}")
            
            # [NEW] D: 중간값, UP: 낮은값, LOW: 높은값 자동 기입
            d_mid = (z1 + z2) / 2
            d_min = min(z1, z2)
            d_max = max(z1, z2)
            
            self.paut_manual_vars['d'].set(f"{d_mid:.2f}")
            self.paut_manual_vars['D_Upper'].set(f"{d_min:.2f}")
            self.paut_manual_vars['D_Lower'].set(f"{d_max:.2f}")

            self._update_paut_auto_loc()
            self.log(f"📐 h 산출 완료 ({peak}% peak -> {db}dB 강하 목표 {target_fsh:.1f}%): {h_calc:.2f} mm (D:{d_mid:.2f}, U:{d_min:.2f}, L:{d_max:.2f})")
        except Exception as e:
            messagebox.showerror("계산 오류", f"입력값을 확인해주세요: {e}")

    def _calculate_paut_l(self):
        """[NEW] L1, L2 값을 이용한 결함 길이(L) 자동 계산"""
        try:
            l1 = float(self.paut_manual_vars['L1'].get() or 0)
            l2 = float(self.paut_manual_vars['L2'].get() or 0)
            l_calc = abs(l1 - l2)
            self.paut_manual_vars['l'].set(f"{l_calc:.2f}")
            self.log(f"📐 L 산출 완료 (L1:{l1} ~ L2:{l2}): {l_calc:.2f} mm")
        except Exception as e:
            messagebox.showerror("계산 오류", f"L1, L2 입력값을 확인해주세요: {e}")

    def _save_paut_manual_to_tree(self):
        """[NEW] 수동 판정 구역의 입력값을 선택된 트리뷰 행(들)에 저장"""
        selected_ids = self.paut_preview_tree.selection()
        if not selected_ids:
            messagebox.showwarning("선택 누락", "저장할 행을 먼저 선택해주세요.")
            return

        try:
            t = self.to_float(self.paut_manual_vars['t'].get() or 0)
            h = self.to_float(self.paut_manual_vars['h'].get() or 0)
            l = self.to_float(self.paut_manual_vars['l'].get() or 0)
            d = self.to_float(self.paut_manual_vars['d'].get() or 0)
            l1 = self.paut_manual_vars['L1'].get()
            l2 = self.paut_manual_vars['L2'].get()
            up = self.paut_manual_vars['D_Upper'].get()
            low = self.paut_manual_vars['D_Lower'].get()
            nat = self.paut_manual_vars['nature'].get()
            rem = "" # Remarks can be added to UI if needed, for now keep existing or clear

            # Run evaluation logic to get the result
            eval_mode = self.paut_eval_mode.get()
            res, loc = self.evaluate_paut_flaw(t, h, l, d, nat, mode=eval_mode)
            
            al = f"{h/l:.3f}" if l > 0 and h > 0 else ""
            at = f"{h/t:.3f}" if t > 0 and h > 0 else ""

            mode_info = self._get_mode_info("PAUT")
            if not mode_info: return
            _, idx_map, data_list, _ = mode_info

            for item_id in selected_ids:
                view_idx = self.paut_preview_tree.index(item_id)
                if 0 <= view_idx < len(idx_map):
                    actual_idx = idx_map[view_idx]
                    item = data_list[actual_idx]
                    item["Th'k(mm)"] = t
                    item['Height(mm)'] = h
                    item['Length(mm)'] = l
                    item['Start'] = l1
                    item['End'] = l2
                    item['Upper'] = up
                    item['Lower'] = low
                    item['Type of Flaw'] = nat
                    item['Evaluation'] = f"{res} ({loc})"
                    item['a/l'] = al
                    item['a/t'] = at
            
            self.populate_preview(data_list, mode="PAUT", switch_tab=False)
            self.log(f"💾 선택된 {len(selected_ids)}개 행에 수동 판정 결과 저장 완료.")
            
            # [NEW] 저장 후 입력 필드 초기화
            self._reset_paut_manual_inputs()
            
        except Exception as e:
            messagebox.showerror("저장 오류", f"입력값을 확인해주세요: {e}")

    def _run_manual_paut_eval(self):
        try:
            t = float(self.paut_manual_vars['t'].get() or 0)
            h = float(self.paut_manual_vars['h'].get() or 0)
            l = float(self.paut_manual_vars['l'].get() or 0)
            d = float(self.paut_manual_vars['d'].get() or 0)
            nat = self.paut_manual_vars['nature'].get()
            eval_mode = self.paut_eval_mode.get()
            
            res, loc = self.evaluate_paut_flaw(t, h, l, d, nat, mode=eval_mode)
            self.paut_manual_vars['loc'].set(loc)
            if res == "Accept":
                self.paut_res_label.config(text=f"{res} (위치: {loc})", fg="white", bg="#27ae60")
            else:
                self.paut_res_label.config(text=f"{res} (위치: {loc})", fg="white", bg="#e74c3c")
        except Exception as e:
            messagebox.showerror("입력 오류", f"입력값을 확인해주세요: {e}")

    def _extract_paut_from_photos(self):
        try:
            if not hasattr(self, 'photo_listbox') or self.photo_listbox.size() == 0:
                messagebox.showwarning("경고", "사진대장 탭에 로드된 사진이 없습니다.")
                return
            
            photo_list = self.photo_listbox.get(0, tk.END)
            self.paut_extracted_data = []
            
            for i, photo_path in enumerate(photo_list):
                filename = os.path.basename(photo_path)
                name_without_ext = os.path.splitext(filename)[0]
                
                # parsing logic: DT2024-016-101-0-1-600A_R04
                line_no = ""
                size = ""
                joint_no = ""
                
                if "_" in name_without_ext:
                    parts = name_without_ext.rsplit("_", 1)
                    line_no = parts[0] # DT2024-016-101-0-1-600A
                    joint_no = parts[1] # R04
                    
                    if "-" in line_no:
                        size = line_no.rsplit("-", 1)[1] # 600A
                else:
                    line_no = name_without_ext
                    
                if i == 0:
                    messagebox.showinfo("첫번째 사진 파싱 확인", f"원본 파일명: {filename}\n추출된 Line No: {line_no}\n추출된 Size: {size}\n추출된 Joint No: {joint_no}")

                    
                item = {
                    'selected': True,
                    'Line No.': line_no,
                    'Joint No.': joint_no,
                    'Size': size,
                    'Th\'k(mm)': "",
                    'Acc': "",
                    'Rej': "",
                    'Start': "",
                    'End': "",
                    'Length(mm)': "",
                    'Upper': "",
                    'Lower': "",
                    'Height(mm)': "",
                    'Type of Flaw': "",
                    'a/l': "", 'a/t': "",
                    'Evaluation': "",
                    'Remarks': ""
                }
                
                item_full = dict(item)
                for k in self.paut_column_keys:
                    if k not in item_full and k != "selected":
                        item_full[k] = ""
                self.paut_extracted_data.append(item_full)
                
            self.log(f"✅ 사진 목록에서 PAUT 데이터 생성 완료: {len(self.paut_extracted_data)} 건")
            self.populate_preview(self.paut_extracted_data, mode="PAUT")
            self.update_date_listbox("PAUT")
            
        except Exception as e:
            self.log(f"❌ 사진에서 PAUT 데이터 추출 중 오류: {e}")
            messagebox.showerror("오류", f"오류가 발생했습니다:\n{e}")

    def _extract_paut_data(self):
        file_path = self.paut_target_file_path.get()
        if not file_path:
            messagebox.showwarning("파일 미선택", "PAUT 데이터 엑셀 파일을 먼저 선택해주세요.")
            return
            
        try:
            self.log(f"📂 PAUT 데이터 불러오는 중: {os.path.basename(file_path)}")
            self.progress['value'] = 20
            
            df = pd.read_excel(file_path)
            cols = df.columns.tolist()
            
            mapping = {
                "date": next((c for c in cols if any(x in str(c).upper() for x in ["DATE", "일자", "날짜"])), None),
                "t": next((c for c in cols if any(x in c.upper() for x in ["THK", "두께", "TH'K", "THICKNESS"])), None),
                "h": next((c for c in cols if any(x in str(c).upper() for x in ["HEIGHT", "높이", "FLAW_HEIGHT"]) or str(c).strip().upper() == "H"), None),
                # [FIX] Prevent 'l' from matching '검사길이' or 'PAUT길이' (Tested Length)
                "l": next((c for c in cols if (any(x in str(c).upper() for x in ["LENGTH", "길이", "FLAW_LENGTH", "LENTH"]) or str(c).strip().upper() == "L") and not any(exclude in str(c).upper() for exclude in ["검사", "PAUT"])), None),
                "d": next((c for c in cols if any(x in str(c).upper() for x in ["DEPTH", "깊이", "FLAW_DEPTH"]) or str(c).strip().upper() == "D"), None),
                "nature": next((c for c in cols if any(x in str(c).upper() for x in ["NAT", "종류", "FLAW_NATURE", "TYPE"])), None),
                "line": next((c for c in cols if any(x in c.upper() for x in ["LINE", "ISO", "DWG", "DRAWING"])), None),
                "joint": next((c for c in cols if any(x in c.upper() for x in ["JOINT", "WELD"])), None),
                "start": next((c for c in cols if any(x in c.upper() for x in ["START", "L1"])), None),
                "end": next((c for c in cols if any(x in c.upper() for x in ["END", "L2"])), None),
                "upper": next((c for c in cols if any(x in c.upper() for x in ["UPPER", "UP", "DEPTH"])), None),
                "lower": next((c for c in cols if any(x in c.upper() for x in ["LOWER", "LOW"])), None),
                "acc": next((c for c in cols if any(x in str(c).upper() for x in ["ACC", "합격"])), None),
                "rej": next((c for c in cols if any(x in str(c).upper() for x in ["REJ", "불합격"])), None),
                "eval": next((c for c in cols if any(x in str(c).upper() for x in ["합부", "EVAL", "판정"])), None),
                "welder": next((c for c in cols if any(x in str(c).upper() for x in ["용접사", "WELDER"])), None),
                "tested_len": next((c for c in cols if any(x in str(c).upper().replace('\n', '').replace(' ', '') for x in ["검사길이", "PAUT길이", "TESTED", "TESTLEN"])), None),
                "size": next((c for c in cols if any(x in str(c).upper() for x in ["SIZE", "크기", "관경"])), None),
                "remarks": next((c for c in cols if any(x in c.upper() for x in ["REMARK", "REMAKER", "비고"])), None)
            }
            
            if not all([mapping["t"], mapping["h"], mapping["l"], mapping["d"]]):
                self.log("⚠️ 필수 컬럼을 모두 찾을 수 없어 자동 매핑에 일부 실패했습니다.")
                
            self.paut_extracted_data = []
            for _, row in df.iterrows():
                acc_raw = str(row.get(mapping["acc"], "")) if mapping["acc"] else ""
                rej_raw = str(row.get(mapping["rej"], "")) if mapping["rej"] else ""
                eval_raw = str(row.get(mapping["eval"], "")) if mapping["eval"] else ""
                
                # Derive ACC/REJ from '합부' if explicit ACC/REJ columns are missing
                if not acc_raw and not rej_raw and eval_raw:
                    if "합격" in eval_raw and "불합격" not in eval_raw:
                        acc_raw = "√"
                    elif "불합격" in eval_raw or "REJ" in eval_raw.upper() or "NG" in eval_raw.upper():
                        rej_raw = "√"
                    elif "ACC" in eval_raw.upper() or "OK" in eval_raw.upper():
                        acc_raw = "√"

                line_val = str(row.get(mapping["line"], "")) if mapping["line"] else ""
                if line_val.lower() == "nan" or not line_val.strip():
                    continue

                # [NEW] Format Date cleanly (remove time)
                raw_date = row.get(mapping["date"], "") if mapping["date"] else ""
                date_str = ""
                if pd.notna(raw_date) and raw_date:
                    try:
                        date_str = pd.to_datetime(raw_date).strftime('%Y-%m-%d')
                    except:
                        date_str = str(raw_date).split(' ')[0]

                # [NEW] Format Size (Append 'A' if missing)
                size_val = str(row.get(mapping["size"], "")) if mapping["size"] else ""
                size_val = str(size_val).replace('.0', '').strip() if pd.notna(size_val) else ""
                
                # [NEW] Auto-map Thickness based on Size if Th'k is missing
                thk_val = str(row.get(mapping["t"], "")) if mapping["t"] else ""
                if not thk_val or str(thk_val).lower() == "nan":
                    # Official GS Power Standard
                    thk_mapping = {
                        "1100": "11.1", "1000": "11.1",
                        "900": "10.3", "850": "10.3",
                        "800": "9.5",
                        "750": "8.7", "700": "8.7", "650": "8.7",
                        "600": "9.5", "550": "9.5",
                        "500": "6.4", "450": "6.4", "400": "6.4", "350": "6.4", "300": "6.4", "250": "6.4", "200": "6.4"
                    }
                    lookup_key = size_val.replace("A", "").strip()
                    thk_val = thk_mapping.get(lookup_key, "")
                
                if size_val and size_val[-1].isdigit():
                    size_val += "A"

                # [NEW] Format Joint No (Pad single digits to 2 digits, handling suffixes like R3-1 -> R03-1)
                import re
                joint_val = str(row.get(mapping["joint"], "")) if mapping["joint"] else ""
                joint_val = joint_val.replace('.0', '').strip() if pd.notna(joint_val) else ""
                joint_val = re.sub(r'^([A-Za-z]*)(\d)(?!\d)', r'\g<1>0\2', joint_val)

                item = {
                    'selected': True,
                    'Date': date_str,
                    'Line No.': line_val,
                    'Joint No.': joint_val,
                    'Size': size_val,
                    'Th\'k(mm)': thk_val,
                    'Acc': acc_raw,
                    'Rej': rej_raw,
                    'Start': str(row.get(mapping["start"], "")) if mapping["start"] else "",
                    'End': str(row.get(mapping["end"], "")) if mapping["end"] else "",
                    'Length(mm)': str(row.get(mapping["l"], "")) if mapping["l"] else "",
                    'Upper': str(row.get(mapping["upper"], "")) if mapping["upper"] else "",
                    'Lower': str(row.get(mapping["lower"], "")) if mapping["lower"] else "",
                    'Height(mm)': str(row.get(mapping["h"], "")) if mapping["h"] else "",
                    'Type of Flaw': "" if acc_raw == "√" else (str(row.get(mapping["nature"], "Slag")) if mapping["nature"] else "Slag"),
                    'a/l': "", 'a/t': "",
                    'Welder': self._normalize_daily_photo_welder(
                        row.get(mapping["welder"], "")
                    ) if mapping["welder"] else "",
                    'Tested Length': str(row.get(mapping["tested_len"], "")) if mapping["tested_len"] else "",
                    'Evaluation': eval_raw,
                    'Remarks': str(row.get(mapping["remarks"], "")) if mapping["remarks"] else ""
                }
                # [NEW] 원본 엑셀의 모든 컬럼 데이터를 유지하여 '컬럼 관리'에서 활용 가능하게 함
                item_full = row.to_dict()
                item_full.update(item)
                from site_apps.central.src.reports.paut_fields import normalize_tcg
                normalize_tcg(item_full)
                # [NEW] 수동 추가된 컬럼(Remark 등)이 추출 시 증발하지 않도록 초기화
                for k in self.paut_column_keys:
                    if k not in item_full and k != "selected":
                        item_full[k] = ""
                self.paut_extracted_data.append(item_full)
            
            self.progress['value'] = 100
            self.sort_by_column("ISO/DWG", mode="PAUT") # Auto sort for PAUT
            self.update_date_listbox("PAUT") # [FIX] Populate date filter
            self.log(f"✅ PAUT 데이터 추출 완료: {len(self.paut_extracted_data)} 건")
            
        except Exception as e:
            self.log(f"❌ PAUT 데이터 추출 오류: {e}")

    def _run_batch_paut_eval(self):
        if not self.paut_extracted_data:
            messagebox.showwarning("데이터 없음", "먼저 데이터를 추출해주세요.")
            return
            
        self.log("🚀 PAUT 일괄 판정 시작...")
        count_ok, count_ng = 0, 0
        
        for item in self.paut_extracted_data:
            t = self.to_float(item.get('Th\'k(mm)', 0))
            h = self.to_float(item.get('Height(mm)', 0))
            l = self.to_float(item.get('Length(mm)', 0))
            d = self.to_float(item.get('Upper', 0))
            nat = item.get('Type of Flaw', 'Slag')
            eval_mode = self.paut_eval_mode.get()
            
            res, loc = self.evaluate_paut_flaw(t, h, l, d, nat, mode=eval_mode)
            item['Evaluation'] = f"{res} ({loc})"
            # a/l, a/t calculation
            if l > 0 and h > 0: item['a/l'] = f"{h/l:.3f}"
            else: item['a/l'] = ""
            if t > 0 and h > 0: item['a/t'] = f"{h/t:.3f}"
            else: item['a/t'] = ""
            
            if res == "Accept": 
                count_ok += 1
                # [NEW] 합격 시 결함 정보 지우기
                item['Type of Flaw'] = ""
                item['a/l'] = ""
                item['a/t'] = ""
            else: 
                count_ng += 1
            
        self.populate_preview(self.paut_extracted_data, mode="PAUT")
        self.log(f"✅ 판정 완료: 총 {len(self.paut_extracted_data)} 건 (합격: {count_ok}, 불합격: {count_ng})")

    def _generate_paut_report(self):
        # PAUT 기준 성적서 값은 실행 중 메모리에 남아 있는 이전 설정보다
        # 우선한다. save_settings() 전에 UI 변수까지 갱신하여 설정 파일이
        # 다시 예전 값으로 덮어써지지 않게 한다.
        reference_settings = {
            'probe': '5L64-A2 / 5L16A10',
            'wedge': 'SA2-N55S-IHC / SA10-N55S',
            'gain': '11 + 6dB / 21.4 + 6dB',
            'offset': '±17 mm',
            'tcg1_p1': '0', 'tcg1_g1': '4.2dB',
            'tcg1_p2': '8.96mm', 'tcg1_g2': '4.5dB',
            'tcg1_p3': '17.13mm', 'tcg1_g3': '7.0dB',
            'tcg1_p4': '24.79mm', 'tcg1_g4': '12.7dB',
            'tcg2_p1': '0', 'tcg2_g1': '3.2dB',
            'tcg2_p2': '7.61mm', 'tcg2_g2': '2.7dB',
            'tcg2_p3': '15.35mm', 'tcg2_g3': '5.3dB',
            'tcg2_p4': '23.17mm', 'tcg2_g4': '8.1dB',
            'set1_qty': '16', 'set1_first': '1', 'set1_last': '16',
            'set2_qty': '16', 'set2_first': '49', 'set2_last': '64',
        }
        if hasattr(self, 'paut_equip_vars'):
            for key, value in reference_settings.items():
                if key in self.paut_equip_vars:
                    self.paut_equip_vars[key].set(value)
                self.config[f'PAUT_EQUIP_{key.upper()}'] = value

        self.save_settings() # Ensure UI -> config sync
        template_path = self.paut_template_file_path.get()
        if not template_path or not os.path.exists(template_path):
            messagebox.showwarning("파일 미선택", "PAUT 성적서 양식 파일을 선택해주세요.")
            return

        final_list = [
            d for d in self.paut_extracted_data
            if d.get('selected', True)
            and d.get('date_filtered', True)
            and d.get('line_filtered', True)
            and d.get('size_filtered', True)
        ]
        if not final_list:
            messagebox.showwarning("항목 미선택", "선택된 데이터가 없습니다.")
            return

        contractors = {str(item.get('_source_contractor', '')).strip() for item in final_list
                       if str(item.get('_source_contractor', '')).strip()}
        if len(contractors) > 1:
            messagebox.showwarning('시공사 확인', '여러 시공사가 포함되어 있습니다. 같은 시공사 자료만 선택해 주세요.\n' + ', '.join(sorted(contractors)))
            return
        if contractors:
            self.paut_contractor.set(next(iter(contractors)))
            self.config['PAUT_CONTRACTOR'] = self.paut_contractor.get()
            self.save_settings()
        self.log(f"🚀 PAUT 성적서 생성 시작 (총 {len(final_list)} 건)...")
        self.progress['value'] = 0
        
        try:
            # keep_vba=True forces openpyxl to write a macro-enabled workbook
            # content type.  When the selected template is a normal .xlsx file,
            # saving that workbook with an .xlsx extension makes Excel reject it
            # as an extension/format mismatch.
            template_ext = os.path.splitext(template_path)[1].lower()
            keep_vba = template_ext == '.xlsm'
            wb = openpyxl.load_workbook(template_path, keep_vba=keep_vba)
            
            # [NEW] Find "을" sheet for data entry
            target_ws = None
            for sheet in wb.worksheets:
                if "을" in sheet.title:
                    target_ws = sheet
                    break
                    
            if not target_ws:
                target_ws = wb.worksheets[0]
                
            ws = target_ws
            ws.title = f"PAUT_Report_001"
            self._prepare_paut_data_stamp(ws)
            
            # [NEW] Write Cover (Gapji) metadata if there is a cover sheet
            if len(wb.worksheets) > 1:
                first_item = final_list[0] if final_list else None
                ws0 = wb.worksheets[0]
                self._write_gapji_metadata(ws0, mode="PAUT", first_item=first_item)

                # 리포트 기본정보의 검사일자를 우선 사용한다. 입력값이 없을 때만
                # 선택 데이터의 실제 검사기간을 대신 표시한다.
                if not self.gapji_exam_date.get().strip():
                    selected_dates = []
                    for report_item in final_list:
                        date_text = str(report_item.get("Date", "")).strip()
                        date_match = re.search(r"\d{4}-\d{2}-\d{2}", date_text)
                        if date_match:
                            selected_dates.append(date_match.group())
                    if selected_dates:
                        start_date = min(selected_dates)
                        end_date = max(selected_dates)
                        exam_period = (
                            start_date if start_date == end_date
                            else f"{start_date}~{end_date}"
                        )
                        self.safe_set_value(ws0, "AA6", exam_period)

                if not str(ws0["AA5"].value or "").strip():
                    self.safe_set_value(ws0, "AA5", "SIT/GI-JY-SRJ-PAUT-001")
                self._prepare_paut_cover_identity(ws0)
                
                # 갑지의 AA4, AA5, AA6 내용을 을지의 Y4, Y5, Y6에 복사
                try:
                    self.safe_set_value(ws, 'Y4', ws0['AA4'].value)
                    self.safe_set_value(ws, 'Y5', ws0['AA5'].value)
                    self.safe_set_value(ws, 'Y6', ws0['AA6'].value)
                except Exception as e:
                    self.log(f"[WARNING] 갑지 데이터 복사 실패: {e}")
            
            self.force_print_settings(ws, context="DATA")
            
            start_row = int(self.config.get('PAUT_START_ROW', 11))
            end_row = int(self.config.get('PAUT_DATA_END_ROW', 40))
            
            curr_row = start_row
            data_font = Font(size=9)
            
            try:
                active_tab_id = self.paut_col_nb.select()
                active_template = self.paut_col_nb.tab(active_tab_id, "text")
            except:
                active_template = "지역난방"

            for i, item in enumerate(final_list):
                if curr_row > end_row:
                    ws = self.prepare_next_paut_sheet(wb, ws.title, i//(end_row-start_row+1))
                    curr_row = start_row
                
                def _get_col(key, default):
                    localized_key = key.replace("PAUT_", f"PAUT_{active_template}_")
                    val = str(self.config.get(localized_key, '')).strip()
                    if not val:
                        val = str(self.config.get(key, '')).strip()
                    return val if val else default

                # [NEW] Combine ACC and REJ to prevent overwriting if they map to the same column
                acc_val = 'ACC' if item.get('Acc', '') in ['√', 'ACC'] else ''
                rej_val = 'REJ' if item.get('Rej', '') in ['√', 'REJ'] else ''
                acc_rej_combined = acc_val if acc_val else rej_val

                # [NEW] Check if there is no defect
                from site_apps.central.src.reports.paut_fields import normalize_tcg
                normalize_tcg(item)
                is_no_defect = (not str(item.get('Type of Flaw', '')).strip()
                                and not str(item.get('Height(mm)', '')).strip())
                if active_template == '지역난방':
                    from site_apps.central.src.reports.paut_report_layout import prepare_indication_row
                    prepare_indication_row(ws, curr_row, not is_no_defect,
                                           has_tcg=str(item.get('TCG MAX(%)', '')).strip() != '')

                # [NEW] Respect user's UI column configuration while providing standard defaults
                mapping = {
                    _get_col('PAUT_COL_NO', 'A'): i+1, 
                    _get_col('PAUT_COL_LINE', 'B'): item.get('Line No.', ''), 
                    _get_col('PAUT_COL_JOINT', 'C'): item.get('Joint No.', ''),
                    _get_col('PAUT_COL_SIZE', 'D'): item.get('Size', ''),
                    _get_col('PAUT_COL_THK', 'E'): item.get('Th\'k(mm)', ''), 
                    _get_col('PAUT_COL_ACC', 'F'): acc_rej_combined,
                    _get_col('PAUT_COL_START', 'G'): item.get('Start', ''),
                    _get_col('PAUT_COL_END', 'H'): item.get('End', ''),
                    _get_col('PAUT_COL_L', 'I'): item.get('Length(mm)', ''),
                    _get_col('PAUT_COL_UP', 'J'): item.get('Upper', ''), 
                    _get_col('PAUT_COL_H', 'K'): item.get('Height(mm)', ''), 
                    _get_col('PAUT_COL_TCG', 'W' if active_template == '지역난방' else 'L'): item.get('TCG MAX(%)', ''),
                    _get_col('PAUT_COL_WELDER', 'M'): item.get('Welder', ''),
                    _get_col('PAUT_COL_TESTED_LEN', 'N'): item.get('Tested Length', ''),
                    _get_col('PAUT_COL_NAT', 'O'): item.get('Type of Flaw', '') or (
                        item.get('Remarks', '') if active_template == '지역난방' else ''),
                    _get_col('PAUT_COL_EVAL', '0'): item.get('Evaluation', ''),
                    _get_col('PAUT_COL_REM', '0'): item.get('Remarks', '')
                }
                
                # [NEW] Explicitly write "NO RECORDABLE INDICATION" to column N if there is no defect
                if is_no_defect:
                    mapping[self.col_to_num('N')] = "NO RECORDABLE INDICATION"
                
                for col_key, val in mapping.items():
                    try:
                        col_idx = self.col_to_num(col_key)
                        if col_idx < 1: continue # Skip if 0 or invalid
                        
                        cell = ws.cell(row=curr_row, column=col_idx, value=val)
                        cell.font = data_font
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                    except Exception:
                        continue
                
                curr_row += 1
                self.progress['value'] = (i+1) / len(final_list) * 100
                
            # [NEW] Add "- 이 하 여 백 -" at the row following the last data row
            if curr_row <= end_row:
                try:
                    ws.merge_cells(start_row=curr_row, start_column=14, end_row=curr_row, end_column=24)
                    blank_cell = ws.cell(row=curr_row, column=14, value="- 이 하 여 백 -")
                    blank_cell.font = Font(name='맑은 고딕', size=10, bold=True)
                    blank_cell.alignment = Alignment(horizontal='center', vertical='center')
                except Exception as e:
                    self.log(f"이하 여백 삽입 오류: {e}")
                
            # [NEW] Add Tested Length Summary at the end
            total_org_m = 0.0
            total_rep_m = 0.0
            processed_joints = set()
            
            for item in final_list:
                try:
                    joint_no = str(item.get('Joint No.', '')).strip().upper()
                    if not joint_no or joint_no in processed_joints:
                        continue
                    
                    val_str = str(item.get('Tested Length', '0')).replace(',', '').strip()
                    num_match = re.search(r"[\d\.]+", val_str)
                    if num_match:
                        val = float(num_match.group())
                        processed_joints.add(joint_no) # Mark as processed only if we successfully parsed length
                        
                        acc_val = 'ACC' if item.get('Acc', '') in ['√', 'ACC'] else ''
                        rej_val = 'REJ' if item.get('Rej', '') in ['√', 'REJ'] else ''
                        acc_rej_combined = acc_val if acc_val else rej_val
                        
                        # Use ACC/REJ to distinguish Original vs Repair
                        if acc_rej_combined == 'REJ' or '불합격' in str(item.get('Rej', '')):
                            total_rep_m += val
                        else:
                            total_org_m += val
                except:
                    pass
            
            total_m = total_org_m + total_rep_m

            org_val = f"{total_org_m:.4f} M" if total_org_m > 0 else " M"
            rep_val = f"{total_rep_m:.4f} M" if total_rep_m > 0 else " M"
            tot_val = f"{total_m:.4f} M" if total_m > 0 else " M"

            from openpyxl.styles import Border, Side
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

            summary_data = [
                ("Original", org_val),
                ("Repair", rep_val),
                ("Total", tot_val)
            ]
            
            aa_col = self.col_to_num('AA')
            ae_col = self.col_to_num('AE')
            
            # [NEW] Check if there is enough space for the 3-row summary.
            # If not, create a new sheet for the summary to prevent overwriting footers.
            if curr_row + 2 > end_row:
                ws = self.prepare_next_paut_sheet(wb, ws.title, len(final_list)//(end_row-start_row+1) + 1)
                curr_row = start_row

            for r_idx, (label, val) in enumerate(summary_data):
                r = curr_row + r_idx
                
                c_label = ws.cell(row=r, column=aa_col, value=label)
                c_label.border = thin_border
                c_label.font = data_font
                c_label.alignment = Alignment(horizontal='center', vertical='center')
                
                c_val = ws.cell(row=r, column=ae_col, value=val)
                c_val.border = thin_border
                c_val.font = data_font
                c_val.alignment = Alignment(horizontal='center', vertical='center')
            
            # [NEW] PAUT Page Numbering Logic (Gapji: AE3, Eulji: AF3)
            photo_pages = 0
            try:
                if hasattr(self, 'paut_photo_pages_var'):
                    photo_pages = int(self.paut_photo_pages_var.get())
            except:
                pass
            
            total_p = len(wb.worksheets) + photo_pages

            # PAUT 보고서 뒤에 붙는 사진대장은 전체 보고서의 다음 번호부터
            # 이어지도록 사진대장 탭의 시작/전체 페이지를 자동 동기화한다.
            # 예: PAUT 2페이지 + 사진대장 3페이지 -> 사진대장 3~5 / 전체 5.
            if photo_pages > 0:
                photo_start = len(wb.worksheets) + 1
                if hasattr(self, "photo_start_page"):
                    self.photo_start_page.set(str(photo_start))
                if hasattr(self, "photo_total_pages"):
                    self.photo_total_pages.set(str(total_p))
                self.log(
                    f"[PAUT] 사진대장 페이지 자동 설정: "
                    f"시작 {photo_start}, 전체 {total_p}"
                )

            for p_idx, s in enumerate(wb.worksheets):
                page_num = p_idx + 1
                p_text = f"Page    {page_num}    of    {total_p}"
                
                # If there are multiple sheets, the first one is the Cover (Gapji)
                if p_idx == 0 and len(wb.worksheets) > 1:
                    self.safe_set_value(s, 'AC3', p_text)
                    try: 
                        if s['AC3'].font: s['AC3'].font = s['AC3'].font.copy(name='바탕', size=9, bold=False)
                        else: s['AC3'].font = Font(name='바탕', size=9, bold=False)
                    except: pass

                    # 장비/TCG 및 ISO 시험 등급 기준값은 갑지에만 적용한다.
                    if hasattr(self, 'paut_equip_vars'):
                        for k, (lbl, cell, default_val) in self.paut_equip_map.items():
                            val = self.paut_equip_vars[k].get()
                            self.safe_set_value(s, cell, val)

                    grade_values = {
                        'A38': 'C',
                        'F38': 'A-Technique',
                        'K38': 'Level 3',
                        'R38': '11 + 6dB / 21.4 + 6dB',
                    }
                    for cell_ref, value in grade_values.items():
                        self.safe_set_value(s, cell_ref, value)

                else:
                    self.safe_set_value(s, 'AA3', p_text)
                    try: 
                        if s['AA3'].font: s['AA3'].font = s['AA3'].font.copy(name='바탕', size=9, bold=False)
                        else: s['AA3'].font = Font(name='바탕', size=9, bold=False)
                    except: pass

                # Fix right border for AJ4:AJ45 (Draw solid thin line)
                try:
                    from site_apps.central.src.reports.paut_report_layout import solid_right_border
                    solid_right_border(s)
                except Exception as e:
                    self.log(f"AJ border error: {e}")

                # 갑지 검사자는 박광복으로 표시하고 등록된 서명을 함께 넣는다.
                self._apply_paut_inspector_signature(s)
                self._make_header_stamp_transparent(s)

            # [NEW] 최종 저장 직전 사용자의 행/열 커스텀 설정을 모든 시트에 적용
            from site_apps.central.src.reports.paut_report_layout import align_report_heights, measure_excel_print_widths
            from site_apps.central.src.reports.paut_report_layout import fit_height_heading
            for p_idx, s in enumerate(wb.worksheets):
                ctx = "COVER" if p_idx == 0 else "DATA"
                self.apply_custom_dimensions(s, ctx)
                self.force_print_settings(s, ctx)
                if p_idx > 0:
                    fit_height_heading(s)

            output_ext = '.xlsm' if keep_vba else '.xlsx'
            output_dir = os.path.join(BASE_DIR, "output") if os.path.exists(os.path.join(BASE_DIR, "output")) else BASE_DIR
            out_path = self._report_save_path(output_dir, output_ext, mode='PAUT')
            if not out_path:
                return
            out_name = os.path.basename(out_path)
            
            if not os.path.exists(os.path.dirname(out_path)): os.makedirs(os.path.dirname(out_path))
            wb.save(out_path)
            adjusted_path = None
            adjusted_wb = None
            try:
                widths, heights = measure_excel_print_widths(
                    out_path, wb.sheetnames, include_heights=True)
                # Saving an image-bearing openpyxl workbook closes its image
                # streams. Reload before saving again, and replace atomically.
                adjusted_wb = openpyxl.load_workbook(out_path, keep_vba=keep_vba)
                align_report_heights(adjusted_wb.worksheets, measured_widths=widths,
                                     measured_heights=heights)
                fd, adjusted_path = tempfile.mkstemp(
                    prefix='paut_print_', suffix=output_ext,
                    dir=os.path.dirname(out_path))
                os.close(fd)
                adjusted_wb.save(adjusted_path)
                os.replace(adjusted_path, out_path)
                adjusted_path = None
            except Exception as exc:
                self.log(f"[PAUT] 인쇄 높이 정렬 실패 — 보정 전 정상 파일 유지: {exc}")
            finally:
                if adjusted_wb is not None:
                    adjusted_wb.close()
                if adjusted_path and os.path.exists(adjusted_path):
                    os.remove(adjusted_path)
            self._record_report_issue('PAUT', wb, out_path, final_list)
            self.log(f"✅ PAUT 성적서 생성 완료: {out_name}")
            os.startfile(os.path.dirname(out_path))
            
        except Exception as e:
            self.log(f"❌ PAUT 성적서 생성 오류: {e}")

    def _sync_paut_photo_page_numbers(self):
        """PAUT 첨부 장수로 사진대장의 시작/전체 페이지를 미리 계산한다."""
        if "photo_start_page" not in self.__dict__ or "photo_total_pages" not in self.__dict__:
            return
        try:
            raw_pages = self.__dict__["paut_photo_pages_var"].get() if "paut_photo_pages_var" in self.__dict__ else self.config.get("PAUT_PHOTO_PAGES", "0")
            photo_pages = int(str(raw_pages).strip() or "0")
        except (TypeError, ValueError):
            return
        if photo_pages <= 0:
            return

        selected_count = sum(
            1 for item in getattr(self, "paut_extracted_data", [])
            if item.get("selected", True)
            and item.get("date_filtered", True)
            and item.get("line_filtered", True)
            and item.get("size_filtered", True)
        )
        start_row = int(self.config.get("PAUT_START_ROW", 11))
        end_row = int(self.config.get("PAUT_DATA_END_ROW", 40))
        rows_per_page = max(1, end_row - start_row + 1)
        data_pages = max(1, math.ceil(selected_count / rows_per_page))
        paut_pages = 1 + data_pages  # 갑지 1장 + 을지

        self.photo_start_page.set(str(paut_pages + 1))
        self.photo_total_pages.set(str(paut_pages + photo_pages))

    def _export_paut_session(self):
        """[NEW] 현재 PAUT 미리보기 데이터를 CSV로 내보내기"""
        if not self.paut_extracted_data:
            messagebox.showwarning("데이터 없음", "내보낼 데이터가 없습니다.")
            return
        
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            title="PAUT 세션 저장",
            initialfile=f"PAUT_Session_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        )
        if not path: return

        try:
            df = pd.DataFrame(self.paut_extracted_data)
            # UTF-8-SIG for Excel compatibility
            df.to_csv(path, index=False, encoding='utf-8-sig')
            self.log(f"💾 PAUT 세션 저장 완료: {os.path.basename(path)}")
            messagebox.showinfo("저장 완료", f"현재 세션이 저장되었습니다.\n{os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("저장 오류", f"세션 저장 중 오류 발생: {e}")

    def _import_paut_session(self):
        """[NEW] 저장된 PAUT 세션 CSV 불러오기"""
        path = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            title="PAUT 세션 불러오기"
        )
        if not path: return

        try:
            df = pd.read_csv(path)
            # Ensure mandatory columns or basic structure
            new_data = df.to_dict('records')
            
            # Simple normalization (handle NaNs from CSV)
            normalized_data = []
            for i, d in enumerate(new_data):
                clean_dict = {}
                for k, v in d.items():
                    if pd.isna(v): clean_dict[k] = ""
                    else: clean_dict[k] = v
                
                # [FIX] Ensure mandatory attributes for UI filtering
                if 'date_filtered' not in clean_dict: clean_dict['date_filtered'] = True
                if 'line_filtered' not in clean_dict: clean_dict['line_filtered'] = True
                if 'size_filtered' not in clean_dict: clean_dict['size_filtered'] = True
                if 'selected' not in clean_dict: clean_dict['selected'] = True
                if 'order_index' not in clean_dict: clean_dict['order_index'] = i
                
                normalized_data.append(clean_dict)
            
            self.paut_extracted_data = normalized_data
            self.file_info_vars['PAUT'].set(f"📂 로드됨: {os.path.basename(path)} (총 {len(normalized_data)}건)")
            self.update_date_listbox("PAUT")
            self.populate_preview(self.paut_extracted_data, mode="PAUT")
            self.log(f"📂 PAUT 세션 불러오기 완료: {os.path.basename(path)} ({len(normalized_data)}건)")
        except Exception as e:
            messagebox.showerror("불러오기 오류", f"세션 불러오기 중 오류 발생: {e}")

    def _reset_paut_manual_inputs(self):
        """[NEW] 개별 판정 입력 필드 초기화 (두께 T 등 공통 설정값 제외)"""
        for k in ['h', 'l', 'd', 'z1', 'z2', 'L1', 'L2', 'D_Upper', 'D_Lower']:
            if k in self.paut_manual_vars:
                self.paut_manual_vars[k].set("")
        self.paut_manual_vars['nature'].set("")
        self.paut_manual_vars['loc'].set("-")
        self.paut_res_label.config(text="결과 대기", fg="black", bg="#f3f4f6")

    def prepare_next_paut_sheet(self, wb, prev_title, page_num):
        source_sheet = wb[prev_title]
        new_sheet = wb.copy_worksheet(source_sheet)
        new_sheet.title = f"PAUT_Report_{page_num+1:03d}"
        
        # [FIX] openpyxl's copy_worksheet does not copy images. Manually copy them to preserve Logos.
        try:
            import io
            import copy
            from openpyxl.drawing.image import Image as OpenpyxlImage
            for idx, img in enumerate(source_sheet._images):
                try:
                    img_bytes = img._data()
                    
                    # Fix original image (since _data() closes its file pointer)
                    fixed_orig = OpenpyxlImage(io.BytesIO(img_bytes))
                    if hasattr(img, 'anchor'): fixed_orig.anchor = copy.copy(img.anchor)
                    source_sheet._images[idx] = fixed_orig
                    
                    # Add to new sheet
                    new_img = OpenpyxlImage(io.BytesIO(img_bytes))
                    if hasattr(img, 'anchor'): new_img.anchor = copy.copy(img.anchor)
                    new_sheet.add_image(new_img)
                except Exception as inner_e:
                    self.log(f"개별 로고 복사 오류: {inner_e}")
        except Exception as e:
            self.log(f"PAUT 로고 복사 오류: {e}")
        
        start_row = int(self.config.get('PAUT_START_ROW', 11))
        end_row = int(self.config.get('PAUT_DATA_END_ROW', 40))
        for r in range(start_row, end_row + 1):
            for c in range(1, 51): 
                try:
                    new_sheet.cell(row=r, column=c).value = None
                except Exception:
                    pass
        
        return new_sheet

    def _apply_paut_inspector_signature(self, ws):
        """PAUT 갑지의 검사자 이름과 서명을 박광복으로 교체한다."""
        inspector_name = str(ws["D45"].value or "").replace(" ", "")
        if inspector_name not in {"박원준", "박광복"}:
            return False

        signature_path = os.path.join(SRC_ROOT or "", "signs", "박광복.png")
        if not os.path.exists(signature_path):
            self.log(f"[WARNING] 검사자 서명 파일을 찾을 수 없습니다: {signature_path}")
            return False

        ws["D45"] = "박 광 복"
        ws["D45"].alignment = Alignment(
            horizontal="center", vertical="center", shrink_to_fit=True
        )

        # 기존 검사자 서명(I44:K45)의 위치와 크기는 양식 그대로 유지한다.
        inspector_anchor = None
        retained_images = []
        for image in ws._images:
            anchor = getattr(image, "anchor", None)
            marker = getattr(anchor, "_from", None)
            if (
                marker is not None
                and 7 <= marker.col <= 10
                and 43 <= marker.row <= 44
            ):
                inspector_anchor = anchor
                continue
            retained_images.append(image)

        if inspector_anchor is None:
            self.log("[WARNING] 기존 검사자 서명 위치를 찾을 수 없습니다.")
            return False

        import copy
        ws._images = retained_images
        try:
            transparent_stream = self._white_background_to_transparent_png(
                signature_path
            )
            signature = XLImage(transparent_stream)
            # workbook 저장 시까지 BytesIO가 닫히지 않도록 참조를 유지한다.
            signature._pmi_source_stream = transparent_stream
        except Exception as exc:
            self.log(f"[WARNING] 검사자 서명 투명 처리 실패: {exc}")
            signature = XLImage(signature_path)
        signature.anchor = copy.copy(inspector_anchor)
        ws.add_image(signature)
        return True

    def _make_header_stamp_transparent(self, ws):
        """보고서 머리글의 붉은 도장에서 흰 배경을 제거한다."""
        import copy

        replaced_images = []
        changed = False
        for image in ws._images:
            anchor = getattr(image, "anchor", None)
            marker = getattr(anchor, "_from", None)
            is_header_stamp = (
                marker is not None
                and 12 <= marker.col <= 22
                and 0 <= marker.row <= 7
                and 0.65 <= (float(image.width) / max(float(image.height), 1.0)) <= 1.35
            )
            if not is_header_stamp:
                replaced_images.append(image)
                continue

            try:
                source_bytes = image._data()
                with PILImage.open(io.BytesIO(source_bytes)) as source:
                    rgba = source.convert("RGBA")
                    transparent_pixels = []
                    for red, green, blue, alpha in rgba.getdata():
                        # 붉은 인영만 남기고 JPEG 압축으로 생긴 흰색/회색 배경도 제거한다.
                        if red > 100 and red - green > 12 and red - blue > 12:
                            transparent_pixels.append((red, green, blue, alpha))
                        else:
                            transparent_pixels.append((red, green, blue, 0))
                    rgba.putdata(transparent_pixels)
                    output = io.BytesIO()
                    rgba.save(output, format="PNG")
                    output.seek(0)

                stamp = XLImage(output)
                stamp.anchor = copy.copy(anchor)
                replaced_images.append(stamp)
                changed = True
            except Exception as exc:
                self.log(f"[WARNING] 도장 배경 투명 처리 실패: {exc}")
                replaced_images.append(image)

        if changed:
            ws._images = replaced_images
        return changed


    def _get_photo_log_stamp(self):
        """PAUT 양식의 도장을 추출해 투명 PNG 데이터로 반환한다."""
        template_path = self.paut_template_file_path.get().strip()
        if not template_path or not os.path.exists(template_path):
            return None

        template_wb = None
        try:
            template_wb = openpyxl.load_workbook(template_path, keep_vba=True)
            for sheet in template_wb.worksheets:
                for image in sheet._images:
                    anchor = getattr(image, "anchor", None)
                    marker = getattr(anchor, "_from", None)
                    is_header_stamp = (
                        marker is not None
                        and 12 <= marker.col <= 22
                        and 0 <= marker.row <= 7
                        and 0.65 <= (float(image.width) / max(float(image.height), 1.0)) <= 1.35
                    )
                    if not is_header_stamp:
                        continue

                    with PILImage.open(io.BytesIO(image._data())) as source:
                        rgba = source.convert("RGBA")
                        pixels = []
                        for red, green, blue, alpha in rgba.getdata():
                            if red > 100 and red - green > 12 and red - blue > 12:
                                pixels.append((red, green, blue, alpha))
                            else:
                                pixels.append((red, green, blue, 0))
                        rgba.putdata(pixels)
                        output = io.BytesIO()
                        rgba.save(output, format="PNG")
                        return output.getvalue()
        except Exception as exc:
            self.log(f"[PhotoLog] 도장 이미지 준비 실패: {exc}")
        finally:
            if template_wb is not None:
                template_wb.close()
        return None

    def col_to_num(self, col_str):
        """[NEW] 엑셀 열 이름(A, B, AA...)을 숫자 인덱스(1, 2, 27...)로 변환"""
        if not col_str: return 0
        s = str(col_str).strip().upper().replace("열", "").replace("행", "").strip()
        if s.isdigit(): return int(s)
        
        res = 0
        for i, char in enumerate(reversed(s)):
            if ord('A') <= ord(char) <= ord('Z'):
                res += (ord(char) - ord('A') + 1) * (26 ** i)
            else:
                return 0 # Invalid
        return res

    def _create_tree_context_menu(self):
        """미리보기 트리뷰용 공통 컨텍스트 메뉴 초기화"""
        self.ctx_menu = tk.Menu(self.root, tearoff=0)
        # 기본 항목 기입 (show_context_menu에서 command 동적 바인딩)
        self.ctx_menu.add_command(label="선택 항목 체크/해제 토글 (Toggle Check)")
        self.ctx_menu.add_command(label="현재 셀 내용만 복사 (Copy Cell)")
        self.ctx_menu.add_command(label="선택 행 전체 복사 (Copy Row)")
        self.ctx_menu.add_command(label="이 컬럼 전체 복사 (Copy Column)")
        self.ctx_menu.add_command(label="셀 내용 붙여넣기 (Paste)")
        self.ctx_menu.add_separator()
        self.ctx_menu.add_command(label="위로 이동 (Move Up)")
        self.ctx_menu.add_command(label="아래로 이동 (Move Down)")
        self.ctx_menu.add_separator()
        self.ctx_menu.add_command(label="행 추가 (Add Row)")
        self.ctx_menu.add_command(label="선택 삭제 (Delete Selected)")
        self.ctx_menu.add_separator()
        self.ctx_menu.add_command(label="선택 항목 ISO 병합 (Merge ISO)")
        self.ctx_menu.add_command(label="선택 항목 Joint 병합 (Merge Joint)")
        self.ctx_menu.add_command(label="선택 항목 ISO 병합 해제 (Ungroup ISO)")
        self.ctx_menu.add_command(label="선택 항목 Joint 병합 해제 (Ungroup Joint)")
        self.ctx_menu.add_separator()
        self.ctx_menu.add_command(label="선택 항목 일괄 변경 (Bulk Update)")
        # PAUT 전용 (기본은 숨김)
        self.ctx_menu.add_separator()
        self.ctx_menu.add_command(label="컬럼 설정 (Column Manager)")

    def _create_entry_context_menu(self):
        self.entry_popup = tk.Menu(self.root, tearoff=0)
        self.entry_popup.add_command(label="복사 (Copy)", command=self._copy_entry)
        self.entry_popup.add_command(label="붙여넣기 (Paste)", command=self._paste_entry)
        self.entry_popup.add_separator()
        self.entry_popup.add_command(label="잘라내기 (Cut)", command=self._cut_entry)
        
        self.root.bind_class("Entry", "<Button-3>", self._show_entry_popup)
        self.root.bind_class("TEntry", "<Button-3>", self._show_entry_popup)

    def _show_entry_popup(self, event):
        self.current_entry = event.widget
        self._popup_selection_range = None
        if self.current_entry.select_present():
            self._popup_selection_range = (self.current_entry.index("sel.first"), self.current_entry.index("sel.last"))
        self.entry_popup.tk_popup(event.x_root, event.y_root)

    def _copy_entry(self):
        try:
            if getattr(self, 'current_entry', None):
                if self.current_entry.select_present():
                    first, last = self.current_entry.index("sel.first"), self.current_entry.index("sel.last")
                elif getattr(self, '_popup_selection_range', None):
                    first, last = self._popup_selection_range
                else:
                    return
                text = self.current_entry.get()[first:last]
                self.root.clipboard_clear()
                self.root.clipboard_append(text)
                self.root.update()
        except Exception: pass

    def _paste_entry(self):
        try:
            text = self.root.clipboard_get()
            if getattr(self, 'current_entry', None):
                if getattr(self, '_popup_selection_range', None):
                    first, last = self._popup_selection_range
                    self.current_entry.delete(first, last)
                    self.current_entry.insert(first, text)
                elif self.current_entry.select_present():
                    self.current_entry.delete("sel.first", "sel.last")
                    self.current_entry.insert(tk.INSERT, text)
                else:
                    self.current_entry.insert(tk.INSERT, text)
        except Exception: pass

    def _cut_entry(self):
        self._copy_entry()
        try:
            if getattr(self, 'current_entry', None):
                if getattr(self, '_popup_selection_range', None):
                    first, last = self._popup_selection_range
                    self.current_entry.delete(first, last)
                elif self.current_entry.select_present():
                    self.current_entry.delete("sel.first", "sel.last")
        except Exception: pass

    def _excel_anchor_to_px(self, anchor, scale):
        """Roughly map Excel cell to A4 pixel coordinates."""
        import re
        m = re.match(r"([A-Z]+)(\d+)", anchor.upper())
        if not m: return 0, 0
        
        col_str, row_str = m.groups()
        # Col to Num (A=0, B=1...)
        col = 0
        for char in col_str: col = col * 26 + (ord(char) - ord('A'))
        row = int(row_str) - 1
        
        # Rough estimates for A4 grid (210mm / 10 cols, 297mm / 50 rows)
        # Adjusting to feel more like typical report templates
        x = (col * 15) * scale # 15mm per col
        y = (row * 6) * scale  # 6mm per row
        return x, y

    def _create_setting_grid(self, parent, context):
        # [REFINED] More organized layout for logo settings
        items = [
            ("SITCO 로고", f"SITCO_{context}"), 
            ("서울검사 로고", f"SEOUL_{context}"), 
            ("바닥글 우측 (Footer Right)", f"FOOTER_{context}"), 
            ("바닥글 좌측 (Footer Left)", f"FOOTER_PT_{context}")
        ]
        
        next_row = 0
        for i, (label, key_prefix) in enumerate(items):
            # [HYPER-REACTIVE] Force columnspan=1 to match weighted col 0 of parent
            short_label = label.split(" ")[0].replace("로고", "")
            block = tk.LabelFrame(parent, text=f" {short_label} ", padx=2, pady=1, background="#ffffff", font=("Malgun Gothic", 8, "bold"))
            # [FIX] Offset row by 1 to leave room for Gapji Info at row 0 if present
            block.grid(row=i + 1, column=0, sticky='ew', pady=1, padx=2)
            
            frame = tk.Frame(block, background="#ffffff")
            frame.pack(fill='x')
            frame.columnconfigure(1, weight=5, minsize=0) # Path entry weight
            for c in [4, 6, 8, 10]: frame.columnconfigure(c, weight=1, minsize=0) # Coords weight
            
            tk.Label(frame, text="P:", width=2, anchor='e', background="#ffffff", font=("Malgun Gothic", 7)).grid(row=0, column=0)
            v_path = tk.StringVar(value=self.config.get(f"{key_prefix}_PATH", ""))
            ttk.Entry(frame, textvariable=v_path, width=15, exportselection=False).grid(row=0, column=1, sticky='ew', padx=1) # [FIX] Width 1 -> 15
            ttk.Button(frame, text="..", width=2, command=lambda v=v_path: self._browse_file(v, [("Images", "*.png;*.jpg;*.jpeg")])).grid(row=0, column=2, padx=1)
            self.setting_vars[f"{key_prefix}_PATH"] = v_path
            
            for idx, (coord, key_suffix) in enumerate([("X", "X"), ("Y", "Y"), ("W", "W"), ("H", "H")]):
                tk.Label(frame, text=f"{coord}:", width=1, anchor='e', background="#ffffff", font=("Malgun Gothic", 7)).grid(row=0, column=3+idx*2)
                v = tk.StringVar(value=str(self.config.get(f"{key_prefix}_{key_suffix}", "0.0")))
                ttk.Entry(frame, textvariable=v, width=6, exportselection=False).grid(row=0, column=4+idx*2, sticky='ew', padx=1) # [FIX] Width 1 -> 6
                self.setting_vars[f"{key_prefix}_{key_suffix}"] = v
            next_row = i + 1
        return next_row
            
    def _create_margin_settings(self, parent, context, use_pack=True):
        # [FIX] Adaptive layout to avoid TclError (grid vs pack conflict)
        frame = ttk.LabelFrame(parent, text=" 인쇄 및 여백 (Margins) ", padding=2)
        if use_pack:
            frame.pack(fill='x', pady=(5, 0))
        else:
            # Inside Notebook tabs, we usually use grid. 
            # We assume it's placed at the bottom of the existing grid.
            frame.grid(row=100, column=0, columnspan=10, sticky='ew', pady=(10, 0))
        
        # Force 0 minsize for elasticity
        for col in range(10): frame.columnconfigure(col, minsize=0)
        # [FIX] Distribute expansion only to entries
        for col in [0, 2, 4, 6, 8]: frame.columnconfigure(col, weight=0) # Labels
        for col in [1, 3, 5, 7, 9]: frame.columnconfigure(col, weight=1) # Entries

        m_items = [("T", "TOP"), ("B", "BOTTOM"), ("L", "LEFT"), ("R", "RIGHT")]
        for i, (sh_lbl, key) in enumerate(m_items):
            ttk.Label(frame, text=f"{sh_lbl}:", font=("Arial", 7)).grid(row=0, column=i*2, sticky='e', padx=1)
            v = tk.StringVar(value=str(self.config.get(f"MARGIN_{context}_{key}", "0.2")))
            ent = ttk.Entry(frame, textvariable=v, width=1)
            ent.grid(row=0, column=i*2+1, sticky='ew', padx=1)
            ent.bind("<Return>", lambda e: self.root.focus_set())
            self.setting_vars[f"MARGIN_{context}_{key}"] = v
            
        ttk.Label(frame, text="S:", font=("Arial", 7)).grid(row=0, column=8, sticky='e', padx=1)
        v_s = tk.StringVar(value=str(self.config.get(f"PRINT_SCALE_{context}", "95")))
        ent_s = ttk.Entry(frame, textvariable=v_s, width=1)
        ent_s.grid(row=0, column=9, sticky='ew', padx=1)
        ent_s.bind("<Return>", lambda e: self.root.focus_set())
        self.setting_vars[f"PRINT_SCALE_{context}"] = v_s

        # [NEW] Print Area Row
        area_f = ttk.Frame(frame)
        area_f.grid(row=1, column=0, columnspan=10, sticky='ew', pady=(2, 0))
        ttk.Label(area_f, text="Area:", font=("Arial", 7)).pack(side='left', padx=1)
        v_a = tk.StringVar(value=str(self.config.get(f"PRINT_AREA_{context}", "")))
        ent_a = ttk.Entry(area_f, textvariable=v_a)
        ent_a.pack(side='left', fill='x', expand=True, padx=1)
        self.setting_vars[f"PRINT_AREA_{context}"] = v_a
        
        # Adjusters Row
        sub = ttk.Frame(frame)
        sub.grid(row=2, column=0, columnspan=10, sticky='ew', pady=2)
        for c in [1, 3, 5, 7]: sub.columnconfigure(c, weight=1, minsize=0)
        for c in [0, 2, 4, 6]: sub.columnconfigure(c, weight=0, minsize=0)
        
        labels = [("Rows:", f"CUSTOM_ROWS_{context}"), ("H:", f"CUSTOM_ROW_HEIGHT_{context}"), 
                  ("Cols:", f"CUSTOM_COLS_{context}"), ("W:", f"CUSTOM_COL_WIDTH_{context}")]
        
        for i, (lbl, key) in enumerate(labels):
            ttk.Label(sub, text=lbl, font=("Arial", 7)).grid(row=0, column=i*2, sticky='e')
            defaultv = "16.5" if "HEIGHT" in key else ("10.0" if "WIDTH" in key else "")
            v = tk.StringVar(value=str(self.config.get(key, defaultv)))
            ent = ttk.Entry(sub, textvariable=v, width=1)
            ent.grid(row=0, column=i*2+1, sticky='ew', padx=1)
            ent.bind("<FocusOut>", lambda e: self.save_settings())
            ent.bind("<Return>", lambda e: [self.root.focus_set(), self.save_settings()])
            self.setting_vars[key] = v

    def _create_column_mapping_ui(self, parent, mode, mapping_items):
        """[NEW] 전 모듈 공용 엑셀 컬럼 매핑 및 이름 설정 UI (스크롤바 추가)"""
        # Outer Frame for Scrollbar
        outer_frame = tk.Frame(parent, background="#f9fafb")
        outer_frame.pack(fill='both', expand=True)
        
        canvas = tk.Canvas(outer_frame, background="#f9fafb", highlightthickness=0)
        scrollbar = ttk.Scrollbar(outer_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, background="#f9fafb")

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        # Mouse wheel support
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        container = ttk.LabelFrame(scrollable_frame, text=f" {mode} 성적서 및 미리보기 컬럼 설정 ", padding=10)
        container.pack(fill='both', expand=True, pady=5, padx=10)
        
        # Grid settings
        container.columnconfigure(2, weight=1)
        container.columnconfigure(5, weight=1)
        
        # Header labels
        ttk.Label(container, text="항목", font=("Malgun Gothic", 8, "bold")).grid(row=0, column=0, sticky='w')
        ttk.Label(container, text="엑셀열 | 표시이름", font=("Malgun Gothic", 8, "bold")).grid(row=0, column=1, columnspan=2, sticky='w')
        ttk.Label(container, text="항목", font=("Malgun Gothic", 8, "bold")).grid(row=0, column=3, sticky='w')
        ttk.Label(container, text="엑셀열 | 표시이름", font=("Malgun Gothic", 8, "bold")).grid(row=0, column=4, columnspan=2, sticky='w')
        
        last_row = 0
        for i, (label, key_idx, def_idx, key_name, def_name, internal_id) in enumerate(mapping_items):
            row = (i // 2) + 1
            last_row = max(last_row, row)
            col_offset = (i % 2) * 3
            
            # 1. Label
            ttk.Label(container, text=label, font=("Malgun Gothic", 8)).grid(row=row, column=col_offset, sticky='e', padx=2, pady=2)
            
            # 2. Excel Col #
            v_idx = tk.StringVar(value=str(self.config.get(key_idx, def_idx)))
            ent_idx = ttk.Entry(container, textvariable=v_idx, width=4)
            ent_idx.grid(row=row, column=col_offset+1, sticky='w', padx=2, pady=2)
            ent_idx.bind("<FocusOut>", lambda e: self.save_settings())
            ent_idx.bind("<Return>", lambda e: [self.root.focus_set(), self.save_settings()])
            self.setting_vars[key_idx] = v_idx
            
            # [NEW] Trace to update config in real-time so changes are active immediately
            v_idx.trace_add("write", lambda *args, var=v_idx, k=key_idx: self.config.update({k: var.get()}))
            
            # 3. Display Name
            v_name = tk.StringVar(value=str(self.config.get(key_name, def_name)))
            ent_name = ttk.Entry(container, textvariable=v_name, width=12)
            ent_name.grid(row=row, column=col_offset+2, sticky='ew', padx=2, pady=2)
            ent_name.bind("<FocusOut>", lambda e: self.save_settings())
            ent_name.bind("<Return>", lambda e: [self.root.focus_set(), self.save_settings()])
            self.setting_vars[key_name] = v_name
            
            # Trace to update config and preview tree heading in real-time
            v_name.trace_add("write", lambda *args, tid=internal_id, m=mode, var=v_name, kn=key_name: self._update_mode_heading(m, tid, var, kn))

        ttk.Label(container, text="* 엑셀열: 알파벳(A, B...) 또는 숫자(1, 2...) 입력 (0=제외)", foreground="gray", font=("Malgun Gothic", 8)).grid(row=last_row+1, column=0, columnspan=6, pady=(10, 2))
        
        # [NEW] Save Button
        btn_frame = ttk.Frame(container)
        btn_frame.grid(row=last_row+2, column=0, columnspan=6, pady=5)
        ttk.Button(btn_frame, text="💾 현재 컬럼 설정 저장", command=self.manual_save_settings).pack()

    def _update_mode_heading(self, mode, internal_id, var, key_name):
        """[NEW] 특정 모드의 미리보기 헤더 이름을 실시간으로 업데이트 및 설정 저장"""
        try:
            new_text = var.get()
            # 1. Update config immediately so it survives refreshes
            self.config[key_name] = new_text
            
            # 2. Update existing tree UI
            tree = None
            if mode == "PMI": tree = self.preview_tree
            elif mode == "RT": tree = self.rt_preview_tree
            elif mode == "KOGAS": tree = self.kogas_preview_tree
            elif mode == "PT": tree = self.pt_preview_tree
            
            if tree:
                try: tree.heading(internal_id, text=new_text)
                except: pass
        except Exception:
            pass

    def _show_paut_check_menu(self):
        tree = self.paut_preview_tree
        try:
            self.paut_check_menu.tk_popup(tree.winfo_pointerx(), tree.winfo_rooty() + 32)
        finally:
            self.paut_check_menu.grab_release()

    def _show_mt_check_menu(self):
        tree = self.mt_preview_tree
        try:
            self.mt_check_menu.tk_popup(tree.winfo_pointerx(), tree.winfo_rooty() + 32)
        finally:
            self.mt_check_menu.grab_release()


    def _show_rt_check_menu(self, mode):
        tree = self._get_mode_info(mode)[0]
        try:
            self._rt_check_menus[mode].tk_popup(tree.winfo_pointerx(), tree.winfo_rooty() + 32)
        finally:
            self._rt_check_menus[mode].grab_release()


    def _refresh_rt_filter_headers(self, mode):
        tree = self._get_mode_info(mode)[0]
        columns = list(tree['columns'])
        for key, flag, candidates in [('Date', 'date_filtered', ('Date',)),
                                       ('Dwg', 'drawing_filtered', ('Dwg', 'ISO/Dwg', 'ISO/DWG'))]:
            column = next((value for value in candidates if value in columns), None)
            if column is None:
                continue
            values = {str(item.get(key, '') or '').strip() for item in self._get_mode_info(mode)[2]}
            active = {str(item.get(key, '') or '').strip() for item in self._get_mode_info(mode)[2] if item.get(flag, True)}
            suffix = f' ({len(active)}개)' if values != active else ''
            tree.heading(column, text=('Drawing No.' if key == 'Dwg' else key) + ' ▾' + suffix,
                         command=lambda k=key, f=flag, c=column: self._choose_rt_filter(mode, k, f, c))
        if mode in getattr(self, '_rt_check_menus', {}) and 'selected' in columns:
            tree.heading('selected', text='V ▾', command=lambda: self._show_rt_check_menu(mode))


    def _choose_rt_filter(self, mode, key, flag, column):
        states = {}
        for item in self._get_mode_info(mode)[2]:
            states.setdefault(str(item.get(key, '') or '').strip(), bool(item.get(flag, True)))
        states = dict(sorted(states.items()))
        panel = tk.Frame(self.root, background='white', borderwidth=1, relief='solid')
        self._position_paut_filter_dropdown(panel, column, 560 if key == 'Dwg' else 420, 420,
                                            tree=self._get_mode_info(mode)[0])
        ttk.Label(panel, text=f'{key} — 여러 항목 체크 후 적용').pack(anchor='w', padx=10, pady=8)
        search = tk.StringVar()
        ttk.Entry(panel, textvariable=search).pack(fill='x', padx=10)
        frame = ttk.Frame(panel)
        frame.pack(fill='both', expand=True, padx=10, pady=6)
        choices = tk.Listbox(frame, font=('Malgun Gothic', 10), exportselection=False)
        choices.pack(side='left', fill='both', expand=True)
        scrollbar = ttk.Scrollbar(frame, command=choices.yview)
        scrollbar.pack(side='right', fill='y')
        choices.configure(yscrollcommand=scrollbar.set)
        visible = []
        def refresh(*_args):
            visible[:] = [value for value in states if search.get().strip().casefold() in value.casefold()]
            choices.delete(0, tk.END)
            for value in visible:
                choices.insert(tk.END, ('[v] ' if states[value] else '[ ] ') + (value or '(비어 있음)'))
        def toggle(event):
            index = choices.nearest(event.y) if event.keysym != 'space' else (
                choices.curselection()[0] if choices.curselection() else -1)
            if 0 <= index < len(visible):
                value = visible[index]
                states[value] = not states[value]
                refresh()
                choices.selection_set(index)
            return 'break'
        def set_all(state):
            for value in states:
                states[value] = state
            refresh()
        def apply():
            for item in self._get_mode_info(mode)[2]:
                item[flag] = states.get(str(item.get(key, '') or '').strip(), False)
            self.update_date_listbox(mode=mode)
            self.populate_preview(self._get_mode_info(mode)[2], switch_tab=False, mode=mode)
            panel.destroy()
        choices.bind('<ButtonRelease-1>', toggle)
        choices.bind('<space>', toggle)
        search.trace_add('write', refresh)
        refresh()
        buttons = ttk.Frame(panel)
        buttons.pack(fill='x', padx=8, pady=6)
        for label, command in [('전체 선택', lambda: set_all(True)), ('전체 해제', lambda: set_all(False)),
                               ('적용', apply), ('취소', panel.destroy)]:
            ttk.Button(buttons, text=label, command=command).pack(side='left', padx=2)
        ttk.Button(panel, text='정렬 ↑ / ↓', command=lambda: self.sort_by_column(key, mode=mode)).pack(fill='x', padx=10, pady=4)


    def on_tree_select(self, event):
        """항목 선택 시 체크박스 상태와 동기화 (현재는 드래그 안정성을 위해 드래그 중에는 무시)"""
        pass

    def toggle_item_selection(self, item_id, mode="PMI"):
        """단일 항목의 체크 상태를 토글"""
        tree, idx_map, data, _ = self._get_mode_info(mode)
        
        view_idx = tree.index(item_id)
        if 0 <= view_idx < len(idx_map):
            actual_idx = idx_map[view_idx]
            data[actual_idx]['selected'] = not data[actual_idx].get('selected', True)
            self.populate_preview(data, switch_tab=False, mode=mode)
            
            # 선택 상태 복구
            num_items = len(tree.get_children())
            if view_idx < num_items:
                new_item = tree.get_children()[view_idx]
                tree.selection_set(new_item)
                tree.focus(new_item)

    def toggle_selected_items(self, mode="PMI"):
        """선택된 모든 항목의 체크 상태를 토글"""
        tree, idx_map, data, _ = self._get_mode_info(mode)
        
        selected_ids = tree.selection()
        if not selected_ids: return
        
        for item_id in selected_ids:
            view_idx = tree.index(item_id)
            if 0 <= view_idx < len(idx_map):
                actual_idx = idx_map[view_idx]
                data[actual_idx]['selected'] = not data[actual_idx].get('selected', True)
        
        self.populate_preview(data, switch_tab=False, mode=mode)
        # 선택 보존
        for sid in selected_ids:
            try: tree.selection_add(sid)
            except: pass

    def on_tree_click_checkbox(self, event):
        """Reserved for future coordinate-based checkbox toggle if needed without blocking drag."""
        pass

    def on_tree_double_click(self, event, mode="PMI"):
        """셀(Cell) 더블클릭 시 엑셀처럼 그 자리에서 바로 수동 편집 가능하게 함"""
        mode_info = self._get_mode_info(mode)
        if not mode_info: return
        tree, idx_map, data_list, col_keys = mode_info

        region = tree.identify_region(event.x, event.y)
        if region != "cell": return
        
        column = tree.identify_column(event.x)
        item_id = tree.identify_row(event.y)
        if not item_id or not column: return
        
        # [DYNAMIC] 컬럼 인덱스를 통해 키 추출
        try:
            col_idx = int(column.replace('#', '')) - 1
            if col_idx < 0 or col_idx >= len(col_keys): return
            key = col_keys[col_idx]
        except: return
        
        # 수정 차단 컬럼 (V, 상태 아이콘 등)
        if key in ["selected", "_status", "V", "ST"]:
            if key in ["selected", "V"]: self.toggle_item_selection(item_id, mode)
            return

        # 현재 값 가져오기
        x, y, w, h = tree.bbox(item_id, column)
        view_idx = tree.index(item_id)
        if not (0 <= view_idx < len(idx_map)): return
        actual_idx = idx_map[view_idx]
        
        curr_val = data_list[actual_idx].get(key, "")
        numeric_keys = ["Th'k(mm)", "Start", "End", "Length(mm)", "Upper", "Lower", "Height(mm)", "TCG MAX(%)", "a/l", "a/t"]
        if (key in numeric_keys or (key not in ["No", "Joint No.", "Line No.", "Joint", "Loc", "Grade", "Date", "Dwg", "ISO", "ISO/DWG", "Type of Flaw", "Evaluation", "Remarks", "Result", "Location", "Nature", "selected", "_status", "ST", "V"])) and self.to_float(curr_val) > 0:
            curr_val = f"{self.to_float(curr_val):.2f}"
            
        entry = ttk.Entry(tree, font=("Malgun Gothic", 10), exportselection=True)
        entry.insert(0, str(curr_val))
        entry.select_range(0, tk.END)
        entry.place(x=x, y=y, width=w, height=h)
        entry.focus_set()
        
        def finish_edit(e=None):
            if not entry.winfo_exists(): return
            new_val = entry.get().strip()
            
            # [PMI 특화] 함량 소수점 처리 및 판정 재계산
            if mode == "PMI" and key not in ["No", "Joint", "Loc", "Grade", "Date", "Dwg", "_status", "ST", "V", "selected", "order_index"]:
                f_val = self.to_float(new_val)
                data_list[actual_idx][key] = f"{f_val:.2f}" if f_val > 0 else ""
                # 함량 변경 시 등급 판정 재수행
                new_grade = self.check_material_grade(data_list[actual_idx])
                if new_grade: data_list[actual_idx]['Grade'] = new_grade
            elif mode == "PAUT" and key in ["Th'k(mm)", "Height(mm)", "Length(mm)", "Type of Flaw"]:
                data_list[actual_idx][key] = new_val
                # [PAUT] Recalculate Evaluation and Ratios
                t = self.to_float(data_list[actual_idx].get('Th\'k(mm)', 0))
                h = self.to_float(data_list[actual_idx].get('Height(mm)', 0))
                l = self.to_float(data_list[actual_idx].get('Length(mm)', 0))
                d = self.to_float(data_list[actual_idx].get('Upper', 0))
                nat = data_list[actual_idx].get('Type of Flaw', 'Slag')
                eval_mode = self.paut_eval_mode.get()
                
                res, _ = self.evaluate_paut_flaw(t, h, l, d, nat, mode=eval_mode)
                data_list[actual_idx]['Evaluation'] = res
                if l > 0 and h > 0: data_list[actual_idx]['a/l'] = f"{h/l:.3f}"
                else: data_list[actual_idx]['a/l'] = ""
                if t > 0 and h > 0: data_list[actual_idx]['a/t'] = f"{h/t:.3f}"
                else: data_list[actual_idx]['a/t'] = ""
                
                # [NEW] 합격 시 결함 정보 지우기 (단, 수동으로 결함 종류를 입력하는 중이면 유지)
                if "Accept" in res and key != "Type of Flaw":
                    data_list[actual_idx]['Type of Flaw'] = ""
                    data_list[actual_idx]['a/l'] = ""
                    data_list[actual_idx]['a/t'] = ""
            else:
                data_list[actual_idx][key] = new_val

            self.populate_preview(data_list, switch_tab=False, mode=mode)
            self.log(f"📝 {mode} {key} 수정: {new_val}")
            entry.destroy()

        entry.bind("<Return>", finish_edit)
        entry.bind("<FocusOut>", lambda e: finish_edit())
        entry.bind("<Escape>", lambda e: entry.destroy())

    def on_delete_key(self, event, mode="PMI"):
        """Delete 키 선택 시 현재 선택된 셀(들)의 내용을 지움 (엑셀 방식)"""
        mode_info = self._get_mode_info(mode)
        if not mode_info: return
        tree, idx_map, data_list, col_keys = mode_info
        
        selected = tree.selection()
        if not selected: return
        
        # 드래그된 컬럼 범위 계산
        s_col = getattr(self, 'start_col', None)
        e_col = getattr(self, 'end_col', None)
        l_col = getattr(self, 'last_clicked_col', None)
        
        if not s_col or not e_col or s_col == e_col:
            target_col_ids = [l_col or s_col or "#4"]
        else:
            try:
                s_idx = int(s_col.replace('#', '')) - 1
                e_idx = int(e_col.replace('#', '')) - 1
                col_start, col_end = min(s_idx, e_idx), max(s_idx, e_idx)
                target_col_ids = [f"#{i+1}" for i in range(col_start, col_end + 1)]
            except:
                target_col_ids = [l_col or s_col or "#4"]
        
        key_map = {f"#{i+1}": k for i, k in enumerate(col_keys)}
        count = 0
        for item_id in selected:
            view_idx = tree.index(item_id)
            if 0 <= view_idx < len(idx_map):
                actual_idx = idx_map[view_idx]
                for c_id in target_col_ids:
                    key = key_map.get(c_id)
                    if key and key not in ["selected", "_status"]:
                        data_list[actual_idx][key] = ""
                        count += 1
        
        if count > 0:
            self.populate_preview(data_list, switch_tab=False, mode=mode)
            self.log(f"🧹 {mode} 데이터 {count}건 내용 삭제 완료")

    def show_context_menu(self, event, mode="PMI"):
        """우클릭 컨텍스트 메뉴 표시 및 현재 모드에 따라 명령 구성"""
        tree = self._get_mode_info(mode)[0]
        col = tree.identify_column(event.x)
        self.last_clicked_col = col 
        item_id = tree.identify_row(event.y)
        
        # 선택 상태 보정 (우클릭한 항목으로 선택 변경)
        if item_id:
            if item_id not in tree.selection():
                tree.selection_set(item_id)
            tree.focus(item_id)

        # 메뉴 명령 동적 업데이트
        self.ctx_menu.entryconfigure("선택 항목 체크/해제 토글 (Toggle Check)", command=lambda: self.toggle_selected_items(mode))
        self.ctx_menu.entryconfigure("현재 셀 내용만 복사 (Copy Cell)", command=lambda: self.copy_cell(mode, target="cell"))
        self.ctx_menu.entryconfigure("선택 행 전체 복사 (Copy Row)", command=lambda: self.copy_cell(mode, target="row"))
        self.ctx_menu.entryconfigure("이 컬럼 전체 복사 (Copy Column)", command=lambda: self.copy_cell(mode, target="column"))
        self.ctx_menu.entryconfigure("셀 내용 붙여넣기 (Paste)", command=lambda: self.paste_cell(mode))
        
        self.ctx_menu.entryconfigure("위로 이동 (Move Up)", command=lambda: self.move_item(-1, mode))
        self.ctx_menu.entryconfigure("아래로 이동 (Move Down)", command=lambda: self.move_item(1, mode))
        
        self.ctx_menu.entryconfigure("행 추가 (Add Row)", command=lambda: self.add_item(mode))
        self.ctx_menu.entryconfigure("선택 삭제 (Delete Selected)", command=lambda: self.delete_item(mode))
        
        # PMI 전용/특화 기능 제어
        is_pmi = (mode == "PMI")
        self.ctx_menu.entryconfigure("선택 항목 ISO 병합 (Merge ISO)", command=lambda: self.merge_selected_iso(mode), state="normal" if is_pmi else "disabled")
        self.ctx_menu.entryconfigure("선택 항목 Joint 병합 (Merge Joint)", command=lambda: self.merge_selected_joint(mode), state="normal" if is_pmi else "disabled")
        self.ctx_menu.entryconfigure("선택 항목 ISO 병합 해제 (Ungroup ISO)", command=lambda: self.ungroup_selected_iso(mode), state="normal" if is_pmi else "disabled")
        self.ctx_menu.entryconfigure("선택 항목 Joint 병합 해제 (Ungroup Joint)", command=lambda: self.ungroup_selected_joint(mode), state="normal" if is_pmi else "disabled")
        self.ctx_menu.entryconfigure("선택 항목 일괄 변경 (Bulk Update)", command=lambda: self.show_bulk_update_dialog(mode))

        # 컬럼 설정 기능 (모든 모드 지원)
        self.ctx_menu.entryconfigure("컬럼 설정 (Column Manager)", command=lambda: self.manage_columns(mode))
        self.ctx_menu.entryconfigure("컬럼 설정 (Column Manager)", state="normal")

        self.ctx_menu.post(event.x_root, event.y_root)

    def update_date_listbox(self, mode="PMI"):
        if mode == "RT":
            listbox = self.rt_date_listbox
            data = self.rt_extracted_data
        elif mode == "KOGAS":
            listbox = self.kogas_date_listbox
            data = self.kogas_extracted_data
        elif mode == "PT":
            listbox = self.pt_date_listbox
            data = self.pt_extracted_data
        elif mode == "MT":
            listbox = self.mt_date_listbox
            data = self.mt_extracted_data
        elif mode == "PAUT":
            listbox = self.paut_date_listbox
            data = self.paut_extracted_data
        else:
            listbox = self.date_listbox
            data = self.extracted_data
        
        if listbox is None or not listbox.winfo_exists(): return
        
        listbox.delete(0, tk.END)
        # [NEW] 데이터 기반으로 날짜별 선택 상태(date_filtered) 수집
        date_status = {}
        for item in data:
            dt = str(item.get('Date', '') or '').strip()
            if not dt:
                continue
            if dt not in date_status:
                date_status[dt] = item.get('date_filtered', True)
        
        dates = sorted(list(date_status.keys()))
        for d in dates:
            is_v = date_status.get(d, True)
            prefix = "[v]" if is_v else "[ ]"
            listbox.insert(tk.END, f"{prefix} {d}")
            if is_v: listbox.select_set(tk.END)
        self.log(
            f"📅 {mode} 날짜 필터 갱신: {listbox.size()}개 "
            f"({', '.join(dates) if dates else '날짜 없음'})"
        )
        if mode == "PMI":
            self.update_pmi_loc_listbox()
        elif mode == "PAUT":
            self.update_paut_filter_listboxes()
        elif mode == 'MT':
            self._refresh_mt_filter_headers()
        elif mode in ('RT', 'KOGAS'):
            self._refresh_rt_filter_headers(mode)

    def _position_paut_filter_dropdown(self, popup, column, width, height, tree=None):
        """Place an in-app filter panel directly below its column header."""
        previous = getattr(self, '_paut_filter_dropdown', None)
        if previous is not None and previous.winfo_exists():
            previous.destroy()
        self._paut_filter_dropdown = popup
        tree = tree if tree is not None else self.paut_preview_tree
        columns = list(tree['displaycolumns'])
        if columns == ['#all']:
            columns = list(tree['columns'])
        widths = [tree.column(key, 'width') for key in columns]
        offset = sum(widths[:columns.index(column)]) if column in columns else 0
        offset -= int(tree.xview()[0] * sum(widths))
        app_width, app_height = self.root.winfo_width(), self.root.winfo_height()
        x = max(0, min(tree.winfo_rootx() - self.root.winfo_rootx() + offset, app_width - width))
        header_height = 32
        found_header = False
        for row_y in range(min(100, tree.winfo_height())):
            region = tree.identify_region(min(20, max(1, tree.winfo_width() - 2)), row_y)
            if region == 'heading':
                found_header = True
            elif found_header:
                header_height = row_y
                break
        y = max(0, tree.winfo_rooty() - self.root.winfo_rooty() + header_height)
        width, height = min(width, app_width), min(height, max(1, app_height - y))
        def outside_click(event):
            if not (popup.winfo_rootx() <= event.x_root < popup.winfo_rootx() + popup.winfo_width()
                    and popup.winfo_rooty() <= event.y_root < popup.winfo_rooty() + popup.winfo_height()):
                popup.destroy()
                return 'break'
        popup.bind('<ButtonPress-1>', outside_click)
        popup.bind('<Escape>', lambda _event: popup.destroy())
        def show_dropdown():
            if not popup.winfo_exists():
                return
            popup.place(x=x, y=y, width=width, height=height)
            popup.update_idletasks()
            popup.lift()
            popup.focus_force()
            popup.grab_set()
        # Build the contents before displaying the panel within the main window.
        popup.after_idle(show_dropdown)

    def _refresh_mt_filter_headers(self):
        tree = self.mt_preview_tree
        columns = list(tree['columns'])
        for key, flag, candidates in [('Date', 'date_filtered', ('Date',)),
                                       ('Dwg', 'drawing_filtered', ('Dwg', 'ISO/Dwg', 'ISO/DWG'))]:
            column = next((value for value in candidates if value in columns), None)
            if column is None:
                continue
            values = {str(item.get(key, '') or '').strip() for item in self.mt_extracted_data}
            active = {str(item.get(key, '') or '').strip() for item in self.mt_extracted_data if item.get(flag, True)}
            suffix = f' ({len(active)}개)' if values != active else ''
            tree.heading(column, text=key + ' ▾' + suffix,
                         command=lambda k=key, f=flag, c=column: self._choose_mt_filter(k, f, c))
        if hasattr(self, 'mt_check_menu') and 'V' in columns:
            tree.heading('V', text='V ▾', command=self._show_mt_check_menu)

    def _choose_mt_filter(self, key, flag, column):
        states = {}
        for item in self.mt_extracted_data:
            states.setdefault(str(item.get(key, '') or '').strip(), bool(item.get(flag, True)))
        states = dict(sorted(states.items()))
        panel = tk.Frame(self.root, background='white', borderwidth=1, relief='solid')
        self._position_paut_filter_dropdown(panel, column, 560 if key == 'Dwg' else 420, 420,
                                            tree=self.mt_preview_tree)
        ttk.Label(panel, text=f'{key} — 여러 항목 체크 후 적용').pack(anchor='w', padx=10, pady=8)
        search = tk.StringVar()
        ttk.Entry(panel, textvariable=search).pack(fill='x', padx=10)
        frame = ttk.Frame(panel)
        frame.pack(fill='both', expand=True, padx=10, pady=6)
        choices = tk.Listbox(frame, font=('Malgun Gothic', 10), exportselection=False)
        choices.pack(side='left', fill='both', expand=True)
        scrollbar = ttk.Scrollbar(frame, command=choices.yview)
        scrollbar.pack(side='right', fill='y')
        choices.configure(yscrollcommand=scrollbar.set)
        visible = []
        def refresh(*_args):
            visible[:] = [value for value in states if search.get().strip().casefold() in value.casefold()]
            choices.delete(0, tk.END)
            for value in visible:
                choices.insert(tk.END, ('[v] ' if states[value] else '[ ] ') + (value or '(비어 있음)'))
        def toggle(event):
            index = choices.nearest(event.y) if event.keysym != 'space' else (
                choices.curselection()[0] if choices.curselection() else -1)
            if 0 <= index < len(visible):
                value = visible[index]
                states[value] = not states[value]
                refresh()
                choices.selection_set(index)
            return 'break'
        def set_all(state):
            for value in states:
                states[value] = state
            refresh()
        def apply():
            for item in self.mt_extracted_data:
                item[flag] = states.get(str(item.get(key, '') or '').strip(), False)
            self.update_date_listbox(mode='MT')
            self.populate_preview(self.mt_extracted_data, switch_tab=False, mode='MT')
            panel.destroy()
        choices.bind('<ButtonRelease-1>', toggle)
        choices.bind('<space>', toggle)
        search.trace_add('write', refresh)
        refresh()
        buttons = ttk.Frame(panel)
        buttons.pack(fill='x', padx=8, pady=6)
        for label, command in [('전체 선택', lambda: set_all(True)), ('전체 해제', lambda: set_all(False)),
                               ('적용', apply), ('취소', panel.destroy)]:
            ttk.Button(buttons, text=label, command=command).pack(side='left', padx=2)
        ttk.Button(panel, text='정렬 ↑ / ↓', command=lambda: self.sort_by_column(key, mode='MT')).pack(fill='x', padx=10, pady=4)

    def _choose_paut_dates(self):
        """Choose multiple dates without changing independent line/size filters."""
        states = {}
        for item in self.paut_extracted_data:
            date = str(item.get('Date', '') or '').strip()
            states.setdefault(date, bool(item.get('date_filtered', True)))
        states = dict(sorted(states.items()))
        tree = self.paut_preview_tree
        dialog = tk.Frame(self.root, background='white', borderwidth=1, relief='solid')
        self._position_paut_filter_dropdown(dialog, 'Date', 420, 440)
        ttk.Label(dialog, text='날짜를 여러 개 체크한 후 적용하세요.').pack(anchor='w', padx=10, pady=8)
        search = tk.StringVar()
        entry = ttk.Entry(dialog, textvariable=search)
        entry.pack(fill='x', padx=10)
        frame = ttk.Frame(dialog)
        frame.pack(fill='both', expand=True, padx=10, pady=8)
        choices = tk.Listbox(frame, exportselection=False, font=('Malgun Gothic', 10))
        choices.pack(side='left', fill='both', expand=True)
        scroll = ttk.Scrollbar(frame, command=choices.yview)
        scroll.pack(side='right', fill='y')
        choices.configure(yscrollcommand=scroll.set)
        visible = []
        def refresh(*_args):
            visible[:] = [date for date in states if search.get().strip() in date]
            choices.delete(0, tk.END)
            for date in visible:
                choices.insert(tk.END, ('[v] ' if states[date] else '[ ] ') + (date or '(날짜 없음)'))
        def toggle(event=None):
            index = choices.nearest(event.y) if event and event.keysym != 'space' else (
                choices.curselection()[0] if choices.curselection() else -1)
            if 0 <= index < len(visible):
                date = visible[index]
                states[date] = not states[date]
                refresh()
                choices.selection_set(index)
            return 'break'
        def set_all(state):
            for date in states:
                states[date] = state
            refresh()
        def apply():
            for item in self.paut_extracted_data:
                item['date_filtered'] = states.get(str(item.get('Date', '') or '').strip(), False)
            self.update_date_listbox(mode='PAUT')
            self.populate_preview(self.paut_extracted_data, switch_tab=False, mode='PAUT')
            self._sync_paut_photo_page_numbers()
            dialog.destroy()
        choices.bind('<ButtonRelease-1>', toggle)
        choices.bind('<space>', toggle)
        search.trace_add('write', refresh)
        refresh()
        buttons = ttk.Frame(dialog)
        buttons.pack(fill='x', padx=10, pady=4)
        for label, command in [('전체 선택', lambda: set_all(True)), ('전체 해제', lambda: set_all(False)),
                               ('적용', apply), ('취소', dialog.destroy)]:
            ttk.Button(buttons, text=label, command=command).pack(side='left', padx=2)
        ttk.Button(dialog, text='날짜 정렬 ↑ / ↓',
                   command=lambda: self.sort_by_column('Date', mode='PAUT')).pack(fill='x', padx=10, pady=(0, 8))
        entry.focus_set()
        dialog.bind('<Escape>', lambda _event: dialog.destroy())

    def update_paut_filter_listboxes(self):
        """Refresh PAUT Line No./Size choices while preserving check states."""
        line_listbox = getattr(self, 'paut_line_listbox', None)
        size_listbox = getattr(self, 'paut_size_listbox', None)
        if not line_listbox or not size_listbox:
            return
        if not line_listbox.winfo_exists() or not size_listbox.winfo_exists():
            return

        def refresh(listbox, key, flag):
            previous = {}
            for idx in range(listbox.size()):
                text = listbox.get(idx)
                value = text.replace("[v] ", "").replace("[ ] ", "")
                previous[value] = text.startswith("[v] ")

            values = sorted({
                str(item.get(key, '') or '').strip()
                for item in self.paut_extracted_data
                if str(item.get(key, '') or '').strip()
            })
            listbox.delete(0, tk.END)
            active_values = set()
            for value in values:
                checked = previous.get(value, True)
                listbox.insert(tk.END, ("[v] " if checked else "[ ] ") + value)
                if checked:
                    active_values.add(value)
            for item in self.paut_extracted_data:
                item_value = str(item.get(key, '') or '').strip()
                item[flag] = item_value in active_values

        refresh(line_listbox, 'Line No.', 'line_filtered')
        refresh(size_listbox, 'Size', 'size_filtered')
        dates = {str(item.get('Date', '') or '').strip() for item in self.paut_extracted_data}
        active_dates = {str(item.get('Date', '') or '').strip() for item in self.paut_extracted_data
                        if item.get('date_filtered', True)}
        suffix = f' ({len(active_dates)}개)' if active_dates != dates else ''
        self.paut_preview_tree.heading('Date', text='Date ▾' + suffix, command=self._choose_paut_dates)
        if hasattr(self, 'paut_check_menu'):
            self.paut_preview_tree.heading('V', text='V ▾', command=self._show_paut_check_menu)
        if hasattr(self, '_choose_paut_line'):
            checked = sum(line_listbox.get(i).startswith('[v] ') for i in range(line_listbox.size()))
            label = self.config.get('PAUT_NAME_LINE', 'Line No.')
            suffix = f' ({checked}개)' if checked != line_listbox.size() else ''
            self.paut_preview_tree.heading('Line No.', text=label + ' ▾' + suffix,
                                           command=self._choose_paut_line)

    def _get_mode_info(self, mode):
        """Helper to get core UI/Data objects for a specific module mode."""
        if mode == "MT":
            return (self.mt_preview_tree, self.mt_item_idx_map, self.mt_extracted_data, self.mt_column_keys)
        if mode == "PMI":
            return (self.preview_tree, self.item_idx_map, self.extracted_data, self.column_keys)
        elif mode == "RT":
            return (self.rt_preview_tree, self.rt_item_idx_map, self.rt_extracted_data, self.rt_column_keys)
        elif mode == "KOGAS":
            return (self.kogas_preview_tree, self.kogas_item_idx_map, self.kogas_extracted_data, self.kogas_column_keys)
        elif mode == "PT":
            return (self.pt_preview_tree, self.pt_item_idx_map, self.pt_extracted_data, self.pt_column_keys)
        elif mode == "PAUT":
            return (self.paut_preview_tree, self.paut_item_idx_map, self.paut_extracted_data, self.paut_column_keys)
        return None

    def copy_cell(self, mode="PMI", target="smart"):
        """선택된 영역의 내용을 엑셀 형식(\t)으로 클립보드에 복사"""
        try:
            mode_info = self._get_mode_info(mode)
            if not mode_info: return
            tree, idx_map, data_list, col_keys = mode_info

            selected = list(tree.selection())
            if not selected: return
            selected.sort(key=lambda x: tree.index(x))

            # [REFINED] 드래그 영역 계산
            s_col = getattr(self, 'start_col', None)
            e_col = getattr(self, 'end_col', None)
            l_col = getattr(self, 'last_clicked_col', None)

            if target == "cell":
                target_col_ids = [l_col or s_col or "#4"]
            elif not s_col or not e_col or s_col == e_col:
                target_col_ids = [l_col or s_col or "#4"]
            else:
                try:
                    s_idx = int(s_col.replace('#', '')) - 1
                    e_idx = int(e_col.replace('#', '')) - 1
                    col_start, col_end = min(s_idx, e_idx), max(s_idx, e_idx)
                    target_col_ids = [f"#{i+1}" for i in range(col_start, col_end + 1)]
                except:
                    target_col_ids = [l_col or s_col or "#4"]

            key_map = {f"#{i+1}": k for i, k in enumerate(col_keys)}
            output = []

            for item_id in selected:
                view_idx = tree.index(item_id)
                if not (0 <= view_idx < len(idx_map)): continue
                actual_idx = idx_map[view_idx]
                item = data_list[actual_idx]
                
                row_vals = []
                for c_id in target_col_ids:
                    key = key_map.get(c_id)
                    if key == "selected":
                        row_vals.append("●" if item.get('selected', True) else "○")
                    elif key == "_status":
                        is_def = False
                        if mode == "PMI" and hasattr(self, 'element_filters'):
                            for f in self.element_filters:
                                fk = f['key'].get().strip()
                                if not fk: continue
                                v = self.to_float(self._get_val_ci(item, fk))
                                fm, fx = self.to_float(f['min'].get()), self.to_float(f['max'].get())
                                if (fm > 0 and v < fm) or (fx > 0 and v > fx):
                                    is_def = True; break
                        row_vals.append("⚠️" if is_def else "✅")
                    else:
                        val = item.get(key, "")
                        if mode == "PMI" and key not in ["No", "Joint", "Loc", "Grade", "Date", "Dwg", "_status", "ST", "V", "selected", "order_index"]:
                            f_val = self.to_float(val)
                            val = f"{f_val:.2f}" if f_val > 0 else ""
                        row_vals.append(str(val))
                
                output.append("\t".join(row_vals))

            self.root.clipboard_clear()
            self.root.clipboard_append("\n".join(output))
            self.log(f"📋 {mode} 복사 완료: {len(selected)}행 x {len(target_col_ids)}열")
        except Exception as e:
            self.log(f"❌ 복사 오류: {e}")

    def paste_cell(self, mode="PMI"):
        """클립보드 내용을 선택된 영역에 스마트하게(Replicate 지원) 붙여넣기"""
        try:
            mode_info = self._get_mode_info(mode)
            if not mode_info: return
            tree, idx_map, data_list, col_keys = mode_info
            
            try: 
                clipboard_val = self.root.clipboard_get()
                if not clipboard_val: return
                # 엑셀 형식(\t, \n) 파싱
                paste_rows = [line.split("\t") for line in clipboard_val.strip().split("\n")]
            except Exception as e:
                self.log(f"⚠️ 클립보드 데이터를 가져올 수 없습니다: {e}")
                return
            
            selected = list(tree.selection())
            if not selected:
                messagebox.showwarning("붙여넣기 오류", "붙여넣을 셀을 선택해주세요.")
                return
            selected.sort(key=lambda x: tree.index(x))
            
            # [SMART AREA] 드래그된 컬럼 범위를 계산
            s_col = getattr(self, 'start_col', None)
            e_col = getattr(self, 'end_col', None)
            l_col = getattr(self, 'last_clicked_col', None)
            
            if not s_col or not e_col or s_col == e_col:
                target_col = l_col or s_col or "#4"
                try:
                    s_idx = int(target_col.replace('#', '')) - 1
                    col_start = s_idx
                    # 단일 셀 선택 시 클립보드 너비만큼 자동 확장
                    max_cols_in_paste = len(paste_rows[0]) if paste_rows else 1
                    col_end = min(len(col_keys) - 1, col_start + max_cols_in_paste - 1)
                except:
                    col_start, col_end = 0, len(col_keys) - 1
            else:
                try:
                    s_idx = int(s_col.replace('#', '')) - 1
                    e_idx = int(e_col.replace('#', '')) - 1
                    col_start, col_end = min(s_idx, e_idx), max(s_idx, e_idx)
                except:
                    col_start, col_end = 0, len(col_keys) - 1
            
            target_col_ids = [f"#{i+1}" for i in range(col_start, col_end + 1)]
            key_map = {f"#{i+1}": k for i, k in enumerate(col_keys)}
            
            # [EXCEL-LIKE REPLICATION]
            # 한 행만 선택했는데 클립보드가 여러 행이면 -> 자동 드래그 (Fill Down)
            all_children = tree.get_children('')
            if len(selected) == 1 and len(paste_rows) > 1:
                curr_idx = all_children.index(selected[0])
                for i in range(1, len(paste_rows)):
                    if curr_idx + i < len(all_children):
                        selected.append(all_children[curr_idx + i])
            
            for r_idx, item_id in enumerate(selected):
                view_idx = tree.index(item_id)
                if not (0 <= view_idx < len(idx_map)): continue
                actual_idx = idx_map[view_idx]
                
                # 행 복제(Replicate): 클립보드 행보다 선택된 행이 많으면 순환하며 붙여넣음
                source_row = paste_rows[r_idx % len(paste_rows)]
                
                for c_idx, c_id in enumerate(target_col_ids):
                    # 열 복제(Replicate): 클립보드 열보다 선택된 열이 많으면 순환하며 붙여넣음
                    cell_val = source_row[c_idx % len(source_row)].strip()
                    key = key_map.get(c_id)
                    
                    if key and key not in ["selected", "_status"]:
                        if mode == "PMI" and key not in ["No", "Joint", "Loc", "Grade", "Date", "Dwg", "_status", "ST", "V", "selected", "order_index"]:
                            f_val = self.to_float(cell_val)
                            data_list[actual_idx][key] = f"{f_val:.2f}" if f_val > 0 else ""
                        else:
                            data_list[actual_idx][key] = cell_val
            
            self.populate_preview(data_list, switch_tab=False, mode=mode)
            self.log(f"📋 {mode} 붙여넣기 완료: {len(selected)}행 x {len(target_col_ids)}열")
        except Exception as e:
            self.log(f"❌ 붙여넣기 오류: {e}")
            messagebox.showerror("붙여넣기 실패", f"데이터를 붙여넣는 중 오류가 발생했습니다:\n{e}")

    def merge_selected_iso(self, mode="PMI"):
        """선택된 항목들의 ISO 번호를 첫 번째 항목의 것으로 통일 (병합 효과)"""
        mode_info = self._get_mode_info(mode)
        if not mode_info: return
        tree, idx_map, data_list, _ = mode_info
        
        selected = tree.selection()
        if len(selected) < 2:
            messagebox.showinfo("알림", "병합할 항목을 2개 이상 선택해주세요.")
            return
        
        selected = sorted(list(selected), key=lambda x: tree.index(x))
        view_idx_0 = tree.index(selected[0])
        actual_idx_0 = idx_map[view_idx_0] if 0 <= view_idx_0 < len(idx_map) else None
        
        k_iso = "ISO" if mode == "PAUT" else "Dwg"
        if actual_idx_0 is not None:
            first_iso = data_list[actual_idx_0].get(k_iso, '')
        else:
            first_iso = ""
        
        for i, item_id in enumerate(selected):
            view_idx = tree.index(item_id)
            if 0 <= view_idx < len(idx_map):
                actual_idx = idx_map[view_idx]
                data_list[actual_idx][k_iso] = first_iso
                if i > 0: data_list[actual_idx]['is_merged_iso'] = True
                else: data_list[actual_idx].pop('is_merged_iso', None)
        
        self.populate_preview(data_list, switch_tab=False, mode=mode)
        self.log(f"🔗 {mode} {len(selected)}개 항목 ISO 병합 완료: {first_iso}")

    def merge_selected_joint(self, mode="PMI"):
        """선택된 항목들의 Joint No를 첫 번째 항목의 것으로 통일 (데이터 완전 변경)"""
        mode_info = self._get_mode_info(mode)
        if not mode_info: return
        tree, idx_map, data_list, _ = mode_info
        
        selected = tree.selection()
        if len(selected) < 2:
            messagebox.showinfo("알림", "병합할 항목을 2개 이상 선택해주세요.")
            return
        
        selected = sorted(list(selected), key=lambda x: tree.index(x))
        view_idx_0 = tree.index(selected[0])
        actual_idx_0 = idx_map[view_idx_0] if 0 <= view_idx_0 < len(idx_map) else None
        
        if actual_idx_0 is not None:
            first_joint = data_list[actual_idx_0].get('Joint', '')
        else:
            first_joint = ""
        
        for i, item_id in enumerate(selected):
            view_idx = tree.index(item_id)
            if 0 <= view_idx < len(idx_map):
                actual_idx = idx_map[view_idx]
                data_list[actual_idx]['Joint'] = first_joint
                if i > 0: data_list[actual_idx]['is_merged_joint'] = True
                else: data_list[actual_idx].pop('is_merged_joint', None)
        
        self.populate_preview(data_list, switch_tab=False, mode=mode)
        self.log(f"🔗 {mode} {len(selected)}개 항목 Joint 병합 완료: {first_joint}")

    def group_selected_joint(self, mode="PMI"):
        """선택된 항목들을 시각적으로만 같은 Joint로 취급하여 묶음 (데이터 유지)"""
        tree = self.rt_preview_tree if mode == "RT" else self.preview_tree
        idx_map = self.rt_item_idx_map if mode == "RT" else self.item_idx_map
        data = self.rt_extracted_data if mode == "RT" else self.extracted_data
        
        selected = tree.selection()
        if len(selected) < 2:
            messagebox.showinfo("알림", "그룹화할 항목을 2개 이상 선택해주세요.")
            return
            
        selected = sorted(list(selected), key=lambda x: tree.index(x))
            
        view_idx_0 = tree.index(selected[0])
        actual_idx_0 = idx_map[view_idx_0] if 0 <= view_idx_0 < len(idx_map) else None
        
        if actual_idx_0 is not None:
            group_val = data[actual_idx_0].get('Joint', '')
        else:
            group_val = ""
            
        for i, item_id in enumerate(selected):
            view_idx = tree.index(item_id)
            if 0 <= view_idx < len(idx_map):
                actual_idx = idx_map[view_idx]
                data[actual_idx]['visual_group_joint'] = group_val
                if i > 0: data[actual_idx]['is_merged_joint'] = True
                else: data[actual_idx].pop('is_merged_joint', None)
                
        self.populate_preview(data, switch_tab=False, mode=mode)
        self.log(f"🔗 {mode} {len(selected)}개 항목 Joint 시각적 그룹화 완료")

    def ungroup_selected_iso(self, mode="PMI"):
        """선택된 항목들의 ISO 시각적 그룹화 연결 끊기"""
        tree = self.rt_preview_tree if mode == "RT" else self.preview_tree
        idx_map = self.rt_item_idx_map if mode == "RT" else self.item_idx_map
        data = self.rt_extracted_data if mode == "RT" else self.extracted_data
        
        selected = tree.selection()
        if not selected: return
            
        for item_id in selected:
            view_idx = tree.index(item_id)
            if 0 <= view_idx < len(idx_map):
                actual_idx = idx_map[view_idx]
                data[actual_idx].pop('visual_group_iso', None)
                data[actual_idx].pop('is_merged_iso', None)
                
        self.populate_preview(data, switch_tab=False, mode=mode)
        self.log(f"✂️ {mode} {len(selected)}개 항목 ISO 병합 해제 완료")

    def ungroup_selected_iso(self, mode="PMI"):
        """선택된 항목들의 시각적 ISO 병합 해제"""
        tree = self.rt_preview_tree if mode == "RT" else self.preview_tree
        idx_map = self.rt_item_idx_map if mode == "RT" else self.item_idx_map
        data = self.rt_extracted_data if mode == "RT" else self.extracted_data
        
        selected = tree.selection()
        if not selected: return
            
        for item_id in selected:
            view_idx = tree.index(item_id)
            if 0 <= view_idx < len(idx_map):
                actual_idx = idx_map[view_idx]
                data[actual_idx].pop('visual_group_iso', None)
                data[actual_idx].pop('is_merged_iso', None)
                
        self.populate_preview(data, switch_tab=False, mode=mode)
        self.log(f"✂️ {mode} {len(selected)}개 항목 ISO 병합 해제 완료")

    def ungroup_selected_joint(self, mode="PMI"):
        """선택된 항목들의 Joint 시각적 그룹화 연결 끊기"""
        tree = self.rt_preview_tree if mode == "RT" else self.preview_tree
        idx_map = self.rt_item_idx_map if mode == "RT" else self.item_idx_map
        data = self.rt_extracted_data if mode == "RT" else self.extracted_data
        
        selected = tree.selection()
        if not selected: return
            
        for item_id in selected:
            view_idx = tree.index(item_id)
            if 0 <= view_idx < len(idx_map):
                actual_idx = idx_map[view_idx]
                data[actual_idx].pop('visual_group_joint', None)
                data[actual_idx].pop('is_merged_joint', None)
                
        self.populate_preview(data, switch_tab=False, mode=mode)
        self.log(f"✂️ {mode} {len(selected)}개 항목 Joint 병합 해제 완료")

    def select_all(self, mode="PMI"):
        """모든 항목 체크"""
        if mode == "RT": data = self.rt_extracted_data
        elif mode == "KOGAS": data = self.kogas_extracted_data
        elif mode == "PT": data = self.pt_extracted_data
        elif mode == "MT": data = self.mt_extracted_data
        elif mode == "PAUT": data = self.paut_extracted_data
        else: data = self.extracted_data
        
        for item in data:
            if (item.get('date_filtered', True) and (mode not in ('MT', 'RT', 'KOGAS') or item.get('drawing_filtered', True))
                    and (mode != "PAUT" or (
                        item.get('line_filtered', True)
                        and item.get('size_filtered', True)))):
                item['selected'] = True
        self.populate_preview(data, switch_tab=False, mode=mode)



    def deselect_all(self, mode="PMI"):
        """모든 항목 체크 해제"""
        if mode == "RT": data = self.rt_extracted_data
        elif mode == "KOGAS": data = self.kogas_extracted_data
        elif mode == "PT": data = self.pt_extracted_data
        elif mode == "MT": data = self.mt_extracted_data
        elif mode == "PAUT": data = self.paut_extracted_data
        else: data = self.extracted_data
        
        for item in data:
            if (item.get('date_filtered', True) and (mode not in ('MT', 'RT', 'KOGAS') or item.get('drawing_filtered', True))
                    and (mode != "PAUT" or (
                        item.get('line_filtered', True)
                        and item.get('size_filtered', True)))):
                item['selected'] = False
        self.populate_preview(data, switch_tab=False, mode=mode)



    def select_deficient_items(self):
        """PMI: 함량 미달인 항목만 자동으로 선택(체크)"""
        if not self.extracted_data or not hasattr(self, 'element_filters'): return
        
        count = 0
        for item in self.extracted_data:
            is_deficient = False
            for f_item in self.element_filters:
                f_key = f_item['key'].get().strip()
                if not f_key: continue
                f_min = round(self.to_float(f_item['min'].get()), 2)
                f_max = round(self.to_float(f_item['max'].get()), 2)
                
                val = round(self.to_float(self._get_val_ci(item, f_key)), 2)
                if (f_min > 0 and val < f_min) or (f_max > 0 and val > f_max):
                    is_deficient = True
                    break
            
            item['selected'] = is_deficient
            if is_deficient: count += 1
            
        self.populate_preview(self.extracted_data, switch_tab=False, mode="PMI")
        self.log(f"✅ 미달 항목 {count}건을 선택했습니다.")

    def reset_pmi_filters(self, refresh=True, clear_search=True):
        """PMI 검색 및 함량 필터 초기화"""
        if clear_search:
            self.pmi_search_loc.set("")
            # 사이드바 지점 필터도 초기화
            if hasattr(self, 'pmi_loc_listbox'):
                for i in range(self.pmi_loc_listbox.size()):
                    val = self.pmi_loc_listbox.get(i)
                    self.pmi_loc_listbox.delete(i)
                    self.pmi_loc_listbox.insert(i, val.replace("[v] ", "[ ] "))
        
        self.pmi_show_deficiency_only.set(False)
        if hasattr(self, 'element_filters'):
            for f in self.element_filters:
                f['min'].set("")
                f['max'].set("")
        if refresh:
            self.populate_preview(self.extracted_data, switch_tab=False, mode="PMI")
        self.log("🧹 PMI 필터가 초기화되었습니다.")

    def update_pmi_loc_listbox(self):
        """현재 데이터의 Test Location 목록을 추출하여 드롭다운 갱신 (날짜 필터 적용된 항목만)"""
        if not hasattr(self, 'pmi_loc_listbox') or not self.extracted_data: return
        
        # 기존 선택 세트 백업
        selected_locs = set()
        for i in range(self.pmi_loc_listbox.size()):
            val = self.pmi_loc_listbox.get(i)
            if val.startswith("[v] "):
                selected_locs.add(val.replace("[v] ", ""))
        
        # 새 목록 추출
        current_locs = sorted(list(set(str(item.get('Loc', '')).strip() for item in self.extracted_data if item.get('Loc') and item.get('date_filtered', True))))
        
        self.pmi_loc_listbox.delete(0, tk.END)
        for loc in current_locs:
            prefix = "[v] " if loc in selected_locs else "[ ] "
            self.pmi_loc_listbox.insert(tk.END, prefix + loc)

    def add_item(self, mode="PMI"):
        """새 데이터 행 추가"""
        if mode == "RT": data = self.rt_extracted_data
        elif mode == "KOGAS": data = self.kogas_extracted_data
        elif mode == "PT": data = self.pt_extracted_data
        elif mode == "MT": data = self.mt_extracted_data
        elif mode == "PAUT": data = self.paut_extracted_data
        else: data = self.extracted_data
        
        new_item = {'selected': True, 'date_filtered': True, 'order_index': len(data)}
        # 기본 키 초기화
        keys = self.kogas_column_keys if mode == "KOGAS" else self.mt_column_keys if mode == "MT" else self.rt_column_keys if mode == "RT" else (self.pt_column_keys if mode == "PT" else (self.paut_column_keys if mode == "PAUT" else self.column_keys))
        for k in keys:
            if k not in new_item: new_item[k] = ""
            
        data.append(new_item)
        self.populate_preview(data, switch_tab=False, mode=mode)
        self.log(f"➕ {mode} 새 데이터 항목 추가 완료")



    def move_item(self, direction, mode="PMI"):
        """선택된 아이템의 순서를 위/아래로 이동 (필터 상태 대응)"""
        # direction이 문자열인 경우 처리 ("up" -> -1, "down" -> 1)
        if isinstance(direction, str):
            direction = -1 if direction.lower() == "up" else 1
        
        mode_info = self._get_mode_info(mode)
        if not mode_info: return
        tree, idx_map, data, _ = mode_info
        
        selected_ids = tree.selection()
        if not selected_ids: return
        
        view_indices = sorted([tree.index(sid) for sid in selected_ids])
        
        if direction == -1: # 위로
            if view_indices[0] == 0: return
        else: # 아래로
            if view_indices[-1] == len(tree.get_children()) - 1: return
            
        # 데이터 리스트(data)에서 위치 교환
        # Treeview 인덱스를 item_idx_map을 통해 실제 데이터 인덱스로 변환
        actual_indices = [idx_map[vi] for vi in view_indices]
        
        # 이동 방향에 따라 처리 순서 결정
        if direction == -1: # 위로
            # 가장 위 아이템의 위쪽 아이템 찾기 (Treeview상에서)
            target_view_idx = view_indices[0] - 1
            target_actual_idx = idx_map[target_view_idx]
            
            # 선택된 아이템들을 한 칸씩 위로 (target_actual_idx 앞으로)
            items_to_move = [data.pop(ai) for ai in reversed(actual_indices)]
            # target_actual_idx가 팝업으로 인해 밀렸을 수 있으므로 다시 계산
            # 하지만 팝업 시 뒤쪽부터 뺐으므로 target_actual_idx 위치는 유지됨
            for item in items_to_move:
                data.insert(target_actual_idx, item)
        else: # 아래로
            # 가장 아래 아이템의 아래쪽 아이템 찾기
            target_view_idx = view_indices[-1] + 1
            target_actual_idx = idx_map[target_view_idx]
            
            # 선택된 아이템들을 한 칸씩 아래로 (target_actual_idx 뒤로)
            items_to_move = [data.pop(ai) for ai in actual_indices]
            # 팝업으로 인해 인덱스가 당겨졌으므로 target_actual_idx에서 len(items_to_move) 만큼 보정하거나 단순히 insert
            insert_pos = target_actual_idx - len(items_to_move) + 1
            for item in reversed(items_to_move):
                data.insert(insert_pos, item)
                
        self.populate_preview(data, switch_tab=False, mode=mode)
        
        # 선택 상태 복원 (Treeview 아이템이 다시 생성되므로 인덱스로 재선택)
        new_children = tree.get_children()
        for vi in view_indices:
            new_pos = max(0, min(len(new_children)-1, vi + direction))
            tree.selection_add(new_children[new_pos])

    def delete_item(self, mode="PMI"):
        """선택된 아이템 삭제"""
        if mode == "RT":
            tree, idx_map, data = self.rt_preview_tree, self.rt_item_idx_map, self.rt_extracted_data
        elif mode == "KOGAS":
            tree, idx_map, data = self.kogas_preview_tree, self.kogas_item_idx_map, self.kogas_extracted_data
        elif mode == "PT":
            tree, idx_map, data = self.pt_preview_tree, self.pt_item_idx_map, self.pt_extracted_data
        elif mode == "MT":
            tree, idx_map, data = self.mt_preview_tree, self.mt_item_idx_map, self.mt_extracted_data
        elif mode == "PAUT":
            tree, idx_map, data = self.paut_preview_tree, self.paut_item_idx_map, self.paut_extracted_data
        else:
            tree, idx_map, data = self.preview_tree, self.item_idx_map, self.extracted_data
        
        selected_ids = tree.selection()
        if not selected_ids: return
        
        if not messagebox.askyesno("삭제 확인", f"선택한 {len(selected_ids)}개 항목을 정말 삭제하시겠습니까?"):
            return
            
        view_indices = sorted([tree.index(sid) for sid in selected_ids], reverse=True)
        for vi in view_indices:
            actual_idx = idx_map[vi]
            data.pop(actual_idx)
            
        self.populate_preview(data, switch_tab=False, mode=mode)
        self.update_date_listbox(mode)
        self.log(f"🗑️ {mode} {len(selected_ids)}개 항목 삭제 완료")


    def clear_all(self, mode="PMI"):
        """모든 데이터 초기화"""
        if messagebox.askyesno("확인", f"모든 {mode} 데이터를 초기화하시겠습니까?"):
            if mode == "RT": self.rt_extracted_data = []
            elif mode == "KOGAS": self.kogas_extracted_data = []
            elif mode == "PT": self.pt_extracted_data = []
            elif mode == "MT": self.mt_extracted_data = []
            elif mode == "PAUT": self.paut_extracted_data = []
            else: self.extracted_data = []
            self.update_date_listbox(mode)
            self.populate_preview([], mode=mode)
            self.log(f"🧹 모든 {mode} 데이터 초기화 완료")


    def export_to_excel(self, mode="PMI"):
        """현재 미리보기 목록(필터링/선택 반영)을 엑셀 파일로 내보냄 (서식 및 병합 시인성 유지)"""
        mode_info = self._get_mode_info(mode)
        if not mode_info: return
        _, _, data, col_keys = mode_info
        
        if not data:
            messagebox.showwarning("알림", f"{mode} 내보낼 데이터가 없습니다.")
            return

        # 1. 엑셀에 저장할 대상 데이터 필터링
        filter_enabled = self.show_selected_only.get()
        final_list = [
            d for d in data
            if d.get('date_filtered', True)
            and (mode not in ('MT', 'RT', 'KOGAS') or d.get('drawing_filtered', True))
            and (mode != "PAUT" or (
                d.get('line_filtered', True) and d.get('size_filtered', True)))
            and (not filter_enabled or d.get('selected', True))
        ]
        
        if not final_list:
            messagebox.showinfo("알림", f"{mode} 현재 조건에 맞는 데이터가 없습니다.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            initialfile=f"{mode}_Export_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            title=f"{mode} 서식 있는 엑셀 파일로 내보내기"
        )
        if not file_path: return

        try:
            from openpyxl import Workbook
            from openpyxl.styles import Alignment, Border, Side, PatternFill, Font
            import openpyxl.utils

            wb = Workbook()
            ws = wb.active
            ws.title = f"{mode}_Preview"

            # [REFINED] Export with consistent headers from mode_info
            if mode == "PMI":
                # PMI는 특수하게 ST, V 헤더 추가
                headers = ["ST", "V", "No", "Date", "ISO/DWG", "Joint No", "Test Location", "Ni", "Cr", "Mo", "Grade"]
            else:
                # RT/PT/PAUT는 Treeview 컬럼 헤더 그대로 사용
                tree = self._get_mode_info(mode)[0]
                headers = [tree.heading(c)['text'] for c in tree['columns']]
            
            export_rows = []
            for item in final_list:
                row = {}
                if mode == "PMI":
                    # Status Icon
                    is_def = False
                    if hasattr(self, 'element_filters'):
                        for f in self.element_filters:
                            fk = f['key'].get().strip()
                            if not fk: continue
                            v = self.to_float(self._get_val_ci(item, fk))
                            fm, fx = self.to_float(f['min'].get()), self.to_float(f['max'].get())
                            if (fm > 0 and v < fm) or (fx > 0 and v > fx):
                                is_def = True; break
                    
                    row["ST"] = "⚠️" if is_def else "✅"
                    row["V"] = "●" if item.get('selected', True) else "○"
                    row["No"] = item.get('No', '')
                    row["Date"] = item.get('Date', '')
                    row["ISO/DWG"] = item.get('Dwg', '')
                    row["Joint No"] = item.get('Joint', '')
                    row["Test Location"] = item.get('Loc', '')
                    row["Ni"] = round(self.to_float(item.get('Ni', 0)), 2)
                    row["Cr"] = round(self.to_float(item.get('Cr', 0)), 2)
                    row["Mo"] = round(self.to_float(item.get('Mo', 0)), 2)
                    row["Grade"] = item.get('Grade', '')
                else:
                    # Generic mapping for other modes
                    # col_keys: ("_status", "selected", "No", ...)
                    for i, h in enumerate(headers):
                        k = col_keys[i] if i < len(col_keys) else None
                        if k == "_status":
                            row[h] = "✅" 
                        elif k == "selected":
                            row[h] = "●" if item.get('selected', True) else "○"
                        elif k:
                            row[h] = item.get(k, "")
                        else:
                            row[h] = ""
                export_rows.append(row)

            for col_idx, header in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col_idx, value=header)
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
                cell.alignment = Alignment(horizontal='center', vertical='center')

            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                                 top=Side(style='thin'), bottom=Side(style='thin'))
            
            for row_idx, row_data in enumerate(export_rows, 2):
                for col_idx, key in enumerate(headers, 1):
                    val = row_data.get(key, "")
                    cell = ws.cell(row=row_idx, column=col_idx, value=val)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = thin_border

            if mode == "PMI":
                column_widths = {1: 5, 2: 5, 3: 8, 4: 12, 5: 25, 6: 15, 7: 15, 8: 8, 9: 8, 10: 8, 11: 15}
            else:
                # RT Column Widths
                column_widths = {1: 5, 2: 8, 3: 12, 4: 25, 5: 15, 6: 10}
                for i in range(7, 22): column_widths[i] = 4 # Defects
                column_widths[22] = 8 # Result
                column_widths[23] = 12 # Welder
                column_widths[24] = 20 # Remarks
                
            for col_idx, width in column_widths.items():
                if col_idx <= len(headers):
                    col_letter = openpyxl.utils.get_column_letter(col_idx)
                    ws.column_dimensions[col_letter].width = width

            wb.save(file_path)
            self.log(f"📊 {mode} 서식 포함 엑셀 내보내기 완료: {os.path.basename(file_path)}")
            messagebox.showinfo("완료", f"화면 형식과 동일하게 엑셀 파일로 저장했습니다.\n(총 {len(export_rows)} 건)")
        except Exception as e:
            self.log(f"❌ {mode} 엑셀 내보내기 오류: {e}")
            messagebox.showerror("오류", f"엑셀 저장 중 오류가 발생했습니다: {e}")
        finally:
            if 'wb' in locals():
                try: wb.close()
                except: pass

    def _browse_dir(self, var):
        current = var.get()
        init_dir = current if current and os.path.exists(current) else RESOURCE_DIR
        path = filedialog.askdirectory(initialdir=init_dir)
        if path: var.set(path)

    def _update_file_info(self, mode, path):
        """Extracts and formats file metadata for the UI header."""
        if not path or not os.path.exists(path):
            self.file_info_vars[mode].set("📄 파일을 선택해주세요.")
            return

        try:
            fname = os.path.basename(path)
            fsize = os.path.getsize(path) / 1024 # KB
            mtime = os.path.getmtime(path)
            dt_str = datetime.datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M')
            
            size_str = f"{fsize:.1f}KB" if fsize < 1024 else f"{fsize/1024:.2f}MB"
            self.file_info_vars[mode].set(f"📄 파일: {fname}  |  📏 크기: {size_str}  |  📅 수정일: {dt_str}")
        except Exception as e:
            self.file_info_vars[mode].set(f"⚠️ 파일 정보 읽기 오류: {str(e)}")

    def _sync_all_file_infos(self):
        """Updates all mode headers from current config paths."""
        modes = [
            ('PMI', self.target_file_path.get()),
            ('RT', self.rt_target_file_path.get()),
            ('KOGAS', self.kogas_target_file_path.get()),
            ('PT', self.pt_target_file_path.get()),
            ('MT', self.mt_target_file_path.get()),
            ('PAUT', self.paut_target_file_path.get())
        ]
        for mode, path in modes:
            self._update_file_info(mode, path)

    def _browse_file(self, var, types):
        current = var.get()
        first_current = current.split('|')[0] if current else ""
        init_dir = os.path.dirname(first_current) if first_current and os.path.exists(first_current) else BASE_DIR
        paths = filedialog.askopenfilenames(initialdir=init_dir, filetypes=types)
        if paths: 
            var.set("|".join(paths))
            # Find which mode this belongs to and update info
            for mode, m_var in [('PMI', self.target_file_path), ('RT', self.rt_target_file_path), 
                               ('KOGAS', self.kogas_target_file_path), ('PT', self.pt_target_file_path), ('MT', self.mt_target_file_path), 
                               ('PAUT', self.paut_target_file_path)]:
                if var == m_var:
                    try: first_path = paths[0]
                    except: first_path = paths
                    self._update_file_info(mode, first_path)
                    break 

    def to_float(self, val):
        try:
            s = str(val).replace('%', '').strip()
            if not s or s.lower() == 'nan': return 0.0
            return float(s)
        except: return 0.0

    def _get_val_ci(self, item, key):
        """대소문자 구분 없이 딕셔너리에서 값 찾기 (Cr vs CR 등)"""
        if not key or not isinstance(item, dict): return None
        if key in item: return item[key]
        k_lower = key.lower()
        for k in item.keys():
            if str(k).lower() == k_lower:
                return item[k]
        return None

    # --- PMI 재질별 표준 화학성분 스펙 및 판정 DB ---
    PMI_MATERIAL_SPECS = {
        "A312-TP304": {
            "name": "A312-TP304",
            "title": "ASTM A312 TP304",
            "elements": ["Cr", "Ni", "Mo", "-", "-"],
            "specs": [
                "A312-TP304: Cr:18.00 ~ 20.00%, Ni:8.00 ~ 11.00%",
                "ER308L: Cr:19.50 ~ 22.00%, Ni:9.00 ~ 11.00%",
                "A182-F304: Cr:18.00 ~ 20.00%, Ni:8.00 ~ 11.00%"
            ],
            "limits": {"Cr": (18.0, 20.0), "Ni": (8.0, 11.0), "Mo": (0.0, 0.5)}
        },
        "A312-TP304L": {
            "name": "A312-TP304L",
            "title": "ASTM A312 TP304L",
            "elements": ["Cr", "Ni", "Mo", "-", "-"],
            "specs": [
                "A312-TP304L: Cr:18.00 ~ 20.00%, Ni:9.00 ~ 13.00%",
                "ER308L: Cr:19.50 ~ 22.00%, Ni:9.00 ~ 11.00%",
                "A182-F304L: Cr:18.00 ~ 20.00%, Ni:9.00 ~ 13.00%"
            ],
            "limits": {"Cr": (18.0, 20.0), "Ni": (9.0, 13.0), "Mo": (0.0, 0.5)}
        },
        "A312-TP316": {
            "name": "A312-TP316",
            "title": "ASTM A312 TP316",
            "elements": ["Ni", "Cr", "Mo", "-", "-"],
            "specs": [
                "A312-TP316: Ni:10.00 ~ 14.00%, Cr:16.00 ~ 18.00%, Mo:2.00 ~ 3.00%",
                "ER316: Ni:11.00 ~ 14.00%, Cr:18.00 ~ 20.00%, Mo:2.00 ~ 3.00%",
                "A182-F316: Ni:10.00 ~ 14.00%, Cr:16.00 ~ 18.00%, Mo:2.00 ~ 3.00%"
            ],
            "limits": {"Ni": (10.0, 14.0), "Cr": (16.0, 18.0), "Mo": (2.0, 3.0)}
        },
        "A312-TP316L": {
            "name": "A312-TP316L",
            "title": "ASTM A312 TP316L",
            "elements": ["Ni", "Cr", "Mo", "-", "-"],
            "specs": [
                "A312- Ni:10.0 ~ 14.0%, Cr:16.0 ~ 18.0%, Mo:2.0 ~ 3.0%",
                "ER316L- Ni:11.0 ~ 14.0%, Cr:18.0 ~ 20.0%, Mo:2.0 ~ 3.0%",
                "A182-Ni:10.0 ~ 15.0%, Cr:16.0 ~ 18.0%, Mo:2.0 ~ 3.0%"
            ],
            "limits": {"Ni": (10.0, 14.0), "Cr": (16.0, 18.0), "Mo": (2.0, 3.0)}
        },
        "A312-TP321": {
            "name": "A312-TP321",
            "title": "ASTM A312 TP321",
            "elements": ["Ni", "Cr", "Ti", "-", "-"],
            "specs": [
                "A312-TP321: Ni:9.00 ~ 12.00%, Cr:17.00 ~ 19.00%, Ti:0.35 ~ 0.70%",
                "ER321: Ni:9.00 ~ 11.00%, Cr:18.50 ~ 20.50%, Ti:0.30 ~ 0.75%",
                "A182-F321: Ni:9.00 ~ 12.00%, Cr:17.00 ~ 19.00%, Ti:0.35 ~ 0.70%"
            ],
            "limits": {"Ni": (9.0, 12.0), "Cr": (17.0, 19.0), "Ti": (0.35, 0.70)}
        },
        "A312-TP347": {
            "name": "A312-TP347",
            "title": "ASTM A312 TP347",
            "elements": ["Ni", "Cr", "Nb", "-", "-"],
            "specs": [
                "A312-TP347: Ni:9.00 ~ 13.00%, Cr:17.00 ~ 19.00%, Nb:0.80 ~ 1.00%",
                "ER347: Ni:9.00 ~ 11.00%, Cr:19.00 ~ 21.50%, Nb:0.75 ~ 1.00%",
                "A182-F347: Ni:9.00 ~ 13.00%, Cr:17.00 ~ 19.00%, Nb:0.80 ~ 1.00%"
            ],
            "limits": {"Ni": (9.0, 13.0), "Cr": (17.0, 19.0), "Nb": (0.80, 1.00)}
        },
        "DUPLEX (S31803)": {
            "name": "DUPLEX (S31803)",
            "title": "ASTM A790 S31803",
            "elements": ["Cr", "Ni", "Mo", "-", "-"],
            "specs": [
                "S31803: Cr:21.00 ~ 23.00%, Ni:4.50 ~ 6.50%, Mo:2.50 ~ 3.50%",
                "ER2209: Cr:21.50 ~ 23.50%, Ni:7.50 ~ 9.50%, Mo:2.50 ~ 3.50%",
                "A182-F51: Cr:21.00 ~ 23.00%, Ni:4.50 ~ 6.50%, Mo:2.50 ~ 3.50%"
            ],
            "limits": {"Cr": (21.0, 23.0), "Ni": (4.5, 6.5), "Mo": (2.5, 3.5)}
        },
        "INCONEL 625": {
            "name": "INCONEL 625",
            "title": "ASTM B444 N06625",
            "elements": ["Ni", "Cr", "Mo", "Nb", "-"],
            "specs": [
                "Inconel 625: Ni:min 58.0%, Cr:20.0 ~ 23.0%, Mo:8.0 ~ 10.0%, Nb:3.15 ~ 4.15%",
                "ERNiCrMo-3: Ni:min 58.0%, Cr:20.0 ~ 23.0%, Mo:8.0 ~ 10.0%, Nb:3.15 ~ 4.15%",
                "B564-N06625: Ni:min 58.0%, Cr:20.0 ~ 23.0%, Mo:8.0 ~ 10.0%, Nb:3.15 ~ 4.15%"
            ],
            "limits": {"Ni": (58.0, 100.0), "Cr": (20.0, 23.0), "Mo": (8.0, 10.0), "Nb": (3.15, 4.15)}
        },
        # --- 구조용 강재 (COLUMN, GIRDER, BEAM / H-BEAM) ---
        "ASTM A992": {
            "name": "ASTM A992",
            "title": "ASTM A992 (BEAM / COLUMN)",
            "elements": ["Mn", "Si", "V", "Cu", "Cr"],
            "specs": [
                "ASTM A992: Mn:0.50 ~ 1.60%, Si:0.15 ~ 0.40%, V:≤0.15%",
                "WELD (E7018): Mn:0.80 ~ 1.60%, Si:0.20 ~ 0.75%",
                "A992 (Residual): Cu:≤0.60%, Ni:≤0.45%, Cr:≤0.35%"
            ],
            "limits": {"Mn": (0.50, 1.60), "Si": (0.15, 0.40), "V": (0.00, 0.15)}
        },
        "ASTM A572 Gr.50": {
            "name": "ASTM A572 Gr.50",
            "title": "ASTM A572 Gr.50 (GIRDER / PLATE)",
            "elements": ["Mn", "Si", "V", "Nb", "Cr"],
            "specs": [
                "A572 Gr.50: Mn:≤1.35%, Si:0.15 ~ 0.40%, V:0.01 ~ 0.15%",
                "WELD (E7018): Mn:0.80 ~ 1.60%, Si:0.20 ~ 0.75%",
                "A572 Gr.50: Nb(Cb):0.005 ~ 0.05%, Cu:≤0.50%"
            ],
            "limits": {"Mn": (0.50, 1.35), "Si": (0.15, 0.40), "V": (0.01, 0.15)}
        },
        "SM355 (SM490)": {
            "name": "SM355 (SM490)",
            "title": "KS D 3515 SM355 (용접구조용)",
            "elements": ["Mn", "Si", "P", "S", "Cu"],
            "specs": [
                "SM355: Mn:≤1.60%, Si:≤0.55%, P:≤0.030%",
                "WELD (E7018/YGW11): Mn:0.80 ~ 1.60%, Si:0.20 ~ 0.75%",
                "SM355: S:≤0.030%, Ceq:≤0.44%"
            ],
            "limits": {"Mn": (0.50, 1.60), "Si": (0.10, 0.55)}
        },
        "SN355": {
            "name": "SN355",
            "title": "KS D 3866 SN355 (건축구조용 H형강)",
            "elements": ["Mn", "Si", "P", "S", "Cr"],
            "specs": [
                "SN355: Mn:≤1.60%, Si:≤0.55%, P:≤0.030%",
                "WELD (E7018/YGW11): Mn:0.80 ~ 1.60%, Si:0.20 ~ 0.75%",
                "SN355: S:≤0.015%, Ceq:≤0.45%"
            ],
            "limits": {"Mn": (0.50, 1.60), "Si": (0.10, 0.55)}
        },
        "ASTM A36": {
            "name": "ASTM A36",
            "title": "ASTM A36 (탄소 구조강)",
            "elements": ["Mn", "Si", "Cu", "P", "S"],
            "specs": [
                "ASTM A36: Mn:0.60 ~ 1.20%, Si:≤0.40%",
                "WELD (E6013/E7018): Mn:0.40 ~ 1.25%, Si:≤0.60%",
                "ASTM A36: P:≤0.030%, S:≤0.030%"
            ],
            "limits": {"Mn": (0.40, 1.20), "Si": (0.10, 0.40)}
        },
        "SS275 (SS400)": {
            "name": "SS275 (SS400)",
            "title": "KS D 3503 SS275 (일반구조용)",
            "elements": ["Mn", "Si", "P", "S", "-"],
            "specs": [
                "SS275: Mn:≤1.40%, Si:≤0.45%",
                "WELD (E4313/E7016): Mn:0.40 ~ 1.20%, Si:≤0.50%",
                "SS275: P:≤0.040%, S:≤0.040%"
            ],
            "limits": {"Mn": (0.30, 1.40), "Si": (0.05, 0.45)}
        },
        # --- Project PROVIDENCE 현장 공식 성분표 규격 (SA-351 / SA-403 / SFA-5.9) ---
        "A351-CF8M": {
            "name": "A351-CF8M",
            "title": "ASTM A351 CF8M (316 VALVE)",
            "elements": ["Cr", "Ni", "Mo", "-", "-"],
            "specs": [
                "A351-CF8M (Valve): Cr:18.00 ~ 21.00%, Ni:9.00 ~ 12.00%, Mo:2.00 ~ 3.00%",
                "ER316L (Weld): Cr:18.00 ~ 20.00%, Ni:11.00 ~ 14.00%, Mo:2.00 ~ 3.00%",
                "A182-F316L (Flange): Cr:16.00 ~ 18.00%, Ni:10.00 ~ 15.00%, Mo:2.00 ~ 3.00%"
            ],
            "limits": {"Cr": (16.0, 21.0), "Ni": (9.0, 15.0), "Mo": (2.0, 3.0)}
        },
        "A351-CF3M": {
            "name": "A351-CF3M",
            "title": "ASTM A351 CF3M (316L VALVE)",
            "elements": ["Cr", "Ni", "Mo", "-", "-"],
            "specs": [
                "A351-CF3M (Valve): Cr:17.00 ~ 21.00%, Ni:9.00 ~ 13.00%, Mo:2.00 ~ 3.00%",
                "ER316L (Weld): Cr:18.00 ~ 20.00%, Ni:11.00 ~ 14.00%, Mo:2.00 ~ 3.00%",
                "A182-F316L (Flange): Cr:16.00 ~ 18.00%, Ni:10.00 ~ 15.00%, Mo:2.00 ~ 3.00%"
            ],
            "limits": {"Cr": (16.0, 21.0), "Ni": (9.0, 15.0), "Mo": (2.0, 3.0)}
        },
        "A351-CF8": {
            "name": "A351-CF8",
            "title": "ASTM A351 CF8 (304 VALVE)",
            "elements": ["Cr", "Ni", "Mo", "-", "-"],
            "specs": [
                "A351-CF8 (Valve): Cr:18.00 ~ 21.00%, Ni:8.00 ~ 11.00%, Mo:≤0.50%",
                "ER308L (Weld): Cr:19.50 ~ 22.00%, Ni:9.00 ~ 11.00%",
                "A182-F304 (Flange): Cr:18.00 ~ 20.00%, Ni:8.00 ~ 11.00%"
            ],
            "limits": {"Cr": (18.0, 22.0), "Ni": (8.0, 11.0), "Mo": (0.0, 0.75)}
        },
        "A403-WP316L": {
            "name": "A403-WP316L",
            "title": "ASTM A403 WP316L (FITTINGS)",
            "elements": ["Ni", "Cr", "Mo", "-", "-"],
            "specs": [
                "A403-WP316L (Fitting): Ni:10.00 ~ 14.00%, Cr:16.00 ~ 18.00%, Mo:2.00 ~ 3.00%",
                "ER316L (Weld): Ni:11.00 ~ 14.00%, Cr:18.00 ~ 20.00%, Mo:2.00 ~ 3.00%",
                "A312-TP316L (Pipe): Ni:10.00 ~ 14.00%, Cr:16.00 ~ 18.00%, Mo:2.00 ~ 3.00%"
            ],
            "limits": {"Ni": (10.0, 14.0), "Cr": (16.0, 20.0), "Mo": (2.0, 3.0)}
        },
        "A403-WP304L": {
            "name": "A403-WP304L",
            "title": "ASTM A403 WP304L (FITTINGS)",
            "elements": ["Cr", "Ni", "Mo", "-", "-"],
            "specs": [
                "A403-WP304L (Fitting): Cr:18.00 ~ 20.00%, Ni:9.00 ~ 13.00%",
                "ER308L (Weld): Cr:19.50 ~ 22.00%, Ni:9.00 ~ 11.00%",
                "A312-TP304L (Pipe): Cr:18.00 ~ 20.00%, Ni:9.00 ~ 13.00%"
            ],
            "limits": {"Cr": (18.0, 20.0), "Ni": (9.0, 13.0), "Mo": (0.0, 0.5)}
        },
        "A182-F316L": {
            "name": "A182-F316L",
            "title": "ASTM A182 F316L (FLANGE)",
            "elements": ["Ni", "Cr", "Mo", "-", "-"],
            "specs": [
                "A182-F316L (Flange): Ni:10.00 ~ 15.00%, Cr:16.00 ~ 18.00%, Mo:2.00 ~ 3.00%",
                "ER316L (Weld): Ni:11.00 ~ 14.00%, Cr:18.00 ~ 20.00%, Mo:2.00 ~ 3.00%",
                "A312-TP316L (Pipe): Ni:10.00 ~ 14.00%, Cr:16.00 ~ 18.00%, Mo:2.00 ~ 3.00%"
            ],
            "limits": {"Ni": (10.0, 15.0), "Cr": (16.0, 18.0), "Mo": (2.0, 3.0)}
        },
        "A182-F304L": {
            "name": "A182-F304L",
            "title": "ASTM A182 F304L (FLANGE)",
            "elements": ["Cr", "Ni", "Mo", "-", "-"],
            "specs": [
                "A182-F304L (Flange): Cr:18.00 ~ 20.00%, Ni:8.00 ~ 13.00%",
                "ER308L (Weld): Cr:19.50 ~ 22.00%, Ni:9.00 ~ 11.00%",
                "A312-TP304L (Pipe): Cr:18.00 ~ 20.00%, Ni:9.00 ~ 13.00%"
            ],
            "limits": {"Cr": (18.0, 20.0), "Ni": (8.0, 13.0), "Mo": (0.0, 0.5)}
        },
        "ER309L (이재질)": {
            "name": "ER309L (이재질)",
            "title": "AWS A5.9 ER309L (DISSIMILAR WELD)",
            "elements": ["Cr", "Ni", "Mo", "-", "-"],
            "specs": [
                "ER309L (Weld): Cr:23.00 ~ 25.00%, Ni:12.00 ~ 14.00%, Mo:≤0.75%",
                "A312-TP316L / TP304L: Cr:16.00 ~ 20.00%, Ni:10.00 ~ 14.00%",
                "Carbon Steel (Base): Fe:Balance, Mn:≤1.60%"
            ],
            "limits": {"Cr": (16.0, 25.0), "Ni": (8.0, 14.0)}
        },
        # --- 용접봉 및 용접재 단독 규격 (AWS A5.9 / A5.18 / ASME SFA-5.9) ---
        "ER316L": {
            "name": "ER316L",
            "title": "AWS A5.9 ER316L (WELD)",
            "elements": ["Cr", "Ni", "Mo", "-", "-"],
            "specs": [
                "ER316L (Weld): Cr:18.00 ~ 20.00%, Ni:11.00 ~ 14.00%, Mo:2.00 ~ 3.00%",
                "A312-TP316L (Pipe): Cr:16.00 ~ 18.00%, Ni:10.00 ~ 14.00%, Mo:2.00 ~ 3.00%",
                "A182-F316L (Flange): Cr:16.00 ~ 18.00%, Ni:10.00 ~ 15.00%, Mo:2.00 ~ 3.00%"
            ],
            "limits": {"Cr": (18.0, 20.0), "Ni": (11.0, 14.0), "Mo": (2.0, 3.0)}
        },
        "ER308L": {
            "name": "ER308L",
            "title": "AWS A5.9 ER308L (WELD)",
            "elements": ["Cr", "Ni", "Mo", "-", "-"],
            "specs": [
                "ER308L (Weld): Cr:19.50 ~ 22.00%, Ni:9.00 ~ 11.00%, Mo:≤0.75%",
                "A312-TP304L (Pipe): Cr:18.00 ~ 20.00%, Ni:9.00 ~ 13.00%",
                "A182-F304L (Flange): Cr:18.00 ~ 20.00%, Ni:9.00 ~ 13.00%"
            ],
            "limits": {"Cr": (19.5, 22.0), "Ni": (9.0, 11.0), "Mo": (0.0, 0.75)}
        },
        "ER2209": {
            "name": "ER2209",
            "title": "AWS A5.9 ER2209 (DUPLEX WELD)",
            "elements": ["Cr", "Ni", "Mo", "-", "-"],
            "specs": [
                "ER2209 (Duplex Weld): Cr:21.50 ~ 23.50%, Ni:7.50 ~ 9.50%, Mo:2.50 ~ 3.50%",
                "S31803 / S32205 (Pipe): Cr:21.00 ~ 23.00%, Ni:4.50 ~ 6.50%, Mo:2.50 ~ 3.50%",
                "A182-F51 (Flange): Cr:21.00 ~ 23.00%, Ni:4.50 ~ 6.50%, Mo:2.50 ~ 3.50%"
            ],
            "limits": {"Cr": (21.5, 23.5), "Ni": (7.5, 9.5), "Mo": (2.5, 3.5)}
        },
        "ER70S-6": {
            "name": "ER70S-6",
            "title": "AWS A5.18 ER70S-6 / E7018 (WELD)",
            "elements": ["Mn", "Si", "P", "S", "Cu"],
            "specs": [
                "ER70S-6 (Weld Wire): Mn:1.40 ~ 1.85%, Si:0.80 ~ 1.15%",
                "E7018 (SMAW): Mn:0.80 ~ 1.60%, Si:0.20 ~ 0.75%",
                "ASTM A992 / SM355: Mn:0.50 ~ 1.60%, Si:0.15 ~ 0.55%"
            ],
            "limits": {"Mn": (0.80, 1.85), "Si": (0.20, 1.15)}
        }
    }

    PMI_MATERIAL_ALIASES = {
        "TP304": "A312-TP304", "SUS304": "A312-TP304", "SS304": "A312-TP304", "304": "A312-TP304",
        "TP304L": "A312-TP304L", "SUS304L": "A312-TP304L", "SS304L": "A312-TP304L", "304L": "A312-TP304L",
        "TP316": "A312-TP316", "SUS316": "A312-TP316", "SS316": "A312-TP316", "316": "A312-TP316",
        "TP316L": "A312-TP316L", "SUS316L": "A312-TP316L", "SS316L": "A312-TP316L", "316L": "A312-TP316L",
        "TP321": "A312-TP321", "SUS321": "A312-TP321", "SS321": "A312-TP321", "321": "A312-TP321",
        "TP347": "A312-TP347", "SUS347": "A312-TP347", "SS347": "A312-TP347", "347": "A312-TP347",
        "DUPLEX": "DUPLEX (S31803)", "S31803": "DUPLEX (S31803)", "S32205": "DUPLEX (S31803)",
        "INCONEL": "INCONEL 625", "INCONEL625": "INCONEL 625",
        "A992": "ASTM A992", "WBEAM": "ASTM A992", "HBEAM": "ASTM A992", "BEAM": "ASTM A992", "COLUMN": "ASTM A992", "GIRDER": "ASTM A992",
        "A572": "ASTM A572 Gr.50", "A572GR50": "ASTM A572 Gr.50", "GR50": "ASTM A572 Gr.50",
        "SM355": "SM355 (SM490)", "SM490": "SM355 (SM490)", "SM355A": "SM355 (SM490)", "SM355B": "SM355 (SM490)",
        "SN355": "SN355", "SN490": "SN355", "SHN355": "SN355",
        "A36": "ASTM A36",
        "SS275": "SS275 (SS400)", "SS400": "SS275 (SS400)",
        "CF8M": "A351-CF8M", "CF3M": "A351-CF3M", "CF8": "A351-CF8", "CF3": "A351-CF8", "VALVE": "A351-CF8M", "351": "A351-CF8M",
        "WP316L": "A403-WP316L", "WP316": "A403-WP316L", "WP304L": "A403-WP304L", "WP304": "A403-WP304L", "FITTING": "A403-WP316L",
        "ELBOW": "A403-WP316L", "TEE": "A403-WP316L", "REDUCER": "A403-WP316L",
        "엘보": "A403-WP316L", "티": "A403-WP316L", "레듀샤": "A403-WP316L",
        "ER309L": "ER309L (이재질)", "309L": "ER309L (이재질)", "309": "ER309L (이재질)", "이재질": "ER309L (이재질)",
        "ER316L": "ER316L", "E316L": "ER316L", "316LWELD": "ER316L", "ER316": "ER316L",
        "ER308L": "ER308L", "E308L": "ER308L", "308L": "ER308L", "ER308": "ER308L",
        "ER2209": "ER2209", "2209": "ER2209", "E2209": "ER2209",
        "ER70S6": "ER70S-6", "ER70S-6": "ER70S-6", "YGW11": "ER70S-6", "YGW12": "ER70S-6", "E7018": "ER70S-6", "WELD": "ER316L",
        "F316L": "A182-F316L", "F316": "A182-F316L", "F304L": "A182-F304L", "F304": "A182-F304L", "FLANGE": "A182-F316L", "A182": "A182-F316L"
    }

    # 기준 1 / 2 / 3 전용 빠른 선택 목록 (선택한 자재에 맞춰 모든 기준에서 자유롭게 선택 가능)
    PMI_COMMON_SPEC_OPTIONS = [
        "(공란 / 비움)",
        # 배관 (PIPE)
        "A312-TP316L", "A312-TP304L", "A312-TP316", "A312-TP304",
        "A312-TP321", "A312-TP347", "DUPLEX (S31803)", "INCONEL 625",
        # 피팅 (ELBOW / TEE / REDUCER)
        "A403-WP316L (피팅/엘보)", "A403-WP304L (피팅/엘보)",
        # 플랜지 (FLANGE)
        "A182-F316L (플랜지)", "A182-F304L (플랜지)",
        # 밸브 (VALVE)
        "A351-CF8M (주조 밸브)", "A351-CF3M (밸브)", "A351-CF8 (304 밸브)",
        # 용접봉 (WELD)
        "ER316L (용접봉)", "ER308L (용접봉)", "ER309L (이재질)", "ER2209", "ER70S-6", "ERNiCrMo-3",
        # 구조용강 (BEAM)
        "ASTM A992", "ASTM A572 Gr.50", "SM355 (SM490)", "SN355", "ASTM A36", "SS275 (SS400)"
    ]
    PMI_SPEC1_OPTIONS = PMI_COMMON_SPEC_OPTIONS
    PMI_SPEC2_OPTIONS = PMI_COMMON_SPEC_OPTIONS
    PMI_SPEC3_OPTIONS = PMI_COMMON_SPEC_OPTIONS

    def _lookup_pmi_material_spec(self, mat_str):
        if not mat_str: return None
        raw = str(mat_str).strip().upper().replace(" ", "").replace("-", "").replace("_", "")
        # 1. Exact match with primary spec keys
        for k, v in self.PMI_MATERIAL_SPECS.items():
            k_clean = k.upper().replace(" ", "").replace("-", "").replace("_", "")
            if raw == k_clean:
                return v
        # 2. Exact match with alias keys
        for alias, target in self.PMI_MATERIAL_ALIASES.items():
            a_clean = alias.upper().replace(" ", "").replace("-", "").replace("_", "")
            if raw == a_clean:
                return self.PMI_MATERIAL_SPECS.get(target)
        # 3. Partial match (longest alias first to prevent short substring false-positive like '316' inside 'WP316L')
        sorted_aliases = sorted(self.PMI_MATERIAL_ALIASES.items(), key=lambda x: len(x[0]), reverse=True)
        for alias, target in sorted_aliases:
            a_clean = alias.upper().replace(" ", "").replace("-", "").replace("_", "")
            if len(a_clean) >= 3 and a_clean in raw:
                return self.PMI_MATERIAL_SPECS.get(target)
        # 4. Partial match with primary spec keys
        sorted_specs = sorted(self.PMI_MATERIAL_SPECS.items(), key=lambda x: len(x[0]), reverse=True)
        for k, v in sorted_specs:
            k_clean = k.upper().replace(" ", "").replace("-", "").replace("_", "")
            if len(k_clean) >= 4 and (k_clean in raw or raw in k_clean):
                return v
        return None

    def _determine_pmi_target_slot(self, spec):
        """새로 선택된 재질/스펙이 들어갈 최적의 기준 슬롯(1, 2, 3)을 순차적으로 결정"""
        m_name = spec["name"]
        is_weld = any(k in m_name for k in ["ER316", "ER308", "ER309", "ER2209", "ER70S", "ERNICRMO", "E7018", "WELD", "용접"])
        
        # 현재 슬롯별 상태 확인 (공란 여부)
        s1 = self.pmi_chem_specs[0].get().strip() if len(self.pmi_chem_specs) > 0 else ""
        s2 = self.pmi_chem_specs[1].get().strip() if len(self.pmi_chem_specs) > 1 else ""
        s3 = self.pmi_chem_specs[2].get().strip() if len(self.pmi_chem_specs) > 2 else ""

        is_s1_empty = (not s1 or s1 == "-")
        is_s2_empty = (not s2 or s2 == "-")
        is_s3_empty = (not s3 or s3 == "-")

        # 1. 용접봉/용접재인 경우: 기준 2(용접부) 우선 배정
        if is_weld:
            if is_s2_empty or (m_name in s2):
                return 2
            elif is_s3_empty:
                return 3
            elif is_s1_empty:
                return 1
            return 2

        # 2. 일반 자재(모재 / 피팅 / 플랜지 등):
        # 2-1. 기준 1이 비어있거나, 기존 기준 1과 동일한 재질 계열이면 기준 1 갱신
        if is_s1_empty or (m_name in s1):
            return 1

        # 2-2. 기준 1에 이미 다른 자재가 들어가 있는 경우:
        # 차례대로 기준 2가 비어있으면 기준 2로 자동 입력
        if is_s2_empty or (m_name in s2):
            return 2

        # 2-3. 기준 2도 차있으면 기준 3으로 자동 입력
        if is_s3_empty or (m_name in s3):
            return 3

        # 모두 차있으면 기존 기준 보호를 위해 덮어쓰지 않음
        return None

    def apply_pmi_material_preset(self, mat_name, fill_all_specs=False):
        spec = self._lookup_pmi_material_spec(mat_name)
        if not spec: return False
        
        # 1. Chemical elements (Ni, Cr, Mo...)
        for idx, elem in enumerate(spec["elements"]):
            if idx < len(self.pmi_chem_elements):
                self.pmi_chem_elements[idx].set(elem)
                
        # 2. Chemical specs (기준 1, 2, 3 차례대로 스마트 순차 자동 입력)
        m_name = spec["name"]
        target_slot = self._determine_pmi_target_slot(spec)
        if target_slot is not None:
            target_idx = target_slot - 1

            if target_idx < len(self.pmi_chem_specs):
                self.pmi_chem_specs[target_idx].set(spec["specs"][0])

            mat_var = getattr(self, f'pmi_spec{target_slot}_mat', None)
            if mat_var:
                matched = next((opt for opt in self.PMI_COMMON_SPEC_OPTIONS if spec["name"] in opt), spec["name"])
                mat_var.set(matched)

        # fill_all_specs=True(명시적 플래그)일 때만 기준 2, 3을 자동 채움
        if fill_all_specs:
            if len(spec.get("specs", [])) > 1 and len(self.pmi_chem_specs) > 1:
                self.pmi_chem_specs[1].set(spec["specs"][1])
                if hasattr(self, 'pmi_spec2_mat'):
                    raw_s2 = spec["specs"][1].split(":")[0].strip()
                    matched_s2 = next((opt for opt in self.PMI_SPEC2_OPTIONS if raw_s2 in opt or opt in raw_s2), raw_s2)
                    self.pmi_spec2_mat.set(matched_s2)
            if len(spec.get("specs", [])) > 2 and len(self.pmi_chem_specs) > 2:
                self.pmi_chem_specs[2].set(spec["specs"][2])
                if hasattr(self, 'pmi_spec3_mat'):
                    raw_s3 = spec["specs"][2].split(":")[0].strip()
                    matched_s3 = next((opt for opt in self.PMI_SPEC3_OPTIONS if raw_s3 in opt or opt in raw_s3), raw_s3)
                    self.pmi_spec3_mat.set(matched_s3)
                
        # 3. element_filters (Min ~ Max)
        if hasattr(self, 'element_filters'):
            self.element_filters.clear()
            for elem_name, (f_min, f_max) in spec["limits"].items():
                min_str = f"{f_min:.2f}" if f_min > 0 else ""
                max_str = f"{f_max:.2f}" if f_max > 0 else ""
                self.element_filters.append({
                    'key': tk.StringVar(value=elem_name),
                    'min': tk.StringVar(value=min_str),
                    'max': tk.StringVar(value=max_str)
                })
            self._update_pmi_filter_ui()
            
        self._update_pmi_spec_labels()
        self.save_settings()
        self.log(f"📋 PMI 재질 기준 적용: {spec['name']}")
        return True
        
        # 4. If extracted data exists, re-evaluate automatically
        if getattr(self, 'extracted_data', None):
            self.evaluate_pmi_data_compliance(show_msg=False)
            
        return True

    # --- 품명(형태)별 추천 재질 및 기본 프리셋 매핑 ---
    PMI_ITEM_PRESETS = {
        "PIPE": {
            "default_material": "A312-TP316L",
            "materials": ["A312-TP316L", "A312-TP304L", "A312-TP316", "A312-TP304", "A312-TP321", "A312-TP347", "DUPLEX (S31803)", "INCONEL 625"]
        },
        "FLANGE": {
            "default_material": "A182-F316L",
            "materials": ["A182-F316L", "A182-F304L"]
        },
        "VALVE": {
            "default_material": "A351-CF8M",
            "materials": ["A351-CF8M", "A351-CF3M", "A351-CF8"]
        },
        "ELBOW": {
            "default_material": "A403-WP316L",
            "materials": ["A403-WP316L", "A403-WP304L"]
        },
        "TEE": {
            "default_material": "A403-WP316L",
            "materials": ["A403-WP316L", "A403-WP304L"]
        },
        "REDUCER": {
            "default_material": "A403-WP316L",
            "materials": ["A403-WP316L", "A403-WP304L"]
        },
        "FITTING": {
            "default_material": "A403-WP316L",
            "materials": ["A403-WP316L", "A403-WP304L"]
        },
        "WELD": {
            "default_material": "ER316L",
            "materials": ["ER316L", "ER308L", "ER309L (이재질)", "ER2209", "ER70S-6"]
        },
        "BEAM/COLUMN": {
            "default_material": "ASTM A992",
            "materials": ["ASTM A992", "ASTM A572 Gr.50", "SM355 (SM490)", "SN355", "ASTM A36", "SS275 (SS400)"]
        }
    }

    def _on_pmi_item_changed(self, event=None):
        """품명 변경 시 해당 품명 형태(파이프, 엘보, 티, 레듀샤, 단조품, 주물, 피팅, 용접봉 등)에 맞는 재질 드롭다운 재정렬 및 기본 재질 자동 연동"""
        if not hasattr(self, 'gapji_item') or not hasattr(self, 'gapji_material'):
            return
        raw_item = self.gapji_item.get().strip().upper()
        if not raw_item: return

        target_cat = None
        if any(k in raw_item for k in ["PIPE", "파이프", "배관", "TUBE"]):
            target_cat = "PIPE"
        elif any(k in raw_item for k in ["ELBOW", "엘보"]):
            target_cat = "ELBOW"
        elif any(k in raw_item for k in ["TEE", "티"]):
            target_cat = "TEE"
        elif any(k in raw_item for k in ["REDUCER", "레듀샤", "레듀서"]):
            target_cat = "REDUCER"
        elif any(k in raw_item for k in ["FLANGE", "플랜지", "단조", "FORG"]):
            target_cat = "FLANGE"
        elif any(k in raw_item for k in ["VALVE", "밸브", "주물", "주조", "CAST"]):
            target_cat = "VALVE"
        elif any(k in raw_item for k in ["FITTING", "피팅", "이음쇠"]):
            target_cat = "FITTING"
        elif any(k in raw_item for k in ["WELD", "용접", "용접봉", "와이어", "WIRE", "ELECTRODE"]):
            target_cat = "WELD"
        elif any(k in raw_item for k in ["BEAM", "COLUMN", "GIRDER", "보", "기둥", "거더", "형강"]):
            target_cat = "BEAM/COLUMN"

        if target_cat and target_cat in self.PMI_ITEM_PRESETS:
            preset = self.PMI_ITEM_PRESETS[target_cat]
            rec_materials = list(preset["materials"])
            if hasattr(self, 'pmi_material_combobox') and self.pmi_material_combobox:
                self.pmi_material_combobox['values'] = rec_materials

            current_mat = self.gapji_material.get().strip()
            # 품명 변경 시: 해당 품명의 전용 재질 목록만 드롭다운에 표시하고 기본 재질 힌트 세팅.
            # 하단 기준 1, 2, 3은 사용자가 [입력] 또는 [➕ 다음 기준에 추가] 버튼을 누를 때 반영되도록 보존!
            if not current_mat or current_mat not in rec_materials:
                new_mat = preset["default_material"]
                self.gapji_material.set(new_mat)
            self._update_pmi_spec_labels()
        else:
            # 기타/미지정 품목인 경우 전체 재질 목록 제공
            if hasattr(self, 'pmi_material_combobox') and self.pmi_material_combobox:
                self.pmi_material_combobox['values'] = list(self.PMI_MATERIAL_SPECS.keys())

    def _on_pmi_material_changed(self, event=None):
        """재질 변경 시 품명 상호 연동 및 원소명 동기화 (하단 기준 슬롯은 [입력] 버튼으로 적용)"""
        mat = self.gapji_material.get().strip()
        if not mat: return

        spec = self._lookup_pmi_material_spec(mat)
        if spec:
            # 원소명(Ni, Cr, Mo...) 동기화
            for idx, elem in enumerate(spec["elements"]):
                if idx < len(self.pmi_chem_elements):
                    self.pmi_chem_elements[idx].set(elem)

            # 품명 역연동
            if hasattr(self, 'gapji_item'):
                cur_item = self.gapji_item.get().strip().upper()
                m_name = spec['name']
                new_item = None
                if "A182" in m_name:
                    new_item = "FLANGE"
                elif "A351" in m_name:
                    new_item = "VALVE"
                elif "A403" in m_name:
                    if cur_item in ("ELBOW", "TEE", "REDUCER"):
                        new_item = cur_item
                    else:
                        new_item = "ELBOW"
                elif any(k in m_name for k in ["ER316L", "ER308L", "ER309L", "ER2209", "ER70S"]):
                    new_item = "WELD"
                elif any(k in m_name for k in ["A992", "A572", "SM355", "SN355", "A36", "SS275"]):
                    new_item = "H-BEAM"
                elif "A312" in m_name:
                    new_item = "PIPE"

                if new_item and (not cur_item or cur_item in ("PIPE", "FLANGE", "VALVE", "FITTING", "ELBOW", "TEE", "REDUCER", "WELD", "H-BEAM", "BEAM", "HOT TAPPING")):
                    self.gapji_item.set(new_item)
                    self.save_settings()

    def _get_pmi_spec_item_title(self, slot_idx, mat_txt="", spec_txt=""):
        """기준 슬롯(1, 2, 3)의 재질 및 규격에 대응하는 품명(Item Name)을 도출"""
        mat_clean = mat_txt.strip()
        spec_clean = spec_txt.strip()
        if not mat_clean and not spec_clean:
            return ""
        if any(k in mat_clean for k in ["공란", "비움", "해당없음"]) and (not spec_clean or spec_clean == "-"):
            return ""

        cur_item = self.gapji_item.get().strip().upper() if hasattr(self, 'gapji_item') else ""
        combined = f"{mat_clean} {spec_clean}".upper()

        # 1. 용접봉/용접부
        if any(k in combined for k in ["ER316", "ER308", "ER309", "ER2209", "ER70S", "ERNICRMO", "E7018", "WELD", "용접"]):
            return "용접봉"

        # 2. 플랜지
        if any(k in combined for k in ["A182", "F316", "F304", "F51", "B564", "FLANGE", "플랜지"]):
            return "FLANGE"

        # 3. 밸브
        if any(k in combined for k in ["A351", "CF8M", "CF3M", "CF8", "VALVE", "밸브"]):
            return "VALVE"

        # 4. 피팅 (엘보/티/레듀샤)
        if any(k in combined for k in ["A403", "WP316", "WP304", "ELBOW", "TEE", "REDUCER", "엘보", "티", "레듀샤", "피팅"]):
            if "TEE" in combined or "티" in combined:
                return "TEE"
            if "REDUCER" in combined or "레듀샤" in combined:
                return "REDUCER"
            return "ELBOW"

        # 5. 형강/구조재
        if any(k in combined for k in ["A992", "A572", "SM355", "SN355", "A36", "SS275", "BEAM", "COLUMN", "GIRDER", "형강"]):
            return "H-BEAM"

        # 6. 배관 (PIPE)
        if any(k in combined for k in ["A312", "TP316", "TP304", "TP321", "TP347", "S31803", "PIPE", "배관"]):
            return "PIPE"

        # 7. 괄호 안에 품명 힌트가 있는 경우 (예: "피팅/엘보", "플랜지", "용접봉")
        if "(" in mat_clean and ")" in mat_clean:
            hint = mat_clean.split("(")[1].split(")")[0].strip()
            if hint: return hint

        # 8. 그 외 기본 재질명에서 품명 유추 불가 시
        if mat_clean and not any(k in mat_clean for k in ["공란", "비움"]):
            return mat_clean.split("(")[0].strip()

        return ""

    def _get_combined_pmi_item_names(self):
        """보고서 발행 및 갑지 표기 시 기준(1~6)에 언급된 모든 품명을 중복 없이 순서대로 수집하여 콤마로 연결"""
        collected = []
        # 1. 활성화된 기준 그룹에서 품명 수집
        grouped = self._collect_pmi_active_criteria() if hasattr(self, '_collect_pmi_active_criteria') else []
        for grp in grouped:
            item_txt = grp.get('item', '').strip()
            if item_txt:
                for part in item_txt.split(','):
                    p = part.strip()
                    if p and p not in collected and not any(k in p for k in ("공란", "비움", "-")):
                        collected.append(p)

        # 2. 기준 슬롯들 직접 검사 (grouped에 누락된 품명이 있는 경우 보완)
        specs_list = getattr(self, 'pmi_chem_specs', [])
        for slot_idx, spec_var in enumerate(specs_list, start=1):
            s_val = spec_var.get().strip()
            if not s_val or s_val == "-": continue
            mat_var = getattr(self, f'pmi_spec{slot_idx}_mat', None)
            m_val = mat_var.get().strip() if mat_var else ""
            item_t = self._get_pmi_spec_item_title(slot_idx, m_val, s_val) if hasattr(self, '_get_pmi_spec_item_title') else ""
            if item_t:
                for part in item_t.split(','):
                    p = part.strip()
                    if p and p not in collected and not any(k in p for k in ("공란", "비움", "-")):
                        collected.append(p)

        # 3. 만약 수집된 것이 없다면 상단 기본 정보 폴백
        if not collected:
            base_item = (self.gapji_item.get().strip() if hasattr(self, 'gapji_item') else '') or self.config.get('GAPJI_ITEM', '')
            if base_item:
                return base_item
            return "PIPE"

        return ", ".join(collected)

    def _create_pmi_spec_slot_ui(self, slot_idx, container):
        """지정된 번호(1~6)의 기준 슬롯 UI 위젯 생성"""
        if not hasattr(self, 'pmi_spec_slots_ui'):
            self.pmi_spec_slots_ui = {}
        if slot_idx in self.pmi_spec_slots_ui:
            return

        idx_zero = slot_idx - 1
        while len(self.pmi_chem_specs) <= idx_zero:
            new_v = tk.StringVar(value=self.config.get(f'PMI_CHEM_SPEC_{len(self.pmi_chem_specs)+1}', ""))
            self.pmi_chem_specs.append(new_v)
            self.setting_vars[f'PMI_CHEM_SPEC_{len(self.pmi_chem_specs)}'] = new_v

        mat_var = getattr(self, f'pmi_spec{slot_idx}_mat', None)
        if not mat_var:
            mat_var = tk.StringVar(value=self.config.get(f'PMI_SPEC{slot_idx}_MAT', ''))
            setattr(self, f'pmi_spec{slot_idx}_mat', mat_var)
            self.setting_vars[f'PMI_SPEC{slot_idx}_MAT'] = mat_var

        s_hdr = tk.Frame(container, background="#ffffff")
        s_hdr.pack(fill='x', pady=(2, 1))
        
        lbl = tk.Label(s_hdr, text=f"기준 {slot_idx}:", background="#ffffff", font=("Malgun Gothic", 8, "bold"))
        lbl.pack(side='left')
        setattr(self, f'pmi_spec{slot_idx}_lbl', lbl)

        ttk.Button(s_hdr, text="✕", width=2, command=lambda i=slot_idx: self.clear_pmi_spec(i)).pack(side='right', padx=(2, 0))
        ttk.Button(s_hdr, text="▼입력", width=4, command=lambda i=slot_idx: self.apply_current_to_spec(i)).pack(side='right', padx=(2, 0))

        options = getattr(self, f'PMI_SPEC{slot_idx}_OPTIONS', self.PMI_COMMON_SPEC_OPTIONS)
        combo = ttk.Combobox(
            s_hdr, textvariable=mat_var,
            values=options, width=20, state='readonly'
        )
        combo.pack(side='right')
        combo.bind('<<ComboboxSelected>>', lambda e, i=slot_idx: self._on_pmi_spec_selected(i))
        setattr(self, f'pmi_spec{slot_idx}_combo', combo)

        entry = ttk.Entry(container, textvariable=self.pmi_chem_specs[idx_zero])
        entry.pack(fill='x', padx=2, pady=(0, 3))

        self.pmi_spec_slots_ui[slot_idx] = {
            'header': s_hdr,
            'lbl': lbl,
            'combo': combo,
            'entry': entry,
            'mat_var': mat_var,
            'spec_var': self.pmi_chem_specs[idx_zero]
        }

    def add_pmi_spec_slot(self, show_msg=True):
        """새로운 기준 슬롯(4, 5, 6)을 동적으로 추가 (최대 6개)"""
        cur_count = len(self.pmi_chem_specs)
        if cur_count >= 6:
            if show_msg:
                messagebox.showinfo("알림", "갑지 성적서 1페이지 인쇄 규격상 기준 슬롯은 최대 6개까지 지원됩니다.")
            return None
        
        new_idx = cur_count + 1
        new_spec_var = tk.StringVar(value=self.config.get(f'PMI_CHEM_SPEC_{new_idx}', ""))
        new_mat_var = tk.StringVar(value=self.config.get(f'PMI_SPEC{new_idx}_MAT', ''))
        self.pmi_chem_specs.append(new_spec_var)
        setattr(self, f'pmi_spec{new_idx}_mat', new_mat_var)
        self.setting_vars[f'PMI_CHEM_SPEC_{new_idx}'] = new_spec_var
        self.setting_vars[f'PMI_SPEC{new_idx}_MAT'] = new_mat_var
        
        if hasattr(self, 'pmi_specs_container') and self.pmi_specs_container.winfo_exists():
            self._create_pmi_spec_slot_ui(new_idx, self.pmi_specs_container)
            self._update_pmi_spec_labels()
            if hasattr(self, 'gapji_item') and self.gapji_item.get():
                self._on_pmi_item_changed()
            try: self.root.update_idletasks()
            except: pass
            self.log(f"➕ 기준 슬롯 {new_idx} 추가 완료 (현재 총 {new_idx}개)")
        return new_idx

    def remove_pmi_spec_slot(self, idx):
        """동적으로 추가된 4~6번 슬롯 제거 (중간 슬롯 삭제 시 뒤쪽 슬롯 값을 앞으로 이동)"""
        cur_total = len(self.pmi_chem_specs)
        if idx <= 3 or idx > cur_total:
            return
        
        # 중간 슬롯 삭제 시 뒤쪽 슬롯 값들을 앞으로 한 칸씩 당김
        for s in range(idx, cur_total):
            next_s = s + 1
            cur_spec = self.pmi_chem_specs[next_s - 1].get()
            cur_mat = getattr(self, f'pmi_spec{next_s}_mat', None)
            cur_mat_val = cur_mat.get() if cur_mat else ""
            
            self.pmi_chem_specs[s - 1].set(cur_spec)
            target_mat = getattr(self, f'pmi_spec{s}_mat', None)
            if target_mat: target_mat.set(cur_mat_val)

        # 맨 마지막 슬롯의 UI 및 변수 파괴
        last_idx = cur_total
        slot_data = getattr(self, 'pmi_spec_slots_ui', {}).get(last_idx)
        if slot_data:
            if slot_data.get('header') and slot_data['header'].winfo_exists():
                slot_data['header'].destroy()
            if slot_data.get('entry') and slot_data['entry'].winfo_exists():
                slot_data['entry'].destroy()
            del self.pmi_spec_slots_ui[last_idx]
        
        self.pmi_chem_specs.pop()
        self.setting_vars.pop(f'PMI_CHEM_SPEC_{last_idx}', None)
        self.setting_vars.pop(f'PMI_SPEC{last_idx}_MAT', None)
        self.config.pop(f'PMI_CHEM_SPEC_{last_idx}', None)
        self.config.pop(f'PMI_SPEC{last_idx}_MAT', None)
        self._update_pmi_spec_labels()
        self.save_settings()
        self.log(f"➖ 기준 슬롯 제거 완료 (현재 총 {len(self.pmi_chem_specs)}개)")

    def reset_pmi_spec_slots_to_default(self, show_msg=False):
        """4번 이상 슬롯을 모두 닫아 기본 3개 상태로 원위치"""
        while len(self.pmi_chem_specs) > 3:
            last_idx = len(self.pmi_chem_specs)
            self.remove_pmi_spec_slot(last_idx)
        for i in (4, 5, 6):
            self.config.pop(f'PMI_CHEM_SPEC_{i}', None)
            self.config.pop(f'PMI_SPEC{i}_MAT', None)
        self._update_pmi_spec_labels()
        self.save_settings()
        self.log("↩️ 기준 슬롯을 기본 3개로 원위치했습니다.")
        if show_msg:
            messagebox.showinfo("알림", "기준 슬롯을 기본 3개로 원위치했습니다.")

    def _update_pmi_spec_labels(self):
        """기준 라벨들을 재질명이 아닌 품명(Item Name)으로 동적 갱신"""
        num_slots = len(getattr(self, 'pmi_chem_specs', []))
        for slot_idx in range(1, num_slots + 1):
            lbl = getattr(self, f'pmi_spec{slot_idx}_lbl', None)
            if not lbl: continue
            mat_var = getattr(self, f'pmi_spec{slot_idx}_mat', None)
            spec_var = self.pmi_chem_specs[slot_idx - 1] if slot_idx <= len(self.pmi_chem_specs) else None
            mat_txt = mat_var.get().strip() if mat_var else ""
            spec_txt = spec_var.get().strip() if spec_var else ""
            item_title = self._get_pmi_spec_item_title(slot_idx, mat_txt, spec_txt)
            if item_title:
                lbl.config(text=f"기준 {slot_idx} ({item_title}):")
            else:
                lbl.config(text=f"기준 {slot_idx}:")

    def add_current_to_next_spec(self, show_msg=True):
        """상단에서 선택된 품명 및 재질을 비어있는 다음 기준 슬롯(1 -> N)에 순차 입력 (부족 시 자동 슬롯 확장)"""
        cur_mat = self.gapji_material.get().strip() if hasattr(self, 'gapji_material') else ""
        cur_item = self.gapji_item.get().strip().upper() if hasattr(self, 'gapji_item') else ""
        if not cur_mat:
            if cur_item:
                cur_mat = self._lookup_pmi_material_spec(cur_item)
                if cur_mat: cur_mat = cur_mat["name"]
        if not cur_mat:
            if show_msg:
                messagebox.showinfo("알림", "상단에서 먼저 재질을 선택해주세요.")
            return

        spec = self._lookup_pmi_material_spec(cur_mat)
        if not spec:
            if show_msg:
                messagebox.showinfo("알림", f"선택한 재질({cur_mat})에 대한 표준 규격을 찾을 수 없습니다.")
            return

        # 빈 슬롯(1 -> N) 찾기
        target_slot = None
        for slot in range(1, len(self.pmi_chem_specs) + 1):
            idx = slot - 1
            val = self.pmi_chem_specs[idx].get().strip()
            if not val or val == "-":
                target_slot = slot
                break

        # 기존 슬롯이 모두 채워진 경우 6개 미만이면 자동으로 슬롯 추가
        if target_slot is None:
            if len(self.pmi_chem_specs) < 6:
                target_slot = self.add_pmi_spec_slot(show_msg=show_msg)
            else:
                if show_msg:
                    messagebox.showinfo(
                        "알림",
                        "기준 슬롯 1~6이 이미 모두 채워져 있습니다.\n"
                        "수정을 원하시는 기준 옆의 [▼입력] 버튼을 누르시면 해당 기준이 즉시 변경됩니다."
                    )
                return

        if target_slot:
            self.apply_current_to_spec(target_slot)
            self.log(f"📋 기준 {target_slot}에 입력 완료: {cur_item} {spec['name']}")

    def apply_current_to_spec(self, slot_idx):
        """상단에서 선택된 품명 및 재질을 지정된 기준 슬롯에 즉시 입력"""
        cur_mat = self.gapji_material.get().strip() if hasattr(self, 'gapji_material') else ""
        cur_item = self.gapji_item.get().strip().upper() if hasattr(self, 'gapji_item') else ""
        if not cur_mat:
            if cur_item:
                cur_mat = self._lookup_pmi_material_spec(cur_item)
                if cur_mat: cur_mat = cur_mat["name"]
        if not cur_mat:
            messagebox.showinfo("알림", "상단에서 먼저 재질을 선택해주세요.")
            return

        spec = self._lookup_pmi_material_spec(cur_mat)
        if not spec: return

        idx = slot_idx - 1
        if idx < len(self.pmi_chem_specs):
            spec_text = spec["specs"][0]
            if cur_item and cur_item not in spec_text.upper():
                spec_text = f"{spec['name']} ({cur_item}): " + spec_text.split(":", 1)[-1].strip()
            self.pmi_chem_specs[idx].set(spec_text)

        mat_var = getattr(self, f'pmi_spec{slot_idx}_mat', None)
        if mat_var:
            matched = next((opt for opt in self.PMI_COMMON_SPEC_OPTIONS if spec["name"] in opt), spec["name"])
            mat_var.set(matched)

        if slot_idx == 1 or not getattr(self, 'element_filters', None):
            for e_idx, elem in enumerate(spec["elements"]):
                if e_idx < len(self.pmi_chem_elements):
                    self.pmi_chem_elements[e_idx].set(elem)
            if hasattr(self, 'element_filters'):
                self.element_filters.clear()
                for elem_name, (f_min, f_max) in spec["limits"].items():
                    self.element_filters.append({
                        'key': tk.StringVar(value=elem_name),
                        'min': tk.StringVar(value=f"{f_min:.2f}" if f_min > 0 else ""),
                        'max': tk.StringVar(value=f"{f_max:.2f}" if f_max > 0 else "")
                    })
                self._update_pmi_filter_ui()

        self._update_pmi_spec_labels()
        self.save_settings()
        self.log(f"📋 기준 {slot_idx}에 선택한 자재({spec['name']})를 입력했습니다.")

    def clear_pmi_spec(self, idx):
        """지정된 기준 슬롯을 공란으로 초기화하거나 4번 이상 슬롯은 즉시 닫기(제거)"""
        slot_idx = idx - 1
        if 0 <= slot_idx < len(self.pmi_chem_specs):
            # 4번 이상 확장 슬롯에서 ✕를 누르면 슬롯 자체를 즉시 닫아 원위치
            if idx > 3:
                # idx번부터 마지막까지 모두 비어있거나 현재 슬롯을 지우면 닫힘
                self.remove_pmi_spec_slot(idx)
                return
            self.pmi_chem_specs[slot_idx].set("")
            mat_var = getattr(self, f'pmi_spec{idx}_mat', None)
            if mat_var: mat_var.set("(공란 / 비움)")
            self.log(f"📋 기준 {idx}을 공란으로 비웠습니다.")
        self._update_pmi_spec_labels()
        self.save_settings()

    def clear_all_pmi_specs(self):
        """모든 기준 슬롯을 비우고 4~6번 슬롯은 기본 3개로 자동 원위치"""
        # 4번 이상 슬롯이 열려있다면 전부 닫아서 3개로 축소
        if hasattr(self, 'reset_pmi_spec_slots_to_default'):
            self.reset_pmi_spec_slots_to_default()
        for i in range(1, 4):
            slot_idx = i - 1
            if slot_idx < len(self.pmi_chem_specs):
                self.pmi_chem_specs[slot_idx].set("")
            mat_var = getattr(self, f'pmi_spec{i}_mat', None)
            if mat_var: mat_var.set("(공란 / 비움)")
        self._update_pmi_spec_labels()
        self.save_settings()
        self.log("📋 모든 기준 슬롯을 비우고 기본 3개로 원위치했습니다.")

    def _on_pmi_spec_selected(self, slot_idx, event=None):
        """기준 슬롯 콤보박스 선택 공통 핸들러"""
        mat_var = getattr(self, f'pmi_spec{slot_idx}_mat', None)
        if not mat_var: return
        val = mat_var.get().strip()
        idx = slot_idx - 1
        if idx >= len(self.pmi_chem_specs): return
        if not val or "비움" in val or val == "-":
            self.pmi_chem_specs[idx].set("")
            self.log(f"📋 기준 {slot_idx}: 공란 설정")
            self._update_pmi_spec_labels()
            self.save_settings()
            return
        spec = self._lookup_pmi_material_spec(val)
        if spec:
            self.pmi_chem_specs[idx].set(spec["specs"][0])
            if slot_idx == 1:
                self.gapji_material.set(spec["name"])
                self.apply_pmi_material_preset(spec["name"])
            self.log(f"📋 기준 {slot_idx} 설정: {spec['name']}")
            self.save_settings()
        self._update_pmi_spec_labels()

    def _on_pmi_spec1_selected(self, event=None):
        self._on_pmi_spec_selected(1, event)

    def _on_pmi_spec2_selected(self, event=None):
        self._on_pmi_spec_selected(2, event)

    def _on_pmi_spec3_selected(self, event=None):
        self._on_pmi_spec_selected(3, event)

    def _collect_pmi_active_criteria(self):
        """활성화된 기준 슬롯들로부터 화학성분 규격을 추출하고, 동일한 함량 기준끼리 품목명을 묶음(그룹화)"""
        import re
        raw_list = []
        specs_list = getattr(self, 'pmi_chem_specs', [])
        for slot_idx, spec_var in enumerate(specs_list, start=1):
            spec_txt = spec_var.get().strip()
            if not spec_txt or spec_txt == "-":
                continue
            mat_var = getattr(self, f'pmi_spec{slot_idx}_mat', None)
            mat_txt = mat_var.get().strip() if mat_var else ""
            item_title = self._get_pmi_spec_item_title(slot_idx, mat_txt, spec_txt)
            if not item_title:
                item_title = (self.gapji_item.get().strip() if hasattr(self, 'gapji_item') else '') or "PIPE"

            spec_obj = self._lookup_pmi_material_spec(mat_txt) or self._lookup_pmi_material_spec(spec_txt)
            spec_name = spec_obj.get('title', spec_obj['name']) if spec_obj else (mat_txt or spec_txt.split(':')[0].strip())

            limits = {}
            if spec_obj and spec_obj.get("limits"):
                limits = dict(spec_obj["limits"])
            else:
                matches = re.findall(r'([A-Za-z]{1,2})\s*:\s*([\d\.]+)\s*(?:~|-)\s*([\d\.]+)', spec_txt)
                for el, mi, ma in matches:
                    limits[el] = (float(mi), float(ma))

            raw_list.append({
                'slot': slot_idx,
                'name': spec_name,
                'item': item_title,
                'limits': limits,
                'spec_text': spec_txt
            })

        # 활성 슬롯에 내용이 없으면 상단 기본 정보 폴백
        if not raw_list:
            mat_input = (self.gapji_material.get().strip() if hasattr(self, 'gapji_material') else '') or self.config.get('GAPJI_MATERIAL', '')
            spec_obj = self._lookup_pmi_material_spec(mat_input) if hasattr(self, '_lookup_pmi_material_spec') else None
            item_name = (self.gapji_item.get().strip() if hasattr(self, 'gapji_item') else '') or self.config.get('GAPJI_ITEM', 'PIPE')
            if spec_obj:
                raw_list.append({
                    'slot': 1,
                    'name': spec_obj.get('title', spec_obj['name']),
                    'item': item_name,
                    'limits': dict(spec_obj.get("limits", {})),
                    'spec_text': spec_obj.get("specs", [""])[0]
                })

        # [원칙 1] 동일 함량 자동 묶기 (그룹화)
        grouped = []
        for item_data in raw_list:
            matched_grp = None
            for grp in grouped:
                if grp['limits'] == item_data['limits'] and grp['limits']:
                    matched_grp = grp
                    break
            if matched_grp:
                cur_items = [x.strip() for x in matched_grp['item'].split(',') if x.strip()]
                if item_data['item'] not in cur_items:
                    cur_items.append(item_data['item'])
                    matched_grp['item'] = ", ".join(cur_items)
                cur_names = [x.strip() for x in matched_grp['name'].split(',') if x.strip()]
                if item_data['name'] not in cur_names:
                    cur_names.append(item_data['name'])
                    matched_grp['name'] = ", ".join(cur_names)
            else:
                grouped.append({
                    'name': item_data['name'],
                    'item': item_data['item'],
                    'limits': item_data['limits'],
                    'spec_text': item_data['spec_text']
                })
        return grouped

    def _apply_pmi_acceptance_criteria(self, ws):
        """갑지 1. Acceptance Criteria (%) 표(24~39행, A~T열)를 [방안 B] 가로 비교 매트릭스 형태로 작성"""
        if ws.cell(row=24, column=1).value is None:
            return

        from openpyxl.styles import Font, Alignment, Border, Side
        from openpyxl.utils.cell import range_boundaries

        def _style_range(ws_obj, cell_range, font=None, alignment=None, border=None):
            min_col, min_row, max_col, max_row = range_boundaries(cell_range)
            for row in ws_obj.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
                for cell in row:
                    if font: cell.font = font
                    if alignment: cell.alignment = alignment
                    if border: cell.border = border

        grouped = self._collect_pmi_active_criteria()
        if not grouped:
            return

        # 1. 기존 24~39행, A~T열의 병합 셀 해제 및 내용 초기화
        to_unmerge = [rng for rng in ws.merged_cells.ranges if 24 <= rng.min_row <= 39 and (rng.min_col <= 20 or rng.max_col <= 20)]
        for rng in to_unmerge:
            try: ws.unmerge_cells(str(rng))
            except: pass
        for r in range(24, 40):
            for c in range(1, 21):
                cell = ws.cell(row=r, column=c)
                cell.value = None
                cell.border = Border()

        num_groups = len(grouped)

        # 스타일 정의
        thin_side = Side(style='thin')
        b_thin = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
        f_title = Font(name='Times New Roman', size=11, bold=True)
        # 기준 개수(열 너비)에 따라 헤더 폰트 크기 세분화 (글자 가려짐 완전 방지):
        # - 1개 기준: 10pt
        # - 2개 기준 (7열 너비): 8.5pt
        # - 3개 이상 기준 (5열 너비): 8pt (긴 규격명과 품목명이 여유 있게 한눈에 표시됨)
        if num_groups >= 3:
            hdr_size = 8
        elif num_groups == 2:
            hdr_size = 8.5
        else:
            hdr_size = 10
        f_hdr = Font(name='Times New Roman', size=hdr_size, bold=True)
        f_data = Font(name='Times New Roman', size=10, bold=False)
        f_note = Font(name='Times New Roman', size=9, bold=False)
        align_c = Alignment(horizontal='center', vertical='center', wrap_text=True)

        elem_full_names = {
            'Cr': 'Chromium', 'Ni': 'Nickel', 'Mo': 'Molybdenum',
            'Mn': 'Manganese', 'Si': 'Silicon', 'V': 'Vanadium',
            'Nb': 'Niobium', 'Ti': 'Titanium', 'Cu': 'Copper',
            'P': 'Phosphorus', 'S': 'Sulfur'
        }

        # 대상 원소 선정 (Cr, Ni, Mo 우선, 최대 3개)
        all_elems = []
        for g in grouped:
            for k in g['limits'].keys():
                if k not in all_elems and k != "-":
                    all_elems.append(k)
        target_elems = []
        for p in ["Cr", "Ni", "Mo"]:
            if p in all_elems:
                target_elems.append(p)
        for e in all_elems:
            if e not in target_elems:
                target_elems.append(e)
        if not target_elems:
            target_elems = ["Cr", "Ni", "Mo"]
        target_elems = target_elems[:3]

        def _format_limit(limits_dict, el_key):
            if el_key in limits_dict:
                l_min, l_max = limits_dict[el_key]
                if l_min > 0 and l_max < 900.0:
                    return f"{l_min:.2f}~{l_max:.2f}"
                elif l_min > 0:
                    return f"Min {l_min:.2f}"
                elif l_max < 900.0:
                    return f"Max {l_max:.2f}"
            return "-"

        def _get_hdr_label(grp):
            name = grp.get('name', '')
            item = grp.get('item', '')
            if item:
                return f"{name}\n({item})"
            return name

        # [케이스 1] 고유 기준 1개: 기존 10열(A~J) 단일 표
        if num_groups == 1:
            grp = grouped[0]
            ws.merge_cells('A24:J26')
            ws['A24'] = f"1. Acceptance Criteria (%)\n{grp['name']} — {grp['item']}"
            _style_range(ws, 'A24:J26', font=f_title, alignment=align_c, border=b_thin)

            ws.row_dimensions[27].height = 20.0
            ws.merge_cells('A27:F27')
            ws['A27'] = 'Element'
            _style_range(ws, 'A27:F27', font=f_hdr, alignment=align_c, border=b_thin)

            ws.merge_cells('G27:J27')
            ws['G27'] = 'Composition, %'
            _style_range(ws, 'G27:J27', font=f_hdr, alignment=align_c, border=b_thin)

            for idx, el_k in enumerate(target_elems, start=28):
                full_n = elem_full_names.get(el_k, el_k)
                val_txt = _format_limit(grp['limits'], el_k)
                ws.merge_cells(f'A{idx}:F{idx}')
                ws[f'A{idx}'] = full_n
                _style_range(ws, f'A{idx}:F{idx}', font=f_data, alignment=align_c, border=b_thin)

                ws.merge_cells(f'G{idx}:J{idx}')
                ws[f'G{idx}'] = val_txt
                _style_range(ws, f'G{idx}:J{idx}', font=f_data, alignment=align_c, border=b_thin)

            ws.merge_cells('A31:J31')
            ws['A31'] = '* Tolerance ±10%'
            _style_range(ws, 'A31:J31', font=f_note, alignment=align_c, border=b_thin)

        # [케이스 2] 고유 기준 2개: [방안 B] 가로 비교 2열 매트릭스 (A~F: 원소, G~M: 기준1, N~T: 기준2)
        elif num_groups == 2:
            ws.merge_cells('A24:T26')
            ws['A24'] = '1. Acceptance Criteria (%)'
            _style_range(ws, 'A24:T26', font=f_title, alignment=align_c, border=b_thin)

            # 27행 헤더 2줄 텍스트가 잘리지 않도록 높이 34pt 지정, 빈 32행에서 상쇄하여 총 세로높이 100% 보존
            ws.row_dimensions[27].height = 34.0
            ws.row_dimensions[32].height = 1.77

            ws.merge_cells('A27:F27')
            ws['A27'] = 'Element'
            _style_range(ws, 'A27:F27', font=f_hdr, alignment=align_c, border=b_thin)

            ws.merge_cells('G27:M27')
            ws['G27'] = _get_hdr_label(grouped[0])
            _style_range(ws, 'G27:M27', font=f_hdr, alignment=align_c, border=b_thin)

            ws.merge_cells('N27:T27')
            ws['N27'] = _get_hdr_label(grouped[1])
            _style_range(ws, 'N27:T27', font=f_hdr, alignment=align_c, border=b_thin)

            for idx, el_k in enumerate(target_elems, start=28):
                full_n = elem_full_names.get(el_k, el_k)
                v1 = _format_limit(grouped[0]['limits'], el_k)
                v2 = _format_limit(grouped[1]['limits'], el_k)

                ws.merge_cells(f'A{idx}:F{idx}')
                ws[f'A{idx}'] = full_n
                _style_range(ws, f'A{idx}:F{idx}', font=f_data, alignment=align_c, border=b_thin)

                ws.merge_cells(f'G{idx}:M{idx}')
                ws[f'G{idx}'] = v1
                _style_range(ws, f'G{idx}:M{idx}', font=f_data, alignment=align_c, border=b_thin)

                ws.merge_cells(f'N{idx}:T{idx}')
                ws[f'N{idx}'] = v2
                _style_range(ws, f'N{idx}:T{idx}', font=f_data, alignment=align_c, border=b_thin)

            ws.merge_cells('A31:T31')
            ws['A31'] = '* Tolerance ±10%'
            _style_range(ws, 'A31:T31', font=f_note, alignment=align_c, border=b_thin)

        # [케이스 3] 고유 기준 3개: [방안 B] 가로 비교 3열 매트릭스 (A~E, F~J, K~O, P~T)
        elif num_groups == 3:
            ws.merge_cells('A24:T26')
            ws['A24'] = '1. Acceptance Criteria (%)'
            _style_range(ws, 'A24:T26', font=f_title, alignment=align_c, border=b_thin)

            # 27행 헤더 2줄 텍스트가 잘리지 않도록 높이 34pt 지정, 빈 32행에서 상쇄하여 총 세로높이 100% 보존
            ws.row_dimensions[27].height = 34.0
            ws.row_dimensions[32].height = 1.77

            ws.merge_cells('A27:E27')
            ws['A27'] = 'Element'
            _style_range(ws, 'A27:E27', font=f_hdr, alignment=align_c, border=b_thin)

            ranges_col = [('F', 'J'), ('K', 'O'), ('P', 'T')]
            for g_i, (c_start, c_end) in enumerate(ranges_col):
                rng_txt = f"{c_start}27:{c_end}27"
                ws.merge_cells(rng_txt)
                ws[f"{c_start}27"] = _get_hdr_label(grouped[g_i])
                _style_range(ws, rng_txt, font=f_hdr, alignment=align_c, border=b_thin)

            for idx, el_k in enumerate(target_elems, start=28):
                full_n = elem_full_names.get(el_k, el_k)
                ws.merge_cells(f'A{idx}:E{idx}')
                ws[f'A{idx}'] = full_n
                _style_range(ws, f'A{idx}:E{idx}', font=f_data, alignment=align_c, border=b_thin)

                for g_i, (c_start, c_end) in enumerate(ranges_col):
                    v = _format_limit(grouped[g_i]['limits'], el_k)
                    rng_txt = f"{c_start}{idx}:{c_end}{idx}"
                    ws.merge_cells(rng_txt)
                    ws[f"{c_start}{idx}"] = v
                    _style_range(ws, rng_txt, font=f_data, alignment=align_c, border=b_thin)

            ws.merge_cells('A31:T31')
            ws['A31'] = '* Tolerance ±10%'
            _style_range(ws, 'A31:T31', font=f_note, alignment=align_c, border=b_thin)

        # [케이스 4] 고유 기준 4~6개: 상단(24~31행) 1차 + 하단(32~39행) 2차 분할 2단 매트릭스
        else:
            tier1 = grouped[:3]
            tier2 = grouped[3:6]

            # 1단 (24~31행)
            ws.merge_cells('A24:T26')
            ws['A24'] = '1. Acceptance Criteria (%)'
            _style_range(ws, 'A24:T26', font=f_title, alignment=align_c, border=b_thin)

            # 헤더 높이 32pt씩 부여하고, 25/26행에서 상쇄하여 총 세로높이 100% 보존
            ws.row_dimensions[25].height = 10.885
            ws.row_dimensions[26].height = 10.885
            ws.row_dimensions[27].height = 32.0

            ws.merge_cells('A27:E27')
            ws['A27'] = 'Element'
            _style_range(ws, 'A27:E27', font=f_hdr, alignment=align_c, border=b_thin)

            ranges_col = [('F', 'J'), ('K', 'O'), ('P', 'T')]
            for g_i in range(len(tier1)):
                c_start, c_end = ranges_col[g_i]
                rng_txt = f"{c_start}27:{c_end}27"
                ws.merge_cells(rng_txt)
                ws[f"{c_start}27"] = _get_hdr_label(tier1[g_i])
                _style_range(ws, rng_txt, font=f_hdr, alignment=align_c, border=b_thin)

            for idx, el_k in enumerate(target_elems, start=28):
                full_n = elem_full_names.get(el_k, el_k)
                ws.merge_cells(f'A{idx}:E{idx}')
                ws[f'A{idx}'] = full_n
                _style_range(ws, f'A{idx}:E{idx}', font=f_data, alignment=align_c, border=b_thin)
                for g_i in range(len(tier1)):
                    c_start, c_end = ranges_col[g_i]
                    v = _format_limit(tier1[g_i]['limits'], el_k)
                    rng_txt = f"{c_start}{idx}:{c_end}{idx}"
                    ws.merge_cells(rng_txt)
                    ws[f"{c_start}{idx}"] = v
                    _style_range(ws, rng_txt, font=f_data, alignment=align_c, border=b_thin)

            ws.merge_cells('A31:T31')
            ws['A31'] = '* Tolerance ±10%'
            _style_range(ws, 'A31:T31', font=f_note, alignment=align_c, border=b_thin)

            # 2단 (32~39행)
            ws.merge_cells('A32:T33')
            ws['A32'] = '1. Acceptance Criteria (%) - Continued'
            _style_range(ws, 'A32:T33', font=f_title, alignment=align_c, border=b_thin)

            ws.row_dimensions[34].height = 32.0
            ws.merge_cells('A34:E34')
            ws['A34'] = 'Element'
            _style_range(ws, 'A34:E34', font=f_hdr, alignment=align_c, border=b_thin)

            for g_i in range(len(tier2)):
                c_start, c_end = ranges_col[g_i]
                rng_txt = f"{c_start}34:{c_end}34"
                ws.merge_cells(rng_txt)
                ws[f"{c_start}34"] = _get_hdr_label(tier2[g_i])
                _style_range(ws, rng_txt, font=f_hdr, alignment=align_c, border=b_thin)

            for idx, el_k in enumerate(target_elems, start=35):
                full_n = elem_full_names.get(el_k, el_k)
                ws.merge_cells(f'A{idx}:E{idx}')
                ws[f'A{idx}'] = full_n
                _style_range(ws, f'A{idx}:E{idx}', font=f_data, alignment=align_c, border=b_thin)
                for g_i in range(len(tier2)):
                    c_start, c_end = ranges_col[g_i]
                    v = _format_limit(tier2[g_i]['limits'], el_k)
                    rng_txt = f"{c_start}{idx}:{c_end}{idx}"
                    ws.merge_cells(rng_txt)
                    ws[f"{c_start}{idx}"] = v
                    _style_range(ws, rng_txt, font=f_data, alignment=align_c, border=b_thin)

            ws.merge_cells('A38:T38')
            ws['A38'] = '* Tolerance ±10%'
            _style_range(ws, 'A38:T38', font=f_note, alignment=align_c, border=b_thin)


    def evaluate_pmi_data_compliance(self, show_msg=True):
        """추출된 PMI 측정 데이터를 현재 재질 기준(또는 element_filters)과 비교하여 합/부(ALL MATCH / MISMATCH) 자동 판정"""
        if not getattr(self, 'extracted_data', None):
            if show_msg:
                messagebox.showinfo("알림", "판정할 PMI 추출 데이터가 없습니다.")
            return

        limits = {}
        if hasattr(self, 'element_filters') and self.element_filters:
            for f in self.element_filters:
                k = f['key'].get().strip()
                if not k: continue
                min_v = self.to_float(f['min'].get())
                max_v = self.to_float(f['max'].get())
                if min_v > 0 or max_v > 0:
                    limits[k] = (min_v, max_v if max_v > 0 else 999.9)

        if not limits:
            spec = self._lookup_pmi_material_spec(self.gapji_material.get().strip())
            if spec:
                limits = spec["limits"]

        if not limits:
            if show_msg:
                messagebox.showwarning("경고", "판정할 원소 기준(Min ~ Max)이 설정되지 않았습니다.\n재질을 선택하거나 성분 필터를 입력해주세요.")
            return

        pass_count = 0
        fail_count = 0
        margin = 0.10 # 10% tolerance (갑지 Acceptance Criteria 규격)

        for item in self.extracted_data:
            mismatches = []
            for elem_k, (f_min, f_max) in limits.items():
                val = self.to_float(self._get_val_ci(item, elem_k))
                tol_min = round(f_min * (1.0 - margin), 2) if f_min > 0 else 0.0
                tol_max = round(f_max * (1.0 + margin), 2) if f_max > 0 else 999.9

                if f_min > 0 and val < tol_min:
                    mismatches.append(f"{elem_k}미달({val:.2f}%)")
                elif f_max > 0 and val > tol_max:
                    mismatches.append(f"{elem_k}초과({val:.2f}%)")

            if not mismatches:
                item['Grade'] = "ALL MATCH"
                item['Result'] = "ALL MATCH"
                item['is_deficient'] = False
                pass_count += 1
            else:
                reason = ", ".join(mismatches)
                item['Grade'] = f"MISMATCH ({reason})"
                item['Result'] = "MISMATCH"
                item['is_deficient'] = True
                fail_count += 1

        self.populate_preview(self.extracted_data, switch_tab=False, mode="PMI")
        msg = f"⚖️ 규격 판정 완료: 전체 {len(self.extracted_data)}건 중 합격(ALL MATCH) {pass_count}건, 부적합(MISMATCH) {fail_count}건"
        self.log(msg)
        if show_msg:
            messagebox.showinfo("판정 결과", msg)

    def show_pmi_element_selector_dialog(self):
        """추출된 데이터 또는 전체 화학 원소 목록에서 리포트에 표기할 원소(최대 5개)를 선택하는 다이얼로그"""
        detected_elements = set()
        if getattr(self, 'extracted_data', None):
            ignore_keys = {"_status", "selected", "No", "Date", "Report No", "Dwg", "Joint", "Loc", "Grade", "Result", "Remarks", "Welder", "order_index"}
            for item in self.extracted_data[:30]:
                for k, v in item.items():
                    if k not in ignore_keys and isinstance(v, (int, float)) and v > 0:
                        detected_elements.add(k)
        
        common_elements = ["Ni", "Cr", "Mo", "Ti", "Nb", "Mn", "Si", "Cu", "V", "W", "Co", "Fe", "Al", "C", "P", "S"]
        ordered_candidates = [e for e in common_elements if e in detected_elements] + [e for e in detected_elements if e not in common_elements] + [e for e in common_elements if e not in detected_elements]
        
        dialog = tk.Toplevel(self.root)
        dialog.title("PMI 출력 원소 선택")
        dialog.geometry("450x340")
        dialog.minsize(420, 300)
        dialog.configure(background="#f9fafb")
        dialog.transient(self.root)
        dialog.grab_set()
        
        main_f = ttk.Frame(dialog, padding=15)
        main_f.pack(fill='both', expand=True)
        
        # 1. 상단 안내
        top_f = ttk.Frame(main_f)
        top_f.pack(fill='x', pady=(0, 10))
        ttk.Label(top_f, text="🔬 성적서 출력 원소 선택 (최대 5개)", font=("Malgun Gothic", 10, "bold")).pack(anchor='w', pady=(0, 3))
        ttk.Label(top_f, text="출력할 화학 원소를 체크하세요. (선택된 순서대로 성적서에 반영됩니다)", font=("Malgun Gothic", 8), foreground="#6b7280").pack(anchor='w')
        
        # 2. 하단 버튼 영역 (창 크기와 무관하게 항상 보이도록 side='bottom'으로 먼저 패킹)
        btn_f = ttk.Frame(main_f)
        btn_f.pack(side='bottom', fill='x', pady=(10, 0))
        
        def apply_selected():
            chosen = [e for e, v in check_vars.items() if v.get()]
            if not chosen:
                messagebox.showwarning("경고", "최소 1개 이상의 원소를 선택해주세요.")
                return
            if len(chosen) > 5:
                messagebox.showwarning("경고", f"성적서 양식상 최대 5개 원소까지만 선택 가능합니다.\n(현재 {len(chosen)}개 선택됨)")
                return
                
            for idx in range(5):
                val = chosen[idx] if idx < len(chosen) else "-"
                self.pmi_chem_elements[idx].set(val)
                
            self.save_settings()
            self.log(f"✅ PMI 출력 원소 설정 완료: {', '.join(chosen)}")
            dialog.destroy()
            
        ttk.Button(btn_f, text=" 적용 (Apply) ", command=apply_selected, width=14).pack(side='right', padx=5)
        ttk.Button(btn_f, text=" 취소 (Cancel) ", command=dialog.destroy, width=14).pack(side='right')

        # 3. 중앙 체크박스 영역 (4열 배치로 콤팩트하게 정렬)
        content_box = ttk.LabelFrame(main_f, text=" 화학 원소 목록 (선택) ", padding=10)
        content_box.pack(fill='both', expand=True, pady=5)
        for col_i in range(4):
            content_box.columnconfigure(col_i, weight=1)

        current_selected = [var.get().strip() for var in self.pmi_chem_elements if var.get().strip() and var.get().strip() != '-']
        check_vars = {}
        for i, elem in enumerate(ordered_candidates[:16]):
            var = tk.BooleanVar(value=(elem in current_selected))
            check_vars[elem] = var
            chk = ttk.Checkbutton(content_box, text=elem, variable=var)
            chk.grid(row=i//4, column=i%4, sticky='w', padx=12, pady=6)

        return dialog

    def check_material_grade(self, row_data):
        """jjchRFIPMI.py에서 이식된 10% 여유치 판정 로직"""
        cr = self.to_float(self._get_val_ci(row_data, 'Cr'))
        ni = self.to_float(self._get_val_ci(row_data, 'Ni'))
        mo = self.to_float(self._get_val_ci(row_data, 'Mo'))
        mn = self.to_float(self._get_val_ci(row_data, 'Mn'))
        
        margin = 0.1 # 10% 여유
        
        # 1. SUS 316 (Cr:16~18 / Ni:10~14 / Mo:2~3)
        if (16.0*(1-margin) <= cr <= 18.0*(1+margin)) and (10.0*(1-margin) <= ni <= 14.0*(1+margin)) and (2.0*(1-margin) <= mo <= 3.0*(1+margin)):
            return "SS316"
        
        # 2. Duplex (Cr:22~23 / Mo:3~3.5 / Ni:4.5~6.5 / Mn:2.0이하)
        if (22.0*(1-margin) <= cr <= 23.0*(1+margin)) and (4.5*(1-margin) <= ni <= 6.5*(1+margin)) and (3.0*(1-margin) <= mo <= 3.5*(1+margin)) and (mn <= 2.2):
            return "DUPLEX"

        # 3. SUS 310 (Cr:24~26 / Ni:19~22)
        if (24.0*(1-margin) <= cr <= 26.0*(1+margin)) and (19.0*(1-margin) <= ni <= 22.0*(1+margin)):
            return "SS310"

        # 4. SUS 304 (Cr:18↑ / Ni:8↑ / Mo:0.5↓ / Mn:2.0↓)
        if (cr >= 16.2) and (ni >= 7.2) and (mo <= 0.55) and (mn <= 2.2):
            return "SS304"

        return None # 판정 불가 시 원본 사용

    def normalize_iso(self, val):
        """정렬 및 그룹화 시와 동일하게 정규화된 값으로 비교하여 도면번호 일관성 유지"""
        s_val = str(val).strip()
        try:
            if re.match(r'^\d+(\.\d+)?$', s_val):
                f_val = float(s_val)
                if f_val == int(f_val): return str(int(f_val))
                return str(f_val)
        except: pass
        return s_val.lower()

    # --- Excel Helper Logic ---

    def find_image_smart(self, keyword, exclude_keyword=None):
        def _search_in_folder(folder_path):
            if not folder_path or not os.path.exists(folder_path): return None
            candidates = glob.glob(os.path.join(folder_path, "*.*"))
            valid_extensions = ['.PNG', '.JPG', '.JPEG', '.BMP', '.GIF']
            # Prefer an exact file-stem match before falling back to substring
            # matching (PMI must not resolve to PMI-1 or PMI갑).
            for path in candidates:
                fname = os.path.basename(path); ext = os.path.splitext(path)[1].upper()
                if ext not in valid_extensions: continue
                if os.path.splitext(fname)[0].upper() == keyword.upper():
                    if exclude_keyword and exclude_keyword.upper() in fname.upper(): continue
                    return path
            for path in candidates:
                fname = os.path.basename(path).upper(); ext = os.path.splitext(path)[1].upper()
                if ext not in valid_extensions: continue 
                if keyword.upper() in fname:
                    if exclude_keyword and exclude_keyword.upper() in fname: continue
                    return path 
            return None

        # 1. UI에서 설정한 폴더에서 먼저 검색
        found = _search_in_folder(self.logo_folder_path.get())
        if found: return found

        # Use the bundled project resources when no logo folder/path was selected.
        # This is required for PMI generation because the data-sheet images are
        # cleared before the configured logos are added again.
        app_home = os.path.abspath(
            os.path.join(os.path.dirname(__file__), '..', '..', '..', '..')
        )
        for resource_dir in (
            os.path.join(app_home, 'resources'),
            os.path.join(os.path.dirname(app_home), 'Assets'),
        ):
            found = _search_in_folder(resource_dir)
            if found:
                return found
        
        # 2. PyInstaller 묶음(실행파일 내부 임시 폴더)에서 검색 (Standalone 지원)
        if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
            found = _search_in_folder(sys._MEIPASS)
            if found: return found
            
        return None

    def place_image_freely(self, ws, img_path, anchor_cell_str, w, h, x_offset, y_offset):
        try:
            if not img_path or not os.path.exists(img_path): return
            
            # [ENHANCED] Robust image processing with terminal & UI logging
            original = PILImage.open(img_path).convert("RGBA")
            w_px = max(5, int(float(w))); h_px = max(5, int(float(h)))
            
            temp_name = f"temp_{int(time.time())}_{os.path.basename(img_path)}.png"
            temp_full_path = os.path.join(tempfile.gettempdir(), temp_name)
            
            resized = original.resize((w_px, h_px), PILImage.Resampling.LANCZOS)
            resized.save(temp_full_path, "PNG")
            
            from openpyxl.drawing.image import Image as XLImage
            img = XLImage(temp_full_path)
            img.width = w_px; img.height = h_px
            from openpyxl.utils import column_index_from_string
            from openpyxl.utils.cell import coordinate_from_string
            from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker, XDRPositiveSize2D
            
            # [FIX] Handle 'None' or empty anchor string
            if not anchor_cell_str or str(anchor_cell_str).strip() == "None" or len(str(anchor_cell_str).strip()) < 2:
                anchor_cell_str = "A1"

            try:
                col_str, row_num = coordinate_from_string(anchor_cell_str)
                col_idx = max(0, column_index_from_string(col_str) - 1)
                row_idx = max(0, row_num - 1) 
            except Exception as e:
                col_idx, row_idx = 0, 0
                anchor_cell_str = "A1"
            
            # [STABLE] Use OneCellAnchor instead of AbsoluteAnchor for better Excel compatibility (prevents corruption)
            # 1 Pixel = 9525 EMU (Standard 96 DPI)
            final_emu_x = max(0, int(float(x_offset) * 9525))
            final_emu_y = max(0, int(float(y_offset) * 9525))
            emu_w = max(9525, int(float(w_px) * 9525))
            emu_h = max(9525, int(float(h_px) * 9525))
            
            # [FIX] Use AnchorMarker (cell + offset) for stable positioning
            marker = AnchorMarker(col=col_idx, colOff=final_emu_x, row=row_idx, rowOff=final_emu_y)
            size = XDRPositiveSize2D(cx=emu_w, cy=emu_h)
            
            img.anchor = OneCellAnchor(_from=marker, ext=size)
            ws.add_image(img)
            
            msg = f"[SUCCESS] 로고 배치 완료: {os.path.basename(img_path)}"
            self.log(f"   {msg}")
        except Exception as e:
            msg = f"[ERROR] 로고 배치 실패: {e}"
            self.log(f"   {msg}")

    @staticmethod
    def _white_background_to_transparent_png(source):
        """이미지의 흰 배경만 제거하고 원본 색상/획은 그대로 유지한다."""
        if isinstance(source, (bytes, bytearray)):
            source = io.BytesIO(source)

        with PILImage.open(source) as original:
            image = original.convert("RGBA")
            converted = []
            for red, green, blue, alpha in image.getdata():
                # 완전한 흰색은 투명하게, 가장자리의 밝은 픽셀은 부드럽게
                # 반투명 처리하여 흰 테두리가 남지 않도록 한다.
                whiteness = min(red, green, blue)
                if whiteness >= 248:
                    new_alpha = 0
                elif whiteness >= 210:
                    new_alpha = int(alpha * (248 - whiteness) / 38)
                else:
                    new_alpha = alpha
                converted.append((red, green, blue, new_alpha))
            image.putdata(converted)

            output = io.BytesIO()
            image.save(output, format="PNG")
            output.seek(0)
            return output

    def _prepare_paut_cover_identity(self, ws):
        """PAUT 갑지 직인/검사자 이름/서명을 투명 이미지로 정리한다."""
        # 원본 양식의 이름 셀 서식을 유지하면서 검사자만 교체한다.
        self.safe_set_value(ws, "D45", "박 광 복")
        self.safe_set_value(ws, "AA44", "한국지역난방공사 중앙지사")

        signature_path = os.path.join(SRC_ROOT or "", "signs", "박광복.png")
        if not os.path.exists(signature_path):
            self.log(f"⚠️ 박광복 서명 파일을 찾을 수 없습니다: {signature_path}")
            return

        reviewer_signature_path = os.path.join(
            SRC_ROOT or "", "signs", "주진철.png"
        )
        if not os.path.exists(reviewer_signature_path):
            self.log(
                f"⚠️ 주진철 서명 파일을 찾을 수 없습니다: "
                f"{reviewer_signature_path}"
            )
            return

        # 원본 양식 기준 그림 위치(0부터 시작):
        # 직인 Q3, 검사자 서명 I44, 판독자 서명 S44.
        replacement_specs = {
            (16, 2): None,
            (8, 43): signature_path,
            (18, 43): reviewer_signature_path,
        }
        replaced_images = []

        for original_image in list(getattr(ws, "_images", [])):
            marker = getattr(getattr(original_image, "anchor", None), "_from", None)
            position = (
                getattr(marker, "col", None),
                getattr(marker, "row", None),
            )
            if position not in replacement_specs:
                replaced_images.append(original_image)
                continue

            try:
                replacement_source = replacement_specs[position]
                if replacement_source:
                    transparent_stream = self._white_background_to_transparent_png(
                        replacement_source
                    )
                else:
                    transparent_stream = self._white_background_to_transparent_png(
                        original_image._data()
                    )

                replacement = XLImage(transparent_stream)
                replacement.anchor = copy.deepcopy(original_image.anchor)
                # workbook 저장 시까지 BytesIO가 닫히지 않게 참조를 유지한다.
                replacement._pmi_source_stream = transparent_stream
                replaced_images.append(replacement)
            except Exception as exc:
                replaced_images.append(original_image)
                self.log(f"⚠️ PAUT 직인/서명 투명화 실패({position}): {exc}")

        ws._images = replaced_images
        self.log(
            "✅ PAUT 갑지 직인 투명화 및 박광복/주진철 투명 서명 교체 완료"
        )

    def _prepare_paut_data_stamp(self, ws):
        """PAUT 두 번째 시트 헤더 직인의 흰 배경만 투명하게 만든다."""
        # 원본 양식 drawing2 기준 직인 위치(0부터 시작): O2.
        stamp_position = (14, 1)
        replaced_images = []
        stamp_replaced = False

        for original_image in list(getattr(ws, "_images", [])):
            marker = getattr(getattr(original_image, "anchor", None), "_from", None)
            position = (
                getattr(marker, "col", None),
                getattr(marker, "row", None),
            )
            if position != stamp_position:
                replaced_images.append(original_image)
                continue

            try:
                transparent_stream = self._white_background_to_transparent_png(
                    original_image._data()
                )
                replacement = XLImage(transparent_stream)
                replacement.anchor = copy.deepcopy(original_image.anchor)
                replacement._pmi_source_stream = transparent_stream
                replaced_images.append(replacement)
                stamp_replaced = True
            except Exception as exc:
                replaced_images.append(original_image)
                self.log(f"⚠️ PAUT 두 번째 시트 직인 투명화 실패: {exc}")

        ws._images = replaced_images
        if stamp_replaced:
            self.log("✅ PAUT 두 번째 시트 헤더 직인 투명화 완료")
        else:
            self.log("⚠️ PAUT 두 번째 시트 헤더 직인을 찾지 못했습니다.")

    def load_daily_work_history(self):
        """사업장별 작업 감독일보에서 확인한 NDT 항목만 읽어 가져옵니다."""
        try:
            try:
                main_tab = self.mode_notebook.tab(self.mode_notebook.select(), "text")
                if "RT" in main_tab:
                    sub_tab_text = self.rt_preview_nb.tab(self.rt_preview_nb.select(), "text")
                    mode = "KOGAS" if "가스공사" in sub_tab_text else "RT"
                elif "PT" in main_tab: mode = "PT"
                elif "MT" in main_tab: mode = "MT"
                elif "PAUT" in main_tab: mode = "PAUT"
                else: mode = "PMI"
            except: mode = "PMI"
            
            from site_apps.central.src.reports.ndt_history_import import select_history_rows
            ndt_list = select_history_rows(self.root, mode)
            if not ndt_list:
                return
            tree, idx_map, data, keys_attr = self._get_mode_info(mode)
            
            # Map NDT data to the tree columns
            source_site = ndt_list[0]['_source_site']
            if data and any(row.get('_source_site') != source_site for row in data):
                if not messagebox.askyesno('사업장 확인',
                    f'현재 보고서에 다른 사업장 또는 출처 미확인 자료가 있습니다.\n{source_site} 자료를 함께 추가할까요?'):
                    return
            existing_ids = {row.get('_source_id') for row in data if row.get('_source_id')}
            added_count = 0
            for r in ndt_list:
                if r.get('_source_id') and r['_source_id'] in existing_ids:
                    continue
                new_row = {k: "" for k in keys_attr}
                new_row['_source_site'] = r['_source_site']
                new_row['_source_path'] = r['_source_path']
                new_row['_source_id'] = r.get('_source_id')
                new_row['_source_contractor'] = r.get('업체', '')
                new_row['_source_tested_length'] = r.get('검사길이', '') or r.get(mode, '')
                new_row['_source_work_type'] = r.get('근무구분', '')
                for quantity_key in ('RT_OR', 'RT_RE'):
                    if quantity_key in r:
                        new_row[quantity_key] = r[quantity_key]
                new_row["Date"] = r.get("Date", "")

                # 중앙지사 작업/감독일보의 '구간'은 RT의 SEC 값이다.
                if mode in ("RT", "KOGAS"):
                    new_row["Sec"] = r.get("구간", "")
                 
                # 라인번호 / ISO / Dwg
                dwg = r.get("라인번호", "")
                if "ISO" in keys_attr: new_row["ISO"] = dwg
                elif "Line No." in keys_attr: new_row["Line No."] = dwg
                elif "Dwg" in keys_attr: new_row["Dwg"] = dwg
                elif "Drawing No." in keys_attr: new_row["Drawing No."] = dwg
                
                # Joint No.
                joint = r.get("Joint No.", "")
                if "Joint" in keys_attr: new_row["Joint"] = joint
                elif "Joint No." in keys_attr: new_row["Joint No."] = joint
                
                # 관경 / Size / T / Th'k
                size = r.get("관경", "")
                thickness = r.get("두께", "")
                if "Size" in keys_attr: new_row["Size"] = size
                if "T" in keys_attr: new_row["T"] = thickness
                if "Th'k(mm)" in keys_attr: new_row["Th'k(mm)"] = thickness
                if not thickness and any(k in keys_attr for k in ('T', "Th'k(mm)")):
                    new_row['Remarks'] = '두께 확인 필요'
                
                # 용접사 / Welder
                welder = self._normalize_daily_photo_welder(
                    r.get("용접사", "")
                )
                if "Welder" in keys_attr: new_row["Welder"] = welder
                
                # 검사길이 / Tested Length
                test_len = r.get("검사길이", "")
                if not test_len and mode in ["PAUT", "MT", "PT"]:
                    test_len = r.get(mode, "")
                if "Tested Length" in keys_attr: new_row["Tested Length"] = test_len
                
                # 구간 / Loc
                # RT의 Film Location은 중앙지사 작업/감독일보
                # '2. 비파괴검사결과서'의 '구간정보' 값을 사용한다.
                if mode in ("RT", "KOGAS"):
                    loc = r.get("구간정보", "") or r.get("구간", "")
                else:
                    loc = r.get("구간", "")
                if "Loc" in keys_attr: new_row["Loc"] = loc
                elif "Location" in keys_attr: new_row["Location"] = loc
                if mode in ("RT", "KOGAS"):
                    from site_apps.central.src.reports.ndt_history_import import rt_film_findings
                    new_row.update(rt_film_findings(r.get("구간정보", "")))
                
                # 결과 / Result / Acc / Rej
                if mode == "PAUT":
                    from site_apps.central.src.reports.ndt_history_import import paut_flaw_findings
                    findings = paut_flaw_findings(r.get("구간정보", ""), thickness)
                    if findings.get('Remarks') and new_row.get('Remarks'):
                        findings['Remarks'] = new_row['Remarks'] + '; ' + findings['Remarks']
                    new_row.update(findings)
                res = r.get("결과", "")
                if "Result" in keys_attr:
                    new_row["Result"] = res
                if res == "합격":
                    if "Acc" in keys_attr: new_row["Acc"] = "√"
                    if "Evaluation" in keys_attr: new_row["Evaluation"] = "Acc"
                elif res == "불합격":
                    if "Rej" in keys_attr: new_row["Rej"] = "X"
                    if "Evaluation" in keys_attr: new_row["Evaluation"] = "Rej"
                
                new_row["selected"] = True
                new_row["_status"] = "new"
                if "No" in keys_attr: new_row["No"] = str(len(data) + 1)
                
                data.append(new_row)
                if r.get('_source_id'):
                    existing_ids.add(r['_source_id'])
                
                row_vals = [new_row.get(k, "") for k in keys_attr]
                row_tags = []
                if new_row.get("Result") == "불합격" or new_row.get("Evaluation") == "Rej":
                    row_tags.append("fail")
                elif new_row.get("Result") == "합격" or new_row.get("Evaluation") == "Acc":
                    row_tags.append("pass")
                
                tree.insert("", "end", values=tuple(row_vals), tags=tuple(row_tags))
                idx_map.append(len(data) - 1)
                added_count += 1
                
            self.update_date_listbox(mode=mode)
            if mode == 'PAUT':
                companies = {row.get('업체', '').strip() for row in ndt_list if row.get('업체', '').strip()}
                if len(companies) == 1 and hasattr(self, 'paut_contractor'):
                    self.paut_contractor.set(next(iter(companies)))
                    self.config['PAUT_CONTRACTOR'] = self.paut_contractor.get()
                    self.save_settings()
                self.populate_preview(data, switch_tab=False, mode=mode)
            elif mode in ('MT', 'RT', 'KOGAS'):
                self.populate_preview(data, switch_tab=False, mode=mode)
            messagebox.showinfo("완료", f"{ndt_list[0]['_source_site']} 작업 감독일보에서 {added_count}개의 {mode} NDT 결과를 성공적으로 불러왔습니다.")
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("오류", f"일보 불러오기 중 오류가 발생했습니다:\n{e}")

    def extract_only(self, show_msg=True):
        """데이터만 추출하여 리스트와 미리보기에 반영 (PMI/RT 대응)"""
        self.save_settings() # [FIX] UI 설정(컬럼명 등)을 즉시 반영하여 리셋 방지
        # 현재 활성 탭에 따른 모드 결정
        try:
            main_tab = self.mode_notebook.tab(self.mode_notebook.select(), "text")
            if "RT" in main_tab:
                # [FIX] 하부 탭 이름을 직접 확인하여 모드 결정
                sub_tab_text = self.rt_preview_nb.tab(self.rt_preview_nb.select(), "text")
                mode = "KOGAS" if "가스공사" in sub_tab_text else "RT"
            elif "PT" in main_tab: mode = "PT"
            elif "MT" in main_tab: mode = "MT"
            elif "PAUT" in main_tab: mode = "PAUT"
            else: mode = "PMI"
        except Exception as e:
            self.log(f"⚠️ 모드 판별 오류: {e}")
            mode = "PMI"

        if mode == "PT":
            target_file_str = self.pt_target_file_path.get()
        elif mode == "MT":
            target_file_str = self.mt_target_file_path.get()
        elif mode == "PAUT":
            target_file_str = self.paut_target_file_path.get()
        elif mode == "RT":
            target_file_str = self.rt_target_file_path.get()
        elif mode == "KOGAS":
            target_file_str = self.kogas_target_file_path.get()
        else:
            target_file_str = self.target_file_path.get()
            
        if not target_file_str:
            messagebox.showwarning("파일 미선택", f"{mode} 데이터 파일을 선택해주세요.")
            return False
            
        target_files = target_file_str.split('|') if '|' in target_file_str else [target_file_str]
        
        if not target_files: return False
        
        success_count = 0
        for i, tf in enumerate(target_files):
            if not tf.strip(): continue
            is_last = (i == len(target_files) - 1)
            if self._extract_single_file(mode, tf.strip(), show_msg=(show_msg and is_last)):
                success_count += 1
                
        return success_count > 0

    def _extract_single_file(self, mode, target_file, show_msg=True):
        self.log(f"🔍 {mode} 데이터 추출 시작: {os.path.basename(target_file)}")
        
        # [NEW] Extract date from filename
        fname = os.path.basename(target_file)
        date_match = re.search(r'(\d{4}[-._]\d{2}[-._]\d{2}|\d{2}[-._]\d{2}[-._]\d{2}|\d{8}|\d{6})', fname)
        extracted_date = date_match.group(0) if date_match else ""
        
        # Standardize formatting
        if extracted_date:
            clean_date = re.sub(r'[-._]', '', extracted_date)
            if len(clean_date) == 8:
                extracted_date = f"{clean_date[:4]}-{clean_date[4:6]}-{clean_date[6:]}"
            elif len(clean_date) == 6:
                extracted_date = f"20{clean_date[:2]}-{clean_date[2:4]}-{clean_date[4:]}"
        
        if not extracted_date:
            extracted_date = datetime.datetime.now().strftime("%Y-%m-%d")
            self.log(f"⚠️ 파일명에서 날짜를 찾지 못해 오늘 날짜({extracted_date})로 설정합니다.")
        else:
            self.log(f"📅 파일 날짜 인식 성공: {extracted_date}")

        def _is_date_like_dwg(value):
            """Reject timestamps that can be mistaken for a drawing number."""
            if value is None or pd.isna(value):
                return False
            if isinstance(value, (datetime.datetime, datetime.date, pd.Timestamp)):
                return True
            text = str(value).strip()
            return bool(re.fullmatch(
                r'\d{4}[-./]\d{1,2}[-./]\d{1,2}(?:[ T]\d{1,2}:\d{2}(?::\d{2})?)?',
                text,
            ))

        self.progress['value'] = 0
        all_extracted_data = []

        # [AUTO COLUMN ADDITION] 
        if mode == "PMI" and hasattr(self, 'element_filters') and self.element_filters:
            new_keys = []
            for f_item in self.element_filters:
                f_k = f_item['key'].get().strip()
                if f_k and f_k.upper() not in ["CR", "NI", "MO", "MN"] and f_k.capitalize() not in self.column_keys and f_k.upper() not in self.column_keys and f_k not in self.column_keys:
                    new_keys.append(f_k)
            
            if new_keys:
                self.column_keys.extend(new_keys)
                current_cols = list(self.preview_tree["columns"])
                current_display = list(self.preview_tree.cget("displaycolumns"))
                if current_display[0] == "#all": current_display = list(current_cols)

                # Save existing configurations to survive the tree["columns"] reset
                configs = {}
                for c in current_cols:
                    configs[c] = {
                        "text": self.preview_tree.heading(c).get("text", c),
                        "width": self.preview_tree.column(c).get("width", 80)
                    }

                for k in new_keys:
                    if k not in current_cols: current_cols.append(k)
                    if k not in current_display: current_display.append(k)
                
                self.preview_tree["columns"] = tuple(current_cols)
                self.preview_tree["displaycolumns"] = tuple(current_display)
                
                # Re-apply all configs + new keys setup
                for c in current_cols:
                    header_txt = configs.get(c, {}).get("text", c)
                    w = configs.get(c, {}).get("width", 80)
                    self.preview_tree.heading(c, text=header_txt, anchor='center', command=lambda _c=c: self.sort_by_column(_c, mode="PMI"))
                    self.preview_tree.column(c, width=w, anchor='center', stretch=False)
        
        def _find_col(df, keywords, exclude=None):
            cleaned_keywords = [re.sub(r'\s+', '', k).upper() for k in keywords]
            # 1. Exact match by keyword priority
            for k in cleaned_keywords:
                for col in df.columns:
                    c_clean = re.sub(r'\s+', '', str(col)).upper()
                    if exclude and any(re.sub(r'\s+', '', ex).upper() in c_clean for ex in exclude): continue
                    if k == c_clean: return col
            # 2. Substring match by keyword priority
            for k in cleaned_keywords:
                for col in df.columns:
                    c_clean = re.sub(r'\s+', '', str(col)).upper()
                    if exclude and any(re.sub(r'\s+', '', ex).upper() in c_clean for ex in exclude): continue
                    if k in c_clean:
                        if "NI" in keywords and ("UNIT" in c_clean or "LINE" in c_clean): continue
                        return col
            return None

        def _get_kw(k): 
            if hasattr(self, 'rt_extract_mappings') and k in self.rt_extract_mappings:
                return [x.strip().upper() for x in self.rt_extract_mappings[k].get().split(',')]
            return []

        try:
            target_input = self.sequence_filter.get().strip()
            target_no_list = [x.strip() for x in target_input.replace(',', ' ').split() if target_input and x.strip()] if target_input else []
            
            with pd.ExcelFile(target_file) as xls:
                for s_idx, sheet_name in enumerate(xls.sheet_names):
                    self.log(f"📄 시트 스캔: {sheet_name}")
                    try: temp_df = pd.read_excel(xls, sheet_name=sheet_name, header=None, nrows=50)
                    except: continue
                    
                    # [지능형 시트 내 날짜 추출]
                    # 의뢰서의 L5, L6 셀(머지된 영역)을 직접 파싱하여 
                    # 파일명 날짜보다 우선하여 해당 시트의 기본 날짜(sheet_level_date)로 활용합니다.
                    sheet_level_date = extracted_date
                    try:
                        for r_pos in [4, 5]: # 0-indexed 4=5행, 5=6행
                            for c_pos in [11, 10, 12]: # 11=L열, 10=K열, 12=M열
                                if r_pos < len(temp_df) and c_pos < len(temp_df.columns):
                                    val = temp_df.iloc[r_pos, c_pos]
                                    if pd.notna(val):
                                        if isinstance(val, (datetime.datetime, datetime.date, pd.Timestamp)) or hasattr(val, 'strftime'):
                                            sheet_level_date = val.strftime('%Y-%m-%d')
                                            break
                                        cleaned = str(val).strip()
                                        # YYYY-MM-DD 혹은 YYYY.MM.DD 등의 형태에서 날짜만 클리닝
                                        cleaned_num = re.sub(r'[^0-9]', '', cleaned)
                                        if len(cleaned_num) == 8:
                                            sheet_level_date = f"{cleaned_num[:4]}-{cleaned_num[4:6]}-{cleaned_num[6:]}"
                                            break
                                        elif len(cleaned_num) == 6:
                                            prefix = "20" if int(cleaned_num[:2]) < 50 else "19"
                                            sheet_level_date = f"{prefix}{cleaned_num[:2]}-{cleaned_num[2:4]}-{cleaned_num[4:]}"
                                            break
                            if sheet_level_date != extracted_date: break
                    except: pass
                    
                    if sheet_level_date != extracted_date:
                        self.log(f"   📅 의뢰서 헤더(L5/L6 셀)에서 날짜 파싱 성공: {sheet_level_date}")
                        
                    # [지능형 시트 내 도면번호 추출]
                    # 의뢰서의 C5, C6 셀(머지된 영역)을 직접 파싱하여
                    # 해당 시트의 기본 도면번호(sheet_level_dwg)로 활용합니다.
                    # PT/MT 모드는 의뢰서 고정 헤더 구조가 아니므로 건너뜁니다.
                    sheet_level_dwg = ""
                    if mode not in ["PT", "MT"]:
                        try:
                            # C5/C6 is a valid fixed drawing-number area only
                            # when a nearby cell explicitly labels it as such.
                            nearby_text = " ".join(
                                str(temp_df.iloc[r_pos, c_pos]).strip().upper()
                                for r_pos in range(3, min(7, len(temp_df)))
                                for c_pos in range(0, min(5, len(temp_df.columns)))
                                if pd.notna(temp_df.iloc[r_pos, c_pos])
                            )
                            has_dwg_label = any(
                                label in nearby_text
                                for label in ["DWG", "DRAWING", "도면번호", "도면 NO", "도면NO"]
                            )
                            if has_dwg_label:
                                for r_pos in [4, 5]: # 0-indexed 4=5행, 5=6행
                                    for c_pos in [2, 1, 3]: # 2=C열, 1=B열, 3=D열
                                        if r_pos < len(temp_df) and c_pos < len(temp_df.columns):
                                            val = temp_df.iloc[r_pos, c_pos]
                                            if pd.notna(val):
                                                if _is_date_like_dwg(val):
                                                    continue
                                                val_str = str(val).strip()
                                                if val_str and val_str.lower() != 'nan' and len(val_str) > 2:
                                                    if val_str.upper() in ["도면번호", "도면", "DWG", "DWG NO", "DWG.NO", "도면 NO", "도면 NO."]:
                                                        continue
                                                    val_upper = val_str.upper()
                                                    if any(k in val_upper for k in ["CO.", "LTD", "INSPECTION", "TESTING", "CORP", "INC", "SEOUL", "주식회사"]):
                                                        continue
                                                    if val_str.count(" ") > 2:
                                                        continue
                                                    sheet_level_dwg = val_str
                                                    break
                                    if sheet_level_dwg: break
                        except: pass

                    if sheet_level_dwg:
                        self.log(f"   📐 의뢰서 헤더(C5/C6 셀)에서 도면번호 파싱 성공: {sheet_level_dwg}")
                    
                    header_idx = None
                    is_fixed_range = False
                    # [NEW] KOGAS 고정 영역 처리 (A9:AB13 헤더, A14:AB25 데이터)
                    if mode == "KOGAS":
                        header_idx = 8 # Excel Row 9
                        is_fixed_range = True
                    else:
                        for i, row in temp_df.iterrows():
                            row_str = str(row.values).upper()
                            if mode == "RT":
                                if any(k in row_str for k in ["FILM", "DEFECT", "IQI", "순번", "용접부"]):
                                    header_idx = i; break
                            elif mode == "PT":
                                if (("LINE" in row_str or "ISO" in row_str or "DWG" in row_str) and ("JOINT" in row_str or "WELD" in row_str)):
                                    header_idx = i; break
                            else:
                                if ("CR" in row_str and "NI" in row_str) or ("CHROMIUM" in row_str):
                                    header_idx = i; break
                    
                    if header_idx is None: continue
                    
                    df = pd.read_excel(xls, sheet_name=sheet_name, header=header_idx)
                    if df.empty: continue

                    if is_fixed_range:
                        # [KOGAS] 9~13행 헤더 병합 (df.columns + df.iloc[0:3])
                        new_cols = []
                        for col_idx in range(len(df.columns)):
                            parts = [str(df.columns[col_idx])]
                            for r_idx in range(0, 4): # 10, 11, 12, 13행
                                val = df.iloc[r_idx, col_idx]
                                if pd.notna(val) and "UNNAMED" not in str(val).upper():
                                    parts.append(str(val))
                            new_cols.append(" ".join(parts).upper())
                        df.columns = new_cols
                        # [KOGAS] 14~25행 데이터 추출 (인덱스 4~15)
                        df = df.iloc[4:16].reset_index(drop=True)
                    else:
                        # [STANDARD] 다중 헤더 병합 (최대 2행까지 지원)
                        row0_str = " ".join([str(x).upper() for x in df.iloc[0] if pd.notna(x)]) if len(df) > 0 else ""
                        row1_str = " ".join([str(x).upper() for x in df.iloc[1] if pd.notna(x)]) if len(df) > 1 else ""
                        
                        merge_count = 0
                        if any(k in row1_str for k in ["ORIGIN", "FACTOR", "WELDER1"]):
                            merge_count = 2
                        elif any(k in row0_str for k in ["ORIGIN", "FACTOR", "WELDER1"]):
                            merge_count = 1
                            
                        if merge_count > 0:
                            self.log(f"   ℹ️ [자동] 다중 헤더 병합 ({merge_count}행) (시트: {sheet_name})")
                            new_cols = []
                            for c_idx, col in enumerate(df.columns):
                                c_txt = str(col) if pd.notna(col) and "Unnamed" not in str(col) else ""
                                v_parts = []
                                for r in range(merge_count):
                                    v = df.iloc[r, c_idx]
                                    if pd.notna(v) and "Unnamed" not in str(v): v_parts.append(str(v).strip())
                                v_txt = " ".join(v_parts)
                                new_cols.append(f"{c_txt} {v_txt}".strip())
                            df.columns = new_cols
                            df = df.iloc[merge_count:].reset_index(drop=True)

                    # --- Column Identification for THIS sheet ---
                    col_date = None
                    if mode in ["RT", "KOGAS"]:
                        def _get_col(config_key, search_keywords):
                            c_idx_str = str(self.config.get(config_key, "")).strip()
                            if c_idx_str:
                                try:
                                    idx = self.col_to_num(c_idx_str)
                                    if 1 <= idx <= len(df.columns):
                                        candidate = df.columns[idx - 1]
                                        
                                        # [지능형 자동 매핑 우선 정책] 
                                        # 오직 사용자가 설정을 변경하지 않고 '기본값' 상태일 때만, 헤더 매칭이 틀렸을 경우 자동 매핑을 시도합니다.
                                        # 사용자가 명시적으로 컬럼을 직접 변경한 경우는 설정을 100% 신뢰하여 그대로 사용합니다.
                                        DEFAULT_COLUMNS = {
                                            "KOGAS_R_COL_NO": "1", "KOGAS_R_COL_DATE": "2", "KOGAS_R_COL_DWG": "3",
                                            "KOGAS_R_COL_JOINT": "4", "KOGAS_R_COL_LOC": "5", "KOGAS_R_COL_THK": "6",
                                            "KOGAS_R_COL_MAT": "7", "KOGAS_R_COL_SIZE": "8", "KOGAS_R_COL_RES": "9",
                                            "KOGAS_R_COL_WELDER": "10", "KOGAS_R_COL_REM": "11", "KOGAS_R_COL_D1": "12",
                                            "RT_COL_NO": "1", "RT_COL_DATE": "2", "RT_COL_DWG": "3", "RT_COL_JOINT": "4",
                                            "RT_COL_LOC": "5", "RT_COL_THK": "6", "RT_COL_MAT": "7", "RT_COL_SIZE": "0",
                                            "RT_COL_ACC": "8", "RT_COL_REJ": "9", "RT_COL_DEG": "10", "RT_COL_RES": "0",
                                        }
                                        is_default = False
                                        default_val = DEFAULT_COLUMNS.get(config_key)
                                        if default_val is not None:
                                            try:
                                                is_default = (self.col_to_num(c_idx_str) == self.col_to_num(default_val))
                                            except:
                                                is_default = (c_idx_str == default_val)
                                                
                                        if is_default:
                                            cand_clean = re.sub(r'\s+', '', str(candidate)).upper()
                                            cleaned_keywords = [re.sub(r'\s+', '', k).upper() for k in search_keywords]
                                            if search_keywords and not any(k in cand_clean for k in cleaned_keywords):
                                                found = _find_col(df, search_keywords)
                                                if found: return found
                                                
                                        return candidate
                                except: pass
                            return _find_col(df, search_keywords)

                        col_prefix = "KOGAS_R" if mode == "KOGAS" else "RT"
                        col_no = _get_col(f"{col_prefix}_COL_NO", _get_kw('No') or ["NO.", "NO", "SEQ", "ITEM", "순번"])
                        col_dwg = _get_col(f"{col_prefix}_COL_DWG", _get_kw('Dwg') or ["ISO", "DWG", "DRAWING", "도면"])
                        col_joint = _get_col(f"{col_prefix}_COL_JOINT", _get_kw('Joint') or ["JOINT", "WELD NO", "J/N", "FILM IDENT", "용접부"])
                        col_loc = _get_col(f"{col_prefix}_COL_LOC", _get_kw('Loc') or ["LOCATION", "POSITION", "FILM LOC"])
                        col_size = _get_col(f"{col_prefix}_COL_SIZE", _get_kw('Size') or ["SIZE", "DIA", "INCH", "구경", "관경"])
                        col_welder = _get_col(f"{col_prefix}_COL_WELDER", _get_kw('Welder') or ["WELDER", "W/N"])
                        col_remarks = _get_col(f"{col_prefix}_COL_REM", _get_kw('Remarks') or ["REMARKS", "REMARK", "비고"])
                        col_date = _get_col(f"{col_prefix}_COL_DATE", _get_kw('Date') or ["DATE", "검사일"])
                        col_t = _get_col(f"{col_prefix}_COL_THK", _get_kw('T') or ["T", "THICK", "THK", "두께"])
                        col_mat = _get_col(f"{col_prefix}_COL_MAT", _get_kw('Mat') or ["MAT", "MATERIAL"])
                        col_weld = _get_col(f"{col_prefix}_COL_WELD", _get_kw('Weld') or ["WELD", "TYPE"])
                        col_iqi = _get_col(f"{col_prefix}_COL_IQI", _get_kw('IQI') or ["IQI"])
                        col_sens = _get_col(f"{col_prefix}_COL_SENS", _get_kw('Sens') or ["SENS", "SENSITIVITY"])
                        col_den = _get_col(f"{col_prefix}_COL_DEN", _get_kw('Den') or ["DEN", "DENSITY"])
                        col_acc = _get_col(f"{col_prefix}_COL_ACC", _get_kw('Acc') or ["ACC", "합격"])
                        col_rej = _get_col(f"{col_prefix}_COL_REJ", _get_kw('Rej') or ["REJ", "불합격"])
                        col_deg = _get_col(f"{col_prefix}_COL_DEG", _get_kw('Deg') or ["DEG", "물성", "수정브랜드", "GRADE"])
                        col_result = _get_col(f"{col_prefix}_COL_RES", _get_kw('Result') or ["RESULT", "판정"])
                        
                        defect_cols = {}
                        for i in range(1, 16):
                            defect_cols[f"D{i}"] = _get_col(f"{col_prefix}_COL_D{i}", [f"D{i}", f"DEFECT{i}", chr(9311 + i), f"{i}"])
                        defect_cols = {k: v for k, v in defect_cols.items() if v}

                        # [NEW] KOGAS Fallback (B column for Joint, E column for Size) ONLY if not found
                        if mode == "KOGAS":
                            if col_joint is None and len(df.columns) > 1: col_joint = df.columns[1]
                            if col_size is None and len(df.columns) > 4: col_size = df.columns[4]
                    elif mode == "PT":
                        # PT requires an exact standalone sequence header.
                        # Do not treat Spool No / Request No / Film No as No.
                        # If absent, the row loop below assigns 1, 2, 3, ... .
                        col_no = find_exact_column(df, ["NO.", "NO", "SEQ", "ITEM", "순번"])
                        # Use the request date from each PT source row instead
                        # of the file name/header-level fallback date.
                        col_date = _find_col(df, ["TEST DATE", "REQUEST DATE", "DATE", "검사일자", "일자"])
                        col_report_no = _find_col(df, ["REPORT NO", "REPORT"])
                        col_dwg = _find_col(df, ["ISO", "LINE", "DWG", "DRAWING"], exclude=["JOINT", "WELD"]) 
                        col_joint = _find_col(df, ["JOINT NO", "JOINT NUMBER"], exclude=["ISO", "LINE", "ITEM"])
                        if not col_joint:
                            col_joint = _find_col(df, ["JOINT", "WELD"], exclude=["ISO", "LINE", "ITEM", "WPS", "REPORT", "DATE", "TYPE"])
                        col_mat = _find_col(df, ["MAT", "MATERIAL", "재질", "CLASS", "M'TL"])
                        col_size = _find_col(df, ["SIZE", "NPS", "DIA", "INCH", "관경", "ORIGIN", "DI ORIGIN"])
                        col_thk = _find_col(df, ["THK", "THICK", "SCH", "두께"])
                        col_welder = _find_col(df, ["WELDER", "ID", "용접사", "WELDER1"])
                        col_wtype = _find_col(df, ["WELD TYPE", "WELDTYPE", "W.TYPE", "TYPE"], exclude=["JOINT"])
                        col_result = _find_col(df, ["RESULT", "결과", "판정"])
                    # Placeholder for Loc to maintain structure
                        col_loc = None 
                    else:
                        col_cr = _find_col(df, ["CR", "CHROMIUM"]); col_ni = _find_col(df, ["NI", "NICKEL"])
                        col_mo = _find_col(df, ["MO", "MOLYBDENUM"]); col_mn = _find_col(df, ["MN", "MANGANESE"])
                        col_no = _find_col(df, ["NO.", "NO", "SEQ", "NUM", "POS", "ITEM"], exclude=["REPORT", "DWG", "DRAWING", "LINE", "WELD", "JOINT", "ORIGIN"])
                        col_date = _find_col(
                            df,
                            ["DATE/TIME", "DATETIME", "DATE", "TIME", "검사일", "검사일자", "측정시간", "일자"],
                        )
                        col_joint = _find_col(df, ["JOINT", "J/N", "JOINT NO", "PUNCH", "WELD NO"])
                        col_loc = _find_col(df, ["LOCATION", "TEST POSITION", "POINT", "AREA", "POSITION"])
                        col_dwg = _find_col(df, ["ISO", "DWG", "DRAWING", "LINE"])
                        col_report_no = _find_col(df, ["REPORT NO", "REPORT"])
                        col_grade_orig = _find_col(df, ["GRADE", "MATERIAL", "SPEC", "TYPE"])
                        
                        # [DYNAMIC] Scan for dynamically added filter elements (e.g., C, P, S)
                        custom_cols = {}
                        if hasattr(self, 'element_filters') and self.element_filters:
                            for f_item in self.element_filters:
                                f_key = f_item['key'].get().strip()
                                if f_key and f_key.upper() not in ["CR", "NI", "MO", "MN"]:
                                    c_found = _find_col(df, [f_key.upper()])
                                    if c_found is not None:
                                        custom_cols[f_key] = c_found

                    # [NEW] Forward-fill (Fill Down) for ISO and Joint
                    last_dwg = ""
                    last_joint = ""
                    last_date = ""

                    def clean_v(v):
                        if pd.isna(v): return ""
                        # 만약 날짜/시간(Timestamp) 형태라면 YYYY-MM-DD 형식으로 클리닝
                        if isinstance(v, (datetime.datetime, datetime.date, pd.Timestamp)) or hasattr(v, 'strftime'):
                            try:
                                return v.strftime('%Y-%m-%d')
                            except: pass
                        if isinstance(v, str):
                            v_strip = v.strip()
                            # [NEW] Preserve leading zeros for purely numeric strings (e.g. "075")
                            if v_strip.startswith('0') and v_strip.isdigit() and len(v_strip) > 1:
                                return v_strip
                                
                        try:
                            if isinstance(v, (int, float)) or (isinstance(v, str) and v.replace('.','').isdigit()):
                                fv = float(v)
                                if fv.is_integer(): return str(int(fv))
                                return str(fv)
                        except: pass
                        # 만약 문자열이고 날짜 형식인 경우 시간 정보(00:00:00) 제거
                        ret_str = str(v).strip()
                        if len(ret_str) >= 10:
                            if ret_str[:4].isdigit() and ret_str[5:7].isdigit() and ret_str[8:10].isdigit():
                                if ret_str[4] == '-' and ret_str[7] == '-':
                                    ret_str = ret_str[:10]
                                elif ret_str[4] == '.' and ret_str[7] == '.':
                                    ret_str = ret_str[:10]
                                elif ret_str[4] == '/' and ret_str[7] == '/':
                                    ret_str = ret_str[:10]
                        return ret_str

                    def process_welder_string(w_str):
                        if not w_str or str(w_str).lower() == 'nan': return ''
                        import re
                        parts = re.split(r'[/,]', str(w_str))
                        processed = []
                        for p in parts:
                            p = p.strip()
                            if len(p) > 3:
                                p = p[-3:]
                            if p and p not in processed:
                                processed.append(p)
                        return "/".join(processed)

                if mode == "KOGAS":
                    # ===== 가스공사 모드: 2개 행을 1개 Joint로 페어링 =====
                    num_rows = len(df)
                    r_idx = 0
                    while r_idx < num_rows:
                        row_top = df.iloc[r_idx]
                        row_bot = df.iloc[r_idx + 1] if (r_idx + 1 < num_rows) else None

                        # 1. 순번 (No)
                        v_raw_no = clean_v(row_top[col_no]) if col_no is not None else ""
                        if not v_raw_no and row_bot is not None and col_no is not None:
                            v_raw_no = clean_v(row_bot[col_no])
                        
                        # 만약 NO 컬럼이 지정되어 있는데 값이 완전히 비어있거나 'nan'인 경우, 데이터가 없는 행이므로 건너뜁니다.
                        if col_no is not None and (not v_raw_no or v_raw_no == 'nan'):
                            r_idx += 2
                            continue
                            
                        # NO 컬럼이 지정되어 있지 않은 경우에만 인덱스 기반으로 번호를 생성합니다.
                        if not v_raw_no or v_raw_no == 'nan':
                            v_raw_no = str((r_idx // 2) + 1)

                        if target_no_list and v_raw_no not in target_no_list:
                            r_idx += 2
                            continue

                        # 키워드 필터링
                        extract_key = self.rt_extract_keyword.get().strip().lower()
                        if extract_key:
                            top_vals = [clean_v(v).lower() for v in row_top.values if pd.notna(v)]
                            bot_vals = [clean_v(v).lower() for v in row_bot.values if pd.notna(v)] if row_bot is not None else []
                            row_str = ' '.join(top_vals + bot_vals)
                            if extract_key not in row_str:
                                r_idx += 2
                                continue

                        # 2. 도면번호 (Dwg) 및 Dwg_Sub
                        curr_dwg = sheet_level_dwg if sheet_level_dwg else (clean_v(row_top[col_dwg]) if col_dwg is not None else '')
                        if (not curr_dwg or curr_dwg == 'nan'):
                            curr_dwg = last_dwg if last_dwg else ''
                        if curr_dwg and curr_dwg != 'nan': last_dwg = curr_dwg

                        curr_dwg_sub = ''
                        if sheet_level_dwg:
                            curr_dwg_sub = sheet_level_dwg
                        elif row_bot is not None and col_dwg is not None:
                            curr_dwg_sub = str(row_bot[col_dwg]).strip()
                        if not curr_dwg_sub or curr_dwg_sub == 'nan':
                            curr_dwg_sub = curr_dwg

                        # 3. 공동 (Joint)
                        curr_joint = clean_v(row_top[col_joint]) if col_joint is not None else ''
                        if (not curr_joint or curr_joint == 'nan') and last_joint: curr_joint = last_joint
                        if curr_joint and curr_joint != 'nan': last_joint = curr_joint
                        elif not curr_joint or curr_joint == 'nan': curr_joint = v_raw_no

                        # 4. 일자 (Date)
                        curr_date = clean_v(row_top[col_date]) if col_date is not None else ''
                        if (not curr_date or curr_date.lower() == 'nan') and last_date: curr_date = last_date
                        if curr_date and curr_date.lower() != 'nan': last_date = curr_date
                        elif not curr_date or curr_date.lower() == 'nan': curr_date = sheet_level_date if sheet_level_date else ''

                        v_report = clean_v(row_top[col_report_no]) if col_report_no is not None else ''
                        if (not v_report or v_report == 'nan') and getattr(self, 'last_report', ''): v_report = self.last_report
                        if v_report and v_report != 'nan': self.last_report = v_report
                        elif not v_report or v_report == 'nan': v_report = ''

                        # 5. 규격 (Size) 및 촬영 매수 (Shots)
                        raw_size = str(row_top[col_size]).strip() if col_size is not None else ''
                        s_match = re.search(r'(\d+\.?\d*)', raw_size)
                        size_val = s_match.group(1) if s_match else raw_size
                        num_shots = self.calculate_rt_shots(size_val)

                        # 6. 용접사 (Welder) 및 Welder_Sub
                        raw_welder = clean_v(row_top[col_welder]) if col_welder is not None else ''
                        v_welder = process_welder_string(raw_welder)

                        raw_welder_sub = ''
                        if row_bot is not None and col_welder is not None:
                            raw_welder_sub = clean_v(row_bot[col_welder])
                        v_welder_sub = process_welder_string(raw_welder_sub)
                        
                        # [NEW] Combine both welders into one string if they are different and not already included
                        v_welder_parts = v_welder.split('/') if v_welder else []
                        if v_welder_sub and v_welder_sub != 'nan':
                            for sub_part in v_welder_sub.split('/'):
                                if sub_part and sub_part not in v_welder_parts:
                                    v_welder_parts.append(sub_part)
                        v_welder = "/".join(v_welder_parts)
                        
                        v_welder_sub = v_welder

                        # 7. 재질 (Mat) 및 Mat_Sub
                        raw_mat = clean_v(row_top[col_mat]) if col_mat is not None else ''
                        raw_mat_sub = ''
                        if row_bot is not None and col_mat is not None:
                            raw_mat_sub = clean_v(row_bot[col_mat])
                        if not raw_mat_sub or raw_mat_sub == 'nan':
                            raw_mat_sub = raw_mat

                        # 8. 두께 (T)
                        raw_t = clean_v(row_top[col_t]) if col_t is not None else ''
                        if raw_t:
                            if '*' in raw_t or 'x' in raw_t.lower():
                                t_match = re.search(r'[*xX]\s*(\d+\.?\d*)\s*[tT]?', raw_t)
                                if t_match:
                                    raw_t = t_match.group(1)
                            else:
                                t_suffix_match = re.search(r'^(\d+\.?\d*)\s*[tT]?$', raw_t)
                                if t_suffix_match:
                                    raw_t = t_suffix_match.group(1)

                        # 9. 위치 (Loc) - 촬영 위치 추출 추가!
                        v_loc = clean_v(row_top[col_loc]) if col_loc is not None else ''
                        if not v_loc or v_loc == 'nan':
                            v_loc = '-'

                        item_data = {
                            'No': v_raw_no, 'Date': curr_date,
                            'Dwg': curr_dwg, 'Dwg_Sub': curr_dwg_sub,
                            'Joint': curr_joint, 'Loc': v_loc,
                            'Acc': clean_v(row_top[col_acc]) if col_acc is not None else '',
                            'Rej': clean_v(row_top[col_rej]) if col_rej is not None else '',
                            'Deg': clean_v(row_top[col_deg]) if col_deg is not None else '',
                            'Welder': v_welder, 'Welder_Sub': v_welder_sub,
                            'Remarks': clean_v(row_top[col_remarks]) if col_remarks is not None else '',
                            'T': raw_t,
                            'Mat': raw_mat, 'Mat_Sub': raw_mat_sub,
                            'Weld': clean_v(row_top[col_weld]) if col_weld is not None else '',
                            'IQI': clean_v(row_top[col_iqi]) if col_iqi is not None else '',
                            'Sens': clean_v(row_top[col_sens]) if col_sens is not None else '',
                            'Den': clean_v(row_top[col_den]) if col_den is not None else '',
                            'Result': clean_v(row_top[col_result]) if col_result is not None else 'ACC',
                            'Size': size_val, 'num_shots': num_shots,
                            'selected': True, 'order_index': len(self.rt_extracted_data) + len(all_extracted_data),
                            '_src': {
                                'sheet': sheet_name,
                                'row': r_idx + 4 + header_idx + 2,
                                'col_result': list(df.columns).index(col_result) + 1 if col_result is not None else None,
                                'col_remarks': list(df.columns).index(col_remarks) + 1 if col_remarks is not None else None,
                                'col_acc': list(df.columns).index(col_acc) + 1 if col_acc is not None else None,
                                'col_rej': list(df.columns).index(col_rej) + 1 if col_rej is not None else None,
                                'col_defects': {k: list(df.columns).index(v) + 1 for k, v in defect_cols.items()}
                            }
                        }
                        # Defects
                        for i in range(1, 16):
                            key = f'D{i}'; src_col = defect_cols.get(key)
                            src_val = str(row_top[src_col]).strip() if src_col is not None else ''
                            item_data[key] = src_val if (src_val and src_val.lower() != 'nan') else ''

                        if num_shots > 1:
                            for shot_idx in range(1, num_shots + 1):
                                shot_item = item_data.copy()
                                shot_item['Loc'] = f'{shot_idx}-{(shot_idx % num_shots) + 1}'
                                for i in range(1, 16):
                                    key = f'D{i}'; c = defect_cols.get(key); val = str(row_top[c]).strip() if c is not None else ''
                                    shot_item[key] = val if (val and val.lower() not in ['nan', ''] and val.lower() in ['v', 'x', 'o', '1', '?']) else ''
                                
                                # [NEW] Clear repetitive Joint number for 2nd shot and beyond
                                if shot_idx > 1:
                                    shot_item['Joint'] = ''
                                    
                                all_extracted_data.append(shot_item)
                        else:
                            all_extracted_data.append(item_data)
                        r_idx += 2

                else:
                    # ===== 표준 모드 (RT, PT, PMI) =====
                    for _, row in df.iterrows():
                        v_raw_no = clean_v(row[col_no]) if col_no is not None else ""
                        raw_dwg = clean_v(row[col_dwg]) if col_dwg is not None else ""
                        if _is_date_like_dwg(raw_dwg):
                            raw_dwg = ""
                        raw_joint = clean_v(row[col_joint]) if col_joint is not None else ""
                        
                        # 행의 주요 데이터가 모두 비어있으면 빈 행으로 간주하여 건너뜀
                        if not v_raw_no and not raw_dwg and not raw_joint:
                            continue
                            
                        # 주요 데이터는 있는데 순번만 비어있으면 자동 부여
                        if not v_raw_no or v_raw_no == "nan":
                            v_raw_no = str(len(all_extracted_data) + 1)
                            
                        if target_no_list and v_raw_no not in target_no_list: continue

                        extract_key = self.rt_extract_keyword.get().strip().lower()
                        if extract_key:
                            row_vals = [clean_v(v).lower() for v in row.values if pd.notna(v)]
                            row_str = ' '.join(row_vals)
                            if extract_key not in row_str:
                                continue

                        curr_dwg = sheet_level_dwg if sheet_level_dwg else raw_dwg
                        if _is_date_like_dwg(curr_dwg):
                            curr_dwg = ''
                        if (not curr_dwg or curr_dwg == 'nan'):
                            curr_dwg = last_dwg if last_dwg else ''
                        if curr_dwg and curr_dwg != 'nan': last_dwg = curr_dwg

                        curr_joint = clean_v(row[col_joint]) if col_joint is not None else ''
                        if (not curr_joint or curr_joint == 'nan') and last_joint: curr_joint = last_joint
                        if curr_joint and curr_joint != 'nan': last_joint = curr_joint
                        elif not curr_joint or curr_joint == 'nan': curr_joint = v_raw_no

                        curr_date = clean_v(row[col_date]) if col_date is not None else ''
                        if (not curr_date or curr_date.lower() == 'nan') and last_date: curr_date = last_date
                        if curr_date and curr_date.lower() != 'nan': last_date = curr_date
                        elif not curr_date or curr_date.lower() == 'nan': curr_date = sheet_level_date if sheet_level_date else ''

                        v_report = clean_v(row[col_report_no]) if col_report_no is not None else ''
                        if (not v_report or v_report == 'nan') and getattr(self, 'last_report', ''): v_report = self.last_report
                        if v_report and v_report != 'nan': self.last_report = v_report
                        elif not v_report or v_report == 'nan': v_report = ''

                        if mode == 'RT':
                            raw_size = str(row[col_size]).strip() if col_size is not None else ''
                            s_match = re.search(r'(\d+\.?\d*)', raw_size)
                            size_val = s_match.group(1) if s_match else raw_size

                            num_shots = self.calculate_rt_shots(size_val)

                            raw_welder = clean_v(row[col_welder]) if col_welder is not None else ''
                            v_welder = process_welder_string(raw_welder)

                            v_date = curr_date

                            item_data = {
                                'No': v_raw_no, 'Date': v_date, 'Report No': v_report,
                                'Dwg': curr_dwg, 'Joint': curr_joint, 'Loc': '-',
                                'Acc': clean_v(row[col_acc]) if col_acc is not None else '',
                                'Rej': clean_v(row[col_rej]) if col_rej is not None else '',
                                'Deg': clean_v(row[col_deg]) if col_deg is not None else '',
                                'Welder': v_welder,
                                'Remarks': clean_v(row[col_remarks]) if col_remarks is not None else '',
                                'T': clean_v(row[col_t]) if col_t is not None else '',
                                'Mat': clean_v(row[col_mat]) if col_mat is not None else '',
                                'Weld': clean_v(row[col_weld]) if col_weld is not None else '',
                                'IQI': clean_v(row[col_iqi]) if col_iqi is not None else '',
                                'Sens': clean_v(row[col_sens]) if col_sens is not None else '',
                                'Den': clean_v(row[col_den]) if col_den is not None else '',
                                'Result': clean_v(row[col_result]) if col_result is not None else 'ACC',
                                'Size': size_val, 'num_shots': num_shots,
                                'selected': True, 'order_index': len(self.rt_extracted_data) + len(all_extracted_data),
                                '_src': {
                                    'sheet': sheet_name,
                                    'row': _ + header_idx + 2,
                                    'col_result': list(df.columns).index(col_result) + 1 if col_result is not None else None,
                                    'col_remarks': list(df.columns).index(col_remarks) + 1 if col_remarks is not None else None,
                                    'col_acc': list(df.columns).index(col_acc) + 1 if col_acc is not None else None,
                                    'col_rej': list(df.columns).index(col_rej) + 1 if col_rej is not None else None,
                                    'col_defects': {k: list(df.columns).index(v) + 1 for k, v in defect_cols.items()}
                                }
                            }
                            for i in range(1, 16):
                                key = f'D{i}'; src_col = defect_cols.get(key)
                                src_val = str(row[src_col]).strip() if src_col is not None else ''
                                item_data[key] = src_val if src_val else ('1' if i <= num_shots else '')

                            if num_shots > 1:
                                for shot_idx in range(1, num_shots + 1):
                                    shot_item = item_data.copy()
                                    shot_item['Loc'] = f'{shot_idx}-{(shot_idx % num_shots) + 1}'
                                    for i in range(1, 16):
                                        key = f'D{i}'; c = defect_cols.get(key); val = str(row[c]).strip() if c is not None else ''
                                        shot_item[key] = '?' if val and val.lower() in ['v', 'x', 'o', '1', '?'] else ''
                                        
                                    # [NEW] Clear repetitive Joint number for 2nd shot and beyond
                                    if shot_idx > 1:
                                        shot_item['Joint'] = ''
                                        
                                    all_extracted_data.append(shot_item)
                            else:
                                all_extracted_data.append(item_data)
                        elif mode == "PT":
                            # 헤더 행 방어: Dwg/Joint 값이 컬럼 키워드이면 스킵
                            _dwg_up = curr_dwg.upper().strip()
                            _jnt_up = curr_joint.upper().strip()
                            _header_dwg = {"ISO", "LINE", "DWG", "DWG NO", "DWG.", "DRAWING", "ISO/DWG", "ISO NO"}
                            _header_jnt = {"JOINT", "JOINT NO", "JOINT NO.", "JOINT NUMBER", "WELD", "WELD NO"}
                            if _dwg_up in _header_dwg and _jnt_up in _header_jnt:
                                continue
                            res_str = str(row[col_result]).upper() if col_result is not None else "ACC"
                            if any(k in res_str for k in ["ACC", "OK", "ACCEPT", "합격"]):
                                item_row = {
                                    'No': v_raw_no, 'Date': curr_date, 'Report No': v_report, 'Dwg': curr_dwg, 'Joint': self.force_two_digit(curr_joint),
                                    'NPS': str(row[col_size]).strip() if col_size is not None else "", 'Thk.': convert_sch_to_thk(str(row[col_size]).strip() if col_size is not None else "", str(row[col_thk]).strip() if col_thk is not None else ""),
                                    'Material': self.fix_material_name(row[col_mat]) if col_mat is not None else "", 'Welder': str(row[col_welder]).strip() if col_welder is not None else "",
                                    'WType': str(row[col_wtype]).strip() if col_wtype is not None else "", 'Result': "Acc", 'selected': True,
                                    'order_index': len(self.pt_extracted_data) + len(all_extracted_data)
                                }
                                for k in self.pt_column_keys:
                                    if k not in item_row and k != "selected": item_row[k] = ""
                                all_extracted_data.append(item_row)
                        else:
                            # PMI-specific Row Data
                            v_cr = self.to_float(row[col_cr])
                            if v_cr > 0 or (v_raw_no != "" and v_raw_no != "nan"):
                                v_ni = self.to_float(row[col_ni])
                                v_mo = self.to_float(row[col_mo]) if col_mo is not None else 0.0
                                v_mn = self.to_float(row[col_mn]) if col_mn is not None else 0.0
                                orig_grade = str(row[col_grade_orig]).strip() if col_grade_orig is not None else ""
                                
                                final_grade = orig_grade
                                if self.auto_verify.get():
                                    detected = self.check_material_grade({'Cr': v_cr, 'Ni': v_ni, 'Mo': v_mo, 'Mn': v_mn})
                                    if detected: final_grade = detected
                                if not final_grade or final_grade == "nan":
                                    final_grade = "SS316" if v_mo >= 1.5 else "SS304"

                                pmi_item = {
                                    'No': v_raw_no, 
                                    'Joint': curr_joint,
                                    'Loc': str(row[col_loc]).strip() if col_loc is not None else "",
                                    'Cr': v_cr, 'Ni': v_ni, 'Mo': v_mo, 'Mn': v_mn,
                                    'Grade': final_grade, 'Dwg': curr_dwg,
                                    'Date': curr_date or sheet_level_date,
                                    'Report No': v_report,
                                    'selected': True,
                                    'order_index': len(self.extracted_data) + len(all_extracted_data)
                                }
                                # Extract dynamically found elements
                                for c_key, c_idx in custom_cols.items():
                                    pmi_item[c_key] = self.to_float(row[c_idx])
                                # [NEW] 기존 수동 추가 컬럼 보존
                                for k in self.column_keys:
                                    if k not in pmi_item and k != "selected":
                                        pmi_item[k] = ""
                                all_extracted_data.append(pmi_item)
                self.progress['value'] = ((s_idx + 1) / len(xls.sheet_names)) * 50

            if not all_extracted_data:
                messagebox.showerror("오류", "추출된 데이터가 없습니다.")
                return False

            # Extraction Mode Filter (PMI Only)
            if mode == "PMI":
                ext_mode = self.extraction_mode.get()
                if ext_mode != "전체":
                    original_count = len(all_extracted_data)
                    if ext_mode == "SS304 만": all_extracted_data = [d for d in all_extracted_data if d['Grade'] == "SS304"]
                    elif ext_mode == "SS316 만": all_extracted_data = [d for d in all_extracted_data if d['Grade'] == "SS316"]
                    elif ext_mode == "DUPLEX 만": all_extracted_data = [d for d in all_extracted_data if d['Grade'] == "DUPLEX"]
                    elif ext_mode == "SS310 만": all_extracted_data = [d for d in all_extracted_data if d['Grade'] == "SS310"]
                    elif ext_mode == "미분류(기타) 만":
                        known_grades = ["SS304", "SS316", "DUPLEX", "SS310"]
                        all_extracted_data = [d for d in all_extracted_data if d['Grade'] not in known_grades]
                    
                    self.log(f"🔍 필터링 ({ext_mode}): {original_count}개 -> {len(all_extracted_data)}개")
                    if not all_extracted_data:
                        messagebox.showinfo("알림", f"'{ext_mode}'에 해당하는 데이터가 없습니다.")
                        return False

            # [NEW] 원소 함량 필터링 적용 (PMI 전용)
            if mode == "PMI" and hasattr(self, 'element_filters') and self.element_filters:
                for f_item in self.element_filters:
                    f_key = f_item['key'].get().strip()
                    if not f_key: continue
                    f_min = self.to_float(f_item['min'].get())
                    f_max = self.to_float(f_item['max'].get())
                    
                    if f_min > 0 or f_max > 0:
                        original_count = len(all_extracted_data)
                        all_extracted_data = [d for d in all_extracted_data if (f_min <= d.get(f_key, 0.0) <= (f_max if f_max > 0 else 999.9))]
                        if len(all_extracted_data) != original_count:
                             self.log(f"🔍 원소 필터 적용 ({f_key}: {f_min}~{f_max}): {original_count} -> {len(all_extracted_data)}건 남음")

            if mode == "PT":
                # [NEW] [Refinement] Duplicate removal for PT (ISO + Joint)
                seen = set()
                unique_data = []
                for d in all_extracted_data:
                    key = (str(d['Dwg']).strip(), str(d['Joint']).strip())
                    if key not in seen:
                        seen.add(key)
                        unique_data.append(d)
                all_extracted_data = unique_data

            if target_no_list:
                all_extracted_data.sort(key=lambda x: target_no_list.index(str(x['No'])) if str(x['No']) in target_no_list else 999999)

            # [CHANGE] Overwrite -> Accumulate + Auto Sort
            if mode == "RT":
                self.rt_extracted_data.extend(all_extracted_data)
                self.update_date_listbox("RT")
                self._update_rt_preview_columns(mode="RT") # [FIX] 표시 이름 보존을 위해 호출
                self.rt_sort_col = "" # [NEW] Force Ascending
                self.sort_by_column("Dwg", mode="RT") 
                total_count = len(self.rt_extracted_data)
            elif mode == "KOGAS":
                self.kogas_extracted_data.extend(all_extracted_data)
                self.update_date_listbox("KOGAS")
                self._update_rt_preview_columns(mode="KOGAS") # [FIX] 표시 이름 보존을 위해 호출
                self.kogas_sort_col = "" # [NEW] Force Ascending
                self.sort_by_column("Dwg", mode="KOGAS")
                total_count = len(self.kogas_extracted_data)
            elif mode == "PT":
                self.pt_extracted_data.extend(all_extracted_data)
                self.update_date_listbox("PT")
                self.pt_sort_col = ""
                self.mt_sort_col = ""
                self.sort_by_column("Dwg", mode="PT") 
                total_count = len(self.pt_extracted_data)
            else:
                self.extracted_data.extend(all_extracted_data)
                self.pmi_sort_col = "" # [NEW] Force Ascending
                self.sort_by_column("Dwg", mode="PMI")
                # Refresh after sorting/populating so the visible PMI sidebar is
                # updated last and cannot be left with its initial empty state.
                self.update_date_listbox("PMI")
                # Run once more through Tk's event loop. This guarantees that the
                # visible widget is updated after all preview redraw callbacks.
                self.root.after_idle(lambda: self.update_date_listbox("PMI"))
                total_count = len(self.extracted_data)
            
            self.progress['value'] = 100
            if show_msg:
                self.log(f"✅ {mode} 데이터 누적 완료 (현재 총 {total_count} 건)")
                messagebox.showinfo("완료", f"데이터가 추가되었습니다.\n현재 목록에는 총 {total_count}건의 데이터가 있습니다.")
            return True
        except Exception as e:
            self.log(f"❌ {mode} 추출 오류: {e}")
            traceback.print_exc()
            return False

    def sync_results_to_request(self):
        """입력된 결과를 원본 의뢰서 파일로 역전송 (RT/KOGAS 대응)"""
        try:
            main_tab = self.mode_notebook.tab(self.mode_notebook.select(), "text")
            if "RT" in main_tab:
                sub_tab = self.rt_preview_nb.tab(self.rt_preview_nb.select(), "text")
                mode = "KOGAS" if "가스공사" in sub_tab else "RT"
            else:
                messagebox.showwarning("경고", "현재 RT 모드에서만 지원됩니다.")
                return

            target_file = self.rt_target_file_path.get() if mode == "RT" else self.kogas_target_file_path.get()
            if not target_file or not os.path.exists(target_file):
                messagebox.showwarning("파일 오류", "원본 의뢰서 파일을 찾을 수 없습니다.")
                return

            data = self.rt_extracted_data if mode == "RT" else self.kogas_extracted_data
            if not data:
                messagebox.showwarning("데이터 오류", "반영할 데이터가 없습니다.")
                return

            if not messagebox.askyesno("확인", f"현재 입력된 결과를 원본 의뢰서({os.path.basename(target_file)})에 반영하시겠습니까?"):
                return

            self.log(f"🔄 의뢰서 결과 반영 시작: {os.path.basename(target_file)}")
            
            import openpyxl
            wb = openpyxl.load_workbook(target_file)
            count = 0
            
            # Group items by source row to handle multi-shot joints
            groups = {}
            for item in data:
                if not item.get('date_filtered', True): continue # [NEW] Apply Date filter
                if not item.get('_src'): continue
                src_key = (item['_src']['sheet'], item['_src']['row'])
                if src_key not in groups: groups[src_key] = []
                groups[src_key].append(item)

            for (sheet_name, row_idx), items in groups.items():
                if sheet_name not in wb.sheetnames: continue
                ws = wb[sheet_name]
                src = items[0]['_src']
                
                # Determine final result (If any item is Rej, result is Rej)
                final_res = "ACC"
                all_remarks = []
                
                if mode == "KOGAS":
                    for it in items:
                        res = str(it.get('Result', '')).upper()
                        if any(x in res for x in ["REJ", "NG", "불합격"]): final_res = "REJ"
                        rem = str(it.get('Remarks', '')).strip()
                        if rem and rem not in all_remarks: all_remarks.append(rem)
                    
                    # KOGAS 비고란은 컬럼 11 (K열)
                    ws.cell(row=row_idx, column=11, value=", ".join(all_remarks) if all_remarks else None)
                    
                    # NDT 의뢰방식 확인 (13:RT, 14:UT, 15:PT)
                    rt_val = ws.cell(row=row_idx, column=13).value
                    ut_val = ws.cell(row=row_idx, column=14).value
                    pt_val = ws.cell(row=row_idx, column=15).value
                    
                    def is_marked(v):
                        if v is None: return False
                        cleaned = str(v).strip().upper()
                        return cleaned in ["O", "V", "X", "●", "합격", "1", "TRUE"]
                    
                    rt_req = is_marked(rt_val)
                    ut_req = is_marked(ut_val)
                    pt_req = is_marked(pt_val)
                    
                    has_rt_films = False
                    for it in items:
                        loc_str = str(it.get('Loc', '')).strip()
                        if re.match(r'^\d+', loc_str):
                            has_rt_films = True
                            break
                    
                    defect_map = {
                        "D1": "C", "D2": "IP", "D3": "LF", "D4": "S", "D5": "P",
                        "D6": "UC", "D7": "RUC", "D8": "BT", "D9": "TI", "D10": "CP",
                        "D11": "RC", "D12": "Mis", "D13": "EP", "D14": "SD", "D15": "Oth"
                    }
                    
                    def get_item_val(it):
                        grade = str(it.get('Deg', '')).strip()
                        if grade.lower() == 'nan':
                            grade = ''
                        checked = []
                        for d_key, abbrev in defect_map.items():
                            val = str(it.get(d_key, '')).strip()
                            if val in ["√", "1", "v", "V", "o", "O", "●", "?"]:
                                checked.append(abbrev)
                        if checked:
                            return f"{grade if grade else '4'}/{','.join(checked)}"
                        else:
                            if any(x in str(it.get('Result', '')).upper() for x in ["ACC", "OK"]):
                                return "1"
                            return grade if grade else "4"

                    # RT (Q~W: 17~23) 반영
                    if rt_req or has_rt_films:
                        for col in range(17, 24):
                            ws.cell(row=row_idx, column=col, value=None)
                        for it in items:
                            loc_str = str(it.get('Loc', '')).strip()
                            film_num = None
                            match = re.match(r'^(\d+)', loc_str)
                            if match:
                                film_num = int(match.group(1))
                            if film_num and 1 <= film_num <= 7:
                                val = get_item_val(it)
                                ws.cell(row=row_idx, column=16 + film_num, value=val)
                                
                    # UT/UP (X: 24) 반영
                    if ut_req:
                        it = items[0] if items else {}
                        val = get_item_val(it)
                        ws.cell(row=row_idx, column=24, value=val)
                        
                    # PT (Y: 25) 반영
                    if pt_req:
                        it = items[0] if items else {}
                        val = get_item_val(it)
                        ws.cell(row=row_idx, column=25, value=val)
                else:
                    has_defects = {k: False for k in src.get('col_defects', {}).keys()}
                    
                    for it in items:
                        res = str(it.get('Result', '')).upper()
                        if any(x in res for x in ["REJ", "NG", "불합격"]): final_res = "REJ"
                        
                        rem = str(it.get('Remarks', '')).strip()
                        if rem and rem not in all_remarks: all_remarks.append(rem)
                        
                        if src.get('col_defects'):
                            for d_key, col_idx in src['col_defects'].items():
                                if it.get(d_key) in ["√", "1", "v", "V", "o", "O"]: has_defects[d_key] = True

                    # Write back to Excel
                    if src.get('col_result'):
                        ws.cell(row=row_idx, column=src['col_result'], value=final_res)
                    
                    if src.get('col_remarks') and all_remarks:
                        ws.cell(row=row_idx, column=src['col_remarks'], value=", ".join(all_remarks))
                    
                    if src.get('col_defects'):
                        for d_key, col_idx in src['col_defects'].items():
                            if has_defects[d_key]:
                                ws.cell(row=row_idx, column=col_idx, value="√")
                
                count += 1

            wb.save(target_file)
            self.log(f"✅ 총 {count}개 행의 결과를 의뢰서에 성공적으로 반영했습니다.")
            messagebox.showinfo("완료", f"{count}개 행의 결과가 의뢰서에 반영되었습니다.")

        except PermissionError:
            msg = f"파일이 Excel에서 열려 있어 저장할 수 없습니다.\n\n'{os.path.basename(target_file)}' 파일을 Excel에서 닫은 후 다시 시도해 주세요."
            self.log(f"❌ 의뢰서 결과 반영 실패: 파일이 열려 있음 (PermissionError)")
            messagebox.showerror("파일 접근 오류", msg)
        except Exception as e:
            self.log(f"❌ 의뢰서 결과 반영 실패: {e}")
            messagebox.showerror("오류", f"의뢰서 결과 반영 중 오류가 발생했습니다: {e}")

    def duplicate_selected_rows(self):
        """선택된 행들을 복제하여 데이터 리스트에 삽입"""
        selected = self.preview_tree.selection()
        if not selected: return
        
        # 시각적 순서대로 정렬 (위에서부터 복제)
        selected = sorted(list(selected), key=lambda x: self.preview_tree.index(x))
        
        new_items = []
        insert_pos = 0
        for item_id in selected:
            view_idx = self.preview_tree.index(item_id)
            actual_idx = self.item_idx_map[view_idx]
            new_item = self.extracted_data[actual_idx].copy()
            # No.는 중복 방지를 위해 비우거나 새로 부여 (여기서는 유지 후 나중에 rename 가능)
            new_items.append(new_item)
            insert_pos = actual_idx
            
        for item in reversed(new_items):
            self.extracted_data.insert(insert_pos + 1, item)
             
        self.populate_preview(self.extracted_data, switch_tab=False)
        self.log(f"👯 {len(new_items)}개 행 복제 완료")

    def save_column_widths(self):
        """현재 트리뷰의 컬럼 너비를 설정 파일에 저장"""
        widths = {}
        for col in self.preview_tree["columns"]:
            widths[col] = self.preview_tree.column(col, "width")
        
        self.config["PMI_COL_WIDTHS"] = widths
        self.save_settings()
        self.log("📏 컬럼 너비 설정 저장 완료")

    def apply_date_filter(self):
        """날짜 리스트박스 선택 상태를 데이터 모델에 반영"""
        selected_dates = []
        for i in range(self.date_listbox.size()):
            val = self.date_listbox.get(i)
            if val.startswith("[v]"):
                selected_dates.append(val.replace("[v] ", "").strip())
        
        if not selected_dates:
            messagebox.showwarning("필터 오류", "최소 하나 이상의 날짜를 선택해주세요.")
            return

        for item in self.extracted_data:
            item['date_filtered'] = (item.get('Date', '') in selected_dates)
        
        self.populate_preview(self.extracted_data, switch_tab=False)
        self.log(f"📅 날짜 필터 적용 완료: {len(selected_dates)}개 날짜 선택됨")

    def show_report_issue_history(self):
        from site_apps.central.src.reports.report_issue_history import show_history
        show_history(self.root)

    def _report_save_path(self, directory, extension, mode=None, choose_directory=True):
        from site_apps.central.src.reports.report_issue_history import report_output_path
        if choose_directory:
            mode = mode or getattr(self, 'current_mode', 'PMI')
            key = f'{mode}_LAST_SAVE_DIR'
            remembered = str(self.config.get(key, '')).strip()
            initial = remembered if remembered and os.path.isdir(remembered) else directory
            directory = filedialog.askdirectory(parent=self.root, title=f'{mode} 보고서 저장 폴더 선택',
                                                initialdir=initial, mustexist=True)
            if not directory:
                self.log(f'{mode} 보고서 저장이 취소되었습니다.')
                return None
            self.config[key] = directory
            self.save_settings()
        return report_output_path(directory, self.gapji_report_no.get(), extension,
                                  fallback='Report_' + datetime.datetime.now().strftime('%Y%m%d_%H%M%S'))

    def _record_report_issue(self, mode, workbook, save_path, items):
        from site_apps.central.src.reports.report_issue_history import record_issue
        try:
            # Read the actual saved cover value rather than a row's imported number.
            cell = {'RT': 'C5', 'KOGAS': 'C5', 'PT': 'P6', 'MT': 'S6', 'PAUT': 'AA5'}.get(mode)
            if cell is None:
                cell = self.config.get('GAPJI_REPORT_NO_CELL', 'B10')
            report_no = workbook.worksheets[0][cell].value
            record_issue(mode, report_no, save_path, items)
            if mode == 'PAUT':
                self.populate_preview(self.paut_extracted_data, switch_tab=False, mode=mode)
            elif mode in ('MT', 'RT', 'KOGAS'):
                self.populate_preview(self._get_mode_info(mode)[2], switch_tab=False, mode=mode)
            self.log(f"발행 이력 저장: {report_no or '(번호 미입력)'} · {len(items)}건")
        except Exception as exc:
            self.log(f"발행 이력 저장 실패: {exc}")
            messagebox.showwarning('발행 이력 저장 실패',
                f'보고서 파일은 생성되었으나 발행 이력을 저장하지 못했습니다.\n{save_path}\n\n{exc}')


    def _create_top_action_bar(self, parent, mode):
        """데이터 미리보기 상단에 모던 플랫 액션 툴바 생성"""
        action_bar = tk.Frame(parent, background="#ffffff", padx=8, pady=6)
        action_bar.pack(fill='x')

        # 1. 성적서 생성 / 시작 (파란색 하이라이트)
        if mode == "PAUT":
            cmd_run = self._generate_paut_report
            text_run = "✨ 성적서 생성"
        else:
            cmd_run = self.run_process
            text_run = "✨ 성적서 생성"

        btn_run = tk.Button(
            action_bar, text=text_run, command=cmd_run,
            font=("Malgun Gothic", 10, "bold"), bg="#2563eb", fg="#ffffff",
            activebackground="#1d4ed8", activeforeground="#ffffff",
            relief="flat", cursor="hand2", padx=12, pady=5, bd=0
        )
        btn_run.pack(side='left', fill='x', expand=True, padx=(0, 6))
        btn_run.bind("<Enter>", lambda e, b=btn_run: b.configure(bg="#1d4ed8"))
        btn_run.bind("<Leave>", lambda e, b=btn_run: b.configure(bg="#2563eb"))

        # 2. 데이터 추출
        if mode == "PAUT":
            cmd_ext = self._extract_paut_data
        else:
            cmd_ext = self.extract_only

        btn_extract = tk.Button(
            action_bar, text="📝 데이터 추출", command=cmd_ext,
            font=("Malgun Gothic", 9, "bold"), bg="#f8fafc", fg="#334155",
            activebackground="#e2e8f0", activeforeground="#1e293b",
            relief="solid", bd=1, cursor="hand2", padx=10, pady=5
        )
        btn_extract.pack(side='left', fill='x', expand=True, padx=(0, 6))
        btn_extract.bind("<Enter>", lambda e, b=btn_extract: b.configure(bg="#e2e8f0"))
        btn_extract.bind("<Leave>", lambda e, b=btn_extract: b.configure(bg="#f8fafc"))

        # 3. 일보 NDT 불러오기
        btn_daily = tk.Button(
            action_bar, text="📋 일보 NDT 불러오기", command=self.load_daily_work_history,
            font=("Malgun Gothic", 9, "bold"), bg="#f8fafc", fg="#334155",
            activebackground="#e2e8f0", activeforeground="#1e293b",
            relief="solid", bd=1, cursor="hand2", padx=10, pady=5
        )
        has_extra = mode in ("RT", "PAUT")
        btn_daily.pack(side='left', fill='x', expand=True, padx=(0, 6) if has_extra else 0)
        btn_daily.bind("<Enter>", lambda e, b=btn_daily: b.configure(bg="#e2e8f0"))
        btn_daily.bind("<Leave>", lambda e, b=btn_daily: b.configure(bg="#f8fafc"))

        # RT 모드 추가 액션: 의뢰서 결과 반영
        if mode == "RT":
            btn_sync = tk.Button(
                action_bar, text="🔄 의뢰서에 결과 반영", command=self.sync_results_to_request,
                font=("Malgun Gothic", 9, "bold"), bg="#f8fafc", fg="#334155",
                activebackground="#e2e8f0", activeforeground="#1e293b",
                relief="solid", bd=1, cursor="hand2", padx=8, pady=5
            )
            btn_sync.pack(side='left', fill='x', expand=True)
            btn_sync.bind("<Enter>", lambda e, b=btn_sync: b.configure(bg="#e2e8f0"))
            btn_sync.bind("<Leave>", lambda e, b=btn_sync: b.configure(bg="#f8fafc"))

        # PAUT 모드 추가 액션: 일괄 판정
        elif mode == "PAUT":
            btn_eval = tk.Button(
                action_bar, text="⚡ 일괄 판정", command=self._run_batch_paut_eval,
                font=("Malgun Gothic", 9, "bold"), bg="#f8fafc", fg="#334155",
                activebackground="#e2e8f0", activeforeground="#1e293b",
                relief="solid", bd=1, cursor="hand2", padx=8, pady=5
            )
            btn_eval.pack(side='left', fill='x', expand=True)
            btn_eval.bind("<Enter>", lambda e, b=btn_eval: b.configure(bg="#e2e8f0"))
            btn_eval.bind("<Leave>", lambda e, b=btn_eval: b.configure(bg="#f8fafc"))

        # 구분선
        tk.Frame(parent, height=1, background="#e2e8f0").pack(fill='x', pady=(0, 4))
        return action_bar

    def run_process(self):
        self.save_settings() # [FIX] Ensure UI settings are saved to config before generating
        
        # [NEW] Ensure config values are correct types for comparison (prevent str vs int errors)
        for k in list(self.config.keys()):
            if k.endswith(('_ROW', '_IDX', '_SIZE')) or any(x in k for x in ['START', 'END', 'PAGE']):
                try: self.config[k] = int(float(self.config[k]))
                except: pass
            elif any(x in k for x in ['MARGIN', 'SCALE', 'RATIO', 'POS', 'HEIGHT', 'WIDTH']):
                try: self.config[k] = float(self.config[k])
                except: pass

        # 현재 활성 탭에 따른 모드 결정
        try:
            tab_text = self.mode_notebook.tab(self.mode_notebook.select(), "text")
            if "RT" in tab_text:
                sub_tab = self.rt_preview_nb.tab(self.rt_preview_nb.select(), "text")
                mode = "KOGAS" if "가스공사" in sub_tab else "RT"
            elif "PT" in tab_text: mode = "PT"
            elif "MT" in tab_text: mode = "MT"
            elif "PAUT" in tab_text: mode = "PAUT"
            else: mode = "PMI"
        except: mode = "PMI"

        if mode == "RT":
            target_file = self.rt_target_file_path.get()
            template_path = self.rt_template_file_path.get()
            data = self.rt_extracted_data
        elif mode == "KOGAS":
            target_file = self.kogas_target_file_path.get()
            template_path = self.kogas_template_file_path.get()
            data = self.kogas_extracted_data
        elif mode == "PT":
            target_file = self.pt_target_file_path.get()
            template_path = self.pt_template_file_path.get()
            data = self.pt_extracted_data
        elif mode == "MT":
            target_file = self.mt_target_file_path.get()
            template_path = self.mt_template_file_path.get()
            data = self.mt_extracted_data
        else:
            target_file = self.target_file_path.get()
            template_path = self.template_file_path.get()
            data = self.extracted_data
        
        if not template_path:
            messagebox.showwarning("파일 미선택", f"{mode} 양식(Template) 파일을 선택해주세요.")
            return
            
        if not target_file and not data:
            messagebox.showwarning("파일 미선택", f"{mode} 데이터 파일(Excel)을 선택하거나, 저장된 데이터를 불러와주세요.")
            return

        # [NEW] 템플릿 파일 존재 여부 및 기본 구조 확인
        if not os.path.exists(template_path):
            messagebox.showerror("오류", f"템플릿 파일을 찾을 수 없습니다:\n{template_path}")
            return
        self.save_settings() # This now captures all setting_vars (Logo X/Y, Print Area, etc.)
        
        # 데이터가 비어있거나 새로 추출이 필요한 경우 수행
        if not data:
            if not self.extract_only(show_msg=False): return
            # Re-fetch data after extraction
            if mode == "PT": data = self.pt_extracted_data
            elif mode == "MT": data = self.mt_extracted_data
            elif mode == "RT": data = self.rt_extracted_data
            elif mode == "KOGAS": data = self.kogas_extracted_data
            else: data = self.extracted_data
            
        # [NEW] 체크된 항목만 필터링 (기본값은 True)
        final_list = [
            d for d in data
            if d.get('selected', True)
            and d.get('date_filtered', True)
            and (mode not in ('MT', 'RT', 'KOGAS') or d.get('drawing_filtered', True))
            and (mode != "PAUT" or (
                d.get('line_filtered', True) and d.get('size_filtered', True)))
        ]
        if not final_list:
            messagebox.showwarning("항목 미선택", f"선택된 {mode} 데이터가 없습니다. 미리보기에서 항목을 체크해주세요.")
            return

        if mode == "RT":
            self._run_rt_process(final_list, template_path, mode="RT")
        elif mode == "KOGAS":
            self._run_rt_process(final_list, template_path, mode="KOGAS")
        elif mode == "PT":
            self._run_pt_process(final_list, template_path)
        elif mode == "MT":
            self._run_mt_process(final_list, template_path)
        else:
            self._run_pmi_process(final_list, template_path)

    def _on_entry_esc(self, event):
        """Removes focus and clears selection on ESC (preserves text)."""
        try:
            widget = event.widget
            if hasattr(widget, 'selection_clear'):
                widget.selection_clear()
            self.root.focus_set()
        except: pass

if __name__ == "__main__":
    root = tk.Tk()
    app = PMIReportApp(root)
    try:
        root.mainloop()
    except KeyboardInterrupt:
        # Ctrl+C from a terminal is a normal shutdown request, not an app error.
        try:
            app.on_closing()
        except tk.TclError:
            # The window may already have been destroyed while handling shutdown.
            pass
