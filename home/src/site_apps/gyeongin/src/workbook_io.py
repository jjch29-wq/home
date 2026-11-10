"""Publish complete, CRC-validated workbooks without in-place writes."""
import os
from pathlib import Path
import shutil
import tempfile
import zipfile

import pandas as pd


def replace_workbook(source, target):
    with zipfile.ZipFile(source) as archive:
        bad = archive.testzip()
        if bad:
            raise ValueError(f'Workbook CRC check failed: {bad}')
    os.replace(source, target)


def write_workbook(target, sheets):
    target = Path(target)
    descriptor, temporary = tempfile.mkstemp(prefix='.sync-save-', suffix='.xlsx', dir=target.parent)
    os.close(descriptor)
    try:
        with pd.ExcelWriter(temporary, engine='openpyxl') as writer:
            for name, frame in sheets.items():
                frame.to_excel(writer, sheet_name=name, index=False)
        replace_workbook(temporary, target)
    finally:
        if os.path.exists(temporary):
            os.remove(temporary)


def copy_workbook(source, target):
    target = Path(target)
    descriptor, temporary = tempfile.mkstemp(prefix='.mirror-save-', suffix='.xlsx', dir=target.parent)
    os.close(descriptor)
    try:
        shutil.copy2(source, temporary)
        replace_workbook(temporary, target)
    finally:
        if os.path.exists(temporary):
            os.remove(temporary)
