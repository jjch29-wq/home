"""PAUT field compatibility for saved sessions and imported spreadsheets."""
import re

TCG_MAX = 'TCG MAX(%)'


def tcg_from_text(text):
    match = re.search(
        r'(?<![A-Za-z])TCG\s*(?:MAX\s*)?(?:\(\s*%\s*\)\s*)?'
        r'[:=]?\s*(\d+(?:\.\d+)?)\s*[%％]', str(text or ''), re.I)
    return match[1] if match else ''


def normalize_tcg(item):
    def present(value):
        return value is not None and str(value).strip().lower() not in ('', 'nan', 'none')

    if present(item.get(TCG_MAX)):
        return item
    for key, value in item.items():
        normalized = re.sub(r'[^A-Z0-9]', '', str(key).upper())
        if normalized in ('TCGMAX', 'TCGMAXPERCENT') and present(value):
            item[TCG_MAX] = value
            return item
    item[TCG_MAX] = next((value for key in ('_source_paut_section', '구간정보', 'Remarks')
                         if (value := tcg_from_text(item.get(key, '')))), '')
    return item
