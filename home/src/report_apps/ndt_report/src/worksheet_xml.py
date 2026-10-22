"""Preserve worksheet XML while keeping print options in schema order."""
import re


def normalize_print_options(xml):
    pattern = r'<printOptions\b[^>]*/>|<printOptions\b[^>]*>.*?</printOptions>'
    matches = re.findall(pattern, xml, flags=re.DOTALL)
    if not matches:
        return xml
    xml = re.sub(pattern, '', xml, flags=re.DOTALL)
    # All of these follow printOptions in CT_Worksheet. mergeCells and
    # phoneticPr precede it. Do not insert directly after sheetData.
    following = re.search(
        r'<(?:pageMargins|pageSetup|headerFooter|rowBreaks|colBreaks|customProperties|'
        r'cellWatches|ignoredErrors|smartTags|drawing|legacyDrawing|legacyDrawingHF|'
        r'picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b', xml
    )
    index = following.start() if following else xml.rfind('</worksheet>')
    if index < 0:
        raise ValueError('worksheet closing tag is missing')
    return xml[:index] + matches[-1] + xml[index:]
