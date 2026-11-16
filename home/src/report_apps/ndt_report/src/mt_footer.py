"""Preserve template footer pictures that openpyxl does not round-trip."""
import os
import posixpath
import re
import tempfile
import zipfile
from lxml import etree as ET

MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
DOCREL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
CONTENT = 'http://schemas.openxmlformats.org/package/2006/content-types'


def _insert_in_schema_order(root, new_element):
    """Insert element into worksheet root strictly following OpenXML XSD sequence."""
    SCHEMA_ORDER = [
        'sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData',
        'sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter',
        'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr',
        'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions',
        'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks',
        'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing',
        'legacyDrawing', 'legacyDrawingHF', 'picture', 'oleObjects', 'drawingHF'
    ]
    new_tag = new_element.tag.split('}')[-1]
    new_order = SCHEMA_ORDER.index(new_tag) if new_tag in SCHEMA_ORDER else 999
    for i, child in enumerate(root):
        c_tag = child.tag.split('}')[-1]
        c_order = SCHEMA_ORDER.index(c_tag) if c_tag in SCHEMA_ORDER else 999
        if c_order > new_order:
            root.insert(i, new_element)
            return
    root.append(new_element)


def resolve(part, target):
    return target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(part), target))


def restore_mt_footer_pictures(template, output):
    with zipfile.ZipFile(template) as archive:
        original = {name: archive.read(name) for name in archive.namelist()}
    with zipfile.ZipFile(output) as archive:
        result = {name: archive.read(name) for name in archive.namelist()}
    types = ET.fromstring(result['[Content_Types].xml'])
    def ensure_type(extension, content_type):
        if not any(node.get('Extension') == extension for node in types):
            ET.SubElement(types, f'{{{CONTENT}}}Default', Extension=extension, ContentType=content_type)
    sheets = sorted((name for name in result if re.fullmatch(r'xl/worksheets/sheet\d+\.xml', name)),
                    key=lambda name: int(re.search(r'(\d+)\.xml$', name)[1]))
    for index, target_sheet in enumerate(sheets, 1):
        source_sheet = f'xl/worksheets/sheet{1 if index == 1 else 2}.xml'
        source_rels = posixpath.dirname(source_sheet) + '/_rels/' + posixpath.basename(source_sheet) + '.rels'
        if source_sheet not in original or source_rels not in original:
            continue
        source_xml = ET.fromstring(original[source_sheet])
        footer_ref = source_xml.find(f'{{{MAIN}}}legacyDrawingHF')
        if footer_ref is None:
            continue
        relationships = ET.fromstring(original[source_rels])
        source_link = next(node for node in relationships if node.get('Id') == footer_ref.get(f'{{{DOCREL}}}id'))
        vml_source = resolve(source_sheet, source_link.get('Target'))
        vml_target = f'xl/drawings/mt_footer_{index}.vml'
        result[vml_target] = original[vml_source]
        vml_rel_source = posixpath.dirname(vml_source) + '/_rels/' + posixpath.basename(vml_source) + '.rels'
        vml_rels = ET.fromstring(original[vml_rel_source])
        for number, link in enumerate(vml_rels, 1):
            if link.get('TargetMode') == 'External':
                continue
            media_source = resolve(vml_source, link.get('Target'))
            extension = posixpath.splitext(media_source)[1]
            media_target = f'xl/media/mt_footer_{index}_{number}{extension}'
            result[media_target] = original[media_source]
            link.set('Target', posixpath.relpath(media_target, posixpath.dirname(vml_target)))
            ensure_type(extension.lstrip('.'), 'image/jpeg' if extension.lower() in ('.jpg', '.jpeg') else 'image/png')
        result[f'xl/drawings/_rels/mt_footer_{index}.vml.rels'] = ET.tostring(vml_rels)
        ensure_type('vml', 'application/vnd.openxmlformats-officedocument.vmlDrawing')
        rel_path = posixpath.dirname(target_sheet) + '/_rels/' + posixpath.basename(target_sheet) + '.rels'
        target_rels = ET.fromstring(result[rel_path]) if rel_path in result else ET.Element(f'{{{REL}}}Relationships', nsmap={None: REL})
        rel_id = 'rIdMTFooter'
        used = {node.get('Id') for node in target_rels}
        while rel_id in used:
            rel_id += '_'
        ET.SubElement(target_rels, f'{{{REL}}}Relationship', Id=rel_id,
                      Type=DOCREL + '/vmlDrawing', Target='../drawings/' + posixpath.basename(vml_target))
        result[rel_path] = ET.tostring(target_rels)
        target_xml = ET.fromstring(result[target_sheet])
        for old in target_xml.findall(f'{{{MAIN}}}legacyDrawingHF'):
            target_xml.remove(old)
        new_ref = ET.Element(f'{{{MAIN}}}legacyDrawingHF')
        new_ref.set(f'{{{DOCREL}}}id', rel_id)
        header_footer = target_xml.find(f'{{{MAIN}}}headerFooter')
        if header_footer is None:
            import copy
            header_footer = copy.deepcopy(source_xml.find(f'{{{MAIN}}}headerFooter'))
            _insert_in_schema_order(target_xml, header_footer)
        else:
            odd_footer = header_footer.find(f'{{{MAIN}}}oddFooter')
            source_hf = source_xml.find(f'{{{MAIN}}}headerFooter')
            source_odd = source_hf.find(f'{{{MAIN}}}oddFooter') if source_hf is not None else None
            if odd_footer is not None and odd_footer.text:
                if '&G' not in odd_footer.text and source_odd is not None and '&G' in (source_odd.text or ''):
                    odd_footer.text += '&R&"바탕,굵게"&10 &G'
            elif odd_footer is None and source_odd is not None:
                import copy
                header_footer.append(copy.deepcopy(source_odd))
        _insert_in_schema_order(target_xml, new_ref)
        result[target_sheet] = ET.tostring(target_xml)
    result['[Content_Types].xml'] = ET.tostring(types)
    handle, temporary = tempfile.mkstemp(suffix='.xlsx', dir=os.path.dirname(os.path.abspath(output)))
    os.close(handle)
    try:
        with zipfile.ZipFile(temporary, 'w', zipfile.ZIP_DEFLATED) as archive:
            for name, data in result.items():
                archive.writestr(name, data)
        os.replace(temporary, output)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)


# Alias for general template footer restoration (PMI, MT, etc.)
restore_template_footer_pictures = restore_mt_footer_pictures
