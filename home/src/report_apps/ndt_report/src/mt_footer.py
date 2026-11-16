"""Preserve template footer pictures that openpyxl does not round-trip.
Uses Python built-in xml.etree.ElementTree (no external lxml dependency required).
"""
import os
import posixpath
import re
import tempfile
import zipfile
import xml.etree.ElementTree as ET

MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
DOCREL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
CONTENT = 'http://schemas.openxmlformats.org/package/2006/content-types'

ET.register_namespace('', MAIN)
ET.register_namespace('r', DOCREL)
ET.register_namespace('rel', REL)
ET.register_namespace('ct', CONTENT)


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
    """Restore template footer pictures (VML drawings & media) to the output excel file.
    If template lacks VML drawings, automatically falls back to clean G_DRIVE template.
    """
    if not os.path.exists(template) or not os.path.exists(output):
        return

    # Fallback to pristine template with VML drawings if the specified template lacks them
    actual_template = template
    has_vml = False
    try:
        with zipfile.ZipFile(template, 'r') as z_chk:
            has_vml = any('vml' in n.lower() for n in z_chk.namelist())
    except Exception:
        pass

    if not has_vml:
        candidate_fallbacks = [
            os.path.join(os.path.dirname(template), 'G_DRIVE_PMI-템플릿보고서.xlsx'),
            os.path.abspath('home/data/G_DRIVE_PMI-템플릿보고서.xlsx'),
        ]
        for cand in candidate_fallbacks:
            if os.path.exists(cand):
                try:
                    with zipfile.ZipFile(cand, 'r') as z_fb:
                        if any('vml' in n.lower() for n in z_fb.namelist()):
                            actual_template = cand
                            has_vml = True
                            break
                except Exception:
                    pass

    if not has_vml:
        return

    with zipfile.ZipFile(actual_template) as archive:
        original = {name: archive.read(name) for name in archive.namelist()}
    with zipfile.ZipFile(output) as archive:
        result = {name: archive.read(name) for name in archive.namelist()}

    types = ET.fromstring(result['[Content_Types].xml'])

    def ensure_type(extension, content_type):
        if not any(node.get('Extension') == extension for node in types):
            ET.SubElement(types, f'{{{CONTENT}}}Default', {'Extension': extension, 'ContentType': content_type})

    sheets = sorted((name for name in result if re.fullmatch(r'xl/worksheets/sheet\d+\.xml', name)),
                    key=lambda name: int(re.search(r'(\d+)\.xml$', name)[1]))

    for index, target_sheet in enumerate(sheets, 1):
        source_sheet = f'xl/worksheets/sheet{1 if index == 1 else 2}.xml'
        source_rels = posixpath.dirname(source_sheet) + '/_rels/' + posixpath.basename(source_sheet) + '.rels'
        if source_sheet not in original or source_rels not in original:
            # Fallback to sheet1 if sheet2 rels not present
            source_sheet = 'xl/worksheets/sheet1.xml'
            source_rels = 'xl/worksheets/_rels/sheet1.xml.rels'
            if source_sheet not in original or source_rels not in original:
                continue

        source_xml = ET.fromstring(original[source_sheet])
        footer_ref = source_xml.find(f'{{{MAIN}}}legacyDrawingHF')
        if footer_ref is None:
            continue

        relationships = ET.fromstring(original[source_rels])
        matching_nodes = [node for node in relationships if node.get('Id') == footer_ref.get(f'{{{DOCREL}}}id')]
        if not matching_nodes:
            continue
        source_link = matching_nodes[0]

        vml_source = resolve(source_sheet, source_link.get('Target'))
        if vml_source not in original:
            continue
        vml_target = f'xl/drawings/mt_footer_{index}.vml'
        result[vml_target] = original[vml_source]

        vml_rel_source = posixpath.dirname(vml_source) + '/_rels/' + posixpath.basename(vml_source) + '.rels'
        if vml_rel_source in original:
            vml_rels = ET.fromstring(original[vml_rel_source])
            for number, link in enumerate(vml_rels, 1):
                if link.get('TargetMode') == 'External':
                    continue
                media_source = resolve(vml_source, link.get('Target'))
                if media_source in original:
                    extension = posixpath.splitext(media_source)[1]
                    media_target = f'xl/media/mt_footer_{index}_{number}{extension}'
                    result[media_target] = original[media_source]
                    link.set('Target', posixpath.relpath(media_target, posixpath.dirname(vml_target)))
                    ensure_type(extension.lstrip('.'), 'image/jpeg' if extension.lower() in ('.jpg', '.jpeg') else 'image/png')
            result[f'xl/drawings/_rels/mt_footer_{index}.vml.rels'] = ET.tostring(vml_rels, encoding='utf-8')

        ensure_type('vml', 'application/vnd.openxmlformats-officedocument.vmlDrawing')

        rel_path = posixpath.dirname(target_sheet) + '/_rels/' + posixpath.basename(target_sheet) + '.rels'
        target_rels = ET.fromstring(result[rel_path]) if rel_path in result else ET.Element(f'{{{REL}}}Relationships')
        rel_id = 'rIdMTFooter'
        used = {node.get('Id') for node in target_rels}
        while rel_id in used:
            rel_id += '_'
        ET.SubElement(target_rels, f'{{{REL}}}Relationship', {
            'Id': rel_id,
            'Type': DOCREL + '/vmlDrawing',
            'Target': '../drawings/' + posixpath.basename(vml_target)
        })
        result[rel_path] = ET.tostring(target_rels, encoding='utf-8')

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
        result[target_sheet] = ET.tostring(target_xml, encoding='utf-8')

    result['[Content_Types].xml'] = ET.tostring(types, encoding='utf-8')

    handle, temporary = tempfile.mkstemp(suffix='.xlsx', dir=os.path.dirname(os.path.abspath(output)))
    os.close(handle)
    try:
        with zipfile.ZipFile(temporary, 'w', zipfile.ZIP_DEFLATED) as archive:
            for name, data in result.items():
                archive.writestr(name, data)
        os.replace(temporary, output)
    finally:
        if os.path.exists(temporary):
            try:
                os.unlink(temporary)
            except Exception:
                pass


# Alias for general template footer restoration (PMI, MT, etc.)
restore_template_footer_pictures = restore_mt_footer_pictures
