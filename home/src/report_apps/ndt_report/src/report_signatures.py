"""Append signatures without resaving or replacing existing Excel drawings."""
import io
import os
import posixpath
import tempfile
import zipfile
from lxml import etree as ET
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from mt_footer import resolve, MAIN, REL, DOCREL, CONTENT


def append_rt_signatures(output, signatures):
    """signatures contains (PNG stream, anchor, maximum width, height)."""
    temporary_book = Workbook()
    for stream, anchor, max_width, max_height in signatures:
        image = Image(stream)
        ratio = min(max_width / image.width, max_height / image.height)
        image.width, image.height = image.width * ratio, image.height * ratio
        temporary_book.active.add_image(image, anchor)
    buffer = io.BytesIO()
    temporary_book.save(buffer)
    with zipfile.ZipFile(buffer) as archive:
        added = {name: archive.read(name) for name in archive.namelist()}
    with zipfile.ZipFile(output) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    sheet_path = 'xl/worksheets/sheet1.xml'
    sheet = ET.fromstring(parts[sheet_path])
    sheet_rels_path = 'xl/worksheets/_rels/sheet1.xml.rels'
    sheet_rels = ET.fromstring(parts[sheet_rels_path]) if sheet_rels_path in parts else ET.Element(f'{{{REL}}}Relationships', nsmap={None: REL})
    drawing_link = next((link for link in sheet_rels if link.get('Type') == DOCREL + '/drawing'), None)
    if drawing_link is None:
        identifier = 'rIdSignatures'
        while any(link.get('Id') == identifier for link in sheet_rels):
            identifier += '_'
        drawing_link = ET.SubElement(sheet_rels, f'{{{REL}}}Relationship', Id=identifier,
                                    Type=DOCREL + '/drawing', Target='../drawings/rt_signatures.xml')
        drawing = ET.SubElement(sheet, f'{{{MAIN}}}drawing')
        drawing.set(f'{{{DOCREL}}}id', identifier)
    drawing_path = resolve(sheet_path, drawing_link.get('Target'))
    drawing_rels_path = posixpath.dirname(drawing_path) + '/_rels/' + posixpath.basename(drawing_path) + '.rels'
    xdr = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
    drawing = ET.fromstring(parts[drawing_path]) if drawing_path in parts else ET.Element(f'{{{xdr}}}wsDr', nsmap={None: xdr})
    relationships = ET.fromstring(parts[drawing_rels_path]) if drawing_rels_path in parts else ET.Element(f'{{{REL}}}Relationships', nsmap={None: REL})
    new_drawing = ET.fromstring(added['xl/drawings/drawing1.xml'])
    new_relationships = ET.fromstring(added['xl/drawings/_rels/drawing1.xml.rels'])
    next_id = max([int(node.get('id', '0')) for node in drawing.iter() if ET.QName(node).localname == 'cNvPr'] or [0]) + 1
    for index, anchor in enumerate(new_drawing, 1):
        for node in anchor.iter():
            if ET.QName(node).localname == 'cNvPr':
                node.set('id', str(next_id))
                next_id += 1
            embed = node.get(f'{{{DOCREL}}}embed')
            if embed:
                source_link = next(link for link in new_relationships if link.get('Id') == embed)
                media_source = resolve('xl/drawings/drawing1.xml', source_link.get('Target'))
                media_path = f'xl/media/rt_signature_{index}.png'
                parts[media_path] = added[media_source]
                identifier = f'rIdRTSignature{index}'
                while any(link.get('Id') == identifier for link in relationships):
                    identifier += '_'
                ET.SubElement(relationships, f'{{{REL}}}Relationship', Id=identifier,
                              Type=DOCREL + '/image', Target=posixpath.relpath(media_path, posixpath.dirname(drawing_path)))
                node.set(f'{{{DOCREL}}}embed', identifier)
        drawing.append(anchor)
    parts[drawing_path] = ET.tostring(drawing)
    parts[drawing_rels_path] = ET.tostring(relationships)
    parts[sheet_rels_path] = ET.tostring(sheet_rels)
    parts[sheet_path] = ET.tostring(sheet)
    types = ET.fromstring(parts['[Content_Types].xml'])
    if not any(node.get('Extension') == 'png' for node in types):
        ET.SubElement(types, f'{{{CONTENT}}}Default', Extension='png', ContentType='image/png')
    if not any(node.get('PartName') == '/' + drawing_path for node in types):
        ET.SubElement(types, f'{{{CONTENT}}}Override', PartName='/' + drawing_path,
                      ContentType='application/vnd.openxmlformats-officedocument.drawing+xml')
    parts['[Content_Types].xml'] = ET.tostring(types)
    handle, temp_path = tempfile.mkstemp(suffix='.xlsx', dir=os.path.dirname(os.path.abspath(output)))
    os.close(handle)
    try:
        with zipfile.ZipFile(temp_path, 'w', zipfile.ZIP_DEFLATED) as archive:
            for name, value in parts.items():
                archive.writestr(name, value)
        os.replace(temp_path, output)
    finally:
        if os.path.exists(temp_path):
            os.unlink(temp_path)
