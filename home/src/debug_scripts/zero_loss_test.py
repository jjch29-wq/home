import zipfile
import xml.etree.ElementTree as ET
import os

def patch_sheet_data(template_xlsx, report_xlsx, output_xlsx):
    # Namespaces
    NS = {
        'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    }
    ET.register_namespace('', NS['main'])
    ET.register_namespace('r', NS['r'])

    # 1. Get data from report
    import openpyxl
    wb_data = openpyxl.load_workbook(report_xlsx, data_only=True)
    ws_data = wb_data.worksheets[0]
    cell_values = {}
    for row in ws_data.iter_rows():
        for cell in row:
            if cell.value is not None:
                cell_values[cell.coordinate] = cell.value

    # 2. Patch template
    import shutil
    shutil.copy2(template_xlsx, output_xlsx)
    
    temp_zip = output_xlsx + ".tmp"
    with zipfile.ZipFile(output_xlsx, 'r') as z_in:
        with zipfile.ZipFile(temp_zip, 'w') as z_out:
            for item in z_in.infolist():
                if item.filename == 'xl/worksheets/sheet1.xml':
                    tree = ET.fromstring(z_in.read(item.filename))
                    sheet_data = tree.find('main:sheetData', NS)
                    
                    # Update existing cells or add new ones?
                    # For now, let's just try to find and replace
                    for cell_node in sheet_data.findall('main:row/main:c', NS):
                        addr = cell_node.get('r')
                        if addr in cell_values:
                            val = cell_values[addr]
                            v_node = cell_node.find('main:v', NS)
                            if v_node is None:
                                v_node = ET.SubElement(cell_node, '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
                            
                            # Simple logic: convert to string
                            # If we want shared strings, it's harder. 
                            # Let's use 'inlineStr' or just raw values for now.
                            cell_node.set('t', 'str') 
                            v_node.text = str(val)
                            del cell_values[addr] # Done
                    
                    # If there are leftovers, we might need to add them. 
                    # But for RT, most cells already exist in template.
                    
                    z_out.writestr(item, ET.tostring(tree, encoding='utf-8', xml_declaration=True))
                else:
                    z_out.writestr(item, z_in.read(item.filename))
                    
    os.remove(output_xlsx)
    os.rename(temp_zip, output_xlsx)

# Test run...
