import codecs
import re

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\monthly_report_exporter.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    text = f.read()

old_generate = r'''    def generate\(self, output_path\):
        self\._aggregate_data\(\)
        
        excel = None
        wb = None
        try:
            excel = win32com\.client\.Dispatch\("Excel\.Application"\)
            excel\.Visible = False
            excel\.DisplayAlerts = False
            
            # Ensure path is absolute
            abs_template = os\.path\.abspath\(self\.template_path\)
            abs_output = os\.path\.abspath\(output_path\)
            
            wb = excel\.Workbooks\.Open\(abs_template\)
            
            # Excel constants
            xlPart = 2
            
            # Search and replace in all worksheets
            for ws in wb\.Worksheets:
                # To maximize performance and reliability, we use Cells\.Replace
                for tag, value in self\.replacements\.items\(\):
                    # Replace requires strings
                    val_str = str\(value\)
                    # Use xlPart so that if a cell is ".*?", it replaces properly
                    ws\.Cells\.Replace\(What=tag, Replacement=val_str, LookAt=xlPart\)
                    
            # Save as format 51 \(xlsx\)
            abs_output = abs_output\.replace\("/", "\\\\"\)
            wb\.SaveAs\(abs_output, 51\)
            
        except Exception as e:
            print\(f"Error generating report: \{e\}"\)
            raise
        finally:
            if wb:
                wb\.Close\(SaveChanges=False\)
            if excel:
                excel\.Quit\(\)'''

new_generate = r'''    def generate(self, output_path):
        self._aggregate_data()
        
        try:
            import openpyxl
            wb = openpyxl.load_workbook(self.template_path)
            
            for ws in wb.worksheets:
                for row in ws.iter_rows():
                    for cell in row:
                        if cell.value and isinstance(cell.value, str):
                            new_val = cell.value
                            for tag, value in self.replacements.items():
                                new_val = new_val.replace(tag, str(value))
                            if new_val != cell.value:
                                cell.value = new_val
                                
            wb.save(output_path)
            wb.close()
            
        except Exception as e:
            print(f"Error generating report: {e}")
            raise'''

if re.search(old_generate, text):
    text = re.sub(old_generate, new_generate, text)
    with codecs.open(file_path, 'w', 'utf-8') as f:
        f.write(text)
    print('Patched generate() to use openpyxl!')
else:
    print('Failed to match generate()!')
