import codecs
import re

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\monthly_report_exporter.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    text = f.read()

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

text = re.sub(r'    def generate\(self, output_path\):.*', new_generate, text, flags=re.DOTALL)

with codecs.open(file_path, 'w', 'utf-8') as f:
    f.write(text)
print('Patched generate()!')
