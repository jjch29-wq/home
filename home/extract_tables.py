from bs4 import BeautifulSoup
import sys

# Change default encoding to utf-8 for print
sys.stdout.reconfigure(encoding='utf-8')

html = open(r'c:\Users\-\OneDrive\바탕 화면\home\Na-aba\home\hwp_output\index.xhtml', encoding='utf-8').read()
soup = BeautifulSoup(html, 'html.parser')
tables = soup.find_all('table')
print(f'Found {len(tables)} tables.')
for i, t in enumerate(tables):
    print(f'\n--- Table {i+1} ---')
    for row in t.find_all('tr'):
        print(' | '.join([cell.get_text(strip=True) for cell in row.find_all(['th','td'])]))
