"""Build a self-contained Apps Script upgrade; only Code.gs must be replaced."""
import hashlib
import json
from pathlib import Path
import zipfile

ROOT=Path(__file__).resolve().parents[1]
SOURCE=ROOT/'home/src/site_apps/central/google_apps_script'
OUTPUT=ROOT/'outputs/central-mobile-gas'


def build():
    OUTPUT.mkdir(parents=True,exist_ok=True)
    script='\n\n'.join((SOURCE/name).read_text(encoding='utf-8')
                       for name in ('Code.gs','Reference.gs','Workspace.gs'))
    html=(SOURCE/'index.html').read_text(encoding='utf-8')
    html=html.replace("<?!= include('MobileShell'); ?>",(SOURCE/'MobileShell.html').read_text(encoding='utf-8'))
    html=html.replace("<?!= include('MethodFilters'); ?>",(SOURCE/'MethodFilters.html').read_text(encoding='utf-8'))
    assert "include('MobileShell')" not in html
    assert script.count('function saveRecord(req)')==1
    assert script.count('function updateRecord(req)')==1
    assert script.count('function deleteRecord(req)')==1
    script=script.replace("HtmlService.createTemplateFromFile('index')",
                          'HtmlService.createTemplate(CENTRAL_WEBAPP_HTML)')
    script+='\n\n// Embedded mobile page: no separate index/MobileShell file is required.\nvar CENTRAL_WEBAPP_HTML = '+json.dumps(html,ensure_ascii=False)+';\n'
    assert "createTemplateFromFile('index')" not in script
    (OUTPUT/'Code.gs').write_text(script,encoding='utf-8')
    (OUTPUT/'index.html').write_text(html,encoding='utf-8')
    guide=(SOURCE/'MOBILE_UPGRADE.md').read_text(encoding='utf-8')
    (OUTPUT/'APPLY_GUIDE.md').write_text(guide,encoding='utf-8')
    manifest={name:hashlib.sha256((OUTPUT/name).read_bytes()).hexdigest()
              for name in ('Code.gs','index.html','APPLY_GUIDE.md')}
    (OUTPUT/'manifest.json').write_text(json.dumps(manifest,indent=2),encoding='utf-8')
    archive=OUTPUT/'central-mobile-gas.zip'
    with zipfile.ZipFile(archive,'w',zipfile.ZIP_DEFLATED) as zipped:
        for name in (*manifest,'manifest.json'):
            zipped.write(OUTPUT/name,name)
    print(f'Package prepared: {archive}')
    return OUTPUT


if __name__=='__main__':build()
