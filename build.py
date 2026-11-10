import html
import json
import pathlib
import re

ROOT = pathlib.Path(__file__).resolve().parent
OUTPUTS = ROOT

# Public builds must never embed a local dataset or any practice history.
for name in ['seed.json', 'vocab-seed.json']:
    data = json.loads((ROOT / 'src' / name).read_text(encoding='utf-8'))
    if data.get('entries') or data.get('ordered'):
        raise RuntimeError('Public seed files must be empty. Use the browser to import a local word list.')

def inline_script(value):
    return re.sub(r'</script', r'<\\/script', value, flags=re.I)

page = (ROOT / 'src/template.html').read_text(encoding='utf-8')
parts = {
    '/* APP_CSS */': (ROOT / 'src/style.css').read_text(encoding='utf-8'),
    '<!-- PUBLIC_IMPORT_AREA -->': (ROOT / 'src/public-import.html').read_text(encoding='utf-8'),
    '/* PUBLIC_DATA_JS */': inline_script((ROOT / 'src/public-data.js').read_text(encoding='utf-8')),
    '/* VOCAB_CSS */': (ROOT / 'src/vocab-style.css').read_text(encoding='utf-8'),
    '<!-- VOCAB_AREA -->': (ROOT / 'src/vocab-template.html').read_text(encoding='utf-8'),
    '/* VOCAB_JSON */': json.dumps(json.loads((ROOT / 'src/vocab-seed.json').read_text(encoding='utf-8')), ensure_ascii=False, separators=(',', ':')).replace('<', '\\u003c'),
    '/* VOCAB_CORE_JS */': inline_script((ROOT / 'src/vocab-core.js').read_text(encoding='utf-8')),
    '/* VOCAB_WORKBOOK_JS */': inline_script((ROOT / 'src/vocab-workbook.js').read_text(encoding='utf-8')),
    '/* VOCAB_BANK_JS */': inline_script((ROOT / 'src/vocab-bank.js').read_text(encoding='utf-8')),
    '/* XLSX_IMPORTANCE_JS */': inline_script((ROOT / 'src/xlsx-importance.js').read_text(encoding='utf-8')),
    '/* VOCAB_APP_JS */': inline_script((ROOT / 'src/vocab-app.js').read_text(encoding='utf-8')),
    '/* SEED_JSON */': json.dumps(json.loads((ROOT / 'src/seed.json').read_text(encoding='utf-8')), ensure_ascii=False, separators=(',', ':')).replace('<', '\\u003c'),
    '/* XLSX_LIBRARY */': inline_script((ROOT / 'vendor/xlsx.full.min.js').read_text(encoding='utf-8')),
    '/* CORE_JS */': inline_script((ROOT / 'src/core.js').read_text(encoding='utf-8')),
    '/* JSZIP_LIBRARY */': inline_script((ROOT / 'vendor/jszip.min.js').read_text(encoding='utf-8')),
    '/* XLSX_LOG_JS */': inline_script((ROOT / 'src/xlsx-log.js').read_text(encoding='utf-8')),
    '/* XLSX_CHAPTERS_JS */': inline_script((ROOT / 'src/xlsx-chapters.js').read_text(encoding='utf-8')),
    '/* XLSX_SYNC_JS */': inline_script((ROOT / 'src/xlsx-sync.js').read_text(encoding='utf-8')),
    '/* LOCAL_FILES_JS */': inline_script((ROOT / 'src/local-files.js').read_text(encoding='utf-8')),
    '/* APP_JS */': inline_script((ROOT / 'src/app.js').read_text(encoding='utf-8')),
    '<!-- SHEETJS_LICENSE -->': html.escape((ROOT / 'vendor/LICENSE-SheetJS.txt').read_text(encoding='utf-8')),
    '<!-- JSZIP_LICENSE -->': html.escape((ROOT / 'vendor/LICENSE-JSZip.txt').read_text(encoding='utf-8')),
}
for marker, content in parts.items():
    assert marker in page, marker
    page = page.replace(marker, content)
assert not re.search(r'/\* (APP_CSS|SEED_JSON|XLSX_LIBRARY|CORE_JS|JSZIP_LIBRARY|XLSX_LOG_JS|XLSX_CHAPTERS_JS|XLSX_SYNC_JS|LOCAL_FILES_JS|APP_JS|VOCAB_\w+) \*/', page)
(ROOT / 'dist').mkdir(exist_ok=True)
OUTPUTS.mkdir(exist_ok=True)
for target in [ROOT / 'dist/index.html', OUTPUTS / 'index.html']:
    target.write_text(page, encoding='utf-8', newline='\n')
print(json.dumps({'file': str(OUTPUTS / 'index.html'), 'bytes': len(page.encode('utf-8')), 'seedEntries': len(json.loads((ROOT / 'src/seed.json').read_text(encoding='utf-8'))['entries'])}, ensure_ascii=False))
