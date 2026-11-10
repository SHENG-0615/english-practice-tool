"""Check the actual public source set, including the bundled HTML, before publication."""
import argparse
import hashlib
import json
import pathlib
import re
import subprocess
import sys

ROOT = pathlib.Path(__file__).resolve().parents[1]
parser = argparse.ArgumentParser()
parser.add_argument('--tracked', action='store_true', help='Require and check the exact Git file list')
args = parser.parse_args()
result = subprocess.run(['git', '-C', str(ROOT), 'ls-files', '-z'], capture_output=True)
if result.returncode == 0 and result.stdout:
    names = result.stdout.decode('utf-8').split('\0')[:-1]
elif args.tracked:
    sys.exit('No tracked review files; stage the intended public files first.')
else:
    names = [str(p.relative_to(ROOT)).replace('\\', '/') for p in ROOT.rglob('*') if p.is_file() and not any(part in {'.git', 'dist', 'node_modules', '.local-runtime', '__pycache__', 'test-results', 'playwright-report'} for part in p.relative_to(ROOT).parts)]

errors = []
roots = {'src', 'vendor', 'scripts', 'tests', '.github'}
top = {'.gitignore', '.gitattributes', 'README.md', 'USAGE.md', 'REVIEW.md', 'THIRD_PARTY_NOTICES.md', 'package.json', 'package-lock.json', 'build.py', 'index.html', '开始练习.cmd', 'local-sync.cs', 'local-sync.ps1', 'third_party_integrity.json'}
for name in names:
    p = pathlib.PurePosixPath(name)
    if len(p.parts) == 1 and name not in top or len(p.parts) > 1 and p.parts[0] not in roots:
        errors.append('Unapproved file: ' + name)
    if p.suffix.lower() in {'.xlsx', '.xls', '.xlsm', '.xlsb', '.csv', '.tsv', '.zip', '.bak', '.tmp', '.log', '.png', '.jpg', '.jpeg', '.webp'}:
        errors.append('Private-data or generated artifact type: ' + name)
    if p.suffix.lower() == '.txt' and name not in {'vendor/LICENSE-SheetJS.txt', 'vendor/LICENSE-JSZip.txt'}:
        errors.append('Unapproved text export: ' + name)
    if any(part in {'.local-runtime', '同步备份', 'private', 'local-data', 'browser-profile'} for part in p.parts):
        errors.append('Local runtime/data path: ' + name)
    if p.suffix.lower() == '.json' and name not in {'src/seed.json', 'src/vocab-seed.json', 'package.json', 'package-lock.json', 'third_party_integrity.json'}:
        errors.append('Unapproved JSON export: ' + name)
    if name.startswith('vendor/') or name == 'scripts/check_public_release.py':
        continue
    data = (ROOT / name).read_text(encoding='utf-8-sig')
    if re.search(r'(?:[A-Za-z]:[\\/]+Users[\\/]+|/Users/|/home/)[^\s\"<>]+', data):
        errors.append('Hardcoded personal home path: ' + name)
    if re.search(r'(?:gh[pousr]_[A-Za-z0-9]{24,}|github_pat_[A-Za-z0-9_]{24,}|-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----)', data):
        errors.append('Credential material: ' + name)

expected_seeds = {
    'src/seed.json': {'version': 1, 'title': '听写与词汇练习', 'entries': [], 'ordered': []},
    'src/vocab-seed.json': {'version': 1, 'entries': [], 'syncCompatible': False}
}
for name, expected in expected_seeds.items():
    if json.loads((ROOT / name).read_text(encoding='utf-8')) != expected:
        errors.append('Nonempty or unexpected seed fields: ' + name)
page = (ROOT / 'index.html').read_text(encoding='utf-8')
for tag, name in [('seedData', 'src/seed.json'), ('vocabularyData', 'src/vocab-seed.json')]:
    match = re.search(r'<script id="' + tag + r'" type="application/json">(.*?)</script>', page, re.S)
    if not match or json.loads(match.group(1)) != expected_seeds[name]:
        errors.append('Bundled page has unexpected embedded data: ' + tag)
if '/* LOCAL_CONFIG_START */null/* LOCAL_CONFIG_END */' not in page:
    errors.append('Bundled page includes a local connection configuration')
if re.search(r'<script[^>]+src\s*=', page, re.I):
    errors.append('Unexpected external script reference')
integrity = json.loads((ROOT / 'third_party_integrity.json').read_text(encoding='utf-8'))
for name, expected in integrity.items():
    actual = hashlib.sha256((ROOT / name).read_bytes()).hexdigest()
    if actual != expected:
        errors.append('Third-party integrity mismatch: ' + name)
if errors:
    print('\n'.join(errors), file=sys.stderr)
    sys.exit(1)
print(json.dumps({'status': 'passed', 'files': len(names), 'embedded_word_count': 0, 'embedded_attempt_count': 0, 'xlsx_files': 0, 'local_connection_tokens': 0, 'tracked_check': bool(result.stdout)}, ensure_ascii=False))
