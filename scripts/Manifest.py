"""Generate integrity inventory from release runtime files only (no profiles/caches)."""
import hashlib
from pathlib import Path
import argparse
parser=argparse.ArgumentParser()
parser.add_argument('--root',default=str(Path(__file__).resolve().parents[1]/'dist'))
args=parser.parse_args()
root=Path(args.root).resolve()
names=['PagePilot.dll','PagePilot.Manager.exe','Python','Browser','Chatbox','bridge','vendor']
files=[]
for name in names:
    item=root/name
    if not item.exists():raise FileNotFoundError(item)
    files.extend([item] if item.is_file() else [p for p in item.rglob('*') if p.is_file() and p.name.lower()!='debug.log' and not {'__pycache__','.git','.pytest_cache'}&set(p.relative_to(item).parts)])
rows=[]
for p in sorted(files):
    with p.open('rb') as f: digest=hashlib.file_digest(f,'sha256').hexdigest()
    rows.append(digest+'  '+p.relative_to(root).as_posix())
(root/'manifest.sha256').write_text('\n'.join(rows)+'\n',encoding='utf-8')
print('Integrity inventory:',len(rows),'runtime files')
