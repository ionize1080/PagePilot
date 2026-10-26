"""Whitelist-only package assembly. No user profiles, credentials or caches."""
import argparse
import hashlib
from pathlib import Path
import zipfile

parser=argparse.ArgumentParser()
parser.add_argument('--output',required=True)
parser.add_argument('--chatbox-source',required=True)
parser.add_argument('--version',default='0.4.1')
args=parser.parse_args()
repo=Path(__file__).resolve().parents[1]
out=Path(args.output).resolve()
out.mkdir(parents=True,exist_ok=True)
skip={'__pycache__','.git','node_modules','.pytest_cache','bin','obj','data'}
def put_tree(z,folder,prefix):
    for p in sorted(folder.rglob('*')):
        if p.is_file() and not any(x in skip for x in p.relative_to(folder).parts):
            z.write(p,str(Path(prefix)/p.relative_to(folder)))
version=args.version
if not __import__('re').fullmatch(r'\d+\.\d+\.\d+(?:-[a-z0-9.]+)?',version):raise ValueError('Invalid release version')
portable=out/f'PagePilot-{version}-portable-win-x64.zip'
with zipfile.ZipFile(portable,'w',zipfile.ZIP_DEFLATED,compresslevel=6) as z:
    for name in ['PagePilot.dll','PagePilot.Manager.exe','manifest.sha256','Python','Browser','Chatbox','bridge','vendor']:
        p=repo/'dist'/name
        if not p.exists():raise FileNotFoundError(p)
        # Runtime node_modules are required; filter only caches here.
        if p.is_dir():
            for f in sorted(p.rglob('*')):
                if f.is_file() and not {'__pycache__','.git','.pytest_cache'}&set(f.relative_to(p).parts):
                    z.write(f,str(Path('PagePilot')/name/f.relative_to(p)))
        else:z.write(p,'PagePilot/'+name)
    for name in ['scripts','docs']:
        put_tree(z,repo/name,'PagePilot/'+name)
    for name in ['README.md','LICENSE','THIRD-PARTY-NOTICES.md','Enable.cmd','Disable.cmd','requirements-lock.txt','runtime.example.json']:
        z.write(repo/name,'PagePilot/'+name)
source=out/f'PagePilot-{version}-source.zip'
with zipfile.ZipFile(source,'w',zipfile.ZIP_DEFLATED,compresslevel=6) as z:
    for name in ['manager','src','shared','bridge','chatbox','vendor','scripts','tests','docs']:
        put_tree(z,repo/name,'PagePilot/'+name)
    for p in repo.iterdir():
        if p.is_file() and p.name!='.gitignore':z.write(p,'PagePilot/'+p.name)
    z.write(args.chatbox_source,'PagePilot/upstream-source/chatbox-v1.20.3.tar.gz')
hashes=[]
for p in [portable,source]:
    digest=hashlib.file_digest(p.open('rb'),'sha256').hexdigest()
    hashes.append(digest+'  '+p.name)
    print(p.name,p.stat().st_size,digest)
(out/f'PagePilot-{version}-SHA256SUMS.txt').write_text('\n'.join(hashes)+'\n',encoding='ascii')
