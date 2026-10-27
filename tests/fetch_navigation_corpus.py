"""Fetch the pinned public DOCX corpus; do not redistribute third-party documents."""
import argparse, concurrent.futures, hashlib, json, pathlib, re, urllib.request, urllib.parse, zipfile, io
p=argparse.ArgumentParser();p.add_argument('--output',default='.temp/corpus/files');p.add_argument('--proxy');a=p.parse_args()
manifest=json.loads((pathlib.Path(__file__).parent/'corpus/manifest.json').read_text(encoding='utf-8'));target=pathlib.Path(a.output);target.mkdir(parents=True,exist_ok=True)
def fetch(item):
    assert re.fullmatch(r'[0-9a-f]{16}\.docx',item['file']) and urllib.parse.urlsplit(item['url']).scheme=='https'
    file=target/item['file']
    if file.exists() and hashlib.sha256(file.read_bytes()).hexdigest()==item['sha256']:return item['file']+' cached'
    opener=urllib.request.build_opener(urllib.request.ProxyHandler({'https':a.proxy}) if a.proxy else urllib.request.ProxyHandler())
    with opener.open(item['url'],timeout=60) as r:data=r.read(16000000)
    if hashlib.sha256(data).hexdigest()!=item['sha256']:raise ValueError('Source changed or download incomplete: '+item['url'])
    with zipfile.ZipFile(io.BytesIO(data)) as archive:assert 'word/document.xml' in archive.namelist()
    file.write_bytes(data);return item['file']+' verified'
with concurrent.futures.ThreadPoolExecutor(max_workers=4) as pool:
    for result in pool.map(fetch,manifest):print(result)
