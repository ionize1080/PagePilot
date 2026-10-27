"""Check selected human-reviewed headings/non-headings and unchanged source files."""
import json,re,hashlib,pathlib,sys,statistics
root=pathlib.Path(__file__).resolve().parent;manifest=json.loads((root/'corpus/manifest.json').read_text(encoding='utf-8'))
values=json.loads(pathlib.Path(sys.argv[1]).read_text(encoding='utf-8'));actual={x['file']:x for x in values};checks=0
key=lambda text:re.sub(r'\s','',text)
for item in manifest:
    headings=[key(h['Text']) for h in actual[item['file']]['headings']]
    for expected in item['expectedHeadings']:
        assert any(key(expected) in text for text in headings),(item['file'],'missing',expected);checks+=1
    for excluded in item['excludedSnippets']:
        assert not any(key(excluded) in text for text in headings),(item['file'],'extra',excluded);checks+=1
    file=pathlib.Path(sys.argv[2])/item['file'];assert hashlib.sha256(file.read_bytes()).hexdigest()==item['sha256'];checks+=1
report={'documents':len(values),'paragraphs':sum(x['paragraphs'] for x in values),'selectedAnnotationAndIntegrityChecks':checks,'medianWordMs':statistics.median(x['totalMs'] for x in values),'maxWordMs':max(x['totalMs'] for x in values),'note':'Includes Word opening, metadata reading, analysis and native range checks; not exhaustive heading accuracy.'}
print(json.dumps(report,ensure_ascii=False,indent=2))
if len(sys.argv)>3:pathlib.Path(sys.argv[3]).write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
