"""Alternate baseline/candidate inventory launches; no Word/profile/network requests."""
import argparse,json,subprocess,time
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('python');p.add_argument('baseline');p.add_argument('candidate');p.add_argument('output');p.add_argument('--samples',type=int,default=30);a=p.parse_args()
samples={'baseline':[],'candidate':[]}
for i in range(a.samples):
    for label in (['baseline','candidate'] if i%2==0 else ['candidate','baseline']):
        script=getattr(a,label);start=time.perf_counter()
        subprocess.run([a.python,script,'--inventory'],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL,check=True,timeout=30)
        samples[label].append((time.perf_counter()-start)*1000)
    Path(a.output).write_text(json.dumps(samples),encoding='utf-8')
    print('completed pair',i+1,flush=True)
