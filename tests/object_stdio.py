"""End-to-end read + manually approved edit against StartupRegression's own fixture."""
import sys,json
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parent))
from mcp_client import Client
root=Path(__file__).resolve().parents[1]
c=Client(root/'dist',json.loads(Path(sys.argv[1]).read_text(encoding='utf-8-sig')))
def call(name,**args):
 r=c.call(name,args)
 assert not r.get('isError'),r
 result=json.loads('\n'.join(x.get('text','') for x in r.get('content',[])))
 assert not result.get('error'),result
 return result
try:
 caps=call('pagepilot_capabilities');assert len(caps['tools'])==132
 context=call('pagepilot_context',start=1,count=1)
 d=context['document'];assert '标题' in d['paragraphs'][0]['text'], 'Only synthetic StartupRegression fixture'
 document=d['path']
 schema=call('pagepilot_objects',document=document,kind='schema');assert 'border' in schema['propertySets']
 objects=call('pagepilot_objects',document=document,kind='range',scope={'start':0,'end':0})
 handle=objects['objects'][0]['handle']
 print('READS PASSED; waiting for UI approval for synthetic insertion',flush=True)
 result=call('pagepilot_object_edit',document=document,handle=handle,operation='text',parameters={'text':'MCP QA '})
 assert result['status']=='completed',result
 readback=call('pagepilot_objects',document=document,kind='range',scope={'start':0,'end':7})
 assert readback['objects'][0]['properties']['text']=='MCP QA '
 wrong=c.call('pagepilot_objects',{'document':'Z:\\does-not-exist\\wrong.docx','kind':'section'})
 assert wrong.get('isError'),wrong
 print('PASS stdio discovery, schema, manual approval, write/readback, wrong-document rejection',flush=True)
finally:c.close()
