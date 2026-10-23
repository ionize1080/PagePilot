"""Actual Word readback for new live object tools; disposable fixture only."""
import json
from pathlib import Path
import win32com.client
from mcp_client import Client,current_session
root=Path(__file__).resolve().parents[1]
app=win32com.client.GetActiveObject('Word.Application');doc=app.ActiveDocument
assert doc.Name=='PagePilot-v020-QA.docx'
document=str(doc.FullName);c=Client(root/'dist',current_session());report=[]
def call(name,**args):
    result=c.call(name,args);assert not result.get('isError'),result
    text='\n'.join(x.get('text','') for x in result.get('content',[]));payload=json.loads(text)
    assert not payload.get('error'),payload
    report.append(dict(tool=name,arguments=args,result=payload));print('PASS',name,args.get('operation',''),flush=True)
    return payload
try:
    caps=call('pagepilot_capabilities');assert len(caps['tools'])==129
    before=call('pagepilot_inspect_objects',document=document)
    for op,collection in [('add_footnote','Footnotes'),('add_endnote','Endnotes')]:
        count=getattr(doc,collection).Count
        call('pagepilot_edit_objects',document=document,operation=op,paragraph=2,text='PagePilot QA annotation')
        assert getattr(doc,collection).Count==count+1
    control=call('pagepilot_edit_objects',document=document,operation='add_content_control',paragraph=3,text='Initial QA value',title='QA content')
    call('pagepilot_edit_objects',document=document,operation='set_content_control',control_id=control['id'],text='Updated QA value')
    inspect=call('pagepilot_inspect_objects',document=document)
    assert any(x['id']==control['id'] and x['text']=='Updated QA value' for x in inspect['contentControls'])
    doc.Paragraphs(1).OutlineLevel=1
    call('pagepilot_edit_objects',document=document,operation='add_toc',paragraph=doc.Paragraphs.Count)
    assert doc.TablesOfContents.Count==1
    call('pagepilot_edit_objects',document=document,operation='update_fields')
    assert '第一章' in doc.TablesOfContents(1).Range.Text
    result=c.call('pagepilot_edit_objects',dict(document=document,operation='invalid'))
    assert result.get('isError');print('PASS invalid object operation rejected')
    (root.parent/'objects-v020.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
finally:c.close()
