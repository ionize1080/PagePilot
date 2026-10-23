"""Real Word + stdio regression matrix; confirm auto approval for this fixture first."""
import json
from pathlib import Path
import win32com.client
from mcp_client import Client,current_session

root=Path(__file__).resolve().parents[1]
app=win32com.client.GetActiveObject("Word.Application")
doc=app.ActiveDocument
assert doc.Name=="PagePilot-v020-QA.docx", "Only synthetic QA document is allowed"
# Reset only our disposable fixture; never touch a user document.
doc.Content.Text="第一章 总则\r第一节 范围\r收入从100增长到120，增长10%。\r第二章 结论\rQA text ABC 2024.\r"
doc.TrackRevisions=False
path=str(doc.FullName)
c=Client(root/"dist",current_session())
results=[]
def call(name,**args):
    r=c.call(name,args)
    assert not r.get("isError"),(name,r)
    text="\n".join(p.get("text","") for p in r.get("content",[]))
    try:
        payload=json.loads(text)
        assert not (isinstance(payload,dict) and (payload.get("error") or payload.get("success") is False)),(name,payload)
    except json.JSONDecodeError:
        assert not text.lower().startswith(("error","failed","cannot")),(name,text)
    results.append({"tool":name,"result":r})
    print("PASS",name,flush=True)
    return r

try:
    call("word_live_get_info")
    call("word_live_find_text",search_text="收入")
    call("word_live_format_text",start_paragraph=1,end_paragraph=1,bold=True,font_size=16)
    assert doc.Paragraphs(1).Range.Font.Bold==-1
    call("word_live_set_paragraph_spacing",paragraph_index=2,space_after_pt=6,line_spacing_rule="exactly",line_spacing=18,alignment="left")
    assert abs(doc.Paragraphs(2).Format.SpaceAfter-6)<0.01
    call("word_live_set_page_layout",margin_top_inches=1,margin_bottom_inches=1,margin_left_inches=1,margin_right_inches=1)
    assert abs(doc.PageSetup.TopMargin-72)<0.1
    call("word_live_add_table",rows=2,cols=2,data=[["项目","金额"],["收入","120"]],style="",position="end")
    assert doc.Tables.Count==1
    call("pagepilot_normalize_tables",document=path,margin_points=4)
    assert abs(doc.Tables(1).LeftPadding-4)<0.01
    call("word_live_format_table",table_index=1,table_alignment="center")
    call("word_live_modify_table",table_index=1,operation="set_cell",row=2,col=2,text="125")
    assert "125" in doc.Tables(1).Cell(2,2).Range.Text
    call("word_live_add_comment",paragraph_index=3,text="核稿：100 增至 120 的增幅应为 20%。",author="PagePilot QA")
    assert doc.Comments.Count>=1
    call("word_live_get_comments")
    call("word_live_add_header_footer",header_text="PagePilot QA",footer_text="Integration test")
    call("word_live_add_page_numbers",include_total=True,prefix="Page ")
    call("pagepilot_apply_hierarchy",document=path,headings=[{"paragraph":1,"level":1},{"paragraph":2,"level":2}],numbering=True)
    assert doc.Paragraphs(1).OutlineLevel==1 and doc.Paragraphs(2).OutlineLevel==2
    assert doc.Paragraphs(2).Range.ListFormat.ListType!=0
    call("word_live_insert_text",text="\rTracked QA sentence.",position="end",track_changes=True)
    assert doc.Revisions.Count>0
    call("word_live_list_revisions")
    call("word_live_get_paragraph_format",start_paragraph=2)
    call("word_live_diagnose_layout")
    pdf=root.parent/"PagePilot-v020-QA.pdf"
    call("convert_to_pdf",filename=path,output_filename=str(pdf))
    assert pdf.exists() and pdf.stat().st_size>1000
    assert app.Documents.Count==1 and app.ActiveDocument.Name=="PagePilot-v020-QA.docx"
    call("word_live_save")
    print("Verified actual Word state for formatting, tables, hierarchy, revisions and PDF export",flush=True)
finally:
    (root.parent/"operations-v020.json").write_text(json.dumps(results,ensure_ascii=False,indent=2),encoding="utf-8")
    c.close()
