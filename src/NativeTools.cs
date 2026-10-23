using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
namespace PagePilot {
 internal static class NativeTools {
  internal static void Backup(dynamic doc){
   // WordOpenXML includes unsaved edits, unlike copying the file on disk.
   string folder=Path.Combine(Connect.Data,"backups");Directory.CreateDirectory(folder);
   string name=Regex.Replace((string)doc.Name,@"[^\p{L}\p{N}._-]","_");
   File.WriteAllText(Path.Combine(folder,DateTime.UtcNow.ToString("yyyyMMdd-HHmmss-fffffff")+"-"+name+".xml"),(string)doc.WordOpenXML);
  }
  internal static object Ascii(dynamic doc){Backup(doc);doc.Content.Font.NameAscii="Times New Roman";return new{formatted=true,preservedFontSizes=true};}
  internal static object Tables(dynamic doc,float margin){Backup(doc);int count=0;foreach(dynamic table in doc.Tables){table.Rows.Alignment=1;table.TopPadding=margin;table.BottomPadding=margin;table.LeftPadding=margin;table.RightPadding=margin;table.Borders.Enable=1;table.Borders.OutsideLineWidth=4;table.Borders.InsideLineWidth=4;count++;}return new{formattedTables=count};}
  static string S(Dictionary<string,object>a,string key,string fallback="")=>a.ContainsKey(key)?Convert.ToString(a[key]):fallback;
  static int I(Dictionary<string,object>a,string key,int fallback=0)=>a.ContainsKey(key)?Convert.ToInt32(a[key]):fallback;
  static float F(Dictionary<string,object>a,string key,float fallback=0)=>a.ContainsKey(key)?Convert.ToSingle(a[key]):fallback;
  internal static object Execute(Connect host,string name,Dictionary<string,object>a){
   if(host.App.Documents.Count==0)throw new Exception("没有打开的文档。");dynamic doc=host.App.ActiveDocument;
   if(name=="pagepilot_context")return host.Context();
   // Bind mutations to the explicit full name returned by context.
   if(S(a,"document")!=(string)doc.FullName)throw new Exception("活动文档与请求不一致，请重新读取上下文。");
   if(name=="pagepilot_inspect_objects"){
    var fields=new List<object>();foreach(dynamic field in doc.Fields)fields.Add(new{code=(string)field.Code.Text,result=(string)field.Result.Text});
    var controls=new List<object>();foreach(dynamic c in doc.ContentControls)controls.Add(new{id=(string)c.ID,title=(string)c.Title,tag=(string)c.Tag,text=(string)c.Range.Text});
    var bookmarks=new List<string>();foreach(dynamic b in doc.Bookmarks)bookmarks.Add((string)b.Name);
    return new{fields,contentControls=controls,bookmarks,footnotes=(int)doc.Footnotes.Count,endnotes=(int)doc.Endnotes.Count,tablesOfContents=(int)doc.TablesOfContents.Count,sections=(int)doc.Sections.Count};
   }
   if(name=="pagepilot_edit_objects"){
    string operation=S(a,"operation");int paragraph=I(a,"paragraph",1);if(paragraph<1||paragraph>(int)doc.Paragraphs.Count)throw new Exception("段落索引无效");
    var supported=new[]{"add_footnote","add_endnote","add_toc","update_fields","add_content_control","set_content_control"};if(!supported.Contains(operation))throw new Exception("不支持的操作");
    Backup(doc);dynamic range=doc.Paragraphs[paragraph].Range;range.Collapse(1);string text=S(a,"text");
    if(operation=="add_footnote")doc.Footnotes.Add(range,Text:text);
    if(operation=="add_endnote")doc.Endnotes.Add(range,Text:text);
    if(operation=="add_toc")doc.TablesOfContents.Add(range,UseHeadingStyles:true,UpperHeadingLevel:1,LowerHeadingLevel:9,UseOutlineLevels:true);
    if(operation=="update_fields"){int errors=(int)doc.Fields.Update();foreach(dynamic toc in doc.TablesOfContents)toc.Update();return new{updated=true,firstFieldError=errors};}
    if(operation=="add_content_control"){dynamic c=doc.ContentControls.Add(0,range);c.Title=S(a,"title","PagePilot");c.Range.Text=text;return new{created=true,id=(string)c.ID};}
    if(operation=="set_content_control"){bool found=false;foreach(dynamic c in doc.ContentControls)if((string)c.ID==S(a,"control_id")){c.Range.Text=text;found=true;break;}if(!found)throw new Exception("未找到内容控件 ID");}
    return new{applied=true,operation};
   }
   if(name=="pagepilot_normalize_ascii")return Ascii(doc);
   if(name=="pagepilot_normalize_tables")return Tables(doc,F(a,"margin_points",4));
   if(name=="pagepilot_insert_caption"){
    int paragraph=I(a,"paragraph");if(paragraph<1||paragraph>(int)doc.Paragraphs.Count)throw new Exception("无效段落索引。");
    string label=S(a,"label","表"),title=S(a,"title");if(label.Length==0||label.Length>30||title.Length>1000)throw new Exception("题注内容过长或标签为空。");
    Backup(doc);try{var existing=host.App.CaptionLabels[label];}catch{host.App.CaptionLabels.Add(label);}
    doc.Paragraphs[paragraph].Range.InsertCaption(label,title,"",I(a,"position",1),false);return new{inserted=true};
   }
   if(name=="pagepilot_set_navigation"){
    var list=new List<Heading>();foreach(var item in (ArrayList)a["headings"]){var h=(Dictionary<string,object>)item;int i=I(h,"paragraph"),level=I(h,"level");if(i<1||i>(int)doc.Paragraphs.Count||level<1||level>9)throw new Exception("无效标题索引或级别。");dynamic p=doc.Paragraphs[i];list.Add(new Heading{Paragraph=i,Start=(int)p.Range.Start,End=(int)p.Range.End-1,Text=((string)p.Range.Text).Trim(),Level=level});}
    host.ShowPane(1);
    host.SetNavigation(list);return new{headings=list.Count,documentModified=false};
   }
   if(name=="pagepilot_apply_hierarchy"){
    var list=new List<Tuple<int,int>>();foreach(var item in (ArrayList)a["headings"]){var h=(Dictionary<string,object>)item;int i=I(h,"paragraph"),level=I(h,"level");if(i<1||i>(int)doc.Paragraphs.Count||level<1||level>9)throw new Exception("无效标题索引或级别。");list.Add(Tuple.Create(i,level));}
    Backup(doc);dynamic template=null;
    if(S(a,"numbering","false").Equals("true",StringComparison.OrdinalIgnoreCase)){
     try{template=doc.ListTemplates["PagePilot.Hierarchy"];}catch{template=doc.ListTemplates.Add(true,"PagePilot.Hierarchy");}for(int n=1;n<=9;n++){dynamic l=template.ListLevels[n];l.NumberFormat=string.Join(".",Enumerable.Range(1,n).Select(x=>"%"+x));l.NumberStyle=0;l.StartAt=1;l.ResetOnHigher=n-1;l.TrailingCharacter=0;l.NumberPosition=0;l.TextPosition=21f*n;}
    }
    foreach(var h in list){dynamic p=doc.Paragraphs[h.Item1];p.OutlineLevel=h.Item2;if(template!=null)p.Range.ListFormat.ApplyListTemplateWithLevel(template,true,0,1,h.Item2);}
    return new{applied=list.Count};
   }
   throw new Exception("Unknown native tool");
  }
 }
}
