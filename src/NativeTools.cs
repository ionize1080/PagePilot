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
  internal static object Ascii(dynamic doc){Backup(doc);string text=(string)doc.Content.Text;int count=0;foreach(Match m in Regex.Matches(text,@"[\u0021-\u007e]+")){dynamic r=doc.Range((int)doc.Content.Start+m.Index,(int)doc.Content.Start+m.Index+m.Length);r.Font.NameAscii="Times New Roman";r.Font.NameOther="Times New Roman";r.Font.Size=10.5f;count++;}return new{formattedRuns=count};}
  internal static object Tables(dynamic doc,float margin){Backup(doc);int count=0;foreach(dynamic table in doc.Tables){table.Rows.Alignment=1;table.TopPadding=margin;table.BottomPadding=margin;table.LeftPadding=margin;table.RightPadding=margin;table.Borders.Enable=1;table.Borders.OutsideLineWidth=4;table.Borders.InsideLineWidth=4;count++;}return new{formattedTables=count};}
  static string S(Dictionary<string,object>a,string key,string fallback="")=>a.ContainsKey(key)?Convert.ToString(a[key]):fallback;
  static int I(Dictionary<string,object>a,string key,int fallback=0)=>a.ContainsKey(key)?Convert.ToInt32(a[key]):fallback;
  static float F(Dictionary<string,object>a,string key,float fallback=0)=>a.ContainsKey(key)?Convert.ToSingle(a[key]):fallback;
  internal static object Execute(Connect host,string name,Dictionary<string,object>a){
   if(host.App.Documents.Count==0)throw new Exception("没有打开的文档。");dynamic doc=host.App.ActiveDocument;
   if(name=="pagepilot_context")return host.Context();
   // Bind mutations to the explicit full name returned by context.
   if(S(a,"document")!=(string)doc.FullName)throw new Exception("活动文档与请求不一致，请重新读取上下文。");
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
     template=doc.ListTemplates.Add(true);for(int n=1;n<=9;n++){dynamic l=template.ListLevels[n];l.NumberFormat=string.Join(".",Enumerable.Range(1,n).Select(x=>"%"+x));l.NumberStyle=0;l.StartAt=1;l.ResetOnHigher=n-1;l.TrailingCharacter=0;l.NumberPosition=0;l.TextPosition=21f*n;}
    }
    foreach(var h in list){dynamic p=doc.Paragraphs[h.Item1];p.OutlineLevel=h.Item2;if(template!=null)p.Range.ListFormat.ApplyListTemplateWithLevel(template,true,0,1,h.Item2);}
    return new{applied=list.Count};
   }
   throw new Exception("Unknown native tool");
  }
 }
}
