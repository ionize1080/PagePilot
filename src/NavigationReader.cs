using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
namespace PagePilot {
 // Runs on Word's STA, yielding between short batches. Only immutable values leave it.
 public static class NavigationReader {
  static string Clean(string text)=>System.Text.RegularExpressions.Regex.Replace((text??"").Trim('\r','\a',' ','\t'),@"^(?:[/\u0001\u0002]*[\f\v])+\s*","").Trim();
  static string Key(string text)=>System.Text.RegularExpressions.Regex.Replace(Clean(text),@"[\s\a\u0001\u0002\u0013\u0014\u0015]","");
  public static void Resolve(object document,IEnumerable<Heading> headings){dynamic doc=document;foreach(var h in headings){dynamic r=doc.Paragraphs[h.Paragraph].Range;h.Start=(int)r.Start;h.End=(int)r.End-1;h.Text=Clean((string)r.Text);try{if((int)r.ListFormat.ListType!=0)h.Text=(string)r.ListFormat.ListString+" "+h.Text;}catch(Exception ex) when(ex is System.Runtime.InteropServices.COMException||ex is Microsoft.CSharp.RuntimeBinder.RuntimeBinderException){}}}
  public static async Task<List<ParagraphInfo>> Read(object document,Action<int,int> progress,CancellationToken cancel){
   dynamic doc=document;var rows=new List<ParagraphInfo>();int count=(int)doc.Paragraphs.Count,index=0,samples=0;
   var tocRanges=new List<Tuple<int,int>>();foreach(dynamic toc in doc.TablesOfContents)tocRanges.Add(Tuple.Create((int)toc.Range.Start,(int)toc.Range.End));
   var watch=System.Diagnostics.Stopwatch.StartNew();
   List<ParagraphInfo> fast=null;int metadataIndex=0;
   try{fast=NavigationSnapshot.ReadFlat((string)doc.Content.WordOpenXML);}catch{fast=null;}
   var keys=fast?.Select(p=>Key(p.Text)).ToArray();
   foreach(dynamic p in doc.Paragraphs){
    cancel.ThrowIfCancellationRequested();dynamic r=p.Range;string text=Clean((string)r.Text);
    int start=(int)r.Start,end=(int)r.End-1;ParagraphInfo row=null;
    // Word may expose extra cell markers/field paragraphs. Align by displayed text,
    // never by an unchecked XML paragraph index; uncertain rows use native metadata.
    if(fast!=null&&!string.IsNullOrWhiteSpace(text)){while(metadataIndex<keys.Length&&keys[metadataIndex].Length==0)metadataIndex++;string key=Key(text);for(int k=metadataIndex;k<Math.Min(fast.Count,metadataIndex+16);k++){if(keys[k]==key){row=fast[k];metadataIndex=k+1;break;}}}
    if(row==null){
     row=new ParagraphInfo{Text=text,OutlineLevel=10};
     try{row.OutlineLevel=(int)p.OutlineLevel;}catch(Exception ex) when(ex is System.Runtime.InteropServices.COMException||ex is Microsoft.CSharp.RuntimeBinder.RuntimeBinderException){}
     if(text.Length>0){
      if(text.Length<=140||row.OutlineLevel<=9){
       try{row.Style=(string)r.Style.NameLocal;row.Size=(double)r.Font.Size;if(row.Size>100)row.Size=(double)r.Characters[1].Font.Size;row.Bold=(int)r.Font.Bold==-1;row.Centered=(int)p.Alignment==1;row.SpaceBefore=(double)p.SpaceBefore;row.InTable=(bool)r.Information[12];}catch(Exception ex) when(ex is System.Runtime.InteropServices.COMException||ex is Microsoft.CSharp.RuntimeBinder.RuntimeBinderException){}
       try{if((int)r.ListFormat.ListType!=0){row.ListLevel=(int)r.ListFormat.ListLevelNumber;row.Number=(string)r.ListFormat.ListString;}}catch(Exception ex) when(ex is System.Runtime.InteropServices.COMException||ex is Microsoft.CSharp.RuntimeBinder.RuntimeBinderException){}
       try{if((int)r.OMaths.Count>0){row.HasMath=true;var prose=new StringBuilder();int pos=start;foreach(dynamic math in r.OMaths){int a=Math.Max(start,(int)math.Range.Start),b=Math.Min(end,(int)math.Range.End);if(a>pos)prose.Append((string)doc.Range(pos,a).Text);pos=Math.Max(pos,b);}if(pos<end)prose.Append((string)doc.Range(pos,end).Text);row.NonMathText=prose.ToString();}}catch(Exception ex) when(ex is System.Runtime.InteropServices.COMException||ex is Microsoft.CSharp.RuntimeBinder.RuntimeBinderException){}
      }else if(samples++<32)try{row.Size=(double)r.Font.Size;if(row.Size>100)row.Size=(double)r.Characters[1].Font.Size;}catch(Exception ex) when(ex is System.Runtime.InteropServices.COMException||ex is Microsoft.CSharp.RuntimeBinder.RuntimeBinderException){}
     }
    }
    row.Paragraph=++index;row.Start=start;row.End=end;row.Text=text;
    row.InToc|=tocRanges.Any(t=>row.Start>=t.Item1&&row.Start<t.Item2);
    rows.Add(row);
    if(watch.ElapsedMilliseconds>=12){progress?.Invoke(index,count);await Task.Delay(1);watch.Restart();}
   }
   NavigationEngine.MarkDocumentTitles(rows);progress?.Invoke(count,count);return rows;
  }
 }
}
