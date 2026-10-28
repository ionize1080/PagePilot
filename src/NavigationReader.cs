using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
namespace PagePilot {
 // Runs on Word's STA, yielding between short batches. Only immutable values leave it.
 public static class NavigationReader {
  sealed class TextParagraph {internal int Start,Length;internal string Key;}
  // Saved DOCX: analyze streaming metadata off-thread, align it with live text,
  // then resolve only headings through Word. Never scan 15,000 native paragraphs
  // merely to display a few hundred navigation destinations.
  public static async Task<List<Heading>> TryReadHeadings(object document,IEnumerable<HeadingRule> rules,NavigationAppearance options,Action<int,int> progress,CancellationToken cancel){
   dynamic doc=document;string path=null;try{if((bool)doc.Saved)path=(string)doc.FullName;}catch{}
   if(path==null||!System.IO.File.Exists(path))return null;
   List<ParagraphInfo> metadata;try{metadata=await Task.Run(()=>NavigationSnapshot.ReadFile(path,cancel),cancel);}catch(OperationCanceledException){throw;}catch{return null;}
   if(metadata==null)return null;
   var candidates=await Task.Run(()=>NavigationEngine.Analyze(metadata,rules,cancel,options),cancel);
   int documentEnd=(int)doc.Content.End;
   var snapshot=await SearchSnapshot.Read(document,0,documentEnd,progress,cancel);
   var paragraphs=new List<TextParagraph>();int start=0;
   for(int i=0;i<snapshot.Text.Length;i++)if(snapshot.Text[i]=='\r'){
    int a=start;while(a<i&&snapshot.Text[a]=='\a')a++;
    paragraphs.Add(new TextParagraph{Start=a,Length=i-a,Key=Key(snapshot.Text.Substring(a,i-a))});start=i+1;
   }
   if(start<snapshot.Text.Length)paragraphs.Add(new TextParagraph{Start=start,Length=snapshot.Text.Length-start,Key=Key(snapshot.Text.Substring(start))});
   var index=paragraphs.Select((p,i)=>new{p.Key,Index=i}).Where(x=>x.Key.Length>0).GroupBy(x=>x.Key).ToDictionary(g=>g.Key,g=>g.Select(x=>x.Index).ToList());
   var located=new Dictionary<int,TextParagraph>();int cursor=0,matched=0,total=0;
   foreach(var row in metadata){
    cancel.ThrowIfCancellationRequested();string key=Key(row.Text);if(key.Length==0)continue;total+=key.Length;
    if(!index.TryGetValue(key,out var positions))continue;
    int found=positions.BinarySearch(cursor);if(found<0)found=~found;if(found==positions.Count)continue;
    int n=positions[found];cursor=n+1;located[row.Paragraph]=paragraphs[n];matched+=key.Length;
   }
   if(total>0&&matched<total*.9||candidates.Any(h=>!located.ContainsKey(h.Paragraph))){Connect.Log("navigation.index.fallback aligned="+matched+"/"+total+" missing="+candidates.Count(h=>!located.ContainsKey(h.Paragraph)));return null;}
   var slice=System.Diagnostics.Stopwatch.StartNew();int done=0;
   foreach(var h in candidates){
    cancel.ThrowIfCancellationRequested();var paragraph=located[h.Paragraph];
    var hit=new Hit{SourceStart=paragraph.Start,SourceLength=paragraph.Length};try{snapshot.Resolve(hit);}catch(InvalidOperationException){return null;}
    h.Start=hit.Start;h.End=hit.Start+hit.Length;
    string live=(string)doc.Range(h.Start,h.End).Text;
    if(live!=snapshot.Text.Substring(paragraph.Start,paragraph.Length))return null;
    // The resolved Range is authoritative. Prefix Paragraphs.Count is both
    // quadratic and inconsistent with the full collection in nested tables.
    // Zero denotes a range-based destination; explicit AI paragraph IDs remain
    // unchanged on the native reader path.
    h.Paragraph=0;h.Text=Clean(live);
    if(slice.ElapsedMilliseconds>=8){progress?.Invoke(++done,candidates.Count);await Task.Delay(1,cancel);slice.Restart();}else done++;
   }
   Connect.Log("navigation.indexed headings="+candidates.Count+" metadata="+metadata.Count+" alignedCharacters="+matched+"/"+total+" positionReads="+snapshot.Reads);
   return candidates;
  }
  public static async Task ResolveAsync(object document,IEnumerable<Heading> headings,CancellationToken cancel){
   dynamic doc=document;var slice=System.Diagnostics.Stopwatch.StartNew();
   foreach(var h in headings){
    cancel.ThrowIfCancellationRequested();dynamic r=h.Start>=0&&h.End>=h.Start?doc.Range(h.Start,h.End):doc.Paragraphs[h.Paragraph].Range;
    h.Start=(int)r.Start;h.End=(int)r.End;h.Text=Clean((string)r.Text);
    try{if((int)r.ListFormat.ListType!=0)h.Text=(string)r.ListFormat.ListString+" "+h.Text;}catch(System.Runtime.InteropServices.COMException){}
    if(slice.ElapsedMilliseconds>=8){await Task.Delay(1,cancel);slice.Restart();}
   }
  }
  static string Clean(string text)=>System.Text.RegularExpressions.Regex.Replace((text??"").Trim('\r','\a',' ','\t'),@"^(?:[/\u0001\u0002]*[\f\v])+\s*","").Trim();
  static string Key(string text)=>System.Text.RegularExpressions.Regex.Replace(Clean(text),@"[\s\a\u0001\u0002\u0013\u0014\u0015]","");
  public static void Resolve(object document,IEnumerable<Heading> headings){dynamic doc=document;foreach(var h in headings){dynamic r=doc.Paragraphs[h.Paragraph].Range;h.Start=(int)r.Start;h.End=(int)r.End-1;h.Text=Clean((string)r.Text);try{if((int)r.ListFormat.ListType!=0)h.Text=(string)r.ListFormat.ListString+" "+h.Text;}catch(Exception ex) when(ex is System.Runtime.InteropServices.COMException||ex is Microsoft.CSharp.RuntimeBinder.RuntimeBinderException){}}}
  public static async Task<List<ParagraphInfo>> Read(object document,Action<int,int> progress,CancellationToken cancel){
   dynamic doc=document;var rows=new List<ParagraphInfo>();int count=(int)doc.Paragraphs.Count,index=0,samples=0;
   var tocRanges=new List<Tuple<int,int>>();foreach(dynamic toc in doc.TablesOfContents)tocRanges.Add(Tuple.Create((int)toc.Range.Start,(int)toc.Range.End));
   var watch=System.Diagnostics.Stopwatch.StartNew();
   List<ParagraphInfo> fast=null;int metadataIndex=0;
   // Disk metadata is a hint, only for a saved document. Live text is still read
   // and aligned below, so an old file can never silently supply Word positions.
   string path=null;try{if((bool)doc.Saved)path=(string)doc.FullName;}catch{}
   if(path!=null&&System.IO.File.Exists(path))try{fast=await Task.Run(()=>NavigationSnapshot.ReadFile(path,cancel),cancel);}catch(OperationCanceledException){throw;}catch{fast=null;}
   cancel.ThrowIfCancellationRequested();watch.Restart();
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
