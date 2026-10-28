using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
namespace PagePilot {
 // A retained, lazy interval index. Word positions are not UTF-16 offsets.
 // Matching uses immutable text on a worker; only requested hit geometry touches
 // Word's STA. No COM proxies or mutable interval nodes cross to the worker.
 internal sealed class SearchSnapshot {
  sealed class Span {
   internal int Start,End,Offset,Length;internal Span Left,Right;internal bool? Linear;
  }
  sealed class ProjectionException:InvalidOperationException {internal ProjectionException():base("此匹配包含无法安全定位的对象，请缩小范围。") {}}
  private object document;private int rangeEnd;private readonly List<Span> spans=new List<Span>();
  internal string Text;internal int Reads;internal long MaxSliceMilliseconds;
  internal static async Task<SearchSnapshot> Read(object document,int start,int end,Action<int,int> progress,CancellationToken cancel){
   dynamic doc=document;var result=new SearchSnapshot{document=document,rangeEnd=end};var text=new StringBuilder();
   var omitted=new List<Tuple<int,int>>();var fields=new Queue<object>();var seen=new HashSet<int>();var slice=Stopwatch.StartNew();
   // Include enclosing fields when a selection starts within code or a result.
   foreach(dynamic field in doc.Fields)fields.Enqueue((object)field);
   while(fields.Count>0){
    cancel.ThrowIfCancellationRequested();dynamic field=fields.Dequeue();dynamic code=field.Code,display=field.Result;
    int a=(int)code.Start-1,b=(int)display.Start,c=(int)display.End;
    if(!seen.Add(a))continue;
    if(a<end&&c>=start){omitted.Add(Tuple.Create(a,b));omitted.Add(Tuple.Create(c,c+1));}
    foreach(dynamic nested in code.Fields)fields.Enqueue((object)nested);
    foreach(dynamic nested in display.Fields)fields.Enqueue((object)nested);
    if(slice.ElapsedMilliseconds>=8){result.MaxSliceMilliseconds=Math.Max(result.MaxSliceMilliseconds,slice.ElapsedMilliseconds);await Task.Delay(1,cancel);slice.Restart();}
   }
   int position=start;var ranges=new List<Tuple<int,int>>();
   foreach(var gap in omitted.OrderBy(g=>g.Item1)){
    int a=Math.Max(start,gap.Item1),b=Math.Min(end,gap.Item2);if(b<=position||a>=end)continue;
    if(a>position)ranges.Add(Tuple.Create(position,a));position=Math.Max(position,b);
   }
   if(position<end)ranges.Add(Tuple.Create(position,end));
   foreach(var range in ranges){
    cancel.ThrowIfCancellationRequested();string value=(string)doc.Range(range.Item1,range.Item2).Text??"";result.Reads++;
    result.spans.Add(new Span{Start=range.Item1,End=range.Item2,Offset=text.Length,Length=value.Length});text.Append(value);
    if(slice.ElapsedMilliseconds>=8){result.MaxSliceMilliseconds=Math.Max(result.MaxSliceMilliseconds,slice.ElapsedMilliseconds);progress?.Invoke(range.Item2-start,end-start);await Task.Delay(1,cancel);slice.Restart();}
   }
   result.Text=text.ToString();return result;
  }
  internal int TextOffset(int position){
   dynamic doc=document;
   foreach(var span in spans){if(position<=span.Start)return span.Offset;if(position<span.End){Reads++;return span.Offset+((string)doc.Range(span.Start,position).Text??"").Length;}}
   return Text.Length;
  }
  internal void Resolve(Hit hit){
   if(hit.PositionResolved)return;
   var slice=Stopwatch.StartNew();
   try{
    hit.Start=Position(hit.SourceStart,false);
    hit.Length=hit.SourceLength==0?0:Position(hit.SourceStart+hit.SourceLength-1,true)-hit.Start;
   }catch(ProjectionException){if(!ResolveWithFind(hit))throw;}
   hit.PositionResolved=true;MaxSliceMilliseconds=Math.Max(MaxSliceMilliseconds,slice.ElapsedMilliseconds);
  }
  bool ResolveWithFind(Hit hit){
   // Form controls/equations can project different strings for partial ranges.
   // In that exceptional branch use Word's own literal finder, and verify both
   // the returned text and its ordinal in the original snapshot. Bounded work;
   // duplicate text can never silently move the hit to another occurrence.
   string needle=Text.Substring(hit.SourceStart,hit.SourceLength),findText=needle.Replace("^","^^");
   if(needle.Length==0||findText.Length>255)return false;
   var span=spans.FirstOrDefault(s=>s.Offset<=hit.SourceStart&&s.Offset+s.Length>hit.SourceStart);if(span==null)return false;
   int finish=hit.SourceStart+hit.SourceLength;
   while(span.Left!=null){if(finish<=span.Right.Offset)span=span.Left;else if(hit.SourceStart>=span.Right.Offset)span=span.Right;else break;}
   int limit=span.End;if(finish>span.Offset+span.Length)limit=rangeEnd;
   dynamic doc=document;int position=span.Start;
   for(int attempt=0;attempt<64&&position<limit;attempt++){
    dynamic range=doc.Range(position,limit);dynamic find=range.Find;find.ClearFormatting();
    if(!(bool)find.Execute(FindText:findText,MatchCase:true,MatchWholeWord:false,MatchWildcards:false,MatchSoundsLike:false,MatchAllWordForms:false,Forward:true,Wrap:0,Format:false))return false;
    int a=(int)range.Start,b=(int)range.End;
    if((string)range.Text==needle&&TextOffset(a)==hit.SourceStart){hit.Start=a;hit.Length=b-a;return true;}
    position=Math.Max(position+1,b);
   }
   return false;
  }
  int Position(int index,bool end){
   if(index==Text.Length)return rangeEnd;
   int lo=0,hi=spans.Count;while(lo<hi){int middle=lo+(hi-lo)/2;if(spans[middle].Offset+spans[middle].Length<=index)lo=middle+1;else hi=middle;}
   if(lo==spans.Count)throw new InvalidOperationException("匹配位置已失效，请重新查找。");
   var span=spans[lo];dynamic doc=document;
   while(true){
    if(!span.Linear.HasValue)span.Linear=span.End-span.Start==span.Length&&Text.IndexOf('\a',span.Offset,span.Length)<0;
    if(span.Linear==true)return span.Start+index-span.Offset+(end?1:0);
    if(span.End-span.Start<=1)return end?span.End:span.Start;
    if(span.Left==null){
     int middle=span.Start+(span.End-span.Start)/2;string left=(string)doc.Range(span.Start,middle).Text??"";Reads++;
     // Cutting through a Word equation/object can change its text projection.
     // Retry at its enclosing paragraph boundary, never accept a guessed offset.
     bool aligned=left.Length<=span.Length&&string.CompareOrdinal(Text,span.Offset,left,0,left.Length)==0;
     if(!aligned){
      dynamic paragraph=doc.Range(middle,middle).Paragraphs.First.Range;
      foreach(int boundary in new[]{(int)paragraph.Start,(int)paragraph.End}){
       if(boundary<=span.Start||boundary>=span.End||boundary==middle)continue;
       string candidate=(string)doc.Range(span.Start,boundary).Text??"";Reads++;
       if(candidate.Length<=span.Length&&string.CompareOrdinal(Text,span.Offset,candidate,0,candidate.Length)==0){middle=boundary;left=candidate;aligned=true;break;}
      }
     }
     if(!aligned)throw new ProjectionException();
     span.Left=new Span{Start=span.Start,End=middle,Offset=span.Offset,Length=left.Length};
     span.Right=new Span{Start=middle,End=span.End,Offset=span.Offset+left.Length,Length=span.Length-left.Length};
    }
    span=index<span.Right.Offset?span.Left:span.Right;
   }
  }
  internal async Task ResolveAll(IEnumerable<Hit> hits,CancellationToken cancel){
   var slice=Stopwatch.StartNew();foreach(var hit in hits){cancel.ThrowIfCancellationRequested();Resolve(hit);if(slice.ElapsedMilliseconds>=8){await Task.Delay(1,cancel);slice.Restart();}}
  }
 }
}
