using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using PagePilot;
public class SnapshotDocument {
 public readonly List<string> Units=new List<string>();public readonly List<object> Fields=new List<object>();
 public SnapshotRange Range(int start,int end)=>new SnapshotRange(this,start,end);
 public SnapshotRange Content=>Range(0,Units.Count);
}
public class SnapshotRange {
 readonly SnapshotDocument doc;public int Start{get;}public int End{get;}public List<object> Fields=new List<object>();
 public SnapshotRange(SnapshotDocument d,int a,int b){doc=d;Start=a;End=b;}
 public string Text=>string.Concat(doc.Units.Skip(Start).Take(End-Start));
}
public class SnapshotField {public SnapshotRange Code{get;set;}public SnapshotRange Result{get;set;}}
class SearchSnapshotRegression {
 static int checks;
 static void Assert(bool value,string label){if(!value)throw new Exception(label);checks++;Console.WriteLine("PASS "+label);}
 static void Main(){try{
  var doc=new SnapshotDocument();for(int i=0;i<20000;i++){doc.Units.Add("");foreach(char ch in "WORD "+i)doc.Units.Add(ch.ToString());doc.Units.Add("\r\a");}
  var map=SearchSnapshot.Read(doc,0,doc.Units.Count,null,CancellationToken.None).GetAwaiter().GetResult();
  var hits=Engine.Find(map.Text,0,@"\d+","new",false);Assert(hits.Count==20000&&map.Reads==1,"20,000 results need one initial text read");
  foreach(var hit in hits.Where((h,i)=>i%251==0)){map.Resolve(hit);Assert(doc.Range(hit.Start,hit.Start+hit.Length).Text==hit.Before,"nonlinear position "+hit.SourceStart);}
  Assert(map.Reads<1800,"only visited interval branches are materialized");int count=map.Reads;foreach(var hit in hits.Where((h,i)=>i%251==0))map.Resolve(hit);Assert(map.Reads==count,"retained positions require no further reads");
  var empty=SearchSnapshot.Read(doc,123,123,null,CancellationToken.None).GetAwaiter().GetResult();var zero=Engine.Find("",0,"^","",false)[0];empty.Resolve(zero);Assert(zero.Start==123&&zero.Length==0,"empty selection zero-width position");
  var fields=new SnapshotDocument();foreach(char ch in "A<CODE>VALUE!Z")fields.Units.Add(ch.ToString());fields.Fields.Add(new SnapshotField{Code=fields.Range(2,6),Result=fields.Range(7,12)});
  var fieldMap=SearchSnapshot.Read(fields,0,fields.Units.Count,null,CancellationToken.None).GetAwaiter().GetResult();Assert(fieldMap.Text=="AVALUEZ","field code and delimiters omitted");var value=Engine.Find(fieldMap.Text,0,"VALUE","",false)[0];fieldMap.Resolve(value);Assert(value.Start==7&&fields.Range(value.Start,value.Start+value.Length).Text=="VALUE","field result maps precisely");
  var cancelled=new CancellationTokenSource();cancelled.Cancel();bool stopped=false;try{SearchSnapshot.Read(doc,0,doc.Units.Count,null,cancelled.Token).GetAwaiter().GetResult();}catch(OperationCanceledException){stopped=true;}Assert(stopped,"read cancellation");
  Console.WriteLine(checks+" checks passed; lazy interval reads="+count);
 }catch(Exception ex){Console.WriteLine(ex);Environment.ExitCode=1;}}
}
