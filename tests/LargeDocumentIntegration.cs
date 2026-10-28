using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Web.Script.Serialization;
using PagePilot;
using Forms=System.Windows.Forms;
class LargeDocumentIntegration {
 static readonly BindingFlags flags=BindingFlags.Instance|BindingFlags.NonPublic|BindingFlags.Public;
 static object Field(object obj,string name)=>obj.GetType().GetField(name,flags).GetValue(obj);
 static object Call(object obj,string name,params object[] args)=>obj.GetType().GetMethods(flags).Single(m=>m.Name==name&&m.GetParameters().Length==args.Length).Invoke(obj,args);
 static T UI<T>(Pane pane,string name)where T:class=>((FrameworkElement)Field(pane,"view")).FindName(name)as T;
 static void Assert(bool ok,string name){if(!ok)throw new Exception(name);Console.WriteLine("PASS "+name);}
 static async Task Until(Func<bool> ready,int seconds=240){var watch=Stopwatch.StartNew();while(!ready()){if(watch.Elapsed.TotalSeconds>seconds)throw new Exception("Timed out");await Task.Delay(20);}}
 [STAThread]static void Main(string[] args){Forms.Application.EnableVisualStyles();var runner=new Forms.Form{Width=450,Height=850,Text="PagePilot large document verification"};runner.Shown+=async(s,e)=>{try{await Run(runner,args);}catch(Exception ex){Console.WriteLine(ex);Environment.ExitCode=1;}finally{runner.Close();}};Forms.Application.Run(runner);}
 static async Task Run(Forms.Form runner,string[] args){
  dynamic app=Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application"));dynamic doc=null;Assert((int)app.Documents.Count==0,"isolated Word");app.Visible=true;app.DisplayAlerts=0;app.AutomationSecurity=3;app.Options.UpdateLinksAtOpen=false;
  var results=new List<object>();var gaps=new List<long>();var pulse=Stopwatch.StartNew();var heartbeat=new Forms.Timer{Interval=15};heartbeat.Tick+=(s,e)=>{gaps.Add(pulse.ElapsedMilliseconds);pulse.Restart();};
  try{
   var files=args.Where(a=>a!="--synthetic").ToList();if(args.Contains("--synthetic"))files.Insert(0,"");
   foreach(string file in files){
    if(file==""){
     doc=app.Documents.Add();doc.SetCompatibilityMode(14);doc.Content.Text=string.Join("\r",Enumerable.Range(0,6000).Select(i=>i%50==0?"第"+(i/50+1)+"章 TEST 标题":"正文 TEST 12345 😀 数据。"))+"\r";
     doc.Hyperlinks.Add(doc.Range(0,3),"https://example.com");dynamic table=doc.Tables.Add(doc.Range((int)doc.Content.End-1,(int)doc.Content.End-1),50,5);table.Range.Text="";
     table.Cell(1,1).Range.Text="CELL TEST 4321";
     doc.Range(60,64).Font.Hidden=-1;
     doc.SaveAs2(Path.GetFullPath(".temp/large-synthetic-047.docx"));
    }else doc=app.Documents.Open(Path.GetFullPath(file),ReadOnly:true,AddToRecentFiles:false,Visible:true);
    string before=(string)doc.Content.Text;int end=(int)doc.Content.End;bool saved=doc.Saved;
    var host=new Connect();typeof(Connect).GetField("App",flags).SetValue(host,app);var pane=new Pane{Dock=Forms.DockStyle.Fill};runner.Controls.Add(pane);Call(pane,"Bind",host,(object)doc,0);((Forms.Timer)Field(pane,"highlightTimer")).Stop();
    UI<TextBox>(pane,"Pattern").Text=@"\d+";
    gaps.Clear();pulse.Restart();heartbeat.Start();var watch=Stopwatch.StartNew();
    await Until(()=>Field(pane,"snapshot")!=null||UI<TextBlock>(pane,"Status").Text.Contains("无法")||UI<TextBlock>(pane,"Status").Text.Contains("变化"));heartbeat.Stop();long searchMs=watch.ElapsedMilliseconds;
    Assert(Field(pane,"snapshot")!=null,"search completed: "+Path.GetFileName(file));
    var hits=(List<Hit>)Field(pane,"found");Assert(hits.Count>0,"matches found");
    var map=Field(pane,"searchSnapshot");Console.WriteLine("SEARCH "+Path.GetFileName(file)+" ms="+searchMs+" hits="+hits.Count+" reads="+Field(map,"Reads")+" maxSlice="+Field(map,"MaxSliceMilliseconds"));
    // Every returned position must be valid; sample actual Word text evenly across
    // the full story, not just a short prefix before the first table or field.
    Assert(hits.All(h=>h.Length>=0),"all match lengths bounded");
    foreach(var hit in hits.Where((h,i)=>i%Math.Max(1,hits.Count/100)==0)){Call(pane,"ResolveHit",hit);AssertText(doc,hit);}
    var searchGaps=gaps.ToArray();var mapping=Field(pane,"searchSnapshot");UI<TextBox>(pane,"Pattern").Text=@"[0-9]{2,}";watch.Restart();await Until(()=>Field(pane,"snapshot")!=null);long cachedMs=watch.ElapsedMilliseconds;Assert(ReferenceEquals(mapping,Field(pane,"searchSnapshot")),"query reuses immutable position map");
    UI<TextBox>(pane,"Replacement").Text="replacement";watch.Restart();await Until(()=>Field(pane,"searchCancel")==null&&!(bool)Field(pane,"updatingReplacement"));Assert(((List<Hit>)Field(pane,"found")).All(h=>h.After=="replacement"),"latest replacement preview");
    pane.Dispose();pane=new Pane{Dock=Forms.DockStyle.Fill};runner.Controls.Add(pane);Call(pane,"Bind",host,(object)doc,1);
    Call(pane,"RefreshOutline");await Task.Delay(80);watch.Restart();UI<Button>(pane,"StopScan").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));await Until(()=>!(bool)Field(pane,"scanning"));long cancelMs=watch.ElapsedMilliseconds;Assert(UI<TextBlock>(pane,"Status").Text.Contains("取消"),"navigation cancellation");
    gaps.Clear();pulse.Restart();heartbeat.Start();watch.Restart();Call(pane,"RefreshOutline");await Until(()=>!(bool)Field(pane,"scanning"));heartbeat.Stop();long navigationMs=watch.ElapsedMilliseconds;
    Assert(Field(pane,"outlineSnapshot")!=null,"navigation complete: "+UI<TextBlock>(pane,"Status").Text);var headings=(List<Heading>)Field(pane,"outline");
    Assert(headings.All(h=>h.Start>=0&&h.End>=h.Start&&h.End<=end),"all heading ranges bounded");
    var navGaps=gaps.ToArray();Assert((string)doc.Content.Text==before&&(bool)doc.Saved==saved,"document remains unchanged");
    var appearance=(NavigationAppearance)Field(host,"NavigationStyle");appearance.VisualStyle="acrobat";host.GetType().GetField("NavigationStyleVersion",flags).SetValue(host,1);Call(pane,"ApplyNavigationAppearance");
    var view=(FrameworkElement)Field(pane,"view");Assert((Thickness)view.Resources["FrameLine"]==new Thickness(0),"Acrobat is borderless");appearance.VisualStyle="colorful";host.GetType().GetField("NavigationStyleVersion",flags).SetValue(host,2);Call(pane,"ApplyNavigationAppearance");Assert((Thickness)view.Resources["FrameLine"]==new Thickness(1),"colorful style restored");
    results.Add(new{file=Path.GetFileName(file),searchMs,cachedMs,navigationMs,cancelMs,matches=hits.Count,headings=headings.Count,reads=Field(map,"Reads"),maxReadSliceMs=Field(map,"MaxSliceMilliseconds"),searchPulseP95=Percentile(searchGaps),searchPulseMax=searchGaps.DefaultIfEmpty().Max(),navigationPulseP95=Percentile(navGaps),navigationPulseMax=navGaps.DefaultIfEmpty().Max()});
    File.WriteAllText(".temp/large-document-results.json",new JavaScriptSerializer().Serialize(results));Console.WriteLine("DOCUMENT "+new JavaScriptSerializer().Serialize(results.Last()));pane.Dispose();doc.Close(0);doc=null;
   }
  }finally{heartbeat.Dispose();if(doc!=null)doc.Close(0);app.Quit();}
 }
 static long Percentile(long[] values)=>values.Length==0?0:values.OrderBy(v=>v).ElementAt(Math.Min(values.Length-1,(int)(values.Length*.95)));
 static void AssertText(dynamic doc,Hit hit){if((string)doc.Range(hit.Start,hit.Start+hit.Length).Text!=hit.Before)throw new Exception("Incorrect mapped text at "+hit.Start);}
}
