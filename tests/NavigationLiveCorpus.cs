using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using Forms=System.Windows.Forms;
using PagePilot;
class NavigationLiveCorpus {
 [STAThread] static void Main(string[] args){var form=new Forms.Form{Width=400,Height=150,Text="PagePilot public corpus QA"};form.Shown+=async(s,e)=>{try{await Run(args);}catch(Exception ex){Console.WriteLine(ex);Environment.ExitCode=1;}finally{form.Close();}};Forms.Application.Run(form);}
 static async Task Run(string[] args){
  var json=new JavaScriptSerializer{MaxJsonLength=int.MaxValue};var manifest=json.Deserialize<List<Dictionary<string,object>>>(File.ReadAllText(args[0]));var output=new List<object>();bool ai=args.Contains("--ai");if(args.Contains("--resume")&&File.Exists(args[2]))output.AddRange(json.Deserialize<List<object>>(File.ReadAllText(args[2])));
  dynamic app=Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application"));dynamic doc=null;if((int)app.Documents.Count!=0)throw new Exception("Requires isolated Word instance");bool links=app.Options.UpdateLinksAtOpen;app.Options.UpdateLinksAtOpen=false;app.DisplayAlerts=0;app.AutomationSecurity=3;
  try{foreach(var m in manifest){
   string file=(string)m["file"];if(args.Contains("--resume")&&File.Exists(args[2])&&json.Deserialize<List<Dictionary<string,object>>>(File.ReadAllText(args[2])).Any(x=>(string)x["file"]==file))continue;if(ai&&!new[]{"bf958006598806b3.docx","7c0f84497eabbe58.docx","8b7d81fefcab770b.docx"}.Contains(file))continue;
   var watch=Stopwatch.StartNew();doc=app.Documents.Open(Path.GetFullPath(Path.Combine(args[1],file)),ConfirmConversions:false,ReadOnly:true,AddToRecentFiles:false,Visible:false,NoEncodingDialog:true);string before=(string)doc.Content.Text;bool saved=doc.Saved;
   var rows=await NavigationReader.Read((object)doc,null,CancellationToken.None);var headings=NavigationEngine.Analyze(rows,Engine.Defaults);long local=watch.ElapsedMilliseconds;
   int ticks=0;List<Heading> actualAi=null;
   if(ai){
    var method=typeof(Pane).Assembly.GetType("PagePilot.NavigationAi").GetMethod("Analyze",BindingFlags.Static|BindingFlags.NonPublic);
    var input=rows.Where(p=>!NavigationEngine.Excluded(p)).Select(p=>new Heading{Paragraph=p.Paragraph,Text=p.Text,Start=p.Start,End=p.End}).ToList();
    var progress=new Progress<string>(t=>{ticks++;Console.WriteLine(file+" "+t+" elapsed="+watch.ElapsedMilliseconds);});
    using(var limit=new CancellationTokenSource(TimeSpan.FromMinutes(8)))actualAi=await (Task<List<Heading>>)method.Invoke(null,new object[]{input,"识别正文标题层级，排除封面、表格数据、公式、正文列表、目录页重复条目。",progress,limit.Token,rows,new NavigationAppearance()});
    if(actualAi.Count==0||ticks==0||actualAi.Any(h=>NavigationEngine.Excluded(rows[h.Paragraph-1])))throw new Exception("Invalid live AI result "+file);
    Console.WriteLine("LIVE AI "+file+" headings="+actualAi.Count+" batches="+ticks);
   }
   NavigationReader.Resolve((object)doc,headings);foreach(var h in headings){if(h.Start<0||h.End<h.Start||h.End>(int)doc.Content.End)throw new Exception("Invalid Word range "+file);}
   bool stable=(string)doc.Content.Text==before;Console.WriteLine("WORD state textStable="+stable+" savedBefore="+saved+" savedAfter="+(bool)doc.Saved);
   if(!stable){before=(string)doc.Content.Text;rows=await NavigationReader.Read((object)doc,null,CancellationToken.None);headings=NavigationEngine.Analyze(rows,Engine.Defaults);NavigationReader.Resolve((object)doc,headings);if((string)doc.Content.Text!=before)throw new Exception("Word document did not settle "+file);}
   output.Add(new{file,paragraphs=rows.Count,headings,ai=actualAi,localMs=local,totalMs=watch.ElapsedMilliseconds});Console.WriteLine("WORD "+file+" rows="+rows.Count+" headings="+headings.Count+" ms="+watch.ElapsedMilliseconds);
   doc.Close(0);doc=null;File.WriteAllText(args[2],json.Serialize(output));
  }}finally{if(doc!=null)doc.Close(0);app.Options.UpdateLinksAtOpen=links;app.Quit();}
  Console.WriteLine("PASS "+output.Count+" live Word documents; AI="+ai);
 }
}
