using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Reflection;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using PagePilot;
class NavigationAiLive {
 class Sample {public string file;public List<ParagraphInfo> rows;}
 static void Main(string[] args){try{Run(args).GetAwaiter().GetResult();}catch(Exception e){Console.WriteLine(e);Environment.ExitCode=1;}}
 static async Task Run(string[] args){
  var json=new JavaScriptSerializer{MaxJsonLength=int.MaxValue};var samples=json.Deserialize<List<Sample>>(File.ReadAllText(args[0]));var output=new List<object>();var method=typeof(ParagraphInfo).Assembly.GetType("PagePilot.NavigationAi").GetMethod("Analyze",BindingFlags.Static|BindingFlags.NonPublic);
  foreach(var sample in samples.Where(s=>new[]{"bf958006598806b3.docx","7c0f84497eabbe58.docx","8b7d81fefcab770b.docx"}.Contains(s.file))){
   var input=sample.rows.Where(p=>!NavigationEngine.Excluded(p)).Select(p=>new Heading{Paragraph=p.Paragraph,Text=p.Text}).ToList();var watch=Stopwatch.StartNew();var updates=new List<string>();
   var progress=new Progress<string>(t=>{lock(updates)updates.Add(t);Console.WriteLine(sample.file+" "+t+" elapsed="+watch.ElapsedMilliseconds);});
   List<Heading> result;using(var timeout=new CancellationTokenSource(TimeSpan.FromMinutes(8)))result=await (Task<List<Heading>>)method.Invoke(null,new object[]{input,"识别正文标题层级，排除封面、表格数据、公式、正文列表、目录页重复条目。",progress,timeout.Token,sample.rows,new NavigationAppearance()});
   if(result.Count==0||result.Any(h=>NavigationEngine.Excluded(sample.rows[h.Paragraph-1])))throw new Exception("Invalid AI classification");
   output.Add(new{sample.file,ms=watch.ElapsedMilliseconds,progress=updates.ToArray(),headings=result});File.WriteAllText(args[1],json.Serialize(output));Console.WriteLine("PASS REAL API "+sample.file+" headings="+result.Count+" ms="+watch.ElapsedMilliseconds);
  }
 }
}
