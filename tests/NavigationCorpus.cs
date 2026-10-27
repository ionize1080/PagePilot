using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Diagnostics;
using System.Xml.Linq;
using System.IO.Compression;
using System.Web.Script.Serialization;
using PagePilot;
class NavigationCorpus {
 static int Main(string[] args){try{
  var json=new JavaScriptSerializer{MaxJsonLength=int.MaxValue};var manifest=json.Deserialize<List<Dictionary<string,object>>>(File.ReadAllText(args[0]));var output=new List<object>();int paragraphs=0,headings=0,checks=0;var total=Stopwatch.StartNew();
  foreach(var item in manifest){string file=(string)item["file"];var watch=Stopwatch.StartNew();List<ParagraphInfo> rows;
   using(var zip=ZipFile.OpenRead(Path.Combine(args[1],file)))rows=NavigationSnapshot.Read(XDocument.Load(zip.GetEntry("word/document.xml").Open()),zip.GetEntry("word/styles.xml")==null?null:XDocument.Load(zip.GetEntry("word/styles.xml").Open()));
   var actual=NavigationEngine.Analyze(rows,Engine.Defaults);
   Func<string,string> key=t=>System.Text.RegularExpressions.Regex.Replace(t,@"\s","");
   if(item.ContainsKey("expectedHeadings"))foreach(string text in (System.Collections.ArrayList)item["expectedHeadings"]){if(!actual.Any(h=>key(h.Text).Contains(key(text))))throw new Exception(file+" missing annotated heading: "+text);checks++;}
   if(item.ContainsKey("excludedSnippets"))foreach(string text in (System.Collections.ArrayList)item["excludedSnippets"]){if(actual.Any(h=>key(h.Text).Contains(key(text))))throw new Exception(file+" included annotated non-heading: "+text);checks++;}
   if(actual.Any(h=>h.Level<1||h.Level>9)||actual.Select(h=>h.Paragraph).Distinct().Count()!=actual.Count||!actual.Select(h=>h.Paragraph).SequenceEqual(actual.Select(h=>h.Paragraph).OrderBy(x=>x)))throw new Exception(file+" invalid hierarchy/order");checks+=3;
   foreach(var p in rows.Where(p=>NavigationEngine.Excluded(p))){if(actual.Any(h=>h.Paragraph==p.Paragraph))throw new Exception(file+" excluded content leaked");checks++;}
   output.Add(new{file,title=item["title"],paragraphs=rows.Count,ms=watch.ElapsedMilliseconds,headings=actual,rows});paragraphs+=rows.Count;headings+=actual.Count;
   Console.WriteLine(file+" paragraphs="+rows.Count+" headings="+actual.Count+" ms="+watch.ElapsedMilliseconds);
  }
  File.WriteAllText(args[2],json.Serialize(output));Console.WriteLine($"CORPUS {manifest.Count} documents, {paragraphs} paragraphs, {headings} headings, {checks} invariants, {total.ElapsedMilliseconds} ms");return 0;
 }catch(Exception ex){Console.WriteLine(ex);return 1;}}
}
