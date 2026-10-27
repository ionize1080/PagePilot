using System;
using System.Linq;
using System.Collections.Generic;
using System.Diagnostics;
using System.Xml.Linq;
using System.IO.Compression;
using PagePilot;
class NavigationRegression {
 static int checks;
 static void Assert(bool ok,string name){if(!ok)throw new Exception(name);checks++;Console.WriteLine("PASS "+name);}
 static ParagraphInfo P(string text,int level=0)=>new ParagraphInfo{Text=text,OutlineLevel=level,Style="Normal",Size=11};
 static List<Heading> Run(params ParagraphInfo[] p)=>NavigationEngine.Analyze(p,Engine.Defaults);
 static void Main(string[] args){
  foreach(var text in new[]{"1 引言","1. Introduction","一、总则","（一）范围","Chapter One Introduction","chapter IV — Methods","PART II: Results","附录 A","Abstract","1.2.3范围","１．２ 设计","IV. Scope","(iv) Detail","A. Background","(a) Detail","Ⅰ 概述","Ⅳ. 方法","① 引言","A、范围","One. Introduction","第一百零二章 范围","第壹章 总则"})Assert(Run(P(text)).Count==1,"Recognize "+text);
  foreach(var text in new[]{"2026 年 10 月","2026.10.27 会议记录","3.5 亿元","192.168.1.1 server","普通正文。","Part of the reason is cost.","1. This item is a sentence; it ends here."})Assert(Run(P(text)).Count==0,"Exclude "+text);
  var longTitle=P(new string('长',180),3);Assert(Run(longTitle)[0].Level==3,"Explicit outline overrides length");
  var styled=P("没有编号的标题");styled.Style="标题 2";Assert(Run(styled)[0].Level==2,"Localized style");
  var toc=P("1. Introduction",1);toc.InToc=true;Assert(Run(toc).Count==0,"TOC duplicate excluded");
  var table=P("1. 数据");table.InTable=true;Assert(Run(table).Count==0,"Table false positive excluded");table.OutlineLevel=2;Assert(Run(table).Count==1,"Explicit table heading retained");
  var list=P("1. ordinary list");list.ListLevel=1;Assert(Run(list).Count==0,"Ordinary automatic lists excluded");list.Bold=true;Assert(Run(list).Count==1,"Formatted numbered heading");
  var levels=Run(P("第一篇 理论"),P("第一章 导论"),P("第一节 范围"));Assert(levels.Select(x=>x.Level).SequenceEqual(new[]{1,2,3}),"Part chapter section hierarchy");
  levels=Run(P("一、总则"),P("（一）范围"),P("1. 对象"),P("（1）细节"));Assert(levels.Select(x=>x.Level).SequenceEqual(new[]{1,2,3,4}),"Chinese numeric hierarchy");
  levels=Run(P("I. Introduction"),P("A. Scope"),P("1. Background"),P("i. Detail"));Assert(levels.Select(x=>x.Level).SequenceEqual(new[]{1,2,3,4}),"Roman letter numeric hierarchy");
  var format=P("未编号标题");format.Size=18;format.Bold=true;Assert(Run(format)[0].Level==1,"Typography inference");
  var custom=new[]{new HeadingRule{Name="自定义",Pattern="^SPECIAL",Level=4}};Assert(NavigationEngine.Analyze(new[]{P("SPECIAL",1)},custom)[0].Level==4,"Custom precedence");
  var appearance=new NavigationAppearance();appearance.Validate();appearance.Levels[0].Size=double.NaN;bool bad=false;try{appearance.Validate();}catch(ArgumentException){bad=true;}Assert(bad,"Appearance rejects invalid size");
  var cancel=new System.Threading.CancellationTokenSource();cancel.Cancel();bad=false;try{NavigationEngine.Analyze(new[]{P("1 Test")},Engine.Defaults,cancel.Token);}catch(OperationCanceledException){bad=true;}Assert(bad,"Analysis cancellation");
  var big=Enumerable.Range(0,50000).Select(i=>P(i%50==0?(i/50+1)+". Title":new string('文',150)+"。")).ToList();var watch=Stopwatch.StartNew();var result=NavigationEngine.Analyze(big,Engine.Defaults);Console.WriteLine("50,000 paragraphs: "+watch.ElapsedMilliseconds+" ms; headings="+result.Count);Assert(watch.Elapsed.TotalSeconds<5,"Bounded large document analysis");
  if(args.Length>0){var rows=ReadDocx(args[0]);var actual=NavigationEngine.Analyze(rows,Engine.Defaults);foreach(var h in actual)Console.WriteLine("DOCX "+h.Paragraph+" L"+h.Level+" "+h.Text);Assert(Enumerable.Range(1,8).All(i=>actual.Any(h=>h.Text.StartsWith(i+" ")&&h.Level==1)),"User ETH document all eight sections");}
  Console.WriteLine(checks+" checks passed");
 }
 static List<ParagraphInfo> ReadDocx(string file){
  XNamespace w="http://schemas.openxmlformats.org/wordprocessingml/2006/main";var rows=new List<ParagraphInfo>();
  using(var z=ZipFile.OpenRead(file)){
   var styles=XDocument.Load(z.GetEntry("word/styles.xml").Open()).Descendants(w+"style").ToDictionary(x=>(string)x.Attribute(w+"styleId"));
   int i=0;foreach(var p in XDocument.Load(z.GetEntry("word/document.xml").Open()).Root.Element(w+"body").Elements(w+"p")){
    var row=P(string.Concat(p.Descendants(w+"t").Select(x=>x.Value)));row.Paragraph=++i;
    var pr=p.Element(w+"pPr");string id=(string)pr?.Element(w+"pStyle")?.Attribute(w+"val");XElement st=null;if(id!=null)styles.TryGetValue(id,out st);
    row.Style=(string)st?.Element(w+"name")?.Attribute(w+"val")??"Normal";
    var ol=pr?.Element(w+"outlineLvl")??st?.Element(w+"pPr")?.Element(w+"outlineLvl");row.OutlineLevel=ol==null?0:(int)ol.Attribute(w+"val")+1;
    var rp=pr?.Element(w+"rPr")??p.Element(w+"r")?.Element(w+"rPr");row.Size=((double?)rp?.Element(w+"sz")?.Attribute(w+"val")??22)/2;row.Bold=rp?.Element(w+"b")!=null;
    row.SpaceBefore=((double?)pr?.Element(w+"spacing")?.Attribute(w+"before")??0)/20;row.Centered=(string)pr?.Element(w+"jc")?.Attribute(w+"val")=="center";rows.Add(row);
   }
  }return rows;
 }
}
