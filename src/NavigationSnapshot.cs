using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.IO;
using System.IO.Compression;
using System.Threading;
using System.Xml;
namespace PagePilot {
 // Shared OOXML metadata extraction for the live Word reader and corpus regression.
 public static class NavigationSnapshot {
  // Read only document/styles parts; never ask Word to serialize embedded media.
  internal static List<ParagraphInfo> ReadFile(string path,CancellationToken cancel){
   cancel.ThrowIfCancellationRequested();
   using(var file=new FileStream(path,FileMode.Open,FileAccess.Read,FileShare.ReadWrite|FileShare.Delete))
   using(var zip=new ZipArchive(file,ZipArchiveMode.Read)){
    var part=zip.GetEntry("word/document.xml");if(part==null)return null;
    XDocument styles=null;var style=zip.GetEntry("word/styles.xml");if(style!=null)using(var stream=style.Open())styles=XDocument.Load(stream);
    using(var stream=part.Open())using(var reader=XmlReader.Create(stream,new XmlReaderSettings{DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null}))return ReadParagraphs(StreamParagraphs(reader,cancel),styles,cancel);
   }
  }
  static IEnumerable<XElement> StreamParagraphs(XmlReader reader,CancellationToken cancel){
   int tables=0;var controls=new Stack<Tuple<int,bool>>();
   while(!reader.EOF){
    cancel.ThrowIfCancellationRequested();
    if(reader.NamespaceURI==W.NamespaceName){
     if(reader.NodeType==XmlNodeType.Element){
      if(reader.LocalName=="ins"||reader.LocalName=="del"||reader.LocalName=="altChunk")throw new InvalidDataException("Live revision metadata required");
      if(reader.LocalName=="tbl"&&!reader.IsEmptyElement)tables++;
      if(reader.LocalName=="sdt"&&!reader.IsEmptyElement)controls.Push(Tuple.Create(reader.Depth,false));
      if(reader.LocalName=="docPartGallery"&&reader.GetAttribute("val",W.NamespaceName)=="Table of Contents"&&controls.Count>0){var parent=controls.Pop();controls.Push(Tuple.Create(parent.Item1,true));}
      if(reader.LocalName=="p"){
       var paragraph=(XElement)XNode.ReadFrom(reader);
       if(paragraph.Descendants().Any(x=>x.Name==W+"del"||x.Name==W+"ins"||x.Name==W+"altChunk"))throw new InvalidDataException("Live revision metadata required");
       if(tables>0)new XElement(W+"tbl",paragraph);
       if(controls.Any(c=>c.Item2))paragraph.AddAnnotation(new TocContext());
       yield return paragraph;continue;
      }
     }else if(reader.NodeType==XmlNodeType.EndElement){if(reader.LocalName=="tbl")tables--;if(reader.LocalName=="sdt"&&controls.Count>0&&controls.Peek().Item1==reader.Depth)controls.Pop();}
    }
    reader.Read();
   }
  }
  sealed class TocContext {}
  static readonly XNamespace W="http://schemas.openxmlformats.org/wordprocessingml/2006/main",M="http://schemas.openxmlformats.org/officeDocument/2006/math";
  static string Val(XElement e)=> (string)e?.Attribute(W+"val");
  static bool On(XElement e)=>e!=null&&!new[]{"0","false","off"}.Contains(Val(e));
  public static List<ParagraphInfo> Read(XDocument document,XDocument styleDocument,CancellationToken cancel=default(CancellationToken)){
   return ReadParagraphs(document.Descendants(W+"body").First().Descendants(W+"p").Where(p=>!p.Ancestors(W+"txbxContent").Any()),styleDocument,cancel);
  }
  static List<ParagraphInfo> ReadParagraphs(IEnumerable<XElement> paragraphs,XDocument styleDocument,CancellationToken cancel){
   var styles=(styleDocument?.Descendants(W+"style")??Enumerable.Empty<XElement>()).Where(x=>x.Attribute(W+"styleId")!=null).GroupBy(x=>(string)x.Attribute(W+"styleId")).ToDictionary(g=>g.Key,g=>g.First());
   var defaults=styleDocument?.Descendants(W+"docDefaults").FirstOrDefault();
   Func<string,string,string,XElement> property=(id,group,name)=>{var visited=new HashSet<string>();while(id!=null&&visited.Add(id)&&styles.TryGetValue(id,out var st)){var v=st.Element(W+group)?.Element(W+name);if(v!=null)return v;id=Val(st.Element(W+"basedOn"));}return defaults?.Element(W+(group+"Default"))?.Element(W+group)?.Element(W+name);};
   var rows=new List<ParagraphInfo>();int i=0;
   foreach(var p in paragraphs){
    cancel.ThrowIfCancellationRequested();
    var pr=p.Element(W+"pPr");string id=Val(pr?.Element(W+"pStyle"))??"Normal";styles.TryGetValue(id,out var st);
    Func<string,XElement> pp=name=>pr?.Element(W+name)??property(id,"pPr",name);
    var text=string.Concat(p.Descendants().Where(x=>!x.Ancestors(W+"txbxContent").Any()&&(x.Name==W+"t"||x.Name==M+"t"||x.Name==W+"tab"||x.Name==W+"br")).Select(x=>x.Name==W+"tab"?"\t":x.Name==W+"br"?"\v":x.Value)).Trim();
    var row=new ParagraphInfo{Paragraph=++i,Start=-1,End=-1,Text=text,Style=Val(st?.Element(W+"name"))??id,OutlineLevel=(int.TryParse(Val(pp("outlineLvl")),out int outline)?outline+1:10),InTable=p.Ancestors(W+"tbl").Any(),Centered=Val(pp("jc"))=="center"};
    row.InToc=p.Annotation<TocContext>()!=null||System.Text.RegularExpressions.Regex.IsMatch(row.Style,@"^(TOC|目录)\s*\d",System.Text.RegularExpressions.RegexOptions.IgnoreCase)||p.Ancestors(W+"sdt").Any(s=>s.Descendants(W+"docPartGallery").Any(g=>Val(g)=="Table of Contents"));
    row.ListLevel=pp("numPr")?.Element(W+"numId")!=null?(int.TryParse(Val(pp("numPr")?.Element(W+"ilvl")),out int ilvl)?ilvl+1:1):0;
    row.SpaceBefore=double.TryParse((string)pp("spacing")?.Attribute(W+"before"),out double space)?space/20:0;
    row.HasMath=p.Descendants(M+"oMath").Any();row.NonMathText=string.Concat(p.Descendants(W+"t").Where(t=>!t.Ancestors(M+"oMath").Any()).Select(t=>t.Value));
    var fontSizes=new Dictionary<double,int>();int bold=0,total=0;
    foreach(var run in p.Descendants(W+"r").Where(r=>!r.Ancestors(M+"oMath").Any()&&!r.Ancestors(W+"txbxContent").Any())){
     int length=string.Concat(run.Elements(W+"t").Select(t=>t.Value)).Trim().Length;if(length==0)continue;
     var rp=run.Element(W+"rPr");string charStyle=Val(rp?.Element(W+"rStyle"));
     var size=rp?.Element(W+"sz")??(charStyle==null?null:property(charStyle,"rPr","sz"))??property(id,"rPr","sz");
     double n=double.TryParse(Val(size),out double half)?half/2:11;if(!fontSizes.ContainsKey(n))fontSizes[n]=0;fontSizes[n]+=length;
     if(On(rp?.Element(W+"b")??(charStyle==null?null:property(charStyle,"rPr","b"))??property(id,"rPr","b")))bold+=length;total+=length;
    }
    row.Size=fontSizes.Count>0?fontSizes.OrderByDescending(k=>k.Value).First().Key:11;row.Bold=total>0&&bold>=total*0.6;rows.Add(row);
   }
   NavigationEngine.MarkDocumentTitles(rows);return rows;
  }
  internal static List<ParagraphInfo> ReadFlat(string xml){
   var pkg=XDocument.Parse(xml);XNamespace p="http://schemas.microsoft.com/office/2006/xmlPackage";
   var parts=pkg.Root.Elements(p+"part").ToDictionary(x=>(string)x.Attribute(p+"name"));
   XElement document=parts["/word/document.xml"].Element(p+"xmlData").Elements().First();
   XElement styles=parts.TryGetValue("/word/styles.xml",out var stylePart)?stylePart.Element(p+"xmlData").Elements().First():null;
   // Revision markup and imported chunks need Word's own displayed-text interpretation.
   if(document.Descendants().Any(x=>x.Name==W+"del"||x.Name==W+"ins"||x.Name==W+"altChunk"))return null;
   return Read(new XDocument(document),styles==null?null:new XDocument(styles));
  }
 }
}
