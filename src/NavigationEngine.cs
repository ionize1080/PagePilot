using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace PagePilot {
 public sealed class ParagraphInfo {
  public int Paragraph, Start, End, OutlineLevel, ListLevel;
  public string Text, Style, Number, NonMathText;
  public double Size, SpaceBefore;
  public bool Bold, Centered, InTable, InToc, HasMath, DocumentTitle;
 }
 // Pure document analysis: no COM objects cross onto the worker thread.
 public static class NavigationEngine {
  const string Cn="零〇一二三四五六七八九十百千万两壹贰叁肆伍陆柒捌玖拾佰仟";
  static readonly Regex DateOrValue=Rx(@"^(?:\d{4}\s*(?:[-/.]\d{1,2}(?:[-/.]\d{1,2})?(?:\s|$)|年)|\d+(?:\.\d+)?\s*(?:%|％|元|万元|亿元|kg\b|km\b)|\d+\.\d+\.\d+\.\d+(?:\s|$))");
  static Regex Rx(string s)=>new Regex(s,RegexOptions.IgnoreCase|RegexOptions.CultureInvariant,TimeSpan.FromMilliseconds(100));
  static readonly Regex ChineseChapter=Rx("^第["+Cn+@"0-9]+\s*([编篇部章节卷])(?:分)?(?:\s*[:：、.．]?\s*.*)$");
  static readonly Regex EnglishChapter=Rx(@"^(part|book|volume|chapter|section|appendix)\s+(?:\d+|[ivxlcdm]+|[a-z]|one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve)(?:\s|[.:：、-]|$)");
  static readonly Regex Decimal=Rx(@"^(\d{1,5}(?:\.\d{1,5}){1,8})(?:[.、:：]?\s*[^\d\s.]|\s+\S)");
  static readonly Regex Arabic=Rx(@"^\d{1,5}(?:[.、:：)]\s*|\s+)\D");
  static readonly Regex Chinese=Rx("^["+Cn+@"]+[、.．:：]\s*\S");
  static readonly Regex ChineseParen=Rx("^[（(]["+Cn+@"]+[）)]\s*\S");
  static readonly Regex ArabicParen=Rx(@"^[（(]\d{1,5}[）)]\s*\S");
  static readonly Regex Roman=Rx(@"^([IVXLCDM]+)(?:[.、:：)]\s*|\s+)\S|^[（(]([IVXLCDM]+)[）)]\s*\S");
  static readonly Regex Letter=Rx(@"^(?:[A-Z][.、:：)]\s*|[（(][A-Z][）)]\s*)\S");
  static readonly Regex Circle=Rx(@"^[①-⑳❶-❿]\s*\S");
  static readonly Regex RomanToken=Rx(@"^(?=[IVXLCDM]+$)M{0,4}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})$");
  static readonly Regex EnglishNumber=Rx(@"^(?:one|two|three|four|five|six|seven|eight|nine|ten|first|second|third|fourth|fifth|sixth|seventh|eighth|ninth|tenth)[.:]\s+\S");
  static readonly Regex Special=Rx(@"^(?:摘要|引言|前言|绪论|结语|结论|总结|致谢|参考文献|附录(?:\s*[A-Z一二三0-9])?|abstract|introduction|preface|conclusion(?:s)?|references|bibliography|acknowledg(?:e)?ments)(?:\s*[:：].*)?$");
  static readonly Regex StyleHeading=Rx(@"^(?:heading|标题)\s*([1-9])(?:\b|$)");
  static readonly Regex SentenceEnd=Rx(@"[。；;!?！？]$|[。；;！？].{8,}$");
  static readonly Regex TitleStyle=Rx(@"^(?:(?:doc(?:ument)?\s+)?(?:title|subtitle)(?:\s*\(.*\))?|副标题|标题|文档标题|主标题)$");
  static readonly Regex DateLine=Rx(@"^(?:\d{4}\s*[–—-]\s*\d{2,4}|\d{1,2}\s+(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4}|(?:January|February|March|April|May|June|July|August|September|October|November|December),?\s+\d{4})$");
  static readonly Regex Caption=Rx(@"^(?:(?:图|表|公式)\s*\d+(?:[.－-]\d+)*|(?:figure|table|equation)\s+\d+)(?:\s|[:：.、]|$)");
  static readonly Regex StructuralExtras=Rx(@"^(?:摘\s*要|目\s*录|学位论文.*声明|foreword|forward|intro|contents|cyflwyniad)$");
  static readonly Regex BodyStyle=Rx(@"body\s*text");
  static readonly Regex CoverLabel=Rx(@"^(?:[中英]文(?:论文)?题[目名]|论文题名|[博硕]\s*士\s*学\s*位\s*论\s*文|指导教师\s*[:：]|candidate\s*[:：]|supervisor\s*[:：]|完\s*成\s*日\s*期)");
  static readonly Regex Formula=Rx(@"[=≠≈≤≥∑∏∫√∞]|[α-ωΑ-Ω].*[+−*/^]");
  public static bool Excluded(ParagraphInfo p,bool includeTitle=false){
   if(p.InToc||string.IsNullOrWhiteSpace(p.Text))return true;
   if(!includeTitle&&(p.DocumentTitle||TitleStyle.IsMatch((p.Style??"").Trim())))return true;
   if(DateLine.IsMatch(p.Text.Trim())||Caption.IsMatch(p.Text.Trim()))return true;
   if(p.Text.Length>180&&SentenceEnd.IsMatch(p.Text.Trim()))return true;
   string prose=p.HasMath?p.NonMathText??"":p.Text;
   // OMML is structural evidence; an inline equation inside a real heading is retained.
   if(p.HasMath&&!Regex.IsMatch(prose,@"[\p{IsCJKUnifiedIdeographs}A-Za-z]{2}"))return true;
   string t=p.Text.Trim();
   if(t.Length<=2&&!Regex.IsMatch(t,@"[\p{IsCJKUnifiedIdeographs}]{2}"))return true;
   if(Formula.IsMatch(t)&&!Regex.IsMatch(prose,@"[\p{IsCJKUnifiedIdeographs}]{2}|[A-Za-z]{3,}"))return true;
   return false;
  }
  static bool Structural(string text){string t=text.Normalize(NormalizationForm.FormKC).Trim();return !DateOrValue.IsMatch(t)&&(ChineseChapter.IsMatch(t)||EnglishChapter.IsMatch(t)||Chinese.IsMatch(t)||Decimal.IsMatch(t)||Arabic.IsMatch(t)||Special.IsMatch(t)||StructuralExtras.IsMatch(t));}
  public static void MarkDocumentTitles(IReadOnlyList<ParagraphInfo> rows){
   foreach(var p in rows.Where(p=>p.Paragraph>0&&p.Paragraph<=150&&p.Size>=14&&CoverLabel.IsMatch(p.Text??"")))p.DocumentTitle=true;
   // Cover material precedes the first section or substantive body paragraph. Do not
   // infer a cover from paragraph index alone (documents can start with a real heading).
   var first=rows.FirstOrDefault(p=>!string.IsNullOrWhiteSpace(p.Text)&&!p.InTable);if(first==null||first.Paragraph<=0||Structural(first.Text))return;
   var boundary=rows.FirstOrDefault(p=>p.Paragraph>first.Paragraph&&!p.InTable&&(Structural(p.Text)||(p.Text.Length>140&&!p.Centered)));
   if(boundary==null||boundary.Paragraph-first.Paragraph>150)return;
   var lead=rows.Where(p=>p.Paragraph>=first.Paragraph&&p.Paragraph<boundary.Paragraph&&!p.InTable&&!string.IsNullOrWhiteSpace(p.Text)).ToList();
   if(lead.Count==0||!lead.Any(p=>p.Bold||p.Centered||p.Size>=14||TitleStyle.IsMatch(p.Style??"")))return;
   foreach(var p in lead.Where(p=>p.Text.Length<=140&&!(p.OutlineLevel>=2&&p.OutlineLevel<=9&&StyleHeading.IsMatch(p.Style??"")&&p.Paragraph>first.Paragraph+1)))p.DocumentTitle=true;
  }
  static bool Short(string text)=>text.Length>0&&text.Length<=140&&!SentenceEnd.IsMatch(text)&&!DateOrValue.IsMatch(text);
  sealed class Candidate {public ParagraphInfo P;public int Level;public string Family;public bool Fixed;}
  public static List<Heading> Analyze(IReadOnlyList<ParagraphInfo> rows,IEnumerable<HeadingRule> rules,System.Threading.CancellationToken token=default(System.Threading.CancellationToken),NavigationAppearance options=null) {
   var custom=rules.Where(r=>r.Name=="自定义").ToArray();
   MarkDocumentTitles(rows);
   var body=rows.Where(p=>p.Size>0&&p.Size<100&&p.Text.Length>140).ToList();
   double bodySize=body.Count==0?11:body.GroupBy(p=>p.Size).OrderByDescending(g=>g.Count()).First().Key;
   var sizes=rows.Where(p=>!p.InTable&&!Excluded(p,options?.IncludeDocumentTitle==true)&&Short(p.Text)&&p.Size>bodySize+1&&p.Bold).Select(p=>p.Size).Distinct().OrderByDescending(x=>x).Take(9).ToList();
   var candidates=new List<Candidate>();
   foreach(var p in rows) {
    token.ThrowIfCancellationRequested();if(Excluded(p,options?.IncludeDocumentTitle==true))continue;
    string t=p.Text.Normalize(NormalizationForm.FormKC).Trim();int level=Engine.Level(t,custom);string family=null;bool fixedLevel=level>0;
    if(level==0&&p.OutlineLevel>=1&&p.OutlineLevel<=9&&(!p.InTable||StyleHeading.IsMatch(p.Style??""))&&!(t.Length>180&&SentenceEnd.IsMatch(t))){level=p.OutlineLevel;fixedLevel=true;}
    var style=StyleHeading.Match((p.Style??"").Normalize(NormalizationForm.FormKC));
    if(level==0&&style.Success){level=int.Parse(style.Groups[1].Value);fixedLevel=true;}
    // Table cells and TOC entries are not implicit headings. Explicit styles still win.
    if(level==0&&!p.InTable&&Short(t)) {
     var cn=ChineseChapter.Match(t);var en=EnglishChapter.Match(t);var number=Decimal.Match(t);
     if(cn.Success){family="cn-"+cn.Groups[1].Value;level=cn.Groups[1].Value=="节"?2:1;}
     else if(en.Success){family="en-"+en.Groups[1].Value.ToLowerInvariant();level=family=="en-section"?2:1;}
     else if(number.Success){level=number.Groups[1].Value.Count(c=>c=='.')+1;family="decimal";fixedLevel=true;}
     else if(Chinese.IsMatch(t)){family="chinese";level=1;}
     else if(ChineseParen.IsMatch(t)){family="chinese-paren";level=2;}
     else if(Roman.IsMatch(t)&&RomanToken.IsMatch(t.TrimStart('(','（').Split(new[]{' ','.','、',':','：',')','）','\t'})[0])){family=char.IsLower(t.TrimStart('(','（')[0])?"roman-lower":"roman";level=1;}
     else if(Letter.IsMatch(t)){family=char.IsLower(t.TrimStart('(','（')[0])?"letter-lower":"letter";level=2;}
     else if(ArabicParen.IsMatch(t)){family="arabic-paren";level=4;}
     else if(Arabic.IsMatch(t)||EnglishNumber.IsMatch(t)||Circle.IsMatch(p.Text)){family="arabic";level=1;}
     else if(Special.IsMatch(t)){level=1;fixedLevel=true;}
     // Formatting is evidence only when supported by spacing/centering or a custom style.
     else if(options?.InferTypography!=false&&!BodyStyle.IsMatch(p.Style??"")&&!t.StartsWith("—")&&p.Size>bodySize+1&&p.Bold) {level=Math.Max(1,sizes.IndexOf(p.Size)+1);fixedLevel=true;}
     else if(options?.InferTypography!=false&&p.Bold&&!p.Centered&&p.SpaceBefore>=6&&t.Length<=80){level=2;fixedLevel=true;}
     if(p.ListLevel>0&&level==0&&Short(p.Number+" "+t)&&(p.Bold||p.SpaceBefore>=6)) {level=p.ListLevel;fixedLevel=true;}
     // Ordinary list items are intentionally conservative; styles/outline remain authoritative.
     if(p.ListLevel>0&&!p.Bold&&p.Size<=bodySize+1&&p.SpaceBefore<6&&!fixedLevel){level=0;family=null;}
    }
    if(level>0)candidates.Add(new Candidate{P=p,Level=level,Family=family,Fixed=fixedLevel});
   }
   bool chapter=candidates.Any(c=>c.Family=="cn-章"||c.Family=="en-chapter");
   bool part=candidates.Any(c=>new[]{"cn-编","cn-篇","cn-部","cn-卷","en-part","en-book","en-volume"}.Contains(c.Family));
   bool chinese=candidates.Any(c=>c.Family=="chinese"),roman=candidates.Any(c=>c.Family=="roman");
   bool letter=candidates.Any(c=>c.Family=="letter");
   int prefix=(part?1:0)+(chapter?1:0);
   foreach(var c in candidates.Where(c=>!c.Fixed)){
    switch(c.Family){
     case "cn-章":case "en-chapter":c.Level=part?2:1;break;
     case "cn-节":case "en-section":c.Level=prefix+1;break;
     case "chinese":c.Level=prefix+1;break;
     case "chinese-paren":c.Level=prefix+(chinese?2:1);break;
     case "arabic":c.Level=prefix+(chinese?3:roman?(letter?3:2):1);break;
     case "arabic-paren":c.Level=prefix+(chinese?4:2);break;
     case "roman":c.Level=prefix+1;break;
     case "letter":c.Level=prefix+(roman?2:1);break;
     case "roman-lower":c.Level=prefix+(roman?4:1);break;
     case "letter-lower":c.Level=prefix+(roman?3:2);break;
    }
   }
   return candidates.Where(c=>c.Level<=(options?.MaxLevel??9)).Select(c=>new Heading{Paragraph=c.P.Paragraph,Start=c.P.Start,End=c.P.End,Text=string.IsNullOrEmpty(c.P.Number)?c.P.Text:c.P.Number+" "+c.P.Text,Level=Math.Min(9,Math.Max(1,c.Level))}).ToList();
  }
 }
 public sealed class NavigationLevelStyle {
  public string Font {get;set;}="Microsoft YaHei UI";
  public double Size {get;set;}=13;
  public bool Bold {get;set;}
  public string Color {get;set;}="#334155";
 }
 public sealed class NavigationAppearance {
  public string VisualStyle {get;set;}="colorful";
  public NavigationAppearance Effective(){
   if(VisualStyle!="acrobat")return this;
   return new NavigationAppearance{VisualStyle="acrobat",Background="#F5F5F5",Accent="#1473E6",HighlightColor=HighlightColor,
    SelectionBackground="#E1EFFF",SelectionForeground="#2C2C2C",RowPadding=5,Indent=18,MaxLines=MaxLines,
    Levels=Enumerable.Range(0,9).Select(i=>new NavigationLevelStyle{Font=Levels[i].Font,Size=13,Bold=false,Color="#4B4B4B"}).ToArray()};
  }
  public string Background {get;set;}="#F8FAFC";
  public string Preset {get;set;}="light";
  public string Accent {get;set;}="#0F766E";
  public string HighlightColor {get;set;}="#B8780A";
  public string SelectionBackground {get;set;}="#DDF3ED";
  public string SelectionForeground {get;set;}="#115E59";
  public int RowPadding {get;set;}=8;
  public int Indent {get;set;}=20;
  public int MaxLines {get;set;}=2;
  public int MaxLevel {get;set;}=9;
  public bool IncludeDocumentTitle {get;set;}=false;
  public bool InferTypography {get;set;}=true;
  public NavigationLevelStyle[] Levels {get;set;}=Enumerable.Range(1,9).Select(i=>new NavigationLevelStyle{Size=i==1?15:13,Bold=i==1,Color=i==1?"#1E293B":"#475569"}).ToArray();
  public void Validate(){
   if(VisualStyle!="colorful"&&VisualStyle!="acrobat")throw new ArgumentException("请选择多彩或 Acrobat 专业风格。");
   foreach(var color in new[]{Accent,HighlightColor,SelectionBackground,SelectionForeground})if(!Regex.IsMatch(color??"",@"^#[0-9a-fA-F]{6}$"))throw new ArgumentException("外观颜色应为 #RRGGBB。");
   if(RowPadding<3||RowPadding>16||Indent<12||Indent>32||MaxLines<1||MaxLines>4||MaxLevel<1||MaxLevel>9)throw new ArgumentException("目录间距、行数或显示级数无效。");
   if(!Regex.IsMatch(Background??"",@"^#[0-9a-fA-F]{6}$")||Levels==null||Levels.Length!=9)throw new ArgumentException("导航外观需要背景颜色与九级样式。");
   foreach(var s in Levels)if(s==null||string.IsNullOrWhiteSpace(s.Font)||s.Font.Length>100||double.IsNaN(s.Size)||s.Size<9||s.Size>32||!Regex.IsMatch(s.Color??"",@"^#[0-9a-fA-F]{6}$"))throw new ArgumentException("字体大小应为 9–32，颜色应为 #RRGGBB。");
  }
 }
}
