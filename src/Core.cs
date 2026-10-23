using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
namespace PagePilot {
 public sealed class HeadingRule { public string Name {get;set;} public string Pattern {get;set;} public int Level {get;set;} }
 public sealed class Heading { public int Paragraph {get;set;} public int Start {get;set;} public int End {get;set;} public int Level {get;set;} public string Text {get;set;} public override string ToString()=>new string(' ',(Level-1)*2)+Text; }
 public sealed class Hit { public int Start {get;set;} public int Length {get;set;} public string Before {get;set;} public string After {get;set;} public override string ToString()=> $"{Start}: {Before.Replace("\r","↵")} → {After.Replace("\r","↵")}"; }
 public static class Engine {
  public static readonly HeadingRule[] Defaults = {
   new HeadingRule{Name="中文章",Pattern=@"^第[一二三四五六七八九十百千万零〇两0-9]+[编篇章]\s*\S.*$",Level=1},
   new HeadingRule{Name="中文节",Pattern=@"^第[一二三四五六七八九十百千万零〇两0-9]+节\s*\S.*$",Level=2},
   new HeadingRule{Name="中文一级",Pattern=@"^[一二三四五六七八九十百]+[、．.]\s*\S.*$",Level=1},
   new HeadingRule{Name="中文二级",Pattern=@"^[（(][一二三四五六七八九十百]+[）)]\s*\S.*$",Level=2},
   new HeadingRule{Name="多级数字",Pattern=@"^\d+(?:\.\d+){1,5}[.、]?\s+\S.*$",Level=0},
   new HeadingRule{Name="数字一级",Pattern=@"^\d+[.、]\s*[^\d\s].*$",Level=1},
   new HeadingRule{Name="数字三级",Pattern=@"^[（(]\d+[）)]\s*\S.*$",Level=3},
   new HeadingRule{Name="英文章",Pattern=@"^(?:Chapter|Part|Appendix)\s+[A-Z0-9IVXLCDM]+\b.*$",Level=1}
  };
  public static Regex Compile(string pattern, bool ignoreCase=false) => new Regex(pattern, RegexOptions.Multiline | (ignoreCase?RegexOptions.IgnoreCase:RegexOptions.None), TimeSpan.FromMilliseconds(500));
  public static int Level(string text,IEnumerable<HeadingRule> rules) {
   text=text.Trim(); if(text.Length==0 || text.Length>160)return 0;
   foreach(var r in rules) if(Compile(r.Pattern,true).IsMatch(text)) return r.Level==0? Math.Min(9,text.TakeWhile(c=>char.IsDigit(c)||c=='.').Count(c=>c=='.')+1):Math.Max(1,Math.Min(9,r.Level));
   return 0;
  }
  public static List<Hit> Find(string text,int offset,string pattern,string replacement,bool ignoreCase) {
   var results=new List<Hit>();
   foreach(Match m in Compile(pattern,ignoreCase).Matches(text.Replace('\r','\n'))) {
    if(results.Count>=10000)throw new InvalidOperationException("匹配超过 10000 项，请缩小范围。");
    if(m.Value.Contains('\a') || m.Value.Contains('\n') || m.Value.Contains('\f'))throw new InvalidOperationException("不支持跨段落或单元格边界替换，请按段落匹配。");
    var after=m.Result(replacement);
    if(after.IndexOfAny(new[]{'\a','\r','\n','\f'})>=0)throw new InvalidOperationException("替换文本不能含段落或单元格控制符。");
    results.Add(new Hit{Start=offset+m.Index,Length=m.Length,Before=m.Value,After=after});
   }
   return results;
  }
 }
}
