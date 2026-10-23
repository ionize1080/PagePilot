using System;
using System.Linq;
using System.Text.RegularExpressions;
using PagePilot;
class Tests {
 static int count;
 static void Assert(bool ok,string name){if(!ok)throw new Exception(name);count++;Console.WriteLine("PASS "+name);}
 static void Main(){
  Assert(Engine.Level("第一章 总则",Engine.Defaults)==1,"Chinese chapter");
  Assert(Engine.Level("（二）工作安排",Engine.Defaults)==2,"Chinese subsection");
  Assert(Engine.Level("1.2.3 Scope",Engine.Defaults)==3,"Decimal depth");
  Assert(Engine.Level("普通段落的正文。",Engine.Defaults)==0,"No body headings");
  Assert(Engine.Level(new string('a',161),Engine.Defaults)==0,"Long body excluded");
  var hits=Engine.Find("甲 ABC 123 abc",10,@"(?<word>[a-z]+)","${word}!",true);
  Assert(hits.Count==2&&hits[0].Start==12&&hits[1].After=="abc!","Unicode offset and named replacement");
  Assert(Engine.Find("aaa",0,@"(?=a)","x",false).Count==3,"Zero width matches terminate");
  Assert(Engine.Find("foo\rbar\rfoo\r",0,"^foo$","x",false).Count==2,"Multiline Word paragraph anchors");
  bool rejected=false;try{Engine.Find("a\rb",0,"a\\nb","",false);}catch(InvalidOperationException){rejected=true;}
  Assert(rejected,"Paragraph boundaries protected");
  rejected=false;try{Engine.Find("a\ab",0,"a.b","",false);}catch(InvalidOperationException){rejected=true;}
  Assert(rejected,"Table boundaries protected");
  bool timeout=false;try{Engine.Compile("^(a+)+$").IsMatch(new string('a',100000)+"!");}catch(RegexMatchTimeoutException){timeout=true;}
  Assert(timeout,"Catastrophic regex times out");
  Console.WriteLine(count+" checks passed");
 }
}
