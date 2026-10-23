using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Forms=System.Windows.Forms;
using PagePilot;
class PhysicalIntegration {
 [System.Runtime.InteropServices.DllImport("user32.dll")] static extern bool SetProcessDPIAware();
 [System.Runtime.InteropServices.DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h);
 static BindingFlags flags=BindingFlags.Instance|BindingFlags.NonPublic|BindingFlags.Public;
 static object Field(object obj,string name)=>obj.GetType().GetField(name,flags).GetValue(obj);
 static object Call(object obj,string name,params object[] args)=>obj.GetType().GetMethods(flags).Single(m=>m.Name==name&&m.GetParameters().Length==args.Length).Invoke(obj,args);
 static T UI<T>(Pane pane,string name) where T:class=>((FrameworkElement)Field(pane,"view")).FindName(name) as T;
 static string Canonical(string xml){var document=System.Xml.Linq.XDocument.Parse(xml);foreach(var e in document.Descendants().Where(e=>e.Name.LocalName=="proofErr"||e.Name.LocalName=="proofState").ToList())e.Remove();return document.ToString();}
 static void Assert(bool value,string name){if(!value)throw new Exception(name);Console.WriteLine("PASS "+name);}
 static async Task Until(Func<bool> check,int seconds=180){var watch=Stopwatch.StartNew();while(!check()){if(watch.Elapsed.TotalSeconds>seconds)throw new Exception("Timeout");await Task.Delay(100);}}
 [STAThread] static void Main(){SetProcessDPIAware();Forms.Application.EnableVisualStyles();var runner=new Forms.Form{Text="PagePilot physical integration",Width=430,Height=850};runner.Shown+=async(s,e)=>{try{await Run(runner);}catch(Exception ex){Console.WriteLine("FAIL "+ex);Environment.ExitCode=1;}finally{runner.Close();}};Forms.Application.Run(runner);}
 static async Task Run(Forms.Form runner){
  dynamic app=Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application"));dynamic doc=null;
  try{
   app.Visible=true;doc=app.Documents.Add();doc.Content.Text="第一章 总则\r本期收入从100增长到120，同比增长10%。\r第一节 范围\rABC 123 测试 ABC\r第二章 说明\r这是正文。\r";
   doc.Content.NoProofing=1;
   var host=new Connect();host.GetType().GetField("App",flags).SetValue(host,app);
   host.GetType().GetField("Dispatcher",flags).SetValue(host,new Forms.Control());
   var pane=new Pane{Dock=Forms.DockStyle.Fill};runner.Controls.Add(pane);Call(pane,"Bind",host,(object)doc,1);
   string before=(string)doc.WordOpenXML;
   UI<TextBox>(pane,"AiRequirement").Text="只识别一级章标题，忽略节和正文。";
   UI<Button>(pane,"Analyze").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
   await Until(()=>Field(pane,"navigationCancel")==null);
   string status=UI<TextBlock>(pane,"Status").Text;Console.WriteLine("AI status: "+status);
   var outline=(List<Heading>)Field(pane,"outline");Assert(status.StartsWith("AI 识别")&&outline.Count==2&&outline[0].Paragraph==1&&outline[1].Paragraph==5,"real model inline navigation exact indices");
   File.WriteAllText(".temp/before.xml",before);File.WriteAllText(".temp/after.xml",(string)doc.WordOpenXML);Assert(Canonical(before)==Canonical((string)doc.WordOpenXML),"navigation preserves document XML");
   UI<Button>(pane,"Analyze").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));UI<Button>(pane,"Analyze").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));await Until(()=>Field(pane,"navigationCancel")==null);Assert(UI<TextBlock>(pane,"Status").Text.Contains("取消"),"navigation cancellation");
   pane.Dispose();pane=new Pane{Dock=Forms.DockStyle.Fill};runner.Controls.Add(pane);Call(pane,"Bind",host,(object)doc,0);
   UI<TextBox>(pane,"Pattern").Text="ABC";UI<TextBox>(pane,"Replacement").Text="XYZ";
   await Until(()=>UI<TextBlock>(pane,"Status").Text.StartsWith("找到"));
   Assert(((List<Hit>)Field(pane,"found")).Count==2,"debounced regex results");
   UI<ListBox>(pane,"Matches").SelectedIndex=1;await Task.Delay(300);Assert((string)app.Selection.Text=="ABC","single-click range navigation");
   string regexBefore=(string)doc.Content.Text;UI<Button>(pane,"Replace").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));Assert(((string)doc.Content.Text).Contains("XYZ")&&!((string)doc.Content.Text).Contains("ABC"),"batch replacement");doc.Undo(1);Assert((string)doc.Content.Text==regexBefore,"single-step undo");
   // Fields invalidate naive string offsets. Exercise a hyperlink before the next hit.
   doc.Hyperlinks.Add(doc.Range(0,3),"https://example.com");UI<TextBox>(pane,"Pattern").Text="AB[C]";
   await Until(()=>UI<TextBlock>(pane,"Status").Text.StartsWith("找到")||UI<TextBlock>(pane,"Status").Text.Contains("映射"));
   Assert(UI<TextBlock>(pane,"Status").Text.StartsWith("找到"),"field-aware regex mapping");
   foreach(var hit in (List<Hit>)Field(pane,"found"))Assert((string)doc.Range(hit.Start,hit.Start+hit.Length).Text==hit.Before,"mapped range content");
   app.Activate();SetForegroundWindow(new IntPtr((int)app.ActiveWindow.Hwnd));doc.Range(0,0).Select();await Task.Delay(600);
   // Paint the actual overlay and inspect its geometry without changing document state.
   var overlay=Field(pane,"highlight");var h=new IntPtr((int)app.ActiveWindow.Hwnd);string xml=(string)doc.WordOpenXML;
   Call(overlay,"Measure",(object)doc,(object)app.ActiveWindow,h,Field(pane,"found"));
   Console.WriteLine("HIGHLIGHT marks="+((System.Collections.ICollection)Field(overlay,"marks")).Count+" foreground="+overlay.GetType().GetMethod("GetForegroundWindow",BindingFlags.Static|BindingFlags.NonPublic).Invoke(null,null)+" owner="+h);foreach(var hit in (List<Hit>)Field(pane,"found")){try{int x=0,y=0,w=0,height=0;app.ActiveWindow.GetPoint(out x,out y,out w,out height,doc.Range(hit.Start,hit.Start+hit.Length));Console.WriteLine("hit "+x+","+y+","+w+","+height);}catch(Exception ex){Console.WriteLine(ex.Message);}}Assert(((System.Collections.ICollection)Field(overlay,"marks")).Count>0,"visible highlight rectangles");Assert((string)doc.WordOpenXML==xml,"highlight preserves XML");
   Console.WriteLine("READY_VISUAL_QA");await Task.Delay(1000);
   pane.Dispose();Console.WriteLine("ALL PHYSICAL TESTS PASSED Word "+app.Version);
  }finally{if(doc!=null)doc.Close(0);app.Quit();}
 }
}
