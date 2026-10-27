using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Forms=System.Windows.Forms;
class StartupRegression {
 [System.Runtime.InteropServices.DllImport("user32.dll")]static extern bool SetProcessDPIAware();
 static BindingFlags flags=BindingFlags.Instance|BindingFlags.NonPublic|BindingFlags.Public;static Assembly assembly;static object host,pane;static dynamic app,doc;static Forms.Form form;static List<double> identity=new List<double>();static bool hold;
 static object Field(object o,string n)=>o.GetType().GetField(n,flags).GetValue(o);
 static object Call(object o,string n,params object[] args)=>o.GetType().GetMethods(flags).Single(m=>m.Name==n&&m.GetParameters().Length==args.Length).Invoke(o,args);
 static T UI<T>(string name)where T:class=>((FrameworkElement)Field(pane,"view")).FindName(name) as T;
 static void Assert(bool b,string label){if(!b)throw new Exception(label);Console.WriteLine("PASS "+label);}
 static async Task Until(Func<bool> done){var w=Stopwatch.StartNew();while(!done()){if(w.Elapsed.TotalSeconds>90)throw new Exception("UI wait timeout");await Task.Delay(20);}}
 [STAThread]static void Main(string[] args){SetProcessDPIAware();hold=args.Contains("--hold");assembly=Assembly.LoadFrom(Path.GetFullPath(args[0]));Forms.Application.EnableVisualStyles();form=new Forms.Form{Text="PagePilot · 实际侧栏回归",Width=940,Height=1600,StartPosition=Forms.FormStartPosition.CenterScreen};form.Shown+=async(s,e)=>{try{await Run(args);}catch(Exception ex){Console.WriteLine(ex);Environment.ExitCode=1;}finally{if(!hold)form.Close();}};form.FormClosed+=(s,e)=>{try{if(host!=null&&optimizedHost())Call(host,"Shutdown");if(doc!=null)doc.Close(0);if(app!=null&&(int)app.Documents.Count==0&&(int)app.ProtectedViewWindows.Count==0)app.Quit();}catch{}};Forms.Application.Run(form);}
 static bool optimizedHost()=>host.GetType().GetMethod("Identity",flags)!=null;
 static async Task Run(string[] args){
  app=Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application"));app.Visible=true;doc=app.Documents.Add();doc.SetCompatibilityMode(14);doc.Content.Text=string.Join("\r",Enumerable.Range(1,1000).Select(i=>i%10==1?"第"+i+"章 标题":"第 "+i+" 段 ABC 测试文本"))+"\r";doc.Content.NoProofing=1;
  host=Activator.CreateInstance(assembly.GetType("PagePilot.Connect"));host.GetType().GetField("App",flags).SetValue(host,app);var dispatch=new Forms.Control();var handle=dispatch.Handle;host.GetType().GetField("Dispatcher",flags).SetValue(host,dispatch);
  bool optimized=host.GetType().GetMethod("Identity",flags)!=null;
  for(int i=0;i<30;i++){var watch=Stopwatch.StartNew();if(optimized)Call(host,"Identity");else Call(host,"Context",1,10000);identity.Add(watch.Elapsed.TotalMilliseconds);Console.WriteLine("IDENTITY "+i+" "+watch.Elapsed.TotalMilliseconds);File.WriteAllText(args[1]+".partial",new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(identity));await Task.Delay(10);}
  File.WriteAllText(args[1],new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(new{optimized,paragraphs=1000,samples=identity,p50=identity.OrderBy(x=>x).ElementAt(15),p95=identity.OrderBy(x=>x).ElementAt(28)}));
  pane=Activator.CreateInstance(assembly.GetType("PagePilot.Pane"));((Forms.Control)pane).Dock=Forms.DockStyle.Fill;form.Controls.Add((Forms.Control)pane);var timer=Stopwatch.StartNew();Call(pane,"Bind",host,(object)doc,1);Call(pane,"SelectPage",1);Console.WriteLine("PANE_RETURN_MS "+timer.Elapsed.TotalMilliseconds);await Until(()=>Field(pane,"outlineSnapshot")!=null);Console.WriteLine("NAV_READY_MS "+timer.Elapsed.TotalMilliseconds);Assert(((ICollection)Field(pane,"outline")).Count==100,"long document outline exact");
  if(!optimized){hold=false;return;}
  ((Forms.Control)pane).Dispose();pane=Activator.CreateInstance(assembly.GetType("PagePilot.Pane"));((Forms.Control)pane).Dock=Forms.DockStyle.Fill;form.Controls.Add((Forms.Control)pane);Call(pane,"Bind",host,(object)doc,0);
  doc.Range(0,100).Select();UI<CheckBox>("SelectionOnly").IsChecked=true;UI<TextBox>("Pattern").Text="ABC";UI<TextBox>("Replacement").Text="XYZ";await Until(()=>UI<TextBlock>("Status").Text.StartsWith("找到"));int n=((ICollection)Field(pane,"found")).Count;Assert(n>1,"original selection has multiple matches");UI<ListBox>("Matches").SelectedIndex=0;await Task.Delay(50);UI<TextBox>("Replacement").Text="REPLACED";Assert(((ICollection)Field(pane,"found")).Count==n,"replacement preview preserves original selected range");
  var refresh=new List<double>();for(int i=0;i<30;i++){var w=Stopwatch.StartNew();UI<TextBox>("Replacement").Text="REPLACE"+i;refresh.Add(w.Elapsed.TotalMilliseconds);}Console.WriteLine("REPLACEMENT_PREVIEW_P95_MS "+refresh.OrderBy(x=>x).ElementAt(28));
  UI<TextBox>("Replacement").Text="XYZ";string before=(string)doc.Content.Text;UI<Button>("Replace").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));Assert(((string)doc.Content.Text).Contains("XYZ"),"replace selected matches");doc.Undo(1);Assert((string)doc.Content.Text==before,"one-step undo");
  if(args.Contains("--settings")||args.Contains("--chat")){Call(host,"StartBridge");File.WriteAllText(args[1]+".session",(string)Call(host,"GetMcpSession"));}
  if(args.Contains("--settings")){Call(host,"OpenSettings",new object[]{null});await Until(()=>Convert.ToDouble(Field(host,"SettingsReadyMs"))>=0);Console.WriteLine("SETTINGS_COLD_MS "+Field(host,"SettingsReadyMs"));var times=new List<double>();for(int i=0;i<30;i++){((Process)Field(host,"settingsProcess")).CloseMainWindow();await Task.Delay(50);var w=Stopwatch.StartNew();Call(host,"OpenSettings",new object[]{null});await Until(()=>Convert.ToDouble(Field(host,"SettingsReadyMs"))>=0);times.Add(Convert.ToDouble(Field(host,"SettingsReadyMs")));await Task.Delay(50);}File.WriteAllText(args[1]+".settings.json",new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(times));Console.WriteLine("SETTINGS_WARM_P95_MS "+times.OrderBy(x=>x).ElementAt(28));}
  if(args.Contains("--chat")){
   var runs=new List<object>();int cycles=args.Contains("--bench")?10:1;
   for(int i=0;i<cycles;i++){Call(host,"LaunchChat");var metrics=(Dictionary<string,double>)Field(host,"AssistantMetrics");await Until(()=>metrics.ContainsKey("input_ready")&&metrics.ContainsKey("mcp_ready"));runs.Add(new Dictionary<string,double>(metrics));Console.WriteLine("CHAT "+i+" "+new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(metrics));File.WriteAllText(args[1]+".chat.json",new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(runs));await Task.Delay(2500);if(i<cycles-1){var process=(Process)Field(host,"chat");process.CloseMainWindow();await Until(()=>process.HasExited);await Task.Delay(500);}}
   var warm=new List<double>();for(int i=0;i<30;i++){var w=Stopwatch.StartNew();Call(host,"LaunchChat");warm.Add(w.Elapsed.TotalMilliseconds);await Task.Delay(30);}File.WriteAllText(args[1]+".warm.json",new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(warm));Console.WriteLine("WARM_P95_MS "+warm.OrderBy(x=>x).ElementAt(28));
  }
  Console.WriteLine("READY_VISUAL_QA");
 }
}
