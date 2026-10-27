using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Forms=System.Windows.Forms;
using PagePilot;
class NavigationIntegration {
 [System.Runtime.InteropServices.DllImport("user32.dll")] static extern bool SetProcessDPIAware();
 static BindingFlags flags=BindingFlags.Instance|BindingFlags.NonPublic|BindingFlags.Public;
 static object Field(object obj,string name)=>obj.GetType().GetField(name,flags).GetValue(obj);
 static object Call(object obj,string name,params object[] args)=>obj.GetType().GetMethods(flags).Single(m=>m.Name==name&&m.GetParameters().Length==args.Length).Invoke(obj,args);
 static T UI<T>(Pane pane,string name) where T:class=>((FrameworkElement)Field(pane,"view")).FindName(name) as T;
 static void Assert(bool value,string name){if(!value)throw new Exception(name);Console.WriteLine("PASS "+name);}
 static async Task Until(Func<bool> check){var watch=Stopwatch.StartNew();while(!check()){if(watch.Elapsed.TotalSeconds>90)throw new Exception("Timeout");await Task.Delay(20);}}
 [STAThread] static void Main(string[] args){SetProcessDPIAware();Forms.Application.EnableVisualStyles();var runner=new Forms.Form{Text="PagePilot navigation QA",Width=420,Height=850};runner.Shown+=async(s,e)=>{try{await Run(runner,args);}catch(Exception ex){Console.WriteLine("FAIL "+ex);Environment.ExitCode=1;}finally{runner.Close();}};Forms.Application.Run(runner);}
 static async Task Run(Forms.Form runner,string[] args){
  dynamic app=Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application"));dynamic doc=null;bool owned=(int)app.Documents.Count==0;
  if(!owned)throw new Exception("Test requires an isolated Word instance");
  try{
   app.Visible=true;doc=app.Documents.Open(Path.GetFullPath(args[0]),ReadOnly:true,AddToRecentFiles:false);
   var host=new Connect();host.GetType().GetField("App",flags).SetValue(host,app);host.GetType().GetField("Dispatcher",flags).SetValue(host,new Forms.Control());
   var pane=new Pane{Dock=Forms.DockStyle.Fill};runner.Controls.Add(pane);var watch=Stopwatch.StartNew();Call(pane,"Bind",host,(object)doc,1);Console.WriteLine("Pane bind ms="+watch.ElapsedMilliseconds);
   Assert(!(bool)Field(pane,"scanning")&&((List<Heading>)Field(pane,"outline")).Count==0,"Binding does not scan");
   var before=(string)doc.Content.Text;watch.Restart();pane.SelectPage(1);await Until(()=>Field(pane,"outlineSnapshot")!=null||UI<TextBlock>(pane,"Status").Text.Contains("失败"));
   Console.WriteLine("Scan ms="+watch.ElapsedMilliseconds+" "+UI<TextBlock>(pane,"Status").Text);var outline=(List<Heading>)Field(pane,"outline");
   foreach(var h in outline)Console.WriteLine("OUTLINE "+h.Paragraph+" L"+h.Level+" "+h.Text);
   bool riemann=Path.GetFileName(args[0]).Contains("黎曼");Assert(riemann?outline.Count==8&&outline.All(h=>h.Level==1):Enumerable.Range(1,8).All(i=>outline.Any(h=>h.Text.StartsWith(i+" ")&&h.Level==1)),"Live Word eight chapter headings; no formulas or cover title");
   Assert((string)doc.Content.Text==before&&(bool)doc.Saved,"Read-only analysis preserves document");
   var target=outline[4];Call(pane,"Navigate",target);Assert((int)app.Selection.Start==target.Start,"Actual Word range navigation");
   var view=(FrameworkElement)Field(pane,"view");view.UpdateLayout();var bitmap=new RenderTargetBitmap((int)view.ActualWidth,(int)view.ActualHeight,96,96,PixelFormats.Pbgra32);bitmap.Render(view);var png=new PngBitmapEncoder();png.Frames.Add(BitmapFrame.Create(bitmap));using(var stream=File.Create(".temp/navigation-v046.png"))png.Save(stream);
   var selected=(TreeViewItem)UI<TreeView>(pane,"Headings").Items[4];selected.IsSelected=true;await Task.Delay(50);
   Assert(((SolidColorBrush)((TextBlock)selected.Header).Foreground).Color==(Color)ColorConverter.ConvertFromString(((NavigationAppearance)Field(host,"NavigationStyle")).SelectionForeground),"Selection foreground follows accessible palette");
   view.UpdateLayout();bitmap=new RenderTargetBitmap((int)view.ActualWidth,(int)view.ActualHeight,96,96,PixelFormats.Pbgra32);bitmap.Render(view);png=new PngBitmapEncoder();png.Frames.Add(BitmapFrame.Create(bitmap));using(var stream=File.Create(".temp/navigation-selected-v046.png"))png.Save(stream);
   pane.Dispose();doc.Close(0);doc=app.Documents.Add();doc.Content.Text="I. Overview\rA. Scope\r1. Background\ri. Detail\rII. Methods\rA. Experiments\r";
   pane=new Pane{Dock=Forms.DockStyle.Fill};runner.Controls.Add(pane);Call(pane,"Bind",host,(object)doc,1);pane.SelectPage(1);await Until(()=>Field(pane,"outlineSnapshot")!=null);
   var tree=UI<TreeView>(pane,"Headings");var root=(TreeViewItem)tree.Items[0];Assert(root.Items.Count==1&&((TreeViewItem)root.Items[0]).Tag==null,"Children initially unloaded");root.IsExpanded=true;await Task.Delay(50);Assert(((TreeViewItem)root.Items[0]).Tag is Heading,"Expansion materializes child headings");
   var appearance=(NavigationAppearance)Field(host,"NavigationStyle");appearance.Background="#E8F1EF";appearance.Levels[0].Font="Arial";appearance.Levels[0].Size=20;appearance.Levels[0].Color="#225544";host.GetType().GetField("NavigationStyleVersion",flags).SetValue(host,1);Call(pane,"ApplyNavigationAppearance");
   Assert(((TextBlock)root.Header).FontSize==20&&((TextBlock)root.Header).FontFamily.Source=="Arial","Appearance applies to materialized nodes");
   doc.Content.Text=string.Join("\r",Enumerable.Range(0,3000).Select(i=>i%30==0?"1. Heading":"A plain body paragraph."))+"\r";
   Call(pane,"RefreshOutline");await Task.Delay(30);UI<Button>(pane,"StopScan").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));await Until(()=>!(bool)Field(pane,"scanning"));Assert(UI<TextBlock>(pane,"Status").Text.Contains("取消"),"Live scan cancellation responsive");
   pane.Dispose();Console.WriteLine("ALL NAVIGATION INTEGRATION CHECKS PASSED Word "+app.Version);
  }finally{if(doc!=null)doc.Close(0);if(owned)app.Quit();}
 }
}
