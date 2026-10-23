using System;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using PagePilot;
using F=System.Windows.Forms;
class ComplexSearch{
 [System.Runtime.InteropServices.DllImport("user32.dll")] static extern bool SetProcessDPIAware();
 [System.Runtime.InteropServices.DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h);
 static BindingFlags flags=BindingFlags.Instance|BindingFlags.NonPublic|BindingFlags.Public;
 [STAThread] static void Main(string[] args){SetProcessDPIAware();var form=new F.Form{Width=650,Height=900};form.Shown+=async(s,e)=>{
 dynamic app=null,doc=null;try{
 app=Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application"));app.Visible=true;doc=app.Documents.Open(args[0],ReadOnly:true);
 var host=new Connect();typeof(Connect).GetField("App",flags).SetValue(host,app);
 var pane=new Pane{Dock=F.DockStyle.Fill};form.Controls.Add(pane);typeof(Pane).GetMethod("Bind",flags).Invoke(pane,new object[]{host,doc,0});
 var view=(FrameworkElement)typeof(Pane).GetField("view",flags).GetValue(pane);
 ((TextBox)view.FindName("Pattern")).Text=@"QA|\d+";
 await Task.Delay(4500);string status=((TextBlock)view.FindName("Status")).Text;Console.WriteLine(status);if(!status.StartsWith("找到"))throw new Exception(status);
 foreach(var hit in (System.Collections.Generic.List<Hit>)typeof(Pane).GetField("found",flags).GetValue(pane)){if((string)doc.Range(hit.Start,hit.Start+hit.Length).Text!=hit.Before)throw new Exception("Invalid mapped text");if(hit.Length!=hit.Before.Length)throw new Exception("Hidden field boundary included");}
 string before=(string)doc.Content.Text;SetForegroundWindow(new IntPtr((int)app.ActiveWindow.Hwnd));
 await Task.Delay(1500);var overlay=typeof(Pane).GetField("highlight",flags).GetValue(pane);
 overlay.GetType().GetMethod("Measure",flags).Invoke(overlay,new object[]{doc,app.ActiveWindow,new IntPtr((int)app.ActiveWindow.Hwnd),typeof(Pane).GetField("found",flags).GetValue(pane)});
 int marks=((System.Collections.ICollection)overlay.GetType().GetField("marks",flags).GetValue(overlay)).Count;
 Console.WriteLine("marks="+marks);if(marks==0)throw new Exception("No highlights");
 Console.WriteLine("READY_COMPLEX_VISUAL");await Task.Delay(3000);if((string)doc.Content.Text!=before)throw new Exception("Highlight modified text");pane.Dispose();
 Console.WriteLine("PASS complex search/highlight");
 }catch(Exception ex){Console.WriteLine(ex);Environment.ExitCode=1;}finally{if(doc!=null)doc.Close(0);if(app!=null)app.Quit();form.Close();}};F.Application.Run(form);}
}
