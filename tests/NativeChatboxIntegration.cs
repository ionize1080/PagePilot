using System;
using System.IO;
using System.Reflection;
using System.Diagnostics;
using System.Threading.Tasks;
using F=System.Windows.Forms;
class NativeChatboxIntegration {
 static BindingFlags Flags=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
 [STAThread]static void Main(string[] args){
  var form=new F.Form{ShowInTaskbar=false,WindowState=F.FormWindowState.Minimized};
  form.Shown+=async(s,e)=>{dynamic word=null,doc=null;object host=null;Process process=null;try{
   string root=Path.GetFullPath(args[0]),qa=Path.Combine(root,".temp","native-chatbox");Directory.CreateDirectory(qa);
   var assembly=Assembly.LoadFrom(Path.Combine(root,"dist","PagePilot.dll"));
   var configType=assembly.GetType("PagePilot.Configuration.RuntimeSettings");object config=Activator.CreateInstance(configType);
   string profile=Path.Combine(qa,"profile-"+Process.GetCurrentProcess().Id);
   foreach(var pair in new[]{new[]{"DataDirectory","data"},new[]{"ProfileDirectory",Path.GetFileName(profile)},new[]{"LogDirectory","logs"}})configType.GetProperty(pair[0]).SetValue(config,Path.Combine(qa,pair[1]));
   assembly.GetType("PagePilot.Configuration.RuntimePaths").GetMethod("Set").Invoke(null,new[]{config,(object)Path.Combine(qa,"runtime.json")});
   word=Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application"));word.Visible=false;word.DisplayAlerts=0;word.AutomationSecurity=3;
   doc=word.Documents.Add();doc.Content.Text="Synthetic research report. Source references are in the PagePilot workspace.\r";
   dynamic end=doc.Range((int)doc.Content.End-1,(int)doc.Content.End-1);
   doc.InlineShapes.AddOLEObject(ClassType:"Package",FileName:Path.Combine(root,".temp","workspace-tests","references","text.pdf"),LinkToFile:false,DisplayAsIcon:true,Range:end);
   doc.SaveAs2(Path.Combine(qa,"report-qa.docx"),16);
   var type=assembly.GetType("PagePilot.Connect");host=Activator.CreateInstance(type);type.GetField("App",Flags).SetValue(host,word);type.GetField("Current",Flags).SetValue(null,host);
   var control=new F.Control();var hwnd=control.Handle;type.GetField("Dispatcher",Flags).SetValue(host,control);
   type.GetField("AutoApprove",Flags).SetValue(host,true);type.GetField("approvalDocument",Flags).SetValue(host,doc);type.GetMethod("StartBridge",Flags).Invoke(host,null);
   var start=new ProcessStartInfo(Path.Combine(root,"dist","Browser","electron.exe"),"\""+Path.Combine(root,"tests","native-chatbox.cjs")+"\""){UseShellExecute=false,CreateNoWindow=true,WindowStyle=ProcessWindowStyle.Hidden,WorkingDirectory=root};
   start.EnvironmentVariables["PAGEPILOT_PIPE"]=(string)type.GetField("PipeName",Flags).GetValue(host);start.EnvironmentVariables["PAGEPILOT_TOKEN"]=(string)type.GetField("Token",Flags).GetValue(host);
   start.EnvironmentVariables["PAGEPILOT_ROOT"]=Path.Combine(root,"dist");start.EnvironmentVariables["PAGEPILOT_PROFILE"]=profile;start.EnvironmentVariables["PAGEPILOT_DATA"]=Path.Combine(qa,"data");start.EnvironmentVariables["PAGEPILOT_LOGS"]=Path.Combine(qa,"logs");start.EnvironmentVariables["PAGEPILOT_QA_KEY_PATH"]=Path.GetFullPath(args[1]);
   process=Process.Start(start);var deadline=DateTime.UtcNow.AddMinutes(8);
   while(!process.HasExited){if(DateTime.UtcNow>deadline){process.Kill();throw new Exception("Actual Chatbox QA timed out");}await Task.Delay(100);}
   if(process.ExitCode!=0)throw new Exception("Actual Chatbox QA exit "+process.ExitCode);
   string text=(string)doc.Content.Text;File.WriteAllText(Path.Combine(qa,"word-result.txt"),text);
   if(!text.Contains("50")||!text.Contains("45")||!text.Contains("Synthetic research report")||text.Length<120||(int)doc.InlineShapes.Count!=1)throw new Exception("Analysis conclusion or original content not verified in real Word");
   doc.Save();Console.WriteLine("PASS actual Chatbox streaming Qwen + workspace analysis + approved MCP writes to real Word");
  }catch(Exception ex){Console.Error.WriteLine(ex);Environment.ExitCode=1;}finally{if(host!=null)try{host.GetType().GetMethod("Shutdown",Flags).Invoke(host,null);}catch{}if(doc!=null)try{doc.Close(0);}catch{}if(word!=null)try{word.Quit(0);}catch{}form.Close();}};
  F.Application.Run(form);
 }
}
