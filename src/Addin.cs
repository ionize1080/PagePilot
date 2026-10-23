using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.IO.Pipes;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Web.Script.Serialization;
using Office=Microsoft.Office.Core;
using Extensibility;
namespace PagePilot {
 [ComVisible(true),Guid("0E057EBB-B35B-4A31-9B65-3DF44E8034CE"),ProgId("PagePilot.Connect"),ClassInterface(ClassInterfaceType.AutoDispatch)]
 public sealed class Connect:IDTExtensibility2,Office.IRibbonExtensibility,Office.ICustomTaskPaneConsumer {
  public static Connect Current;
  internal dynamic App;
  internal Control Dispatcher;
  internal bool AutoApprove;
  internal string Token;
  internal string PipeName;
  internal static string Root=>Path.GetDirectoryName(typeof(Connect).Assembly.Location);
  internal static string Data=>Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"PagePilot");
  internal static JavaScriptSerializer Json()=>new JavaScriptSerializer{MaxJsonLength=16*1024*1024};
  private Office.ICTPFactory factory;
  private readonly Dictionary<string,Office.CustomTaskPane> panes=new Dictionary<string,Office.CustomTaskPane>();
  private bool pendingSettings;
  private NamedPipeServerStream listener;
  private volatile bool stopping;
  private Process chat;
  internal static void Log(string message){try{Directory.CreateDirectory(Data);File.AppendAllText(Path.Combine(Data,"host.log"),DateTime.UtcNow.ToString("o")+" "+message+Environment.NewLine);}catch{}}
  public void OnConnection(object app,ext_ConnectMode mode,object addin,ref Array custom){
   Log("OnConnection");
   Current=this;App=app;Dispatcher=new Control();var h=Dispatcher.Handle;
   ((dynamic)addin).Object=this;
   try{StartBridge();Log("Connected Word "+App.Version);}catch(Exception e){Log(e.ToString());}
  }
  public void OnDisconnection(ext_DisconnectMode mode,ref Array custom){Shutdown();}
  public void OnBeginShutdown(ref Array custom){Shutdown();}
  public void OnStartupComplete(ref Array custom){}
  public void OnAddInsUpdate(ref Array custom){}
  private void Shutdown(){stopping=true;listener?.Dispose();listener=null;try{if(chat!=null&&!chat.HasExited)chat.CloseMainWindow();}catch{} Dispatcher?.Dispose();panes.Clear();Current=null;}
  public void CTPFactoryAvailable(Office.ICTPFactory value){factory=value;}
  public string GetCustomUI(string ribbonId)=>@"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'><ribbon><tabs><tab id='PagePilot' label='PagePilot'><group id='ppTools' label='文档增强'><button id='ppAI' label='AI 助手' imageMso='HappyFace' size='large' onAction='OpenAI'/><button id='ppRegex' label='正则查找替换' imageMso='FindDialog' size='large' onAction='OpenRegex'/><button id='ppNav' label='自动导航' imageMso='DocumentMap' size='large' onAction='OpenNavigation'/><button id='ppSettings' label='设置' imageMso='PropertySheet' size='large' onAction='OpenSettings'/></group></tab></tabs></ribbon></customUI>";
  public void OpenAI(Office.IRibbonControl c){if(Dispatcher.InvokeRequired){Dispatcher.Invoke(new Action(()=>OpenAI(c)));return;}try{LaunchChat();}catch(Exception e){MessageBox.Show(e.Message,"PagePilot");}}
  public void LaunchAssistant(){OpenAI(null);}
  public void OpenRegex(Office.IRibbonControl c){ShowPane(0);}
  public void OpenNavigation(Office.IRibbonControl c){ShowPane(1);}
  public void OpenSettings(Office.IRibbonControl c){if(Dispatcher.InvokeRequired){Dispatcher.Invoke(new Action(()=>OpenSettings(c)));return;}pendingSettings=true;OpenAI(null);}
  internal void SetNavigation(List<Heading> headings){var pane=panes[(int)App.ActiveWindow.Hwnd+":1"];((Pane)pane.ContentControl).SetOutline(headings);}
  public string GetMcpSession(){return Json().Serialize(new{pipe=PipeName,token=Token});}
  public void ShowPane(int page){if(Dispatcher.InvokeRequired){Dispatcher.Invoke(new Action(()=>ShowPane(page)));return;}try{
   if(page==2){OpenSettings(null);return;}
   Log("ShowPane "+page);
   if(App.Documents.Count==0)throw new InvalidOperationException("请先打开文档。");
   int hwnd=(int)App.ActiveWindow.Hwnd;string key=hwnd+":"+page;
   if(!panes.TryGetValue(key,out var pane)){
    if(factory==null)throw new InvalidOperationException("Word 尚未初始化任务窗格，请重新启动 Word。");
    pane=factory.CreateCTP("PagePilot.Pane",page==0?"PagePilot · 正则替换":"PagePilot · 自动导航",App.ActiveWindow);Log("Created pane");
    pane.DockPosition=Office.MsoCTPDockPosition.msoCTPDockPositionRight;pane.Width=(int)(370*Native.Dpi((IntPtr)hwnd));panes[key]=pane;
    ((Pane)pane.ContentControl).Bind(this,App.ActiveDocument,page);Log("Bound pane");
   }
   ((Pane)pane.ContentControl).SelectPage(page);pane.Visible=true;
  }catch(Exception e){Log(e.ToString());MessageBox.Show(e.Message,"PagePilot");}}
  internal void LaunchChat(){
   if(chat!=null&&!chat.HasExited){Native.SetForegroundWindow(chat.MainWindowHandle);return;}
   string executable=Path.Combine(Root,"Browser","electron.exe"),entry=Path.Combine(Root,"Chatbox");
   if(!File.Exists(executable))throw new FileNotFoundException("Chatbox 运行组件缺失，请使用完整便携包。");
   var p=new ProcessStartInfo(executable,"\""+entry+"\""){UseShellExecute=false,WorkingDirectory=Root};
   p.EnvironmentVariables["PAGEPILOT_PIPE"]=PipeName;p.EnvironmentVariables["PAGEPILOT_TOKEN"]=Token;
   p.EnvironmentVariables["PAGEPILOT_ROOT"]=Root;chat=Process.Start(p);
  }
  private void StartBridge(){
   byte[] bytes=new byte[32];using(var rng=RandomNumberGenerator.Create())rng.GetBytes(bytes);Token=BitConverter.ToString(bytes).Replace("-","").ToLowerInvariant();
   PipeName="PagePilot-"+Guid.NewGuid().ToString("N");
   new Thread(()=>{while(!stopping){try{
    var security=new PipeSecurity();security.AddAccessRule(new PipeAccessRule(WindowsIdentity.GetCurrent().User,PipeAccessRights.FullControl,AccessControlType.Allow));
    var pipe=new NamedPipeServerStream(PipeName,PipeDirection.InOut,10,PipeTransmissionMode.Byte,PipeOptions.Asynchronous,4096,4096,security);
    listener=pipe;pipe.WaitForConnection();ThreadPool.QueueUserWorkItem(_=>Serve(pipe));
   }catch{if(!stopping)Log("Named pipe stopped");break;}}}){IsBackground=true}.Start();
  }
  private void Serve(NamedPipeServerStream stream){using(stream){try{
   using(var reader=new StreamReader(stream,Encoding.UTF8,false,4096,true))using(var writer=new StreamWriter(stream,new UTF8Encoding(false),4096,true){AutoFlush=true}){
    var b=new StringBuilder();int ch;while((ch=reader.Read())!=-1&&ch!='\n'){if(b.Length>=8*1024*1024)throw new Exception("Request too large");b.Append((char)ch);}
    var req=Json().Deserialize<Dictionary<string,object>>(b.ToString());
    if(!req.TryGetValue("token",out var token)||!FixedEquals(Convert.ToString(token),Token))throw new Exception("Unauthorized");
    object result=null;Exception fault=null;
    Dispatcher.Invoke(new Action(()=>{try{result=Handle(req);}catch(Exception e){fault=e;}}));
    writer.WriteLine(Json().Serialize(fault==null?new{result,error=(string)null}:new{result=(object)null,error=fault.Message}));
   }
  }catch(Exception e){Log("Bridge: "+e.GetType().Name);}}}
  private static bool FixedEquals(string a,string b){if(a==null||b==null||a.Length!=b.Length)return false;int d=0;for(int i=0;i<a.Length;i++)d|=a[i]^b[i];return d==0;}
  internal object Handle(Dictionary<string,object> req){
   string action=Convert.ToString(req["action"]);
   if(action=="context")return Context();
   if(action=="approve"){
    string name=Convert.ToString(req["name"]),args=Convert.ToString(req["arguments"]);
    bool allowed=AutoApprove||Approval(name,args);
    Log("Tool "+name+" "+(allowed?"approved":"denied"));return new{approved=allowed};
   }
   if(action=="ping"){bool open=pendingSettings;pendingSettings=false;return new{version="0.1.0",word=Convert.ToString(App.Version),openSettings=open};}
   if(action=="settings_read")return new{autoApprove=AutoApprove,headingRules=ReadRulesText()};
   if(action=="settings_write"){
    string rules=Convert.ToString(req["headingRules"]);Pane.ParseRules(rules);Directory.CreateDirectory(Data);File.WriteAllText(Path.Combine(Data,"heading-rules.txt"),rules);
    AutoApprove=Convert.ToBoolean(req["autoApprove"]);return new{saved=true};
   }
   if(action=="checkpoint"){
    if(App.Documents.Count==0 || Convert.ToString(req["document"])!=(string)App.ActiveDocument.FullName)throw new Exception("活动文档已变化。");
    NativeTools.Backup(App.ActiveDocument);return new{saved=true};
   }
   if(action=="export_pdf"){
    string path=Path.GetFullPath(Convert.ToString(req["filename"])),output=Path.GetFullPath(Convert.ToString(req["output"]));
    dynamic target=null;bool opened=false;
    foreach(dynamic d in App.Documents)if(string.Equals((string)d.FullName,path,StringComparison.OrdinalIgnoreCase)){target=d;break;}
    try{if(target==null){target=App.Documents.Open(path,ReadOnly:true,Visible:false);opened=true;}target.ExportAsFixedFormat(output,17);return new{success=true,path=output};}
    finally{if(opened&&target!=null)target.Close(0);}
   }
   if(action=="native")return NativeTools.Execute(this,Convert.ToString(req["name"]),Json().Deserialize<Dictionary<string,object>>(Convert.ToString(req["arguments"])));
   throw new Exception("Unknown action");
  }
  internal static string ReadRulesText(){string file=Path.Combine(Data,"heading-rules.txt");return File.Exists(file)?File.ReadAllText(file):"";}
  internal object Context(){
   if(App.Documents.Count==0)return new{document=(object)null};
   dynamic d=App.ActiveDocument;var paragraphs=new List<object>();
   for(int i=1;i<=Math.Min((int)d.Paragraphs.Count,10000);i++){dynamic p=d.Paragraphs[i];paragraphs.Add(new{index=i,start=(int)p.Range.Start,end=(int)p.Range.End,text=(string)p.Range.Text,level=(int)p.OutlineLevel});}
   return new{document=new{name=(string)d.Name,path=(string)d.FullName,saved=(bool)d.Saved,paragraphs,selection=new{start=(int)App.Selection.Start,end=(int)App.Selection.End},truncated=(int)d.Paragraphs.Count>10000}};
  }
  internal bool Approval(string name,string args){float dpi=Native.Dpi((IntPtr)(int)App.ActiveWindow.Hwnd);using(var form=new Form{Text="PagePilot · 批准文档操作",Width=(int)(640*dpi),Height=(int)(480*dpi),Font=new System.Drawing.Font("Microsoft YaHei UI",9),StartPosition=FormStartPosition.CenterScreen,TopMost=true}){
   form.Controls.Add(new TextBox{Dock=DockStyle.Fill,Multiline=true,ReadOnly=true,ScrollBars=ScrollBars.Both,Text=name+Environment.NewLine+args});
   var bar=new FlowLayoutPanel{Dock=DockStyle.Bottom,Height=(int)(46*dpi),FlowDirection=FlowDirection.RightToLeft};
   var yes=new Button{Text="批准此次操作",Width=(int)(140*dpi),Height=(int)(32*dpi),DialogResult=DialogResult.OK};var no=new Button{Text="拒绝",Width=(int)(90*dpi),Height=(int)(32*dpi),DialogResult=DialogResult.Cancel};bar.Controls.Add(yes);bar.Controls.Add(no);form.Controls.Add(bar);form.CancelButton=no;
   return form.ShowDialog()==DialogResult.OK;
  }}
 }
 internal static class Native{
  [DllImport("user32.dll")]internal static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")]private static extern uint GetDpiForWindow(IntPtr h);
  internal static float Dpi(IntPtr h){try{return Math.Max(1,GetDpiForWindow(h)/96f);}catch{return 1;}}
 }
}
