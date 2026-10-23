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
  private bool pendingAI,pendingProvider;private string pendingPrompt;
  private Process settingsProcess;private Office.IRibbonUI ribbon;private Office.CustomTaskPane aiPane;private AiPane aiControl;private IntPtr aiHandle;
  private int aiOwner;internal int Theme=0,NavigationGap=36;internal string AiMode="window";private string approvalDocument;
  public void RibbonLoaded(Office.IRibbonUI value){ribbon=value;}
  public object RibbonImage(Office.IRibbonControl control){IntPtr window=IntPtr.Zero;try{window=(IntPtr)(int)App.ActiveWindow.Hwnd;}catch{}return RibbonIcons.Picture(control.Id,window,Theme);}
  private void LoadPreferences(){try{var p=Json().Deserialize<Dictionary<string,object>>(File.ReadAllText(Path.Combine(Data,"preferences.json")));Theme=Convert.ToInt32(p["theme"]);NavigationGap=Convert.ToInt32(p["navigationGap"]);AiMode=Convert.ToString(p["aiMode"]);}catch{}}
  private void SavePreferences(){Directory.CreateDirectory(Data);File.WriteAllText(Path.Combine(Data,"preferences.json"),Json().Serialize(new{theme=Theme,navigationGap=NavigationGap,aiMode=AiMode}));}
  internal void QueuePrompt(string text){pendingPrompt=text;LaunchChat();}
  private void SetAiMode(string mode){
   if(mode!="sidebar"&&mode!="window")throw new Exception("无效显示模式");
   AiMode=mode;SavePreferences();if(aiHandle==IntPtr.Zero)return;
   if(mode=="sidebar"){
    if(App.Documents.Count==0)throw new Exception("请先打开文档");int hwnd=(int)App.ActiveWindow.Hwnd;
    if(aiPane!=null&&aiOwner!=hwnd){aiControl.Detach();aiPane.Delete();aiPane=null;}
    if(aiPane==null){aiPane=factory.CreateCTP("PagePilot.AiPane","PagePilot · AI 助手",App.ActiveWindow);aiPane.DockPosition=Office.MsoCTPDockPosition.msoCTPDockPositionRight;aiPane.Width=(int)(480*Native.Dpi((IntPtr)hwnd));aiControl=(AiPane)aiPane.ContentControl;aiOwner=hwnd;}
    aiPane.Visible=true;aiControl.Attach(aiHandle);
   }else if(aiPane!=null){aiControl.Detach();aiPane.Visible=false;}
  }
  private object ApprovalState(){string current=App.Documents.Count==0?"":(string)App.ActiveDocument.FullName;if(approvalDocument!=current)AutoApprove=false;return new{automatic=AutoApprove,mode=AiMode};}
  private NamedPipeServerStream listener;
  private volatile bool stopping;
  private Process chat;
  internal static void Log(string message){try{Directory.CreateDirectory(Data);File.AppendAllText(Path.Combine(Data,"host.log"),DateTime.UtcNow.ToString("o")+" "+message+Environment.NewLine);}catch{}}
  public void OnConnection(object app,ext_ConnectMode mode,object addin,ref Array custom){
   Log("OnConnection");
   Current=this;App=app;LoadPreferences();Dispatcher=new Control();var h=Dispatcher.Handle;
   ((dynamic)addin).Object=this;
   try{StartBridge();Log("Connected Word "+App.Version);}catch(Exception e){Log(e.ToString());}
  }
  public void OnDisconnection(ext_DisconnectMode mode,ref Array custom){Shutdown();}
  public void OnBeginShutdown(ref Array custom){Shutdown();}
  public void OnStartupComplete(ref Array custom){}
  public void OnAddInsUpdate(ref Array custom){}
  private void Shutdown(){stopping=true;aiControl?.Detach();try{if(settingsProcess!=null&&!settingsProcess.HasExited)settingsProcess.CloseMainWindow();}catch{}listener?.Dispose();listener=null;try{if(chat!=null&&!chat.HasExited)chat.CloseMainWindow();}catch{} Dispatcher?.Dispose();panes.Clear();Current=null;}
  public void CTPFactoryAvailable(Office.ICTPFactory value){factory=value;}
  public string GetCustomUI(string ribbonId)=>@"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='RibbonLoaded'><ribbon><tabs><tab id='PagePilot' label='PagePilot'><group id='ppTools' label='文档增强'><button id='ppAI' label='AI 助手' getImage='RibbonImage' size='large' onAction='OpenAI'/><button id='ppRegex' label='正则查找替换' getImage='RibbonImage' size='large' onAction='OpenRegex'/><button id='ppNav' label='自动导航' getImage='RibbonImage' size='large' onAction='OpenNavigation'/><button id='ppSettings' label='设置' getImage='RibbonImage' size='large' onAction='OpenSettings'/></group></tab></tabs></ribbon></customUI>";
  public void OpenAI(Office.IRibbonControl c){if(Dispatcher.InvokeRequired){Dispatcher.Invoke(new Action(()=>OpenAI(c)));return;}try{LaunchChat();}catch(Exception e){MessageBox.Show(e.Message,"PagePilot");}}
  public void LaunchAssistant(){OpenAI(null);}
  public void OpenRegex(Office.IRibbonControl c){ShowPane(0);}
  public void OpenNavigation(Office.IRibbonControl c){ShowPane(1);}
  public void OpenSettings(Office.IRibbonControl c){if(Dispatcher.InvokeRequired){Dispatcher.Invoke(new Action(()=>OpenSettings(c)));return;}try{if(settingsProcess!=null&&!settingsProcess.HasExited){Native.SetForegroundWindow(settingsProcess.MainWindowHandle);return;}settingsProcess=StartElectron("settings-entry.cjs");}catch(Exception e){MessageBox.Show(e.Message,"PagePilot");}}
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
   if(chat!=null&&!chat.HasExited){pendingAI=true;if(AiMode=="sidebar")SetAiMode("sidebar");return;}
   chat=StartElectron(null);
  }
  private Process StartElectron(string entryFile){
   string executable=Path.Combine(Root,"Browser","electron.exe"),entry=entryFile==null?Path.Combine(Root,"Chatbox"):Path.Combine(Root,"Chatbox",entryFile);
   if(!File.Exists(executable))throw new FileNotFoundException("Chatbox 运行组件缺失，请使用完整便携包。");
   var p=new ProcessStartInfo(executable,"\""+entry+"\""){UseShellExecute=false,WorkingDirectory=Root};
   p.EnvironmentVariables["PAGEPILOT_PIPE"]=PipeName;p.EnvironmentVariables["PAGEPILOT_TOKEN"]=Token;
   p.EnvironmentVariables["PAGEPILOT_ROOT"]=Root;return Process.Start(p);
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
    ApprovalState();bool allowed=AutoApprove||Approval(name,args);
    Log("Tool "+name+" "+(allowed?"approved":"denied"));return new{approved=allowed};
   }
   if(action=="ping"){
    bool forAI=req.ContainsKey("role")&&Convert.ToString(req["role"])=="ai";bool activate=forAI&&pendingAI,provider=forAI&&pendingProvider;string prompt=forAI?pendingPrompt:null;
    if(forAI){pendingAI=false;pendingProvider=false;pendingPrompt=null;}
    return new{version="0.2.0",activate,provider,prompt,state=ApprovalState()};
   }
   if(action=="ai_ready"){aiHandle=new IntPtr(Convert.ToInt64(req["handle"]));SetAiMode(AiMode);return ApprovalState();}
   if(action=="ai_mode"){SetAiMode(Convert.ToString(req["mode"]));return ApprovalState();}
   if(action=="ai_approval"){
    bool enabled=Convert.ToBoolean(req["enabled"]);ApprovalState();
    if(enabled&&!AutoApprove){if(App.Documents.Count==0)throw new Exception("请先打开文档");
     if(!ConfirmAutomatic())return ApprovalState();
     approvalDocument=(string)App.ActiveDocument.FullName;
    }AutoApprove=enabled;return ApprovalState();
   }
   if(action=="settings_provider"){pendingProvider=true;LaunchChat();return new{opened=true};}
   if(action=="settings_read")return new{theme=Theme,navigationGap=NavigationGap,aiMode=AiMode,headingRules=ReadRulesText()};
   if(action=="settings_write"){
    string rules=Convert.ToString(req["headingRules"]),mode=Convert.ToString(req["aiMode"]);Pane.ParseRules(rules);
    int theme=Convert.ToInt32(req["theme"]),gap=Convert.ToInt32(req["navigationGap"]);if(theme<0||theme>2||gap<0||gap>160||(mode!="window"&&mode!="sidebar"))throw new Exception("设置值超出范围");
    Directory.CreateDirectory(Data);File.WriteAllText(Path.Combine(Data,"heading-rules.txt"),rules);Theme=theme;NavigationGap=gap;SetAiMode(mode);ribbon?.Invalidate();return new{saved=true};
   }
   if(action=="capabilities")return new{wordVersion=(string)App.Version,majorVersion=(int)double.Parse((string)App.Version,System.Globalization.CultureInfo.InvariantCulture),documentOpen=App.Documents.Count>0,protectedDocument=App.Documents.Count>0?(int)App.ActiveDocument.ProtectionType!=-1:false};
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
  private bool ConfirmAutomatic(){float dpi=Native.Dpi((IntPtr)(int)App.ActiveWindow.Hwnd);using(var form=new Form{Text="PagePilot · 确认自动批准",Width=(int)(560*dpi),Height=(int)(290*dpi),Font=new System.Drawing.Font("Microsoft YaHei UI",10),StartPosition=FormStartPosition.CenterScreen,TopMost=true,FormBorderStyle=FormBorderStyle.FixedDialog,MaximizeBox=false,MinimizeBox=false}){
   form.Controls.Add(new Label{Dock=DockStyle.Fill,Padding=new Padding((int)(24*dpi)),Text="启用后，AI 可直接修改当前文档的文字、格式、表格和结构，无需逐项批准。操作前会保存备份。\n\n仅对当前文档、本次 Word 会话有效。切换文档或重启会恢复手动批准。\n\n确认启用自动批准？"});
   var bar=new FlowLayoutPanel{Dock=DockStyle.Bottom,Height=(int)(56*dpi),Padding=new Padding((int)(8*dpi)),FlowDirection=FlowDirection.RightToLeft};
   var yes=new Button{Text="启用自动批准",DialogResult=DialogResult.Yes,Width=(int)(140*dpi),Height=(int)(32*dpi)};var no=new Button{Text="保持手动",DialogResult=DialogResult.No,Width=(int)(110*dpi),Height=(int)(32*dpi)};bar.Controls.Add(yes);bar.Controls.Add(no);form.Controls.Add(bar);form.AcceptButton=no;form.CancelButton=no;form.Shown+=(s,e)=>no.Focus();return form.ShowDialog()==DialogResult.Yes;
  }}
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
