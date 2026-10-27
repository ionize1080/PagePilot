using Microsoft.Win32;
using PagePilot.Configuration;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PagePilot.Deployment {
 internal static partial class Manager {
  static string configFile;static string operationId=Guid.NewGuid().ToString("N");
  static void Event(string name,string detail="",string level="info")=>RuntimeLog.Write("manager",name,"operation="+operationId+" "+detail,level);
  internal static bool Elevated=>new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator);
  [DllImport("advapi32.dll",SetLastError=true)]static extern bool OpenProcessToken(IntPtr process,uint access,out IntPtr token);
  [DllImport("kernel32.dll")]static extern bool CloseHandle(IntPtr handle);
  internal static bool Owned(Process p){
   try{if(p.SessionId!=Process.GetCurrentProcess().SessionId)return false;if(!OpenProcessToken(p.Handle,8,out var token))return false;try{using(var id=new WindowsIdentity(token))return id.User==WindowsIdentity.GetCurrent().User;}finally{CloseHandle(token);}}catch{return false;}
  }
  internal static Process[] WordProcesses()=>Process.GetProcessesByName("WINWORD").Where(p=>{try{return p.SessionId==Process.GetCurrentProcess().SessionId;}catch{return false;}}).ToArray();
  internal static void CloseSelected(Process[] processes,bool force){
   foreach(var p in processes){try{if(p.HasExited)continue;if(!Owned(p))throw new InvalidOperationException("不能关闭其他账户或无法确认归属的 Word 进程 PID "+p.Id);
    Event(force?"word.force_close.request":"word.close.request","pid="+p.Id);
    if(force)p.Kill();else p.CloseMainWindow();
   }catch(InvalidOperationException){if(!p.HasExited)throw;}}
   var until=DateTime.UtcNow.AddSeconds(force?10:12);
   while(DateTime.UtcNow<until&&processes.Any(p=>!p.HasExited))Thread.Sleep(100);
   if(processes.Any(p=>!p.HasExited))throw new InvalidOperationException("Word 尚未退出，可能正在等待保存确认。请切换到 Word 完成保存，或选择强制关闭。");
   Event("word.closed","count="+processes.Length+" forced="+force);
  }
  static bool IsRegistered(RegistryView view){
   string expected=testPrefix==""?AssemblyName.GetAssemblyName(Dll).FullName:"PagePilot, Version=0.4.5.0, Culture=neutral, PublicKeyToken=null";
   using(var u=Store(view)){
    for(int i=0;i<Classes.Length;i++){
     using(var k=u.OpenSubKey(StorePath(@"Software\Classes\CLSID\"+Ids[i]+@"\InprocServer32")))
      if(k==null||Convert.ToString(k.GetValue(""))!="mscoree.dll"||Convert.ToString(k.GetValue("ThreadingModel"))!="Both"||Convert.ToString(k.GetValue("Class"))!=Classes[i]||Convert.ToString(k.GetValue("Assembly"))!=expected||Convert.ToString(k.GetValue("RuntimeVersion"))!="v4.0.30319"||!string.Equals(Convert.ToString(k.GetValue("CodeBase")),new Uri(Dll).ToString(),StringComparison.OrdinalIgnoreCase))return false;
     using(var k=u.OpenSubKey(StorePath(@"Software\Classes\"+Classes[i]+@"\CLSID")))if(Convert.ToString(k?.GetValue(""))!=Ids[i])return false;
    }
    using(var k=u.OpenSubKey(StorePath(Addin)))return Convert.ToString(k?.GetValue("LoadBehavior"))=="3";
   }
  }
  static void Exclusive(Action action){
   using(var mutex=new Mutex(false,"Global\\PagePilot.Deploy.Registry")){
    bool entered=false;try{try{entered=mutex.WaitOne(0);}catch(AbandonedMutexException){entered=true;}if(!entered)throw new Exception("另一部署器正在操作，请等待它完成。");operationId=Guid.NewGuid().ToString("N");action();}finally{if(entered)mutex.ReleaseMutex();}
   }
  }
  static void ReadyConfig(){if(RuntimePaths.Warning!=null)throw new Exception("运行配置无效，请在“存储与日志”修正后重试。"+RuntimePaths.Warning);}
  static string Status(bool doctor){
   var sb=new StringBuilder();sb.AppendLine("PagePilot 0.4.5 · 部署诊断");sb.AppendLine("生成时间 UTC："+DateTime.UtcNow.ToString("o"));
   using(var id=WindowsIdentity.GetCurrent()){sb.AppendLine("账户："+id.Name+" · SID "+id.User);sb.AppendLine("权限："+(Elevated?"管理员（Word 以管理员运行时请选择整台电脑范围）":"普通用户"));}
   sb.AppendLine("系统："+Environment.OSVersion+" · "+(Environment.Is64BitOperatingSystem?"64":"32")+" 位");sb.AppendLine("版本目录："+root);sb.AppendLine("部署范围："+ScopeLabel);
   try{string word=DetectWord();var view=WordView();sb.AppendLine("Word："+word);sb.AppendLine("Word 版本："+FileVersionInfo.GetVersionInfo(word).FileVersion+" · "+(view==RegistryView.Registry64?"64":"32")+" 位");
    using(var machine=RegistryKey.OpenBaseKey(RegistryHive.LocalMachine,RegistryView.Registry32))using(var framework=machine.OpenSubKey(@"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"))sb.AppendLine(".NET Framework Release："+Convert.ToString(framework?.GetValue("Release")));
    foreach(var v in new[]{RegistryView.Registry32,RegistryView.Registry64}){
     foreach(var hive in new[]{RegistryHive.CurrentUser,RegistryHive.LocalMachine})using(var source=RegistryKey.OpenBaseKey(hive,v)){
      using(var k=source.OpenSubKey(Addin))sb.AppendLine(hive+" · "+v+" · LoadBehavior="+Convert.ToString(k?.GetValue("LoadBehavior")));
      using(var k=source.OpenSubKey(@"Software\Classes\CLSID\"+Ids[0]+@"\InprocServer32"))sb.AppendLine(hive+" · "+v+" · CodeBase="+Convert.ToString(k?.GetValue("CodeBase")));
     }
     using(var classes=RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot,v))using(var server=classes.OpenSubKey(@"CLSID\"+Ids[0]+@"\InprocServer32"))sb.AppendLine(v+" · HKCR 合并视图（不代表管理员 COM 已可激活）="+Convert.ToString(server?.GetValue("CodeBase")));
    }
    sb.AppendLine("注册核验："+(IsRegistered(view)?"此目录的全部类与启动标记匹配":"未注册此版本或注册不完整"));
    string ver=FileVersionInfo.GetVersionInfo(word).FileMajorPart+".0";
    foreach(var hive in new[]{RegistryHive.CurrentUser,RegistryHive.LocalMachine})using(var k=RegistryKey.OpenBaseKey(hive,view)){
     foreach(string area in new[]{"DisabledItems","CrashingAddinList"})using(var sub=k.OpenSubKey(@"Software\Microsoft\Office\"+ver+@"\Word\Resiliency\"+area))sb.AppendLine(hive+" · "+area+" 记录数="+(sub?.ValueCount??0)+"（只读检查，不清除其他加载项记录）");
     using(var policy=k.OpenSubKey(@"Software\Policies\Microsoft\Office\"+ver+@"\Word\Resiliency\AddinList"))sb.AppendLine(hive+" · PagePilot 策略="+Convert.ToString(policy?.GetValue("PagePilot.Connect")));
    }
   }catch(Exception ex){sb.AppendLine("注册或环境检查失败："+ex.Message);Event("diagnostic.environment.failed",ex.ToString(),"error");}
   try{ValidatePackage(doctor);sb.AppendLine(doctor?"完整运行文件校验通过":"关键运行文件校验通过");}catch(Exception ex){sb.AppendLine("运行包："+ex.Message);}
   if(File.Exists(Dll+":Zone.Identifier"))sb.AppendLine("提示：DLL 存在下载区域标记；请检查文件属性中的来源限制。");
   foreach(var process in WordProcesses())try{sb.AppendLine("Word 进程 PID="+process.Id+" · 当前用户="+Owned(process)+" · 启动 UTC="+process.StartTime.ToUniversalTime().ToString("o")+" · "+(Loaded(process)?"此版本已连接且功能区已加载":"尚无此版本的完整加载回执"));}catch(Exception e){sb.AppendLine("进程检查失败："+e.Message);}
   if(WordProcesses().Length==0)sb.AppendLine("Word 未运行。注册完成不等于已经加载，请点击“启动 Word 并验证”。");
   var settings=RuntimePaths.Current;sb.AppendLine("配置文件："+configFile);sb.AppendLine("规则与备份："+settings.DataDirectory);sb.AppendLine("模型与会话："+settings.ProfileDirectory);sb.AppendLine("日志目录："+settings.LogDirectory);sb.AppendLine("日志级别："+settings.LogLevel);sb.AppendLine("部署事务："+JournalDirectory);
   if(RuntimePaths.Warning!=null)sb.AppendLine("配置错误，当前显示默认路径："+RuntimePaths.Warning);
   sb.AppendLine("若无 PagePilot 选项卡：文件 → 选项 → 加载项 → 管理 COM 加载项；并检查“禁用项目”。管理器不会清空 Office 禁用列表。");
   if(doctor)sb.AppendLine(EnvironmentReport());
   Event("diagnostic.completed","full="+doctor);return sb.ToString();
  }
  static bool Loaded(Process process){
   try{var receipt=json.Deserialize<HostReceipt>(File.ReadAllText(HostHealth.ReceiptPath(RuntimePaths.Current.LogDirectory,process.Id)));return HostHealth.Matches(receipt,process.Id,process.StartTime.ToUniversalTime().Ticks,Dll)&&receipt.UserSid==WindowsIdentity.GetCurrent().User.Value;}catch{return false;}
  }
  static async Task<string> VerifyWord(){
   ReadyConfig();CheckOtherScope(WordView());if(!IsRegistered(WordView()))throw new Exception("请先启用此版本，再启动验证。");
   if(WordProcesses().Length>0){if(WordProcesses().Any(Loaded))return "已验证：当前版本已连接 Word，且功能区已加载。\r\n\r\n"+Status(false);throw new Exception("检测到尚未验证的 Word 进程。请先保存并关闭 Word，再重新启动验证。");}
   Event("verify.begin");using(var launched=Process.Start(new ProcessStartInfo(DetectWord(),"/w"){UseShellExecute=true})){
    for(int i=0;i<150;i++){await Task.Delay(200);if(WordProcesses().Any(Loaded)){Event("verify.completed","connection and ribbon acknowledged");return "已验证：PagePilot 已连接，功能区已加载。\r\n\r\n"+Status(false);}}
   }
   Event("verify.timeout","30 seconds; Word left open for inspection","error");throw new Exception("30 秒内没有收到完整加载回执。Word 保持打开，请检查首次启动提示或禁用项目，再导出诊断包。");
  }
  internal static void CopyTree(string source,string destination,HashSet<string> skip=null){
   string from=Path.GetFullPath(source).TrimEnd('\\')+"\\",to=Path.GetFullPath(destination).TrimEnd('\\')+"\\";
   if(string.Equals(from,to,StringComparison.OrdinalIgnoreCase)||!Directory.Exists(from))return;
   if(to.StartsWith(from,StringComparison.OrdinalIgnoreCase)||from.StartsWith(to,StringComparison.OrdinalIgnoreCase))throw new Exception("新旧数据目录不能相互包含，请选择独立目录。");
   CopyDirectory(new DirectoryInfo(from),to,skip??new HashSet<string>(StringComparer.OrdinalIgnoreCase));
  }
  static void CopyDirectory(DirectoryInfo source,string destination,HashSet<string> skip){
   if((source.Attributes&FileAttributes.ReparsePoint)!=0)throw new IOException("不复制重解析点："+source.FullName);
   if(Directory.Exists(destination)&&(File.GetAttributes(destination)&FileAttributes.ReparsePoint)!=0)throw new IOException("目标目录不能是重解析点");
   Directory.CreateDirectory(destination);
   foreach(var f in source.GetFiles()){if(skip.Contains(f.Name)||f.Name.EndsWith(".tmp",StringComparison.OrdinalIgnoreCase))continue;if((f.Attributes&FileAttributes.ReparsePoint)!=0)throw new IOException("不复制链接文件");string dest=Path.Combine(destination,f.Name);if(!File.Exists(dest))f.CopyTo(dest,false);}
   foreach(var d in source.GetDirectories())if(!skip.Contains(d.Name))CopyDirectory(d,Path.Combine(destination,d.Name),skip);
  }
  static void SaveConfiguration(RuntimeSettings settings,string file,bool copy){
   Exclusive(()=>{Closed();settings.Normalize(file);var old=RuntimePaths.Current;foreach(var folder in new[]{settings.DataDirectory,settings.ProfileDirectory,settings.LogDirectory})RuntimeSettings.EnsureWritable(folder);
    if(copy){Event("config.copy.begin");CopyTree(old.DataDirectory,settings.DataDirectory,new HashSet<string>(new[]{"deployment","logs","runtime.json","runtime.json.previous","host.log"},StringComparer.OrdinalIgnoreCase));CopyTree(old.ProfileDirectory,settings.ProfileDirectory,new HashSet<string>(new[]{"Cache","Code Cache","GPUCache","DawnCache","Crashpad","SingletonLock","lockfile"},StringComparer.OrdinalIgnoreCase));Event("config.copy.completed");}
    settings.Save(file,true);RuntimePaths.Set(settings,Path.GetFullPath(file));configFile=RuntimePaths.ConfigFile;Event("config.saved","file="+configFile+" data="+settings.DataDirectory+" profile="+settings.ProfileDirectory+" logs="+settings.LogDirectory+" copy="+copy);
   });
  }
  static string ExportDiagnostics(string destination){
   string report=Status(true);Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(destination)));Event("diagnostic.export.begin");
   using(var stream=new FileStream(destination,FileMode.CreateNew))using(var zip=new ZipArchive(stream,ZipArchiveMode.Create)){
    Action<string,string> add=(name,content)=>{var e=zip.CreateEntry(name,CompressionLevel.Optimal);using(var w=new StreamWriter(e.Open(),new UTF8Encoding(false)))w.Write(RuntimeLog.Redact(content));};
    add("diagnostic.txt",report);
    var folder=RuntimePaths.Current.LogDirectory;if(Directory.Exists(folder))foreach(string f in Directory.GetFiles(folder)){
     string name=Path.GetFileName(f);if(!System.Text.RegularExpressions.Regex.IsMatch(name,@"^(manager|host|electron-\d+|bridge-\d+)\.jsonl(\.[123])?$"))continue;
     try{using(var fs=new FileStream(f,FileMode.Open,FileAccess.Read,FileShare.ReadWrite|FileShare.Delete))using(var reader=new StreamReader(fs))add("logs/"+name,reader.ReadToEnd());}catch(IOException e){add("logs/"+name+".read-error.txt",e.Message);}
    }
    try{if(File.Exists(JournalPath))add("deployment/last.json",File.ReadAllText(JournalPath));}catch(UnauthorizedAccessException){add("deployment/read-error.txt","机器级回退记录需要管理员权限读取。");}
   }
   Event("diagnostic.export.completed","file="+destination);return destination;
  }
  static void OpenLogs(){var directory=RuntimePaths.Current.LogDirectory;Directory.CreateDirectory(directory);Event("logs.open");Process.Start(new ProcessStartInfo("explorer.exe","\""+directory.TrimEnd('\\')+"\""){UseShellExecute=true});}
  [STAThread]static int Main(string[] args){
   Application.EnableVisualStyles();Application.SetCompatibleTextRenderingDefault(false);string output=null;
   try{
    string command="ui",explicitConfig=null;bool force=false,acceptLoss=false,verify=false;
    for(int i=0;i<args.Length;i++){
     switch(args[i].ToLowerInvariant()){
      case "--replace-existing":replaceExisting=true;break;
      case "--scope":SelectScope(Next(args,ref i).ToLowerInvariant());break;
      case "--root":root=Path.GetFullPath(Next(args,ref i)).TrimEnd(Path.DirectorySeparatorChar);break;
      case "--config":explicitConfig=Path.GetFullPath(Next(args,ref i));break;
      case "--output":output=Path.GetFullPath(Next(args,ref i));break;
      case "--force-close-word":force=true;break;
      case "--accept-unsaved-loss":acceptLoss=true;break;
      case "--verify":verify=true;break;
      default:if(command!="ui"||i!=0)throw new Exception("未知参数："+args[i]);command=args[i].TrimStart('-','/').ToLowerInvariant();break;
     }
    }
    RuntimePaths.Reload(explicitConfig);configFile=RuntimePaths.ConfigFile;Event("manager.start","version=0.4.5 command="+command+" scope="+ScopeName+" elevated="+Elevated+" sid="+WindowsIdentity.GetCurrent().User+" os="+Environment.OSVersion);
    using(var process=Process.GetCurrentProcess())Event("environment.manager","bitness="+(Environment.Is64BitProcess?64:32)+" CLR="+Environment.Version+" "+ProcessSecurity(process));
    if(force&&!acceptLoss)throw new Exception("命令行强制关闭需要同时提供 --accept-unsaved-loss，未保存的 Word 修改会丢失。");
    if(command=="ui"){RunUi();return 0;}
    string result=null;
    switch(command){
     case "self-test":SelfTest();result="PASS isolated deployment regression";break;
     case "status":case "doctor":result=Status(command=="doctor");break;
     case "open-logs":OpenLogs();break;
     case "export-diagnostics":if(output==null)throw new Exception("请用 --output 指定新的 ZIP 文件路径");ExportDiagnostics(output);output=null;break;
     case "enable":case "disable":case "rollback":
      ReadyConfig();if(command=="enable")CheckReplacementMode();RequireScopePermission();Exclusive(()=>{if(force)CloseSelected(WordProcesses(),true);Closed();if(command=="enable"&&explicitConfig!=null)RuntimePaths.Current.Save(explicitConfig,true);if(command=="rollback")Rollback();else Change(command,WordView());});
      result=command=="enable"?"注册已完成，尚未验证 Word 加载。请启动 Word 并验证。":"操作完成";
      if(verify&&command=="enable")result=VerifyWord().GetAwaiter().GetResult();break;
     case "verify":result=VerifyWord().GetAwaiter().GetResult();break;
     default:throw new Exception("支持 ui、enable、disable、rollback、verify、status、doctor、open-logs、export-diagnostics、self-test");
    }
    if(output!=null){Directory.CreateDirectory(Path.GetDirectoryName(output));File.WriteAllText(output,result??"完成",new UTF8Encoding(false));}
    else if(result!=null&&command!="self-test")MessageBox.Show(result,"PagePilot",MessageBoxButtons.OK,MessageBoxIcon.Information);
    Event("manager.completed",command);return 0;
   }catch(Exception e){Event("manager.failed",e.ToString(),"error");if(output!=null){Directory.CreateDirectory(Path.GetDirectoryName(output));File.WriteAllText(output,e.ToString(),new UTF8Encoding(false));}else MessageBox.Show(e.Message,"PagePilot · 未完成",MessageBoxButtons.OK,MessageBoxIcon.Warning);return 1;}
  }
  static string Next(string[] args,ref int index){if(++index>=args.Length)throw new Exception("参数缺少值");return args[index];}
 }
}
