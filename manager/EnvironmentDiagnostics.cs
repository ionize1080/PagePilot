using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;

namespace PagePilot.Deployment {
 internal static partial class Manager {
  [DllImport("advapi32.dll",SetLastError=true)]static extern bool GetTokenInformation(IntPtr token,int kind,IntPtr buffer,int size,out int needed);
  static string ProcessSecurity(Process process){
   if(!OpenProcessToken(process.Handle,8,out var token))return "无法读取权限 Win32="+Marshal.GetLastWin32Error();
   try{
    string owner;using(var id=new WindowsIdentity(token))owner=id.User.Value;
    Func<int,string> read=kind=>{
     GetTokenInformation(token,kind,IntPtr.Zero,0,out int size);if(size<=0)return "未知 Win32="+Marshal.GetLastWin32Error();
     IntPtr buffer=Marshal.AllocHGlobal(size);
     try{if(!GetTokenInformation(token,kind,buffer,size,out size))return "未知 Win32="+Marshal.GetLastWin32Error();
      if(kind==25)return new SecurityIdentifier(Marshal.ReadIntPtr(buffer)).Value;
      int value=Marshal.ReadInt32(buffer);return kind==20?(value==0?"否":"是"):value==1?"Default（无拆分令牌）":value==2?"Full（提升）":value==3?"Limited（受限）":value.ToString();
     }finally{Marshal.FreeHGlobal(buffer);}
    };
    return "SID="+owner+" · 已提升="+read(20)+" · 令牌="+read(18)+" · 完整性 SID="+read(25);
   }finally{CloseHandle(token);}
  }
  static void DiagnosticSection(StringBuilder report,string title,Action action){
   report.AppendLine("\r\n["+title+"]");try{action();}catch(Exception ex){report.AppendLine("收集失败："+ex.GetType().Name+" HRESULT=0x"+ex.HResult.ToString("X8")+" "+ex.Message);}
  }
  static string EnvironmentReport(){
   var sb=new StringBuilder();sb.AppendLine("PagePilot 0.4.5 · 扩展环境诊断 · UTC "+DateTime.UtcNow.ToString("o"));
   DiagnosticSection(sb,"系统与管理器",()=>{
    sb.AppendLine("系统位数="+(Environment.Is64BitOperatingSystem?64:32)+" · 管理器位数="+(Environment.Is64BitProcess?64:32)+" · CLR="+Environment.Version);
    using(var process=Process.GetCurrentProcess()){sb.AppendLine("Manager PID="+process.Id+" · Session="+process.SessionId+" · "+ProcessSecurity(process));}
    sb.AppendLine("范围="+ScopeName+" · 程序目录="+root+" · 回退目录="+JournalDirectory);
    using(var h=RegistryKey.OpenBaseKey(RegistryHive.LocalMachine,RegistryView.Registry64))using(var k=h.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion"))
     foreach(string name in new[]{"ProductName","DisplayVersion","ReleaseId","CurrentBuildNumber","UBR","InstallationType"})sb.AppendLine(name+"="+Convert.ToString(k?.GetValue(name)));
   });
   DiagnosticSection(sb,"UAC",()=>{
    using(var h=RegistryKey.OpenBaseKey(RegistryHive.LocalMachine,RegistryView.Registry64))using(var k=h.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"))
     foreach(string name in new[]{"EnableLUA","FilterAdministratorToken","ConsentPromptBehaviorAdmin","PromptOnSecureDesktop"})sb.AppendLine(name+"="+Convert.ToString(k?.GetValue(name,"未设置")));
    sb.AppendLine("值只供判断，不自动改变系统策略。完整性 SID 常见尾号：8192=中，12288=高，16384=系统。");
   });
   foreach(var view in new[]{RegistryView.Registry32,RegistryView.Registry64}){
    DiagnosticSection(sb,".NET "+view,()=>{
     using(var h=RegistryKey.OpenBaseKey(RegistryHive.LocalMachine,view))using(var k=h.OpenSubKey(@"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"))
      foreach(string name in new[]{"Release","Version","Install","InstallPath"})sb.AppendLine(name+"="+Convert.ToString(k?.GetValue(name)));
    });
    foreach(var hive in new[]{RegistryHive.CurrentUser,RegistryHive.LocalMachine})DiagnosticSection(sb,hive+" / "+view+" / PagePilot",()=>{
     using(var h=RegistryKey.OpenBaseKey(hive,view)){
      using(var addin=h.OpenSubKey(Addin))sb.AppendLine("Word 加载项：存在="+(addin!=null)+" · LoadBehavior="+Convert.ToString(addin?.GetValue("LoadBehavior"))+" · FriendlyName="+Convert.ToString(addin?.GetValue("FriendlyName")));
      for(int i=0;i<Ids.Length;i++){
       sb.AppendLine(Classes[i]+" "+Ids[i]);
       using(var k=h.OpenSubKey(@"Software\Classes\CLSID\"+Ids[i]+@"\InprocServer32")){
        sb.AppendLine("  存在="+(k!=null));foreach(string name in new[]{"","Assembly","Class","CodeBase","RuntimeVersion","ThreadingModel"})sb.AppendLine("  "+(name==""?"Server":name)+"="+Convert.ToString(k?.GetValue(name)));
       }
       using(var k=h.OpenSubKey(@"Software\Classes\"+Classes[i]+@"\CLSID"))sb.AppendLine("  ProgID→CLSID="+Convert.ToString(k?.GetValue("")));
      }
     }
    });
   }
   DiagnosticSection(sb,"运行文件",()=>{
    foreach(string relative in new[]{"PagePilot.dll","PagePilot.Manager.exe","Python/python.exe","Browser/electron.exe"}){
     string path=Path.Combine(root,relative);var file=new FileInfo(path);sb.AppendLine(relative+" · 存在="+file.Exists);
     if(!file.Exists)continue;
     sb.AppendLine("  大小="+file.Length+" · 修改 UTC="+file.LastWriteTimeUtc.ToString("o")+" · 版本="+FileVersionInfo.GetVersionInfo(path).FileVersion);
     try{using(var stream=File.OpenRead(path))sb.AppendLine("  可读取=True");}catch(Exception ex){sb.AppendLine("  读取失败="+ex.Message);}
     // Only collect ZoneId; ReferrerUrl/HostUrl may contain private download links.
     try{foreach(string line in File.ReadAllLines(path+":Zone.Identifier").Where(l=>l.StartsWith("ZoneId=",StringComparison.OrdinalIgnoreCase)))sb.AppendLine("  下载区域 "+line);}
     catch(FileNotFoundException){sb.AppendLine("  下载区域标记=无");}catch(Exception ex){sb.AppendLine("  下载区域检查="+ex.GetType().Name);}
    }
    sb.AppendLine("DLL 程序集="+AssemblyName.GetAssemblyName(Dll).FullName);
   });
   DiagnosticSection(sb,"Word 进程（所有会话）",()=>{
    var processes=Process.GetProcessesByName("WINWORD");sb.AppendLine("进程数="+processes.Length);
    foreach(var process in processes)using(process)DiagnosticSection(sb,"WINWORD PID="+process.Id,()=>{
     sb.AppendLine("Session="+process.SessionId+" · 当前账户与会话="+Owned(process)+" · 启动 UTC="+process.StartTime.ToUniversalTime().ToString("o"));
     sb.AppendLine(ProcessSecurity(process));sb.AppendLine("此版本完整加载回执="+Loaded(process));
     try{sb.AppendLine("实际进程文件="+process.MainModule.FileName);}catch(Exception ex){sb.AppendLine("进程路径无法读取="+ex.GetType().Name);}
    });
   });
   DiagnosticSection(sb,"Office 与加载策略",()=>{
    string word=DetectWord();var version=FileVersionInfo.GetVersionInfo(word);string office=version.FileMajorPart+".0";
    sb.AppendLine("Word="+word+" · 文件版本="+version.FileVersion+" · 位数="+(WordView()==RegistryView.Registry32?32:64));
    foreach(var hive in new[]{RegistryHive.CurrentUser,RegistryHive.LocalMachine})using(var h=RegistryKey.OpenBaseKey(hive,WordView())){
     foreach(string branch in new[]{@"Software\Microsoft\Office\",@"Software\Policies\Microsoft\Office\"}){
      using(var k=h.OpenSubKey(branch+office+@"\Word\Security"))foreach(string name in new[]{"DisableAllAddins","RequireAddinSig"})sb.AppendLine(hive+" "+branch+" "+name+"="+Convert.ToString(k?.GetValue(name)));
      using(var k=h.OpenSubKey(branch+office+@"\Word\Resiliency\AddinList"))sb.AppendLine(hive+" "+branch+" AddinList/PagePilot.Connect="+Convert.ToString(k?.GetValue("PagePilot.Connect")));
     }
     using(var k=h.OpenSubKey(@"Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers"))
      if(k!=null)foreach(string name in k.GetValueNames().Where(n=>string.Equals(Path.GetFileName(n),"WINWORD.EXE",StringComparison.OrdinalIgnoreCase)||string.Equals(Path.GetFileName(n),"PagePilot.Manager.exe",StringComparison.OrdinalIgnoreCase)))sb.AppendLine(hive+" 兼容性="+name+" → "+Convert.ToString(k.GetValue(name)));
     using(var k=h.OpenSubKey(@"Software\Microsoft\Office\"+office+@"\Word\Resiliency\DisabledItems")){
      sb.AppendLine(hive+" DisabledItems 数="+(k?.ValueCount??0));
      if(k!=null)foreach(string name in k.GetValueNames())if(k.GetValue(name) is byte[] bytes)
       sb.AppendLine("  记录="+name+" · 长度="+bytes.Length+" · 含 PagePilot="+(Encoding.Unicode.GetString(bytes).IndexOf("PagePilot",StringComparison.OrdinalIgnoreCase)>=0||Encoding.ASCII.GetString(bytes).IndexOf("PagePilot",StringComparison.OrdinalIgnoreCase)>=0));
     }
    }
   });
   DiagnosticSection(sb,"近期应用程序错误（48 小时，最多检查最新 400 条，返回 20 条）",()=>{
    using(var log=new EventLog("Application")){
     int total=log.Entries.Count,found=0;DateTime since=DateTime.Now.AddHours(-48);
     for(int i=total-1;i>=Math.Max(0,total-400)&&found<20;i--){
      var entry=log.Entries[i];if(entry.TimeGenerated<since)break;
      if(entry.EntryType!=EventLogEntryType.Error&&entry.EntryType!=EventLogEntryType.Warning)continue;
      string message=entry.Message;
      if(message.IndexOf("PagePilot",StringComparison.OrdinalIgnoreCase)<0&&message.IndexOf("WINWORD",StringComparison.OrdinalIgnoreCase)<0)continue;
      found++;sb.AppendLine(entry.TimeGenerated.ToUniversalTime().ToString("o")+" · "+entry.Source+" · EventId="+entry.InstanceId);
      // Do not include arbitrary event payloads: only known loader exception/fault lines.
      foreach(string line in message.Split('\n').Where(l=>new[]{"Exception","异常","Faulting module","故障模块","出错模块","HRESULT","0x","PagePilot.dll","mscoree.dll","clr.dll"}.Any(w=>l.IndexOf(w,StringComparison.OrdinalIgnoreCase)>=0)).Take(12))sb.AppendLine(line.Length>500?line.Substring(0,500):line);
     }
     sb.AppendLine("匹配事件数="+found+"（没有事件不能证明加载成功）");
    }
   });
   return sb.ToString();
  }
 }
}
