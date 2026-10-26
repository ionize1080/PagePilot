using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web.Script.Serialization;

namespace PagePilot.Configuration {
 public sealed class RuntimeSettings {
  public int SchemaVersion {get;set;}=1;
  public string DataDirectory {get;set;}=DefaultData;
  public string ProfileDirectory {get;set;}=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),"PagePilot");
  public string LogDirectory {get;set;}=Path.Combine(DefaultData,"logs");
  public string LogLevel {get;set;}="debug";
  public static string DefaultData=>Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"PagePilot");
  public static string DefaultFile=>Path.Combine(DefaultData,"runtime.json");
  public const string LocatorKey=@"Software\PagePilot\Runtime";
  public static string ResolveFile(){using(var k=Registry.CurrentUser.OpenSubKey(LocatorKey))return Convert.ToString(k?.GetValue("ConfigFile")) is string p&&!string.IsNullOrWhiteSpace(p)?Path.GetFullPath(p):DefaultFile;}
  public static RuntimeSettings Load(string file){
   if(!File.Exists(file)){if(!string.Equals(Path.GetFullPath(file),DefaultFile,StringComparison.OrdinalIgnoreCase))throw new FileNotFoundException("找不到指定配置文件",file);return new RuntimeSettings();}
   if(new FileInfo(file).Length>65536)throw new InvalidDataException("运行配置文件超过 64 KB");
   var v=new JavaScriptSerializer().Deserialize<RuntimeSettings>(File.ReadAllText(file,Encoding.UTF8));
   if(v==null)throw new InvalidDataException("配置文件内容为空");v.Normalize(file);return v;
  }
  public void Normalize(string file){
   if(SchemaVersion!=1)throw new InvalidDataException("不支持此配置版本");
   DataDirectory=ResolveDirectory(DataDirectory,file);ProfileDirectory=ResolveDirectory(ProfileDirectory,file);LogDirectory=ResolveDirectory(LogDirectory,file);
   LogLevel=(LogLevel??"info").ToLowerInvariant();if(LogLevel!="info"&&LogLevel!="debug")throw new InvalidDataException("日志级别只能为 info 或 debug");
  }
  static string ResolveDirectory(string path,string file){if(string.IsNullOrWhiteSpace(path))throw new InvalidDataException("目录不能为空");path=Environment.ExpandEnvironmentVariables(path);if(path.Contains("%"))throw new InvalidDataException("目录包含未解析的环境变量");return Path.GetFullPath(Path.IsPathRooted(path)?path:Path.Combine(Path.GetDirectoryName(Path.GetFullPath(file)),path)).TrimEnd(Path.DirectorySeparatorChar)+Path.DirectorySeparatorChar;}
  public static void EnsureWritable(string folder){Directory.CreateDirectory(folder);string probe=Path.Combine(folder,".pagepilot-write-"+Guid.NewGuid().ToString("N"));using(var s=new FileStream(probe,FileMode.CreateNew,FileAccess.Write,FileShare.None,1,FileOptions.DeleteOnClose))s.WriteByte(1);}
  public void Save(string file,bool remember){Normalize(file);foreach(string p in new[]{DataDirectory,ProfileDirectory,LogDirectory,Path.GetDirectoryName(Path.GetFullPath(file))})EnsureWritable(p);WriteAtomic(file,new JavaScriptSerializer().Serialize(this));if(remember)using(var k=Registry.CurrentUser.CreateSubKey(LocatorKey))k.SetValue("ConfigFile",Path.GetFullPath(file));}
  public static void WriteAtomic(string file,string value){Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(file)));string temp=file+"."+Guid.NewGuid().ToString("N")+".tmp";try{File.WriteAllText(temp,value,new UTF8Encoding(false));if(File.Exists(file))File.Replace(temp,file,file+".previous",true);else File.Move(temp,file);}finally{if(File.Exists(temp))File.Delete(temp);}}
 }
 public static class RuntimePaths {
  static readonly object Gate=new object();static RuntimeSettings current;
  public static string ConfigFile {get;private set;}
  public static string Warning {get;private set;}
  public static RuntimeSettings Current {get{lock(Gate){if(current==null)Reload();return current;}}}
  public static void Reload(string explicitFile=null){lock(Gate){Warning=null;try{ConfigFile=explicitFile??RuntimeSettings.ResolveFile();current=RuntimeSettings.Load(ConfigFile);}catch(Exception e){Warning=e.GetType().Name+": "+e.Message;current=new RuntimeSettings();}}}
  public static void Set(RuntimeSettings value,string file){lock(Gate){current=value;ConfigFile=file;Warning=null;}}
 }
 public static class RuntimeLog {
  public static void Write(string component,string eventName,string detail="",string level="info"){
   try{var settings=RuntimePaths.Current;if(level=="debug"&&settings.LogLevel!="debug")return;WriteTo(settings.LogDirectory,component,eventName,detail,level);}catch(Exception e){try{WriteTo(Path.Combine(RuntimeSettings.DefaultData,"logs"),component,"log.fallback",e.GetType().Name+"; "+eventName+" "+detail,"error");}catch{}}
  }
  static void WriteTo(string folder,string component,string eventName,string detail,string level){
   Directory.CreateDirectory(folder);string file=Path.Combine(folder,component+".jsonl");string key;using(var sha=SHA256.Create())key=BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(Path.GetFullPath(file).ToUpperInvariant()))).Replace("-","");
   using(var mutex=new Mutex(false,"Local\\PagePilot.Log."+key)){
    bool entered=false;try{try{entered=mutex.WaitOne(2000);}catch(AbandonedMutexException){entered=true;}if(!entered)return;
     if(File.Exists(file)&&new FileInfo(file).Length>4*1024*1024){for(int i=3;i>=1;i--){string target=file+"."+i,source=i==1?file:file+"."+(i-1);if(File.Exists(target))File.Delete(target);if(File.Exists(source))File.Move(source,target);}}
     detail=Redact(detail??"");if(detail.Length>12000)detail=detail.Substring(0,12000)+" [truncated]";
     var row=new{utc=DateTime.UtcNow.ToString("o"),pid=Process.GetCurrentProcess().Id,component,level,eventName,detail};
     File.AppendAllText(file,new JavaScriptSerializer().Serialize(row)+Environment.NewLine,new UTF8Encoding(false));
    }finally{if(entered)mutex.ReleaseMutex();}
   }
  }
  public static string Redact(string s)=>Regex.Replace(s,@"(?i)(\b(?:api[_-]?key|token|authorization)\b\s*[=:]\s*)([^\s,;]+)","$1[redacted]");
 }
 public sealed class HostReceipt {
  public int Pid{get;set;} public long StartUtcTicks{get;set;} public string AssemblyPath{get;set;} public string Version{get;set;} public string WordVersion{get;set;} public string UserSid{get;set;} public bool Elevated{get;set;} public bool Connected{get;set;} public bool RibbonLoaded{get;set;} public string UpdatedUtc{get;set;}
 }
 public static class HostHealth {
  static bool connected,ribbon;static string wordVersion="";
  public static void Reset(string assembly,string version){connected=false;ribbon=false;Update(assembly,version);}
  public static string ReceiptPath(string folder,int pid)=>Path.Combine(folder,"state","word-"+pid+".json");
  public static void Update(string assembly,string version,string word=null,bool ready=false,bool ribbonReady=false){
   try{connected|=ready;ribbon|=ribbonReady;if(word!=null)wordVersion=word;using(var p=Process.GetCurrentProcess())using(var identity=WindowsIdentity.GetCurrent()){
    var receipt=new HostReceipt{Pid=p.Id,StartUtcTicks=p.StartTime.ToUniversalTime().Ticks,AssemblyPath=assembly,Version=version,WordVersion=wordVersion,UserSid=identity.User.Value,Elevated=new WindowsPrincipal(identity).IsInRole(WindowsBuiltInRole.Administrator),Connected=connected,RibbonLoaded=ribbon,UpdatedUtc=DateTime.UtcNow.ToString("o")};
    RuntimeSettings.WriteAtomic(ReceiptPath(RuntimePaths.Current.LogDirectory,p.Id),new JavaScriptSerializer().Serialize(receipt));}
   }catch(Exception e){RuntimeLog.Write("host","health.failed",e.GetType().Name,"error");}
  }
  public static bool Matches(HostReceipt r,int pid,long start,string assembly)=>r!=null&&r.Pid==pid&&r.StartUtcTicks==start&&string.Equals(Path.GetFullPath(r.AssemblyPath),Path.GetFullPath(assembly),StringComparison.OrdinalIgnoreCase)&&r.Connected&&r.RibbonLoaded;
 }
}
