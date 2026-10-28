using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using System.Threading.Tasks;
using PagePilot.Configuration;
namespace PagePilot.Deployment {
 public sealed class Value {public string Name{get;set;}public int Kind{get;set;}public string[] Data{get;set;}}
 public sealed class Tree {public List<Value> Values{get;set;}=new List<Value>();public Dictionary<string,Tree> Children{get;set;}=new Dictionary<string,Tree>();}
 public sealed class Entry {public string Path{get;set;}public int View{get;set;}public Tree Before{get;set;}public Tree After{get;set;}}
 public sealed class Journal {public List<Entry> ReplacedUserEntries{get;set;}public string UserSid{get;set;}public string Scope{get;set;}public string Root{get;set;}public string Operation{get;set;}public List<Entry> Entries{get;set;}public string State{get;set;}}
 internal static partial class Manager {
  static readonly string[] Classes={"PagePilot.Connect","PagePilot.AiPane","PagePilot.Pane"};
  static readonly string[] Ids={"{0E057EBB-B35B-4A31-9B65-3DF44E8034CE}","{246818DC-B9A5-4328-9E71-69575A825B92}","{79CEEFC8-67F4-4F55-9475-24C0681BFBBD}"};
  const string Addin=@"Software\Microsoft\Office\Word\Addins\PagePilot.Connect";
  static string root=AppDomain.CurrentDomain.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar);
  static string data=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"PagePilot","deployment");
  static readonly JavaScriptSerializer json=new JavaScriptSerializer{MaxJsonLength=16*1024*1024};
  static string testPrefix="";static bool testRollbackFailure=false;
  static RegistryKey User(RegistryView v)=>RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,v);
  static string Key(string path)=>testPrefix+path;
  static string Dll=>Path.Combine(root,"PagePilot.dll");
  static void Closed(){if(testPrefix==""&&(machineScope?Process.GetProcessesByName("WINWORD"):WordProcesses()).Length>0)throw new Exception(machineScope?"Word 仍在运行。机器级操作要求所有账户、所有会话的 Word 退出；请先保存并关闭。":"Word 仍在运行。请先保存并关闭。");}
  static string DetectWord(){
   foreach(var view in new[]{RegistryView.Registry64,RegistryView.Registry32})foreach(var hive in new[]{RegistryHive.LocalMachine,RegistryHive.CurrentUser})using(var baseKey=RegistryKey.OpenBaseKey(hive,view))using(var key=baseKey.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE")){
    string path=Convert.ToString(key?.GetValue(""));if(File.Exists(path))return path;
   }
   foreach(var view in new[]{RegistryView.Registry64,RegistryView.Registry32})using(var machine=RegistryKey.OpenBaseKey(RegistryHive.LocalMachine,view)){
    using(var click=machine.OpenSubKey(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration")){string install=Convert.ToString(click?.GetValue("InstallationPath"));if(!string.IsNullOrEmpty(install)){string path=Path.Combine(install,"root","Office16","WINWORD.EXE");if(File.Exists(path))return path;}}
    foreach(string version in new[]{"16.0","15.0","14.0"})using(var install=machine.OpenSubKey(@"SOFTWARE\Microsoft\Office\"+version+@"\Word\InstallRoot")){string folder=Convert.ToString(install?.GetValue("Path"));if(!string.IsNullOrEmpty(folder)&&File.Exists(Path.Combine(folder,"WINWORD.EXE")))return Path.Combine(folder,"WINWORD.EXE");}
   }
   throw new Exception("未找到已安装的 Microsoft Word。请先运行一次 Word，完成 Office 的安装配置。");
  }
  static RegistryView WordView(){string word=DetectWord();using(var stream=File.OpenRead(word))using(var reader=new BinaryReader(stream)){stream.Position=0x3c;int offset=reader.ReadInt32();stream.Position=offset;if(reader.ReadUInt32()!=0x4550)throw new Exception("Word 可执行文件格式无效");ushort machine=reader.ReadUInt16();if(machine==0x8664)return RegistryView.Registry64;if(machine==0x14c)return RegistryView.Registry32;throw new Exception("当前 Word 架构不受支持");}}
  static bool SameLocation(string stored,string path){try{return string.Equals(new Uri(stored).LocalPath,Path.GetFullPath(path),StringComparison.OrdinalIgnoreCase);}catch{return false;}}
  static string CodeBase(RegistryView view){using(var u=Store(view))using(var k=u.OpenSubKey(StorePath(@"Software\Classes\CLSID\"+Ids[0]+@"\InprocServer32")))return Convert.ToString(k?.GetValue("CodeBase"));}
  static Tree Read(RegistryKey parent,string path){using(var key=parent.OpenSubKey(path)){if(key==null)return null;var t=new Tree();foreach(string name in key.GetValueNames().OrderBy(x=>x,StringComparer.Ordinal)){var kind=key.GetValueKind(name);object v=key.GetValue(name,null,RegistryValueOptions.DoNotExpandEnvironmentNames);t.Values.Add(new Value{Name=name,Kind=(int)kind,Data=kind==RegistryValueKind.Binary||kind==RegistryValueKind.None?new[]{Convert.ToBase64String((byte[])v)}:kind==RegistryValueKind.MultiString?(string[])v:new[]{Convert.ToString(v,System.Globalization.CultureInfo.InvariantCulture)}});}foreach(string child in key.GetSubKeyNames().OrderBy(x=>x,StringComparer.Ordinal))t.Children[child]=Read(key,child);return t;}}
  static void Write(RegistryKey parent,string path,Tree tree){parent.DeleteSubKeyTree(path,false);if(tree==null)return;using(var k=parent.CreateSubKey(path)){foreach(var v in tree.Values){var kind=(RegistryValueKind)v.Kind;object value=kind==RegistryValueKind.Binary||kind==RegistryValueKind.None?(object)Convert.FromBase64String(v.Data[0]):kind==RegistryValueKind.MultiString?v.Data:kind==RegistryValueKind.DWord?(object)int.Parse(v.Data[0]):kind==RegistryValueKind.QWord?(object)long.Parse(v.Data[0]):v.Data[0];k.SetValue(v.Name,value,kind);}foreach(var child in tree.Children)Write(k,child.Key,child.Value);}}
  static bool Same(Tree a,Tree b)=>json.Serialize(a)==json.Serialize(b);
  static List<Entry> Snapshot(RegistryView view){using(var user=Store(view))return RegistrationPaths.Select(p=>new Entry{Path=p,View=(int)view,Before=Read(user,StorePath(p))}).ToList();}
  static void Ownership(RegistryView view,bool sameDirectory){using(var user=Store(view)){
   for(int i=0;i<Ids.Length;i++)using(var key=user.OpenSubKey(StorePath(@"Software\Classes\CLSID\"+Ids[i]+@"\InprocServer32"))){if(key==null)continue;if(Convert.ToString(key.GetValue("Class"))!=Classes[i])throw new Exception("CLSID 已由其他组件占用，拒绝覆盖。");if(sameDirectory&&!SameLocation(Convert.ToString(key.GetValue("CodeBase")),Dll))throw new Exception("当前启用的是其他目录的版本，请从该版本运行管理器。");}
   for(int i=0;i<Classes.Length;i++)using(var k=user.OpenSubKey(StorePath(@"Software\Classes\"+Classes[i]+@"\CLSID"))){if(k!=null&&Convert.ToString(k.GetValue(""))!=Ids[i])throw new Exception("ProgID 已由其他组件占用。");}
   using(var k=user.OpenSubKey(StorePath(Addin)))if(k!=null&&Convert.ToString(k.GetValue("FriendlyName"))!="PagePilot")throw new Exception("加载项注册名称冲突。");
  }}
  static void ValidatePackage(bool full){
   foreach(string path in new[]{"PagePilot.dll","PagePilot.Manager.exe","Python/python.exe","Browser/electron.exe","Chatbox/bootstrap.cjs","Chatbox/native/dist/main/main.js","bridge/server.py","vendor/word-mcp-live/word_document_server/main.py"})if(!File.Exists(Path.Combine(root,path)))throw new Exception("运行包不完整，缺少 "+path);
   if(AssemblyName.GetAssemblyName(Dll).Name!="PagePilot")throw new Exception("加载项程序集身份不正确");
   string manifest=Path.Combine(root,"manifest.sha256");if(!File.Exists(manifest))throw new Exception("缺少 manifest.sha256，请使用完整发布包。");
   foreach(string line in File.ReadLines(manifest)){if(line.Length<67)throw new Exception("完整性清单格式无效");string relative=line.Substring(66).Replace('/',Path.DirectorySeparatorChar),path=Path.GetFullPath(Path.Combine(root,relative));if(!path.StartsWith(root+Path.DirectorySeparatorChar,StringComparison.OrdinalIgnoreCase))throw new Exception("清单路径越界");if(!full&&relative!="PagePilot.dll"&&relative!="PagePilot.Manager.exe")continue;using(var stream=File.OpenRead(path))using(var sha=SHA256.Create())if(BitConverter.ToString(sha.ComputeHash(stream)).Replace("-","").ToLowerInvariant()!=line.Substring(0,64))throw new Exception("文件校验失败："+relative);}
  }
  static void Atomic(string path,Journal journal){Directory.CreateDirectory(Path.GetDirectoryName(path));string temp=path+".tmp";File.WriteAllText(temp,json.Serialize(journal),new UTF8Encoding(false));SecureJournalFile(temp);if(File.Exists(path))File.Replace(temp,path,path+".previous",true);else File.Move(temp,path);}
  static string JournalPath=>Path.Combine(JournalDirectory,"last.json");
  static void RecoverPending(){if(!File.Exists(JournalPath))return;var j=ReadJournal();if(j.State=="pending")throw new Exception("检测到上次部署中断。请先选择“回退注册”，恢复后再启用。");}
  static void Change(string operation,RegistryView view,bool injectFailure=false){
   Event("deploy.begin",operation+" root="+root+" scope="+ScopeName+" registry="+view);
   RequireScopePermission();if(operation=="enable")CheckReplacementMode();Closed();PrepareJournalDirectory();RecoverPending();
   bool replacing=operation=="enable"&&replaceExisting;
   List<Entry> previousUser=null;
   if(replacing){
    using(var target=Store(view))CheckReplacementOwnership(target,StorePath);
    using(var user=User(view))CheckReplacementOwnership(user,Key);
    previousUser=UserSnapshot(view);if(previousUser.All(e=>e.Before==null))previousUser=null;
   }else{if(operation=="enable")CheckOtherScope(view);Ownership(view,operation=="disable");}
   if(operation=="enable"&&testPrefix=="")ValidatePackage(true);
   if(operation=="enable"&&!injectFailure&&previousUser==null&&IsRegistered(view)){Event("deploy.unchanged","已经注册此版本，保留上一份回退记录");return;}
   var j=new Journal{Scope=ScopeName,Root=root,Operation=operation,State="pending",Entries=Snapshot(view),ReplacedUserEntries=previousUser,UserSid=previousUser==null?null:System.Security.Principal.WindowsIdentity.GetCurrent().User.Value};
   if(operation=="disable"&&j.Entries.All(e=>e.Before==null)){Event("deploy.unchanged","此范围已停用，保留回退记录");return;}
   if(File.Exists(JournalPath)){string archive=Path.Combine(JournalDirectory,DateTime.UtcNow.ToString("yyyyMMdd-HHmmss-fffffff")+".json");File.Copy(JournalPath,archive);SecureJournalFile(archive);}Atomic(JournalPath,j);
   try{
    if(replacing){
     if(previousUser!=null)using(var old=User(view))foreach(var entry in previousUser){old.DeleteSubKeyTree(Key(entry.Path),false);if(Read(old,Key(entry.Path))!=null)throw new IOException("旧的当前用户注册未能清除。");}
     using(var target=Store(view))foreach(var entry in j.Entries)target.DeleteSubKeyTree(StorePath(entry.Path),false);
     Event("registry.replaced","view="+view+" previousUser="+(previousUser!=null)+" sid="+j.UserSid);
    }
    using(var user=Store(view)){
    if(operation=="disable"){foreach(var e in j.Entries)user.DeleteSubKeyTree(StorePath(e.Path),false);}
    else{
     string assembly=testPrefix==""?AssemblyName.GetAssemblyName(Dll).FullName:"PagePilot, Version=0.4.7.0, Culture=neutral, PublicKeyToken=null";
     for(int i=0;i<Classes.Length;i++){
      string path=@"Software\Classes\CLSID\"+Ids[i];using(var k=user.CreateSubKey(StorePath(path)))k.SetValue("",Classes[i]);
      // Match RegAsm's CodeBase representation. CLR COM activation treats escaped spaces differently.
      using(var k=user.CreateSubKey(StorePath(path+@"\InprocServer32"))){k.SetValue("","mscoree.dll");k.SetValue("ThreadingModel","Both");k.SetValue("Class",Classes[i]);k.SetValue("Assembly",assembly);k.SetValue("RuntimeVersion","v4.0.30319");k.SetValue("CodeBase",new Uri(Dll).ToString());}
      using(var k=user.CreateSubKey(StorePath(path+@"\ProgId")))k.SetValue("",Classes[i]);
      using(var k=user.CreateSubKey(StorePath(@"Software\Classes\"+Classes[i]+@"\CLSID")))k.SetValue("",Ids[i]);
      if(i>0){user.CreateSubKey(StorePath(path+@"\Control")).Dispose();user.CreateSubKey(StorePath(path+@"\Implemented Categories\{40FC6ED4-2438-11CF-A3DB-080036F12502}")).Dispose();}
      Event("registry.class.written",Classes[i]+" view="+view);
      if(injectFailure&&i==1)throw new IOException("Injected registry failure");
     }
     using(var k=user.CreateSubKey(StorePath(Addin))){k.SetValue("FriendlyName","PagePilot");k.SetValue("Description","PagePilot · "+ScopeLabel+"准便携部署");k.SetValue("LoadBehavior",3,RegistryValueKind.DWord);}
    }
    foreach(var e in j.Entries)e.After=Read(user,StorePath(e.Path));
   }
   if(operation=="enable"&&!IsRegistered(view))throw new Exception("注册回读失败");j.State="completed";Atomic(JournalPath,j);Event("deploy.completed",operation+"; registration verified; Word load not yet verified");
   }catch(Exception ex){Event("deploy.failed",ex.ToString(),"error");foreach(var e in j.Entries)using(var u=Store((RegistryView)e.View))Write(u,StorePath(e.Path),e.Before);RestoreMigratedUser(j);j.State="rolled_back";Atomic(JournalPath,j);Event("deploy.recovered","原注册已恢复");throw;}
  }
  static bool MatchesAfter(Entry entry,Tree current){
   if(Same(current,entry.After))return true;
   if(entry.Path!=Addin||current==null||entry.After==null)return false;
   var actual=current.Values.FirstOrDefault(v=>v.Name=="LoadBehavior");var expected=entry.After.Values.FirstOrDefault(v=>v.Name=="LoadBehavior");
   // Word clears the connected bit after a load failure. Permit only this known transition.
   if(actual?.Kind!=4||expected?.Kind!=4||actual.Data.Length!=1||expected.Data.Length!=1||actual.Data[0]!="2"||expected.Data[0]!="3")return false;
   actual.Data=new[]{"3"};return Same(current,entry.After);
  }
  static void Rollback(){Event("rollback.begin","scope="+ScopeName);RequireScopePermission();Closed();PrepareJournalDirectory();if(!File.Exists(JournalPath))throw new Exception("没有可回退的部署记录");var j=ReadJournal();ValidateMigrationAfter(j);if(j.State!="completed"&&j.State!="pending")throw new Exception("上次部署已回退");foreach(var e in j.Entries)using(var u=Store((RegistryView)e.View)){if(j.State=="completed"&&!MatchesAfter(e,Read(u,StorePath(e.Path))))throw new Exception("注册已被其他操作修改，拒绝覆盖。请运行诊断。");}
   if(j.ReplacedUserEntries==null&&j.Entries.Any(e=>e.Before!=null))CheckOtherScope((RegistryView)j.Entries[0].View);
   // Before restoring an older version, make sure its DLL still exists.
   foreach(var e in j.Entries.Concat(j.ReplacedUserEntries??new List<Entry>())){Tree t=e.Before;if(t!=null&&t.Children.TryGetValue("InprocServer32",out var server)){var code=server.Values.FirstOrDefault(v=>v.Name=="CodeBase");if(code!=null&&testPrefix==""&&!File.Exists(new Uri(code.Data[0]).LocalPath))throw new Exception("上一版本已移动或删除，无法回退。请恢复其目录。");}}
   j.State="pending";Atomic(JournalPath,j);
   try{int restored=0;foreach(var e in j.Entries)using(var u=Store((RegistryView)e.View)){
    Write(u,StorePath(e.Path),e.Before);if(!Same(Read(u,StorePath(e.Path)),e.Before))throw new Exception("回退回读失败，请保留部署记录并运行诊断");
    if(testPrefix!=""&&testRollbackFailure&&++restored==2)throw new IOException("Injected rollback interruption");
   }RestoreMigratedUser(j);j.State="rolled_back";Atomic(JournalPath,j);Event("rollback.completed");}
   catch(Exception ex){Event("rollback.failed",ex.ToString(),"error");throw;}
  }
  static void SelfTest(){string originalData=data,originalRoot=root;testPrefix=@"Software\PagePilot\DeploymentTests\"+Guid.NewGuid().ToString("N")+@"\";data=Path.Combine(Path.GetTempPath(),"PagePilot-Deployment-"+Guid.NewGuid().ToString("N"));var view=RegistryView.Registry64;
   try{root=Path.Combine(data,"v1");Change("enable",view);string v1=CodeBase(view);root=Path.Combine(data,"v2");Change("enable",view);if(CodeBase(view)==v1)throw new Exception("Upgrade failed");Rollback();if(CodeBase(view)!=v1)throw new Exception("Rollback failed");try{Change("enable",view,true);throw new Exception("Failure injection missing");}catch(IOException){}if(CodeBase(view)!=v1)throw new Exception("Failed update rollback failed");root=Path.Combine(data,"v1");Change("disable",view);if(CodeBase(view)!="")throw new Exception("Disable failed");Rollback();if(CodeBase(view)!=v1)throw new Exception("Disable rollback failed");File.WriteAllText(Path.Combine(originalRoot,"manager-selftest.txt"),"PASS enable / upgrade / rollback / failure injection / disable / restore; isolated HKCU test subtree only");}
   finally{using(var u=User(view))u.DeleteSubKeyTree(testPrefix.TrimEnd('\\'),false);testPrefix="";data=originalData;root=originalRoot;}
  }
 }
}
