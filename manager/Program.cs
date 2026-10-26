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
namespace PagePilot.Deployment {
 public sealed class Value {public string Name{get;set;}public int Kind{get;set;}public string[] Data{get;set;}}
 public sealed class Tree {public List<Value> Values{get;set;}=new List<Value>();public Dictionary<string,Tree> Children{get;set;}=new Dictionary<string,Tree>();}
 public sealed class Entry {public string Path{get;set;}public int View{get;set;}public Tree Before{get;set;}public Tree After{get;set;}}
 public sealed class Journal {public string Root{get;set;}public string Operation{get;set;}public List<Entry> Entries{get;set;}public string State{get;set;}}
 internal static class Manager {
  static readonly string[] Classes={"PagePilot.Connect","PagePilot.AiPane","PagePilot.Pane"};
  static readonly string[] Ids={"{0E057EBB-B35B-4A31-9B65-3DF44E8034CE}","{246818DC-B9A5-4328-9E71-69575A825B92}","{79CEEFC8-67F4-4F55-9475-24C0681BFBBD}"};
  const string Addin=@"Software\Microsoft\Office\Word\Addins\PagePilot.Connect";
  static string root=AppDomain.CurrentDomain.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar);
  static string data=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"PagePilot","deployment");
  static readonly JavaScriptSerializer json=new JavaScriptSerializer{MaxJsonLength=16*1024*1024};
  static string testPrefix="";
  static RegistryKey User(RegistryView v)=>RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,v);
  static string Key(string path)=>testPrefix+path;
  static string Dll=>Path.Combine(root,"PagePilot.dll");
  static void Closed(){if(testPrefix==""&&Process.GetProcessesByName("WINWORD").Length>0)throw new Exception("请先保存并关闭所有 Word 窗口，再切换或停用版本。");}
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
  static string CodeBase(RegistryView view){using(var u=User(view))using(var k=u.OpenSubKey(Key(@"Software\Classes\CLSID\"+Ids[0]+@"\InprocServer32")))return Convert.ToString(k?.GetValue("CodeBase"));}
  static Tree Read(RegistryKey parent,string path){using(var key=parent.OpenSubKey(path)){if(key==null)return null;var t=new Tree();foreach(string name in key.GetValueNames().OrderBy(x=>x,StringComparer.Ordinal)){var kind=key.GetValueKind(name);object v=key.GetValue(name,null,RegistryValueOptions.DoNotExpandEnvironmentNames);t.Values.Add(new Value{Name=name,Kind=(int)kind,Data=kind==RegistryValueKind.Binary||kind==RegistryValueKind.None?new[]{Convert.ToBase64String((byte[])v)}:kind==RegistryValueKind.MultiString?(string[])v:new[]{Convert.ToString(v,System.Globalization.CultureInfo.InvariantCulture)}});}foreach(string child in key.GetSubKeyNames().OrderBy(x=>x,StringComparer.Ordinal))t.Children[child]=Read(key,child);return t;}}
  static void Write(RegistryKey parent,string path,Tree tree){parent.DeleteSubKeyTree(path,false);if(tree==null)return;using(var k=parent.CreateSubKey(path)){foreach(var v in tree.Values){var kind=(RegistryValueKind)v.Kind;object value=kind==RegistryValueKind.Binary||kind==RegistryValueKind.None?(object)Convert.FromBase64String(v.Data[0]):kind==RegistryValueKind.MultiString?v.Data:kind==RegistryValueKind.DWord?(object)int.Parse(v.Data[0]):kind==RegistryValueKind.QWord?(object)long.Parse(v.Data[0]):v.Data[0];k.SetValue(v.Name,value,kind);}foreach(var child in tree.Children)Write(k,child.Key,child.Value);}}
  static bool Same(Tree a,Tree b)=>json.Serialize(a)==json.Serialize(b);
  static List<Entry> Snapshot(RegistryView view){var paths=Ids.Select(id=>@"Software\Classes\CLSID\"+id).Concat(Classes.Select(c=>@"Software\Classes\"+c)).Concat(new[]{Addin});using(var user=User(view))return paths.Select(p=>new Entry{Path=p,View=(int)view,Before=Read(user,Key(p))}).ToList();}
  static void Ownership(RegistryView view,bool sameDirectory){using(var user=User(view)){
   for(int i=0;i<Ids.Length;i++)using(var key=user.OpenSubKey(Key(@"Software\Classes\CLSID\"+Ids[i]+@"\InprocServer32"))){if(key==null)continue;if(Convert.ToString(key.GetValue("Class"))!=Classes[i])throw new Exception("CLSID 已由其他组件占用，拒绝覆盖。");if(sameDirectory&&!SameLocation(Convert.ToString(key.GetValue("CodeBase")),Dll))throw new Exception("当前启用的是其他目录的版本，请从该版本运行管理器。");}
   for(int i=0;i<Classes.Length;i++)using(var k=user.OpenSubKey(Key(@"Software\Classes\"+Classes[i]+@"\CLSID"))){if(k!=null&&Convert.ToString(k.GetValue(""))!=Ids[i])throw new Exception("ProgID 已由其他组件占用。");}
   using(var k=user.OpenSubKey(Key(Addin)))if(k!=null&&Convert.ToString(k.GetValue("FriendlyName"))!="PagePilot")throw new Exception("加载项注册名称冲突。");
  }}
  static void ValidatePackage(bool full){
   foreach(string path in new[]{"PagePilot.dll","PagePilot.Manager.exe","Python/python.exe","Browser/electron.exe","Chatbox/bootstrap.cjs","Chatbox/native/dist/main/main.js","bridge/server.py","vendor/word-mcp-live/word_document_server/main.py"})if(!File.Exists(Path.Combine(root,path)))throw new Exception("运行包不完整，缺少 "+path);
   if(AssemblyName.GetAssemblyName(Dll).Name!="PagePilot")throw new Exception("加载项程序集身份不正确");
   string manifest=Path.Combine(root,"manifest.sha256");if(!File.Exists(manifest))throw new Exception("缺少 manifest.sha256，请使用完整发布包。");
   foreach(string line in File.ReadLines(manifest)){if(line.Length<67)throw new Exception("完整性清单格式无效");string relative=line.Substring(66).Replace('/',Path.DirectorySeparatorChar),path=Path.GetFullPath(Path.Combine(root,relative));if(!path.StartsWith(root+Path.DirectorySeparatorChar,StringComparison.OrdinalIgnoreCase))throw new Exception("清单路径越界");if(!full&&relative!="PagePilot.dll"&&relative!="PagePilot.Manager.exe")continue;using(var stream=File.OpenRead(path))using(var sha=SHA256.Create())if(BitConverter.ToString(sha.ComputeHash(stream)).Replace("-","").ToLowerInvariant()!=line.Substring(0,64))throw new Exception("文件校验失败："+relative);}
  }
  static void Atomic(string path,Journal journal){Directory.CreateDirectory(Path.GetDirectoryName(path));string temp=path+".tmp";File.WriteAllText(temp,json.Serialize(journal),new UTF8Encoding(false));if(File.Exists(path))File.Replace(temp,path,path+".previous",true);else File.Move(temp,path);}
  static string JournalPath=>Path.Combine(data,"last.json");
  static void RecoverPending(){if(!File.Exists(JournalPath))return;var j=json.Deserialize<Journal>(File.ReadAllText(JournalPath));if(j.State=="pending")throw new Exception("检测到上次部署中断。请先选择“回退注册”，恢复后再启用。");}
  static void Change(string operation,RegistryView view,bool injectFailure=false){
   Closed();RecoverPending();Ownership(view,operation=="disable");if(operation=="enable"&&testPrefix=="")ValidatePackage(true);
   var j=new Journal{Root=root,Operation=operation,State="pending",Entries=Snapshot(view)};
   if(File.Exists(JournalPath))File.Copy(JournalPath,Path.Combine(data,DateTime.UtcNow.ToString("yyyyMMdd-HHmmss-fffffff")+".json"));Atomic(JournalPath,j);
   try{using(var user=User(view)){
    if(operation=="disable"){foreach(var e in j.Entries)user.DeleteSubKeyTree(Key(e.Path),false);}
    else{
     string assembly=testPrefix==""?AssemblyName.GetAssemblyName(Dll).FullName:"PagePilot, Version=0.4.0.0, Culture=neutral, PublicKeyToken=null";
     for(int i=0;i<Classes.Length;i++){
      string path=@"Software\Classes\CLSID\"+Ids[i];using(var k=user.CreateSubKey(Key(path)))k.SetValue("",Classes[i]);
      using(var k=user.CreateSubKey(Key(path+@"\InprocServer32"))){k.SetValue("","mscoree.dll");k.SetValue("ThreadingModel","Both");k.SetValue("Class",Classes[i]);k.SetValue("Assembly",assembly);k.SetValue("RuntimeVersion","v4.0.30319");k.SetValue("CodeBase",Dll);}
      using(var k=user.CreateSubKey(Key(path+@"\ProgId")))k.SetValue("",Classes[i]);
      using(var k=user.CreateSubKey(Key(@"Software\Classes\"+Classes[i]+@"\CLSID")))k.SetValue("",Ids[i]);
      if(i>0){user.CreateSubKey(Key(path+@"\Control")).Dispose();user.CreateSubKey(Key(path+@"\Implemented Categories\{40FC6ED4-2438-11CF-A3DB-080036F12502}")).Dispose();}
      if(injectFailure&&i==1)throw new IOException("Injected registry failure");
     }
     using(var k=user.CreateSubKey(Key(Addin))){k.SetValue("FriendlyName","PagePilot");k.SetValue("Description","PagePilot · 当前用户半便携部署");k.SetValue("LoadBehavior",3,RegistryValueKind.DWord);}
    }
    foreach(var e in j.Entries)e.After=Read(user,Key(e.Path));
   }
   if(operation=="enable"&&!SameLocation(CodeBase(view),Dll))throw new Exception("注册回读失败");j.State="completed";Atomic(JournalPath,j);
   }catch{foreach(var e in j.Entries)using(var u=User((RegistryView)e.View))Write(u,Key(e.Path),e.Before);j.State="rolled_back";Atomic(JournalPath,j);throw;}
  }
  static void Rollback(){Closed();if(!File.Exists(JournalPath))throw new Exception("没有可回退的部署记录");var j=json.Deserialize<Journal>(File.ReadAllText(JournalPath));if(j.State!="completed"&&j.State!="pending")throw new Exception("上次部署已回退");foreach(var e in j.Entries)using(var u=User((RegistryView)e.View)){if(j.State=="completed"&&!Same(Read(u,Key(e.Path)),e.After))throw new Exception("注册已被其他操作修改，拒绝覆盖。请运行诊断。");}
   // Before restoring an older version, make sure its DLL still exists.
   foreach(var e in j.Entries){Tree t=e.Before;if(t!=null&&t.Children.TryGetValue("InprocServer32",out var server)){var code=server.Values.FirstOrDefault(v=>v.Name=="CodeBase");if(code!=null&&testPrefix==""&&!File.Exists(new Uri(code.Data[0]).LocalPath))throw new Exception("上一版本已移动或删除，无法回退。请恢复其目录。");}}
   foreach(var e in j.Entries)using(var u=User((RegistryView)e.View)){Write(u,Key(e.Path),e.Before);if(!Same(Read(u,Key(e.Path)),e.Before))throw new Exception("回退回读失败，请保留部署记录并运行诊断");}j.State="rolled_back";Atomic(JournalPath,j);
  }
  static string Status(bool doctor){var sb=new StringBuilder();sb.AppendLine("PagePilot 管理器 0.4.0");sb.AppendLine("当前目录："+root);try{var view=WordView();string word=DetectWord();sb.AppendLine("Word："+FileVersionInfo.GetVersionInfo(word).FileVersion+" · "+(view==RegistryView.Registry64?"64":"32")+" 位");sb.AppendLine("注册路径："+(CodeBase(view)==""?"未启用":CodeBase(view)));using(var u=User(view))using(var k=u.OpenSubKey(Key(Addin)))sb.AppendLine("加载状态 LoadBehavior："+Convert.ToString(k?.GetValue("LoadBehavior")));}catch(Exception e){sb.AppendLine(e.Message);}try{ValidatePackage(doctor);sb.AppendLine(doctor?"完整文件校验通过":"关键文件校验通过");}catch(Exception e){sb.AppendLine("需要处理："+e.Message);}sb.AppendLine("配置与会话：%APPDATA%\\PagePilot");sb.AppendLine("本地规则、日志、文档备份：%LOCALAPPDATA%\\PagePilot");sb.AppendLine("启用、停用及回退不会删除这些用户数据。");if(Process.GetProcessesByName("WINWORD").Length>0)sb.AppendLine("Word 正在运行，关闭后才能切换版本。");if(doctor)sb.AppendLine("若 Word 已禁用加载项，请在 Word 加载项界面检查原因后手动恢复。管理器不修改 DisabledItems。");return sb.ToString();}
  static void SelfTest(){string originalData=data,originalRoot=root;testPrefix=@"Software\PagePilot\DeploymentTests\"+Guid.NewGuid().ToString("N")+@"\";data=Path.Combine(Path.GetTempPath(),"PagePilot-Deployment-"+Guid.NewGuid().ToString("N"));var view=RegistryView.Registry64;
   try{root=Path.Combine(data,"v1");Change("enable",view);string v1=CodeBase(view);root=Path.Combine(data,"v2");Change("enable",view);if(CodeBase(view)==v1)throw new Exception("Upgrade failed");Rollback();if(CodeBase(view)!=v1)throw new Exception("Rollback failed");try{Change("enable",view,true);throw new Exception("Failure injection missing");}catch(IOException){}if(CodeBase(view)!=v1)throw new Exception("Failed update rollback failed");root=Path.Combine(data,"v1");Change("disable",view);if(CodeBase(view)!="")throw new Exception("Disable failed");Rollback();if(CodeBase(view)!=v1)throw new Exception("Disable rollback failed");File.WriteAllText(Path.Combine(originalRoot,"manager-selftest.txt"),"PASS enable / upgrade / rollback / failure injection / disable / restore; isolated HKCU test subtree only");}
   finally{using(var u=User(view))u.DeleteSubKeyTree(testPrefix.TrimEnd('\\'),false);testPrefix="";data=originalData;root=originalRoot;}
  }
  [STAThread]static int Main(string[] args){Application.EnableVisualStyles();try{
   for(int i=0;i<args.Length;i++)if(args[i]=="--root"&&i+1<args.Length)root=Path.GetFullPath(args[++i]).TrimEnd(Path.DirectorySeparatorChar);
   string command=args.Length>0?args[0].TrimStart('-','/').ToLowerInvariant():"ui";
   if(command=="self-test"){SelfTest();return 0;}
   if(command=="status"||command=="doctor"){string report=Status(command=="doctor");int output=Array.IndexOf(args,"--output");if(output>=0&&output+1<args.Length)File.WriteAllText(Path.GetFullPath(args[output+1]),report,new UTF8Encoding(false));else MessageBox.Show(report,"PagePilot · 状态");return 0;}
   if(command=="enable"||command=="disable"){Change(command,WordView());return 0;}if(command=="rollback"){Rollback();return 0;}if(command!="ui")throw new Exception("支持 enable、disable、status、doctor、rollback、--root、--output。");
   float scale;using(var graphics=Graphics.FromHwnd(IntPtr.Zero))scale=graphics.DpiX/96f;Func<int,int> px=value=>(int)(value*scale);
   using(var form=new Form{Text="PagePilot · 版本与部署",Width=px(760),Height=px(560),MinimumSize=new Size(px(640),px(480)),Font=new Font("Microsoft YaHei UI",10),StartPosition=FormStartPosition.CenterScreen,AutoScaleMode=AutoScaleMode.None}){
    var panel=new FlowLayoutPanel{Dock=DockStyle.Bottom,Height=px(112),Padding=new Padding(px(12)),AutoScroll=true};var report=new TextBox{Dock=DockStyle.Fill,Multiline=true,ReadOnly=true,ScrollBars=ScrollBars.Vertical,BorderStyle=BorderStyle.None,Text=Status(false)};var body=new Panel{Dock=DockStyle.Fill,Padding=new Padding(px(18))};body.Controls.Add(report);form.Controls.Add(body);form.Controls.Add(panel);form.Shown+=(s,e)=>{report.SelectionStart=0;report.SelectionLength=0;panel.Controls[0].Focus();};
    Action<string,Func<string>> button=(label,action)=>{var b=new Button{Text=label,AutoSize=true,Height=px(38),Padding=new Padding(px(10),px(4),px(10),px(4)),Margin=new Padding(px(4))};b.Click+=async(s,e)=>{panel.Enabled=false;report.Text="正在处理，请稍候…";try{string result=await Task.Run(action);report.Text=result??Status(false);report.SelectionStart=0;report.SelectionLength=0;}catch(Exception ex){report.Text=Status(false);MessageBox.Show(form,ex.Message,"PagePilot",MessageBoxButtons.OK,MessageBoxIcon.Information);}finally{panel.Enabled=true;}};panel.Controls.Add(b);};
    button("启用此版本",()=>{Change("enable",WordView());return null;});button("停用此版本",()=>{Change("disable",WordView());return null;});button("回退注册",()=>{Rollback();return null;});button("检查完整文件",()=>Status(true));
    var choose=new Button{Text="选择版本目录",AutoSize=true,Padding=new Padding(px(10),px(4),px(10),px(4)),Margin=new Padding(px(4))};choose.Click+=(s,e)=>{using(var dialog=new FolderBrowserDialog{Description="选择已完整解压的 PagePilot 版本目录",SelectedPath=root})if(dialog.ShowDialog(form)==DialogResult.OK){root=Path.GetFullPath(dialog.SelectedPath).TrimEnd(Path.DirectorySeparatorChar);report.Text=Status(false);}};panel.Controls.Add(choose);Application.Run(form);
   }return 0;
  }catch(Exception e){string message=e.Message;int output=Array.IndexOf(args,"--output");if(output>=0&&output+1<args.Length)File.WriteAllText(Path.GetFullPath(args[output+1]),message);else MessageBox.Show(message,"PagePilot · 未完成",MessageBoxButtons.OK,MessageBoxIcon.Warning);return 1;}}
 }
}
