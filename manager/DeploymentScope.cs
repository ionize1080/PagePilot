using Microsoft.Win32;
using System;
using System.IO;
using System.Linq;
using System.Security.AccessControl;
using System.Security.Principal;

namespace PagePilot.Deployment {
 internal static partial class Manager {
  static bool machineScope;
  static string ScopeName=>machineScope?"machine":"user";
  static string ScopeLabel=>machineScope?"整台电脑（管理员环境）":"当前用户";
  static string JournalDirectory=>machineScope
   ?(testPrefix!=""?Path.Combine(data,"machine"):Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),"PagePilot.Deployment"))
   :data;
  // Tests exercise both stores in independent HKCU subtrees, never real COM keys.
  static RegistryKey Store(RegistryView view)=>RegistryKey.OpenBaseKey(machineScope&&testPrefix==""?RegistryHive.LocalMachine:RegistryHive.CurrentUser,view);
  static string StorePath(string path)=>testPrefix+(machineScope&&testPrefix!=""?@"Machine\":"")+path;
  static string[] RegistrationPaths=>Ids.Select(id=>@"Software\Classes\CLSID\"+id).Concat(Classes.Select(c=>@"Software\Classes\"+c)).Concat(new[]{Addin}).ToArray();
  static void SelectScope(string value){
   if(value!="user"&&value!="machine")throw new ArgumentException("--scope 只能是 user 或 machine。");
   machineScope=value=="machine";
  }
  static void RequireScopePermission(){
   if(machineScope&&testPrefix==""&&!Elevated)throw new InvalidOperationException("整台电脑注册、停用和回退需要管理员权限。请关闭管理器，右键 PagePilot.Manager.exe → 以管理员身份运行，并选择“整台电脑（管理员环境）”。");
  }
  static void CheckOtherScope(RegistryView view){
   using(var other=RegistryKey.OpenBaseKey(machineScope||testPrefix!=""?RegistryHive.CurrentUser:RegistryHive.LocalMachine,view)){
    foreach(string path in RegistrationPaths){
     string otherPath=testPrefix+(!machineScope&&testPrefix!=""?@"Machine\":"")+path;
     using(var key=other.OpenSubKey(otherPath))if(key!=null)
      throw new InvalidOperationException("另一部署范围仍有 PagePilot 注册，可能遮蔽本次加载。请选择整台电脑范围，勾选“替换旧注册”后启用，即可备份并清除冲突注册。");
    }
   }
  }
  static readonly SecurityIdentifier AdminSid=new SecurityIdentifier(WellKnownSidType.BuiltinAdministratorsSid,null);
  static readonly SecurityIdentifier SystemSid=new SecurityIdentifier(WellKnownSidType.LocalSystemSid,null);
  static void RejectLinkedPath(string path){
   for(var dir=new DirectoryInfo(path);dir!=null;dir=dir.Parent)
    if(dir.Exists&&(dir.Attributes&FileAttributes.ReparsePoint)!=0)throw new IOException("部署记录路径不能包含链接或重解析点："+dir.FullName);
  }
  static void PrepareJournalDirectory(){
   if(!machineScope||testPrefix!=""){Directory.CreateDirectory(JournalDirectory);return;}
   // Elevated rollback must not consume a journal writable by an ordinary account.
   RejectLinkedPath(JournalDirectory);
   if(!Directory.Exists(JournalDirectory)){
    var acl=new DirectorySecurity();acl.SetAccessRuleProtection(true,false);acl.SetOwner(AdminSid);
    foreach(var sid in new[]{AdminSid,SystemSid})acl.AddAccessRule(new FileSystemAccessRule(sid,FileSystemRights.FullControl,InheritanceFlags.ContainerInherit|InheritanceFlags.ObjectInherit,PropagationFlags.None,AccessControlType.Allow));
    Directory.CreateDirectory(JournalDirectory,acl);
   }
   CheckJournalSecurity(new DirectoryInfo(JournalDirectory).GetAccessControl());
   foreach(var file in new DirectoryInfo(JournalDirectory).GetFileSystemInfos()){
    if((file.Attributes&FileAttributes.ReparsePoint)!=0||file is DirectoryInfo)throw new IOException("部署记录目录包含非预期项目："+file.Name);
    CheckJournalSecurity(((FileInfo)file).GetAccessControl());
   }
  }
  static void CheckJournalSecurity(FileSystemSecurity acl){
   var owner=acl.GetOwner(typeof(SecurityIdentifier));
   if(!owner.Equals(AdminSid)&&!owner.Equals(SystemSid))throw new IOException("机器级部署记录所有者异常，请由管理员检查 "+JournalDirectory);
   const FileSystemRights writes=FileSystemRights.Write|FileSystemRights.Delete|FileSystemRights.DeleteSubdirectoriesAndFiles|FileSystemRights.ChangePermissions|FileSystemRights.TakeOwnership;
   var rules=acl.GetAccessRules(true,true,typeof(SecurityIdentifier));
   if(rules.Count==0)throw new IOException("机器级部署记录缺少受限访问规则。");
   foreach(FileSystemAccessRule rule in rules)
    if(rule.AccessControlType==AccessControlType.Allow&&(rule.FileSystemRights&writes)!=0&&!rule.IdentityReference.Equals(AdminSid)&&!rule.IdentityReference.Equals(SystemSid))
     throw new IOException("机器级部署记录允许非管理员写入，拒绝使用："+JournalDirectory);
  }
  static void SecureJournalFile(string path){
   if(!machineScope||testPrefix!="")return;
   var acl=new FileInfo(path).GetAccessControl();acl.SetOwner(AdminSid);File.SetAccessControl(path,acl);
  }
  static Journal ReadJournal(){
   var j=json.Deserialize<Journal>(File.ReadAllText(JournalPath));
   if(j==null||(j.Scope??"user")!=ScopeName||!ValidEntries(j.Entries)||
      (j.ReplacedUserEntries!=null&&(!machineScope||string.IsNullOrEmpty(j.UserSid)||!ValidEntries(j.ReplacedUserEntries)||
       j.ReplacedUserEntries[0].View!=j.Entries[0].View||j.ReplacedUserEntries.Any(e=>e.After!=null))))
    throw new IOException("部署记录范围或注册路径无效，拒绝回退。");
   return j;
  }
 }
}
