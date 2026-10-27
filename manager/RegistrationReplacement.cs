using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Principal;

namespace PagePilot.Deployment {
 internal static partial class Manager {
  static bool replaceExisting;
  static void CheckReplacementMode(){
   if(replaceExisting&&!machineScope)throw new InvalidOperationException("替换旧注册用于整台电脑部署，请选择“整台电脑（管理员环境）”。");
  }
  static List<Entry> UserSnapshot(RegistryView view){
   using(var user=User(view))return RegistrationPaths.Select(p=>new Entry{Path=p,View=(int)view,Before=Read(user,Key(p))}).ToList();
  }
  static void CheckReplacementOwnership(RegistryKey store,Func<string,string> path){
   // Permit incomplete PagePilot registration, but never remove an explicitly foreign class.
   for(int i=0;i<Ids.Length;i++){
    using(var k=store.OpenSubKey(path(@"Software\Classes\CLSID\"+Ids[i]+@"\InprocServer32"))){
     string name=Convert.ToString(k?.GetValue("Class"));
     if(!string.IsNullOrEmpty(name)&&name!=Classes[i])throw new IOException("PagePilot CLSID 被其他组件占用，无法替换："+Ids[i]);
    }
    using(var k=store.OpenSubKey(path(@"Software\Classes\"+Classes[i]+@"\CLSID"))){
     string id=Convert.ToString(k?.GetValue(""));
     if(!string.IsNullOrEmpty(id)&&!string.Equals(id,Ids[i],StringComparison.OrdinalIgnoreCase))throw new IOException("PagePilot ProgID 指向其他组件，无法替换："+Classes[i]);
    }
   }
  }
  static void ValidateMigrationOwner(Journal journal){
   if(journal.ReplacedUserEntries!=null&&journal.UserSid!=WindowsIdentity.GetCurrent().User.Value)
    throw new InvalidOperationException("此回退包含原账户的当前用户注册，请使用执行替换的 Windows 账户以管理员身份回退，避免恢复到其他账户。");
  }
  static void ValidateMigrationAfter(Journal journal){
   ValidateMigrationOwner(journal);
   if(journal.ReplacedUserEntries==null||journal.State!="completed")return;
   foreach(var entry in journal.ReplacedUserEntries)using(var user=User((RegistryView)entry.View))
    if(!Same(Read(user,Key(entry.Path)),entry.After))throw new IOException("替换后当前用户的 PagePilot 注册又被修改，拒绝覆盖；请导出诊断包。");
  }
  static void RestoreMigratedUser(Journal journal){
   ValidateMigrationOwner(journal);
   if(journal.ReplacedUserEntries==null)return;
   foreach(var entry in journal.ReplacedUserEntries)using(var user=User((RegistryView)entry.View)){
    Write(user,Key(entry.Path),entry.Before);
    if(!Same(Read(user,Key(entry.Path)),entry.Before))throw new IOException("原账户注册恢复回读失败，保留记录以便重试回退。");
   }
   Event("registry.user.restored","sid="+journal.UserSid);
  }
  static bool ValidEntries(List<Entry> entries){
   return entries!=null&&entries.Count==RegistrationPaths.Length&&entries.All(e=>e!=null)&&
    entries.Select(e=>e.Path).Distinct(StringComparer.OrdinalIgnoreCase).Count()==RegistrationPaths.Length&&
    entries.All(e=>RegistrationPaths.Contains(e.Path,StringComparer.Ordinal)&&(e.View==(int)RegistryView.Registry32||e.View==(int)RegistryView.Registry64))&&
    entries.Select(e=>e.View).Distinct().Count()==1;
  }
 }
}
