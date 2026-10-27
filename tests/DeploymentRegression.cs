using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Web.Script.Serialization;
using Microsoft.Win32;
using PagePilot.Configuration;
using PagePilot.Deployment;

public static class DeploymentTests {
 static int count;static readonly BindingFlags Flags=BindingFlags.NonPublic|BindingFlags.Static;
 static void Assert(bool result,string label){if(!result)throw new Exception(label);Console.WriteLine("PASS "+label);count++;}
 static void Reject(Action a,string label){try{a();}catch{Assert(true,label);return;}throw new Exception("Expected rejection: "+label);}
 static void Set(string field,object value)=>typeof(Manager).GetField(field,Flags).SetValue(null,value);
 static object Call(string name,params object[] args){try{return typeof(Manager).GetMethod(name,Flags).Invoke(null,args);}catch(TargetInvocationException e){throw e.InnerException;}}
 public static int Main(string[] args){
  if(args.Contains("--child")){Thread.Sleep(60000);return 0;}
  string folder=Path.Combine(Path.GetTempPath(),"PagePilot-deploy-tests-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(folder);
  try{
   var s=new RuntimeSettings{DataDirectory="data",ProfileDirectory="profiles",LogDirectory="logs",LogLevel="debug"};string file=Path.Combine(folder,"runtime.json");s.Save(file,false);var loaded=RuntimeSettings.Load(file);RuntimePaths.Set(loaded,file);
   Assert(Path.GetFullPath(loaded.DataDirectory).TrimEnd('\\')==Path.Combine(folder,"data"),"relative paths resolved from configuration file");
   Assert(RuntimeSettings.Load(file).LogLevel=="debug","configuration roundtrip");
   File.WriteAllText(Path.Combine(folder,"bad.json"),"{invalid");Reject(()=>RuntimeSettings.Load(Path.Combine(folder,"bad.json")),"invalid JSON rejected");
   Reject(()=>new RuntimeSettings{SchemaVersion=99}.Save(file,false),"unknown config version rejected before save");
   Reject(()=>new RuntimeSettings{LogDirectory="%PAGEPILOT_MISSING_VARIABLE%"}.Save(file,false),"unresolved environment path rejected");
   string a=Path.Combine(folder,"source"),b=Path.Combine(folder,"destination");Directory.CreateDirectory(a);Directory.CreateDirectory(b);File.WriteAllText(Path.Combine(a,"config.json"),"old");File.WriteAllText(Path.Combine(b,"config.json"),"existing");File.WriteAllText(Path.Combine(a,"rules.txt"),"rules");Manager.CopyTree(a,b);
   Assert(File.ReadAllText(Path.Combine(b,"config.json"))=="existing"&&File.ReadAllText(Path.Combine(b,"rules.txt"))=="rules","migration does not overwrite existing files");
   Reject(()=>Manager.CopyTree(a,Path.Combine(a,"nested")),"nested destination rejected");Assert(File.Exists(Path.Combine(a,"config.json")),"migration preserves originals");
   RuntimeLog.Write("manager","test.event","token=secret apiKey=another");string log=File.ReadAllText(Path.Combine(loaded.LogDirectory,"manager.jsonl"));Assert(!log.Contains("secret")&&!log.Contains("another")&&log.Contains("test.event"),"operational log redacts token fields");
   File.WriteAllText(Path.Combine(loaded.LogDirectory,"manager.jsonl"),new string('x',4*1024*1024+1));RuntimeLog.Write("manager","rotated");Assert(File.Exists(Path.Combine(loaded.LogDirectory,"manager.jsonl.1")),"log rotates at size limit");
   var receipt=new HostReceipt{Pid=10,StartUtcTicks=123,AssemblyPath=Path.Combine(folder,"PagePilot.dll"),Connected=true,RibbonLoaded=false};
   Assert(!HostHealth.Matches(receipt,10,123,receipt.AssemblyPath),"connection alone is not ribbon verification");receipt.RibbonLoaded=true;
   Assert(HostHealth.Matches(receipt,10,123,receipt.AssemblyPath),"matching connection and ribbon accepted");
   Assert(!HostHealth.Matches(receipt,10,124,receipt.AssemblyPath),"stale receipt rejected after PID reuse");
   Assert(!HostHealth.Matches(receipt,10,123,Path.Combine(folder,"other.dll")),"different deployed version rejected");
   foreach(bool machine in new[]{false,true})foreach(var view in new[]{RegistryView.Registry32,RegistryView.Registry64}){
    Set("machineScope",machine);string label=(machine?"machine-":"user-")+view;
    string prefix=@"Software\PagePilot\DeploymentRegression\"+Guid.NewGuid().ToString("N")+@"\";Set("testPrefix",prefix);Set("data",Path.Combine(folder,label));Directory.CreateDirectory(Path.Combine(folder,label));
    try{
     Set("root",Path.Combine(folder,"v1"));Call("Change","enable",view,false);Assert((bool)Call("IsRegistered",view),view+" registration validates all classes");
     string journal=Path.Combine(folder,label,machine?"machine":"","last.json"),before=File.ReadAllText(journal);Call("Change","enable",view,false);Assert(File.ReadAllText(journal)==before,"repeated enable preserves rollback journal "+view);
     Set("root",Path.Combine(folder,"v2"));Call("Change","enable",view,false);Call("Rollback");Set("root",Path.Combine(folder,"v1"));Assert((bool)Call("IsRegistered",view),"upgrade rollback "+view);
     Set("root",Path.Combine(folder,"v2"));Reject(()=>Call("Change","enable",view,true),"injected registry failure "+view);Set("root",Path.Combine(folder,"v1"));Assert((bool)Call("IsRegistered",view),"failure restores registry "+view);
     Call("Change","disable",view,false);Assert(!(bool)Call("IsRegistered",view),"disable "+view);Call("Rollback");Assert((bool)Call("IsRegistered",view),"restore after disable "+view);
     Set("root",Path.Combine(folder,"version 2 中文"));Call("Change","enable",view,false);Assert(((string)Call("CodeBase",view)).Contains("version 2 中文"),"RegAsm-compatible space and Unicode CodeBase "+view);
     using(var user=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view))using(var key=user.OpenSubKey(prefix+(machine?@"Machine\":"")+@"Software\Microsoft\Office\Word\Addins\PagePilot.Connect",true))key.SetValue("LoadBehavior",2,RegistryValueKind.DWord);
     Call("Rollback");Set("root",Path.Combine(folder,"v1"));Assert((bool)Call("IsRegistered",view),"rollback after Word clears connected bit "+view);
     // Old managers may have left a journal without Scope. Only user mode accepts it.
     string validJournal=File.ReadAllText(journal);var serializer=new JavaScriptSerializer();var parsed=serializer.Deserialize<Journal>(validJournal);
     parsed.Scope=null;File.WriteAllText(journal,serializer.Serialize(parsed));
     if(machine)Reject(()=>Call("ReadJournal"),"machine rejects legacy user journal");else Assert(Call("ReadJournal")!=null,"legacy user journal remains readable");
     File.WriteAllText(journal,validJournal);
     parsed=serializer.Deserialize<Journal>(validJournal);parsed.Entries[0].Path=@"Software\Unrelated";File.WriteAllText(journal,serializer.Serialize(parsed));
     Reject(()=>Call("Rollback"),"rollback rejects foreign registry path "+label);File.WriteAllText(journal,validJournal);
     Set("root",Path.Combine(folder,"wrong"));Reject(()=>Call("Change","disable",view,false),"old directory cannot disable another version "+label);Set("root",Path.Combine(folder,"v1"));
     Call("Change","disable",view,false);string disabled=File.ReadAllText(journal);Call("Change","disable",view,false);Assert(File.ReadAllText(journal)==disabled,"repeated disable preserves rollback "+label);Call("Rollback");
     Set("machineScope",!machine);Reject(()=>Call("Change","enable",view,false),"opposite scope shadowing rejected "+label);Set("machineScope",machine);
     Assert((bool)Call("IsRegistered",view),"opposite scope rejection preserves original "+label);
     // An interrupted transaction can restore the original state.
     Call("Change","disable",view,false);parsed=serializer.Deserialize<Journal>(File.ReadAllText(journal));parsed.State="pending";File.WriteAllText(journal,serializer.Serialize(parsed));
     Reject(()=>Call("Change","enable",view,false),"pending transaction blocks deployment "+label);Call("Rollback");Assert((bool)Call("IsRegistered",view),"pending transaction recovery "+label);
     Call("Change","disable",view,false);Set("testRollbackFailure",true);Reject(()=>Call("Rollback"),"interrupted rollback recorded "+label);Set("testRollbackFailure",false);
     Assert(serializer.Deserialize<Journal>(File.ReadAllText(journal)).State=="pending","partial rollback remains recoverable "+label);Call("Rollback");Assert((bool)Call("IsRegistered",view),"partial rollback retry succeeds "+label);
     Set("configFile",file);string export=Path.Combine(folder,label+".zip");Call("ExportDiagnostics",export);using(var z=System.IO.Compression.ZipFile.OpenRead(export)){Assert(z.GetEntry("diagnostic.txt")!=null&&z.Entries.All(e=>!e.FullName.Contains("config.json")&&!e.FullName.Contains("preferences")),"diagnostic whitelist "+view);}
    }finally{using(var k=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view))k.DeleteSubKeyTree(prefix.TrimEnd('\\'),false);}
   }
   foreach(var view in new[]{RegistryView.Registry32,RegistryView.Registry64}){
    string prefix=@"Software\PagePilot\ReplacementRegression\"+Guid.NewGuid().ToString("N")+@"\";
    string dir=Path.Combine(folder,"replacement-"+view);Set("testPrefix",prefix);Set("data",dir);Set("replaceExisting",false);Set("machineScope",false);Set("root",Path.Combine(folder,"old-user"));
    var serializer=new JavaScriptSerializer();
    try{
     Call("Change","enable",view,false);var oldUser=(System.Collections.Generic.List<Entry>)Call("UserSnapshot",view);
     Set("machineScope",true);Set("replaceExisting",true);Set("root",Path.Combine(folder,"new-machine"));
     // Seed an obsolete machine subkey; replacement must remove the whole old class tree.
     using(var user=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view))using(var k=user.CreateSubKey(prefix+@"Machine\Software\Classes\CLSID\{0E057EBB-B35B-4A31-9B65-3DF44E8034CE}\InprocServer32\0.1.0.0"))k.SetValue("CodeBase","stale");
     var oldMachine=(System.Collections.Generic.List<Entry>)Call("Snapshot",view);
     Reject(()=>Call("Change","enable",view,true),"replacement failure injected "+view);
     Assert(serializer.Serialize(Call("UserSnapshot",view))==serializer.Serialize(oldUser),"failed replacement restores user registration exactly "+view);
     Assert(serializer.Serialize(Call("Snapshot",view))==serializer.Serialize(oldMachine),"failed replacement restores machine registration exactly "+view);
     Call("Change","enable",view,false);Assert((bool)Call("IsRegistered",view),"replacement enables new machine version "+view);
     Assert(((System.Collections.Generic.List<Entry>)Call("UserSnapshot",view)).All(e=>e.Before==null),"replacement removes user shadow keys "+view);
     using(var user=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view))using(var k=user.OpenSubKey(prefix+@"Machine\Software\Classes\CLSID\{0E057EBB-B35B-4A31-9B65-3DF44E8034CE}\InprocServer32\0.1.0.0"))Assert(k==null,"replacement removes obsolete version subkeys "+view);
     string path=Path.Combine(dir,"machine","last.json"),saved=File.ReadAllText(path);Call("Change","enable",view,false);Assert(File.ReadAllText(path)==saved,"repeat replacement preserves migration rollback point "+view);
     var journal=serializer.Deserialize<Journal>(saved);Assert(journal.ReplacedUserEntries.Count==7,"migration snapshot records all user keys "+view);
     journal.UserSid="S-1-5-18";File.WriteAllText(path,serializer.Serialize(journal));Reject(()=>Call("Rollback"),"migration rollback rejects a different account "+view);File.WriteAllText(path,saved);
     using(var user=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view))using(var k=user.CreateSubKey(prefix+@"Software\Microsoft\Office\Word\Addins\PagePilot.Connect"))k.SetValue("LoadBehavior",2);
     Reject(()=>Call("Rollback"),"migration rollback detects later user override "+view);
     using(var user=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view))user.DeleteSubKeyTree(prefix+@"Software\Microsoft\Office\Word\Addins\PagePilot.Connect");
     Set("testRollbackFailure",true);Reject(()=>Call("Rollback"),"migration rollback interruption injected "+view);Set("testRollbackFailure",false);Call("Rollback");
     Assert(serializer.Serialize(Call("UserSnapshot",view))==serializer.Serialize(oldUser),"migration rollback restores user registration exactly "+view);
     Assert(serializer.Serialize(Call("Snapshot",view))==serializer.Serialize(oldMachine),"migration rollback restores machine registration exactly "+view);
     Call("Change","enable",view,false);Call("Change","disable",view,false);
     Assert(!(bool)Call("IsRegistered",view)&&((System.Collections.Generic.List<Entry>)Call("UserSnapshot",view)).All(e=>e.Before==null),"disable after migration leaves both scopes inactive "+view);
     Call("Rollback");Assert((bool)Call("IsRegistered",view),"rollback of disable restores machine version "+view);
     using(var user=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view))using(var k=user.CreateSubKey(prefix+@"Software\Classes\CLSID\{0E057EBB-B35B-4A31-9B65-3DF44E8034CE}\InprocServer32"))k.SetValue("Class","OtherVendor.Connect");
     string beforeForeign=serializer.Serialize(Call("Snapshot",view));Reject(()=>Call("Change","enable",view,false),"replacement rejects explicit foreign class "+view);
     Assert(serializer.Serialize(Call("Snapshot",view))==beforeForeign,"foreign conflict leaves target intact "+view);
     Set("machineScope",false);Reject(()=>Call("Change","enable",view,false),"replacement requires machine scope "+view);
    }finally{Set("replaceExisting",false);Set("testRollbackFailure",false);using(var user=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view))user.DeleteSubKeyTree(prefix.TrimEnd((char)92),false);}
   }
   var admins=new System.Security.Principal.SecurityIdentifier(System.Security.Principal.WellKnownSidType.BuiltinAdministratorsSid,null);
   var acl=new System.Security.AccessControl.DirectorySecurity();acl.SetOwner(admins);acl.SetAccessRuleProtection(true,false);
   acl.AddAccessRule(new System.Security.AccessControl.FileSystemAccessRule(admins,System.Security.AccessControl.FileSystemRights.FullControl,System.Security.AccessControl.AccessControlType.Allow));
   Call("CheckJournalSecurity",acl);Assert(true,"admin-owned protected journal ACL accepted");
   var everyone=new System.Security.Principal.SecurityIdentifier(System.Security.Principal.WellKnownSidType.WorldSid,null);
   acl.AddAccessRule(new System.Security.AccessControl.FileSystemAccessRule(everyone,System.Security.AccessControl.FileSystemRights.Write,System.Security.AccessControl.AccessControlType.Allow));
   Reject(()=>Call("CheckJournalSecurity",acl),"world-writable machine journal rejected");
   acl.SetOwner(everyone);Reject(()=>Call("CheckJournalSecurity",acl),"untrusted machine journal owner rejected");
   Set("testPrefix","");Set("machineScope",true);
   if(!Manager.Elevated)Reject(()=>Call("RequireScopePermission"),"non-elevated machine mutation rejected");
   Set("machineScope",false);Call("RequireScopePermission");
   string environment=(string)Call("EnvironmentReport");
   Assert(environment.Contains("已提升=")&&environment.Contains("EnableLUA=")&&environment.Contains("Registry32")&&environment.Contains("Registry64")&&environment.Contains("完整性 SID="),"environment includes process privileges, UAC and both registry views");
   Assert(!environment.Contains("ReferrerUrl=")&&!environment.Contains("HostUrl="),"download provenance URLs omitted");
   Call("Exclusive",new Action(()=>Assert(true,"global deployment lock acquired without elevation")));
   Reject(()=>Call("SelectScope","invalid"),"invalid scope rejected");
   using(var child=Process.Start(new ProcessStartInfo(Assembly.GetExecutingAssembly().Location,"--child"){UseShellExecute=false,CreateNoWindow=true})){
    try{Assert(Manager.Owned(child),"test process owner verified");Manager.CloseSelected(new[]{child},true);Assert(child.HasExited,"force close ends only explicitly supplied test child");}finally{if(!child.HasExited)child.Kill();}
   }
   Console.WriteLine("PASS "+count+" deployment checks; bitness="+(Environment.Is64BitProcess?64:32)+"; fixtures="+folder);return 0;
  }catch(Exception e){Console.WriteLine(e);return 1;}
 }
}
