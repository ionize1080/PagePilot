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
   foreach(var view in new[]{RegistryView.Registry32,RegistryView.Registry64}){
    string prefix=@"Software\PagePilot\DeploymentRegression\"+Guid.NewGuid().ToString("N")+@"\";Set("testPrefix",prefix);Set("data",Path.Combine(folder,view.ToString()));Directory.CreateDirectory(Path.Combine(folder,view.ToString()));
    try{
     Set("root",Path.Combine(folder,"v1"));Call("Change","enable",view,false);Assert((bool)Call("IsRegistered",view),view+" registration validates all classes");
     string journal=Path.Combine(folder,view.ToString(),"last.json"),before=File.ReadAllText(journal);Call("Change","enable",view,false);Assert(File.ReadAllText(journal)==before,"repeated enable preserves rollback journal "+view);
     Set("root",Path.Combine(folder,"v2"));Call("Change","enable",view,false);Call("Rollback");Set("root",Path.Combine(folder,"v1"));Assert((bool)Call("IsRegistered",view),"upgrade rollback "+view);
     Set("root",Path.Combine(folder,"v2"));Reject(()=>Call("Change","enable",view,true),"injected registry failure "+view);Set("root",Path.Combine(folder,"v1"));Assert((bool)Call("IsRegistered",view),"failure restores registry "+view);
     Call("Change","disable",view,false);Assert(!(bool)Call("IsRegistered",view),"disable "+view);Call("Rollback");Assert((bool)Call("IsRegistered",view),"restore after disable "+view);
     Set("root",Path.Combine(folder,"version 2 中文"));Call("Change","enable",view,false);Assert(((string)Call("CodeBase",view)).Contains("version 2 中文"),"RegAsm-compatible space and Unicode CodeBase "+view);
     using(var user=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view))using(var key=user.OpenSubKey(prefix+@"Software\Microsoft\Office\Word\Addins\PagePilot.Connect",true))key.SetValue("LoadBehavior",2,RegistryValueKind.DWord);
     Call("Rollback");Set("root",Path.Combine(folder,"v1"));Assert((bool)Call("IsRegistered",view),"rollback after Word clears connected bit "+view);
     Set("configFile",file);string export=Path.Combine(folder,view+".zip");Call("ExportDiagnostics",export);using(var z=System.IO.Compression.ZipFile.OpenRead(export)){Assert(z.GetEntry("diagnostic.txt")!=null&&z.Entries.All(e=>!e.FullName.Contains("config.json")&&!e.FullName.Contains("preferences")),"diagnostic whitelist "+view);}
    }finally{using(var k=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view))k.DeleteSubKeyTree(prefix.TrimEnd('\\'),false);}
   }
   Set("testPrefix","");
   using(var child=Process.Start(new ProcessStartInfo(Assembly.GetExecutingAssembly().Location,"--child"){UseShellExecute=false,CreateNoWindow=true})){
    try{Assert(Manager.Owned(child),"test process owner verified");Manager.CloseSelected(new[]{child},true);Assert(child.HasExited,"force close ends only explicitly supplied test child");}finally{if(!child.HasExited)child.Kill();}
   }
   Console.WriteLine("PASS "+count+" deployment checks; fixtures="+folder);return 0;
  }catch(Exception e){Console.WriteLine(e);return 1;}
 }
}
