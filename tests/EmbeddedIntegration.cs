using System;
using System.IO;
using System.Reflection;
using System.Collections;
using System.Collections.Generic;
using System.Web.Script.Serialization;
class EmbeddedIntegration {
 static readonly JavaScriptSerializer Json=new JavaScriptSerializer{MaxJsonLength=int.MaxValue};
 [STAThread] static int Main(string[] args){
  dynamic word=null,doc=null;try{
   string root=Path.GetFullPath(args[0]),folder=Path.Combine(root,".temp","workspace-tests"),references=Path.Combine(folder,"references");
   var assembly=Assembly.LoadFrom(Path.Combine(root,"src","bin","Release","net48","PagePilot.dll"));
   word=Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application"));word.Visible=false;word.DisplayAlerts=0;word.AutomationSecurity=3;
   doc=word.Documents.Add();doc.Content.Text="PagePilot embedded reference QA\rRevenue 120; growth 20 percent.\r";
   dynamic r=doc.Range((int)doc.Content.End-1,(int)doc.Content.End-1);
   doc.InlineShapes.AddOLEObject(ClassType:"Package",FileName:Path.Combine(references,"text.pdf"),LinkToFile:false,DisplayAsIcon:true,Range:r);
   doc.Content.InsertParagraphAfter();r=doc.Range((int)doc.Content.End-1,(int)doc.Content.End-1);
   doc.InlineShapes.AddPicture(FileName:Path.Combine(references,"chart.png"),LinkToFile:false,SaveWithDocument:true,Range:r);
   doc.SaveAs2(Path.Combine(folder,"embedded-live.docx"),16);
   string before=(string)doc.WordOpenXML;string beforeText=(string)doc.Content.Text;bool saved=(bool)doc.Saved;
   object host=Activator.CreateInstance(assembly.GetType("PagePilot.Connect"));host.GetType().GetField("App",BindingFlags.Instance|BindingFlags.NonPublic).SetValue(host,word);
   var native=assembly.GetType("PagePilot.NativeTools").GetMethod("Execute",BindingFlags.Static|BindingFlags.NonPublic);
   var a=new Dictionary<string,object>{{"document",(string)doc.FullName},{"kind","ole"},{"count",10}};
   object inspected=native.Invoke(null,new[]{host,(object)"pagepilot_objects",a});
   var parsed=Json.Deserialize<Dictionary<string,object>>(Json.Serialize(inspected));
   var objects=(ArrayList)parsed["objects"];if(objects.Count!=1)throw new Exception("Expected one real embedded PDF");
   string handle=(string)((Dictionary<string,object>)objects[0])["handle"];
   var extracted=native.Invoke(null,new[]{host,(object)"pagepilot_embedded_snapshot",new Dictionary<string,object>{{"document",(string)doc.FullName},{"handle",handle}}});
   var snapshot=Json.Deserialize<Dictionary<string,object>>(Json.Serialize(extracted));
   string target=Path.Combine(folder,"actual-ole-snapshot.xml");File.Copy((string)snapshot["packagePath"],target,true);File.Delete((string)snapshot["packagePath"]);
   File.WriteAllText(Path.Combine(folder,"ole-before.xml"),before);File.WriteAllText(Path.Combine(folder,"ole-after.xml"),(string)doc.WordOpenXML);Console.WriteLine("Save state before="+saved+" after="+(bool)doc.Saved);if(beforeText!=(string)doc.Content.Text || saved!=(bool)doc.Saved)throw new Exception("Read changed Word document");
   Console.WriteLine("PASS real Word Package PDF snapshot; original text and save state unchanged");
   doc.Content.InsertAfter("Changed after enumeration");bool rejected=false;
   try{native.Invoke(null,new[]{host,(object)"pagepilot_embedded_snapshot",new Dictionary<string,object>{{"document",(string)doc.FullName},{"handle",handle}}});}catch(TargetInvocationException){rejected=true;}
   if(!rejected)throw new Exception("Stale handle accepted");Console.WriteLine("PASS stale embedded object handle rejected");
   doc.Close(0);doc=null;
   doc=word.Documents.Open(Path.Combine(references,"report.docx"),ReadOnly:true,AddToRecentFiles:false,Visible:false);
   doc.SaveAs2(Path.Combine(references,"legacy.doc"),0);doc.Close(0);doc=null;
   Console.WriteLine("PASS generated actual Word 97-2003 DOC fixture");
   return 0;
  }catch(Exception e){Console.Error.WriteLine(e);return 1;}finally{if(doc!=null)try{doc.Close(0);}catch{}if(word!=null)try{word.Quit(0);}catch{}}
 }
}
