using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using PagePilot;
using F=System.Windows.Forms;
class ApprovalRegression{
 static BindingFlags flags=BindingFlags.Instance|BindingFlags.NonPublic|BindingFlags.Public;
 static bool Automatic(object o)=>(bool)o.GetType().GetProperty("automatic").GetValue(o);
 static IEnumerable<F.Control> Children(F.Control c){foreach(F.Control child in c.Controls){yield return child;foreach(var next in Children(child))yield return next;}}
 [STAThread]static void Main(){var runner=new F.Form();runner.Shown+=(s,e)=>{dynamic app=null,doc=null;try{
 app=Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application"));doc=app.Documents.Add();app.Visible=true;
 var host=new Connect();typeof(Connect).GetField("App",flags).SetValue(host,app);var handle=typeof(Connect).GetMethod("Handle",flags);
 var request=new Dictionary<string,object>{{"action","ai_approval"},{"enabled",true}};
 bool accept=false;int reentered=0;
 var timer=new F.Timer{Interval=100};timer.Tick+=(ts,te)=>{var modal=F.Application.OpenForms.Cast<F.Form>().FirstOrDefault(f=>f.Text.Contains("确认自动批准"));if(modal==null)return;
 if(!accept){var nested=handle.Invoke(host,new object[]{request});if(Automatic(nested))throw new Exception("Reentry enabled approval");reentered++;modal.Close();}
 else ((F.Button)Children(modal).First(c=>c is F.Button&&c.Text=="启用自动批准")).PerformClick();};
 timer.Start();var result=handle.Invoke(host,new object[]{request});if(Automatic(result)||reentered!=1)throw new Exception("Cancel or reentry failed");Console.WriteLine("PASS cancel and confirmation reentry");
 accept=true;result=handle.Invoke(host,new object[]{request});if(!Automatic(result))throw new Exception("Enable failed");Console.WriteLine("PASS explicit enable");timer.Stop();timer.Dispose();
 dynamic other=app.Documents.Add();result=typeof(Connect).GetMethod("ApprovalState",flags).Invoke(host,null);if(Automatic(result))throw new Exception("Document isolation failed");other.Close(0);result=typeof(Connect).GetMethod("ApprovalState",flags).Invoke(host,null);if(Automatic(result))throw new Exception("Approval carried back");Console.WriteLine("PASS document identity reset");
 }catch(Exception ex){Console.WriteLine(ex);Environment.ExitCode=1;}finally{if(doc!=null)doc.Close(0);if(app!=null)app.Quit();runner.Close();}};F.Application.Run(runner);}
}
