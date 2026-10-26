using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web.Script.Serialization;
class ObjectRegression {
 static Assembly assembly;static dynamic app,doc;static object host;static Type tools;static JavaScriptSerializer json=new JavaScriptSerializer{MaxJsonLength=int.MaxValue};
 static object Call(string method,params object[] args)=>tools.GetMethod(method,BindingFlags.Static|BindingFlags.NonPublic).Invoke(null,args);
 static Dictionary<string,object> Args(string s)=>json.Deserialize<Dictionary<string,object>>(s);
 static Dictionary<string,object> Result(object value)=>Args(json.Serialize(value));
 static void Assert(bool b,string label){if(!b)throw new Exception(label);Console.WriteLine("PASS "+label);}
 static void Run(string tool,string payload){var a=Args(payload);a["document"]=(string)doc.FullName;var r=Result(Call("Execute",host,tool,a));if(r.ContainsKey("status"))Assert((string)r["status"]=="completed",tool+" "+payload+" "+json.Serialize(r));}
 static string Handle(string kind,string scope="{}",int index=0){var a=Args("{\"kind\":\""+kind+"\",\"scope\":"+scope+"}");a["document"]=(string)doc.FullName;var r=Result(Call("Execute",host,"pagepilot_objects",a));var rows=(ArrayList)r["objects"];Assert(rows.Count>index,"enumerate "+kind);return (string)((Dictionary<string,object>)rows[index])["handle"];}
 static void Edit(string kind,string op,string extra="",string scope="{}",int index=0){for(int attempt=0;;attempt++){string h=Handle(kind,scope,index);try{Run("pagepilot_object_edit","{\"handle\":\""+h+"\",\"operation\":\""+op+"\""+extra+"}");break;}catch(Exception e){if(attempt>=2||!e.ToString().Contains("句柄失效"))throw;System.Threading.Thread.Sleep(500);}}}
 static void Create(string kind,string parameters,string scope="{\"start\":0,\"end\":0}"){if(scope=="end"){int end=(int)doc.Content.End-1;scope="{\"start\":"+end+",\"end\":"+end+"}";}Run("pagepilot_object_create","{\"kind\":\""+kind+"\",\"parameters\":"+parameters+",\"scope\":"+scope+"}");}
 [STAThread]static void Main(string[] args){
  try{assembly=Assembly.LoadFrom(Path.GetFullPath(args[0]));tools=assembly.GetType("PagePilot.NativeTools");app=Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application"));app.Visible=true;doc=app.Documents.Add();doc.SetCompatibilityMode(14);doc.Content.Text="First heading\rBody ABC ABC\rLast heading\r";doc.Content.NoProofing=1;host=Activator.CreateInstance(assembly.GetType("PagePilot.Connect"));host.GetType().GetField("App",BindingFlags.Instance|BindingFlags.NonPublic).SetValue(host,app);
   Create("table","{\"rows\":2,\"columns\":2}");Edit("cell","text",",\"parameters\":{\"text\":\"Cell one\"}");Assert(((string)doc.Tables[1].Cell(1,1).Range.Text).Contains("Cell one"),"cell text readback");Edit("cell","split_cell",",\"parameters\":{\"rows\":1,\"columns\":2}");Assert((int)doc.Tables[1].Range.Cells.Count==5,"split merged grid count");Edit("cell","set",",\"properties\":{\"VerticalAlignment\":1,\"TopPadding\":6}");Assert((int)doc.Tables[1].Range.Cells[1].VerticalAlignment==1,"cell alignment");
   int end=(int)doc.Content.End-1;Create("section_break","{\"type\":2}","{\"start\":"+end+",\"end\":"+end+"}");Assert((int)doc.Sections.Count==2,"section inserted at explicit range");Edit("section","set",",\"properties\":{\"PageSetup.DifferentFirstPageHeaderFooter\":-1,\"PageSetup.LeftMargin\":80}");Edit("header","set",",\"properties\":{\"Exists\":true}","{\"section\":1}",1);Edit("header","text",",\"parameters\":{\"text\":\"First page header\"}","{\"section\":1}",1);Assert(((string)doc.Sections[1].Headers[2].Range.Text).Contains("First page header"),"first header readback");
   end=(int)doc.Content.End-1;string point="{\"start\":"+end+",\"end\":"+end+"}";
   Create("field","{\"text\":\"DATE\"}","end");Edit("field","set",",\"properties\":{\"Locked\":true}");Assert((bool)doc.Fields[1].Locked,"field lock");Edit("field","set",",\"properties\":{\"Locked\":false}");Edit("field","update");
   Create("control","{\"type\":8,\"title\":\"Check\"}","end");Edit("control","set",",\"properties\":{\"Checked\":true}");Assert((bool)doc.ContentControls[1].Checked,"checkbox set");Edit("control","delete");Create("control","{\"type\":4}","end");Edit("control","dropdown_add",",\"parameters\":{\"text\":\"Option\",\"value\":\"1\"}");Edit("control","dropdown_select",",\"parameters\":{\"index\":1}");Assert(((string)doc.ContentControls[1].Range.Text)=="Option","dropdown selected");Edit("control","delete");
   Create("footnote","{\"text\":\"Note\"}","end");Edit("footnote","text",",\"parameters\":{\"text\":\"Updated note\"}");Assert(((string)doc.Footnotes[1].Range.Text).Contains("Updated"),"footnote edit");
   Create("equation","{\"text\":\"x^2+y^2\"}","end");Edit("equation","linearize");Edit("equation","build_up");Assert((int)doc.OMaths.Count==1,"equation roundtrip");doc.Content.InsertParagraphAfter();
   Create("textbox","{\"text\":\"Box\",\"width\":200}","end");Edit("shape","set",",\"properties\":{\"Left\":30,\"Width\":180}");Edit("shape","text",",\"parameters\":{\"text\":\"Updated textbox\"}");Assert(((string)doc.Shapes[1].TextFrame.TextRange.Text).Contains("Updated"),"textbox text");
   Create("bookmark","{\"name\":\"QABookmark\"}","end");Assert((bool)doc.Bookmarks.Exists("QABookmark"),"bookmark add");Edit("bookmark","delete");Create("style","{\"name\":\"QAStyle\"}");
   string stale=Handle("section");doc.Sections[1].PageSetup.RightMargin=91;bool rejected=false;try{Run("pagepilot_object_edit","{\"handle\":\""+stale+"\",\"properties\":{\"PageSetup.LeftMargin\":100}}");}catch{rejected=true;}Assert(rejected,"format-only stale handle rejected");
   float margin=(float)doc.Sections[1].PageSetup.LeftMargin;rejected=false;try{Edit("section","set",",\"properties\":{\"PageSetup.LeftMargin\":100,\"Execute\":1}");}catch{rejected=true;}Assert(rejected&&(float)doc.Sections[1].PageSetup.LeftMargin==margin,"all properties validated before write");
   Edit("paragraph","tab_add",",\"parameters\":{\"position\":72,\"alignment\":0}","{}",1);Edit("paragraph","tab_clear","","{}",1);

   string image=Path.GetFullPath("chatbox/icons/0-ppAI.png");Create("picture",json.Serialize(new{path=image}),"end");Edit("inline","set",",\"properties\":{\"Width\":48,\"AlternativeText\":\"QA image\"}");Assert((float)doc.InlineShapes[1].Width==48,"image resize");Edit("inline","replace_image",",\"parameters\":"+json.Serialize(new{path=image}));Assert((string)doc.InlineShapes[1].AlternativeText=="QA image","image replacement preserves alt text");
   int chartEnd=(int)doc.Content.End-1;dynamic chart=doc.InlineShapes.AddChart(51,doc.Range(chartEnd,chartEnd));Edit("chart","set",",\"properties\":{\"HasTitle\":true}");Edit("chart","set",",\"properties\":{\"ChartTitle.Text\":\"QA chart\"}");Edit("chart","series_set",",\"parameters\":{\"index\":1,\"name\":\"QA\",\"values\":[10,20],\"categories\":[\"A\",\"B\"]}");Assert((string)chart.Chart.ChartTitle.Text=="QA chart","chart title readback");
   Create("toc","{\"upper\":1,\"lower\":2}","end");Edit("toc","set",",\"properties\":{\"LowerHeadingLevel\":3}");Edit("toc","update");Assert((int)doc.TablesOfContents[1].LowerHeadingLevel==3,"TOC update settings");
   doc.Content.InsertParagraphAfter();doc.Content.InsertParagraphAfter();Create("table","{\"rows\":2,\"columns\":2}","end");int table=(int)doc.Tables.Count;string ts="{\"table\":"+table+"}";
   Edit("table","row_add",",\"parameters\":{\"before\":2}","{}",table-1);Edit("table","column_add",",\"parameters\":{\"before\":2}","{}",table-1);Assert((int)doc.Tables[table].Rows.Count==3&&(int)doc.Tables[table].Columns.Count==3,"row and column insertion");
   Edit("row","set",",\"properties\":{\"HeadingFormat\":-1,\"AllowBreakAcrossPages\":0}",ts);Edit("column","set",",\"properties\":{\"Width\":80}",ts);Edit("border","set",",\"properties\":{\"LineStyle\":1,\"LineWidth\":12,\"Color\":255}",ts);Assert((int)doc.Tables[table].Borders[1].Color==255,"table border readback");Edit("column","delete","",ts,2);Edit("row","delete","",ts,2);Edit("section","columns_set",",\"parameters\":{\"count\":2}");Assert((int)doc.Sections[1].PageSetup.TextColumns.Count==2,"section columns");
   Edit("equation","equation_text",",\"parameters\":{\"text\":\"a^2+b^2=c^2\"}");Assert(((string)doc.OMaths[1].Range.Text).Normalize(System.Text.NormalizationForm.FormKC).Contains("c"),"existing equation text replacement");
   Console.WriteLine("ALL OBJECT TESTS PASSED Word "+app.Version);
  }catch(Exception e){Console.WriteLine(e);Environment.ExitCode=1;}finally{if(doc!=null)doc.Close(0);if(app!=null&&(int)app.Documents.Count==0&&(int)app.ProtectedViewWindows.Count==0)app.Quit();}
 }
}
