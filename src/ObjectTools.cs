using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
namespace PagePilot {
 // Explicit, sparse property edits. Never exposes arbitrary COM invocation or OLE activation.
 internal static class ObjectTools {
  sealed class Ticket {internal object Doc,Value;internal string Kind,Hash;}
  static readonly Dictionary<string,Ticket> tickets=new Dictionary<string,Ticket>();
  static readonly Dictionary<string,string[]> properties=new Dictionary<string,string[]> {
   {"section",new[]{"PageSetup.Orientation","PageSetup.PageWidth","PageSetup.PageHeight","PageSetup.TopMargin","PageSetup.BottomMargin","PageSetup.LeftMargin","PageSetup.RightMargin","PageSetup.HeaderDistance","PageSetup.FooterDistance","PageSetup.Gutter","PageSetup.SectionStart","PageSetup.DifferentFirstPageHeaderFooter","PageSetup.OddAndEvenPagesHeaderFooter","PageSetup.VerticalAlignment","PageSetup.TextColumns.Spacing","PageSetup.TextColumns.EvenlySpaced"}},
   {"header",new[]{"LinkToPrevious","Exists","PageNumbers.RestartNumberingAtSection","PageNumbers.StartingNumber","PageNumbers.NumberStyle"}},
   {"footer",new[]{"LinkToPrevious","Exists","PageNumbers.RestartNumberingAtSection","PageNumbers.StartingNumber","PageNumbers.NumberStyle"}},
   {"table",new[]{"Title","Descr","AllowAutoFit","TopPadding","BottomPadding","LeftPadding","RightPadding","Spacing","Rows.Alignment","Rows.AllowBreakAcrossPages","Rows.Height","Rows.HeightRule","Rows.HeadingFormat","Borders.Enable"}},
   {"cell",new[]{"Width","Height","HeightRule","VerticalAlignment","TopPadding","BottomPadding","LeftPadding","RightPadding","FitText","WordWrap","Shading.BackgroundPatternColor"}},
   {"row",new[]{"Height","HeightRule","AllowBreakAcrossPages","HeadingFormat","Alignment","LeftIndent","SpaceBetweenColumns"}},
   {"column",new[]{"Width","PreferredWidth","PreferredWidthType"}},
   {"border",new[]{"LineStyle","LineWidth","Color","Visible"}},
   {"field",new[]{"Code.Text","Locked"}},
   {"bookmark",new string[0]},
   {"control",new[]{"Title","Tag","LockContents","LockContentControl","Checked","DateDisplayFormat","DateDisplayLocale","MultiLine","Color","Appearance"}},
   {"footnote",new string[0]}, {"endnote",new string[0]},
   {"toc",new[]{"UpperHeadingLevel","LowerHeadingLevel","UseHeadingStyles","UseOutlineLevels","UseHyperlinks","IncludePageNumbers","RightAlignPageNumbers","TabLeader"}},
   {"inline",new[]{"Width","Height","LockAspectRatio","AlternativeText","Title","ScaleWidth","ScaleHeight","PictureFormat.CropLeft","PictureFormat.CropRight","PictureFormat.CropTop","PictureFormat.CropBottom","PictureFormat.Brightness","PictureFormat.Contrast"}},
   {"shape",new[]{"Name","Width","Height","Left","Top","Rotation","LockAspectRatio","AlternativeText","Title","WrapFormat.Type","WrapFormat.DistanceTop","WrapFormat.DistanceBottom","WrapFormat.DistanceLeft","WrapFormat.DistanceRight","RelativeHorizontalPosition","RelativeVerticalPosition","LockAnchor","LayoutInCell","Line.Visible","Line.Weight","Line.ForeColor.RGB","Fill.Visible","Fill.ForeColor.RGB","TextFrame.MarginLeft","TextFrame.MarginRight","TextFrame.MarginTop","TextFrame.MarginBottom","PictureFormat.CropLeft","PictureFormat.CropRight","PictureFormat.CropTop","PictureFormat.CropBottom"}},
   {"chart",new[]{"ChartType","HasTitle","ChartTitle.Text","HasLegend","Legend.Position","ChartStyle"}},
   {"ole",new[]{"AlternativeText","Width","Height","OLEFormat.IconLabel","LinkFormat.AutoUpdate","LinkFormat.Locked"}},
   {"equation",new[]{"Type","Justification"}},
   {"paragraph",new[]{"Format.LeftIndent","Format.RightIndent","Format.FirstLineIndent","Format.SpaceBefore","Format.SpaceAfter","Format.LineSpacingRule","Format.LineSpacing","Format.KeepWithNext","Format.KeepTogether","Format.PageBreakBefore","Format.WidowControl","Format.Alignment","OutlineLevel"}},
   {"style",new[]{"NameLocal","BaseStyle","NextParagraphStyle","AutomaticallyUpdate","Font.Name","Font.NameAscii","Font.NameFarEast","Font.Size","Font.Bold","Font.Italic","Font.Color","ParagraphFormat.LeftIndent","ParagraphFormat.RightIndent","ParagraphFormat.FirstLineIndent","ParagraphFormat.SpaceBefore","ParagraphFormat.SpaceAfter","ParagraphFormat.LineSpacingRule","ParagraphFormat.LineSpacing","ParagraphFormat.Alignment"}},
   {"range",new[]{"Font.Name","Font.NameAscii","Font.NameFarEast","Font.Size","Font.Bold","Font.Italic","Font.Color","ParagraphFormat.Alignment","ParagraphFormat.LeftIndent","ParagraphFormat.RightIndent","ParagraphFormat.FirstLineIndent"}}
  };
  static string S(Dictionary<string,object>a,string k,string d="")=>a.ContainsKey(k)?Convert.ToString(a[k],CultureInfo.InvariantCulture):d;
  static int I(Dictionary<string,object>a,string k,int d=0)=>a.ContainsKey(k)?Convert.ToInt32(a[k],CultureInfo.InvariantCulture):d;
  static Dictionary<string,object> D(Dictionary<string,object>a,string k)=>a.ContainsKey(k)&&a[k]!=null?(Dictionary<string,object>)a[k]:new Dictionary<string,object>();
  internal static string Hash(dynamic doc){
   var xml=System.Xml.Linq.XDocument.Parse((string)doc.WordOpenXML);string w="http://schemas.openxmlformats.org/wordprocessingml/2006/main",w14="http://schemas.microsoft.com/office/word/2010/wordml";
   foreach(var e in xml.Descendants().Where(x=>x.Name.NamespaceName==w&&new[]{"rsids","proofErr","proofState"}.Contains(x.Name.LocalName)).ToList())e.Remove();
   foreach(var attr in xml.Descendants().Attributes().Where(x=>(x.Name.NamespaceName==w&&x.Name.LocalName.StartsWith("rsid"))||(x.Name.NamespaceName==w14&&(x.Name.LocalName=="paraId"||x.Name.LocalName=="textId"))).ToList())attr.Remove();
   using(var sha=SHA256.Create())return Convert.ToBase64String(sha.ComputeHash(Encoding.UTF8.GetBytes(xml.ToString(System.Xml.Linq.SaveOptions.DisableFormatting))));
  }
  static object Get(object obj,string path){foreach(string part in path.Split('.'))obj=obj.GetType().InvokeMember(part,BindingFlags.GetProperty,null,obj,null,CultureInfo.InvariantCulture);return obj;}
  static void Set(object obj,string path,object value){var parts=path.Split('.');for(int i=0;i<parts.Length-1;i++)obj=Get(obj,parts[i]);obj.GetType().InvokeMember(parts.Last(),BindingFlags.SetProperty,null,obj,new[]{value},CultureInfo.InvariantCulture);}
  static dynamic Range(dynamic doc,Dictionary<string,object>s){
   string area=S(s,"area","main");dynamic range;
   if(area=="main")range=doc.Content.Duplicate;
   else if(area=="header"||area=="footer"){
    int sec=I(s,"section",1),type=I(s,"type",1);CheckIndex(sec,(int)doc.Sections.Count);CheckIndex(type,3);
    dynamic h=area=="header"?doc.Sections[sec].Headers[type]:doc.Sections[sec].Footers[type];
    if(!(bool)h.Exists)throw new Exception("目标页眉/页脚不存在，请先在 section 设置首页/奇偶页选项并启用目标。");range=h.Range.Duplicate;
   }else if(area=="story"){
    int type=I(s,"type",1),chain=I(s,"chain",1);if(type<1||type>17||chain<1)throw new Exception("无效 story 范围");range=doc.StoryRanges[type].Duplicate;
    for(int i=1;i<chain;i++){range=range.NextStoryRange;if(range==null)throw new Exception("story chain 不存在");}
   }else if(area=="shape") {CheckIndex(I(s,"index"),(int)doc.Shapes.Count);range=doc.Shapes[I(s,"index")].TextFrame.TextRange.Duplicate;}
   else if(area=="page") {int n=I(s,"index",1);CheckIndex(n,(int)doc.ComputeStatistics(2));range=doc.GoTo(What:1,Which:1,Count:n);int start=(int)range.Start;int end=n<(int)doc.ComputeStatistics(2)?(int)doc.GoTo(What:1,Which:1,Count:n+1).Start:(int)doc.Content.End;range=doc.Range(start,end);}
   else throw new Exception("area 必须是 main/header/footer/story/shape/page");
   int first=(int)range.Start,last=(int)range.End,startPos=I(s,"start",first),endPos=I(s,"end",last);
   if(startPos<first||endPos<startPos||endPos>last)throw new Exception("范围超出目标文本区域");range.SetRange(startPos,endPos);return range;
  }
  static void CheckIndex(int n,int count){if(n<1||n>count)throw new Exception("索引超出范围 1.."+count);}
  internal static object Schema(int major,int compatibility=15){return new{minimumWord=14,propertySets=properties,scope=new{areas=new[]{"main","header","footer","story","shape","page"},coordinates="Word character offsets within selected story; section/type/chain/index are 1-based; header type: 1 primary, 2 first page, 3 even pages"},modern=new{commentThreads=major>=15,repeatingSectionControl=major>=15&&compatibility>=15,checkboxControl=major>=14&&compatibility>=14,controlColorAndAppearance=major>=15},operations=new[]{"set","text","delete","update","unlink","linearize","build_up","equation_text","split_cell","merge_cell","tab_add","tab_clear","dropdown_add","dropdown_delete","dropdown_select","convert_inline","convert_float","replace_image","series_set","series_delete","row_add","column_add","autofit","columns_set"},createKinds=new[]{"field","bookmark","control","footnote","endnote","toc","equation","textbox","shape","picture","table","section_break","page_break","style"},limitations=new[]{"No OLE activation, embedded application automation or arbitrary COM execution.","Object handles expire after any document change; inspect again.","Page ranges depend on current pagination; page setup belongs to sections.","Property support is probed on each object; missing interfaces are reported.","Ranges ending in table-cell markers cannot be replaced wholesale; select an inner range."}};}
  static IEnumerable<object> Items(dynamic doc,string kind,Dictionary<string,object>s){
   if(kind=="section"){foreach(dynamic x in doc.Sections)yield return x;yield break;}
   if(kind=="style"){foreach(dynamic x in doc.Styles)yield return x;yield break;}
   if(kind=="header"||kind=="footer"){CheckIndex(I(s,"section",1),(int)doc.Sections.Count);dynamic sec=doc.Sections[I(s,"section",1)];for(int t=1;t<=3;t++)yield return kind=="header"?sec.Headers[t]:sec.Footers[t];yield break;}
   dynamic r=Range(doc,s),collection=null;
   if(kind=="range"){yield return r;yield break;}
   if(kind=="row"||kind=="column"||kind=="border"){
    CheckIndex(I(s,"table",1),(int)r.Tables.Count);dynamic table=r.Tables[I(s,"table",1)];
    if(kind=="row")collection=table.Rows;else if(kind=="column")collection=table.Columns;else{dynamic target=table;if(s.ContainsKey("cell")){CheckIndex(I(s,"cell"),(int)table.Range.Cells.Count);target=table.Range.Cells[I(s,"cell")];}collection=target.Borders;}
    foreach(dynamic item in collection)yield return item;yield break;
   }
   switch(kind){case "table":collection=r.Tables;break;case "cell":CheckIndex(I(s,"table",1),(int)r.Tables.Count);collection=r.Tables[I(s,"table",1)].Range.Cells;break;case "field":collection=r.Fields;break;case "bookmark":collection=r.Bookmarks;break;case "control":collection=r.ContentControls;break;case "footnote":collection=r.Footnotes;break;case "endnote":collection=r.Endnotes;break;case "toc":foreach(dynamic toc in doc.TablesOfContents)if((bool)toc.Range.InRange(r))yield return toc;yield break;case "inline":collection=r.InlineShapes;break;case "shape":try{collection=r.ShapeRange;}catch{yield break;}break;case "equation":collection=r.OMaths;break;case "paragraph":collection=r.Paragraphs;break;case "chart":case "ole":
    foreach(dynamic x in r.InlineShapes){if(kind=="chart"){bool has=false;try{has=(int)x.HasChart!=0;}catch{}if(has)yield return x.Chart;}else if((int)x.Type==1||(int)x.Type==2)yield return x;}
    object shapes=null;try{shapes=r.ShapeRange;}catch{}if(shapes!=null)foreach(dynamic x in (IEnumerable)shapes){if(kind=="chart"){bool has=false;try{has=(int)x.HasChart!=0;}catch{}if(has)yield return x.Chart;}else if((int)x.Type==7||(int)x.Type==10)yield return x;}yield break;
    default:throw new Exception("不支持的对象类型");}
   foreach(dynamic x in collection)yield return x;
  }
  internal static object Inspect(dynamic doc,Dictionary<string,object>a){
   string kind=S(a,"kind","range");if(kind=="schema")return Schema((int)double.Parse((string)doc.Application.Version,CultureInfo.InvariantCulture),(int)doc.CompatibilityMode);
   if(kind=="stories"){var stories=new List<object>();foreach(dynamic root in doc.StoryRanges){dynamic r=root;int chain=0;while(r!=null){stories.Add(new{area="story",type=(int)r.StoryType,chain=++chain,start=(int)r.Start,end=(int)r.End});r=r.NextStoryRange;}}return new{stories};}
   if(!properties.ContainsKey(kind))throw new Exception("未知对象类型");int start=I(a,"start",1),count=I(a,"count",50);if(start<1||count<1||count>200)throw new Exception("start>=1，count=1..200");
   if(tickets.Count+count>256)tickets.Clear();string hash=Hash(doc);var issued=new List<Ticket>();var rows=new List<object>();int index=0;bool more=false;
   foreach(dynamic obj in Items(doc,kind,D(a,"scope"))){index++;if(index<start)continue;if(rows.Count>=count){more=true;break;}
    var values=new Dictionary<string,object>();var unavailable=new List<string>();
    bool absent=(kind=="header"||kind=="footer")&&!(bool)obj.Exists;
    foreach(string prop in properties[kind]){if(absent&&prop!="Exists"){unavailable.Add(prop);continue;}try{object v=Get(obj,prop);if(v==null||v is string||v is bool||v is IConvertible)values[prop]=v;}catch{unavailable.Add(prop);}}
    foreach(string prop in new[]{"Name","ID","Type","Range.Start","Range.End","Range.StoryType","Range.Text","Code.Text","Result.Text","TextFrame.TextRange.Text","OLEFormat.ProgID","LinkFormat.SourceFullName","RowIndex","ColumnIndex","NestingLevel"})try{if(absent)continue;object v=Get(obj,prop);if(v is string str&&str.Length>4000)v=str.Substring(0,4000);if(v is string||v is IConvertible)values[prop]=v;}catch{}
    if(kind=="range"){string content=(string)obj.Text??"";values["text"]=content.Length>16000?content.Substring(0,16000):content;values["textTruncated"]=content.Length>16000;values["start"]=(int)obj.Start;values["end"]=(int)obj.End;}
    if(kind=="chart"){var series=new List<object>();try{int n=(int)obj.SeriesCollection().Count;for(int i=1;i<=n;i++){dynamic ss=obj.SeriesCollection(i);series.Add(new{index=i,name=(string)ss.Name,values=(object)ss.Values,categories=(object)ss.XValues});}values["series"]=series;}catch{unavailable.Add("series");}}
    string id=Guid.NewGuid().ToString("N");var issuedTicket=new Ticket{Doc=doc,Value=obj,Kind=kind,Hash=hash};tickets[id]=issuedTicket;issued.Add(issuedTicket);
    rows.Add(new{index,handle=id,properties=values,unavailableProperties=unavailable});
   }
   hash=Hash(doc);foreach(var issuedTicket in issued)issuedTicket.Hash=hash;
   return new{kind,objects=rows,nextStart=more?(int?)(start+rows.Count):null,requiresFreshInspectionAfterEdit=true};
  }
  static object ConvertValue(object original,object value){if(value==null)throw new Exception("属性不能为 null");if(original is string){if(!(value is string)||((string)value).Length>100000)throw new Exception("无效文本属性");return value;}if(original is bool){if(!(value is bool))throw new Exception("需要布尔值");return value;}double n=Convert.ToDouble(value,CultureInfo.InvariantCulture);if(double.IsNaN(n)||double.IsInfinity(n))throw new Exception("需要有限数值");return Convert.ChangeType(value,original.GetType(),CultureInfo.InvariantCulture);}
  static object Mutate(dynamic doc,string operation,Action action){
   string backup=NativeTools.Backup(doc),before=Hash(doc);bool recording=false;
   try{doc.Application.UndoRecord.StartCustomRecord("PagePilot "+operation);recording=true;action();doc.Application.UndoRecord.EndCustomRecord();recording=false;tickets.Clear();return new{status="completed",operation,backup};}
   catch(Exception ex){if(recording)try{doc.Application.UndoRecord.EndCustomRecord();}catch{}bool rollback=false;try{if(Hash(doc)==before)rollback=true;else{doc.Undo(1);rollback=Hash(doc)==before;}}catch{}tickets.Clear();return new{status=rollback?"not_applied":"recovery_required",operation,error=ex.GetBaseException().Message,backup,rollbackVerified=rollback};}
  }
  internal static object Edit(dynamic doc,Dictionary<string,object>a){
   if((bool)doc.ReadOnly||(int)doc.ProtectionType!=-1)throw new Exception("文档只读或受保护，无法编辑。");
   string handle=S(a,"handle"),op=S(a,"operation","set");if(!tickets.TryGetValue(handle,out var ticket))throw new Exception("对象句柄失效：已过期，请重新枚举。");if(!object.Equals(ticket.Doc,(object)doc))throw new Exception("对象句柄失效：目标文档已切换。");if(ticket.Hash!=Hash(doc))throw new Exception("对象句柄失效：文档内容或格式已变化，请重新枚举。");
   dynamic obj=ticket.Value;string kind=ticket.Kind;var p=D(a,"parameters");Action action;
   if(op=="set"){
    var patch=D(a,"properties");if(patch.Count==0)throw new Exception("缺少属性");var validated=new Dictionary<string,object>();int major=(int)double.Parse((string)doc.Application.Version,CultureInfo.InvariantCulture);
    foreach(var entry in patch){if(!properties[kind].Contains(entry.Key))throw new Exception("不允许修改属性 "+entry.Key);if(kind=="control"&&major<15&&(entry.Key=="Color"||entry.Key=="Appearance"))throw new Exception("此属性需要 Word 2013+");if(kind=="style"&&(entry.Key=="BaseStyle"||entry.Key=="NextParagraphStyle")){string name=entry.Value as string;if(string.IsNullOrWhiteSpace(name))throw new Exception("需要已有样式名称");object existing=doc.Styles[name];validated[entry.Key]=name;}else validated[entry.Key]=ConvertValue(Get(obj,entry.Key),entry.Value);}
    action=()=>{foreach(var entry in validated)Set(obj,entry.Key,entry.Value);};
   }else if(op=="text"){
    string text=S(p,"text");dynamic r=kind=="range"?obj:kind=="shape"?obj.TextFrame.TextRange:obj.Range;
    if(!new[]{"range","shape","cell","control","footnote","endnote","bookmark","header","footer"}.Contains(kind))throw new Exception("此对象不支持文本替换");
    r=r.Duplicate;if(kind=="cell")r.End=(int)r.End-1;
    if(((string)r.Text??"").Contains('\a')||text.Contains('\a'))throw new Exception("不能整体覆盖含单元格边界的范围");action=()=>r.Text=text;
   }else if(op=="delete"){
    if(!new[]{"table","row","column","field","bookmark","control","footnote","endnote","toc","inline","shape","equation","style"}.Contains(kind))throw new Exception("此对象不支持删除");
    action=()=>{if(kind=="control")obj.Delete(false);else if(kind=="equation")obj.Range.Delete();else obj.Delete();};
   }else if(op=="update"&&(kind=="field"||kind=="toc"))action=()=>{if(kind=="field"&&!(bool)obj.Update())throw new Exception("域更新失败");if(kind=="toc")obj.Update();};
   else if(op=="unlink"&&kind=="field")action=()=>obj.Unlink();
   else if(op=="build_up"&&kind=="equation")action=()=>obj.BuildUp();
   else if(op=="linearize"&&kind=="equation")action=()=>obj.Linearize();
   else if(op=="equation_text"&&kind=="equation"){string text=S(p,"text");if(text.Length==0)throw new Exception("公式不能为空");dynamic r=obj.Range.Duplicate;int first=(int)r.Start;action=()=>{obj.Remove();r.Text=text;r.SetRange(first,first+text.Length);doc.OMaths.Add(r).OMaths[1].BuildUp();};}
   else if(op=="split_cell"&&kind=="cell"){int rows=I(p,"rows",1),cols=I(p,"columns",1);if(rows<1||cols<1||rows*cols>1000)throw new Exception("拆分尺寸无效");action=()=>obj.Split(rows,cols);}
   else if(op=="merge_cell"&&kind=="cell"){if(!tickets.TryGetValue(S(p,"other_handle"),out var other)||other.Kind!="cell"||other.Hash!=ticket.Hash||!object.Equals(other.Doc,ticket.Doc))throw new Exception("第二单元格句柄失效");dynamic b=other.Value;if(!object.Equals(obj.Range.Tables[1],b.Range.Tables[1]))throw new Exception("只能合并同一表格的单元格");action=()=>obj.Merge(b);}
   else if((op=="tab_add"||op=="tab_clear")&&(kind=="paragraph"||kind=="range")){dynamic tabs=kind=="paragraph"?obj.Format.TabStops:obj.ParagraphFormat.TabStops;float pos=Convert.ToSingle(p.ContainsKey("position")?p["position"]:0);if(float.IsNaN(pos)||float.IsInfinity(pos)||pos<0||pos>1584)throw new Exception("制表位位置无效");int align=I(p,"alignment"),leader=I(p,"leader");if(align<0||align>4||leader<0||leader>5)throw new Exception("制表位属性无效");action=()=>{if(op=="tab_clear")tabs.ClearAll();else tabs.Add(pos,align,leader);};}
   else if(op.StartsWith("dropdown_")&&kind=="control"){
    if((int)obj.Type!=3&&(int)obj.Type!=4)throw new Exception("需要组合框或下拉列表控件");int i=I(p,"index",1);if(op!="dropdown_add")CheckIndex(i,(int)obj.DropdownListEntries.Count);
    if(op=="dropdown_add"){string text=S(p,"text"),value=S(p,"value",text);if(text.Length==0)throw new Exception("列表文字不能为空");action=()=>obj.DropdownListEntries.Add(text,value);}
    else if(op=="dropdown_delete")action=()=>obj.DropdownListEntries[i].Delete();else if(op=="dropdown_select")action=()=>obj.DropdownListEntries[i].Select();else throw new Exception("未知列表操作");
   }else if(op=="convert_inline"&&kind=="shape")action=()=>obj.ConvertToInlineShape();
   else if(op=="convert_float"&&kind=="inline")action=()=>obj.ConvertToShape();
   else if(op=="replace_image"&&kind=="inline"){
    if((int)obj.Type!=3&&(int)obj.Type!=4)throw new Exception("需要图片对象");string path=System.IO.Path.GetFullPath(S(p,"path"));if(!System.IO.File.Exists(path))throw new Exception("图片不存在");float width=(float)obj.Width,height=(float)obj.Height;string alt=(string)obj.AlternativeText;dynamic r=obj.Range.Duplicate;action=()=>{obj.Delete();r.Collapse(1);dynamic image=r.InlineShapes.AddPicture(path,false,true,r);image.LockAspectRatio=0;image.Width=width;image.Height=height;image.AlternativeText=alt;};

   }else if(op=="replace_image"&&kind=="shape"){
    if((int)obj.Type!=11&&(int)obj.Type!=13)throw new Exception("需要浮动图片对象");string path=System.IO.Path.GetFullPath(S(p,"path"));if(!System.IO.File.Exists(path))throw new Exception("图片不存在");dynamic anchor=obj.Anchor.Duplicate;var preserve=new Dictionary<string,object>();foreach(string key in new[]{"Left","Top","Width","Height","Rotation","AlternativeText","Title","WrapFormat.Type","RelativeHorizontalPosition","RelativeVerticalPosition","LockAnchor","LayoutInCell"})preserve[key]=Get(obj,key);action=()=>{obj.Delete();dynamic image=doc.Shapes.AddPicture(path,false,true,0,0,Convert.ToSingle(preserve["Width"]),Convert.ToSingle(preserve["Height"]),anchor);foreach(var value in preserve)Set(image,value.Key,value.Value);};
   }else if(op=="series_set"&&kind=="chart"){
    int i=I(p,"index",1);CheckIndex(i,(int)obj.SeriesCollection().Count+1);var values=p.ContainsKey("values")?p["values"] as ArrayList:null;var cats=p.ContainsKey("categories")?p["categories"] as ArrayList:null;
    if(values==null||cats==null||values.Count==0||values.Count!=cats.Count||values.Count>10000)throw new Exception("需要等长的 values/categories");double[] data=values.Cast<object>().Select(x=>Convert.ToDouble(x,CultureInfo.InvariantCulture)).ToArray();if(data.Any(x=>double.IsNaN(x)||double.IsInfinity(x)))throw new Exception("图表数据必须是有限数值");object[] labels=cats.Cast<object>().Select(x=>(object)Convert.ToString(x)).ToArray();string name=S(p,"name","Series");action=()=>{dynamic series=i>(int)obj.SeriesCollection().Count?obj.SeriesCollection().NewSeries():obj.SeriesCollection(i);series.Name=name;series.Values=data;series.XValues=labels;};
   }else if(op=="series_delete"&&kind=="chart"){int i=I(p,"index");CheckIndex(i,(int)obj.SeriesCollection().Count);action=()=>obj.SeriesCollection(i).Delete();}
   else if((op=="row_add"||op=="column_add")&&kind=="table"){dynamic list=op=="row_add"?obj.Rows:obj.Columns;int before=I(p,"before",(int)list.Count+1);CheckIndex(before,(int)list.Count+1);action=()=>{if(before>(int)list.Count)list.Add();else list.Add(list[before]);};}
   else if(op=="autofit"&&kind=="table"){int behavior=I(p,"behavior",0);if(behavior<0||behavior>2)throw new Exception("behavior: 0 fixed, 1 content, 2 window");action=()=>obj.AutoFitBehavior(behavior);}
   else if(op=="columns_set"&&kind=="section"){int count=I(p,"count",1);if(count<1||count>12)throw new Exception("栏数必须为 1..12");action=()=>obj.PageSetup.TextColumns.SetCount(count);}
   else throw new Exception("此类型不支持操作 "+op);
   return Mutate(doc,op,action);
  }
  internal static object Create(dynamic doc,Dictionary<string,object>a){
   if((bool)doc.ReadOnly||(int)doc.ProtectionType!=-1)throw new Exception("文档只读或受保护，无法编辑。");
   string kind=S(a,"kind");var p=D(a,"parameters");dynamic range=Range(doc,D(a,"scope"));Action action;string text=S(p,"text");
   switch(kind){
    case "field":if(string.IsNullOrWhiteSpace(text))throw new Exception("缺少域代码");action=()=>range.Fields.Add(range,-1,text,true);break;
    case "bookmark":string name=S(p,"name");if(!System.Text.RegularExpressions.Regex.IsMatch(name,@"^[A-Za-z_][A-Za-z0-9_]{0,39}$")||(bool)doc.Bookmarks.Exists(name))throw new Exception("书签名无效或已存在");action=()=>doc.Bookmarks.Add(name,range);break;
    case "control":int type=I(p,"type",0),major=(int)double.Parse((string)doc.Application.Version,CultureInfo.InvariantCulture);if(type<0||type>9||(type==8&&(int)doc.CompatibilityMode<14)||(type==9&&(major<15||(int)doc.CompatibilityMode<15)))throw new Exception("控件类型不受此 Word 版本支持");action=()=>{dynamic c=doc.ContentControls.Add(type,range);c.Title=S(p,"title","PagePilot");c.Tag=S(p,"tag");if(text.Length>0&&(type==0||type==1))c.Range.Text=text;};break;
    case "footnote":action=()=>doc.Footnotes.Add(range,Text:text);break;
    case "endnote":action=()=>doc.Endnotes.Add(range,Text:text);break;
    case "toc":int upper=I(p,"upper",1),lower=I(p,"lower",9);if(upper<1||lower<upper||lower>9)throw new Exception("目录级别必须在 1..9");action=()=>doc.TablesOfContents.Add(range,UseHeadingStyles:true,UpperHeadingLevel:upper,LowerHeadingLevel:lower,UseOutlineLevels:true);break;
    case "equation":if(text.Length==0)throw new Exception("公式不能为空");action=()=>{int first=(int)range.Start;range.Text=text;range.SetRange(first,first+text.Length);dynamic math=doc.OMaths.Add(range);math.OMaths[1].BuildUp();};break;
    case "textbox":case "shape":int shape=I(p,"type",1);if(shape<1||shape>183)throw new Exception("形状类型超出范围");float width=Number(p,"width",144,1,1584),height=Number(p,"height",72,1,1584),left=Number(p,"left",0,-1584,1584),top=Number(p,"top",0,-1584,1584);action=()=>{dynamic o=kind=="textbox"?doc.Shapes.AddTextbox(1,left,top,width,height,range):doc.Shapes.AddShape(shape,left,top,width,height,range);if(text.Length>0)o.TextFrame.TextRange.Text=text;};break;
    case "picture":string path=System.IO.Path.GetFullPath(S(p,"path"));if(!System.IO.File.Exists(path))throw new Exception("图片不存在");action=()=>range.InlineShapes.AddPicture(path,false,true,range);break;
    case "table":int rows=I(p,"rows",2),cols=I(p,"columns",2);if(rows<1||cols<1||cols>63||rows*cols>10000)throw new Exception("表格尺寸无效");action=()=>doc.Tables.Add(range,rows,cols);break;
    case "section_break":int breakType=I(p,"type",2);if(!new[]{2,3,4,5}.Contains(breakType))throw new Exception("分节符 type: 2 next page, 3 continuous, 4 even, 5 odd");if((int)range.Start!=(int)range.End)throw new Exception("分节符需要折叠范围");action=()=>range.InsertBreak(breakType);break;
    case "page_break":if((int)range.Start!=(int)range.End)throw new Exception("分页符需要折叠范围");action=()=>range.InsertBreak(7);break;
    case "style":string styleName=S(p,"name");int styleType=I(p,"type",1);if(styleName.Length<1||styleName.Length>253||styleType<1||styleType>4)throw new Exception("样式参数无效");foreach(dynamic st in doc.Styles)if((string)st.NameLocal==styleName)throw new Exception("样式已存在");action=()=>doc.Styles.Add(styleName,styleType);break;
    default:throw new Exception("不支持创建类型 "+kind);
   }
   return Mutate(doc,"create_"+kind,action);
  }
  static float Number(Dictionary<string,object>p,string key,float fallback,float min,float max){float v=p.ContainsKey(key)?Convert.ToSingle(p[key]):fallback;if(float.IsNaN(v)||float.IsInfinity(v)||v<min||v>max)throw new Exception("无效参数 "+key);return v;}
 }
}
