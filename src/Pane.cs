using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;
using System.Windows.Forms.Integration;
namespace PagePilot {
 [ComVisible(true),Guid("79CEEFC8-67F4-4F55-9475-24C0681BFBBD"),ProgId("PagePilot.Pane"),ClassInterface(ClassInterfaceType.AutoDispatch)]
 public sealed class Pane:System.Windows.Forms.UserControl {
  private Connect host;private dynamic doc;private int page,appliedTheme=-1;private long windowHandle;
  private FrameworkElement view;private TextBox pattern,replacement;private CheckBox ignore,selection;
  private ListBox hits;private TreeView headings;private bool rebuilding;private TextBlock status;
  private List<Hit> found=new List<Hit>();private string snapshot;private int snapshotStart,snapshotEnd,snapshotDocumentEnd;private bool snapshotSelection;
  private int? lockedStart,lockedEnd;private string cachedText,cachedQuery;private int[] cachedPositions,cachedEnds;private bool cachedIgnore,aiOutline;private int outlineVersion;
  private bool scanning;private int appearanceVersion=-1;private System.Threading.CancellationTokenSource scanCancel;
  private int idleTicks;private int outlineEnd;private List<Heading> outline=new List<Heading>();private string outlineSnapshot,ruleSnapshot;
  private readonly System.Windows.Forms.Timer searchTimer=new System.Windows.Forms.Timer{Interval=350};private int searchVersion;private System.Threading.CancellationTokenSource searchCancel;private SearchHighlight highlight;private readonly System.Windows.Forms.Timer highlightTimer=new System.Windows.Forms.Timer{Interval=180};
  private readonly System.Windows.Forms.Timer timer=new System.Windows.Forms.Timer{Interval=1800};
  public Pane(){AutoScaleMode=System.Windows.Forms.AutoScaleMode.None;}
  private T Find<T>(string name) where T:class=>view.FindName(name) as T;
  private void On(string name,Action action){Find<Button>(name).Click+=(s,e)=>{Connect.Log("pane.command page="+page+" action="+name);Guard(action);};}
  internal void Bind(Connect h,object document,int mode){
   host=h;doc=document;page=mode;windowHandle=WordWindows.Handle(h.App.ActiveWindow).ToInt64();
   using(var stream=typeof(Pane).Assembly.GetManifestResourceStream("PagePilot.Views."+(mode==0?"Regex":"Navigation")+".xaml"))view=(FrameworkElement)XamlReader.Load(stream);
   using(var stream=typeof(Pane).Assembly.GetManifestResourceStream("PagePilot.Views.Modern.xaml"))view.Resources.MergedDictionaries.Add((ResourceDictionary)XamlReader.Load(stream));
   Controls.Add(new ElementHost{Dock=System.Windows.Forms.DockStyle.Fill,Child=view});status=Find<TextBlock>("Status");
   On("Settings",()=>host.OpenSettings(null));
   ApplyTheme();ApplyNavigationAppearance();
   if(page==0){
    pattern=Find<TextBox>("Pattern");replacement=Find<TextBox>("Replacement");ignore=Find<CheckBox>("IgnoreCase");selection=Find<CheckBox>("SelectionOnly");hits=Find<ListBox>("Matches");
    On("Preview",()=>{lockedStart=null;lockedEnd=null;FindMatches();});On("Replace",Replace);
    highlight=new SearchHighlight{MarkColor=System.Drawing.ColorTranslator.FromHtml(host.NavigationStyle.HighlightColor)};highlightTimer.Tick+=(s,e)=>{try{if(Visible&&WordWindows.IsWindowVisible(Handle)&&snapshot!=null&&host.App.Documents.Count>0&&object.Equals(host.App.ActiveDocument,doc)&&WordWindows.Handle(host.App.ActiveWindow).ToInt64()==windowHandle)highlight.Update(doc,host.App.ActiveWindow,new IntPtr(windowHandle),found);else highlight.Hide();}catch{highlight.Hide();}};highlightTimer.Start();
    searchTimer.Tick+=(s,e)=>{searchTimer.Stop();FindMatches();};
    pattern.TextChanged+=(s,e)=>ScheduleSearch();replacement.TextChanged+=(s,e)=>UpdateReplacement();
    ignore.Click+=(s,e)=>ScheduleSearch();selection.Click+=(s,e)=>{lockedStart=null;lockedEnd=null;ScheduleSearch();};
    hits.SelectionChanged+=(s,e)=>Guard(()=>{var item=hits.SelectedItem as Hit;if(item!=null){EnsureSnapshot();doc.Activate();doc.Range(item.Start,item.Start+item.Length).Select();}});
   }else{
    headings=Find<TreeView>("Headings");On("Refresh",RefreshOutline);On("StopScan",()=>{scanCancel?.Cancel();outlineVersion++;status.Text="分析已取消";});On("Analyze",AIOutline);On("Apply",ApplyOutline);
    headings.SelectedItemChanged+=(s,e)=>{if(!rebuilding)Guard(()=>{var item=(headings.SelectedItem as TreeViewItem)?.Tag as Heading;if(item!=null)Navigate(item);});};

   }
   timer.Tick+=(s,e)=>{if(Visible&&WordWindows.IsWindowVisible(Handle)){try{ApplyTheme();ApplyNavigationAppearance();RefreshBinding();bool inspect=++idleTicks%3==0;
     if(page==1&&!scanning&&navigationCancel==null&&outlineSnapshot!=null&&((int)doc.Content.End!=outlineEnd||Connect.ReadRulesText()!=ruleSnapshot)){status.Text="目录待更新 · 点击分析刷新（含样式变化）";}
     if(page==0&&snapshot!=null&&((int)doc.Content.End!=snapshotDocumentEnd||(inspect&&(string)doc.Content.Text!=snapshot)))ScheduleSearch();}catch{}}};timer.Start();
  }
  private void ApplyTheme(){if(appliedTheme==host.Theme)return;appliedTheme=host.Theme;
   view.Resources["Accent"]=new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(host.Theme==1?"#4671D6":host.Theme==2?"#515E6F":"#238B73"));
   using(var bitmap=RibbonIcons.Draw("ppSettings",24,host.Theme))using(var memory=new MemoryStream()){bitmap.Save(memory,System.Drawing.Imaging.ImageFormat.Png);memory.Position=0;var source=new System.Windows.Media.Imaging.BitmapImage();source.BeginInit();source.CacheOption=System.Windows.Media.Imaging.BitmapCacheOption.OnLoad;source.StreamSource=memory;source.EndInit();source.Freeze();Find<Button>("Settings").Content=new Image{Source=source,Width=18,Height=18};}
  }
  private void RefreshBinding(){
   if(host.App.Documents.Count==0||WordWindows.Handle(host.App.ActiveWindow).ToInt64()!=windowHandle)return;
   dynamic current=host.App.ActiveDocument;bool same=false;
   try{same=object.Equals(doc,current)&&(string)doc.FullName==(string)current.FullName;}catch{}
   if(same)return;
   navigationCancel?.Cancel();scanCancel?.Cancel();outlineVersion++;aiOutline=false;lockedStart=null;lockedEnd=null;searchVersion++;doc=current;snapshot=null;outlineSnapshot=null;found.Clear();outline.Clear();
   if(hits!=null)hits.ItemsSource=null;if(headings!=null)headings.Items.Clear();
   status.Text="已切换文档 · 请重新预览";
  }
  public void SelectPage(int index){Guard(()=>{if(page==1&&outlineSnapshot==null&&!scanning)BeginInvoke(new Action(RefreshOutline));});}
  private void Guard(Action action){try{RefreshBinding();action();}catch(Exception e){PagePilot.Configuration.RuntimeLog.Write("host","pane.failed",e.GetType().Name+" stack="+e.StackTrace,"error");status.Text=e.Message;System.Windows.Forms.MessageBox.Show(e.Message,"PagePilot");}}
  private void ScheduleSearch(){searchCancel?.Cancel();highlight?.Hide();searchVersion++;snapshot=null;searchTimer.Stop();searchTimer.Start();}
  private async void FindMatches(){
   searchCancel?.Cancel();var cancel=new System.Threading.CancellationTokenSource();searchCancel=cancel;int version=++searchVersion;var watch=System.Diagnostics.Stopwatch.StartNew();PagePilot.Configuration.RuntimeLog.Write("host","regex.search.begin","revision="+version,"debug");
   try{
    if(string.IsNullOrEmpty(pattern.Text)){found.Clear();hits.ItemsSource=null;status.Text="输入表达式开始查找";return;}
    dynamic range=doc.Content;
    if(selection.IsChecked==true){if(!lockedStart.HasValue){dynamic chosen=host.App.Selection.Range;if((int)chosen.StoryType!=1)throw new Exception("正则侧栏仅支持正文；其他区域请用对象工具。");lockedStart=(int)chosen.Start;lockedEnd=(int)chosen.End;}if(lockedEnd>(int)doc.Content.End)throw new Exception("原选区已失效，请点击查找预览重新选择。");range=doc.Range(lockedStart.Value,lockedEnd.Value);}
    if(selection.IsChecked==true&&!object.Equals(host.App.ActiveDocument,doc))throw new Exception("请激活此窗格对应的文档后再查找。");
    string full=(string)doc.Content.Text,text=(string)range.Text,query=pattern.Text,replace=replacement.Text;
    int start=(int)range.Start,end=(int)range.End;bool selected=selection.IsChecked==true,ic=ignore.IsChecked==true;
    // Simple stories have a direct UTF-16 mapping. Complex stories require explicit Word positions.
    int[] positions=null,ends=null;
    if((int)range.Fields.Count>0||(int)range.ContentControls.Count>0||end-start!=text.Length){
     var map=new List<int>();var finish=new List<int>();var actual=new System.Text.StringBuilder();
     var fieldStarts=new Dictionary<int,int>();var fieldEnds=new Dictionary<int,int>();
     foreach(dynamic field in doc.Fields){fieldStarts[(int)field.Code.Start-1]=(int)field.Result.Start;fieldEnds[(int)field.Result.End+1]=(int)field.Result.End;}
     foreach(dynamic paragraph in range.Paragraphs){
      dynamic pr=paragraph.Range;int ps=Math.Max(start,(int)pr.Start),pe=Math.Min(end,(int)pr.End);if(pe<=ps)continue;
      dynamic part=doc.Range(ps,pe);string pt=(string)part.Text??"";
      if((int)part.Fields.Count==0&&(int)part.ContentControls.Count==0&&pe-ps==pt.Length){for(int k=0;k<pt.Length;k++){map.Add(ps+k);finish.Add(ps+k+1);}actual.Append(pt);}
      else foreach(dynamic c in part.Characters){string t=(string)c.Text??"";int cs=(int)c.Start,ce=(int)c.End;while(fieldStarts.ContainsKey(cs))cs=fieldStarts[cs];while(fieldEnds.ContainsKey(ce))ce=fieldEnds[ce];
       for(int k=0;k<t.Length;k++){map.Add(cs+Math.Min(k,ce-cs-1));finish.Add(Math.Min(cs+k+1,ce));actual.Append(t[k]);}}
     }
     map.Add(end);
     if(actual.ToString()!=text){
      map.Clear();finish.Clear();actual.Clear();
      foreach(dynamic c in range.Characters){string t=(string)c.Text??"";int cs=(int)c.Start,ce=(int)c.End;while(fieldStarts.ContainsKey(cs))cs=fieldStarts[cs];while(fieldEnds.ContainsKey(ce))ce=fieldEnds[ce];for(int k=0;k<t.Length;k++){map.Add(cs+Math.Min(k,ce-cs-1));finish.Add(Math.Min(cs+k+1,ce));actual.Append(t[k]);}}
      map.Add(end);if(actual.ToString()!=text)throw new Exception("此范围包含无法安全映射的域，请缩小查找范围。");
     }
     positions=map.ToArray();ends=finish.ToArray();
    }
    status.Text="正在查找…";
    var result=await System.Threading.Tasks.Task.Run(()=>Engine.Find(text,0,query,replace,ic,cancel.Token));
    if(version!=searchVersion||IsDisposed)return;
    foreach(var hit in result){int a=hit.Start,b=a+hit.Length;hit.Start=positions==null?start+a:positions[a];hit.Length=(positions==null?start+b:(a==b?positions[a]:ends[b-1]))-hit.Start;}
    cachedText=text;cachedQuery=query;cachedIgnore=ic;cachedPositions=positions;cachedEnds=ends;
    snapshot=full;snapshotDocumentEnd=(int)doc.Content.End;snapshotStart=start;snapshotEnd=end;snapshotSelection=selected;found=result;
    hits.ItemsSource=null;hits.ItemsSource=found;status.Text=$"找到 {found.Count:N0} 项 · 单击跳转";PagePilot.Configuration.RuntimeLog.Write("host","regex.search.completed","revision="+version+" matches="+found.Count+" ms="+watch.ElapsedMilliseconds,"debug");
   }catch(OperationCanceledException){PagePilot.Configuration.RuntimeLog.Write("host","regex.search.cancelled","revision="+version,"debug");}catch(Exception e){PagePilot.Configuration.RuntimeLog.Write("host","regex.search.failed",e.GetType().Name+" stack="+e.StackTrace,"error");if(version==searchVersion){snapshot=null;hits.ItemsSource=null;status.Text=e.Message;}}finally{if(searchCancel==cancel)searchCancel=null;cancel.Dispose();}
  }
  private void UpdateReplacement(){
   if(snapshot==null||cachedText==null){ScheduleSearch();return;}
   try{var result=Engine.Find(cachedText,0,cachedQuery,replacement.Text,cachedIgnore);
    foreach(var hit in result){int a=hit.Start,b=a+hit.Length;hit.Start=cachedPositions==null?snapshotStart+a:cachedPositions[a];hit.Length=(cachedPositions==null?snapshotStart+b:(a==b?cachedPositions[a]:cachedEnds[b-1]))-hit.Start;}
    found=result;hits.ItemsSource=found;status.Text=$"找到 {found.Count:N0} 项 · 原查找范围保持不变";
   }catch(Exception e){snapshot=null;hits.ItemsSource=null;status.Text=e.Message;}
  }
  private void EnsureSnapshot(){if(snapshot==null||(string)doc.Content.Text!=snapshot)throw new Exception("文档内容已变化，请重新查找预览。");}
  private void Replace(){
   EnsureSnapshot();if(found.Count==0)return;if(snapshotSelection!=(selection.IsChecked==true))throw new Exception("查找范围已变化，请重新预览。");
   // Validate every target before the first write. Never leave partial replacements.
   foreach(var item in found)if((string)doc.Range(item.Start,item.Start+item.Length).Text!=item.Before)throw new Exception("匹配位置已变化，请重新查找。");
   EnsureSnapshot();NativeTools.Backup(doc);int count=0;bool recording=false;
   try{
    host.App.UndoRecord.StartCustomRecord("PagePilot 正则替换");recording=true;
    foreach(var item in found.AsEnumerable().Reverse()){doc.Range(item.Start,item.Start+item.Length).Text=item.After;count++;}
   }catch{
    if(recording){host.App.UndoRecord.EndCustomRecord();recording=false;}
    if(count>0)doc.Undo(1);
    throw;
   }finally{if(recording)host.App.UndoRecord.EndCustomRecord();snapshot=null;hits.ItemsSource=null;}
   status.Text=$"已替换 {count:N0} 项 · 可一步撤销";Connect.Log("regex.replace.completed count="+count);
  }

  internal static List<HeadingRule> ParseRules(string text){
   var list=new List<HeadingRule>();foreach(string line in text.Replace("\r","").Split('\n')){if(string.IsNullOrWhiteSpace(line))continue;var parts=line.Split(new[]{'|'},2);if(parts.Length!=2||!int.TryParse(parts[0],out int level)||level<1||level>9)throw new Exception("规则格式应为 1|^标题正则，级别为 1–9。");Engine.Compile(parts[1]);list.Add(new HeadingRule{Name="自定义",Level=level,Pattern=parts[1]});}list.AddRange(Engine.Defaults);return list;
  }
  private async void RefreshOutline(){
   scanCancel?.Cancel();var cancel=new System.Threading.CancellationTokenSource();scanCancel=cancel;
   int version=++outlineVersion;object document=doc;scanning=true;
   Find<Button>("StopScan").Visibility=Visibility.Visible;Find<Button>("Apply").IsEnabled=false;
   try{
    StartWork("正在读取文档结构");await System.Threading.Tasks.Task.Delay(20);cancel.Token.ThrowIfCancellationRequested();
    string rulesText=Connect.ReadRulesText();var rules=ParseRules(rulesText);var rows=new List<ParagraphInfo>();
    string before=(string)doc.Content.Text;int end=(int)doc.Content.End,count=(int)doc.Paragraphs.Count;
    var total=System.Diagnostics.Stopwatch.StartNew();
    rows=await NavigationReader.Read((object)doc,(i,n)=>ReportWork($"读取结构 · {i:N0} / {n:N0} 段",i,n),cancel.Token);
    ReportWork("正在归纳标题层级");
    var result=await System.Threading.Tasks.Task.Run(()=>NavigationEngine.Analyze(rows,rules,cancel.Token,host.NavigationStyle),cancel.Token);
    if(version!=outlineVersion||IsDisposed||!object.Equals(document,doc))return;
    if((string)doc.Content.Text!=before||(int)doc.Content.End!=end){status.Text="文档已变化，请重新分析";return;}
    NavigationReader.Resolve((object)doc,result);aiOutline=false;outline=result;outlineSnapshot=before;outlineEnd=end;ruleSnapshot=rulesText;RenderOutline();ApplyNavigationAppearance();
    status.Text=$"{outline.Count:N0} 个标题 · 单击定位 · 展开加载子目录";
    Connect.Log("navigation.scan.completed headings="+outline.Count+" paragraphs="+count+" ms="+total.ElapsedMilliseconds);
   }catch(OperationCanceledException){if(!IsDisposed&&version==outlineVersion)status.Text="分析已取消";}
   catch(Exception e){PagePilot.Configuration.RuntimeLog.Write("host","navigation.scan.failed",e.GetType().Name+" stack="+e.StackTrace,"error");if(!IsDisposed&&version==outlineVersion)status.Text="目录读取失败："+e.Message;}
   finally{if(scanCancel==cancel){scanCancel=null;scanning=false;if(!IsDisposed){Find<Button>("StopScan").Visibility=Visibility.Collapsed;Find<Button>("Apply").IsEnabled=navigationCancel==null;if(navigationCancel==null)FinishWork();}}cancel.Dispose();}
  }
  sealed class OutlineNode {internal Heading Heading;internal List<OutlineNode> Children=new List<OutlineNode>();}
  private void RenderOutline(){
   rebuilding=true;try{
    headings.Items.Clear();var roots=new List<OutlineNode>();var stack=new Stack<OutlineNode>();
    foreach(var item in outline){while(stack.Count>0&&stack.Peek().Heading.Level>=item.Level)stack.Pop();
     var node=new OutlineNode{Heading=item};if(stack.Count>0)stack.Peek().Children.Add(node);else roots.Add(node);stack.Push(node);}
    foreach(var root in roots)headings.Items.Add(CreateNode(root));
   }finally{rebuilding=false;}
  }
  private TreeViewItem CreateNode(OutlineNode model){
   var item=model.Heading;var label=new TextBlock{Text=item.Text.Length>120?item.Text.Substring(0,120)+"…":item.Text,TextWrapping=TextWrapping.Wrap};
   var node=new TreeViewItem{Header=label,Tag=item,Padding=new Thickness(4,host.NavigationStyle.RowPadding,8,host.NavigationStyle.RowPadding),HorizontalContentAlignment=HorizontalAlignment.Stretch,ToolTip="第 "+item.Paragraph+" 段 · "+item.Level+" 级\n"+item.Text};
   label.SetBinding(TextBlock.MaxWidthProperty,new System.Windows.Data.Binding("ActualWidth"){Source=headings,Converter=new NavigationWidth(),ConverterParameter=Tuple.Create(item.Level,host.NavigationStyle.Indent)});
   StyleLabel(label,item.Level);node.Selected+=(s,e)=>{if(e.OriginalSource==node)StyleLabel(label,item.Level,true);};node.Unselected+=(s,e)=>{if(e.OriginalSource==node)StyleLabel(label,item.Level);};label.MouseLeftButtonDown+=(s,e)=>{if(node.IsSelected)Guard(()=>Navigate(item));};
   if(model.Children.Count>0){node.Items.Add(new TreeViewItem{Header="加载子目录…"});bool loaded=false;
    node.Expanded+=(s,e)=>{if(e.OriginalSource!=node||loaded)return;loaded=true;node.Items.Clear();foreach(var child in model.Children)node.Items.Add(CreateNode(child));};}
   return node;
  }
  sealed class NavigationWidth:System.Windows.Data.IValueConverter {
   public object Convert(object value,Type target,object parameter,System.Globalization.CultureInfo culture)=>Math.Max(40,(double)value-64-(((Tuple<int,int>)parameter).Item1-1)*((Tuple<int,int>)parameter).Item2);
   public object ConvertBack(object value,Type target,object parameter,System.Globalization.CultureInfo culture)=>throw new NotSupportedException();
  }
  private void StyleLabel(TextBlock label,int level,bool selected=false){var style=host.NavigationStyle.Levels[Math.Max(0,Math.Min(8,level-1))];
   label.LineHeight=style.Size*1.5;label.MaxHeight=style.Size*1.5*host.NavigationStyle.MaxLines;label.TextTrimming=TextTrimming.CharacterEllipsis;
   label.FontFamily=new System.Windows.Media.FontFamily(style.Font);label.FontSize=style.Size;label.FontWeight=style.Bold?FontWeights.Bold:FontWeights.Normal;
   label.Foreground=new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(selected?host.NavigationStyle.SelectionForeground:style.Color));}
  private void ApplyNavigationAppearance(){
   if(appearanceVersion==host.NavigationStyleVersion)return;appearanceVersion=host.NavigationStyleVersion;
   var a=host.NavigationStyle;var background=(System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(a.Background);
   bool dark=background.R*0.2126+background.G*0.7152+background.B*0.0722<120;
   Action<string,string> brush=(key,color)=>view.Resources[key]=new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(color));
   brush("Canvas",a.Background);brush("Accent",a.Accent);brush("Selected",a.SelectionBackground);brush("SelectedInk",a.SelectionForeground);
   var accent=(System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(a.Accent);brush("AccentInk",accent.R*0.2126+accent.G*0.7152+accent.B*0.0722>150?"#102A26":"#FFFFFF");
   if(highlight!=null){highlight.MarkColor=System.Drawing.ColorTranslator.FromHtml(a.HighlightColor);highlight.Invalidate();}
   brush("Ink",dark?"#EDF2F7":"#1E293B");brush("Muted",dark?"#BAC8D9":"#526279");brush("Surface",dark?"#243143":"#FFFFFF");brush("Line",dark?"#46566D":"#D6DFE9");brush("Hover",dark?"#2D3E50":"#EDF2F7");
   view.Resources["TreeIndent"]=new Thickness(a.Indent,0,0,0);
   if(headings==null)return;
   Action<ItemsControl> apply=null;apply=parent=>{foreach(TreeViewItem n in parent.Items){if(n.Tag is Heading h&&n.Header is TextBlock t){StyleLabel(t,h.Level,n.IsSelected);n.Padding=new Thickness(4,a.RowPadding,8,a.RowPadding);t.SetBinding(TextBlock.MaxWidthProperty,new System.Windows.Data.Binding("ActualWidth"){Source=headings,Converter=new NavigationWidth(),ConverterParameter=Tuple.Create(h.Level,a.Indent)});}apply(n);}};apply(headings);
  }
  private readonly System.Windows.Forms.Timer workTimer=new System.Windows.Forms.Timer{Interval=1000};
  private System.Diagnostics.Stopwatch workWatch;private string workPhase;
  private void StartWork(string phase){workWatch=System.Diagnostics.Stopwatch.StartNew();workTimer.Stop();workTimer.Tick-=WorkTick;workTimer.Tick+=WorkTick;workTimer.Start();ReportWork(phase);}
  private void WorkTick(object sender,EventArgs e){if(!IsDisposed&&workWatch!=null)status.Text=workPhase+$" · 已用 {workWatch.Elapsed.TotalSeconds:F0} 秒";}
  private void ReportWork(string phase,int done=0,int total=0){workPhase=phase;var bar=Find<ProgressBar>("WorkProgress");if(bar!=null){bar.Visibility=Visibility.Visible;bar.IsIndeterminate=total<=0;bar.Maximum=Math.Max(1,total);bar.Value=done;}WorkTick(null,EventArgs.Empty);}
  private void FinishWork(){workTimer.Stop();workWatch=null;var bar=Find<ProgressBar>("WorkProgress");if(bar!=null)bar.Visibility=Visibility.Collapsed;}
  private void Navigate(Heading item){
   if((string)doc.Content.Text!=outlineSnapshot){if(aiOutline)throw new Exception("AI 目录已过期，请重新识别。");RefreshOutline();return;}
   doc.Activate();dynamic range=doc.Range(item.Start,item.Start);range.Select();dynamic window=host.App.ActiveWindow;window.ScrollIntoView(range,true);
   var viewport=ReadingPosition.Viewport(WordWindows.Handle(window));int target=viewport.Top+(int)(host.NavigationGap*WordWindows.Dpi(WordWindows.Handle(window)));
   for(int i=0;i<10;i++){int left=0,top=0,width=0,height=0;try{window.GetPoint(out left,out top,out width,out height,range);}catch{break;}
    int delta=top-target;if(Math.Abs(delta)<12)break;int lines=Math.Max(1,Math.Min(12,Math.Abs(delta)/Math.Max(18,height)));
    if(delta>0)window.SmallScroll(Down:lines);else window.SmallScroll(Up:lines);
    int next=0;window.GetPoint(out left,out next,out width,out height,range);if(next==top)break;
   }
  }
  private System.Threading.CancellationTokenSource navigationCancel;
  private async void AIOutline(){
   if(navigationCancel!=null){navigationCancel.Cancel();return;}
   scanCancel?.Cancel();outlineVersion++;
   var cancel=new System.Threading.CancellationTokenSource();navigationCancel=cancel;var button=Find<Button>("Analyze");
   button.Content="取消识别";Find<Button>("Apply").IsEnabled=false;Find<Button>("Refresh").IsEnabled=false;
   try{
    StartWork("准备 AI 识别");await System.Threading.Tasks.Task.Delay(20);
    object document=doc;string before=(string)doc.Content.Text;
    var metadata=await NavigationReader.Read((object)doc,(i,n)=>ReportWork($"读取结构 · {i:N0} / {n:N0} 段",i,n),cancel.Token);
    var rows=metadata.Where(p=>!NavigationEngine.Excluded(p,host.NavigationStyle.IncludeDocumentTitle)).Select(p=>new Heading{Paragraph=p.Paragraph,Start=p.Start,End=p.End,Text=p.Text}).ToList();
    var progress=new Progress<string>(text=>{if(!IsDisposed&&!cancel.IsCancellationRequested&&navigationCancel==cancel)ReportWork(text);});
    var result=await NavigationAi.Analyze(rows,Find<TextBox>("AiRequirement").Text,progress,cancel.Token,metadata,host.NavigationStyle);
    cancel.Token.ThrowIfCancellationRequested();if(IsDisposed)return;
    if(!object.Equals(document,doc)||(string)doc.Content.Text!=before)throw new Exception("文档已变化，请重新识别。");
    if(Find<CheckBox>("NavigationAutomatic").IsChecked!=true&&!host.Approval("更新导航目录（不修改文档）",result.Count+" 个标题")){status.Text="已取消更新目录";return;}
    if((string)doc.Content.Text!=before)throw new Exception("文档已变化，请重新识别。");
    NavigationReader.Resolve((object)doc,result);SetOutline(result);
   }catch(OperationCanceledException){if(!IsDisposed)status.Text="识别已取消";}
   catch(Exception e){Connect.Log("Navigation AI: "+e.GetType().Name);if(!IsDisposed)status.Text="识别失败："+(e is System.Net.Http.HttpRequestException?"模型服务连接失败，请检查网络与配置。":e.Message);}
   finally{navigationCancel=null;cancel.Dispose();if(!IsDisposed){FinishWork();button.Content="AI 识别";Find<Button>("Apply").IsEnabled=true;Find<Button>("Refresh").IsEnabled=true;}}
  }
  internal void SetOutline(List<Heading> items){scanCancel?.Cancel();outlineVersion++;aiOutline=true;outline=items;outlineSnapshot=(string)doc.Content.Text;outlineEnd=(int)doc.Content.End;ruleSnapshot=Connect.ReadRulesText();RenderOutline();status.Text=$"AI 识别 {outline.Count} 个标题 · 文档未修改";}
  private void ApplyOutline(){if((string)doc.Content.Text!=outlineSnapshot)throw new Exception("文档已变化，请刷新结构。");NativeTools.Backup(doc);foreach(var item in outline)doc.Paragraphs[item.Paragraph].OutlineLevel=item.Level;status.Text="已应用大纲级别 · 可使用 Word 原生导航";}
  protected override void Dispose(bool disposing){if(disposing){outlineVersion++;scanCancel?.Cancel();navigationCancel?.Cancel();searchCancel?.Cancel();workTimer.Dispose();highlightTimer.Dispose();highlight?.Dispose();searchVersion++;timer.Dispose();searchTimer.Dispose();}base.Dispose(disposing);}
 }
}
