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
  private Connect host;private dynamic doc;private int page,windowHandle,appliedTheme=-1;
  private FrameworkElement view;private TextBox pattern,replacement;private CheckBox ignore,selection;
  private ListBox hits;private TreeView headings;private bool rebuilding;private TextBlock status;
  private List<Hit> found=new List<Hit>();private string snapshot;private int snapshotStart,snapshotEnd;private bool snapshotSelection;
  private List<Heading> outline=new List<Heading>();private string outlineSnapshot,ruleSnapshot;
  private readonly System.Windows.Forms.Timer timer=new System.Windows.Forms.Timer{Interval=1800};
  public Pane(){AutoScaleMode=System.Windows.Forms.AutoScaleMode.None;}
  private T Find<T>(string name) where T:class=>view.FindName(name) as T;
  private void On(string name,Action action){Find<Button>(name).Click+=(s,e)=>Guard(action);}
  internal void Bind(Connect h,object document,int mode){
   host=h;doc=document;page=mode;windowHandle=(int)h.App.ActiveWindow.Hwnd;
   using(var stream=typeof(Pane).Assembly.GetManifestResourceStream("PagePilot.Views."+(mode==0?"Regex":"Navigation")+".xaml"))view=(FrameworkElement)XamlReader.Load(stream);
   Controls.Add(new ElementHost{Dock=System.Windows.Forms.DockStyle.Fill,Child=view});status=Find<TextBlock>("Status");
   On("Settings",()=>host.OpenSettings(null));
   ApplyTheme();
   if(page==0){
    pattern=Find<TextBox>("Pattern");replacement=Find<TextBox>("Replacement");ignore=Find<CheckBox>("IgnoreCase");selection=Find<CheckBox>("SelectionOnly");hits=Find<ListBox>("Matches");
    On("Preview",FindMatches);On("Replace",Replace);
    hits.MouseDoubleClick+=(s,e)=>Guard(()=>{var item=hits.SelectedItem as Hit;if(item!=null){EnsureSnapshot();doc.Activate();doc.Range(item.Start,item.Start+item.Length).Select();}});
   }else{
    headings=Find<TreeView>("Headings");On("Refresh",RefreshOutline);On("Analyze",AIOutline);On("Apply",ApplyOutline);
    headings.SelectedItemChanged+=(s,e)=>{if(!rebuilding)Guard(()=>{var item=(headings.SelectedItem as TreeViewItem)?.Tag as Heading;if(item!=null)Navigate(item);});};

   }
   timer.Tick+=(s,e)=>{if(Visible){try{ApplyTheme();RefreshBinding();if(page==1&&((string)doc.Content.Text!=outlineSnapshot||Connect.ReadRulesText()!=ruleSnapshot))RefreshOutline();}catch{}}};timer.Start();
  }
  private void ApplyTheme(){if(appliedTheme==host.Theme)return;appliedTheme=host.Theme;
   view.Resources["Accent"]=new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(host.Theme==1?"#4671D6":host.Theme==2?"#515E6F":"#238B73"));
   using(var bitmap=RibbonIcons.Draw("ppSettings",24,host.Theme))using(var memory=new MemoryStream()){bitmap.Save(memory,System.Drawing.Imaging.ImageFormat.Png);memory.Position=0;var source=new System.Windows.Media.Imaging.BitmapImage();source.BeginInit();source.CacheOption=System.Windows.Media.Imaging.BitmapCacheOption.OnLoad;source.StreamSource=memory;source.EndInit();source.Freeze();Find<Button>("Settings").Content=new Image{Source=source,Width=18,Height=18};}
  }
  private void RefreshBinding(){
   if(host.App.Documents.Count==0||(int)host.App.ActiveWindow.Hwnd!=windowHandle)return;
   dynamic current=host.App.ActiveDocument;bool same=false;
   try{same=object.Equals(doc,current)&&(string)doc.FullName==(string)current.FullName;}catch{}
   if(same)return;
   doc=current;snapshot=null;outlineSnapshot=null;found.Clear();outline.Clear();
   if(hits!=null)hits.ItemsSource=null;if(headings!=null)headings.Items.Clear();
   status.Text="已切换文档 · 请重新预览";
  }
  public void SelectPage(int index){Guard(()=>{if(page==1&&outlineSnapshot==null)RefreshOutline();});}
  private void Guard(Action action){try{RefreshBinding();action();}catch(Exception e){status.Text=e.Message;System.Windows.Forms.MessageBox.Show(e.Message,"PagePilot");}}
  private void FindMatches(){
   dynamic range=selection.IsChecked==true?host.App.Selection.Range:doc.Content;
   if(selection.IsChecked==true&&!object.Equals(host.App.ActiveDocument,doc))throw new Exception("请激活此窗格对应的文档后再查找。");
   snapshot=(string)doc.Content.Text;snapshotStart=(int)range.Start;snapshotEnd=(int)range.End;snapshotSelection=selection.IsChecked==true;
   found=Engine.Find((string)range.Text,snapshotStart,pattern.Text,replacement.Text,ignore.IsChecked==true);
   hits.ItemsSource=null;hits.ItemsSource=found;status.Text=$"找到 {found.Count} 项 · 双击结果定位";
  }
  private void EnsureSnapshot(){if(snapshot==null||(string)doc.Content.Text!=snapshot)throw new Exception("文档内容已变化，请重新查找预览。");}
  private void Replace(){
   EnsureSnapshot();if(found.Count==0)return;if(snapshotSelection!=(selection.IsChecked==true))throw new Exception("查找范围已变化，请重新预览。");
   var check=Engine.Find((string)doc.Range(snapshotStart,snapshotEnd).Text,snapshotStart,pattern.Text,replacement.Text,ignore.IsChecked==true);
   if(Connect.Json().Serialize(check)!=Connect.Json().Serialize(found))throw new Exception("查找或替换条件已变化，请重新预览。");
   EnsureSnapshot();NativeTools.Backup(doc);int count=0;
   try{foreach(var item in found.AsEnumerable().Reverse()){dynamic r=doc.Range(item.Start,item.Start+item.Length);if((string)r.Text!=item.Before)throw new Exception("匹配位置已变化。");r.Text=item.After;count++;}}
   finally{snapshot=null;hits.ItemsSource=null;status.Text=$"已替换 {count} 项 · 已保存操作前备份";}
  }
  internal static List<HeadingRule> ParseRules(string text){
   var list=new List<HeadingRule>();foreach(string line in text.Replace("\r","").Split('\n')){if(string.IsNullOrWhiteSpace(line))continue;var parts=line.Split(new[]{'|'},2);if(parts.Length!=2||!int.TryParse(parts[0],out int level)||level<1||level>9)throw new Exception("规则格式应为 1|^标题正则，级别为 1–9。");Engine.Compile(parts[1]);list.Add(new HeadingRule{Name="自定义",Level=level,Pattern=parts[1]});}list.AddRange(Engine.Defaults);return list;
  }
  private void RefreshOutline(){
   ruleSnapshot=Connect.ReadRulesText();var rules=ParseRules(ruleSnapshot);outline=new List<Heading>();outlineSnapshot=(string)doc.Content.Text;
   int count=(int)doc.Paragraphs.Count;for(int i=1;i<=count;i++){dynamic p=doc.Paragraphs[i];string text=((string)p.Range.Text).Trim('\r','\a',' ','\t');int level=Engine.Level(text,rules);if(level>0)outline.Add(new Heading{Paragraph=i,Start=(int)p.Range.Start,End=(int)p.Range.End-1,Level=level,Text=text});}
   RenderOutline();status.Text=$"识别 {outline.Count} 个标题 · 单击跳转";
  }
  private void RenderOutline(){
   rebuilding=true;try{headings.Items.Clear();var stack=new Stack<TreeViewItem>();foreach(var item in outline){
    while(stack.Count>0&&((Heading)stack.Peek().Tag).Level>=item.Level)stack.Pop();
    var node=new TreeViewItem{Header=new TextBlock{Text=item.Text,TextWrapping=TextWrapping.Wrap,MaxWidth=290},Tag=item,IsExpanded=true,Padding=new Thickness(4,7,4,7),ToolTip="第 "+item.Paragraph+" 段 · "+item.Level+" 级"};
    ((TextBlock)node.Header).MouseLeftButtonDown+=(s,e)=>{if(node.IsSelected)Guard(()=>Navigate(item));};
    if(stack.Count>0)stack.Peek().Items.Add(node);else headings.Items.Add(node);stack.Push(node);
   }}finally{rebuilding=false;}
  }
  private void Navigate(Heading item){
   if((string)doc.Content.Text!=outlineSnapshot){RefreshOutline();return;}
   doc.Activate();dynamic range=doc.Range(item.Start,item.Start);range.Select();dynamic window=host.App.ActiveWindow;window.ScrollIntoView(range,true);
   var viewport=ReadingPosition.Viewport((IntPtr)(int)window.Hwnd);int target=viewport.Top+(int)(host.NavigationGap*Native.Dpi((IntPtr)(int)window.Hwnd));
   for(int i=0;i<10;i++){int left=0,top=0,width=0,height=0;try{window.GetPoint(out left,out top,out width,out height,range);}catch{break;}
    int delta=top-target;if(Math.Abs(delta)<12)break;int lines=Math.Max(1,Math.Min(12,Math.Abs(delta)/Math.Max(18,height)));
    if(delta>0)window.SmallScroll(Down:lines);else window.SmallScroll(Up:lines);
    int next=0;window.GetPoint(out left,out next,out width,out height,range);if(next==top)break;
   }
  }
  private void AIOutline(){doc.Activate();string requirement=Find<TextBox>("AiRequirement").Text;
   host.QueuePrompt("分析当前 Word 文档结构。要求："+(string.IsNullOrWhiteSpace(requirement)?"识别全部标题层级":requirement)+"。先读取 pagepilot_capabilities 和 pagepilot_context，再调用 pagepilot_set_navigation，按段落索引和 1–9 级给出结果。不要修改文档文字、格式或大纲级别。");status.Text="要求已放入 AI 输入框 · 发送后生成目录";
  }
  internal void SetOutline(List<Heading> items){outline=items;outlineSnapshot=(string)doc.Content.Text;ruleSnapshot=Connect.ReadRulesText();RenderOutline();status.Text=$"AI 识别 {outline.Count} 个标题 · 文档未修改";}
  private void ApplyOutline(){if((string)doc.Content.Text!=outlineSnapshot)throw new Exception("文档已变化，请刷新结构。");NativeTools.Backup(doc);foreach(var item in outline)doc.Paragraphs[item.Paragraph].OutlineLevel=item.Level;status.Text="已应用大纲级别 · 可使用 Word 原生导航";}
  protected override void Dispose(bool disposing){if(disposing)timer.Dispose();base.Dispose(disposing);}
 }
}
