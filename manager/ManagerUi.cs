using PagePilot.Configuration;
using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PagePilot.Deployment {
 internal static partial class Manager {
  static void RunUi(){
   float scale;using(var g=Graphics.FromHwnd(IntPtr.Zero))scale=g.DpiX/96f;Func<int,int> px=x=>(int)Math.Round(x*scale);
   var ink=Color.FromArgb(32,42,47);var green=Color.FromArgb(22,116,95);var muted=Color.FromArgb(89,102,110);var border=Color.FromArgb(216,224,221);
   var area=Screen.PrimaryScreen.WorkingArea;
   using(var form=new Form{Text="PagePilot · 部署与诊断",Size=new Size(Math.Min(px(980),area.Width-40),Math.Min(px(760),area.Height-40)),MinimumSize=new Size(Math.Min(px(800),area.Width-40),Math.Min(px(600),area.Height-40)),StartPosition=FormStartPosition.CenterScreen,BackColor=Color.White,ForeColor=ink,Font=new Font("Microsoft YaHei UI",10),AutoScaleMode=AutoScaleMode.None}){
    bool busy=false;
    var shell=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=2,RowCount=1};shell.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,px(176)));shell.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));form.Controls.Add(shell);
    var rail=new FlowLayoutPanel{Dock=DockStyle.Fill,FlowDirection=FlowDirection.TopDown,WrapContents=false,BackColor=Color.FromArgb(241,246,244),Padding=new Padding(px(18),px(28),px(12),px(12))};shell.Controls.Add(rail,0,0);
    rail.Controls.Add(new Label{Text="PagePilot",AutoSize=true,Font=new Font(form.Font.FontFamily,18,FontStyle.Bold),ForeColor=green,Margin=new Padding(px(5),0,0,px(6))});
    rail.Controls.Add(new Label{Text="部署管理器  0.4.5",AutoSize=true,ForeColor=muted,Font=new Font(form.Font.FontFamily,9),Margin=new Padding(px(5),0,0,px(30))});
    var pages=new Panel{Dock=DockStyle.Fill};shell.Controls.Add(pages,1,0);
    Func<string,string,TableLayoutPanel> page=(title,desc)=>{
     var scroll=new Panel{Dock=DockStyle.Fill,AutoScroll=true,BackColor=Color.White,Visible=false};
     var layout=new TableLayoutPanel{Dock=DockStyle.Top,AutoSize=true,ColumnCount=1,Padding=new Padding(px(28),px(25),px(28),px(24))};layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));scroll.Controls.Add(layout);pages.Controls.Add(scroll);
     layout.Controls.Add(new Label{Text=title,AutoSize=true,Font=new Font(form.Font.FontFamily,22,FontStyle.Bold),Margin=new Padding(0,0,0,px(10))});
     layout.Controls.Add(new Label{Text=desc,AutoSize=true,ForeColor=muted,MaximumSize=new Size(px(650),0),Margin=new Padding(0,0,0,px(24))});
     return layout;
    };
    Func<string,bool,Button> button=(text,primary)=>new Button{Text=text,AutoSize=true,MinimumSize=new Size(px(100),px(38)),FlatStyle=FlatStyle.Flat,BackColor=primary?green:Color.White,ForeColor=primary?Color.White:ink,Padding=new Padding(px(12),px(5),px(12),px(5)),Margin=new Padding(0,0,px(10),px(10)),Cursor=Cursors.Hand,AccessibleName=text};
    Func<FlowLayoutPanel> row=()=>new FlowLayoutPanel{AutoSize=true,Dock=DockStyle.Top,WrapContents=true,Margin=new Padding(0)};
    Action<TableLayoutPanel,string> section=(layout,title)=>layout.Controls.Add(new Label{Text=title,AutoSize=true,Font=new Font(form.Font.FontFamily,11,FontStyle.Bold),Margin=new Padding(0,px(14),0,px(10))});
    var deploy=page("让 Word 准备就绪","选择注册范围。管理员 Word 请使用整台电脑范围；程序仍直接从解压目录运行。");
    var location=new Label{AutoSize=true,MaximumSize=new Size(px(640),0),Margin=new Padding(0,0,0,px(12)),ForeColor=muted};deploy.Controls.Add(location);
    var account=new Label{AutoSize=true,MaximumSize=new Size(px(640),0),Margin=new Padding(0,0,0,px(18))};deploy.Controls.Add(account);
    var summary=new Label{Text="正在检查…",AutoSize=true,MaximumSize=new Size(px(640),0),BackColor=Color.FromArgb(238,247,243),Padding=new Padding(px(14)),Margin=new Padding(0,0,0,px(20))};deploy.Controls.Add(summary);
    var scopePicker=new ComboBox{DropDownStyle=ComboBoxStyle.DropDownList,Width=px(340),AccessibleName="部署范围",Margin=new Padding(0,0,0,px(12))};
    scopePicker.Items.AddRange(new object[]{"当前用户（普通权限 Word）","整台电脑（管理员环境，需管理员权限）"});scopePicker.SelectedIndex=machineScope?1:0;deploy.Controls.Add(scopePicker);
    deploy.Controls.Add(new Label{Text="整台电脑范围会让其他账户的 Word 也可发现此加载项。替换时也清理当前账户的旧注册，并一并备份供回退。个人数据保留；移动前先停用。",AutoSize=true,MaximumSize=new Size(px(640),0),ForeColor=muted,Margin=new Padding(0,0,0,px(12))});
    var replaceRegistration=new CheckBox{Text="替换旧注册（备份后删除旧注册，再启用此版本）",AutoSize=true,Checked=replaceExisting,Enabled=machineScope,MaximumSize=new Size(px(640),0),Margin=new Padding(0,0,0,px(10)),AccessibleName="替换旧注册"};deploy.Controls.Add(replaceRegistration);
    replaceRegistration.CheckedChanged+=(s,e)=>replaceExisting=replaceRegistration.Checked;
    var actions=row();var enable=button("启用此版本",true);var verify=button("启动 Word 并验证",false);actions.Controls.AddRange(new Control[]{enable,verify});deploy.Controls.Add(actions);
    section(deploy,"版本操作");var secondary=row();var choose=button("选择版本目录",false);var disable=button("停用此版本",false);var rollback=button("回退软件版本",false);secondary.Controls.AddRange(new Control[]{choose,disable,rollback});deploy.Controls.Add(secondary);
    section(deploy,"Word 进程");var wordRow=row();var close=button("请求保存并关闭",false);var kill=button("强制关闭 Word",false);kill.ForeColor=Color.FromArgb(180,35,24);wordRow.Controls.AddRange(new Control[]{close,kill});deploy.Controls.Add(wordRow);
    deploy.Controls.Add(new Label{Text="强制关闭会丢失未保存的修改。只处理当前账户、当前桌面会话的 Word。部署中选择关闭后，会继续完成部署。",AutoSize=true,MaximumSize=new Size(px(640),0),ForeColor=muted,Margin=new Padding(0,0,0,px(12))});
    var feedback=new Label{Text="选择操作开始。日志会记录每一步。",AutoSize=true,MaximumSize=new Size(px(640),0),ForeColor=muted,Margin=new Padding(0,px(16),0,0)};deploy.Controls.Add(feedback);
    var storage=page("存储与日志","配置文件可以放在自选位置。保存后，下次启动 Word、AI 和设置时生效。");
    Func<string,string,TextBox> pathField=(label,value)=>{
     storage.Controls.Add(new Label{Text=label,AutoSize=true,Font=new Font(form.Font.FontFamily,10,FontStyle.Bold),Margin=new Padding(0,px(8),0,px(5))});
     var line=new TableLayoutPanel{AutoSize=true,Dock=DockStyle.Top,ColumnCount=2,Margin=new Padding(0,0,0,px(8))};line.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));line.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
     var input=new TextBox{Text=value,Dock=DockStyle.Top,Margin=new Padding(0,px(5),px(8),0),AccessibleName=label};var browse=button("浏览…",false);browse.MinimumSize=new Size(px(76),px(32));line.Controls.Add(input,0,0);line.Controls.Add(browse,1,0);storage.Controls.Add(line);
     browse.Click+=(s,e)=>{if(label=="运行配置文件"){using(var dialog=new SaveFileDialog{Title="选择或新建运行配置文件",Filter="JSON 配置|*.json",FileName=input.Text,OverwritePrompt=false})if(dialog.ShowDialog(form)==DialogResult.OK)input.Text=dialog.FileName;}else using(var dialog=new FolderBrowserDialog{Description=label,SelectedPath=input.Text})if(dialog.ShowDialog(form)==DialogResult.OK)input.Text=dialog.SelectedPath;};return input;
    };
    var conf=pathField("运行配置文件",configFile);var dat=pathField("本地规则与文档备份",RuntimePaths.Current.DataDirectory);var profile=pathField("AI 模型配置与会话",RuntimePaths.Current.ProfileDirectory);var logs=pathField("日志存放目录",RuntimePaths.Current.LogDirectory);
    var debug=new CheckBox{Text="详细日志（包含高频状态与耗时，不记录正文和密钥）",Checked=RuntimePaths.Current.LogLevel=="debug",AutoSize=true,Margin=new Padding(0,px(8),0,px(8))};storage.Controls.Add(debug);
    var migrate=new CheckBox{Text="复制旧配置与会话到新目录，保留原文件且不覆盖已有文件",Checked=true,AutoSize=true,MaximumSize=new Size(px(630),0),Margin=new Padding(0,0,0,px(16))};storage.Controls.Add(migrate);
    var configActions=row();var save=button("保存并使用配置",true);var load=button("读取所选配置",false);configActions.Controls.AddRange(new Control[]{save,load});storage.Controls.Add(configActions);
    var configNote=new Label{Text="目录变更前请关闭 Word。部署回退记录固定保留在默认本地目录，避免更换配置后找不到回退点。",AutoSize=true,MaximumSize=new Size(px(640),0),ForeColor=muted,Margin=new Padding(0,px(8),0,0)};storage.Controls.Add(configNote);
    var diagnostic=page("检查加载情况","注册、Word 连接、功能区加载分别检查。诊断包不包含模型密钥、会话记录或文档副本。");
    var diagActions=row();var refresh=button("完整诊断",true);var openLogs=button("打开日志文件夹",false);var export=button("导出诊断包",false);diagActions.Controls.AddRange(new Control[]{refresh,openLogs,export});diagnostic.Controls.Add(diagActions);
    var report=new TextBox{Multiline=true,ReadOnly=true,ScrollBars=ScrollBars.Both,WordWrap=false,Dock=DockStyle.Top,Height=px(360),BackColor=Color.FromArgb(248,250,249),BorderStyle=BorderStyle.FixedSingle,Font=new Font("Microsoft YaHei UI",9),AccessibleName="部署诊断详情",Margin=new Padding(0,px(12),0,0)};diagnostic.Controls.Add(report);
    var navs=new System.Collections.Generic.List<Button>();
    Action<TableLayoutPanel,string> nav=(layout,title)=>{var b=button(title,false);b.BackColor=rail.BackColor;b.Width=px(136);b.AutoSize=false;b.FlatAppearance.BorderSize=0;b.TextAlign=ContentAlignment.MiddleLeft;b.Click+=(s,e)=>{foreach(Control p in pages.Controls)p.Visible=p==layout.Parent;layout.Parent.BringToFront();foreach(var n in navs){n.BackColor=rail.BackColor;n.ForeColor=ink;}b.BackColor=Color.FromArgb(218,235,226);b.ForeColor=green;Event("ui.page",title);};navs.Add(b);rail.Controls.Add(b);};
    nav(deploy,"部署与验证");nav(storage,"存储与日志");nav(diagnostic,"诊断详情");
    var quickLogs=button("打开日志文件夹",false);quickLogs.FlatAppearance.BorderSize=0;quickLogs.BackColor=rail.BackColor;quickLogs.Width=px(136);quickLogs.AutoSize=false;quickLogs.Margin=new Padding(0,px(30),0,0);rail.Controls.Add(quickLogs);
    Action update=()=>{location.Text="版本目录\r\n"+root;account.Text=System.Security.Principal.WindowsIdentity.GetCurrent().Name+"   /   "+(Elevated?"管理员运行：可使用整台电脑注册":"普通权限：仅可修改当前用户注册");account.ForeColor=Elevated?Color.FromArgb(157,83,0):muted;try{summary.Text=IsRegistered(WordView())&&WordProcesses().Any(Loaded)?"已验证：Word 已连接，PagePilot 功能区已加载。":IsRegistered(WordView())?"已注册此版本 · 等待 Word 实际加载验证":"尚未启用此版本 · 请先完成部署";}catch(Exception ex){summary.Text=ex.Message;}};
    scopePicker.SelectedIndexChanged+=(s,e)=>{machineScope=scopePicker.SelectedIndex==1;replaceRegistration.Enabled=machineScope;if(!machineScope)replaceRegistration.Checked=false;Event("scope.selected",ScopeName);update();};
    Func<string,Func<Task<string>>,Task> run=async(name,action)=>{if(busy)return;busy=true;pages.Enabled=false;rail.Enabled=false;feedback.Text=name+"…";summary.Text=name+"，请稍候…";Event("ui.action.begin",name);var watch=Stopwatch.StartNew();try{string message=await action();if(message!=null){report.Text=message;report.SelectionStart=0;report.SelectionLength=0;}feedback.Text=name+"已完成";Event("ui.action.completed",name+" ms="+watch.ElapsedMilliseconds);}catch(Exception ex){feedback.Text=name+"未完成："+ex.Message;Event("ui.action.failed",name+" "+ex,"error");report.Text=ex+"\r\n\r\n"+await Task.Run(()=>Status(false));MessageBox.Show(form,ex.Message,"PagePilot · 未完成",MessageBoxButtons.OK,MessageBoxIcon.Warning);}finally{update();busy=false;pages.Enabled=true;rail.Enabled=true;}};
    Func<bool> confirmForce=()=>MessageBox.Show(form,"强制关闭会丢失所有未保存的 Word 修改。\r\n只结束当前账户、当前会话的 Word 进程。\r\n\r\n确认强制关闭？","强制关闭 Word",MessageBoxButtons.YesNo,MessageBoxIcon.Warning,MessageBoxDefaultButton.Button2)==DialogResult.Yes;
    Func<bool?> closure=()=>{if(WordProcesses().Length==0)return false;var answer=MessageBox.Show(form,"部署需要 Word 完全退出。\r\n\r\n“是”：请求 Word 正常关闭，请完成保存提示。\r\n“否”：选择强制关闭，未保存修改可能丢失。\r\n“取消”：返回。","先关闭 Word",MessageBoxButtons.YesNoCancel,MessageBoxIcon.Information,MessageBoxDefaultButton.Button3);if(answer==DialogResult.Cancel)return null;if(answer==DialogResult.No)return confirmForce()?(bool?)true:null;return false;};
    Action<Button,string,string> mutation=(b,label,command)=>b.Click+=async(s,e)=>{try{RequireScopePermission();}catch(Exception ex){MessageBox.Show(form,ex.Message,"PagePilot");return;}bool? force=closure();if(force==null)return;await run(label,()=>Task.Run(()=>{ReadyConfig();Exclusive(()=>{CloseSelected(WordProcesses(),force.Value);if(command=="rollback")Rollback();else Change(command,WordView());});return Status(false);}));};
    mutation(enable,"启用此版本","enable");mutation(disable,"停用此版本","disable");mutation(rollback,"回退软件版本","rollback");
    verify.Click+=async(s,e)=>await run("验证 Word 加载",VerifyWord);
    close.Click+=async(s,e)=>await run("请求关闭 Word",()=>Task.Run(()=>{Exclusive(()=>CloseSelected(WordProcesses(),false));return Status(false);}));
    kill.Click+=async(s,e)=>{if(confirmForce())await run("强制关闭 Word",()=>Task.Run(()=>{Exclusive(()=>CloseSelected(WordProcesses(),true));return Status(false);}));};
    choose.Click+=(s,e)=>{using(var dialog=new FolderBrowserDialog{Description="选择完整解压的 PagePilot 目录",SelectedPath=root})if(dialog.ShowDialog(form)==DialogResult.OK){root=Path.GetFullPath(dialog.SelectedPath).TrimEnd('\\');Event("root.selected",root);update();}};
    save.Click+=async(s,e)=>{string file=conf.Text;var settings=new RuntimeSettings{DataDirectory=dat.Text,ProfileDirectory=profile.Text,LogDirectory=logs.Text,LogLevel=debug.Checked?"debug":"info"};bool copy=migrate.Checked;await run("保存存储配置",()=>Task.Run(()=>{SaveConfiguration(settings,file,copy);return Status(false);}));configNote.Text=RuntimePaths.Warning??"当前配置："+configFile+"。请重新启动 Word 后使用。";};
    load.Click+=async(s,e)=>{string file=conf.Text;RuntimeSettings settings=null;await run("读取配置",()=>Task.Run(()=>{settings=RuntimeSettings.Load(file);Event("config.preview",file);return "已读取配置，尚未保存为当前配置。";}));if(settings!=null){dat.Text=settings.DataDirectory;profile.Text=settings.ProfileDirectory;logs.Text=settings.LogDirectory;debug.Checked=settings.LogLevel=="debug";}};
    refresh.Click+=async(s,e)=>await run("完整诊断",()=>Task.Run(()=>Status(true)));
    Action openLogFolder=()=>{try{OpenLogs();}catch(Exception ex){Event("logs.open.failed",ex.GetType().Name,"error");MessageBox.Show(form,ex.Message);}};openLogs.Click+=(s,e)=>openLogFolder();quickLogs.Click+=(s,e)=>openLogFolder();
    export.Click+=async(s,e)=>{using(var dialog=new SaveFileDialog{Filter="诊断压缩包|*.zip",FileName="PagePilot-diagnostics-"+DateTime.Now.ToString("yyyyMMdd-HHmmss")+".zip",OverwritePrompt=true})if(dialog.ShowDialog(form)==DialogResult.OK){if(File.Exists(dialog.FileName)){MessageBox.Show(form,"请选择新的文件名，避免覆盖已有诊断包。");return;}await run("导出诊断包",()=>Task.Run(()=>"诊断包已保存："+ExportDiagnostics(dialog.FileName)));}};
    form.FormClosing+=(s,e)=>{if(busy){e.Cancel=true;return;}Event("manager.closed");};form.Shown+=(s,e)=>{navs[0].PerformClick();update();enable.Focus();};Application.Run(form);
   }
  }
 }
}
