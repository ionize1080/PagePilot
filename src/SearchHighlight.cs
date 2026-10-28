using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Windows.Forms;
namespace PagePilot {
 // Viewport-only decorations. Multiply a pale tint into captured pixels: black stays
 // black, unlike a translucent window painted over the glyphs. Never edits Word.
 internal sealed class SearchHighlight:Form {
  [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
  sealed class PhysicalPixels:IDisposable {
   readonly IntPtr previous;
   internal PhysicalPixels(){try{previous=SetThreadDpiAwarenessContext(new IntPtr(-4));}catch(EntryPointNotFoundException){}}
   public void Dispose(){if(previous!=IntPtr.Zero)SetThreadDpiAwarenessContext(previous);}
  }
  protected override void CreateHandle(){using(new PhysicalPixels())base.CreateHandle();}
  [DllImport("user32.dll")] static extern bool IsIconic(IntPtr h);
  [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h,out uint process);
  [DllImport("user32.dll")] static extern IntPtr GetAncestor(IntPtr h,uint flags);
  delegate void WinEvent(IntPtr hook,uint evt,IntPtr hwnd,int obj,int child,uint thread,uint time);
  [DllImport("user32.dll")] static extern IntPtr SetWinEventHook(uint min,uint max,IntPtr module,WinEvent callback,uint process,uint thread,uint flags);
  [DllImport("user32.dll")] static extern bool UnhookWinEvent(IntPtr hook);
  private readonly List<Rectangle> marks=new List<Rectangle>();private readonly HashSet<Rectangle> unique=new HashSet<Rectangle>();
  private readonly Timer frames=new Timer{Interval=16};private PaintListener listener;private WinEvent eventCallback;private IntPtr locationHook,scrollHook,owner,viewportHandle;
  private dynamic document,window;private List<Hit> current;private Rectangle viewport;private int stage,probe,first,last,cursor,anchor,forwardStart,misses;private bool painting,measuring,suspended=true;
  internal Color MarkColor {get;set;}=Color.FromArgb(184,120,10);
  private Bitmap surface;private SearchSnapshot mapping;
  internal void SetMapping(SearchSnapshot value){if(!ReferenceEquals(mapping,value)){mapping=value;geometry.Clear();}}
  private readonly Dictionary<int,Rectangle> geometry=new Dictionary<int,Rectangle>();
  private Rectangle captureBounds;private ulong? stableFingerprint;
  internal long TranslationFrames {get;private set;} internal long CapturedPixels {get;private set;}
  internal long PresentedFrames {get;private set;} internal long MaxFrameMilliseconds {get;private set;}
  private bool geometryDirty;
  private long invalidatedAt;
  internal long MeasurementSlices {get;private set;} internal long MaxSliceMilliseconds {get;private set;}
  protected override bool ShowWithoutActivation=>true;
  protected override CreateParams CreateParams{get{var p=base.CreateParams;p.ExStyle|=0x08000000|0x20|0x80;return p;}}
  internal SearchHighlight(){StartPosition=FormStartPosition.Manual;AutoScaleMode=AutoScaleMode.None;FormBorderStyle=FormBorderStyle.None;ShowInTaskbar=false;BackColor=Color.Magenta;TransparencyKey=Color.Magenta;DoubleBuffered=true;Opacity=1;frames.Tick+=(s,e)=>Frame();}
  internal new void Hide(){suspended=true;frames.Stop();stage=0;base.Hide();}
  internal void Update(dynamic doc,dynamic activeWindow,IntPtr activeOwner,List<Hit> hits){using(new PhysicalPixels())UpdateCore(doc,activeWindow,activeOwner,hits);}
  private void UpdateCore(dynamic doc,dynamic activeWindow,IntPtr activeOwner,List<Hit> hits){
   if(hits.Count==0||IsIconic(activeOwner)||GetForegroundWindow()!=activeOwner){Hide();return;}
   bool changed=suspended||owner!=activeOwner||!ReferenceEquals(current,hits)||!object.Equals(document,doc);suspended=false;
   document=doc;window=activeWindow;current=hits;
   if(owner!=activeOwner){Unlisten();owner=activeOwner;Listen();changed=true;}
   var area=ReadingPosition.Viewport(owner);if(area!=viewport){viewport=area;changed=true;}
   if(viewportHandle==IntPtr.Zero||!WordWindows.IsWindow(viewportHandle)){Unlisten();Listen();changed=true;}
   if(changed){geometry.Clear();InvalidateGeometry();}
  }
  private void Listen(){
   viewportHandle=ReadingPosition.ViewportHandle(owner);GetWindowThreadProcessId(owner,out uint pid);
   // Subclass only inside Word's process. WinEvent covers scrollbar/zoom and test hosts.
   if(pid==(uint)Process.GetCurrentProcess().Id&&viewportHandle!=IntPtr.Zero)try{listener=new PaintListener(viewportHandle,InvalidateGeometry);}catch{listener=null;}
   eventCallback=(hook,evt,hwnd,obj,child,thread,time)=>{if(!IsDisposed&&hwnd!=IntPtr.Zero&&(hwnd==owner||GetAncestor(hwnd,2)==owner)&&(evt!=0x800B||obj==0||obj==-4))InvalidateGeometry();};
   locationHook=SetWinEventHook(0x800B,0x800B,IntPtr.Zero,eventCallback,pid,0,0);scrollHook=SetWinEventHook(0x12,0x13,IntPtr.Zero,eventCallback,pid,0,0);
  }
  private void Unlisten(){listener?.ReleaseHandle();listener=null;if(locationHook!=IntPtr.Zero)UnhookWinEvent(locationHook);if(scrollHook!=IntPtr.Zero)UnhookWinEvent(scrollHook);locationHook=scrollHook=IntPtr.Zero;viewportHandle=IntPtr.Zero;}
  private void InvalidateGeometry(){
   if(IsDisposed||suspended||painting||measuring||current==null||current.Count==0)return;
   // Coalesce paint/input bursts without continually restarting partial geometry.
   // Never keep a captured glyph visible after the underlying Word view moves.
   bool wasVisible=Visible;painting=true;if(wasVisible)base.Hide();painting=false;
   if(!geometryDirty){invalidatedAt=Stopwatch.GetTimestamp();geometryDirty=true;}
   stableFingerprint=null;frames.Start();
  }
  private void Frame(){using(new PhysicalPixels())FrameCore();}
  private void FrameCore(){
   if(IsDisposed||current==null||GetForegroundWindow()!=owner||IsIconic(owner)){Hide();return;}
   var elapsed=Stopwatch.StartNew();
   try{
    if(geometryDirty){geometryDirty=false;stage=0;probe=0;first=int.MaxValue;last=-1;cursor=0;marks.Clear();unique.Clear();}
    if(stage!=3){MeasureSlice();return;}
    if(marks.Count==0){frames.Stop();return;}
    painting=true;CaptureSurface();var fingerprint=SurfaceFingerprint();
    // A pair of sparse samples detects Word's asynchronous smooth-scroll paint.
    // No fixed 100 ms debounce and no blocking DwmFlush on Word's input thread.
    if(stableFingerprint!=fingerprint){stableFingerprint=fingerprint;return;}
    frames.Stop();Bounds=viewport;
    var region=new Region();region.MakeEmpty();foreach(var mark in marks)region.Union(mark);
    var old=Region;Region=region;old?.Dispose();
    if(!Visible)Show(new Win32Owner(owner));Invalidate();PresentedFrames++;
   }catch{Hide();}finally{painting=false;MaxFrameMilliseconds=Math.Max(MaxFrameMilliseconds,elapsed.ElapsedMilliseconds);}
  }
  private bool TranslateGeometry(){
   if(geometry.Count<2)return false;
   var keys=new List<int>(geometry.Keys);keys.Sort();int a=keys[keys.Count/3],b=keys[Math.Min(keys.Count-1,keys.Count*2/3)];
   if(a>=current.Count||b>=current.Count||!GetRectangle(a,out var firstRect)||!GetRectangle(b,out var lastRect))return false;
   var oldFirst=geometry[a];var oldLast=geometry[b];
   int dx=firstRect.X-oldFirst.X,dy=firstRect.Y-oldFirst.Y;
   if(firstRect.Size!=oldFirst.Size||lastRect.Size!=oldLast.Size||lastRect.X-oldLast.X!=dx||lastRect.Y-oldLast.Y!=dy||Math.Abs(dy)>viewport.Height*3/4||Math.Abs(dx)>viewport.Width/2)return false;
   int low=int.MaxValue,high=-1;
   foreach(int key in keys){var r=geometry[key];r.Offset(dx,dy);if(AddMark(r)){geometry[key]=r;low=Math.Min(low,key);high=Math.Max(high,key);}else geometry.Remove(key);}
   if(high<0)return false;
   first=current[low].Start;last=current[high].Start+current[high].Length;anchor=low;cursor=low-1;forwardStart=high+1;misses=0;stage=1;TranslationFrames++;return true;
  }
  private void MeasureSlice(){
   var budget=Stopwatch.StartNew();measuring=true;try{
    if(stage==0){
     viewport=ReadingPosition.Viewport(owner);
     if(TranslateGeometry())return;geometry.Clear();
     int[] xs={2,1,3,2,1,3,2,1,3};int[] ys={0,0,0,4,4,4,2,2,2};
     while(probe<xs.Length&&budget.ElapsedMilliseconds<4){int n=probe++;try{dynamic r=window.RangeFromPoint(viewport.Left+viewport.Width*xs[n]/4,viewport.Top+10+(viewport.Height-20)*ys[n]/4);if((int)r.StoryType==1){first=Math.Min(first,(int)r.Start);last=Math.Max(last,(int)r.End);}}catch{}}
     if(probe<xs.Length)return;if(last<0){stage=3;return;}
     try{first=(int)document.Range(first,first).Paragraphs.First.Range.Start;last=(int)document.Range(last,last).Paragraphs.Last.Range.End;}catch{}
     int visibleStart=mapping==null?first:mapping.TextOffset(first);
     int lo=0,hi=current.Count;while(lo<hi){int mid=lo+(hi-lo)/2;if((mapping==null?current[mid].Start+current[mid].Length:current[mid].SourceStart+current[mid].SourceLength)<visibleStart)lo=mid+1;else hi=mid;}anchor=lo;forwardStart=lo;cursor=lo-1;misses=0;stage=1;
    }
    // Sampling supplies an anchor, not a hard visibility boundary. Walk outwards
    // until off-screen matches are reached, so large page margins do not hide hits.
    while(stage==1&&budget.ElapsedMilliseconds<4&&marks.Count<400){
     if(cursor<0||misses>=4){cursor=forwardStart;misses=0;stage=2;break;}
     misses=TryMark(cursor--)?0:misses+1;
    }
    while(stage==2&&budget.ElapsedMilliseconds<4&&marks.Count<400){
     if(cursor>=current.Count||misses>=4){stage=3;break;}
     misses=TryMark(cursor++)?0:misses+1;
    }
    if(marks.Count>=400)stage=3;
   }finally{measuring=false;MeasurementSlices++;MaxSliceMilliseconds=Math.Max(MaxSliceMilliseconds,budget.ElapsedMilliseconds);}
  }
  private bool GetRectangle(int index,out Rectangle rectangle){
   rectangle=Rectangle.Empty;try{var hit=current[index];mapping?.Resolve(hit);int x=0,y=0,w=0,h=0;
    window.GetPoint(out x,out y,out w,out h,document.Range(hit.Start,hit.Start+hit.Length));
    rectangle=new Rectangle(x-2-viewport.Left,y-1-viewport.Top,Math.Max(3,w)+4,Math.Max(2,h)+3);return true;
   }catch{return false;}
  }
  private bool AddMark(Rectangle rectangle){var r=Rectangle.Intersect(new Rectangle(Point.Empty,viewport.Size),rectangle);if(r.Width<=0||r.Height<=0)return false;if(unique.Add(r))marks.Add(r);return true;}
  private bool TryMark(int index){if(!GetRectangle(index,out var r))return false;if(!AddMark(r))return false;geometry[index]=r;return true;}
  // A deterministic probe for the Word integration harness; production uses Frame().
  internal void Measure(dynamic doc,dynamic activeWindow,IntPtr activeOwner,List<Hit> hits){suspended=false;document=doc;window=activeWindow;owner=activeOwner;current=hits;viewport=ReadingPosition.Viewport(owner);InvalidateGeometry();geometryDirty=false;stage=0;probe=0;marks.Clear();unique.Clear();var limit=Stopwatch.StartNew();while(stage!=3&&limit.ElapsedMilliseconds<250)MeasureSlice();frames.Stop();CaptureSurface();Bounds=viewport;}
  private void CaptureSurface(){
   captureBounds=Rectangle.Empty;foreach(var mark in marks)captureBounds=captureBounds.IsEmpty?mark:Rectangle.Union(captureBounds,mark);
   if(captureBounds.IsEmpty)return;
   if(surface==null||surface.Size!=captureBounds.Size){surface?.Dispose();surface=new Bitmap(captureBounds.Width,captureBounds.Height,PixelFormat.Format32bppRgb);}
   // Only copy matching tiles. The unused pixels between tiles are never read or
   // presented; neither CPU work nor capture bandwidth scales with monitor area.
   using(var g=Graphics.FromImage(surface))foreach(var mark in marks){
    g.CopyFromScreen(new Point(viewport.Left+mark.Left,viewport.Top+mark.Top),new Point(mark.Left-captureBounds.Left,mark.Top-captureBounds.Top),mark.Size,CopyPixelOperation.SourceCopy);
    CapturedPixels+=(long)mark.Width*mark.Height;
   }
  }
  private ulong SurfaceFingerprint(){
   var data=surface.LockBits(new Rectangle(Point.Empty,surface.Size),ImageLockMode.ReadOnly,PixelFormat.Format32bppRgb);
   try{ulong hash=14695981039346656037UL;unchecked{foreach(var mark in marks){
    int x=mark.X-captureBounds.X,y=mark.Y-captureBounds.Y;
    for(int row=0;row<mark.Height;row+=2)for(int col=0;col<mark.Width;col+=4)
     hash=(hash^(uint)Marshal.ReadInt32(data.Scan0,(y+row)*data.Stride+(x+col)*4))*1099511628211UL;
   }}return hash;}finally{surface.UnlockBits(data);}
  }
  protected override void OnPaint(PaintEventArgs e){
   if(surface==null)return;
   using(var attributes=new ImageAttributes())using(var pen=new Pen(MarkColor,1)){
    var tint=new ColorMatrix();tint.Matrix00=.65f+.35f*MarkColor.R/255f;tint.Matrix11=.65f+.35f*MarkColor.G/255f;tint.Matrix22=.65f+.35f*MarkColor.B/255f;
    attributes.SetColorMatrix(tint);
    foreach(var r in marks){var clip=Rectangle.Intersect(r,captureBounds.IsEmpty?new Rectangle(Point.Empty,surface.Size):captureBounds);if(clip.IsEmpty)continue;
     e.Graphics.DrawImage(surface,clip,clip.X-captureBounds.X,clip.Y-captureBounds.Y,clip.Width,clip.Height,GraphicsUnit.Pixel,attributes);
     e.Graphics.DrawRectangle(pen,clip.Left,clip.Top,Math.Max(1,clip.Width-1),Math.Max(1,clip.Height-1));
    }
   }
  }
  protected override void Dispose(bool disposing){if(disposing){frames.Dispose();Unlisten();surface?.Dispose();surface=null;mapping=null;geometry.Clear();document=null;window=null;current=null;}base.Dispose(disposing);}
  sealed class PaintListener:NativeWindow {
   readonly Action changed;internal PaintListener(IntPtr handle,Action action){changed=action;AssignHandle(handle);}
   protected override void WndProc(ref Message m){int msg=m.Msg;if(msg==0x20A||msg==0x20E||msg==0x114||msg==0x115||msg==0x100||msg==5||msg==3)changed();base.WndProc(ref m);if(msg==0xF)changed();}
  }
 }
}
