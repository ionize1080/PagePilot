using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
namespace PagePilot {
 // A click-through owned overlay. Search never changes document formatting or undo history.
 internal sealed class SearchHighlight:Form {
  [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] static extern bool IsIconic(IntPtr h);
  private readonly List<Rectangle> marks=new List<Rectangle>();
  protected override bool ShowWithoutActivation=>true;
  protected override CreateParams CreateParams{get{var p=base.CreateParams;p.ExStyle|=0x08000000|0x20|0x80;return p;}}
  internal SearchHighlight(){FormBorderStyle=FormBorderStyle.None;ShowInTaskbar=false;BackColor=Color.Magenta;TransparencyKey=Color.Magenta;DoubleBuffered=true;Opacity=0.45;}
  internal void Update(dynamic doc,dynamic window,IntPtr owner,List<Hit> hits){
   if(hits.Count==0||IsIconic(owner)||GetForegroundWindow()!=owner){Hide();return;}
   Measure(doc,window,owner,hits);if(marks.Count==0){Hide();return;}if(!Visible)Show(new Win32Owner(owner));Invalidate();
  }
  internal void Measure(dynamic doc,dynamic window,IntPtr owner,List<Hit> hits){
   var viewport=ReadingPosition.Viewport(owner);int first=int.MaxValue,last=-1;
   // Sample the visible text area. RangeFromPoint does not move the insertion point.
   for(int row=0;row<5;row++)for(int col=1;col<=3;col++)try{dynamic r=window.RangeFromPoint(viewport.Left+viewport.Width*col/4,viewport.Top+8+(viewport.Height-16)*row/4);if((int)r.StoryType==1){first=Math.Min(first,(int)r.Start);last=Math.Max(last,(int)r.End);}}catch{}
   marks.Clear();if(last<0){Hide();return;}
   // Include the edge paragraphs; GetPoint clips candidates to the actual viewport.
   try{first=(int)doc.Range(first,first).Paragraphs.First.Range.Start;last=(int)doc.Range(last,last).Paragraphs.Last.Range.End;}catch{}
   int lo=0,hi=hits.Count;while(lo<hi){int mid=lo+(hi-lo)/2;if(hits[mid].Start+hits[mid].Length<first)lo=mid+1;else hi=mid;}
   var elapsed=System.Diagnostics.Stopwatch.StartNew();
   for(int i=lo;i<hits.Count&&hits[i].Start<=last&&marks.Count<600&&elapsed.ElapsedMilliseconds<35;i++){
    var hit=hits[i];try{int x=0,y=0,w=0,h=0;window.GetPoint(out x,out y,out w,out h,doc.Range(hit.Start,hit.Start+hit.Length));
     var rect=Rectangle.Intersect(viewport,new Rectangle(x,y,Math.Max(3,w),Math.Max(2,h)));if(rect.IsEmpty)continue;rect.Offset(-viewport.Left,-viewport.Top);marks.Add(rect);
    }catch{}
   }
   Bounds=viewport;
  }
  protected override void OnPaint(PaintEventArgs e){using(var fill=new SolidBrush(Color.Gold))foreach(var r in marks)e.Graphics.FillRectangle(fill,r);}
 }
}
