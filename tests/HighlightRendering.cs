using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using PagePilot;
class HighlightRendering {
 [DllImport("user32.dll")]static extern IntPtr SendMessage(IntPtr h,int msg,IntPtr w,IntPtr l);
 static BindingFlags flags=BindingFlags.Instance|BindingFlags.Public|BindingFlags.NonPublic;
 [STAThread]static void Main(){try{
  var type=typeof(Heading).Assembly.GetType("PagePilot.SearchHighlight");using(var overlay=(Form)Activator.CreateInstance(type,true)){
   if(overlay.Opacity!=1)throw new Exception("Alpha overlay would fade glyphs");
   if(overlay.StartPosition!=FormStartPosition.Manual)throw new Exception("Windows default placement would override the measured viewport on first show");
   var marks=(List<Rectangle>)type.GetField("marks",flags).GetValue(overlay);marks.Add(new Rectangle(10,10,80,30));var color=Color.FromArgb(65,95,205);type.GetProperty("MarkColor",flags).SetValue(overlay,color);
   using(var bitmap=new Bitmap(100,60))using(var graphics=Graphics.FromImage(bitmap)){
    var source=new Bitmap(100,60);using(var original=Graphics.FromImage(source)){original.Clear(Color.White);using(var black=new SolidBrush(Color.Black))original.FillRectangle(black,20,20,40,8);using(var gray=new SolidBrush(Color.FromArgb(90,90,90)))original.FillRectangle(gray,20,30,40,4);}
    type.GetField("surface",flags).SetValue(overlay,source);graphics.Clear(Color.Magenta);
    type.GetMethod("OnPaint",flags).Invoke(overlay,new object[]{new PaintEventArgs(graphics,new Rectangle(0,0,100,60))});
    if(bitmap.GetPixel(30,22).ToArgb()!=Color.Black.ToArgb())throw new Exception("Text pixels changed");
    var fill=bitmap.GetPixel(55,15);if(fill.ToArgb()==Color.White.ToArgb()||fill.ToArgb()==Color.Magenta.ToArgb()||fill.B<=fill.R)throw new Exception("Custom tinted background missing");
    var grayPixel=bitmap.GetPixel(30,31);if(grayPixel.R>90||grayPixel.G>90||grayPixel.B>90)throw new Exception("Antialiased text washed out");
    var before=bitmap.GetPixel(55,15);marks.Add(new Rectangle(10,10,80,30));type.GetMethod("OnPaint",flags).Invoke(overlay,new object[]{new PaintEventArgs(graphics,new Rectangle(0,0,100,60))});if(bitmap.GetPixel(55,15)!=before)throw new Exception("Overlapping highlights accumulate color");
    if(bitmap.GetPixel(10,10).ToArgb()!=color.ToArgb())throw new Exception("Custom highlight color not rendered");
    bitmap.Save(".temp/highlight-pixel-test.png");
   }
   using(var window=new Form()){var handle=window.Handle;int invalidations=0;var nested=type.GetNestedType("PaintListener",BindingFlags.NonPublic);var listener=(NativeWindow)Activator.CreateInstance(nested,flags,null,new object[]{handle,new Action(()=>invalidations++)},null);
    foreach(int msg in new[]{0x20A,0x20E,0x114,0x115,0x100,5,3})SendMessage(handle,msg,IntPtr.Zero,IntPtr.Zero);
    if(invalidations<7)throw new Exception("Scroll/resize/keyboard invalidation missing");listener.ReleaseHandle();int before=invalidations;SendMessage(handle,0x20A,IntPtr.Zero,IntPtr.Zero);if(before!=invalidations)throw new Exception("Hook not released");
   }
  }
  Console.WriteLine("PASS filled custom-color background / unchanged black pixels / no gray washout / no overlapping tint accumulation / wheel, horizontal scroll, scrollbar, keyboard, resize hooks / cleanup");
 }catch(Exception e){Console.WriteLine(e);Environment.ExitCode=1;}}
}
