using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
namespace PagePilot {
 internal static class ReadingPosition {
  private delegate bool Visitor(IntPtr h,IntPtr p);
  [StructLayout(LayoutKind.Sequential)]private struct Rect {public int left,top,right,bottom;}
  [DllImport("user32.dll")]private static extern bool EnumChildWindows(IntPtr h,Visitor visitor,IntPtr p);
  [DllImport("user32.dll")]private static extern int GetClassName(IntPtr h,StringBuilder s,int n);
  [DllImport("user32.dll")]private static extern bool GetWindowRect(IntPtr h,out Rect r);
  [DllImport("user32.dll")]private static extern bool IsWindowVisible(IntPtr h);
  internal static IntPtr ViewportHandle(IntPtr owner){IntPtr result=IntPtr.Zero;long area=0;EnumChildWindows(owner,(h,p)=>{var name=new StringBuilder(80);GetClassName(h,name,80);if(name.ToString()=="_WwG"&&IsWindowVisible(h)){GetWindowRect(h,out var r);long a=(long)(r.right-r.left)*(r.bottom-r.top);if(a>area){area=a;result=h;}}return true;},IntPtr.Zero);return result;}
  internal static Rectangle Viewport(IntPtr owner){
   GetWindowRect(owner,out var fallback);var result=Rectangle.FromLTRB(fallback.left,fallback.top+180,fallback.right,fallback.bottom);long area=0;
   EnumChildWindows(owner,(h,p)=>{var name=new StringBuilder(80);GetClassName(h,name,80);if(name.ToString()=="_WwG"&&IsWindowVisible(h)){GetWindowRect(h,out var r);long a=(long)(r.right-r.left)*(r.bottom-r.top);if(a>area){area=a;result=Rectangle.FromLTRB(r.left,r.top,r.right,r.bottom);}}return true;},IntPtr.Zero);return result;
  }
 }
}
