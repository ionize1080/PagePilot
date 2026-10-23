using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
namespace PagePilot {
 internal static class WordWindows {
  delegate bool Callback(IntPtr h,IntPtr p);
  [DllImport("user32.dll")] static extern bool EnumWindows(Callback cb,IntPtr p);
  [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr h,StringBuilder b,int n);
  [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr h,StringBuilder b,int n);
  [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h,out uint pid);
  [DllImport("user32.dll")] internal static extern bool IsWindow(IntPtr h);
  [DllImport("user32.dll")] internal static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] static extern uint GetDpiForWindow(IntPtr h);
  static bool missingHwnd;
  internal static IntPtr Handle(dynamic window){
   if(!missingHwnd)try { return new IntPtr((int)window.Hwnd); } catch(Microsoft.CSharp.RuntimeBinder.RuntimeBinderException){missingHwnd=true;} catch(COMException){}
   string caption=(string)window.Caption;var candidates=new List<IntPtr>();var matches=new List<IntPtr>();
   uint current=(uint)Process.GetCurrentProcess().Id;
   EnumWindows((h,p)=>{uint pid;GetWindowThreadProcessId(h,out pid);if(pid!=current)return true;
    var cls=new StringBuilder(64);GetClassName(h,cls,64);if(cls.ToString()!="OpusApp")return true;
    candidates.Add(h);var title=new StringBuilder(1024);GetWindowText(h,title,1024);string t=title.ToString();
    if(t==caption||t.StartsWith(caption+" - ",StringComparison.Ordinal)||t.StartsWith(caption+" [",StringComparison.Ordinal)||t.StartsWith(caption+" (",StringComparison.Ordinal))matches.Add(h);
    return true;},IntPtr.Zero);
   if(matches.Count==1)return matches[0];if(candidates.Count==1)return candidates[0];
   throw new InvalidOperationException("无法唯一确定 Word 窗口，请激活目标文档后重试。");
  }
  internal static float Dpi(IntPtr h){try{uint d=GetDpiForWindow(h);if(d>0)return d/96f;}catch(EntryPointNotFoundException){}using(var g=Graphics.FromHwnd(IntPtr.Zero))return g.DpiX/96f;}
 }
 internal sealed class Win32Owner:System.Windows.Forms.IWin32Window { public IntPtr Handle{get;private set;} internal Win32Owner(IntPtr h){Handle=h;} }
}
