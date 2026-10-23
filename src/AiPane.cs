using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;




namespace PagePilot
{
    [ComVisible(true), Guid("246818DC-B9A5-4328-9E71-69575A825B92"), ProgId("PagePilot.AiPane")]
    public sealed class AiPane : UserControl
    {
        private IntPtr child,owner;
        private int originalStyle,originalExtendedStyle;
        [StructLayout(LayoutKind.Sequential)] private struct WindowRect { public int left,top,right,bottom; }
        private WindowRect floatingBounds;
        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr hwnd,out WindowRect rect);
        [DllImport("user32.dll")] private static extern bool GetClientRect(IntPtr hwnd,out WindowRect rect);
        [DllImport("user32.dll")] private static extern bool ClientToScreen(IntPtr hwnd,ref System.Drawing.Point point);
        [DllImport("gdi32.dll")] private static extern IntPtr CreateRectRgn(int left,int top,int right,int bottom);
        [DllImport("gdi32.dll")] private static extern bool DeleteObject(IntPtr handle);
        [DllImport("user32.dll")] private static extern int SetWindowRgn(IntPtr hwnd,IntPtr region,bool redraw);
        [StructLayout(LayoutKind.Sequential)] private struct FrameMargins { public int left,right,top,bottom; }
        [DllImport("dwmapi.dll")] private static extern int DwmExtendFrameIntoClientArea(IntPtr hwnd,ref FrameMargins margins);
        [DllImport("dwmapi.dll")] private static extern int DwmSetWindowAttribute(IntPtr hwnd,int attribute,ref int value,int size);
        private static void SetFrame(IntPtr hwnd, bool docked) { int policy=docked?1:0; DwmSetWindowAttribute(hwnd,2,ref policy,4); }
        private readonly Timer tracking=new Timer{Interval=120};
        private System.Drawing.Rectangle lastBounds;
        [DllImport("user32.dll")] private static extern int GetWindowLong(IntPtr window,int index);
        [DllImport("user32.dll")] private static extern int SetWindowLong(IntPtr window,int index,int value);
        [DllImport("user32.dll",EntryPoint="SetWindowLongPtrW")] private static extern IntPtr SetWindowLongPtr(IntPtr window,int index,IntPtr value);
        [DllImport("user32.dll")] private static extern bool SetWindowPos(IntPtr window,IntPtr after,int x,int y,int width,int height,uint flags);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
        [DllImport("user32.dll")] private static extern bool IsIconic(IntPtr window);
        [DllImport("user32.dll")] private static extern bool IsZoomed(IntPtr window);
        [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr window,int command);
        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr window,uint flags);
        private static void Owner(IntPtr window,IntPtr owner){if(IntPtr.Size==8)SetWindowLongPtr(window,-8,owner);else SetWindowLong(window,-8,owner.ToInt32());}
        internal IntPtr HostWindow => IsHandleCreated?GetAncestor(Handle,2):IntPtr.Zero;
        public AiPane()
        {
            AutoScaleMode=AutoScaleMode.Dpi;BackColor=System.Drawing.Color.White;
            tracking.Tick+=(s,e)=>Fit();
        }
        internal void Attach(IntPtr window)
        {
            if(child==window){Fit();return;}
            if(IsIconic(window)||IsZoomed(window))ShowWindow(window,9);
            child=window;originalStyle=GetWindowLong(child,-16);originalExtendedStyle=GetWindowLong(child,-20);
            GetWindowRect(child,out floatingBounds);
            // Keep Chromium as an owned top-level HWND. SetParent/WS_CHILD lets
            // Word's accelerator loop consume text before Chromium receives it.
            owner=GetAncestor(Handle,2);Owner(child,owner);
            SetWindowLong(child,-16,(originalStyle & ~0x00C40000 & ~0x40000000)|unchecked((int)0x80000000));
            SetWindowLong(child,-20,originalExtendedStyle & ~0x00000301);SetFrame(child,true);
            var margins=new FrameMargins();DwmExtendFrameIntoClientArea(child,ref margins);
            SetWindowPos(child,IntPtr.Zero,0,0,0,0,0x37); // Recalculate non-client metrics before measuring.
            lastBounds=System.Drawing.Rectangle.Empty;Fit();tracking.Start();
        }
        internal void Detach()
        {
            tracking.Stop();if(child==IntPtr.Zero)return;
            SetWindowRgn(child,IntPtr.Zero,true);
            Owner(child,IntPtr.Zero);SetWindowLong(child,-16,originalStyle);SetWindowLong(child,-20,originalExtendedStyle);SetFrame(child,false);
            var window=child;child=IntPtr.Zero;
            SetWindowPos(window,IntPtr.Zero,floatingBounds.left,floatingBounds.top,Math.Max(320,floatingBounds.right-floatingBounds.left),Math.Max(240,floatingBounds.bottom-floatingBounds.top),0x34);ShowWindow(window,4);
        }
        private void Fit()
        {
            if(child==IntPtr.Zero||!IsHandleCreated)return;
            if(!Visible||!IsWindowVisible(Handle)||IsIconic(owner)){ShowWindow(child,0);return;}
            var bounds=RectangleToScreen(ClientRectangle);
            if(bounds.Width<20||bounds.Height<20){ShowWindow(child,0);return;}
            if(bounds!=lastBounds){
                // Word 2010/Win10 can retain invisible resize/non-client margins
                // after Electron changes the frame. Align the CLIENT area, then clip
                // those margins; do not assume outer HWND bounds equal web content.
                GetWindowRect(child,out var outer);GetClientRect(child,out var client);
                var origin=new System.Drawing.Point();ClientToScreen(child,ref origin);
                int left=origin.X-outer.left,top=origin.Y-outer.top;
                int extraWidth=Math.Max(0,outer.right-outer.left-(client.right-client.left));
                int extraHeight=Math.Max(0,outer.bottom-outer.top-(client.bottom-client.top));
                SetWindowPos(child,IntPtr.Zero,bounds.X-left,bounds.Y-top,bounds.Width+extraWidth,bounds.Height+extraHeight,0x34);
                IntPtr region=CreateRectRgn(left,top,left+bounds.Width,top+bounds.Height);
                if(SetWindowRgn(child,region,true)==0)DeleteObject(region); // Windows owns successful region.
                lastBounds=bounds;
            }
            ShowWindow(child,4);
        }
        protected override void OnResize(EventArgs e){base.OnResize(e);Fit();}
        protected override void OnVisibleChanged(EventArgs e){base.OnVisibleChanged(e);Fit();}
        protected override void Dispose(bool disposing){if(disposing){Detach();tracking.Dispose();}base.Dispose(disposing);}
    }
}
