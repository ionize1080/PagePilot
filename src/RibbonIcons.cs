using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace PagePilot
{
    public static class RibbonIcons
    {
        private sealed class PictureConverter : System.Windows.Forms.AxHost
        { private PictureConverter():base(""){} public static object Convert(Bitmap image)=>GetIPictureDispFromPicture(image); }
        public static object Picture(string id,IntPtr window,int theme)=>PictureConverter.Convert(Get(id,window,theme));
        private static readonly Dictionary<string, Bitmap> cache = new Dictionary<string, Bitmap>();
        public static Bitmap Get(string id, IntPtr window, int theme=0)
        {
            int dpi = 96; try { uint value = GetDpiForWindow(window); if (value > 0) dpi = (int)value; } catch (EntryPointNotFoundException) { }
            int size = Math.Max(32, Math.Min(96, (int)Math.Round(32 * dpi / 96.0)));
            string key = id + ":" + size + ":" + theme;
            if (!cache.TryGetValue(key, out var bitmap)) cache[key] = bitmap = Draw(id, size, theme);
            return bitmap;
        }
        public static Bitmap Draw(string id, int size, int theme=0)
        {
            var image = new Bitmap(size*4, size*4, PixelFormat.Format32bppArgb); image.SetResolution(96, 96);
            using (var g = Graphics.FromImage(image))
            using (var pen = new Pen(theme==1?Color.FromArgb(43,67,112):theme==2?Color.FromArgb(49,56,67):Color.FromArgb(38,83,74), 1.65f) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
            using (var accent = new SolidBrush(theme==1?Color.FromArgb(70,113,214):theme==2?Color.FromArgb(81,94,111):Color.FromArgb(35,153,125)))
            using (var faint = new SolidBrush(theme==1?Color.FromArgb(211,223,249):theme==2?Color.FromArgb(216,222,228):Color.FromArgb(201,234,224)))
            {
                g.Clear(Color.Transparent); g.SmoothingMode = SmoothingMode.AntiAlias; g.PixelOffsetMode = PixelOffsetMode.HighQuality; g.ScaleTransform(size*4 / 32f, size*4 / 32f);
                if (id == "ppAI")
                {
                    g.FillRectangle(faint,5,3,20,26);g.DrawRectangle(pen,5,3,20,26);
                    g.DrawLine(pen,9,10,19,10);g.DrawLine(pen,9,15,16,15);
                    g.FillPolygon(accent,new[]{new Point(23,11),new Point(26,18),new Point(32,21),new Point(26,24),new Point(23,31),new Point(20,24),new Point(14,21),new Point(20,18)});

                }
                else if (id=="ppRegex") { g.FillEllipse(faint,3,3,20,20);g.DrawEllipse(pen,3,3,20,20);using(var p=new Pen(((SolidBrush)accent).Color,3){StartCap=LineCap.Round,EndCap=LineCap.Round})g.DrawLine(p,21,21,29,29);DrawLabel(g,".*",accent,new RectangleF(6,7,15,12),12); }
                else if (id == "ppSettings") {
                    var points=new List<PointF>();
                    for(int tooth=0;tooth<8;tooth++)for(int edge=0;edge<4;edge++){
                        double angle=tooth*Math.PI/4+(edge-1.5)*Math.PI/16;
                        float radius=(edge==1||edge==2)?13f:10f;
                        points.Add(new PointF(16+(float)Math.Cos(angle)*radius,16+(float)Math.Sin(angle)*radius));
                    }
                    using(var gear=new GraphicsPath()){gear.AddPolygon(points.ToArray());gear.AddEllipse(11,11,10,10);g.FillPath(faint,gear);g.DrawPath(pen,gear);}
                    using(var ring=new Pen(((SolidBrush)accent).Color,2f))g.DrawEllipse(ring,11,11,10,10);
                }
                else { g.DrawRectangle(pen, 5, 3, 21, 26); g.DrawLine(pen, 10, 3, 10, 29); g.FillPolygon(accent, new[] { new Point(17, 3), new Point(23, 3), new Point(23, 17), new Point(20, 14), new Point(17, 17) }); g.DrawLine(pen, 14, 22, 22, 22); }
            }
            var output=new Bitmap(size,size,PixelFormat.Format32bppArgb);
            using(var scaled=Graphics.FromImage(output)){scaled.InterpolationMode=InterpolationMode.HighQualityBicubic;scaled.PixelOffsetMode=PixelOffsetMode.HighQuality;scaled.DrawImage(image,new Rectangle(0,0,size,size));}
            image.Dispose();return output;
        }
        private static void DrawLabel(Graphics graphics,string text,Brush brush,RectangleF box,float size,string family="Segoe UI")
        {
            // Center the actual glyph outlines, not font side bearings or DrawString layout padding.
            using(var font=new FontFamily(family))using(var glyph=new GraphicsPath())using(var matrix=new Matrix())
            {
                glyph.AddString(text,font,(int)FontStyle.Bold,size,PointF.Empty,StringFormat.GenericTypographic);
                var bounds=glyph.GetBounds();float scale=Math.Min(1,Math.Min(box.Width/bounds.Width,box.Height/bounds.Height));
                matrix.Translate(-bounds.X,-bounds.Y);matrix.Scale(scale,scale,MatrixOrder.Append);
                matrix.Translate(box.X+(box.Width-bounds.Width*scale)/2,box.Y+(box.Height-bounds.Height*scale)/2,MatrixOrder.Append);glyph.Transform(matrix);graphics.FillPath(brush,glyph);
            }
        }
        public static void Release() { foreach (var image in cache.Values) image.Dispose(); cache.Clear(); }
        [DllImport("user32.dll")] private static extern uint GetDpiForWindow(IntPtr hwnd);
    }
}
