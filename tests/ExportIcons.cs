using System;
using System.IO;
using System.Drawing.Imaging;
using PagePilot;
class ExportIcons {
 static void Main(string[] args){Directory.CreateDirectory(args[0]);for(int theme=0;theme<3;theme++)foreach(string id in new[]{"ppAI","ppRegex","ppNav","ppSettings"})using(var bitmap=RibbonIcons.Draw(id,64,theme))bitmap.Save(Path.Combine(args[0],theme+"-"+id+".png"),ImageFormat.Png);}
}
