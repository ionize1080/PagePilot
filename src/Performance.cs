using System;
using System.Diagnostics;
namespace PagePilot {
 internal sealed class Performance : IDisposable {
  readonly string stage; readonly Stopwatch watch=Stopwatch.StartNew();
  internal Performance(string name){stage=name;}
  public void Dispose(){Connect.Log("PERF "+stage+" ms="+watch.Elapsed.TotalMilliseconds.ToString("F3",System.Globalization.CultureInfo.InvariantCulture));}
 }
}
