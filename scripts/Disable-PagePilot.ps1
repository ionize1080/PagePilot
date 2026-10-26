param([string]$Root = (Split-Path -Parent $PSScriptRoot))
$ErrorActionPreference='Stop'
$manager=Join-Path ([IO.Path]::GetFullPath($Root)) 'PagePilot.Manager.exe'
if (-not (Test-Path -LiteralPath $manager)) { throw 'Extract the complete portable package first.' }
# Keep all deployment entry points on the same validated, logged transaction path.
$process=Start-Process -FilePath $manager -ArgumentList @('disable', '--root', ('"'+$Root.TrimEnd('\')+'"')) -PassThru -Wait
if ($process.ExitCode -ne 0) { throw 'PagePilot deployment did not complete. Open the manager for diagnostics.' }
