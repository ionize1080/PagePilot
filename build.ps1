param([string]$OfficePiaPath='C:\Program Files (x86)\Microsoft Visual Studio\Shared\Visual Studio Tools for Office\PIA\Office14\Office.dll')
$ErrorActionPreference='Stop'
& dotnet build (Join-Path $PSScriptRoot 'src\PagePilot.csproj') -c Release "-p:OfficePiaPath=$OfficePiaPath"
if($LASTEXITCODE -ne 0){throw 'Build failed'}
New-Item -ItemType Directory -Force (Join-Path $PSScriptRoot 'dist') | Out-Null
Copy-Item (Join-Path $PSScriptRoot 'src\bin\Release\net48\PagePilot.dll') (Join-Path $PSScriptRoot 'dist')

& dotnet build (Join-Path $PSScriptRoot "manager\PagePilot.Manager.csproj") -c Release
if($LASTEXITCODE -ne 0){throw "Manager build failed"}
Copy-Item (Join-Path $PSScriptRoot "manager\bin\Release\net48\PagePilot.Manager.exe") (Join-Path $PSScriptRoot "dist")
