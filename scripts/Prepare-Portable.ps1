param([string]$BuildPython='python',[string]$OfficePiaPath='C:\Program Files (x86)\Microsoft Visual Studio\Shared\Visual Studio Tools for Office\PIA\Office14\Office.dll')
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
$temp=Join-Path $repo '.temp'
$dist=Join-Path $repo 'dist'
New-Item -ItemType Directory -Force $temp,$dist | Out-Null
function Fetch-Pinned($Url,$Name,$Sha) {
 $file=Join-Path $temp $Name
 if(-not(Test-Path -LiteralPath $file)){Invoke-WebRequest $Url -OutFile $file}
 if((Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash -ne $Sha){throw "Checksum mismatch: $Name"}
 return $file
}
& (Join-Path $repo 'build.ps1') -OfficePiaPath $OfficePiaPath
$py=Fetch-Pinned 'https://www.python.org/ftp/python/3.12.10/python-3.12.10-embed-amd64.zip' 'python.zip' '4ACBED6DD1C744B0376E3B1CF57CE906F9DC9E95E68824584C8099A63025A3C3'
Expand-Archive -LiteralPath $py -DestinationPath (Join-Path $dist 'Python') -Force
& $BuildPython -m pip install --only-binary=:all: --target (Join-Path $dist 'Python\Lib\site-packages') -r (Join-Path $repo 'requirements-lock.txt')
if($LASTEXITCODE -ne 0){throw 'Python dependency preparation failed'}
@('python312.zip','.','Lib\site-packages','import site') | Set-Content (Join-Path $dist 'Python\python312._pth') -Encoding ASCII
$electron=Fetch-Pinned 'https://github.com/electron/electron/releases/download/v44.7.0/electron-v44.7.0-win32-x64.zip' 'electron.zip' 'EEE30DC8FA1F5EA95490E59F44E46EA68DD24C6E93D22FACF70FE5C2D4C2665C'
Expand-Archive -LiteralPath $electron -DestinationPath (Join-Path $dist 'Browser') -Force
$chatbox=Fetch-Pinned 'https://github.com/chatboxai/chatbox/releases/download/v1.20.3/Chatbox-1.20.3-Setup.exe' 'Chatbox-1.20.3-Setup.exe' '874171C6C7358FF3915F01412EB036154A5DED7A76C4663FB0F592D2C5C23B13'
& npm ci --prefix (Join-Path $repo 'chatbox') --ignore-scripts --no-audit --no-fund
if($LASTEXITCODE -ne 0){throw 'Chatbox build dependency preparation failed'}
& node (Join-Path $repo 'chatbox\prepare.cjs') $chatbox (Join-Path $dist 'Chatbox') (Join-Path $repo 'chatbox\node_modules')
if($LASTEXITCODE -ne 0){throw 'Chatbox extraction failed'}
foreach($name in @('bridge','vendor','scripts')) {Copy-Item (Join-Path $repo $name) $dist -Recurse -Force}
foreach($name in @('Enable.cmd','Disable.cmd','README.md','LICENSE','THIRD-PARTY-NOTICES.md')) {Copy-Item (Join-Path $repo $name) $dist -Force}
Write-Output 'Portable runtime prepared. No installer was executed.'
