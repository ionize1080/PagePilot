param([string]$Root = (Split-Path -Parent $PSScriptRoot))
$ErrorActionPreference='Stop'
$dll=[IO.Path]::GetFullPath((Join-Path $Root 'PagePilot.dll'))
if (-not (Test-Path -LiteralPath $dll)) { throw 'PagePilot.dll is missing. Extract the complete portable package first.' }
$assembly=[Reflection.AssemblyName]::GetAssemblyName($dll).FullName
$codebase=([Uri]$dll).AbsoluteUri
$entries=@(
 @{Id='{0E057EBB-B35B-4A31-9B65-3DF44E8034CE}';Class='PagePilot.Connect'},
 @{Id='{79CEEFC8-67F4-4F55-9475-24C0681BFBBD}';Class='PagePilot.Pane'}
)
foreach($view in @([Microsoft.Win32.RegistryView]::Registry32,[Microsoft.Win32.RegistryView]::Registry64)) {
 $user=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser,$view)
 try {
  foreach($entry in $entries) {
   $base='Software\Classes\CLSID\'+$entry.Id
   $key=$user.CreateSubKey($base);$key.SetValue('',$entry.Class);$key.Dispose()
   $key=$user.CreateSubKey($base+'\InprocServer32')
   $key.SetValue('','mscoree.dll');$key.SetValue('ThreadingModel','Both');$key.SetValue('Class',$entry.Class)
   $key.SetValue('Assembly',$assembly);$key.SetValue('RuntimeVersion','v4.0.30319');$key.SetValue('CodeBase',$codebase);$key.Dispose()
   $key=$user.CreateSubKey($base+'\ProgId');$key.SetValue('',$entry.Class);$key.Dispose()
   $key=$user.CreateSubKey('Software\Classes\'+$entry.Class+'\CLSID');$key.SetValue('',$entry.Id);$key.Dispose()
   if($entry.Class -eq 'PagePilot.Pane') {
    $user.CreateSubKey($base+'\Control').Dispose()
    $user.CreateSubKey($base+'\Implemented Categories\{40FC6ED4-2438-11CF-A3DB-080036F12502}').Dispose()
   }
  }
  $key=$user.CreateSubKey('Software\Microsoft\Office\Word\Addins\PagePilot.Connect')
  $key.SetValue('FriendlyName','PagePilot');$key.SetValue('Description','AI Powered Word · Chatbox 1.20.3')
  $key.SetValue('LoadBehavior',3,[Microsoft.Win32.RegistryValueKind]::DWord);$key.Dispose()
 } finally {$user.Dispose()}
}
Write-Output 'PagePilot enabled for the current user. Restart Word. No installer or Uninstall entry was created.'
