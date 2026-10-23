$ErrorActionPreference='Stop'
# Only fixed PagePilot-owned registration keys are removed; settings/documents remain.
$keys=@('Software\Microsoft\Office\Word\Addins\PagePilot.Connect',
 'Software\Classes\PagePilot.Connect','Software\Classes\PagePilot.Pane',
 'Software\Classes\CLSID\{0E057EBB-B35B-4A31-9B65-3DF44E8034CE}',
 'Software\Classes\CLSID\{79CEEFC8-67F4-4F55-9475-24C0681BFBBD}')
foreach($view in @([Microsoft.Win32.RegistryView]::Registry32,[Microsoft.Win32.RegistryView]::Registry64)) {
 $user=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser,$view)
 try {foreach($key in $keys){$user.DeleteSubKeyTree($key,$false)}}finally{$user.Dispose()}
}
Write-Output 'PagePilot disabled. Restart Word before moving or deleting the portable folder.'
