$v_VCServer = ''
$v_ADDomain = ''
$v_UserName = ''
$v_Password = ''

#load VMware Snap
Add-PSSnapin -Name "VMware.VimAutomation.Core" -ErrorAction SilentlyContinue
Connect-VIServer -Server $v_VCServer -User $v_ADDomain\$v_UserName -Password $v_Password | Out-Null

import-module c:\scripts\commands.psm1