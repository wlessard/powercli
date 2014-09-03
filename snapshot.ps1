$v_VCServer = ''
$v_ADDomain = ''
$v_UserName = ''
$v_Password = ''
$mail_To = ''
$mail_From = ''
$mail_Subject = 'Daily Report'
$mail_Server = ''
$list = ''
$count = 0


#load VMware Snap
Add-PSSnapin -Name "VMware.VimAutomation.Core" -ErrorAction SilentlyContinue

#Connect to vCenter
Connect-VIServer -Server $v_VCServer -User $v_ADDomain\$v_UserName -Password $v_Password | Out-Null

#Remove all Snapshots older than 3 days
Get-VM | Get-Snapshot | Where {$_.Created -lt (Get-Date).AddDays(-3)} | Remove-Snapshot -Confirm:$false

#Remove all Snapshots created from Avamar
Get-VM | Get-Snapshot | Where {$_.Name -like "Avamar*"} | Remove-Snapshot -Confirm:$false

#Find all VM's needing consolidation and consolidate them
Get-VM | Where-Object {$_.Extensiondata.Runtime.ConsolidationNeeded} |
ForEach-Object {
  $count++
  $_.ExtensionData.ConsolidateVMDisks_Task()
  $list = $list + $_.name + "<br>"
}

#Create and Send email for daily report
$Body = "The script consolidated $count VM(s)<br>"
$Body = $Body + $list

send-mailmessage -To $mail_To -Subject $mail_Subject -Body $Body  -SmtpServer $mail_Server -From $mail_From -BodyAsHtml

Disconnect-viServer -Server $v_VCServer -Confirm:$false
