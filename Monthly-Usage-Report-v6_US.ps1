

#Script Settings
#-------------------------

$Script_version = 'v6_US'

$v_VCServer = ''
$v_ADDomain = ''
$v_UserName = ''
$v_Password = ''
$s_Location = ''
$s_cluster = ''
$v_MailTo = ''
$v_MailCC = ''
#---------------------------

$Email_Body = ""

#load VMware Snap
Add-PSSnapin -Name "VMware.VimAutomation.Core" -ErrorAction SilentlyContinue
Connect-VIServer -Server $v_VCServer -User $v_ADDomain\$v_UserName -Password $v_Password | Out-Null

#
# Get Cluster Metrics
#
$i_total_hosts = 0
$i_total_vcpus = 0
$i_total_vcpus_ha = 0
$i_total_memoryMB = 0
$i_total_cpumhz = 0
$i_total_usage_cpumhz = 0
$i_total_hosts = 0
$i_total_cpumhz_ha = 0
$i_total_memorymb_ha = 0
$i_total_vms_memconsumed = 0
$i_total_vms_memoverhead = 0
$i_total_vms_memgranted  = 0

$Email_HostBody = "</table><br><br><table cellspacing='10'>"
$Email_HostBody += "<tr><th align=left>Host Name</th><th align=left>CPU Usage</th><th align=left>Memory Granted</th><th align=left>% Memory Commited</th><th align=left>VM Count</th></tr>"
$Email_HostBody += "<tr><th align=left>-----------------------------</th><th align=right>-----------------</th><th align=right>-------------------</th><th align=right>-------------------</th><th align=right>-----------------</th></tr>"

$g_vmhosts = get-vmhost -location $s_cluster

foreach ($s in $g_vmhosts)
{
	$i_total_hosts++
	$i_total_vcpus = $i_total_vcpus + $s.numcpu
	$i_total_cpumhz = $i_total_cpumhz + $s.cputotalmhz
	$i_total_usage_cpumhz = $i_total_usage_cpumhz + $s.cpuusagemhz
	$i_total_memorymb = $i_total_memorymb + $s.memorytotalmb
	$Consumed_stats = (get-stat -Entity $s -Stat mem.consumed.average -Realtime -MaxSamples 5 | Measure-Object -Property Value -Average).Average
	$Overhead_stats = (get-stat -Entity $s -Stat mem.overhead.average -Realtime -MaxSamples 5 | Measure-Object -Property Value -Average).Average
	$Granted_stats = (get-stat -Entity $s -Stat mem.granted.average -Realtime -MaxSamples 5 | Measure-Object -Property Value -Average).Average
	$TotalMem = [Math]::Round($s.memorytotalmb,2)
	$i_total_vms_memconsumed = $i_total_vms_memconsumed + $Consumed_stats;
	$i_total_vms_memoverhead = $i_total_vms_memoverhead + $Overhead_stats;
	$i_total_vms_memgranted  = $i_total_vms_memgranted + $Granted_stats;
	$VM2Host_Count = (Get-VM -Location $s).Count
	$Granted_statsGB = [Math]::Round(($Granted_stats/1024)/1024)
	$CpuPercent = [Math]::Round(($s.cpuusagemhz/$s.cputotalmhz)*100,2)
	$MemPercent = [Math]::Round((($Granted_stats/1024)/$s.memorytotalmb)*100,2)

	$Email_HostBody += "<tr><td>$s</td><td align=right>$CpuPercent %</td><td align=right>$Granted_statsGB</td><td align=right>$MemPercent %</td><td align=right>$VM2Host_Count</td></tr>" 
}

#Round up on Memory

$i_total_memorymb = [Math]::Round($i_total_memorymb,0)
$i_total_mem_com_overmb = [Math]::Round(($i_total_vms_memconsumed + $i_total_vms_memoverhead)/1024,0)
$i_total_vms_memgranted = [Math]::Round($i_total_vms_memgranted / 1024,0)


#Change Cpu and Memory to take Host Fail into consideration (N+1)
$i_total_vcpus_ha = $i_total_vcpus - $s.numcpu
$i_total_cpumhz_ha = $i_total_cpumhz - $s.cputotalmhz
$i_total_memorymb_ha =[Math]::Round($i_total_memorymb - $s.memorytotalmb,0)

#
# Get VM Metrics
#
$i_total_vms = 0
$i_total_vms_poweredon = 0
$i_total_vms_poweredoff = 0
$i_total_vms_vcpu = 0
$i_total_vms_grantedmb = 0


$g_vms = get-vm -location $s_cluster
foreach ($s in $g_vms)
{
	if ($s.powerstate -eq "poweredon")
	{
		$i_total_vms_poweredon++
		$i_total_vms_vcpu = $i_total_vms_vcpu + $s.numcpu
		$i_total_vms_grantedmb = $i_total_vms_grantedmb + $s.memorymb
	}
	else
	{
		$i_total_vms_poweredoff++
	}
}

$Email_BodyDB = "</table><br><br><table cellspacing='10'>"
$Email_BodyDB += "<tr><th align=left>Datastore Name</th><th align=left>Percent Used</th></tr>"
$Email_BodyDB += "<tr><th>--------------</th><th>------------</th></tr>"

if ((Get-DatastoreCluster).Count -ne 0) {

	Get-DatastoreCluster | Sort-Object Name | %{

		$DBName = $_.name
		$CapGB = [math]::Round($_.CapacityGB,0)
		$FreeGB = [math]::Round(($_.FreeSpaceGB),0)
		$PercUsed = [math]::Round(100 - (($FreeGB / $CapGB)*100),0)
		$Email_BodyDB += "<tr><td>$DBName</td><td align=right>$PercUsed %</td></tr>" 
	}

} else {

	$DSCapacitySum = [math]::Round((get-datastore |where { $_ -notlike "*scratch*"} | Measure-object -Property 'CapacityGB' -Sum | select -expand Sum)/1000,2)
	$DSFreeSum     = [math]::Round((get-datastore |where { $_ -notlike "*scratch*"} | Measure-object -Property 'FreespaceGB' -Sum | select -expand Sum)/1000,2)
	$DSUsedGB =  [math]::Round(($DSCapacitySum - $DSFreeSum),2)
	$DSPercUsed = [math]::Round(100*$DSUsedGB/$DSCapacitySum,2)
	$Email_BodyDB += "<tr><td></td><td align=right>$DSPercUsed %</td></tr>"

}
	
$i_total_vms = $i_total_vms_poweredon + $i_total_vms_poweredoff
$i_avg_cpu = $i_total_vms_vcpu / $i_total_vms_poweredon
$i_avg_mem = $i_total_mem_com_overmb / $i_total_vms_poweredon
$i_avg_grant_mem = $i_total_vms_grantedmb / $i_total_vms_poweredon
$i_memory_ratio = $i_total_vms_memgranted / $i_total_memorymb_ha
$i_memory_consumed_ratio = $i_total_mem_com_overmb / $i_total_memorymb_ha
$i_cpu_usage = $i_total_usage_cpumhz / $i_total_cpumhz_ha
$i_vcpu_ratio = $i_total_vms_vcpu / $i_total_vcpus_ha

$Email_Body += "<b>Analysis of $s_cluster cluster as of $(Get-Date)</b><br><br>" 
$Email_Body += "<Table>"
$Email_Body += "<tr><th align=left>Total hosts:</th><td align=right>$i_total_hosts</td></tr>"
$Email_Body += "<tr><th align=left>Total VMs:</th><td align=right>`{0:n0}</td></tr>" -f $i_total_vms
$Email_Body += "<tr><th align=left>Total VMs Powered On:</th><td align=right>`{0:n0}</td></tr>" -f $i_total_vms_poweredon
$Email_Body += "<tr><th align=left>Total Cluster CPU MHz:</th><td align=right>`{0:n0}</td></tr>" -f $i_total_cpumhz
$Email_Body += "<tr><th align=left>Total Cluster Memory (MB):</th><td align=right>`{0:n0}</td></tr>" -f $i_total_memorymb
$Email_Body += "<tr><th align=left>Total Cluster CPU MHz (HA):</th><td align=right>`{0:n0}</td></tr>" -f $i_total_cpumhz_ha
$Email_Body += "<tr><th align=left>Total Cluster Memory (MB)(HA):</th><td align=right>`{0:n0}</td></tr>" -f $i_total_memorymb_ha
$Email_Body += "<tr><th align=left>Total Consumed CPU MHz Usage:</th><td align=right>`{0:n0}</td></tr>" -f $i_total_usage_cpumhz
$Email_Body += "<tr><th align=left>Total Consumed Memory (MB):</th><td align=right>`{0:n0}</td></tr>" -f $i_total_mem_com_overmb
$Email_Body += "<tr><th align=left>Total Granted Memory (MB) :</th><td align=right>`{0:n0}</td></tr>" -f $i_total_vms_memgranted
$Email_Body += "<tr><th align=left>Total pCPU Granted:</th><td align=right>`{0:n0}</td></tr>" -f $i_total_vcpus
$Email_Body += "<tr><th align=left>Total vCPU Configured:</th><td align=right>`{0:n0}</td></tr>" -f $i_total_vms_vcpu
$Email_Body += "<tr><th align=left>AVG vCPU (Pwr On):</th><td align=right>`{0:n1}</td></tr>" -f $i_avg_cpu
$Email_Body += "<tr><th align=left>AVG Consumed Memory (MB)(Pwr On):</th><td align=right>`{0:n0}</td></tr>" -f $i_avg_mem
$Email_Body += "<tr><th align=left>AVG Granted Memory (MB) (Pwr On):</th><td align=right>`{0:n0}</td></tr>" -f $i_avg_grant_mem
$Email_Body += "<tr><th align=left>Percent Memory Commited (HA):</th><td align=right>`{0:P2}</td></tr>" -f $i_memory_ratio
$Email_Body += "<tr><th align=left>Percent Memory Consumed (HA):</th><td align=right>`{0:P2}</td></tr>" -f $i_memory_consumed_ratio
$Email_Body += "<tr><th align=left>Percent CPU Committed (HA):</th><td align=right>`{0:P0}</td></tr>" -f $i_cpu_usage
$Email_Body += "<tr><th align=left>vCPU : pCPU Ratio (HA):</th><td align=right>`{0:n2}:1</td></tr>" -f $i_vcpu_ratio
$Email_Body += $Email_BodyDB
$Email_Body += $Email_HostBody
$Email_Body += "</Table>"
$Email_Body += "<br><font color=red><b>Capgemini - BMC - ESX Capacity Report $script_version</b></font>"

$v_vCenterSettings = Get-View -Id 'OptionManager-VpxSettings'
$v_MailSender = ($v_vCenterSettings.Setting | Where-Object {$_.Key -eq "mail.sender"}).Value
$v_MailSmtpSrv = ($v_vCenterSettings.Setting | Where-Object {$_.Key -eq "mail.smtp.server"}).Value

Send-MailMessage -SmtpServer $v_MailSmtpSrv -From $v_MailSender -To $v_MailTo -CC $v_MailCC -Subject "BMC - ESX Cluster Monthly Usage Report - $s_Location - $(Get-Date -Format yyyy-MM-%d)" -BodyAsHtml $Email_Body
Disconnect-VIServer -Confirm:$False