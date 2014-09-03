#Script Settings
#-------------------------
$v_VCServer = ''
$v_ADDomain = ''
$v_UserName = ''
$v_Password = ''
$s_Location = ''
$s_cluster  = ''
$s_csvname = 'www.cvs'

#load VMware Snap
Add-PSSnapin -Name "VMware.VimAutomation.Core" -ErrorAction SilentlyContinue

Connect-VIServer -Server $v_VCServer -User $v_ADDomain\$v_UserName -Password $v_Password | Out-Null

#Get cluster and all host HBA information and change format from Binary to hex
$list = Get-cluster $s_cluster | Get-VMhost | Get-VMHostHBA -Type FibreChannel | Select VMHost,Device,@{N="WWN";E={"{0:X}" -f $_.PortWorldWideName}} | Sort VMhost,Device

#Go through each row and put : between every 2 digits
foreach ($item in $list){
   $item.wwn = (&{for ($i=0;$i -lt $item.wwn.length;$i+=2)   
                    {     
                        $item.wwn.substring($i,2)   
                    }}) -join':' 
}

#Output CSV to current directory.
$list | export-csv -NoTypeInformation $s_csvname