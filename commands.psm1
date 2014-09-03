function ss {
	<#
	.SYNOPSIS
	List all Snapshots
	.DESCRIPTION
	List all Snapshots showing only vm, name, created
#>

	get-vm | get-snapshot | select vm,name,created
}

function  cc {
	<#
	.SYNOPSIS
	List all consolidations
	.DESCRIPTION
	List all consolidations
#>
	
	Get-VM | where {$_.ExtensionData.Runtime.consolidationNeeded} | Select Name
}