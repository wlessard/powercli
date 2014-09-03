###########################################################################
#
# TITLE: excel-ds-report.ps1
# PURPOSE: Generate an Excel DataStore Usage Report
# AUTHOR: crighter
# EMAIL: christopher.righter@capgemini.com
# SYNOPSIS: This scripts connects to vCenter in the Mrl/Phx
# 			Data Centers and generates an usage report for all BMC
#			Data Stores.
#
# VERSION HISTORY:
# 1.0 2/26/2013 - Initial release
# 2.0 4/15/2013 - Refactoring script for more flexibility
###########################################################################
# Global Variables

# Report Run Date
$v_Date = Get-Date -Format M.d.yyyy
# Define output directory
$v_RptDir = 'C:\Reports'

# Report Name
$v_BmcRptName = "phx_bmc_monthly_rpt_$v_Date.xlsx"

# Define decimal places
$v_Digits = 2

# Define AD credentials
$v_ADDomain = ''
$v_UserName = ''
$v_Password = ''

# Define vCenter Servers
$v_BmcVcSrv = 'localhost'

# Define Excel Sheetnames
$v_SheetName1 = 'PHX-DC-DS-RPT'
$v_SheetName2 = 'PHX-DC-VM-RPT'

###########################################################################
# Functions
Function f_LoadAutoSnapin
{
	param($v_PSSnapinName)
	if (!(Get-PSSnapin | where {$_.NAME -eq $v_PSSnapinName})) {
		Add-PSSnapin -Name $v_PSSnapinName
	}
}
Function f_GenerateReport
{
	f_ConnectVcntr
	
}

Function f_ConnectVcntr
{
	Connect-VIServer $v_BmcVcSrv -User $v_ADDomain\$v_UserName -Password $v_Password | Out-Null
	f_SetSMTPSrv
}

Function f_SetSMTPSrv
{
	$v_vCenterSettings = Get-View -Id 'OptionManager-VpxSettings'
	$v_MailSender = ($v_vCenterSettings.Setting | Where-Object {$_.Key -eq "mail.sender"}).Value
	$v_MailSmtpSrv = ($v_vCenterSettings.Setting | Where-Object {$_.Key -eq "mail.smtp.server"}).Value
	f_CollectDSMetrics
}

Function f_CollectDSMetrics
{
	$a_DSMetrics = Get-Datastore | where {$_.Name -notmatch "local"  -and $_.Name -notmatch "nas"} | Sort-Object
	$a_VmMetrics = Get-VM | Sort-Object
	f_CreateExcel
}

Function f_CreateExcel
{
	$v_Excel = new-object -comobject excel.application
 	$v_Excel.visible = $false
	$v_Excel.displayalerts = $false
	$v_Workbook = $v_Excel.workbooks.add()
 	$v_Workbook.Worksheets.Item(3).Delete()
 	$v_Workbook.WorkSheets.item(1).Name = $v_SheetName1
 	$v_Workbook.WorkSheets.item(2).Name = $v_SheetName2
	$v_Sheet1 = $v_Workbook.WorkSheets.Item($v_SheetName1)
 	$v_Sheet2 = $v_Workbook.WorkSheets.Item($v_SheetName2)
	
	$v_Row = 1
	$v_Column = 1
 	$v_Sheet1.Cells.Item($v_Row,$v_Column) = 'DSName'
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.Bold = $true
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.ColorIndex = 2
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Interior.ColorIndex = 23
	$v_Column++
	$v_Sheet1.Cells.Item($v_Row,$v_Column) = 'CapacityGB'
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.Bold = $true
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.ColorIndex = 2
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Interior.ColorIndex = 23
	$v_Column++
	$v_Sheet1.Cells.Item($v_Row,$v_Column) = 'ConsumedSpaceGB'
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.Bold = $true
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.ColorIndex = 2
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Interior.ColorIndex = 23
	$v_Column++
	$v_Sheet1.Cells.Item($v_Row,$v_Column) = 'FreeSpaceGB'
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.Bold = $true
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.ColorIndex = 2
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Interior.ColorIndex = 23
	$v_Column++
	$v_Sheet1.Cells.Item($v_Row,$v_Column) = 'PercentFree'
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.Bold = $true
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.ColorIndex = 2
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Interior.ColorIndex = 23	
	
	$v_Column = 1
	$v_Sheet2.Cells.Item($v_Row,$v_Column) = 'VirtualMachine'
	$v_Sheet2.Cells.Item($v_Row,$v_Column).Font.Bold = $true
	$v_Sheet2.Cells.Item($v_Row,$v_Column).Font.ColorIndex = 2
	$v_Sheet2.Cells.Item($v_Row,$v_Column).Interior.ColorIndex = 23
	$v_Column++
	$v_Sheet2.Cells.Item($v_Row,$v_Column) = 'ProvisionedSpaceGB'
	$v_Sheet2.Cells.Item($v_Row,$v_Column).Font.Bold = $true
	$v_Sheet2.Cells.Item($v_Row,$v_Column).Font.ColorIndex = 2
	$v_Sheet2.Cells.Item($v_Row,$v_Column).Interior.ColorIndex = 23
	$v_Column++
	$v_Sheet2.Cells.Item($v_Row,$v_Column) = 'ConsumedSpaceGB'
	$v_Sheet2.Cells.Item($v_Row,$v_Column).Font.Bold = $true
	$v_Sheet2.Cells.Item($v_Row,$v_Column).Font.ColorIndex = 2
	$v_Sheet2.Cells.Item($v_Row,$v_Column).Interior.ColorIndex = 23
	$v_Column++
	$v_Sheet2.Cells.Item($v_Row,$v_Column) = 'vSphereHost'
	$v_Sheet2.Cells.Item($v_Row,$v_Column).Font.Bold = $true
	$v_Sheet2.Cells.Item($v_Row,$v_Column).Font.ColorIndex = 2
	$v_Sheet2.Cells.Item($v_Row,$v_Column).Interior.ColorIndex = 23
	f_GatherDSMetrics
}
	
Function f_GatherDSMetrics
{
	$a_DSMetrics | where {$_.Name -match "iso"} | ForEach-Object{
	$v_Row = 2
	$v_Column = 1
	$v_Sheet1.Cells.Item($v_Row,$v_Column) = $_.Name
	$v_Column++
	# Capacity (GB)
	$v_Sheet1.Cells.Item($v_Row,$v_Column) = (($_.CapacityMB /1024),$v_Digits)
	$v_Column++
	# Consumed Space (GB)
	$v_Sheet1.Cells.Item($v_Row,$v_Column) = [Math]::Round(($_.CapacityMB /1024 - $_.FreeSpaceMB /1024),$v_Digits)
	$v_Column++
	# FreeSpace (GB)
	$v_Sheet1.Cells.Item($v_Row,$v_Column) = [Math]::Round($_.FreeSpaceMB /1024,$v_Digits)
	$v_Column++
	# Percent Free
	$v_Sheet1.Cells.Item($v_Row,$v_Column) = [Math]::Round(100*$_.FreeSpaceMB/$_.CapacityMB,$v_Digits)
	$v_Column++
	}
	# Increment to next row and reset column
	$v_Row++
	$v_Column=1
	$v_Sheet1.Cells.Item($v_Row,$v_Column) = "PHX ISO DS Total"
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.Bold = $true
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.ColorIndex = 2
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Interior.ColorIndex = 23
	$v_Column++
		
	$v_UsedRange = $v_Sheet1.UsedRange
	$v_ISOMaxRows = $v_UsedRange.rows.count
	$v_ISODSSumRow = $v_ISOMaxRows
	$v_ISODSCPSum = $v_Sheet1.Range("B2:B$v_ISOMaxRows")
	$v_ISODSCSSum = $v_Sheet1.Range("C2:C$v_ISOMaxRows")
	$v_ISODSFSSum = $v_Sheet1.Range("D2:D$v_ISOMaxRows")
	$v_ISODSPFAvg = $v_Sheet1.Range("E2:E$v_ISOMaxRows")
	$v_functions = $v_excel.WorksheetFunction
	$v_Sheet1.Cells.Item($v_ISODSSumRow,$v_Column) = $v_functions.Sum($v_ISODSCPSum)
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.Bold = $true
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.ColorIndex = 2
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Interior.ColorIndex = 23
	$v_Column++
	$v_Sheet1.cells.item($v_ISODSSumRow,$v_Column) = $v_functions.Sum($v_ISODSCSSum)
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.Bold = $true
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.ColorIndex = 2
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Interior.ColorIndex = 23
	$v_Column++
	$v_Sheet1.cells.item($v_ISODSSumRow,$v_Column) = $v_functions.Sum($v_ISODSFSSum)
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.Bold = $true
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.ColorIndex = 2
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Interior.ColorIndex = 23
	$v_Column++
	$v_Sheet1.cells.item($v_ISODSSumRow,$v_Column) = [Math]::Round($v_functions.Average($v_ISODSPFAvg),$v_Digits)
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.Bold = $true
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.ColorIndex = 2
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Interior.ColorIndex = 23
	$v_Row++
	$v_Column = 1
	f_GatherMGTMetrics
}
	
Function f_GatherMGTMetrics
{
	$a_DSMetrics | where {$_.Name -match "mgmt"} | ForEach-Object{
	$v_Sheet1.Cells.Item($v_Row,$v_Column) = $_.Name
	$v_Column++
	# Capacity (GB)
	$v_Sheet1.Cells.Item($v_Row,$v_Column) = (($_.CapacityMB /1024),$v_Digits)
	$v_Column++
	# Consumed Space (GB)
	$v_Sheet1.Cells.Item($v_Row,$v_Column) = [Math]::Round(($_.CapacityMB /1024 - $_.FreeSpaceMB /1024),$v_Digits)
	$v_Column++
	# FreeSpace (GB)
	$v_Sheet1.Cells.Item($v_Row,$v_Column) = [Math]::Round($_.FreeSpaceMB /1024,$v_Digits)
	$v_Column++
	# Percent Free
	$v_Sheet1.Cells.Item($v_Row,$v_Column) = [Math]::Round(100*$_.FreeSpaceMB/$_.CapacityMB,$v_Digits)
	$v_Column++
	# Increment to next row and reset column
	$v_Row++
	$v_Column=1
	}
	$v_Sheet1.Cells.Item($v_Row,$v_Column) = "PHX MGT DS Total"
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.Bold = $true
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.ColorIndex = 2
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Interior.ColorIndex = 23
	$v_Column++
	
	$v_UsedRange = $v_Sheet1.UsedRange
	$v_MGTMaxRows = $v_UsedRange.rows.count
	$v_MGTDSSumRow = $v_MGTMaxRows
	$v_MGTDSumRowStart = $v_ISOMaxRows + 1
	
	$v_MGTDSCPSum = $v_Sheet1.Range("B$v_MGTDSumRowStart`:B$v_MGTMaxRows")
	$v_MGTDSCSSum = $v_Sheet1.Range("C$v_MGTDSumRowStart`:C$v_MGTMaxRows")
	$v_MGTDSFSSum = $v_Sheet1.Range("D$v_MGTDSumRowStart`:D$v_MGTMaxRows")
	$v_MGTDSPFAvg = $v_Sheet1.Range("E$v_MGTDSumRowStart`:E$v_MGTMaxRows")
	$v_functions = $v_excel.WorksheetFunction
	$v_Sheet1.Cells.Item($v_MGTDSSumRow,$v_Column) = $v_functions.Sum($v_MGTDSCPSum)
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.Bold = $true
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.ColorIndex = 2
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Interior.ColorIndex = 23
	$v_Column++
	$v_Sheet1.cells.item($v_MGTDSSumRow,$v_Column) = $v_functions.Sum($v_MGTDSCSSum)
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.Bold = $true
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.ColorIndex = 2
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Interior.ColorIndex = 23
	$v_Column++
	$v_Sheet1.cells.item($v_MGTDSSumRow,$v_Column) = $v_functions.Sum($v_MGTDSFSSum)
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.Bold = $true
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.ColorIndex = 2
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Interior.ColorIndex = 23
	$v_Column++
	$v_Sheet1.cells.item($v_MGTDSSumRow,$v_Column) = [Math]::Round($v_functions.Average($v_MGTDSPFAvg),$v_Digit)
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.Bold = $true
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.ColorIndex = 2
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Interior.ColorIndex = 23
	$v_Row++
	$v_Column = 1	
	f_Gather_Prod_Metrics
}

Function f_Gather_Prod_Metrics
{
	$a_DSMetrics | where {$_.Name -match "cust"} | ForEach-Object{
	$v_Sheet1.Cells.Item($v_Row,$v_Column) = $_.Name
	$v_Column++
	# Capacity (GB)
	$v_Sheet1.Cells.Item($v_Row,$v_Column) = (($_.CapacityMB /1024),$v_Digits)
	$v_Column++
	# Consumed Space (GB)
	$v_Sheet1.Cells.Item($v_Row,$v_Column) = [Math]::Round(($_.CapacityMB /1024 - $_.FreeSpaceMB /1024),$v_Digits)
	$v_Column++
	# FreeSpace (GB)
	$v_Sheet1.Cells.Item($v_Row,$v_Column) = [Math]::Round($_.FreeSpaceMB /1024,$v_Digits)
	# Percent Free
	$v_Column++
	$v_Sheet1.Cells.Item($v_Row,$v_Column) = [Math]::Round(100*$_.FreeSpaceMB/$_.CapacityMB,$v_Digits)
	$v_Column++
	$v_Row++
	$v_Column=1
	}
	# Increment to next row and reset column
	$v_Sheet1.Cells.Item($v_Row,$v_Column) = "PHX Prod DS Total"
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.Bold = $true
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.ColorIndex = 2
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Interior.ColorIndex = 23
	$v_Column++
	
	$v_UsedRange = $v_Sheet1.UsedRange
	$v_ProdMaxRows = $v_UsedRange.rows.count
	$v_ProdDSSumRow = $v_ProdMaxRows
	$v_ProdDSumRowStart = $v_MGTMaxRows + 1
	
	$v_ProdDSCPSum = $v_Sheet1.Range("B$v_ProdDSumRowStart`:B$v_ProdMaxRows")
	$v_ProdDSCSSum = $v_Sheet1.Range("C$v_ProdDSumRowStart`:C$v_ProdMaxRows")
	$v_ProdDSFSSum = $v_Sheet1.Range("D$v_ProdDSumRowStart`:D$v_ProdMaxRows")
	$v_ProdDSPFAvg = $v_Sheet1.Range("E$v_ProdDSumRowStart`:E$v_ProdMaxRows")
	$v_functions = $v_excel.WorksheetFunction
	$v_Sheet1.Cells.Item($v_ProdDSSumRow,$v_Column) = $v_functions.Sum($v_ProdDSCPSum)
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.Bold = $true
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.ColorIndex = 2
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Interior.ColorIndex = 23
	$v_Column++
	$v_Sheet1.cells.item($v_ProdDSSumRow,$v_Column) = $v_functions.Sum($v_ProdDSCSSum)
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.Bold = $true
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.ColorIndex = 2
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Interior.ColorIndex = 23
	$v_Column++
	$v_Sheet1.cells.item($v_ProdDSSumRow,$v_Column) = $v_functions.Sum($v_ProdDSFSSum)
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.Bold = $true
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.ColorIndex = 2
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Interior.ColorIndex = 23
	$v_Column++
	$v_Sheet1.cells.item($v_ProdDSSumRow,$v_Column) = [Math]::Round($v_functions.Average($v_ProdDSPFAvg),$v_Digits)
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.Bold = $true
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Font.ColorIndex = 2
	$v_Sheet1.Cells.Item($v_Row,$v_Column).Interior.ColorIndex = 23
	$v_Row++
	$v_Column = 1
	f_GatherVmMetrics
}	
Function f_GatherVmMetrics
{
	$v_Row = 2
	$v_Column = 1
	$a_VmMetrics |  ForEach {
	# Datastore Name
	$v_Sheet2.Cells.Item($v_Row,$v_Column) = $_.Name
	$v_Column++
	# Provisioned Space (GB)
	$v_Sheet2.Cells.Item($v_Row,$v_Column) = [Math]::Round(($_.ProvisionedSpaceGB),$v_Digits)
	$v_Column++
	# Used Space (GB)
	$v_Sheet2.Cells.Item($v_Row,$v_Column) = [Math]::Round(($_.UsedSpaceGB),$v_Digits)
	$v_Column++
	# VMHost (GB)
	$v_Sheet2.Cells.Item($v_Row,$v_Column) = $_.VMHost.Name
	$v_Column++
	# Increment to next row and reset column
	$v_Row++
	$v_Column=1
	}
    f_ConditionalFormat
}

Function f_ConditionalFormat
{
	$v_UsedRange = $v_Sheet1.UsedRange
	$v_UsedRange.EntireColumn.Autofit() | Out-Null
	$v_UsedRange.Borders.LineStyle = 1 
	$v_UsedRange.Borders.Weight = 2
	
	$v_CFMaxRows = $v_UsedRange.Rows.Count
	$v_Selection = $v_Sheet1.Range("E1`:E$v_CFMaxRows")
	$v_Selection
	$v_Selection.FormatConditions.AddIconSetCondition() | Out-Null
	$v_Selection.FormatConditions.Item($($v_Selection.FormatConditions.Count)).SetFirstPriority()
	$v_Selection.FormatConditions.Item(1).ReverseOrder=$false
	$v_Selection.FormatConditions.Item(1).ShowIconOnly=$false
	$v_Selection.FormatConditions.Item(1).IconSet=$xlIconSet::xl3TrafficLights1
	$v_Selection.FormatConditions.Item(1).IconCriteria.Item(2).Type=$xlConditionValues::xlConditionValueNumber
	$v_Selection.FormatConditions.Item(1).IconCriteria.Item(2).Value=15
	$v_Selection.FormatConditions.Item(1).IconCriteria.Item(2).Operator=7
	$v_Selection.FormatConditions.Item(1).IconCriteria.Item(3).Type=$xlConditionValues::xlConditionValueNumber
	$v_Selection.FormatConditions.Item(1).IconCriteria.Item(3).Value=30
	$v_Selection.FormatConditions.Item(1).IconCriteria.Item(3).Operator=5

	$v_UsedRange = $v_Sheet2.UsedRange
	$v_UsedRange.EntireColumn.Autofit() | Out-Null
	$v_UsedRange.Borders.LineStyle = 1 
	$v_UsedRange.Borders.Weight = 2
	f_SaveReport
}

Function f_SaveReport
{
	# Set Excel SaveAs File Format
	$v_FileFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault

	# Save Excel Workbook
	$v_Workbook.SaveAs(“$v_RptDir\$v_BmcRptName",$v_FileFormat)

	$v_Workbook.Close()
	$v_Excel.Quit()
	[System.Runtime.InteropServices.Marshal]::ReleaseComObject($v_Excel)
	f_EmailReport
}
Function f_EmailReport
{
	Send-MailMessage -From $v_MailSender -To "bmcadmins.nar@capgemini.com" `
	-Subject "PHX BMC DS/VM Report for $(Get-Date -Format M)" `
	-Body "Automated Monthly Report: PHX BMC Datastores Usage and Virtual Machines Count" `
	-Attachments "$v_RptDir\$v_BmcRptName" -SmtpServer $v_MailSmtpSrv
}
f_LoadAutoSnapin -v_PSSnapinName "VMware.VimAutomation.Core"

f_GenerateReport