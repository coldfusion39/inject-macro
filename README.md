# Inject-Macro
Inject VBA macro code into Excel documents

## Summary ##
Inject-Macro allows for injection of VBA macros into Microsoft Excel documents; specifically targeting the 97-2003 '.xls' Excel file format due to their ability to contain VBA macros without having a '.xlsm' file extension.

Inject-Macro requires an Excel file, specified with '-Excel', and a plain text VBA macro file, specified with '-Macro'. The macro will be injected into the Excel document and file metadata such as the last author will be removed. This is intended to be a quick way to prepare a templated Excel document with a macro payload.

If the '-Infect' flag is given, the supplied VBA macro will be injected into all Excel documents found in the specified '-Excel' directory path. Ideally this would be used to establish a low level form of persistence. The Excel 'Security' registry keys are disabled and not re-enabled on exit. This removes that pesky 'Macros have been disabled.' warning, and executes the macro without prompting the user. Additionally, the 'LastAccessTime', 'LastWriteTime' and 'Author' file properties of the Excel document are initially copied and replaced after injection to make the file appear untouched. For clean-up, the location of all injected Excel documents are written to 'excel_inject.log'. 

If the '-Clean' flag is given, the 'excel_inject.log' file must be in the same directory as Inject-Macro.ps1. The macros will be removed from the injected Excel documents and the registry keys will be re-enabled.

## Requirements ##
Excel and PowerShell 2.0 or greater are the only requirements for [Inject-Macro.ps1](https://github.com/coldfusion39/inject-macro/blob/master/Inject-Macro.ps1)

## Examples ##
Use [Inject-Macro.ps1](https://github.com/coldfusion39/inject-macro/blob/master/Inject-Macro.ps1) to inject a VBA macro into a single Excel document

`C:\PS> Inject-Macro -Excel C:\Users\Test\Excel.xls -Macro C:\temp\Macro.vba`

Use [Inject-Macro.ps1](https://github.com/coldfusion39/inject-macro/blob/master/Inject-Macro.ps1) to recursively search a directory for all '.xls' files and inject a VBA macro into each document

`C:\PS> Inject-Macro -Excel C:\Users\ -Macro C:\temp\Macro.vba -Infect`

Use [Inject-Macro.ps1](https://github.com/coldfusion39/inject-macro/blob/master/Inject-Macro.ps1) to remove the previously injected VBA macros from the targeted Excel documents

`C:\PS> Inject-Macro -Clean`

## Credits ##
Special Thanks:

 * Jeff McCutchan - jamcut ([@jamcut](https://twitter.com/jamcut))

 * Spencer McIntyre - zeroSteiner ([@zeroSteiner](https://twitter.com/zeroSteiner))
