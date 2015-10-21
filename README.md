# Inject-Macro
Inject VBA macro code into Excel and Word documents

## Summary ##
Inject-Macro allows for the injection of VBA macros into Microsoft Excel and Word documents; specifically targeting the 97-2003 '.xls' and '.doc' file format due to their ability to contain VBA macros without having a '.xlsm' or '.docm' file extension.

Inject-Macro requires an Excel or Word file, specified with '-Doc', and a plain text VBA macro file, specified with '-Macro'. The macro will be injected into the document and file metadata such as the last author will be removed. This is intended to be a quick way to prepare a templated Excel or Word document with a macro payload.

#### -Infect ####
If the '-Infect' flag is given, the supplied VBA macro will be injected into all Excel or Word documents found in the user specified '-Doc' directory path. Inject-Macro will read the first line of the user supplied macro and look for 'Sub Auto_Open()' or 'Sub AutoOpen()'. Excel uses 'Sub Auto_Open()' to automatically run macro code when the documet is opened; Word uses 'Sub AutoOpen()'. This will determine if the macro will be injected into Excel or Word documents.

The VBA 'Security' registry keys are disabled and not re-enabled on exit when using the '-Infect' flag. This removes that pesky 'Macros have been disabled.' warning, and executes the macro without prompting the user.

Additionally, the 'LastAccessTime', 'LastWriteTime' and 'Author' file properties of the document are initially copied and replaced after injection to make the file appear untouched. Ideally this would be used to establish a low level form of persistence.

For clean-up, the location of all injected documents are written to '$env:temp\inject.log' when running Inject-Macro with the '-Infect' flag.

#### -Clean ####
If the '-Clean' flag is given, the VBA macro code will be removed from the documents and the registry keys will be re-enabled.

## Requirements ##
Excel and/or Word and PowerShell 2.0 or greater are the only requirements for Inject-Macro

## Examples ##
Inject the VBA macro 'Excel_Macro.vba' into the Excel document 'Excel.xls'

`C:\PS> Inject-Macro -Doc C:\Users\Test\Excel.xls -Macro C:\temp\Excel_Macro.vba`

Inject the VBA macro 'Word_Macro.vba' into the Word document 'Word.doc'

`C:\PS> Inject-Macro -Doc C:\Users\Test\Word.doc -Macro C:\temp\Word_Macro.vba`

Inject the VBA macro 'Macro.vba' into all Excel or Word documents found in 'C:\Users\' recursively.

`C:\PS> Inject-Macro -Doc C:\Users\ -Macro C:\temp\Macro.vba -Infect`

Remove the injected VBA macro code from all documents found in 'inject.log'.

`C:\PS> Inject-Macro -Clean`

## Credits ##
Special Thanks:
 * Jeff McCutchan - jamcut ([@jamcut](https://twitter.com/jamcut))
 * Spencer McIntyre - zeroSteiner ([@zeroSteiner](https://twitter.com/zeroSteiner))
