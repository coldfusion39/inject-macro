# Inject-Macro
Inject VBA macro code into Excel documents

## Summary ##
Inject-Macro allows for injection of VBA macros into Microsoft Office Excel files; specifically targeting Excel 97-2003 '.xls' files due to the ability of these files to contain VBA macros without having a '.xlsm' file extension.

Inject-Macro has been implemented in PowerShell and Python running on Windows. Currently `inject-macro.py` only works on Windows systems running Python 2.7. This is because the Excel files are modified through the Windows COM interface using the [comtypes](https://github.com/enthought/comtypes/) Python library.

## Files ##
* [inject-macro.py](https://github.com/coldfusion39/inject-macro/blob/master/inject-macro.py): Python script to inject VBA macro code into Excel documents
* [Inject-Macro.ps1](https://github.com/coldfusion39/inject-macro/blob/master/examples/Inject-Macro.ps1): PowerShell script to inject VBA macro code into Excel documents

## Requirements ##
Excel and PowerShell 2.0 or greater are the only requirements for [Inject-Macro.ps1](https://github.com/coldfusion39/inject-macro/blob/master/examples/Inject-Macro.ps1)

Other than Microsoft Excel and Python 2.7, [inject-macro.py](https://github.com/coldfusion39/inject-macro/blob/master/inject-macro.py) has the following requirements:
* [comtypes](https://github.com/enthought/comtypes/)
* [colorama](https://github.com/tartley/colorama) (Optional)

## Setup ##
The following shows the __quickest__ way to install Python 2.7, pip, easy_install, and the above required dependencies:

* Download and install [Python 2.7](https://www.python.org/downloads/release/python-2710/), make sure to select the option to add 'python.exe' to your system path
* Download and install [easy_install](https://bootstrap.pypa.io/ez_setup.py): `C:\> python ez_setup.py`
* Download and install [pip](https://bootstrap.pypa.io/get-pip.py): `C:\> python get-pip.py`
* Install [comtypes](https://github.com/enthought/comtypes/) and [colorama](https://github.com/tartley/colorama) using pip: `C:\> pip install comtypes colorama`

## Examples ##
__Python__

Use [inject-macro.py](https://github.com/coldfusion39/inject-macro/blob/master/inject-macro.py) to inject the VBA macro 'macro_code' into 'Excel_01.xls'

`python inject-macro.py -f C:\Excel_01.xls -m C:\macro_code`

Use [inject-macro.py](https://github.com/coldfusion39/inject-macro/blob/master/inject-macro.py) to copy the Excel document 'Excel_01.xls' and inject the VBA macro 'macro_code' into the new document 'Excel_02.xls'

`python inject-macro.py -f C:\Excel_01.xls -m C:\macro_code -o Excel_02`

---

__PowerShell__

Use [Inject-Macro.ps1](https://github.com/coldfusion39/inject-macro/blob/master/examples/Inject-Macro.ps1) to inject the VBA macro 'macro_code' into 'Excel_01.xls'

`.\Inject-Macro.ps1 -Excel C:\Excel_01.xls -Macro C:\macro_code`

Use [Inject-Macro.ps1](https://github.com/coldfusion39/inject-macro/blob/master/examples/Inject-Macro.ps1) recursively search 'C:\Users\' for all '.xls' files and inject 'macro_code' into each document

`.\Inject-Macro.ps1 -Excel C:\Users\ -Macro C:\macro_code -Infect`

## Credits ##
Special Thanks:

 * Jeff McCutchan - jamcut ([@jamcut](https://twitter.com/jamcut))

 * Spencer McIntyre - zeroSteiner ([@zeroSteiner](https://twitter.com/zeroSteiner))
