#!/usr/bin/env python
# Copyright (c) 2015, Brandan [coldfusion]
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.
import argparse
import os
import sys
from comtypes.client import CreateObject
from _winreg import *

parser = argparse.ArgumentParser(description='Inject VBA macro code into an Excel document')
parser.add_argument('-f', '--file', help='Path of the target Excel document', required=True)
parser.add_argument('-m', '--macro', help='Path of the VBA macro file you want to inject', required=True)
parser.add_argument('-o', '--outfile', help='Output file (instead of injecting into original Excel document)', required=False)
args = parser.parse_args()

sys.coinit_flags = 0

def main():
	# Check if running on windows
	if os.name != 'nt':
		print_warn('This script can only run on Windows!')
		sys.exit()

	if '.xls' in args.file:
		vba_trust('disable')
		try:
			inject_macro()
			vba_trust('enable')
		except:
			print_warn('Something went wrong!')
			vba_trust('enable')
	else:
		print_warn('File is not an .xls Excel document!')

def print_warn(msg):
	try:
		from colorama import init, Fore, Back, Style
		init(autoreset=True)
		print(Fore.RED + Style.BRIGHT + msg)
	except:
		print msg

# Disable macro security YES this is needed to create the macro, (it gets re-enabled)
def vba_trust(state):
	# Get installed Excel version number
	XLS = CreateObject("Excel.Application")
	excel_version = (XLS.version).encode('ascii', 'ignore')
	key_val = "Software\\Microsoft\\Office\\{version}\\Excel\\Security".format(version=excel_version)

	if state == 'disable':
		value = 1
	else:
		value = 0

	try:
		key = OpenKey(HKEY_CURRENT_USER, key_val, 0, KEY_ALL_ACCESS)
		SetValueEx(key, "AccessVBOM", 0, REG_DWORD, value)
		SetValueEx(key, "VBAWarnings", 0, REG_DWORD, value)
		CloseKey(key)
	except:
		print_warn('Could not open registry!')
		sys.exit()

	XLS.Quit()

# Inject macro into Excel document
def inject_macro():
	# Prevent Excel from popping up in the foreground
	XLS = CreateObject("Excel.Application", dynamic=True)
	XLS.DisplayAlerts = False
	XLS.Visible = False
	XLS.ScreenUpdating = False
	XLS.UserControl = False
	XLS.Interactive = False

	# Process Excel and macro file location
	excel_full_path = os.path.realpath(args.file)
	macro_full_path = os.path.realpath(args.macro)
	outfile_full_path = os.path.split(excel_full_path)

	workbook = XLS.Workbooks.Open(excel_full_path)
	VBA = workbook.VBProject.VBComponents.Add(1)	
	VBA.CodeModule.AddFromFile(macro_full_path)

	# Remove document metadata
#	xlRemoveDocType = "Microsoft.Office.Interop.Excel.XlRemoveDocInfoType"
#	workbook.RemoveDocumentInformation()

	# Save the document
	if args.outfile and '.' in args.outfile:
		file_base, ext_base = (args.outfile).split('.')
		outfile = file_base
		injected_document = "{outdir}\{outfile}.xls".format(outdir=outfile_full_path[0], outfile=outfile)
	else:
		injected_document = excel_full_path

	# Clean up before exiting
	workbook.SaveAs(injected_document)
	XLS.Workbooks.Close()
	XLS.Quit()

	# Makes output cleaner
	os.system('cls')
	print 'Macro was sucessfully injected!'
	print "File written to: {excel_out}".format(excel_out=injected_document)
	print_warn('Remember, the injected VBA macro is NOT password protected!')

if __name__ == '__main__':
	main()