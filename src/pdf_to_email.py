import xlwings as xw
import datetime as dt
import os
import win32api
import subprocess

def pdf_email():
	mainbook = xw.Book.caller()
	active_sheet = mainbook.sheets.active

	#Save page as pdf. File saved to Documnets directory
	owner = active_sheet.range('A1').value
	today = dt.datetime.now().strftime("%d-%m-%Y")
	documents_dir = os.path.expanduser('~\Documents')
	filepath = os.path.join(documents_dir,f"Summary - {owner} {today}.pdf")
	active_sheet.api.ExportAsFixedFormat(0,filepath)

	#Open default email client
	win32api.ShellExecute(0,'open','mailto:',None,None ,0)

	#Open file location and highlight file
	subprocess.Popen(f'explorer /select,"{filepath}"')


