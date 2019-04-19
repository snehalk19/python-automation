'''Lots of data is present in an excel sheet and gets added regularly. It is time consuming to open every excel sheet and refresh the connections to get the updated data.
This python script reduces the work of refreshing the excel connections manually.
Following actions are performed by the script:
1.	The excel workbook is opened.
2.	The connection is refreshed.
3.	The workbook is saved.
4.	The workbook is closed.
To run the code you must have following things installed:
1.	python (This script is written in python version 3.7)
2.	win32 package
Disable the enable background refresh option in Data->Refresh All->Refresh Control to avoid the pop-up option.
'''


import win32com.client
import shutil

SourcePathName = 'D:/update'	
FileName = 'hi.xlsx'

Application = win32com.client.Dispatch("Excel.Application")

Application.Visible = 1

Workbook = Application.Workbooks.Open(SourcePathName + '/' + FileName)

Workbook.RefreshAll()

Workbook.Save()
 
Workbook.Close(True)

Application.Quit()
	
