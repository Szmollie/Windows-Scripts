'===========================================================
' This VBS turns XLSX into XLSM and Runs the Filtering Macro
'===========================================================
Dim args, objExcel, objWorkbook, todaydFileToReadS, actKPIpathTXT

Set todaydFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile("P:\5DaysFPY\Scripts\todayd.txt", 1)
todayd = todaydFileToRead.Readline()
todaydFileToRead.Close

Set actKPIpathTXT = CreateObject("Scripting.FileSystemObject").OpenTextFile("P:\5DaysFPY\Scripts\actKPIpath.txt", 1)
actKPIpath = actKPIpathTXT.Readline()
actKPIpathTXT.Close

Set args = WScript.Arguments
Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open(actKPIpath)


objExcel.Visible = True
objExcel.DisplayAlerts = False
objWorkbook.SaveAs "P:\5DaysFPY\Database\DQ400_Rep_TP_DB.xlsm", 52

objExcel.ActiveWorkbook.Close(0)
objExcel.Quit
'======================================




