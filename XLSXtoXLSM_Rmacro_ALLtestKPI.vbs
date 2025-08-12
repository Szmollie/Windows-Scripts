'===========================================================
' This VBS turns XLSX into XLSM and Runs the Filtering Macro
'===========================================================
Dim args, objExcel, objWorkbook, todaydFileToReadS, actKPIpathTXT

Set todaydFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile("P:\ALLtestKPI\Scripts\todayd.txt", 1)
todayd = todaydFileToRead.Readline()
todaydFileToRead.Close

Set actKPIpathTXT = CreateObject("Scripting.FileSystemObject").OpenTextFile("P:\ALLtestKPI\Scripts\actKPIpath.txt", 1)
actKPIpath = actKPIpathTXT.Readline()
actKPIpathTXT.Close

Set args = WScript.Arguments
Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open(actKPIpath)


objExcel.Visible = True
objExcel.DisplayAlerts = False
objWorkbook.SaveAs "P:\ALLtestKPI\Database\ALLtestKPI_Rep_TP_DB.xlsm", 52

objExcel.Workbooks(1).Close
objExcel.Quit
'======================================




