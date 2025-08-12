'===========================================================
' This VBS Opens an Excel file in the Excel application and runs a macro
'===========================================================
On Error Resume Next

Dim wb_x00Path, MacroPath, MacroName, objExcel

Set myUserRead = CreateObject("Scripting.FileSystemObject").OpenTextFile("P:\UserInfo\myuser.txt", 1)
myuser = myUserRead.Readline()
myUserRead.Close

' Define the file path and macro details
wb_x00Path = "https://vitesco.sharepoint.com/teams/team_10035494/Shared Documents/General/FF2&FADQ/KPI tracking_2024/Daily FPY DQx00_testdev.xlsm"
MacroPath = "C:\Users\" & myuser & "\AppData\Roaming\Microsoft\Excel\XLSTART\PERSONAL.XLSB"
MacroName = "PERSONAL.XLSB!FiveDaysFPY"

' Create an instance of Excel
Set objExcel = CreateObject("Excel.Application")


' Open the main workbook
objExcel.Workbooks.Open wb_x00Path
If Err.Number <> 0 Then
    WScript.Echo "Error opening main workbook: " & Err.Description
    WScript.Quit 1
End If

' Turn off screen updating and automatic calculations
objExcel.ScreenUpdating = False
objExcel.Calculation = -4135 ' xlCalculationManual


' Make Excel visible
objExcel.Visible = True

' Open the workbook containing the macro
objExcel.Workbooks.Open MacroPath
If Err.Number <> 0 Then
    WScript.Echo "Error opening macro workbook: " & Err.Description
    WScript.Quit 1
End If

' Run the macro
objExcel.Run MacroName
If Err.Number <> 0 Then
    WScript.Echo "Error running macro: " & Err.Description
    WScript.Quit 1
End If

' Save and close the main workbook
objExcel.Workbooks(1).Save
objExcel.Workbooks(1).Close

' Open the main workbook
objExcel.Workbooks.Open wb_x00Path
If Err.Number <> 0 Then
    WScript.Echo "Error opening main workbook: " & Err.Description
    WScript.Quit 1
End If

' Save and close the macro workbook
objExcel.Workbooks(2).Save
'objExcel.Workbooks(2).Close

' Turn screen updating and automatic calculations back on
objExcel.ScreenUpdating = True
objExcel.Calculation = -4105 ' xlCalculationAutomatic

' Quit Excel
'objExcel.Quit
