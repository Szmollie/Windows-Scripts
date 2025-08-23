On Error Resume Next

Dim wb_x00Path, MacroPath, MacroName, objExcel, wbMain, wbMacro, myuser

' Read username from file
Set myUserRead = CreateObject("Scripting.FileSystemObject").OpenTextFile("P:\UserInfo\myuser.txt", 1)
myuser = myUserRead.Readline()
myUserRead.Close

' Define paths
wb_x00Path = "P:\PinSwap\ASB\Database\Template.xlsx"
MacroPath = "C:\Users\" & myuser & "\AppData\Roaming\Microsoft\Excel\XLSTART\PERSONAL.XLSB"
MacroName = "PERSONAL.XLSB!ASB_PinSwap_Evaluator"

' Create Excel instance
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
objExcel.ScreenUpdating = False
'objExcel.Calculation = -4135 ' xlCalculationManual

' Open main workbook
Set wbMain = objExcel.Workbooks.Open(wb_x00Path)
If wbMain Is Nothing Then
    WScript.Echo "Error opening main workbook: " & Err.Description
    objExcel.Quit
    WScript.Quit 1
End If

' Open macro workbook
Set wbMacro = objExcel.Workbooks.Open(MacroPath)
If wbMacro Is Nothing Then
    WScript.Echo "Error opening macro workbook: " & Err.Description
    wbMain.Close False
    objExcel.Quit
    WScript.Quit 1
End If

' Run macro
objExcel.Run MacroName
If Err.Number <> 0 Then
    WScript.Echo "Error running macro: " & Err.Description
    wbMacro.Close False
    wbMain.Close False
    objExcel.Quit
    WScript.Quit 1
End If

' Restore Excel settings
objExcel.ScreenUpdating = True
'objExcel.Calculation = -4105 ' xlCalculationAutomatic

' Save and close workbooks
wbMain.Close True
wbMacro.Close False

' Quit Excel
objExcel.Quit
