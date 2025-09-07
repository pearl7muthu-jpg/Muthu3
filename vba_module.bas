
Attribute VB_Name = "ActivityLogger"
Option Explicit

Sub RecordActivity()
    Dim wsLog As Worksheet, wsData As Worksheet
    Dim emp As String, act As String
    Dim last As Long

    Set wsLog = ThisWorkbook.Sheets("Logger")
    Set wsData = ThisWorkbook.Sheets("Data")

    emp = Trim(wsLog.Range("B2").Value)
    act = Trim(wsLog.Range("B3").Value)

    If emp = "" Or act = "" Then
        MsgBox "Please select an Employee and an Action before recording.", vbExclamation, "Missing data"
        Exit Sub
    End If

    last = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row + 1
    wsData.Cells(last, 1).Value = emp
    wsData.Cells(last, 2).Value = act
    wsData.Cells(last, 3).Value = Now()

    wsLog.Range("B4").Value = wsData.Cells(last, 3).Value
    MsgBox "Recorded '" & act & "' for " & emp & " at " & Format(wsData.Cells(last, 3).Value, "yyyy-mm-dd hh:nn:ss"), vbInformation, "Recorded"
End Sub

' Optional helper: Adds named ranges for dropdowns (run once)
Sub CreateNamedRanges()
    Dim wb As Workbook
    Set wb = ThisWorkbook
    On Error Resume Next
    wb.Names.Add Name:="EmployeeList", RefersTo:="=Employees!$A$2:$A$" & ThisWorkbook.Sheets("Employees").Cells(Rows.Count,1).End(xlUp).Row
    wb.Names.Add Name:="ActionList", RefersTo:="=Actions!$A$2:$A$" & ThisWorkbook.Sheets("Actions").Cells(Rows.Count,1).End(xlUp).Row
    On Error GoTo 0
    MsgBox "Named ranges created: EmployeeList and ActionList", vbInformation
End Sub
