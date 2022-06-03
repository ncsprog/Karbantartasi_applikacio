Attribute VB_Name = "FunctionAdatfelvételLista7_G4"
Option Explicit

Sub AdatfelvételLista7()
' - Teljes Költség - '

Sheets("transfer_gazdasági").Select

Columns("q:q").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.Row
Dim ALoszlop As String
ALoszlop = "q"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

Range("a1", ALkoord).Select
    ActiveWorkbook.Worksheets("transfer_gazdasági").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("transfer_gazdasági").Sort.SortFields.Add Key:=Range("q1"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortOnValues
            
            With ActiveWorkbook.Worksheets("transfer_gazdasági").Sort
        .SetRange Range("A2", ALkoord)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Dim Teljes_költség As Long
Teljes_költség = Application.WorksheetFunction.Sum(Range("q2", ALkoord))
AppWindow.TextBox96.Value = "Teljes költség: " & Teljes_költség & " Ft"


Dim rngList As Range
Set rngList = Munka10.Range("a1", ALkoord)
AppWindow.ListBox26.List = rngList.Value

Sheets("Start").Select
Range("b2").Select

End Sub
