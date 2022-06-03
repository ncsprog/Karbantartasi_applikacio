Attribute VB_Name = "FunctionAdatfelvételLista9_R1"
Option Explicit

Sub AdatfelvételLista9_R1()
' - Rendelkezés - '
Sheets("transfer_rendelkezés").Select

Columns("r:r").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.Row
Dim ALoszlop As String
ALoszlop = "r"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

Range("a1", ALkoord).Select
    ActiveWorkbook.Worksheets("transfer_rendelkezés").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("transfer_rendelkezés").Sort.SortFields.Add Key:=Range("r1"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortOnValues
           
            With ActiveWorkbook.Worksheets("transfer_rendelkezés").Sort
        .SetRange Range("R2", ALkoord)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Dim rngList As Range
Set rngList = Munka14.Range("a1", ALkoord)
AppWindow.ListBox31.List = rngList.Value

Sheets("Start").Select
Range("b2").Select
End Sub
