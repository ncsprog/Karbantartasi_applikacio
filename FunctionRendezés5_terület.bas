Attribute VB_Name = "FunctionRendezés5_terület"
Option Explicit

Sub Rendezés5()

Munka12.Select
Columns("p:p").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.Row
Dim ALoszlop As String
ALoszlop = "p"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw


Munka12.Range("p2", ALkoord).Select
    Munka12.Sort.SortFields.Clear
    Munka12.Sort.SortFields.Add Key:=Range("p2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    'ActiveWorkbook.Worksheets("Munka12").Sort.SortFields.Clear
    'ActiveWorkbook.Worksheets("Munka12").Sort.SortFields.Add Key:=Range("b1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With Munka12.Sort
        .SetRange Range("p2", ALkoord)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With


End Sub
