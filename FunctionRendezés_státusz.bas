Attribute VB_Name = "FunctionRendezés_státusz"
Option Explicit

Sub Rendezés()
' - státuszok - '

Munka12.Select
Columns("b:b").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.row
Dim ALoszlop As String
ALoszlop = "b"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw


Munka12.Range("b2", ALkoord).Select
    Munka12.Sort.SortFields.Clear
    Munka12.Sort.SortFields.Add Key:=Range("b2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    'ActiveWorkbook.Worksheets("Munka12").Sort.SortFields.Clear
    'ActiveWorkbook.Worksheets("Munka12").Sort.SortFields.Add Key:=Range("b1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With Munka12.Sort
        .SetRange Range("b2", ALkoord)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub
