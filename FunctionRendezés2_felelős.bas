Attribute VB_Name = "FunctionRendezés2_felelõs"
Option Explicit

Sub Rendezés2()
' - felelõsök - '

Munka12.Select
Columns("d:d").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.row
Dim ALoszlop As String
ALoszlop = "d"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw


Munka12.Range("d2", ALkoord).Select
    Munka12.Sort.SortFields.Clear
    Munka12.Sort.SortFields.Add Key:=Range("d2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
   With Munka12.Sort
        .SetRange Range("d2", ALkoord)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
