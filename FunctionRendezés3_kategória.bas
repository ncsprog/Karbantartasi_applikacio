Attribute VB_Name = "FunctionRendez�s3_kateg�ria"
Option Explicit

Sub Rendez�s3()

Munka12.Select
Columns("j:j").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.Row
Dim ALoszlop As String
ALoszlop = "j"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw


Munka12.Range("j2", ALkoord).Select
    Munka12.Sort.SortFields.Clear
    Munka12.Sort.SortFields.Add Key:=Range("j2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With Munka12.Sort
        .SetRange Range("j2", ALkoord)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub
