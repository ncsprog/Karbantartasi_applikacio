Attribute VB_Name = "FunctionAdatfelv�telLista9_R1"
Option Explicit

Sub Adatfelv�telLista9_R1()
'Jelsz�Rejt�s2
' - Rendelkez�s - '
Sheets("transfer_rendelkez�s").Select

Columns("r:r").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.row
Dim ALoszlop As String
ALoszlop = "r"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

Range("a1", ALkoord).Select
    ActiveWorkbook.Worksheets("transfer_rendelkez�s").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("transfer_rendelkez�s").Sort.SortFields.Add Key:=Range("r1"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortOnValues
           
            With ActiveWorkbook.Worksheets("transfer_rendelkez�s").Sort
        .SetRange Range("R2", ALkoord)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'Dim �ll�sid� As Long
'�ll�sid� = Application.WorksheetFunction.Sum(Range("r2", ALkoord))
'AppWindow.TextBox96.Value = "�ll�sid�: " & �ll�sid� & " Ft"


Dim rngList As Range
Set rngList = Munka14.Range("a1", ALkoord)
AppWindow.ListBox31.List = rngList.Value

Sheets("Start").Select
Range("b2").Select
'Jelsz�Rejt�s
End Sub
