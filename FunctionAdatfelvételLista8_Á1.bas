Attribute VB_Name = "functionAdatfelv�telLista8_�1"
Option Explicit

Sub Adatfelv�telLista8_�1()
'Jelsz�Rejt�s2

Sheets("transfer_kulcsg�p").Select

Columns("r:r").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.row
Dim ALoszlop As String
ALoszlop = "r"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

Range("a1", ALkoord).Select
    ActiveWorkbook.Worksheets("transfer_kulcsg�p").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("transfer_kulcsg�p").Sort.SortFields.Add Key:=Range("r1"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortOnValues
           
            With ActiveWorkbook.Worksheets("transfer_kulcsg�p").Sort
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
Set rngList = Munka11.Range("a1", ALkoord)
AppWindow.ListBox27.List = rngList.Value

Sheets("Start").Select
Range("b2").Select

'Jelsz�Rejt�s
End Sub
