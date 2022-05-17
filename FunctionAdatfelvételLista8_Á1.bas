Attribute VB_Name = "functionAdatfelvételLista8_Á1"
Option Explicit

Sub AdatfelvételLista8_Á1()
'JelszóRejtés2

Sheets("transfer_kulcsgép").Select

Columns("r:r").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.row
Dim ALoszlop As String
ALoszlop = "r"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

Range("a1", ALkoord).Select
    ActiveWorkbook.Worksheets("transfer_kulcsgép").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("transfer_kulcsgép").Sort.SortFields.Add Key:=Range("r1"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortOnValues
           
            With ActiveWorkbook.Worksheets("transfer_kulcsgép").Sort
        .SetRange Range("R2", ALkoord)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'Dim Állásidõ As Long
'Állásidõ = Application.WorksheetFunction.Sum(Range("r2", ALkoord))
'AppWindow.TextBox96.Value = "Állásidõ: " & Állásidõ & " Ft"


Dim rngList As Range
Set rngList = Munka11.Range("a1", ALkoord)
AppWindow.ListBox27.List = rngList.Value

Sheets("Start").Select
Range("b2").Select

'JelszóRejtés
End Sub
