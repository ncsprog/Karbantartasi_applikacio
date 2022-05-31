Attribute VB_Name = "FunctionAdatfelvételLista6_G3"
Option Explicit

Sub AdatfelvételLista6()
'JelszóRejtés2
' - Külsõ költség - '

Sheets("transfer_gazdasági").Select

Columns("p:p").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.Row
Dim ALoszlop As String
ALoszlop = "p"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

Range("a1", ALkoord).Select
    ActiveWorkbook.Worksheets("transfer_gazdasági").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("transfer_gazdasági").Sort.SortFields.Add Key:=Range("p1"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortOnValues
            
            With ActiveWorkbook.Worksheets("transfer_gazdasági").Sort
        .SetRange Range("A2", ALkoord)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Dim költség As Long
költség = Application.WorksheetFunction.Sum(Range("p2", ALkoord))
AppWindow.TextBox95.Value = "Külsõ költség: " & költség & " Ft"


Dim rngList As Range
Set rngList = Munka10.Range("a1", ALkoord)
AppWindow.ListBox25.List = rngList.Value

Sheets("Start").Select
Range("b2").Select

'JelszóRejtés
End Sub
