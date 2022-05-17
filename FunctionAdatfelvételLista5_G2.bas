Attribute VB_Name = "FunctionAdatfelvételLista5_G2"
Option Explicit

Sub AdatfelfételLista5()
'JelszóRejtés2
' - Bérköltség - '

Sheets("transfer_gazdasági").Select

Columns("o:o").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.row
Dim ALoszlop As String
ALoszlop = "o"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

Range("a1", ALkoord).Select
    ActiveWorkbook.Worksheets("transfer_gazdasági").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("transfer_gazdasági").Sort.SortFields.Add Key:=Range("o1"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortOnValues
            
            With ActiveWorkbook.Worksheets("transfer_gazdasági").Sort
        .SetRange Range("A2", ALkoord)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Dim Bérköltség As Long
Bérköltség = Application.WorksheetFunction.Sum(Range("o2", ALkoord))
AppWindow.TextBox94.Value = "Bérköltség: " & Bérköltség & " Ft"


Dim rngList As Range
Set rngList = Munka10.Range("a1", ALkoord)
AppWindow.ListBox24.List = rngList.Value

Sheets("Start").Select
Range("b2").Select

'JelszóRejtés
End Sub
