Attribute VB_Name = "FunctionAdatfelv�telLista4_G1"
Option Explicit

Sub Adatfelv�telLista4()
' - Anyagk�lts�g - '

Sheets("transfer_gazdas�gi").Select

Columns("n:n").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.Row
Dim ALoszlop As String
ALoszlop = "n"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

Range("a1", ALkoord).Select
    ActiveWorkbook.Worksheets("transfer_gazdas�gi").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("transfer_gazdas�gi").Sort.SortFields.Add Key:=Range("n1"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortOnValues
            
            With ActiveWorkbook.Worksheets("transfer_gazdas�gi").Sort
        .SetRange Range("A2", ALkoord)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    

Dim Aktsg As Long
Aktsg = Application.WorksheetFunction.Sum(Range("n2", ALkoord))
AppWindow.TextBox93.Value = "Anyagk�lts�g: " & Aktsg & " Ft"

Dim rngList As Range
Set rngList = Munka10.Range("a1", ALkoord)
AppWindow.ListBox23.List = rngList.Value


Sheets("Start").Select
Range("b2").Select

End Sub
