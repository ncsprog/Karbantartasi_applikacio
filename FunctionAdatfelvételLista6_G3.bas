Attribute VB_Name = "FunctionAdatfelv�telLista6_G3"
Option Explicit

Sub Adatfelv�telLista6()
'Jelsz�Rejt�s2
' - K�ls� k�lts�g - '

Sheets("transfer_gazdas�gi").Select

Columns("p:p").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.Row
Dim ALoszlop As String
ALoszlop = "p"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

Range("a1", ALkoord).Select
    ActiveWorkbook.Worksheets("transfer_gazdas�gi").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("transfer_gazdas�gi").Sort.SortFields.Add Key:=Range("p1"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortOnValues
            
            With ActiveWorkbook.Worksheets("transfer_gazdas�gi").Sort
        .SetRange Range("A2", ALkoord)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Dim k�lts�g As Long
k�lts�g = Application.WorksheetFunction.Sum(Range("p2", ALkoord))
AppWindow.TextBox95.Value = "K�ls� k�lts�g: " & k�lts�g & " Ft"


Dim rngList As Range
Set rngList = Munka10.Range("a1", ALkoord)
AppWindow.ListBox25.List = rngList.Value

Sheets("Start").Select
Range("b2").Select

'Jelsz�Rejt�s
End Sub
