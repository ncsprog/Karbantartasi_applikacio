Attribute VB_Name = "FunctionAdatfelv�telLista5_G2"
Option Explicit

Sub Adatfelf�telLista5()
'Jelsz�Rejt�s2
' - B�rk�lts�g - '

Sheets("transfer_gazdas�gi").Select

Columns("o:o").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.row
Dim ALoszlop As String
ALoszlop = "o"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

Range("a1", ALkoord).Select
    ActiveWorkbook.Worksheets("transfer_gazdas�gi").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("transfer_gazdas�gi").Sort.SortFields.Add Key:=Range("o1"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortOnValues
            
            With ActiveWorkbook.Worksheets("transfer_gazdas�gi").Sort
        .SetRange Range("A2", ALkoord)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Dim B�rk�lts�g As Long
B�rk�lts�g = Application.WorksheetFunction.Sum(Range("o2", ALkoord))
AppWindow.TextBox94.Value = "B�rk�lts�g: " & B�rk�lts�g & " Ft"


Dim rngList As Range
Set rngList = Munka10.Range("a1", ALkoord)
AppWindow.ListBox24.List = rngList.Value

Sheets("Start").Select
Range("b2").Select

'Jelsz�Rejt�s
End Sub
