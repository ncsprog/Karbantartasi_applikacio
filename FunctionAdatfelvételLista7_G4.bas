Attribute VB_Name = "FunctionAdatfelv�telLista7_G4"
Option Explicit

Sub Adatfelv�telLista7()
' - Teljes K�lts�g - '

Sheets("transfer_gazdas�gi").Select

Columns("q:q").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.Row
Dim ALoszlop As String
ALoszlop = "q"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

Range("a1", ALkoord).Select
    ActiveWorkbook.Worksheets("transfer_gazdas�gi").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("transfer_gazdas�gi").Sort.SortFields.Add Key:=Range("q1"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortOnValues
            
            With ActiveWorkbook.Worksheets("transfer_gazdas�gi").Sort
        .SetRange Range("A2", ALkoord)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Dim Teljes_k�lts�g As Long
Teljes_k�lts�g = Application.WorksheetFunction.Sum(Range("q2", ALkoord))
AppWindow.TextBox96.Value = "Teljes k�lts�g: " & Teljes_k�lts�g & " Ft"


Dim rngList As Range
Set rngList = Munka10.Range("a1", ALkoord)
AppWindow.ListBox26.List = rngList.Value

Sheets("Start").Select
Range("b2").Select

End Sub
