Attribute VB_Name = "FunctionAdatfelvételLista13"
Option Explicit

Sub AdatfelvételLista13()
' - Applikációból származó adatok - '

' - kiürítit a transfert - '
Sheets("szûrõ_transfer").Select
Range("a1:xx10000") = ""
' - átmásolja a kezelendõ adatokat - '
Sheets("adatok").Select
Columns("w:w").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.Row
Dim ALoszlop As String
ALoszlop = "w"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw
Range("a1", ALkoord).Copy
Sheets("szûrõ_transfer").Select
Range("a1").PasteSpecial xlPasteValues

' - rendezi dátum szerint csökkenõbe -'

Columns("w:w").Select
Selection.End(xlDown).Select
Dim Dtrw As Long
Dtrw = ActiveCell.Row
Dim Dtoszlop As String
Dtoszlop = "w"
Dim Dtkoord As String
Dtkoord = Dtoszlop & Dtrw

Range("a1", Dtkoord).Select
    ActiveWorkbook.Worksheets("szûrõ_transfer").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("szûrõ_transfer").Sort.SortFields.Add Key:=Range("c1"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortOnValues
            
            With ActiveWorkbook.Worksheets("szûrõ_transfer").Sort
        .SetRange Range("A2", Dtkoord)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Dim rngList As Range
Set rngList = Munka6.Range("a1", ALkoord)
AppWindow.ListBox33.List = rngList.Value


Sheets("Start").Select
Range("b2").Select




End Sub
