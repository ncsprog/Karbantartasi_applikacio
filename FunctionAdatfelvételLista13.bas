Attribute VB_Name = "FunctionAdatfelv�telLista13"
Option Explicit

Sub Adatfelv�telLista13()
' - Applik�ci�b�l sz�rmaz� adatok - '

' - ki�r�tit a transfert - '
Sheets("sz�r�_transfer").Select
Range("a1:xx10000") = ""
' - �tm�solja a kezelend� adatokat - '
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
Sheets("sz�r�_transfer").Select
Range("a1").PasteSpecial xlPasteValues

' - rendezi d�tum szerint cs�kken�be -'

Columns("w:w").Select
Selection.End(xlDown).Select
Dim Dtrw As Long
Dtrw = ActiveCell.Row
Dim Dtoszlop As String
Dtoszlop = "w"
Dim Dtkoord As String
Dtkoord = Dtoszlop & Dtrw

Range("a1", Dtkoord).Select
    ActiveWorkbook.Worksheets("sz�r�_transfer").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("sz�r�_transfer").Sort.SortFields.Add Key:=Range("c1"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortOnValues
            
            With ActiveWorkbook.Worksheets("sz�r�_transfer").Sort
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
