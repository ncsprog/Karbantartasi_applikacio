Attribute VB_Name = "FunctionAdatfelv�telLista11"
Option Explicit

Sub Adatfelv�telLista11()
' - Felel�s�k - '
Sheets("alapadatok").Select
Columns("d:d").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.Row
Dim ALoszlop As String
ALoszlop = "d"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

' - Lista ki�r�s - '

Dim rngList As Range
Set rngList = Munka12.Range("c2", ALkoord)
AppWindow.ListBox30.List = rngList.Value

Sheets("Start").Select
Range("b2").Select
End Sub
