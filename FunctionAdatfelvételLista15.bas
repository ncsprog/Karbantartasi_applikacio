Attribute VB_Name = "FunctionAdatfelv�telLista15"
Option Explicit

Sub Adatfelv�telLista15()

Sheets("alapadatok").Select
Columns("p:p").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.Row
Dim ALoszlop As String
ALoszlop = "p"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

' - Lista ki�r�s - '

Dim rngList As Range
Set rngList = Munka12.Range("o2", ALkoord)
AppWindow.ListBox38.List = rngList.Value

Sheets("Start").Select
Range("b2").Select

End Sub
