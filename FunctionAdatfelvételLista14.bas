Attribute VB_Name = "FunctionAdatfelv�telLista14"
Option Explicit

Sub Adatfelv�telLista14()

Sheets("alapadatok").Select
Columns("j:j").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.Row
Dim ALoszlop As String
ALoszlop = "j"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

' - Lista ki�r�s - '

Dim rngList As Range
Set rngList = Munka12.Range("i2", ALkoord)
AppWindow.ListBox37.List = rngList.Value

Sheets("Start").Select
Range("b2").Select

End Sub
