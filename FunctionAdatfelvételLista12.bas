Attribute VB_Name = "FunctionAdatfelv�telLista12"
Option Explicit

Sub Adatfelv�telLista12()
' - �sszl�tsz�m - '

Sheets("alapadatok").Select
Columns("g:g").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.Row
Dim ALoszlop As String
ALoszlop = "g"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

' - Lista ki�r�s - '

Dim rngList As Range
Set rngList = Munka12.Range("f2", ALkoord)
AppWindow.ListBox34.List = rngList.Value

Sheets("Start").Select
Range("b2").Select

End Sub
