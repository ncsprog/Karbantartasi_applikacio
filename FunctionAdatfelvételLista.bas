Attribute VB_Name = "FunctionAdatfelv�telLista"
Option Explicit

Sub Adatfelv�telLista()
'Jelsz�Rejt�s2
' - Lista koordin�ta - '

Sheets("adatok").Select
Columns("q:q").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.row
Dim ALoszlop As String
ALoszlop = "q"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

' - Lista ki�r�s - '

Dim rngList As Range
Set rngList = Munka1.Range("a1", ALkoord)
AppWindow.ListBox7.List = rngList.Value

Sheets("Start").Select
Range("b2").Select
'Jelsz�Rejt�s
End Sub

