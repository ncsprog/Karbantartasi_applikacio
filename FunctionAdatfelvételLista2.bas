Attribute VB_Name = "FunctionAdatfelv�telLista2"
Option Explicit

Sub Adatfelv�telLista2()

' - Lista koordin�ta - '

Sheets("adatok").Select
Columns("u:u").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.row
Dim ALoszlop As String
ALoszlop = "u"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

' - Lista ki�r�s - '

Dim rngList As Range
Set rngList = Munka1.Range("a1", ALkoord)
AppWindow.ListBox20.List = rngList.Value


End Sub
