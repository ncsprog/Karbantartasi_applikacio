Attribute VB_Name = "FunctionAdatfelv�telLista3"
Option Explicit

Sub Adatfelv�telLista3()
'Jelsz�Rejt�s2
' - Lista koordin�ta - '

Sheets("adatok").Select
Columns("u:u").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.Row
Dim ALoszlop As String
ALoszlop = "u"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

' - Lista ki�r�s - '

Dim rngList As Range
Set rngList = Munka1.Range("a1", ALkoord)
AppWindow.ListBox22.List = rngList.Value

Sheets("Start").Select
Range("b2").Select
'Jelsz�Rejt�s
End Sub
