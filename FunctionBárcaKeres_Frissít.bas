Attribute VB_Name = "FunctionB�rcaKeres_Friss�t"
Option Explicit

Sub B�rcaKeres_Friss�t()
' - Lista koordin�ta - '

Sheets("n�vsor").Select
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
Set rngList = Munka3.Range("a1", ALkoord)
AppWindow.ListBox35.List = rngList.Value

Sheets("Start").Select
Range("b2").Select

End Sub
