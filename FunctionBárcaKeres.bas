Attribute VB_Name = "FunctionB�rcaKeres"
Option Explicit

Sub B�rcaKeres()
'Jelsz�Rejt�s2
' - Lista koordin�ta - '

Sheets("n�vsor").Select
Columns("d:d").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.row
Dim ALoszlop As String
ALoszlop = "d"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

' - Lista ki�r�s - '

Dim rngList As Range
Set rngList = Munka3.Range("a1", ALkoord)
AppWindow.ListBox21.List = rngList.Value

Sheets("Start").Select
Range("b2").Select

Worksheets("n�vsor").Visible = False
'Jelsz�Rejt�s
End Sub
