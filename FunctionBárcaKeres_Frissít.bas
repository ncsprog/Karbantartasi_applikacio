Attribute VB_Name = "FunctionBárcaKeres_Frissít"
Option Explicit

Sub BárcaKeres_Frissít()
'JelszóRejtés2
' - Lista koordináta - '

Sheets("névsor").Select
Columns("d:d").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.Row
Dim ALoszlop As String
ALoszlop = "d"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

' - Lista kiírás - '

Dim rngList As Range
Set rngList = Munka3.Range("a1", ALkoord)
AppWindow.ListBox35.List = rngList.Value

Sheets("Start").Select
Range("b2").Select

'Worksheets("névsor").Visible = False
'JelszóRejtés
End Sub
