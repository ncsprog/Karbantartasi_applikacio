Attribute VB_Name = "FunctionG�pKeres_Friss�t"
Option Explicit

Sub G�pKeres()
' - Friss�t�s - '

' - Lista koordin�ta - '

Sheets("g�pek").Select
Columns("c:c").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.Row
Dim ALoszlop As String
ALoszlop = "c"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

' - Lista ki�r�s - '

Dim rngList As Range
Set rngList = Munka4.Range("a1", ALkoord)
AppWindow.ListBox36.List = rngList.Value

Sheets("Start").Select
Range("b2").Select

End Sub
