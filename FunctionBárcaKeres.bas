Attribute VB_Name = "FunctionBárcaKeres"
Option Explicit

Sub BárcaKeres()

' - Lista koordináta - '

Sheets("névsor").Select
Columns("d:d").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.row
Dim ALoszlop As String
ALoszlop = "d"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

' - Lista kiírás - '

Dim rngList As Range
Set rngList = Munka3.Range("a1", ALkoord)
AppWindow.ListBox21.List = rngList.Value

End Sub
