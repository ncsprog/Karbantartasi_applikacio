Attribute VB_Name = "FunctionAdatfelv�telLista10"
Option Explicit

Sub Adatfelv�telLista10()
' - St�tuszok kezel�se - '
Sheets("alapadatok").Select
Columns("b:b").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.Row
Dim ALoszlop As String
ALoszlop = "b"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

' - Lista ki�r�s - '

Dim rngList As Range
Set rngList = Munka12.Range("a2", ALkoord)
AppWindow.ListBox29.List = rngList.Value

Sheets("Start").Select
Range("b2").Select

End Sub
