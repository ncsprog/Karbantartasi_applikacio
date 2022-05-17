Attribute VB_Name = "FunctionAdatfelvételLista"
Option Explicit

Sub AdatfelvételLista()
'JelszóRejtés2
' - Lista koordináta - '

Sheets("adatok").Select
Columns("q:q").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.row
Dim ALoszlop As String
ALoszlop = "q"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

' - Lista kiírás - '

Dim rngList As Range
Set rngList = Munka1.Range("a1", ALkoord)
AppWindow.ListBox7.List = rngList.Value

Sheets("Start").Select
Range("b2").Select
'JelszóRejtés
End Sub

