Attribute VB_Name = "FunctionAdatfelvételLista3"
Option Explicit

Sub AdatfelvételLista3()
'JelszóRejtés2
' - Lista koordináta - '

Sheets("adatok").Select
Columns("u:u").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.Row
Dim ALoszlop As String
ALoszlop = "u"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

' - Lista kiírás - '

Dim rngList As Range
Set rngList = Munka1.Range("a1", ALkoord)
AppWindow.ListBox22.List = rngList.Value

Sheets("Start").Select
Range("b2").Select
'JelszóRejtés
End Sub
