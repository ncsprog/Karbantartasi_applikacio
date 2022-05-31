Attribute VB_Name = "FunctionAdatfelvételLista12"
Option Explicit

Sub AdatfelvételLista12()
' - Összlétszám - '

Sheets("alapadatok").Select
Columns("g:g").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.Row
Dim ALoszlop As String
ALoszlop = "g"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

' - Lista kiírás - '

Dim rngList As Range
Set rngList = Munka12.Range("f2", ALkoord)
AppWindow.ListBox34.List = rngList.Value

Sheets("Start").Select
Range("b2").Select

End Sub
