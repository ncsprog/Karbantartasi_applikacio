Attribute VB_Name = "FunctionAdatfelvételLista16"
Option Explicit

Sub AdatfelvételLista16()

Sheets("alapadatok").Select
Columns("m:m").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.Row
Dim ALoszlop As String
ALoszlop = "m"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

' - Lista kiírás - '

Dim rngList As Range
Set rngList = Munka12.Range("l2", ALkoord)
AppWindow.ListBox39.List = rngList.Value

Sheets("Start").Select
Range("b2").Select

End Sub
