Attribute VB_Name = "FunctionAdatfelvételLista14"
Option Explicit

Sub AdatfelvételLista14()

Sheets("alapadatok").Select
Columns("j:j").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.Row
Dim ALoszlop As String
ALoszlop = "j"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

' - Lista kiírás - '

Dim rngList As Range
Set rngList = Munka12.Range("i2", ALkoord)
AppWindow.ListBox37.List = rngList.Value

Sheets("Start").Select
Range("b2").Select

End Sub
