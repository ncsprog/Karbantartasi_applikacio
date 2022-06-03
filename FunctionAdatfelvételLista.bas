Attribute VB_Name = "FunctionAdatfelvételLista"
Option Explicit

Sub AdatfelvételLista()
' - Lista koordináta - '
Sheets("adatok").Select
Columns("a:a").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.Row
Dim ALoszlop As String
ALoszlop = "q"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw
Dim Koord0 As String
Koord0 = "a" & Munka12.Range("aa3").Value
' - Lista kiírás - '
Dim rngList As Range
Set rngList = Munka1.Range(Koord0, ALkoord)
AppWindow.ListBox7.List = rngList.Value

Sheets("Start").Select
Range("b2").Select
End Sub

