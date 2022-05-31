Attribute VB_Name = "FunctionIDgenerálás"
Option Explicit

Sub IDgenerálás()
'JelszóRejtés2

Sheets("adatok").Select
Columns("a:a").Select
Selection.End(xlDown).Select
Dim IDnr As Long
IDnr = ActiveCell + 1
Dim IDrw As Long
IDrw = ActiveCell.Row + 1
Dim IDoszlop As String
IDoszlop = "a"
Dim IDkoord As String
IDkoord = IDoszlop & IDrw
Range(IDkoord) = IDnr


Sheets("Start").Select
Range("b2").Select
'JelszóRejtés
End Sub
