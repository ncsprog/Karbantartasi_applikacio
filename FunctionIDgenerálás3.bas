Attribute VB_Name = "FunctionIDgener�l�s3"
Option Explicit

Sub IDgener�l�s3()

Sheets("Megbesz�l�s").Select
Columns("a:a").Select
Selection.End(xlDown).Select
Dim IDnr As Long
IDnr = ActiveCell + 1
Dim IDrw As Long
IDrw = ActiveCell.row + 1
Dim IDoszlop As String
IDoszlop = "a"
Dim IDkoord As String
IDkoord = IDoszlop & IDrw
Range(IDkoord) = IDnr


Sheets("Start").Select
Range("b2").Select
End Sub
