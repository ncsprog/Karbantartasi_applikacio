Attribute VB_Name = "FunctionL�tsz�m�sszes�t�s"
Option Explicit

Sub L�tsz�m�sszes�t�s()
'Jelsz�Rejt�s2

Sheets("l�tsz�m").Select
Columns("c:c").Select
Selection.End(xlDown).Select
Dim L�_rw As Long
L�_rw = ActiveCell.Row
Dim L�_Koszlop As String
L�_Koszlop = "c"
Dim L�_Voszlop As String
L�_Voszlop = "k"
Dim kezd As String
kezd = L�_Koszlop & L�_rw
Dim v�g As String
v�g = L�_Voszlop & L�_rw



Dim l�tsz As Integer
l�tsz = Application.WorksheetFunction.Sum(Range(kezd, v�g))

Sheets("l�tsz�m").Select
Columns("m:m").Select
Selection.End(xlDown).Select
Dim szumRw As Long
szumRw = ActiveCell.Row + 1
Dim szumC As String
szumC = "m"
Dim szumRng As String
szumRng = szumC & szumRw

Range(szumRng).Value = l�tsz

Sheets("Start").Select
Range("b2").Select
'Jelsz�Rejt�s
End Sub
