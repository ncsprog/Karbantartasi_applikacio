Attribute VB_Name = "FunctionL�tsz�m�sszes�t�s"
Option Explicit

Sub L�tsz�m�sszes�t�s()


Sheets("L�tsz�m").Select
Columns("a:a").Select
Selection.End(xlDown).Select
Dim L�_rw As Long
L�_rw = ActiveCell.row
Dim L�_Koszlop As String
L�_Koszlop = "c"
Dim L�_Voszlop As String
L�_Voszlop = "af"
Dim kezd As String
kezd = L�_Koszlop & L�_rw
Dim v�g As String
v�g = L�_Voszlop & L�_rw


Dim l�tsz�m As Integer
l�tsz�m = Application.WorksheetFunction.Sum(Range(kezd, v�g))
AppWindow.TextBox80.Value = l�tsz�m

End Sub
